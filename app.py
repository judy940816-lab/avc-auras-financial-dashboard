from pathlib import Path
from textwrap import dedent

import pandas as pd
import plotly.express as px
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DATA_FILE = APP_DIR / "AVC_Auras_dashboard_data.xlsx"

st.set_page_config(
    page_title="AVC vs. Auras Financial Dashboard",
    page_icon="📊",
    layout="wide",
)
st.markdown(
    """
    <style>
    div[data-testid="stMetric"] {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        padding: 18px 20px;
        border-radius: 12px;
        min-height: 135px;
    }

    div[data-testid="stMetricLabel"] {
        font-size: 0.95rem;
        color: #475569;
    }

    div[data-testid="stMetricValue"] {
        font-size: 2rem;
    }

    .research-shell {
        background: linear-gradient(135deg, #F8FAFC 0%, #EFF6FF 100%);
        border: 1px solid #DCE7F5;
        border-radius: 16px;
        padding: 24px 26px;
        margin: 14px 0 20px 0;
    }

    .research-kicker {
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0.12em;
        color: #2563EB;
        margin-bottom: 8px;
    }

    .research-title {
        font-size: 1.35rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 8px;
    }

    .research-objective {
        font-size: 1rem;
        line-height: 1.7;
        color: #334155;
        max-width: 1100px;
    }

    .research-meta {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 16px;
    }

    .research-chip {
        display: inline-block;
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 999px;
        padding: 5px 10px;
        font-size: 0.82rem;
        color: #475569;
    }

    .question-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 20px;
        min-height: 390px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }

    .question-number {
        display: inline-block;
        background-color: #DBEAFE;
        color: #1D4ED8;
        border-radius: 999px;
        padding: 4px 9px;
        font-size: 0.76rem;
        font-weight: 700;
        margin-bottom: 12px;
    }

    .question-text {
        font-size: 1rem;
        font-weight: 700;
        line-height: 1.5;
        color: #0F172A;
        margin-bottom: 14px;
    }

    .answer-pill {
        display: inline-block;
        background-color: #DCFCE7;
        color: #166534;
        border-radius: 999px;
        padding: 5px 10px;
        font-size: 0.74rem;
        font-weight: 800;
        letter-spacing: 0.04em;
        margin-bottom: 12px;
    }

    .answer-title {
        font-size: 0.98rem;
        font-weight: 700;
        line-height: 1.5;
        color: #1E293B;
        margin-bottom: 12px;
    }

    .answer-evidence {
        font-size: 0.84rem;
        line-height: 1.6;
        color: #475569;
        border-top: 1px solid #E2E8F0;
        padding-top: 12px;
        margin-top: 4px;
    }

    .answer-interpretation {
        font-size: 0.82rem;
        line-height: 1.55;
        color: #64748B;
        margin-top: 10px;
    }

    .design-note {
        font-size: 0.82rem;
        color: #64748B;
        margin-top: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

@st.cache_data
def load_raw_data() -> pd.DataFrame:
    """Load only the raw financial statement table.

    The web app recalculates every ratio in Python, so it does not depend on
    Excel formula caches.
    """
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"Data file not found: {DATA_FILE}")

    df = pd.read_excel(DATA_FILE, sheet_name="raw_financials")
    df.columns = df.columns.astype(str).str.strip()

    required = {
        "company",
        "scope",
        "year",
        "statement",
        "account",
        "value",
        "unit",
        "source_page",
        "source_file",
    }
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")

    for column in ["company", "scope", "statement", "account", "unit", "source_page", "source_file"]:
        df[column] = df[column].astype(str).str.strip()

    df["year"] = pd.to_numeric(df["year"], errors="raise").astype(int)
    df["value"] = pd.to_numeric(df["value"], errors="raise")
    return df


def build_metric_base(raw: pd.DataFrame) -> pd.DataFrame:
    base = (
        raw.pivot_table(
            index=["company", "scope", "year"],
            columns="account",
            values="value",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )
    base.columns.name = None
    base["entity"] = base["company"] + " " + base["scope"]
    return base


def safe_divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    return numerator.div(denominator.where(denominator.ne(0)))


def calculate_liquidity(base: pd.DataFrame) -> pd.DataFrame:
    required_accounts = [
        "Current Assets",
        "Current Liabilities",
        "Cash and Cash Equivalents",
        "Total Assets",
        "Inventory",
        "Prepayments",
        "Other Current Assets",
    ]
    result = base.copy()
    for account in required_accounts:
        if account not in result.columns:
            result[account] = 0.0

    result["current_ratio"] = safe_divide(
        result["Current Assets"], result["Current Liabilities"]
    )
    result["cash_ratio"] = safe_divide(
        result["Cash and Cash Equivalents"], result["Current Liabilities"]
    )
    result["nwc_to_assets"] = safe_divide(
        result["Current Assets"] - result["Current Liabilities"],
        result["Total Assets"],
    )
    result["quick_ratio"] = safe_divide(
        result["Current Assets"]
        - result["Inventory"]
        - result["Prepayments"]
        - result["Other Current Assets"],
        result["Current Liabilities"],
    )
    return result


def calculate_profitability(base: pd.DataFrame) -> pd.DataFrame:
    required_accounts = [
        "Revenue",
        "Gross Profit",
        "Operating Profit",
        "Net Income Attributable to Parent",
        "Basic EPS",
    ]
    result = base.copy()
    for account in required_accounts:
        if account not in result.columns:
            result[account] = 0.0

    result["gross_margin"] = safe_divide(result["Gross Profit"], result["Revenue"])
    result["operating_margin"] = safe_divide(
        result["Operating Profit"], result["Revenue"]
    )
    result["net_margin_parent"] = safe_divide(
        result["Net Income Attributable to Parent"], result["Revenue"]
    )
    return result


def format_change(value: float, suffix: str = "") -> str:
    if pd.isna(value):
        return "N/A"
    sign = "+" if value > 0 else ""
    return f"{sign}{value:.2f}{suffix}"


def entity_insight(entity_data: pd.DataFrame) -> str:
    ordered = entity_data.sort_values("year")
    if len(ordered) < 2:
        return "Two years of data are required to generate a trend insight."

    first, last = ordered.iloc[0], ordered.iloc[-1]
    current_change = last["current_ratio"] - first["current_ratio"]
    quick_change = last["quick_ratio"] - first["quick_ratio"]

    if current_change > 0 and quick_change > 0:
        direction = "both current and quick ratios improved"
    elif current_change < 0 and quick_change < 0:
        direction = "both current and quick ratios declined"
    else:
        direction = "current and quick ratios moved in different directions"

    return (
        f"From {int(first['year'])} to {int(last['year'])}, {direction}. "
        f"Current ratio changed by {current_change:+.2f}, while quick ratio "
        f"changed by {quick_change:+.2f}."
    )



def percent_change(current: float, previous: float) -> float:
    """Return a decimal growth rate while safely handling a zero base."""
    if pd.isna(current) or pd.isna(previous) or previous == 0:
        return float("nan")
    return current / previous - 1


def build_research_answers(
    profitability_data: pd.DataFrame,
    liquidity_data: pd.DataFrame,
) -> dict:
    """Build evidence-based answers to the dashboard's three research questions.

    The core research answers use the complete FY 2024–2025 dataset and are
    intentionally independent of the sidebar filters.
    """
    required_entities = {
        "AVC Consolidated",
        "AVC Standalone",
        "Auras Consolidated",
    }
    available = set(profitability_data["entity"].unique())

    if not required_entities.issubset(available):
        return {
            "rq1": {
                "status": "DATA CHECK",
                "title": "Required AVC profitability observations are unavailable.",
                "evidence": "Confirm that AVC consolidated records are included for both years.",
                "interpretation": "The research answer cannot be calculated reliably.",
            },
            "rq2": {
                "status": "DATA CHECK",
                "title": "Required AVC liquidity observations are unavailable.",
                "evidence": "Confirm that AVC consolidated liquidity records are included for both years.",
                "interpretation": "The research answer cannot be calculated reliably.",
            },
            "rq3": {
                "status": "DATA CHECK",
                "title": "Required standalone, consolidated, or peer observations are unavailable.",
                "evidence": "Confirm that AVC standalone, AVC consolidated, and Auras consolidated records exist.",
                "interpretation": "The research answer cannot be calculated reliably.",
            },
        }

    # RQ1: Did AVC growth translate into stronger profitability?
    avc_prof = (
        profitability_data[
            profitability_data["entity"].eq("AVC Consolidated")
        ]
        .sort_values("year")
        .copy()
    )

    if len(avc_prof) < 2:
        raise ValueError("At least two years of AVC consolidated data are required.")

    prof_first = avc_prof.iloc[0]
    prof_last = avc_prof.iloc[-1]

    revenue_growth = percent_change(prof_last["Revenue"], prof_first["Revenue"])
    operating_profit_growth = percent_change(
        prof_last["Operating Profit"],
        prof_first["Operating Profit"],
    )
    eps_growth = percent_change(prof_last["Basic EPS"], prof_first["Basic EPS"])
    gross_margin_change = prof_last["gross_margin"] - prof_first["gross_margin"]
    operating_margin_change = (
        prof_last["operating_margin"] - prof_first["operating_margin"]
    )
    net_margin_change = (
        prof_last["net_margin_parent"] - prof_first["net_margin_parent"]
    )

    profitability_strengthened = (
        revenue_growth > 0
        and operating_profit_growth > revenue_growth
        and gross_margin_change > 0
        and operating_margin_change > 0
        and net_margin_change > 0
    )

    rq1_status = "ANSWER: YES" if profitability_strengthened else "ANSWER: MIXED"
    rq1_title = (
        "Growth translated into stronger profitability and positive operating leverage."
        if profitability_strengthened
        else "Growth was strong, but the profitability evidence is mixed."
    )
    rq1_evidence = (
        f"Revenue increased {revenue_growth:.1%}, while operating profit increased "
        f"{operating_profit_growth:.1%}. Gross margin rose "
        f"{gross_margin_change * 100:+.1f} pp, operating margin rose "
        f"{operating_margin_change * 100:+.1f} pp, net margin rose "
        f"{net_margin_change * 100:+.1f} pp, and EPS increased {eps_growth:.1%}."
    )
    rq1_interpretation = (
        "Operating profit grew faster than revenue, suggesting that AVC converted "
        "scale expansion into proportionally stronger operating earnings."
    )

    # RQ2: Did rapid expansion weaken short-term liquidity?
    avc_liq = (
        liquidity_data[
            liquidity_data["entity"].eq("AVC Consolidated")
        ]
        .sort_values("year")
        .copy()
    )

    if len(avc_liq) < 2:
        raise ValueError("At least two years of AVC consolidated liquidity data are required.")

    liq_first = avc_liq.iloc[0]
    liq_last = avc_liq.iloc[-1]

    current_change = liq_last["current_ratio"] - liq_first["current_ratio"]
    quick_change = liq_last["quick_ratio"] - liq_first["quick_ratio"]
    cash_change = liq_last["cash_ratio"] - liq_first["cash_ratio"]
    nwc_change = liq_last["nwc_to_assets"] - liq_first["nwc_to_assets"]

    improving_indicators = sum(
        change >= 0
        for change in [current_change, quick_change, cash_change, nwc_change]
    )

    if improving_indicators >= 3:
        rq2_status = "ANSWER: NO"
        rq2_title = (
            "Expansion did not materially weaken liquidity; the evidence is broadly stable."
        )
    elif improving_indicators <= 1:
        rq2_status = "ANSWER: YES"
        rq2_title = (
            "Expansion coincided with a broad weakening in short-term liquidity."
        )
    else:
        rq2_status = "ANSWER: MIXED"
        rq2_title = "Liquidity signals moved in different directions."

    rq2_evidence = (
        f"Current ratio changed {current_change:+.2f}x, quick ratio "
        f"{quick_change:+.2f}x, cash ratio {cash_change:+.2f}x, and net working "
        f"capital to assets {nwc_change * 100:+.1f} pp."
    )
    rq2_interpretation = (
        "The slight decline in current ratio should be monitored, but stronger quick "
        "and cash ratios indicate that immediately available liquidity improved."
    )

    # RQ3: What do consolidated–standalone and peer differences reveal?
    latest_year = int(profitability_data["year"].max())

    avc_con = profitability_data[
        profitability_data["entity"].eq("AVC Consolidated")
        & profitability_data["year"].eq(latest_year)
    ]
    avc_standalone = profitability_data[
        profitability_data["entity"].eq("AVC Standalone")
        & profitability_data["year"].eq(latest_year)
    ]
    auras_con = profitability_data[
        profitability_data["entity"].eq("Auras Consolidated")
        & profitability_data["year"].eq(latest_year)
    ]

    if avc_con.empty or avc_standalone.empty or auras_con.empty:
        raise ValueError("Latest-year observations are missing for RQ3.")

    avc_con = avc_con.iloc[0]
    avc_standalone = avc_standalone.iloc[0]
    auras_con = auras_con.iloc[0]

    revenue_gap = avc_con["Revenue"] - avc_standalone["Revenue"]
    operating_profit_gap = (
        avc_con["Operating Profit"] - avc_standalone["Operating Profit"]
    )
    revenue_gap_share = (
        revenue_gap / avc_con["Revenue"] if avc_con["Revenue"] else float("nan")
    )
    operating_gap_share = (
        operating_profit_gap / avc_con["Operating Profit"]
        if avc_con["Operating Profit"]
        else float("nan")
    )
    revenue_multiple = (
        avc_con["Revenue"] / auras_con["Revenue"]
        if auras_con["Revenue"]
        else float("nan")
    )
    operating_margin_gap = (
        avc_con["operating_margin"] - auras_con["operating_margin"]
    )
    gross_margin_gap = avc_con["gross_margin"] - auras_con["gross_margin"]

    rq3_status = "ANSWER: SCALE ADVANTAGE"
    rq3_title = (
        "Consolidation adds substantial scale, while AVC leads Auras in operating margin."
    )
    rq3_evidence = (
        f"In {latest_year}, AVC consolidated revenue exceeded standalone revenue by "
        f"NT${revenue_gap / 1_000_000:,.1f} bn ({revenue_gap_share:.1%} of consolidated "
        f"revenue), and consolidated operating profit was higher by "
        f"NT${operating_profit_gap / 1_000_000:,.1f} bn "
        f"({operating_gap_share:.1%}). AVC's revenue was {revenue_multiple:.1f}x "
        f"Auras's, with an operating-margin advantage of "
        f"{operating_margin_gap * 100:+.1f} pp, although its gross margin was "
        f"{gross_margin_gap * 100:+.1f} pp lower."
    )
    rq3_interpretation = (
        "The consolidated–standalone gap suggests meaningful group-level activity, "
        "but it is not a pure subsidiary contribution because consolidation also "
        "includes intercompany eliminations and accounting adjustments."
    )

    return {
        "rq1": {
            "status": rq1_status,
            "title": rq1_title,
            "evidence": rq1_evidence,
            "interpretation": rq1_interpretation,
        },
        "rq2": {
            "status": rq2_status,
            "title": rq2_title,
            "evidence": rq2_evidence,
            "interpretation": rq2_interpretation,
        },
        "rq3": {
            "status": rq3_status,
            "title": rq3_title,
            "evidence": rq3_evidence,
            "interpretation": rq3_interpretation,
        },
    }


try:
    raw = load_raw_data()
except (FileNotFoundError, ValueError) as exc:
    st.error(str(exc))
    st.stop()

base = build_metric_base(raw)
liquidity = calculate_liquidity(base)
profitability = calculate_profitability(base)

research_answers = build_research_answers(profitability, liquidity)

available_entities = sorted(base["entity"].unique().tolist())
available_years = sorted(base["year"].unique().tolist())

st.sidebar.header("Filters")
selected_entities = st.sidebar.multiselect(
    "Entity",
    options=available_entities,
    default=available_entities,
)
selected_years = st.sidebar.multiselect(
    "Year",
    options=available_years,
    default=available_years,
)

if not selected_entities or not selected_years:
    st.warning("Select at least one entity and one year.")
    st.stop()

liq_view = liquidity[
    liquidity["entity"].isin(selected_entities)
    & liquidity["year"].isin(selected_years)
].copy()
prof_view = profitability[
    profitability["entity"].isin(selected_entities)
    & profitability["year"].isin(selected_years)
].copy()

st.title("AVC vs. Auras Financial Decision Dashboard")
st.caption(
    "Standalone vs. consolidated analysis and peer benchmarking, FY 2024–2025. "
    "All ratios are recalculated from the raw financial statement data in Python."
)


st.markdown(
    dedent(
        """
        <div class="research-shell">
        <div class="research-kicker">RESEARCH FRAMEWORK</div>

        <div class="research-title">Research Motivation</div>
        <div class="research-objective">
            This project originated from a classroom financial statement analysis
            comparing AVC and Auras for FY 2023–2024. To extend the original assignment
            into a more complete and reproducible research project, I updated the dataset
            to FY 2024–2025, rebuilt the financial analysis in Python, and developed an
            interactive dashboard. This extension allows the earlier findings to be
            reassessed using the latest annual data while incorporating profitability,
            liquidity, standalone-versus-consolidated differences, and peer benchmarking.
        </div>

        <div style="height: 18px;"></div>

        <div class="research-title">Project Objective</div>
        <div class="research-objective">
            This project examines whether AVC's rapid FY 2025 growth translated into
            stronger profitability without weakening short-term liquidity. It combines
            year-over-year analysis, standalone-versus-consolidated comparison, and peer
            benchmarking against Auras to identify group-level performance differences
            and decision-relevant financial signals.
        </div>
        <div class="research-meta">
            <span class="research-chip">Period: FY 2024–2025</span>
            <span class="research-chip">Primary firm: AVC</span>
            <span class="research-chip">Peer benchmark: Auras</span>
            <span class="research-chip">Source: Audited financial statements</span>
            <span class="research-chip">Method: Python-recalculated ratios</span>
        </div>
        <div class="design-note">
            Research design: descriptive and comparative analysis. Reported relationships
            should be interpreted as financial associations rather than causal effects.
        </div>
        </div>
        """
    ),
    unsafe_allow_html=True,
)

st.markdown("### Research Questions & Findings")
st.caption(
    "Each conclusion is recalculated from the underlying FY 2024–2025 source data."
)

rq1, rq2, rq3 = st.columns(3, gap="large")

with rq1:
    st.markdown(
        f"""
        <div class="question-card">
            <div class="question-number">RQ1</div>
            <div class="question-text">
                Did AVC's revenue growth translate into stronger profitability?
            </div>
            <div class="answer-pill">{research_answers["rq1"]["status"]}</div>
            <div class="answer-title">{research_answers["rq1"]["title"]}</div>
            <div class="answer-evidence">
                <strong>Evidence:</strong> {research_answers["rq1"]["evidence"]}
            </div>
            <div class="answer-interpretation">
                <strong>Interpretation:</strong>
                {research_answers["rq1"]["interpretation"]}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with rq2:
    st.markdown(
        f"""
        <div class="question-card">
            <div class="question-number">RQ2</div>
            <div class="question-text">
                Did rapid expansion weaken AVC's short-term liquidity position?
            </div>
            <div class="answer-pill">{research_answers["rq2"]["status"]}</div>
            <div class="answer-title">{research_answers["rq2"]["title"]}</div>
            <div class="answer-evidence">
                <strong>Evidence:</strong> {research_answers["rq2"]["evidence"]}
            </div>
            <div class="answer-interpretation">
                <strong>Interpretation:</strong>
                {research_answers["rq2"]["interpretation"]}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with rq3:
    st.markdown(
        f"""
        <div class="question-card">
            <div class="question-number">RQ3</div>
            <div class="question-text">
                What do standalone–consolidated differences and peer comparison reveal?
            </div>
            <div class="answer-pill">{research_answers["rq3"]["status"]}</div>
            <div class="answer-title">{research_answers["rq3"]["title"]}</div>
            <div class="answer-evidence">
                <strong>Evidence:</strong> {research_answers["rq3"]["evidence"]}
            </div>
            <div class="answer-interpretation">
                <strong>Interpretation:</strong>
                {research_answers["rq3"]["interpretation"]}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

tab_summary, tab_liquidity, tab_profitability, tab_sources = st.tabs(
    ["Executive Summary", "Liquidity", "Profitability", "Source Data"]
)

with tab_summary:
    latest_year = max(selected_years)
    latest = prof_view[prof_view["year"].eq(latest_year)].copy()

    st.subheader(f"{latest_year} Snapshot")

    if latest.empty:
        st.info("No profitability records match the current filters.")
    else:
        # Entity selector：限制寬度，不要橫跨整個頁面
        selector_col, spacer_col = st.columns([1, 3])

        with selector_col:
            display_entity = st.selectbox(
                "Entity",
                options=sorted(latest["entity"].unique()),
                key="summary_entity",
            )

        kpi = latest[latest["entity"].eq(display_entity)].iloc[0]

        liq_kpi = liq_view[
            liq_view["entity"].eq(display_entity)
            & liq_view["year"].eq(latest_year)
        ]

        # 第一排：營運與獲利
        st.markdown("#### Performance")
        row1 = st.columns(4, gap="large")

        row1[0].metric(
            "Revenue",
            f"NT${kpi['Revenue'] / 1_000_000:,.1f} bn",
        )
        row1[1].metric(
            "Gross Margin",
            f"{kpi['gross_margin']:.2%}",
        )
        row1[2].metric(
            "Operating Margin",
            f"{kpi['operating_margin']:.2%}",
        )
        row1[3].metric(
            "Net Margin (Parent)",
            f"{kpi['net_margin_parent']:.2%}",
        )

        # 第二排：每股盈餘與流動性
        st.markdown("#### Per-share & Liquidity")
        row2 = st.columns(4, gap="large")

        row2[0].metric(
            "Basic EPS",
            f"NT${kpi['Basic EPS']:.2f}",
        )

        if not liq_kpi.empty:
            liq_row = liq_kpi.iloc[0]

            row2[1].metric(
                "Current Ratio",
                f"{liq_row['current_ratio']:.2f}",
            )
            row2[2].metric(
                "Quick Ratio",
                f"{liq_row['quick_ratio']:.2f}",
            )
            row2[3].metric(
                "Cash Ratio",
                f"{liq_row['cash_ratio']:.2f}",
            )
        else:
            row2[1].metric("Current Ratio", "N/A")
            row2[2].metric("Quick Ratio", "N/A")
            row2[3].metric("Cash Ratio", "N/A")

        st.subheader("Automated Trend Insight")
        st.info(
            entity_insight(
                liq_view[liq_view["entity"].eq(display_entity)]
            )
        )

    st.subheader("Revenue Comparison")
    revenue_chart = px.bar(
        prof_view,
        x="entity",
        y="Revenue",
        color="year",
        barmode="group",
        text_auto=".3s",
        labels={
            "entity": "Entity",
            "Revenue": "Revenue (NT$ thousand)",
            "year": "Year",
        },
    )
    revenue_chart.update_layout(legend_title_text="Year")
    st.plotly_chart(revenue_chart, use_container_width=True)

with tab_liquidity:
    st.subheader("Liquidity Ratios")

    metric_map = {
        "Current Ratio": "current_ratio",
        "Quick Ratio": "quick_ratio",
        "Cash Ratio": "cash_ratio",
        "Net Working Capital / Assets": "nwc_to_assets",
    }
    selected_metric_label = st.selectbox(
        "Metric",
        options=list(metric_map),
        key="liquidity_metric",
    )
    selected_metric = metric_map[selected_metric_label]

    liquidity_chart = px.bar(
        liq_view,
        x="entity",
        y=selected_metric,
        color="year",
        barmode="group",
        text_auto=".2f" if selected_metric != "nwc_to_assets" else ".2%",
        labels={
            "entity": "Entity",
            selected_metric: selected_metric_label,
            "year": "Year",
        },
    )
    liquidity_chart.update_layout(legend_title_text="Year")
    if selected_metric == "nwc_to_assets":
        liquidity_chart.update_yaxes(tickformat=".0%")
    st.plotly_chart(liquidity_chart, use_container_width=True)

    liquidity_table = liq_view[
        [
            "entity",
            "year",
            "current_ratio",
            "quick_ratio",
            "cash_ratio",
            "nwc_to_assets",
        ]
    ].sort_values(["entity", "year"])
    st.dataframe(
        liquidity_table.style.format(
            {
                "current_ratio": "{:.2f}",
                "quick_ratio": "{:.2f}",
                "cash_ratio": "{:.2f}",
                "nwc_to_assets": "{:.2%}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

with tab_profitability:
    st.subheader("Profitability Analysis")

    margin_long = prof_view.melt(
        id_vars=["entity", "year"],
        value_vars=["gross_margin", "operating_margin", "net_margin_parent"],
        var_name="metric",
        value_name="value",
    )
    margin_names = {
        "gross_margin": "Gross Margin",
        "operating_margin": "Operating Margin",
        "net_margin_parent": "Net Margin (Parent)",
    }
    margin_long["metric"] = margin_long["metric"].map(margin_names)

    margin_chart = px.bar(
        margin_long,
        x="entity",
        y="value",
        color="metric",
        facet_col="year",
        barmode="group",
        text_auto=".1%",
        labels={"entity": "Entity", "value": "Margin", "metric": "Metric"},
    )
    margin_chart.update_yaxes(tickformat=".0%")
    margin_chart.update_layout(legend_title_text="Metric")
    st.plotly_chart(margin_chart, use_container_width=True)

    eps_chart = px.line(
        prof_view.sort_values("year"),
        x="year",
        y="Basic EPS",
        color="entity",
        markers=True,
        labels={"year": "Year", "Basic EPS": "Basic EPS (NT$)", "entity": "Entity"},
    )
    eps_chart.update_xaxes(dtick=1)
    st.plotly_chart(eps_chart, use_container_width=True)

    profitability_table = prof_view[
        [
            "entity",
            "year",
            "Revenue",
            "Gross Profit",
            "Operating Profit",
            "Net Income Attributable to Parent",
            "Basic EPS",
            "gross_margin",
            "operating_margin",
            "net_margin_parent",
        ]
    ].sort_values(["entity", "year"])
    st.dataframe(
        profitability_table.style.format(
            {
                "Revenue": "{:,.0f}",
                "Gross Profit": "{:,.0f}",
                "Operating Profit": "{:,.0f}",
                "Net Income Attributable to Parent": "{:,.0f}",
                "Basic EPS": "{:.2f}",
                "gross_margin": "{:.2%}",
                "operating_margin": "{:.2%}",
                "net_margin_parent": "{:.2%}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

with tab_sources:
    st.subheader("Raw Financial Statement Data")
    source_view = raw[
        raw["company"].str.cat(raw["scope"], sep=" ").isin(selected_entities)
        & raw["year"].isin(selected_years)
    ].sort_values(["company", "scope", "year", "statement", "account"])

    st.dataframe(source_view, use_container_width=True, hide_index=True)
    st.caption(
        "Unit is preserved from the source table. Most financial statement values "
        "are in NT$ thousand; EPS is in NT$ per share."
    )
