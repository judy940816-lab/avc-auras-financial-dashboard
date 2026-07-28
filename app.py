from pathlib import Path

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


try:
    raw = load_raw_data()
except (FileNotFoundError, ValueError) as exc:
    st.error(str(exc))
    st.stop()

base = build_metric_base(raw)
liquidity = calculate_liquidity(base)
profitability = calculate_profitability(base)

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
        display_entity = st.selectbox(
            "Entity for KPI cards",
            options=sorted(latest["entity"].unique()),
            key="summary_entity",
        )
        kpi = latest[latest["entity"].eq(display_entity)].iloc[0]
        liq_kpi = liq_view[
            liq_view["entity"].eq(display_entity)
            & liq_view["year"].eq(latest_year)
        ]

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Revenue", f"NT${kpi['Revenue'] / 1_000_000:,.1f} bn")
        c2.metric("Gross Margin", f"{kpi['gross_margin']:.2%}")
        c3.metric("Operating Margin", f"{kpi['operating_margin']:.2%}")
        c4.metric("Net Margin (Parent)", f"{kpi['net_margin_parent']:.2%}")
        c5.metric("Basic EPS", f"NT${kpi['Basic EPS']:.2f}")

        if not liq_kpi.empty:
            q1, q2, q3 = st.columns(3)
            liq_row = liq_kpi.iloc[0]
            q1.metric("Current Ratio", f"{liq_row['current_ratio']:.2f}")
            q2.metric("Quick Ratio", f"{liq_row['quick_ratio']:.2f}")
            q3.metric("Cash Ratio", f"{liq_row['cash_ratio']:.2f}")

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
        labels={"entity": "Entity", "Revenue": "Revenue (NT$ thousand)", "year": "Year"},
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
