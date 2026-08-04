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

    .prior-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 5px solid #3B82F6;
        border-radius: 14px;
        padding: 20px 20px 18px 20px;
        min-height: 360px;
        margin-bottom: 18px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }

    .prior-number {
        font-size: 0.76rem;
        font-weight: 800;
        letter-spacing: 0.08em;
        color: #2563EB;
        margin-bottom: 8px;
    }

    .prior-question {
        font-size: 1.04rem;
        font-weight: 750;
        line-height: 1.5;
        color: #0F172A;
        margin-bottom: 14px;
    }

    .prior-label {
        font-size: 0.72rem;
        font-weight: 800;
        letter-spacing: 0.06em;
        color: #64748B;
        margin-top: 12px;
        margin-bottom: 4px;
    }

    .prior-text {
        font-size: 0.86rem;
        line-height: 1.6;
        color: #475569;
    }

    .prior-status {
        display: inline-block;
        border-radius: 999px;
        padding: 5px 10px;
        font-size: 0.72rem;
        font-weight: 800;
        letter-spacing: 0.04em;
        margin-top: 14px;
    }

    .prior-confirmed {
        background-color: #DCFCE7;
        color: #166534;
    }

    .prior-mixed {
        background-color: #FEF3C7;
        color: #92400E;
    }

    .prior-gap {
        background-color: #EDE9FE;
        color: #5B21B6;
    }

    .prior-note {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 14px 16px;
        color: #475569;
        font-size: 0.86rem;
        line-height: 1.6;
        margin-bottom: 18px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

@st.cache_data
def load_raw_data(file_signature: tuple[int, int]) -> pd.DataFrame:
    """Load only the raw financial statement table.

    ``file_signature`` contains the Excel file's modification time and size.
    It is used as a cache key so Streamlit reloads the data whenever the Excel
    file is replaced, even when this function's code has not changed.
    """
    del file_signature
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


def calculate_cash_flow_quality(base: pd.DataFrame) -> pd.DataFrame:
    """Calculate cash-conversion and free-cash-flow metrics.

    Capital Expenditure is stored as a positive cash-outflow amount in the
    source data, so free cash flow is Operating Cash Flow minus Capex.
    """
    required_accounts = [
        "Operating Cash Flow",
        "Capital Expenditure",
        "Net Income",
        "Revenue",
        "Inventory",
        "Accounts Receivable",
    ]
    result = base.copy()
    for account in required_accounts:
        if account not in result.columns:
            result[account] = 0.0

    result["free_cash_flow"] = (
        result["Operating Cash Flow"] - result["Capital Expenditure"]
    )
    result["cash_conversion_ratio"] = safe_divide(
        result["Operating Cash Flow"], result["Net Income"]
    )
    result["free_cash_flow_margin"] = safe_divide(
        result["free_cash_flow"], result["Revenue"]
    )

    # Only keep entity-years for which cash-flow data were provided.
    available = (
        result["Operating Cash Flow"].ne(0)
        | result["Capital Expenditure"].ne(0)
        | result["Net Income"].ne(0)
    )
    return result[available].copy()


def format_change(value: float, suffix: str = "") -> str:
    if pd.isna(value):
        return "N/A"
    sign = "+" if value > 0 else ""
    return f"{sign}{value:.2f}{suffix}"



def format_growth_delta(current: float, previous: float) -> str | None:
    """Format a year-over-year percentage change for st.metric."""
    change = percent_change(current, previous)
    if pd.isna(change):
        return None
    return f"{change:+.1%} YoY"


def format_pp_delta(current: float, previous: float) -> str | None:
    """Format a margin change in percentage points."""
    if pd.isna(current) or pd.isna(previous):
        return None
    return f"{(current - previous) * 100:+.1f} pp"


def format_ratio_delta(current: float, previous: float) -> str | None:
    """Format a ratio change in x units."""
    if pd.isna(current) or pd.isna(previous):
        return None
    return f"{current - previous:+.2f}x"


def build_validation_report(
    raw_data: pd.DataFrame,
) -> tuple[dict, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Validate uniqueness, completeness, module coverage, and traceability."""

    data = raw_data.copy()
    data["entity"] = data["company"] + " " + data["scope"]

    key_columns = ["company", "scope", "year", "account"]
    duplicate_mask = data.duplicated(key_columns, keep=False)
    duplicate_rows = data.loc[duplicate_mask].sort_values(key_columns)

    required_fields = [
        "company",
        "scope",
        "year",
        "statement",
        "account",
        "value",
        "unit",
        "source_page",
        "source_file",
    ]
    missing_field_mask = pd.Series(False, index=data.index)
    for column in required_fields:
        if column in ["year", "value"]:
            missing_field_mask |= data[column].isna()
        else:
            normalized = data[column].astype(str).str.strip().str.lower()
            missing_field_mask |= normalized.isin(["", "nan", "none"])

    trace_page = (
        data["source_page"].astype(str).str.strip().str.lower()
        .map(lambda value: value not in {"", "nan", "none"})
    )
    trace_file = (
        data["source_file"].astype(str).str.strip().str.lower()
        .map(lambda value: value not in {"", "nan", "none"})
    )
    traceable_mask = trace_page & trace_file

    module_requirements = {
        "Liquidity": {
            "Current Assets",
            "Current Liabilities",
            "Cash and Cash Equivalents",
            "Total Assets",
            "Inventory",
            "Other Current Assets",
        },
        "Profitability": {
            "Revenue",
            "Gross Profit",
            "Operating Profit",
            "Net Income Attributable to Parent",
            "Basic EPS",
        },
    }
    cash_flow_requirements = {
        "Operating Cash Flow",
        "Capital Expenditure",
        "Net Income",
        "Accounts Receivable",
    }

    coverage_rows = []
    for (entity, year), group in data.groupby(["entity", "year"], sort=True):
        accounts = set(group["account"])

        row = {"entity": entity, "year": int(year)}
        for module, required_accounts in module_requirements.items():
            missing_accounts = sorted(required_accounts.difference(accounts))
            row[module] = (
                "Passed"
                if not missing_accounts
                else "Missing: " + ", ".join(missing_accounts)
            )

        if entity == "AVC Consolidated":
            missing_cash = sorted(cash_flow_requirements.difference(accounts))
            row["Cash Flow Quality"] = (
                "Passed"
                if not missing_cash
                else "Missing: " + ", ".join(missing_cash)
            )
        else:
            row["Cash Flow Quality"] = "Not applicable"

        coverage_rows.append(row)

    coverage = pd.DataFrame(coverage_rows)
    coverage_failures = coverage[
        coverage[["Liquidity", "Profitability", "Cash Flow Quality"]]
        .apply(
            lambda row: any(
                str(value).startswith("Missing:")
                for value in row
            ),
            axis=1,
        )
    ]

    source_rows = []
    for (entity, source_file), group in data.groupby(
        ["entity", "source_file"],
        sort=True,
    ):
        pages = sorted(
            set(group["source_page"].astype(str)),
            key=lambda value: (
                int("".join(character for character in value if character.isdigit()) or 0),
                value,
            ),
        )
        years = sorted(group["year"].unique().tolist())
        source_rows.append(
            {
                "entity": entity,
                "source_file": source_file,
                "years": ", ".join(str(year) for year in years),
                "pages": ", ".join(pages),
                "records": len(group),
            }
        )
    source_coverage = pd.DataFrame(source_rows)

    summary = {
        "total_records": len(data),
        "duplicate_records": int(duplicate_mask.sum()),
        "missing_required_values": int(missing_field_mask.sum()),
        "traceable_records": int(traceable_mask.sum()),
        "traceability_rate": float(traceable_mask.mean()) if len(data) else 0.0,
        "coverage_failures": len(coverage_failures),
    }
    summary["passed"] = (
        summary["duplicate_records"] == 0
        and summary["missing_required_values"] == 0
        and summary["coverage_failures"] == 0
        and summary["traceability_rate"] == 1.0
    )

    return summary, coverage, duplicate_rows, source_coverage


def build_source_reference(
    raw_data: pd.DataFrame,
    entity_accounts: dict[str, list[str]],
) -> str:
    """Return a compact source-file and page reference for a finding."""

    data = raw_data.copy()
    data["entity"] = data["company"] + " " + data["scope"]
    references = []

    for entity, accounts in entity_accounts.items():
        rows = data[
            data["entity"].eq(entity)
            & data["account"].isin(accounts)
        ]
        for source_file, group in rows.groupby("source_file", sort=True):
            pages = sorted(
                set(group["source_page"].astype(str)),
                key=lambda value: (
                    int(
                        "".join(
                            character
                            for character in value
                            if character.isdigit()
                        )
                        or 0
                    ),
                    value,
                ),
            )
            references.append(
                f"{source_file} ({', '.join(pages)})"
            )

    if not references:
        return "Source traceability unavailable."
    return "Source: " + "; ".join(dict.fromkeys(references))


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
    cash_flow_data: pd.DataFrame,
) -> dict:
    """Build evidence-based answers to the dashboard's research questions.

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
            "rq4": {
                "status": "DATA CHECK",
                "title": "Required cash-flow observations are unavailable.",
                "evidence": "Confirm that AVC consolidated cash-flow records are included for both years.",
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

    # RQ4: Did rapid growth convert into stable operating cash flow?
    avc_cash = (
        cash_flow_data[
            cash_flow_data["entity"].eq("AVC Consolidated")
        ]
        .sort_values("year")
        .copy()
    )

    if len(avc_cash) < 2:
        rq4_status = "DATA CHECK"
        rq4_title = "Two years of AVC consolidated cash-flow data are required."
        rq4_evidence = (
            "Operating cash flow, capital expenditure, and net income must be "
            "available for FY 2024 and FY 2025."
        )
        rq4_interpretation = (
            "The cash-conversion question cannot yet be evaluated reliably."
        )
    else:
        cash_first = avc_cash.iloc[0]
        cash_last = avc_cash.iloc[-1]

        net_income_growth = percent_change(
            cash_last["Net Income"], cash_first["Net Income"]
        )
        ocf_growth = percent_change(
            cash_last["Operating Cash Flow"],
            cash_first["Operating Cash Flow"],
        )
        inventory_growth = percent_change(
            cash_last["Inventory"], cash_first["Inventory"]
        )
        receivables_growth = percent_change(
            cash_last["Accounts Receivable"],
            cash_first["Accounts Receivable"],
        )
        revenue_cash_growth = percent_change(
            cash_last["Revenue"], cash_first["Revenue"]
        )
        conversion_change = (
            cash_last["cash_conversion_ratio"]
            - cash_first["cash_conversion_ratio"]
        )
        fcf_margin_change = (
            cash_last["free_cash_flow_margin"]
            - cash_first["free_cash_flow_margin"]
        )

        cash_conversion_strengthened = (
            cash_last["Operating Cash Flow"] > 0
            and cash_last["free_cash_flow"] > 0
            and ocf_growth > net_income_growth
            and conversion_change > 0
            and fcf_margin_change > 0
            and inventory_growth < revenue_cash_growth
            and receivables_growth < revenue_cash_growth
        )

        rq4_status = (
            "ANSWER: YES"
            if cash_conversion_strengthened
            else "ANSWER: MIXED"
        )
        rq4_title = (
            "Growth converted into substantially stronger operating and free cash flow."
            if cash_conversion_strengthened
            else "Cash generation improved, but the evidence is not uniformly strong."
        )
        rq4_evidence = (
            f"Operating cash flow increased {ocf_growth:.1%}, compared with "
            f"net income growth of {net_income_growth:.1%}. Cash conversion "
            f"improved from {cash_first['cash_conversion_ratio']:.2f}x to "
            f"{cash_last['cash_conversion_ratio']:.2f}x. Free cash flow rose "
            f"from NT${cash_first['free_cash_flow'] / 1_000_000:,.1f} bn to "
            f"NT${cash_last['free_cash_flow'] / 1_000_000:,.1f} bn, while its "
            f"margin increased {fcf_margin_change * 100:+.1f} pp. Inventory "
            f"grew {inventory_growth:.1%} and receivables grew "
            f"{receivables_growth:.1%}, both below revenue growth of "
            f"{revenue_cash_growth:.1%}."
        )
        rq4_interpretation = (
            "The 2025 results indicate that accounting earnings were converted "
            "into cash more effectively, even after expansion-related capital "
            "spending. The conclusion remains descriptive because it covers only "
            "two annual observations, and capex is approximated by cash paid to "
            "acquire property, plant and equipment."
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
        "rq4": {
            "status": rq4_status,
            "title": rq4_title,
            "evidence": rq4_evidence,
            "interpretation": rq4_interpretation,
        },
    }


try:
    file_stat = DATA_FILE.stat()
    data_file_signature = (file_stat.st_mtime_ns, file_stat.st_size)
    raw = load_raw_data(data_file_signature)
except (FileNotFoundError, ValueError) as exc:
    st.error(str(exc))
    st.stop()

base = build_metric_base(raw)
liquidity = calculate_liquidity(base)
profitability = calculate_profitability(base)
cash_flow_quality = calculate_cash_flow_quality(base)

research_answers = build_research_answers(
    profitability,
    liquidity,
    cash_flow_quality,
)


(
    validation_summary,
    validation_coverage,
    validation_duplicates,
    source_coverage,
) = build_validation_report(raw)

rq1_source = build_source_reference(
    raw,
    {
        "AVC Consolidated": [
            "Revenue",
            "Gross Profit",
            "Operating Profit",
            "Net Income Attributable to Parent",
            "Basic EPS",
        ]
    },
)
rq2_source = build_source_reference(
    raw,
    {
        "AVC Consolidated": [
            "Current Assets",
            "Current Liabilities",
            "Cash and Cash Equivalents",
            "Inventory",
            "Prepayments",
            "Other Current Assets",
            "Total Assets",
        ]
    },
)
rq3_source = build_source_reference(
    raw,
    {
        "AVC Consolidated": ["Revenue", "Operating Profit", "Gross Profit"],
        "AVC Standalone": ["Revenue", "Operating Profit", "Gross Profit"],
        "Auras Consolidated": ["Revenue", "Operating Profit", "Gross Profit"],
    },
)
rq4_source = build_source_reference(
    raw,
    {
        "AVC Consolidated": [
            "Operating Cash Flow",
            "Capital Expenditure",
            "Net Income",
            "Revenue",
            "Inventory",
            "Accounts Receivable",
        ]
    },
)

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
cash_view = cash_flow_quality[
    cash_flow_quality["entity"].isin(selected_entities)
    & cash_flow_quality["year"].isin(selected_years)
].copy()

st.title("AVC vs. Auras Financial Decision Dashboard")
st.caption(
    "Standalone vs. consolidated analysis and peer benchmarking, FY 2024–2025. "
    "All ratios are recalculated from the raw financial statement data in Python."
)


def get_prior_status(question_key: str) -> tuple[str, str]:
    """Translate a calculated answer into a prior-question status label."""

    calculated_status = research_answers[question_key]["status"]

    status_map = {
        "rq1": {
            "ANSWER: YES": ("CONFIRMED & STRENGTHENED", "prior-confirmed"),
            "ANSWER: MIXED": ("PARTIALLY CONFIRMED", "prior-mixed"),
            "DATA CHECK": ("DATA CHECK REQUIRED", "prior-gap"),
        },
        "rq2": {
            "ANSWER: NO": ("PARTIALLY EASED", "prior-confirmed"),
            "ANSWER: YES": ("CONCERN PERSISTED", "prior-mixed"),
            "ANSWER: MIXED": ("MIXED EVIDENCE", "prior-mixed"),
            "DATA CHECK": ("DATA CHECK REQUIRED", "prior-gap"),
        },
        "rq3": {
            "ANSWER: SCALE ADVANTAGE": (
                "CONFIRMED AT THE SCALE LEVEL",
                "prior-confirmed",
            ),
            "DATA CHECK": ("DATA CHECK REQUIRED", "prior-gap"),
        },
        "rq4": {
            "ANSWER: YES": ("CONFIRMED & STRENGTHENED", "prior-confirmed"),
            "ANSWER: MIXED": ("MIXED EVIDENCE", "prior-mixed"),
            "DATA CHECK": ("DATA CHECK REQUIRED", "prior-gap"),
        },
    }

    return status_map.get(question_key, {}).get(
        calculated_status,
        ("REVIEW REQUIRED", "prior-gap"),
    )


rq1_prior_status, rq1_prior_class = get_prior_status("rq1")
rq2_prior_status, rq2_prior_class = get_prior_status("rq2")
rq3_prior_status, rq3_prior_class = get_prior_status("rq3")
rq4_prior_status, rq4_prior_class = get_prior_status("rq4")


def render_prior_question(
    number: str,
    question: str,
    prior_finding: str,
    current_evidence: str,
    status: str,
    status_class: str,
    updated_interpretation: str,
    source_note: str,
    final_label: str = "UPDATED INTERPRETATION",
) -> None:
    """Render one prior-question card using native Streamlit components."""

    with st.container(border=True):
        st.markdown(f"**{number}**")
        st.markdown(f"#### {question}")

        st.markdown("##### 2023–2024 FINDING")
        st.write(prior_finding)

        st.markdown("##### 2024–2025 EVIDENCE")
        st.write(current_evidence)

        st.markdown(
            f'<span class="prior-status {status_class}">{status}</span>',
            unsafe_allow_html=True,
        )

        st.markdown(f"##### {final_label}")
        st.write(updated_interpretation)
        st.caption(source_note)


(
    tab_research,
    tab_summary,
    tab_liquidity,
    tab_profitability,
    tab_cash_flow,
    tab_memo,
    tab_methodology,
    tab_sources,
) = st.tabs(
    [
        "Research Overview",
        "Executive Summary",
        "Liquidity",
        "Profitability",
        "Cash Flow Quality",
        "Decision Memo",
        "Methodology & Validation",
        "Source Data",
    ]
)


with tab_research:
    research_overview_html = (
        '<div class="research-shell">'
        '<div class="research-kicker">RESEARCH FRAMEWORK</div>'
        '<div class="research-title">Research Motivation</div>'
        '<div class="research-objective">'
        'This project originated from a classroom financial statement analysis '
        'comparing AVC and Auras for FY 2023–2024. To extend the original '
        'assignment into a more complete and reproducible research project, I '
        'updated the dataset to FY 2024–2025, rebuilt the financial analysis in '
        'Python, and developed an interactive dashboard. This extension allows '
        'the earlier findings to be reassessed using the latest annual data while '
        'incorporating profitability, liquidity, standalone-versus-consolidated '
        'differences, and peer benchmarking.'
        '</div>'
        '<div style="height:18px;"></div>'
        '<div class="research-title">Project Objective</div>'
        '<div class="research-objective">'
        "This project examines whether AVC's rapid FY 2025 growth translated "
        'into stronger profitability, resilient short-term liquidity, and '
        'improved cash generation. It combines year-over-year analysis, '
        'standalone-versus-consolidated comparison, peer benchmarking against '
        'Auras, and cash-flow-quality analysis to identify decision-relevant '
        'financial signals.'
        '</div>'
        '<div class="research-meta">'
        '<span class="research-chip">Period: FY 2024–2025</span>'
        '<span class="research-chip">Primary firm: AVC</span>'
        '<span class="research-chip">Peer benchmark: Auras</span>'
        '<span class="research-chip">Source: Audited financial statements</span>'
        '<span class="research-chip">Method: Python-recalculated ratios</span>'
        '</div>'
        '<div class="design-note">'
        'Research design: descriptive and comparative analysis. Reported '
        'relationships should be interpreted as financial associations rather '
        'than causal effects.'
        '</div>'
        '</div>'
    )
    st.markdown(research_overview_html, unsafe_allow_html=True)
        st.markdown("### Research Snapshot")

    question_col, data_col, findings_col = st.columns(3, gap="large")

    with question_col:
        with st.container(border=True):
            st.markdown("#### Research Question")
            st.write(
                "Did AVC's rapid FY 2025 growth translate into stronger "
                "profitability, resilient short-term liquidity, and improved "
                "cash generation, and how did its performance compare with Auras?"
            )

    with data_col:
        with st.container(border=True):
            st.markdown("#### Data")
            st.markdown(
                f"""
                - Period: **FY 2024–2025**
                - AVC consolidated statements
                - AVC standalone statements
                - Auras consolidated statements
                - **{validation_summary['total_records']}** source-traceable records
                - Audited financial statements
                """
            )

    with findings_col:
        with st.container(border=True):
            st.markdown("#### Key Findings")
            st.markdown(
                f"""
                - **Profitability:** {research_answers['rq1']['title']}
                - **Liquidity:** {research_answers['rq2']['title']}
                - **Cash Flow:** {research_answers['rq4']['title']}
                """
            )

    st.divider()

    st.markdown("### Prior Questions Revisited: FY 2024–2025 Update")
    st.info(
        "The FY 2023–2024 classroom report did not present formal research "
        "questions. The questions below reconstruct its central analytical "
        "issues from the prior conclusions and reassess them using FY 2024–2025 "
        "evidence. They should therefore be treated as inferred questions rather "
        "than verbatim questions from the original report."
    )

    prior_row1_col1, prior_row1_col2 = st.columns(2, gap="large")

    with prior_row1_col1:
        render_prior_question(
            number="PRIOR QUESTION 1",
            question="Would AVC's profitability improvement continue beyond FY 2024?",
            prior_finding=(
                "Revenue, margins, and EPS improved, but the prior report noted "
                "that long-term earnings stability still required monitoring."
            ),
            current_evidence=research_answers["rq1"]["evidence"],
            status=rq1_prior_status,
            status_class=rq1_prior_class,
            updated_interpretation=research_answers["rq1"]["interpretation"],
            source_note=rq1_source,
        )

    with prior_row1_col2:
        render_prior_question(
            number="PRIOR QUESTION 2",
            question="Would rapid expansion continue to weaken short-term liquidity?",
            prior_finding=(
                "Current, quick, cash, and net-working-capital ratios declined, "
                "indicating tighter liquidity and higher working-capital pressure."
            ),
            current_evidence=research_answers["rq2"]["evidence"],
            status=rq2_prior_status,
            status_class=rq2_prior_class,
            updated_interpretation=(
                research_answers["rq2"]["interpretation"]
                + " Cash-flow conversion improved substantially, reducing concerns "
                "about immediate liquidity pressure. However, inventory turnover "
                "and receivables turnover are still required to determine whether "
                "working-capital efficiency also improved."
            ),
            source_note=rq2_source,
        )

    prior_row2_col1, prior_row2_col2 = st.columns(2, gap="large")

    with prior_row2_col1:
        render_prior_question(
            number="PRIOR QUESTION 3",
            question=(
                "Would standalone–consolidated differences remain economically meaningful?"
            ),
            prior_finding=(
                "The group showed greater operating scale, while consolidated "
                "efficiency and financing characteristics differed from the "
                "standalone entity."
            ),
            current_evidence=research_answers["rq3"]["evidence"],
            status=rq3_prior_status,
            status_class=rq3_prior_class,
            updated_interpretation=(
                research_answers["rq3"]["interpretation"]
                + " The current evidence confirms material group-level scale, "
                "but does not isolate subsidiary operating efficiency."
            ),
            source_note=rq3_source,
        )

    with prior_row2_col2:
        render_prior_question(
            number="PRIOR QUESTION 4",
            question="Did rapid growth convert into stable operating cash flow?",
            prior_finding=(
                "The prior report identified cash conversion—not accounting "
                "profitability—as the central sustainability risk, particularly "
                "as inventory, receivables, and expansion spending increased."
            ),
            current_evidence=research_answers["rq4"]["evidence"],
            status=rq4_prior_status,
            status_class=rq4_prior_class,
            updated_interpretation=research_answers["rq4"]["interpretation"],
            source_note=rq4_source,
        )

    st.caption(
        "Interpretive note: 'Confirmed' means the new evidence is directionally "
        "consistent with the prior finding; it does not imply causal proof."
    )

    with st.expander("Limitations and interpretation boundaries"):
        st.markdown(
            """
            - The analysis covers only FY 2024–2025, so the results should not be
              treated as a long-term trend.
            - Auras is the primary peer benchmark; one peer cannot represent the
              entire thermal-management industry.
            - The consolidated-minus-standalone gap is not a pure subsidiary
              contribution because consolidation includes intercompany eliminations
              and accounting adjustments.
            - The findings are descriptive financial associations and do not establish
              causal effects.
            - Capital expenditure is approximated by cash paid to acquire property,
              plant and equipment.
            """
        )


with tab_summary:
    latest_year = max(selected_years)
    latest = prof_view[prof_view["year"].eq(latest_year)].copy()

    st.subheader(f"{latest_year} Snapshot")

    if latest.empty:
        st.info("No profitability records match the current filters.")
    else:
        selector_col, spacer_col = st.columns([1, 3])

        with selector_col:
            display_entity = st.selectbox(
                "Entity",
                options=sorted(latest["entity"].unique()),
                key="summary_entity",
            )

        kpi = latest[latest["entity"].eq(display_entity)].iloc[0]

        prof_history = (
            profitability[profitability["entity"].eq(display_entity)]
            .sort_values("year")
            .copy()
        )
        previous_prof = prof_history[prof_history["year"].lt(latest_year)]
        previous_kpi = previous_prof.iloc[-1] if not previous_prof.empty else None

        liq_kpi = liquidity[
            liquidity["entity"].eq(display_entity)
            & liquidity["year"].eq(latest_year)
        ]
        liq_history = (
            liquidity[liquidity["entity"].eq(display_entity)]
            .sort_values("year")
            .copy()
        )
        previous_liq_data = liq_history[liq_history["year"].lt(latest_year)]
        previous_liq = (
            previous_liq_data.iloc[-1]
            if not previous_liq_data.empty
            else None
        )

        st.markdown("#### Performance")
        row1 = st.columns(4, gap="large")

        row1[0].metric(
            "Revenue",
            f"NT${kpi['Revenue'] / 1_000_000:,.1f} bn",
            delta=(
                format_growth_delta(kpi["Revenue"], previous_kpi["Revenue"])
                if previous_kpi is not None
                else None
            ),
        )
        row1[1].metric(
            "Gross Margin",
            f"{kpi['gross_margin']:.2%}",
            delta=(
                format_pp_delta(
                    kpi["gross_margin"],
                    previous_kpi["gross_margin"],
                )
                if previous_kpi is not None
                else None
            ),
        )
        row1[2].metric(
            "Operating Margin",
            f"{kpi['operating_margin']:.2%}",
            delta=(
                format_pp_delta(
                    kpi["operating_margin"],
                    previous_kpi["operating_margin"],
                )
                if previous_kpi is not None
                else None
            ),
        )
        row1[3].metric(
            "Net Margin (Parent)",
            f"{kpi['net_margin_parent']:.2%}",
            delta=(
                format_pp_delta(
                    kpi["net_margin_parent"],
                    previous_kpi["net_margin_parent"],
                )
                if previous_kpi is not None
                else None
            ),
        )

        st.markdown("#### Per-share & Liquidity")
        row2 = st.columns(4, gap="large")

        row2[0].metric(
            "Basic EPS",
            f"NT${kpi['Basic EPS']:.2f}",
            delta=(
                format_growth_delta(
                    kpi["Basic EPS"],
                    previous_kpi["Basic EPS"],
                )
                if previous_kpi is not None
                else None
            ),
        )

        if not liq_kpi.empty:
            liq_row = liq_kpi.iloc[0]

            row2[1].metric(
                "Current Ratio",
                f"{liq_row['current_ratio']:.2f}",
                delta=(
                    format_ratio_delta(
                        liq_row["current_ratio"],
                        previous_liq["current_ratio"],
                    )
                    if previous_liq is not None
                    else None
                ),
            )
            row2[2].metric(
                "Quick Ratio",
                f"{liq_row['quick_ratio']:.2f}",
                delta=(
                    format_ratio_delta(
                        liq_row["quick_ratio"],
                        previous_liq["quick_ratio"],
                    )
                    if previous_liq is not None
                    else None
                ),
            )
            row2[3].metric(
                "Cash Ratio",
                f"{liq_row['cash_ratio']:.2f}",
                delta=(
                    format_ratio_delta(
                        liq_row["cash_ratio"],
                        previous_liq["cash_ratio"],
                    )
                    if previous_liq is not None
                    else None
                ),
            )
        else:
            row2[1].metric("Current Ratio", "N/A")
            row2[2].metric("Quick Ratio", "N/A")
            row2[3].metric("Cash Ratio", "N/A")

        st.subheader("Decision Memo")
        memo_profit, memo_liquidity, memo_cash = st.columns(3, gap="large")

        with memo_profit:
            st.success(
                "**Profitability**\n\n"
                + research_answers["rq1"]["title"]
            )
            st.caption(rq1_source)

        with memo_liquidity:
            st.info(
                "**Liquidity**\n\n"
                + research_answers["rq2"]["title"]
            )
            st.caption(rq2_source)

        with memo_cash:
            if research_answers["rq4"]["status"] == "ANSWER: YES":
                st.success(
                    "**Cash Flow**\n\n"
                    + research_answers["rq4"]["title"]
                )
            else:
                st.warning(
                    "**Cash Flow**\n\n"
                    + research_answers["rq4"]["title"]
                )
            st.caption(rq4_source)

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

with tab_cash_flow:
    st.subheader("Cash Flow Quality")
    st.caption(
        "This module evaluates whether AVC's accounting earnings converted into "
        "operating cash and free cash flow. Capital expenditure is approximated "
        "by cash paid to acquire property, plant and equipment."
    )

    if cash_view.empty:
        st.info(
            "No cash-flow records match the current filters. Select AVC Consolidated "
            "and at least one available year."
        )
    else:
        cash_entities = sorted(cash_view["entity"].unique())
        cash_entity = st.selectbox(
            "Entity",
            options=cash_entities,
            key="cash_flow_entity",
        )

        cash_entity_data = (
            cash_view[cash_view["entity"].eq(cash_entity)]
            .sort_values("year")
            .copy()
        )
        latest_cash = cash_entity_data.iloc[-1]
        previous_cash = (
            cash_entity_data.iloc[-2]
            if len(cash_entity_data) >= 2
            else None
        )

        cash_kpis = st.columns(4, gap="large")
        cash_kpis[0].metric(
            "Operating Cash Flow",
            f"NT${latest_cash['Operating Cash Flow'] / 1_000_000:,.1f} bn",
            delta=(
                format_growth_delta(
                    latest_cash["Operating Cash Flow"],
                    previous_cash["Operating Cash Flow"],
                )
                if previous_cash is not None
                else None
            ),
        )
        cash_kpis[1].metric(
            "Free Cash Flow",
            f"NT${latest_cash['free_cash_flow'] / 1_000_000:,.1f} bn",
            delta=(
                format_growth_delta(
                    latest_cash["free_cash_flow"],
                    previous_cash["free_cash_flow"],
                )
                if previous_cash is not None
                else None
            ),
        )
        cash_kpis[2].metric(
            "Cash Conversion Ratio",
            f"{latest_cash['cash_conversion_ratio']:.2f}x",
            delta=(
                format_ratio_delta(
                    latest_cash["cash_conversion_ratio"],
                    previous_cash["cash_conversion_ratio"],
                )
                if previous_cash is not None
                else None
            ),
        )
        cash_kpis[3].metric(
            "Free Cash Flow Margin",
            f"{latest_cash['free_cash_flow_margin']:.1%}",
            delta=(
                format_pp_delta(
                    latest_cash["free_cash_flow_margin"],
                    previous_cash["free_cash_flow_margin"],
                )
                if previous_cash is not None
                else None
            ),
        )
        st.caption(rq4_source)

        flow_long = cash_entity_data[
            [
                "year",
                "Net Income",
                "Operating Cash Flow",
                "Capital Expenditure",
                "free_cash_flow",
            ]
        ].rename(
            columns={
                "Net Income": "Net Income",
                "Operating Cash Flow": "Operating Cash Flow",
                "Capital Expenditure": "Capital Expenditure",
                "free_cash_flow": "Free Cash Flow",
            }
        ).melt(
            id_vars="year",
            var_name="metric",
            value_name="value",
        )
        flow_long["value_bn"] = flow_long["value"] / 1_000_000

        flow_chart = px.bar(
            flow_long,
            x="year",
            y="value_bn",
            color="metric",
            barmode="group",
            text_auto=".1f",
            labels={
                "year": "Year",
                "value_bn": "NT$ billion",
                "metric": "Metric",
            },
            title="Earnings, Operating Cash Flow, Capex, and Free Cash Flow",
        )
        flow_chart.update_xaxes(dtick=1)
        flow_chart.update_layout(legend_title_text="Metric")
        st.plotly_chart(flow_chart, use_container_width=True)

        if len(cash_entity_data) >= 2:
            first_cash = cash_entity_data.iloc[0]
            last_cash = cash_entity_data.iloc[-1]

            growth_data = pd.DataFrame(
                {
                    "metric": [
                        "Revenue",
                        "Net Income",
                        "Operating Cash Flow",
                        "Inventory",
                        "Accounts Receivable",
                    ],
                    "growth": [
                        percent_change(last_cash["Revenue"], first_cash["Revenue"]),
                        percent_change(
                            last_cash["Net Income"], first_cash["Net Income"]
                        ),
                        percent_change(
                            last_cash["Operating Cash Flow"],
                            first_cash["Operating Cash Flow"],
                        ),
                        percent_change(
                            last_cash["Inventory"], first_cash["Inventory"]
                        ),
                        percent_change(
                            last_cash["Accounts Receivable"],
                            first_cash["Accounts Receivable"],
                        ),
                    ],
                }
            )

            growth_chart = px.bar(
                growth_data,
                x="metric",
                y="growth",
                text_auto=".1%",
                labels={"metric": "Metric", "growth": "FY 2024–2025 Growth"},
                title="Growth Quality: Operating Scale vs. Working Capital",
            )
            growth_chart.update_yaxes(tickformat=".0%")
            st.plotly_chart(growth_chart, use_container_width=True)

            st.success(
                f"{research_answers['rq4']['status']}: "
                f"{research_answers['rq4']['title']}"
            )
            st.write(research_answers["rq4"]["interpretation"])

        cash_table = cash_entity_data[
            [
                "entity",
                "year",
                "Revenue",
                "Net Income",
                "Operating Cash Flow",
                "Capital Expenditure",
                "free_cash_flow",
                "cash_conversion_ratio",
                "free_cash_flow_margin",
                "Inventory",
                "Accounts Receivable",
            ]
        ].copy()

        st.dataframe(
            cash_table.style.format(
                {
                    "Revenue": "{:,.0f}",
                    "Net Income": "{:,.0f}",
                    "Operating Cash Flow": "{:,.0f}",
                    "Capital Expenditure": "{:,.0f}",
                    "free_cash_flow": "{:,.0f}",
                    "cash_conversion_ratio": "{:.2f}x",
                    "free_cash_flow_margin": "{:.2%}",
                    "Inventory": "{:,.0f}",
                    "Accounts Receivable": "{:,.0f}",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.caption(
            "Financial statement amounts are shown in NT$ thousand in the table. "
            "Positive capital expenditure values represent cash outflows."
        )


with tab_methodology:
    st.subheader("Methodology & Data Validation")
    st.caption(
        "This page documents how the financial statements were converted into "
        "a reproducible analytical dataset and verifies whether the required "
        "records are complete, unique, and traceable."
    )

    st.markdown("### Reproducible Data Pipeline")
    pipeline_columns = st.columns(4, gap="large")

    pipeline_steps = [
        (
            "1. Source Collection",
            "Audited standalone and consolidated financial statements.",
        ),
        (
            "2. Account Mapping",
            "Financial statement line items mapped into consistent account names.",
        ),
        (
            "3. Long-form Database",
            "Company, scope, year, statement, account, value, unit, page, and file.",
        ),
        (
            "4. Python Analysis",
            "Ratios, growth rates, cash-flow quality, charts, and research findings.",
        ),
    ]

    for column, (title, description) in zip(pipeline_columns, pipeline_steps):
        with column:
            with st.container(border=True):
                st.markdown(f"**{title}**")
                st.write(description)

    st.markdown("### Validation Summary")
    validation_metrics = st.columns(4, gap="large")
    validation_metrics[0].metric(
        "Records Loaded",
        f"{validation_summary['total_records']:,}",
    )
    validation_metrics[1].metric(
        "Duplicate Records",
        f"{validation_summary['duplicate_records']:,}",
        delta=(
            "Passed"
            if validation_summary["duplicate_records"] == 0
            else "Review required"
        ),
        delta_color=(
            "off"
            if validation_summary["duplicate_records"] == 0
            else "inverse"
        ),
    )
    validation_metrics[2].metric(
        "Missing Required Values",
        f"{validation_summary['missing_required_values']:,}",
        delta=(
            "Passed"
            if validation_summary["missing_required_values"] == 0
            else "Review required"
        ),
        delta_color=(
            "off"
            if validation_summary["missing_required_values"] == 0
            else "inverse"
        ),
    )
    validation_metrics[3].metric(
        "Source Traceability",
        f"{validation_summary['traceability_rate']:.0%}",
        delta=(
            "Passed"
            if validation_summary["traceability_rate"] == 1.0
            else "Incomplete"
        ),
        delta_color=(
            "off"
            if validation_summary["traceability_rate"] == 1.0
            else "inverse"
        ),
    )

    if validation_summary["passed"]:
        st.success(
            "Validation status: PASSED. Core module accounts are available, "
            "records are unique, required values are complete, and all records "
            "retain source-file and source-page references."
        )
    else:
        st.warning(
            "Validation status: REVIEW REQUIRED. Inspect the coverage and "
            "duplicate-record tables below."
        )

    st.markdown("### Entity–Year Module Coverage")
    st.dataframe(
        validation_coverage,
        use_container_width=True,
        hide_index=True,
    )

    if not validation_duplicates.empty:
        with st.expander("Duplicate records requiring review"):
            st.dataframe(
                validation_duplicates,
                use_container_width=True,
                hide_index=True,
            )

    st.markdown("### Source Coverage")
    st.dataframe(
        source_coverage,
        use_container_width=True,
        hide_index=True,
    )

    with st.expander("Core formulas and calculation definitions"):
        st.markdown(
            """
            **Liquidity**
            - Current Ratio = Current Assets / Current Liabilities
            - Quick Ratio = (Current Assets − Inventory − Prepayments − Other
              Current Assets) / Current Liabilities
            - Cash Ratio = Cash and Cash Equivalents / Current Liabilities
            - NWC / Assets = (Current Assets − Current Liabilities) / Total Assets

            **Profitability**
            - Gross Margin = Gross Profit / Revenue
            - Operating Margin = Operating Profit / Revenue
            - Net Margin (Parent) = Net Income Attributable to Parent / Revenue

            **Cash-flow quality**
            - Cash Conversion Ratio = Operating Cash Flow / Net Income
            - Free Cash Flow = Operating Cash Flow − Capital Expenditure
            - Free Cash Flow Margin = Free Cash Flow / Revenue
            """
        )

    st.info(
        "Reproducibility note: the app reads only the raw_financials worksheet "
        "and recalculates all analytical metrics in Python. It does not rely on "
        "Excel formula caches."
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
        "are in NT$ thousand; EPS is in NT$ per share. The source_page and "
        "source_file columns provide record-level traceability to the audited "
        "financial statements."
    )
