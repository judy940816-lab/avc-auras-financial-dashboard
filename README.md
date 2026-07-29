# AVC vs. Auras Financial Decision Dashboard

**Live dashboard:** https://avc-auras-financial-dashboard.streamlit.app/

An interactive financial research project that extends a classroom comparison
of AVC and Auras from FY 2023–2024 to FY 2024–2025.

The project converts audited financial statement data into a reproducible
Python workflow and reassesses the earlier findings through profitability,
liquidity, standalone-versus-consolidated comparison, peer benchmarking, and
cash-flow-quality analysis.

## Research questions revisited

1. Did AVC's profitability improvement continue beyond FY 2024?
2. Did rapid expansion continue to weaken short-term liquidity?
3. Did standalone–consolidated differences remain economically meaningful?
4. Did rapid growth convert into stable operating cash flow?

## Key FY 2025 findings

- AVC consolidated revenue increased approximately **94.6%** year over year.
- Operating profit increased approximately **154.6%**, exceeding revenue
  growth and indicating positive operating leverage.
- Gross margin, operating margin, and net margin all improved.
- Immediate liquidity was broadly stable: the current ratio declined slightly,
  while the quick and cash ratios improved.
- Operating cash flow increased approximately **309.2%**.
- Cash conversion improved from approximately **1.05x to 1.88x**.
- Free cash flow increased from approximately **NT$5.1 billion to
  NT$32.1 billion**.

## Dashboard pages

- **Research Overview** — motivation, objective, prior questions, updated
  findings, source traceability, and limitations.
- **Executive Summary** — KPI cards with year-over-year changes and a
  decision memo.
- **Liquidity** — current, quick, cash, and net-working-capital ratios.
- **Profitability** — revenue, margins, and EPS comparison.
- **Cash Flow Quality** — operating cash flow, free cash flow, cash conversion,
  and growth-quality analysis.
- **Methodology & Validation** — data pipeline, uniqueness checks, account
  coverage, source traceability, and formulas.
- **Source Data** — the long-form financial statement database.

## Data pipeline

```text
Audited Financial Statements
→ Manual Account Mapping
→ Long-form Financial Database
→ Python Ratio Calculation
→ Data Validation
→ Interactive Visualization
→ Research Interpretation
```

The app reads only the `raw_financials` worksheet and recalculates all ratios in
Python. It does not rely on Excel formula caches.

## Data structure

```text
company | scope | year | statement | account | value | unit | source_page | source_file
```

The source page and source file are retained for each record so findings can be
traced back to the audited financial statements.

## Limitations

- The analysis covers only FY 2024–2025 and should not be treated as a
  long-term trend.
- Auras is the primary peer benchmark; one peer cannot represent the entire
  thermal-management industry.
- The consolidated-minus-standalone gap is not a pure subsidiary contribution
  because consolidation includes intercompany eliminations and accounting
  adjustments.
- The results are descriptive and do not establish causal effects.
- Capital expenditure is approximated by cash paid to acquire property, plant
  and equipment.

## Technology stack

- Python
- Streamlit
- pandas
- Plotly
- openpyxl
- Excel

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Repository structure

```text
avc_auras_financial_dashboard/
├── app.py
├── AVC_Auras_dashboard_data.xlsx
├── README.md
├── requirements.txt
├── .gitignore
└── .streamlit/
    └── config.toml
```
