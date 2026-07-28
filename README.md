# AVC vs. Auras Financial Decision Dashboard

A Streamlit dashboard comparing:

- AVC consolidated, FY 2024–2025
- AVC standalone, FY 2024–2025
- Auras consolidated, FY 2024–2025

The app reads the `raw_financials` worksheet and recalculates all ratios in Python.

## Included pages

- Executive Summary
- Liquidity
- Profitability
- Source Data

## Run locally

1. Install Python.
2. Open a terminal in this project folder.
3. Install the required packages:

```bash
pip install -r requirements.txt
```

4. Start the web app:

```bash
streamlit run app.py
```

Streamlit will display a local web address in the terminal.

## Project structure

```text
avc_auras_streamlit_dashboard/
├── app.py
├── requirements.txt
├── README.md
├── data/
│   └── AVC_Auras_dashboard_data.xlsx
└── .streamlit/
    └── config.toml
```

## Data design

The app uses the long-form `raw_financials` sheet:

```text
company | scope | year | statement | account | value | unit | source_page | source_file
```

This design makes it easy to add new years, companies, and accounts later without rewriting the dashboard.
