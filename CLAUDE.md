# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

UK House Price Index (HPI) data analytics project. Source data is `PowerBIData.xlsx` (sheet: `HPIAdmins`) covering 6 English metropolitan regions from Jan 1995 – Oct 2016. The cleaned working dataset is `HPIAdmins_v1.xlsx`.

Original Power BI file: `HPI_Admins.pbix` — edited via Power BI Desktop, not via code.

## Python Pipeline

All scripts run with plain `python <script>.py` — no build system, no virtual environment configured.

| Script | Purpose |
|---|---|
| `explore.py` | Profile shape, nulls, duplicates, outliers on `HPIAdmins` sheet |
| `validate.py` | Data quality validation — flags issues and prints summary |
| `clean.py` | Produces `HPIAdmins_v1.xlsx`: drops 100% null columns, renames `%` columns, adds `DataGapFlag` |
| `build_dashboard.py` | Generates `20260515/HPIAdmins.html` — single combined dashboard |
| `build_dashboards_all.py` | Generates four audience-specific dashboards in `20260515/` |
| `build_ppt_pdf.py` | Generates `HPI_Dashboard_Presentation.pptx` and `HPI_Dashboard_UserGuide.pdf` in `20260515/Presentation/` |

Required packages: `pandas`, `openpyxl`, `python-pptx`, `reportlab`.

## Data Notes

- `IndexSA` and `AveragePriceSA` are 100% null — already dropped in `HPIAdmins_v1.xlsx`.
- `%Change` columns have structural nulls for the first 1 month (MoM) and first 12 months (YoY) — expected, not errors.
- `DataGapFlag = True` on 12 rows (Nov–Dec 2016, all 6 regions) where `SalesVolume = 0`.
- Column names were sanitised: `1m%Change` → `1mPctChange`, etc.

## Output Structure (`20260515/`)

Four HTML dashboards (self-contained, no server needed):
- `CEO_HPIAdmins.html` — headline KPIs + 21-yr trend, no filters
- `ExCo_HPIAdmins.html` — tabbed: Market / Regional / Segments / Growth
- `Mgr_HPIAdmins.html` — region + year-range filters, MoM/YoY charts, ranking table
- `Ops_HPIAdmins.html` — full data console, 1,572-row searchable/sortable table

Presentation materials (`20260515/Presentation/`):
- `HPI_Dashboard_Presentation.pptx` — 9-slide intro deck
- `HPI_Dashboard_UserGuide.pdf` — detailed per-dashboard instructions + data dictionary

## Dashboard Architecture

All HTML dashboards embed data as inline JSON (no external data calls). Chart.js 4.4 is loaded from CDN. Each dashboard is fully self-contained after that initial CDN fetch.

The Python generation scripts (`build_dashboard.py`, `build_dashboards_all.py`) compute all aggregates at generation time and serialize them into the HTML as `const` JS variables. To refresh dashboards with new data, re-run the relevant build script against an updated `HPIAdmins_v1.xlsx`.

## Installed Skills

| Skill | Invoke | Use when |
|---|---|---|
| `analyze` | `/analyze` | Ad-hoc data questions & trend investigations |
| `build-dashboard` | `/build-dashboard` | Generate interactive HTML dashboards |
| `create-viz` | `/create-viz` | Publication-quality Python visualizations |
| `data-context-extractor` | `/data-context-extractor` | Generate company-specific data skills |
| `data-visualization` | `/data-visualization` | Python charts (matplotlib/seaborn/plotly) |
| `explore-data` | `/explore-data` | Profile a dataset's shape, quality, anomalies |
| `frontend-design` | `/frontend-design` | Production-grade HTML/CSS/JS interfaces |
| `sql-queries` | `/sql-queries` | Cross-dialect SQL generation |
| `statistical-analysis` | `/statistical-analysis` | Descriptive stats, outlier detection |
| `validate-data` | `/validate-data` | QA analysis before sharing |
| `write-query` | `/write-query` | Optimised SQL with best practices |
