# AI Equity Research & Forecasting Engine

A Python command-line research prototype that combines market and fundamental data with structured AI analysis. It evaluates a selected ticker through specialist “seats,” creates an adversarial bull/bear/challenger debate, and produces probabilistic outlooks for three horizons. Forecasts are saved locally and can be scored later against observed prices.

> **Research tool only.** This project does not provide personalized financial advice, execute trades, or establish that its model-generated probabilities are calibrated. Data coverage, freshness, accuracy, API availability, and usage rights vary by source. Verify material facts independently before making financial decisions.

## What it does

1. Reads the available ticker list and refreshed Capital IQ values from the `Report` worksheet in an Excel workbook.
2. Fetches Yahoo Finance price history, selected company data, earnings information, and headlines. It also attempts to retrieve 2-year and 10-year Treasury yields from FRED.
3. Builds per-seat inputs and source-status metadata. Seats with missing, failed, or unconfigured sources are marked `NO_READ` rather than given an invented directional opinion. Partial data can still be analyzed, with prompts limited to supplied metrics.
4. Runs specialist analyses, a bull/bear/challenger debate, and a judge that returns forecasts for `1_week`, `3_months`, and `1_year`.
5. Stores each horizon forecast in `forecasts.sqlite3`. When a target date is due, it looks for the first available Yahoo Finance daily close on or after that date (up to seven calendar days late) and calculates realized return and a Brier score for the predicted probability of a positive return.

Forecast direction must match its probability: `UP` above 55%, `DOWN` below 45%, and `EVEN` from 45% through 55%, inclusive. Backtest results are historical summaries, not proof of future performance or calibration.

## Current coverage and limitations

| Area | Current inputs / behavior |
|---|---|
| Price chart | Yahoo Finance daily history; technical metrics include moving averages, 30-day realized volatility, volume, RSI, and limited comparison with SPY. |
| Financials | Yahoo Finance selected valuation, cash, debt, free-cash-flow, and margin fields, with Capital IQ workbook values overriding fields populated in Excel. |
| Analyst targets | Capital IQ workbook values when supplied. |
| Earnings estimates | Yahoo Finance selected EPS and earnings-surprise data, supplemented by Capital IQ workbook fields. |
| Macro economy | Optional FRED 2-year and 10-year Treasury yields and their spread. The spread is context, not ticker-specific directional proof. |
| News | Yahoo Finance headlines when available. |
| Insider trades | Yahoo descriptions are not verified SEC Form 4 transaction codes; the council deliberately assigns no directional signal from them. |
| Congress trades, SEC filings, institutional flows, options, related markets | No dedicated verified feed is currently wired in. Capital IQ aggregate institutional ownership can provide ownership context, but it is not real-time buying/selling flow. |

The council defines twelve seats, but a seat definition does not mean its data source is implemented. Unsupported or unavailable sources will show as `NO_READ`. Do not interpret an empty seat as evidence that no relevant activity occurred.

## Requirements

- Python 3.10 or later (developed and tested with Python 3.13).
- An OpenAI API key for the AI analysis. 
- A refreshed Capital IQ Excel workbook for Capital IQ fields. You must have appropriate rights to use and store the data.
- A FRED API key is optional; without it, macro data is unavailable.

## Setup and run (Windows PowerShell)

Open PowerShell in the project directory, then install the third-party packages:

```powershell
py -m pip install openai pydantic httpx pandas yfinance
```

Set the OpenAI key for the current PowerShell window. Do not paste a real key into source files, README content, screenshots, or commits:

```powershell
$env:OPENAI_API_KEY = "YOUR_OPENAI_API_KEY"
```

Optionally configure FRED:

```powershell
$env:FRED_API_KEY = "YOUR_FRED_API_KEY"
```

Place a workbook named `report.xlsx` beside `main.py`, refresh its Capital IQ formulas in Excel, and save it. Then start the CLI:

```powershell
py main.py
```

The ticker prompt is built from non-empty tickers in the workbook's `Report` worksheet. Enter one listed symbol to analyze just that stock, or enter `ALL` to analyze every listed ticker sequentially. A failed ticker is reported and the batch continues; a failure summary is printed at the end. Running all tickers can take a while and may generate substantial OpenAI API usage/costs. The workbook must be saved after the Capital IQ refresh because this program reads cached formula results; it does not calculate Capital IQ formulas.


## Capital IQ workbook

The worksheet is named `Report`. Put headers in row 1 and one ticker per row beginning in row 2. The importer recognizes the fields currently in the excel:

```text
ticker,as_of_date,target_mean,target_high,target_low,number_of_analysts,recommendation_key,forward_eps,forward_revenue,eps_growth_pct,revenue_growth_pct,pe_ratio,price_to_book,free_cash_flow,market_cap,total_debt,cash_and_equivalents,operating_margins_pct,ebitda_margin_pct,held_percent_institutions_pct
```

`ticker` and `as_of_date` are required for each usable row. Use `YYYY-MM-DD` for the date. Numeric fields may be blank; populated Capital IQ values override matching Yahoo values, while blanks leave Yahoo values unchanged. Percentage fields use percentage points (for example, `31.5` means 31.5%). `ebitda_margin_pct` is EBITDA margin, not operating margin. If you add columns, verify the Capital IQ item and returned units in the add-in and update both the importer and its tests.

## Project layout

- `main.py` — CLI, ingestion-to-council orchestration, and forecast scoring trigger.
- `ingestion.py` — Yahoo/FRED collection, Capital IQ XLSX parsing, per-seat payloads, and source metadata.
- `council.py` — structured output models, seat definitions and validation, AI orchestration, debate, and forecast generation.
- `forecast_store.py` — SQLite persistence and outcome/Brier-score calculations.
- `tests/test_pipeline_regressions.py` — local regression tests, including workbook parsing, validation, source routing, and forecast storage.
- `report.xlsx` — local Capital IQ input workbook; do not publish unless you have confirmed you are allowed to redistribute its contents.
- `forecasts.sqlite3` — local forecast history created by the application; it may contain ticker choices and generated analysis.

## Tests

Run the regression suite from the project directory:

```powershell
py -m unittest discover -s tests -v
```

Tests are local and should not require live Yahoo, FRED, Capital IQ, or OpenAI calls. The project does not currently include a dependency lockfile, packaging configuration, or CI workflow.

## Roadmap and improvements

### Priority 1 — safe and reproducible GitHub publication

- Add a `.gitignore` for Python caches, virtual environments, `.env` files, generated SQLite databases, and local workbooks.
- Do not commit `report.xlsx` or Capital IQ cached results unless the license explicitly permits public redistribution. Publish a blank/template workbook or schema documentation instead, and provide the real workbook privately to authorized users.
- Check the complete repository history and files for API keys, personal data, and licensed data before pushing. If a key was exposed, revoke and replace it.
- Add `requirements.txt` (or `pyproject.toml`) with tested dependency versions, a supported Python range, and repeatable setup instructions.
- Choose a project license only after deciding what permissions you want to grant; do not assume that public visibility automatically grants reuse rights.

### Priority 2 — data reliability and observability

- Add clear per-source timestamps, currency/units, freshness checks, and structured error reporting.
- Verify Yahoo Finance fields and behavior against representative symbols; treat Yahoo data as unofficial and subject to change.
- Add explicit source adapters for validated SEC Form 4 data before enabling directional insider analysis.
- Add targeted, documented integrations only where available and licensed: SEC filings, news, congressional disclosures, options, and institutional holdings. Distinguish quarterly 13F holdings from real-time flows.
- Validate workbook headers, ticker uniqueness, dates, currencies, and formula refresh state; provide a workbook template or validation command.

### Priority 3 — forecast quality and testing

- Add integration tests using recorded, sanitized fixtures and mock external services; test source failures, stale data, and partial data.
- Add CI to install supported Python versions, run tests, and perform linting/type checks.
- Evaluate forecast accuracy by horizon, ticker, time period, and baseline; track sample counts and compare with simple benchmarks before claiming calibration.
- Make benchmark and return definitions explicit, including dividends, splits, currencies, trading-calendar handling, and the effect of missing prices.
- Consider storing reproducibility metadata (input snapshot, model/provider, prompt/schema version, source timestamps) with each forecast. Avoid describing records as immutable or cryptographically hashed unless that is implemented.

### Priority 4 — usability and maintainability

- Separate configuration from code and validate required API keys before starting paid model calls.
- Add command-line options for ticker, workbook path, dry-run/data-only mode, and forecast reporting.
- Add a forecast-history/reporting command and database migration/versioning strategy.
- Separate providers and analysis orchestration behind small interfaces so sources and model providers can be changed and tested independently.
- Add logging, retry/backoff policies, API cost tracking, and concurrency/rate limits appropriate to each provider.
