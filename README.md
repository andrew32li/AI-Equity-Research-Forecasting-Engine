# AI Equity Research & Forecasting Engine

A Python command-line research prototype that combines market and fundamental data with structured AI analysis. It evaluates a selected ticker through specialist “seats,” creates an adversarial bull/bear/challenger debate, and produces probabilistic outlooks for three horizons. Forecasts are saved locally and can be scored later against observed prices.

> **Research tool only.** This project does not provide personalized financial advice, execute trades, or establish that its model-generated probabilities are calibrated. Data coverage, freshness, accuracy, API availability, and usage rights vary by source. Verify material facts independently before making financial decisions.

## What it does

1. Reads the available ticker list and refreshed Capital IQ values from the `Report` worksheet in an Excel workbook.
2. Fetches Yahoo Finance price history, financial and earnings data, institutional ownership/short-interest metrics, options-chain snapshots, news, and sector ETF proxy performance. Google News RSS is used as a headline fallback. It also attempts to retrieve 2-year and 10-year Treasury yields from FRED.
3. Uses SEC EDGAR company submissions and Form 4 XML for verified insider transactions and recent 10-K/10-Q/8-K excerpts. It scans recent House Clerk PTR PDFs and can read manually imported Senate disclosures.
4. Builds per-seat inputs and source-status metadata. Seats with missing, failed, or unconfigured sources are marked `NO_READ` rather than given an invented directional opinion. Partial data can still be analyzed, with prompts limited to supplied metrics.
5. Runs specialist analyses, a bull/bear/challenger debate, and a judge that returns forecasts for `1_week`, `3_months`, and `1_year`.
6. Stores each horizon forecast in SQLite. When a target date is due, it looks for the first available Yahoo Finance daily close on or after that date (up to seven calendar days late) and calculates realized return and a Brier score for the predicted probability of a positive return.

Forecast direction must match its probability: `UP` above 55%, `DOWN` below 45%, and `EVEN` from 45% through 55%, inclusive. Backtest results are historical summaries, not proof of future performance or calibration.

## Current coverage and limitations

| Area | Current inputs / behavior |
|---|---|
| Price chart | Yahoo Finance daily history; technical metrics include moving averages, 30-day realized volatility, volume, RSI, and limited comparison with SPY. |
| Financials | Yahoo Finance selected valuation, cash, debt, free-cash-flow, and margin fields; Capital IQ workbook values override matching populated fields. |
| Analyst targets | Capital IQ workbook values when supplied. |
| Earnings estimates | Yahoo Finance selected EPS and earnings-surprise data, supplemented by Capital IQ workbook fields. |
| Macro economy | Optional FRED 2-year and 10-year Treasury yields and their spread. The spread is context, not ticker-specific directional proof. |
| News | Yahoo Finance headlines, with Google News RSS fallback. RSS/search coverage is not guaranteed; headlines may be duplicated or refer to similarly named entities. |
| Insider trades | SEC EDGAR Form 4 filings from the most recent 90 days. Only non-derivative transaction codes `P` (open-market purchase) and `S` (open-market sale) are counted. Form 4 reporting is delayed and not a real-time tape. |
| Congress trades | House Clerk index/PTR PDFs, scanning up to 75 most recent reports filed during the last 90 days, plus optional Senate CSV import. House symbol/company-name matching is best effort; confirm every match in its linked report. Trade values are ranges and reports can be delayed. No match is not proof of no trade. |
| SEC filings | Recent 10-K/10-Q and up to three 8-K excerpts from EDGAR, for up to one year of filings. Excerpts are truncated and are not a substitute for reading the complete filing. |
| Institutional flows | Capital IQ aggregate ownership when supplied; Yahoo Finance institutional/insider ownership and short-interest fields when available. Ownership is not buying/selling flow; short-interest data can be stale. |
| Options market | Yahoo Finance option-chain snapshot for up to the nearest three expirations. Ratios and implied volatility are snapshots and do not show trade direction, unusual flow, or gamma walls. Availability varies. |
| Related markets | Yahoo Finance sector classification and an ETF proxy for recognized sectors, compared with SPY over available periods. ETF returns are proxies, not official index returns; supply-chain, currency, and commodity relationships are not covered. |

Yahoo Finance endpoints and Google News RSS are best-effort public services and may change or restrict access. SEC automated requests require a descriptive contact User-Agent and should be made responsibly. Data and reports are not guaranteed complete, current, or suitable for redistribution. Do not interpret `NO_READ` or an empty result as evidence that no relevant activity occurred.

## Requirements

- Python 3.10 or later (developed and tested with Python 3.13).
- An OpenAI API key for the AI analysis. 
- A contact email for SEC EDGAR requests, configured locally as `SEC_CONTACT_EMAIL` or `SEC_USER_AGENT`.
- A refreshed Capital IQ Excel workbook for Capital IQ fields. You must have appropriate rights to use and store the data.
- A FRED API key is optional; without it, macro data is unavailable.

## Setup and run (Windows PowerShell)

Open PowerShell in the project directory, then install the third-party packages:

```powershell
py -m pip install -r requirements.txt
```

Set the OpenAI key for the current PowerShell window. Do not paste a real key into source files, README content, screenshots, or commits:

```powershell
$env:OPENAI_API_KEY = "YOUR_OPENAI_API_KEY"
```

Configure SEC EDGAR with the email you provided for this purpose. Keep it in your local shell/environment, not in source files or commits:

```powershell
$env:SEC_CONTACT_EMAIL = "your-contact-email@example.com"
```

Optionally configure FRED:

```powershell
$env:FRED_API_KEY = "YOUR_FRED_API_KEY"
```

### Congressional trade data

The House Clerk reports are fetched automatically, but the scan is capped to the 75 most recent House PTR reports filed in the preceding 90 days to bound download volume. Review any matched filing link before relying on a match.

The Senate portal does not expose a documented general-purpose API and automated access may be blocked. To import Senate reports manually, copy `congress_trades.example.csv` to `congress_trades.csv` in the project folder and add verified records from the official Senate disclosure portal. Required CSV columns:

```text
ticker,chamber,transaction_date,disclosure_date,asset_description,transaction_type,amount_range,source_url
```

Use ISO `YYYY-MM-DD` dates, `Senate` as the chamber, and the reported amount range as text. The importer keeps records from the last 90 days that include a source URL. The local CSV is excluded from Git. To use another file path, set `CONGRESS_TRADES_CSV_PATH`.

Place a workbook named `report.xlsx` beside `main.py`, refresh its Capital IQ formulas in Excel, and save it. Then start the CLI:

```powershell
py main.py
```

The ticker prompt is built from non-empty tickers in the workbook's `Report` worksheet. Enter one listed symbol to analyze just that stock, or enter `ALL` to analyze every listed ticker sequentially. A single-ticker run creates a timestamped `equity_research_report_<TICKER>_*.docx` in the project directory. Batch runs create a timestamped `equity_research_report_*.docx` containing forecast summaries, seat votes, debate, reasoning, and any ticker errors. A failed ticker is reported and the batch continues; a failure summary is printed at the end. Running all tickers can take a while and may generate substantial OpenAI API usage/costs. The workbook must be saved after the Capital IQ refresh because this program reads cached formula results; it does not calculate Capital IQ formulas.


## Capital IQ workbook

The worksheet is named `Report`. Put headers in row 1 and one ticker per row beginning in row 2. The importer recognizes the fields currently in the excel:

```text
ticker,as_of_date,target_mean,target_high,target_low,number_of_analysts,recommendation_key,forward_eps,forward_revenue,eps_growth_pct,revenue_growth_pct,pe_ratio,price_to_book,free_cash_flow,market_cap,total_debt,cash_and_equivalents,operating_margins_pct,ebitda_margin_pct,held_percent_institutions_pct
```

`ticker` and `as_of_date` are required for each usable row. Use `YYYY-MM-DD` for the date. Numeric fields may be blank; populated Capital IQ values override matching Yahoo values, while blanks leave Yahoo values unchanged. Percentage fields use percentage points (for example, `31.5` means 31.5%). `ebitda_margin_pct` is EBITDA margin, not operating margin. If you add columns, verify the Capital IQ item and returned units in the add-in and update both the importer and its tests.

## Project layout

- `main.py` — CLI, ingestion-to-council orchestration, and forecast scoring trigger.
- `ingestion.py` — Yahoo/FRED/SEC/House disclosure collection, Senate CSV import, Capital IQ XLSX parsing, per-seat payloads, and source metadata.
- `council.py` — structured output models, seat definitions and validation, AI orchestration, debate, and forecast generation.
- `forecast_store.py` — SQLite persistence and outcome/Brier-score calculations.
- `forecast_db/forecasts.sqlite3` — local forecast history created by the application; it may contain ticker choices and generated analysis.
- `tests/test_pipeline_regressions.py` — local regression tests, including workbook parsing, validation, source routing, and forecast storage.
- `report.xlsx` — local Capital IQ input workbook; do not publish unless you have confirmed you are allowed to redistribute its contents.

## Tests

Run the regression suite from the project directory:

```powershell
py -m unittest discover -s tests -v
```

Tests are local and should not require live Yahoo, FRED, SEC, Congress, Capital IQ, or OpenAI calls. Install dependencies with `py -m pip install -r requirements.txt`.

## Roadmap and improvements

### Priority 1 — repository and release hygiene

- Before each push, check the complete repository history and files for API keys, personal data, and licensed data. If a key was exposed, revoke and replace it.
- Add a CI workflow that installs `requirements.txt` and runs the local test suite.
- Define and test the supported Python version range in CI.
- Choose a project license only after deciding what permissions you want to grant; do not assume that public visibility automatically grants reuse rights.

### Priority 2 — source reliability and observability

- Expand mocked-fixture tests for SEC submissions/Form 4 XML, filing excerpts, House index/PDF parsing, RSS parsing, and Yahoo options payloads.
- Verify Yahoo Finance fields and behavior against representative symbols; treat Yahoo data as unofficial and subject to change.
- Add freshness checks and clearly surfaced per-source timestamps, units, and errors; distinguish empty data from unavailable or incomplete data.
- Keep Senate CSV entries linked to their official disclosures and review House ticker/name matches against the linked reports. Neither source is complete or real-time.
- Validate workbook ticker uniqueness, currencies, and formula refresh state; provide a validation command or a clean sample workbook.
- Do not describe institutional ownership as a trading flow; obtain a separately validated 13F source before analyzing quarter-over-quarter holdings changes.

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
