"""
main.py - Interactive CLI Runner for Ticker Council
"""

import asyncio
from datetime import date, datetime, timedelta
from typing import List, Optional

import yfinance as yf

from council import TickerCouncil
from forecast_store import ForecastStore
from ingestion import DataIngestionPipeline


def _lookup_daily_closes(ticker: str, start: date, end: date):
    history = yf.Ticker(ticker).history(
        start=start.isoformat(),
        end=(end + timedelta(days=1)).isoformat(),
    )
    return [
        (timestamp.date(), float(close))
        for timestamp, close in history["Close"].dropna().items()
    ] if not history.empty else []


async def run_pipeline(
    symbol: str, ingestor: Optional[DataIngestionPipeline] = None
):
    symbol = symbol.strip().upper()
    if not symbol:
        print("[Error] Ticker symbol cannot be empty.")
        return

    print(f"\n{'=' * 60}")
    print(f"STARTING COUNCIL EVALUATION FOR: {symbol}")
    print(f"{'=' * 60}")

    # 1. Fetch live market and regulatory feeds
    ingestor = ingestor or DataIngestionPipeline()
    print(f"[*] Ingesting live market data for {symbol}...\n")
    market_data = await ingestor.fetch_all(symbol)

    # CLI Observability: Print ingestion status
    print("--- INGESTION METADATA ---")
    sources = market_data.get("_metadata", {}).get("sources", {})
    for seat, meta in sources.items():
        status = meta.get("status", "unknown")
        source_name = meta.get("source")
        as_of_date = meta.get("as_of_date")
        details = []
        if source_name:
            details.append(source_name)
        if as_of_date:
            details.append(f"as of {as_of_date}")
        detail_text = f" ({', '.join(details)})" if details else ""
        error = meta.get("error")
        err_str = f" | Error: {error}" if error else ""
        print(f"[{status.upper():<14}] {seat:<20}{detail_text}{err_str}")
    print("-" * 26 + "\n")

    # 2. Feed the ingested matrix into the multi-agent council
    council = TickerCouncil()
    verdict = await council.run(symbol, market_data)

    store = ForecastStore()
    collected_at = market_data.get("_metadata", {}).get("collected_at")
    created_at = datetime.fromisoformat(collected_at) if collected_at else None
    price_data = market_data.get("Price_Chart", {})
    baseline_price = price_data.get("metrics", {}).get("current_price")
    if baseline_price is None and price_data.get("closes"):
        baseline_price = price_data["closes"][-1]
    stored_count = store.record_verdict(verdict, baseline_price, created_at)
    print(f"[*] Saved {stored_count} forecast(s) to {store.database_path}.")

    scoring = store.score_due_forecasts(_lookup_daily_closes)
    brier = scoring["average_brier_score"]
    brier_text = f"{brier:.4f}" if brier is not None else "not available yet"
    print(
        f"[*] Backtest: scored {scoring['scored']} of {scoring['due']} due forecast(s); "
        f"average Brier score: {brier_text}.\n"
    )

    # 3. Print out structured results
    print("\n" + "=" * 60)
    print(f"FINAL COUNCIL VERDICT: {verdict.ticker}")
    print("=" * 60)

    print("\n--- SPECIALIST SEAT VOTES ---")
    for seat in verdict.seat_breakdown:
        print(f"[{seat.signal.value:<7}] {seat.seat_name:<20} | Conf: {seat.confidence:.2f} | {seat.summary}")

    print("\n--- ADVERSARIAL DEBATE ---")
    print(f"[Bull Case]:\n{verdict.debate.bull_case}\n")
    print(f"[Bear Case]:\n{verdict.debate.bear_case}\n")
    print(f"[Challenger Critique]:\n{verdict.debate.challenger_critique}\n")

    print("--- CALIBRATED PROBABILISTIC FORECASTS ---")
    for fc in verdict.forecasts:
        print(f"Horizon: {fc.horizon:<10} | Signal: {fc.direction.value:<4} | Prob Up: {fc.probability_up * 100:.1f}%")
        print(f"  Reasoning: {fc.reasoning}\n")


async def run_all_tickers(
    tickers: List[str], ingestor: Optional[DataIngestionPipeline] = None
) -> List[tuple[str, str]]:
    failures: List[tuple[str, str]] = []
    print(f"[*] Running {len(tickers)} ticker(s) sequentially.")
    for index, ticker in enumerate(tickers, start=1):
        print(f"\n[*] Batch progress: {index}/{len(tickers)}")
        try:
            await run_pipeline(ticker, ingestor)
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
            failures.append((ticker, error))
            print(f"[Error] {ticker} failed: {error}")

    print("\n--- BATCH SUMMARY ---")
    print(f"Completed: {len(tickers) - len(failures)}/{len(tickers)}")
    if failures:
        print("Failed tickers:")
        for ticker, error in failures:
            print(f"- {ticker}: {error}")
    else:
        print("All tickers completed successfully.")
    return failures


def prompt_ticker(available_tickers: Optional[List[str]] = None) -> str:
    if available_tickers:
        print(f"Tickers in the Capital IQ Report worksheet: {', '.join(available_tickers)}")
    while True:
        prompt = (
            "Enter a ticker symbol or ALL to evaluate every ticker: "
            if available_tickers
            else "Enter ticker symbol (e.g. AAPL, NVDA, MSFT): "
        )
        symbol = input(prompt).strip().upper()
        if symbol:
            if available_tickers and symbol == "ALL":
                return symbol
            if available_tickers and symbol not in available_tickers:
                print("Ticker is not listed in the Report worksheet. Update report.xlsx or choose a listed ticker.")
                continue
            return symbol
        print("Input cannot be empty. Please enter a valid stock ticker.")


if __name__ == "__main__":
    pipeline = DataIngestionPipeline()
    try:
        ticker_list = pipeline.capital_iq_tickers()
    except (OSError, ValueError) as exc:
        raise SystemExit(f"[Error] {exc}") from exc
    if not ticker_list:
        raise SystemExit("[Error] No tickers found in the Report worksheet of report.xlsx.")
    ticker_input = prompt_ticker(ticker_list)
    if ticker_input == "ALL":
        asyncio.run(run_all_tickers(ticker_list, pipeline))
    else:
        asyncio.run(run_pipeline(ticker_input, pipeline))