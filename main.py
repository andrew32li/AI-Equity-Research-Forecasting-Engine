"""
main.py - Interactive CLI Runner for Ticker Council
"""

import asyncio
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import List, Optional
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

import yfinance as yf

from council import CouncilVerdict, TickerCouncil
from forecast_store import ForecastStore
from ingestion import DataIngestionPipeline


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("w", WORD_NS)
ET.register_namespace("", PACKAGE_REL_NS)
ET.register_namespace("ct", CONTENT_TYPES_NS)


def _word_tag(name: str) -> str:
    return f"{{{WORD_NS}}}{name}"


def _add_word_paragraph(
    parent: ET.Element,
    text: str,
    *,
    font_size: Optional[int] = None,
    bold: bool = False,
    color: Optional[str] = None,
    page_break_before: bool = False,
) -> None:
    paragraph = ET.SubElement(parent, _word_tag("p"))
    if page_break_before:
        paragraph_properties = ET.SubElement(paragraph, _word_tag("pPr"))
        ET.SubElement(paragraph_properties, _word_tag("pageBreakBefore"))

    run = ET.SubElement(paragraph, _word_tag("r"))
    if font_size is not None or bold or color:
        run_properties = ET.SubElement(run, _word_tag("rPr"))
        if bold:
            ET.SubElement(run_properties, _word_tag("b"))
        if font_size is not None:
            ET.SubElement(
                run_properties, _word_tag("sz"), {_word_tag("val"): str(font_size)}
            )
        if color:
            ET.SubElement(
                run_properties, _word_tag("color"), {_word_tag("val"): color}
            )

    lines = str(text).split("\n")
    for index, line in enumerate(lines):
        if index:
            ET.SubElement(run, _word_tag("br"))
        text_node = ET.SubElement(run, _word_tag("t"))
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_node.text = line


def _add_word_table(parent: ET.Element, headers: List[str], rows: List[List[str]]) -> None:
    table = ET.SubElement(parent, _word_tag("tbl"))
    properties = ET.SubElement(table, _word_tag("tblPr"))
    borders = ET.SubElement(properties, _word_tag("tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        ET.SubElement(
            borders,
            _word_tag(edge),
            {
                _word_tag("val"): "single",
                _word_tag("sz"): "4",
                _word_tag("space"): "0",
                _word_tag("color"): "B7C9D6",
            },
        )

    for row_index, values in enumerate([headers, *rows]):
        row = ET.SubElement(table, _word_tag("tr"))
        for value in values:
            cell = ET.SubElement(row, _word_tag("tc"))
            if row_index == 0:
                cell_properties = ET.SubElement(cell, _word_tag("tcPr"))
                ET.SubElement(
                    cell_properties,
                    _word_tag("shd"),
                    {
                        _word_tag("val"): "clear",
                        _word_tag("color"): "auto",
                        _word_tag("fill"): "1F4E78",
                    },
                )
            _add_word_paragraph(
                cell,
                str(value),
                font_size=20,
                bold=row_index == 0,
                color="FFFFFF" if row_index == 0 else None,
            )


def write_batch_report(
    verdicts: List[CouncilVerdict],
    failures: List[tuple[str, str]],
    output_path: Path,
    report_title: str = "Batch Research Report",
) -> Path:
    document = ET.Element(_word_tag("document"))
    body = ET.SubElement(document, _word_tag("body"))

    _add_word_paragraph(
        body,
        "AI Equity Research & Forecasting Engine",
        font_size=36,
        bold=True,
        color="1F4E78",
    )
    _add_word_paragraph(body, report_title, font_size=28, bold=True)
    _add_word_paragraph(
        body,
        f"Generated: {datetime.now().astimezone().strftime('%Y-%m-%d %H:%M %Z')}",
    )
    _add_word_paragraph(
        body,
        f"Completed successfully: {len(verdicts)} of "
        f"{len(verdicts) + len(failures)} ticker(s).",
        bold=True,
    )
    _add_word_paragraph(
        body,
        "Research only; forecasts are AI-generated estimates, not financial advice "
        "or empirically established calibrated probabilities.",
        color="7F6000",
    )

    if verdicts:
        _add_word_paragraph(body, "Forecast Summary", font_size=26, bold=True)
        summary_rows = [
            [
                verdict.ticker,
                forecast.horizon,
                forecast.direction.value,
                f"{forecast.probability_up:.1%}",
            ]
            for verdict in verdicts
            for forecast in verdict.forecasts
        ]
        _add_word_table(
            body,
            ["Ticker", "Horizon", "Direction", "Probability Up"],
            summary_rows,
        )

        _add_word_paragraph(body, "Detailed Results", font_size=26, bold=True)
        for verdict in verdicts:
            _add_word_paragraph(
                body, verdict.ticker, font_size=30, bold=True, color="1F4E78"
            )
            _add_word_paragraph(body, "Specialist Seat Votes", font_size=22, bold=True)
            _add_word_table(
                body,
                ["Seat", "Signal", "Confidence", "Summary"],
                [
                    [
                        opinion.seat_name,
                        opinion.signal.value,
                        f"{opinion.confidence:.2f}",
                        opinion.summary,
                    ]
                    for opinion in verdict.seat_breakdown
                ],
            )

            _add_word_paragraph(body, "Adversarial Debate", font_size=22, bold=True)
            _add_word_paragraph(body, f"Bull Case: {verdict.debate.bull_case}")
            _add_word_paragraph(body, f"Bear Case: {verdict.debate.bear_case}")
            _add_word_paragraph(
                body,
                f"Challenger Critique: {verdict.debate.challenger_critique}",
            )

            _add_word_paragraph(body, "Forecasts", font_size=22, bold=True)
            _add_word_table(
                body,
                ["Horizon", "Direction", "Probability Up", "Reasoning"],
                [
                    [
                        forecast.horizon,
                        forecast.direction.value,
                        f"{forecast.probability_up:.1%}",
                        forecast.reasoning,
                    ]
                    for forecast in verdict.forecasts
                ],
            )

    if failures:
        _add_word_paragraph(body, "Tickers with Errors", font_size=26, bold=True)
        _add_word_table(body, ["Ticker", "Error"], [[ticker, error] for ticker, error in failures])

    section_properties = ET.SubElement(body, _word_tag("sectPr"))
    ET.SubElement(
        section_properties,
        _word_tag("pgSz"),
        {_word_tag("w"): "12240", _word_tag("h"): "15840"},
    )
    ET.SubElement(
        section_properties,
        _word_tag("pgMar"),
        {
            _word_tag("top"): "900",
            _word_tag("right"): "900",
            _word_tag("bottom"): "900",
            _word_tag("left"): "900",
            _word_tag("header"): "450",
            _word_tag("footer"): "450",
            _word_tag("gutter"): "0",
        },
    )

    content_types = ET.Element(f"{{{CONTENT_TYPES_NS}}}Types")
    ET.SubElement(
        content_types,
        f"{{{CONTENT_TYPES_NS}}}Default",
        {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"},
    )
    ET.SubElement(
        content_types,
        f"{{{CONTENT_TYPES_NS}}}Default",
        {"Extension": "xml", "ContentType": "application/xml"},
    )
    ET.SubElement(
        content_types,
        f"{{{CONTENT_TYPES_NS}}}Override",
        {
            "PartName": "/word/document.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        },
    )

    package_relationships = ET.Element(f"{{{PACKAGE_REL_NS}}}Relationships")
    ET.SubElement(
        package_relationships,
        f"{{{PACKAGE_REL_NS}}}Relationship",
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "Target": "word/document.xml",
        },
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output_path, "w", ZIP_DEFLATED) as report:
        report.writestr(
            "[Content_Types].xml",
            ET.tostring(content_types, encoding="utf-8", xml_declaration=True),
        )
        report.writestr(
            "_rels/.rels",
            ET.tostring(package_relationships, encoding="utf-8", xml_declaration=True),
        )
        report.writestr(
            "word/document.xml",
            ET.tostring(document, encoding="utf-8", xml_declaration=True),
        )
    return output_path


def write_ticker_report(
    verdict: CouncilVerdict, report_directory: Optional[Path] = None
) -> Path:
    output_directory = report_directory or Path(__file__).resolve().parent
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    report_path = (
        output_directory
        / f"equity_research_report_{verdict.ticker}_{timestamp}.docx"
    )
    write_batch_report(
        [verdict],
        [],
        report_path,
        report_title="Equity Research Report",
    )
    print(f"[*] Word report saved to: {report_path}")
    return report_path


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
) -> Optional[CouncilVerdict]:
    symbol = symbol.strip().upper()
    if not symbol:
        print("[Error] Ticker symbol cannot be empty.")
        return None

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
    return verdict


async def run_all_tickers(
    tickers: List[str],
    ingestor: Optional[DataIngestionPipeline] = None,
    report_directory: Optional[Path] = None,
) -> List[tuple[str, str]]:
    failures: List[tuple[str, str]] = []
    verdicts: List[CouncilVerdict] = []
    print(f"[*] Running {len(tickers)} ticker(s) sequentially.")
    for index, ticker in enumerate(tickers, start=1):
        print(f"\n[*] Batch progress: {index}/{len(tickers)}")
        try:
            verdict = await run_pipeline(ticker, ingestor)
            if verdict is None:
                raise RuntimeError("Ticker run completed without producing a verdict.")
            verdicts.append(verdict)
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
    output_directory = report_directory or Path(__file__).resolve().parent
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    report_path = output_directory / f"equity_research_report_{timestamp}.docx"
    write_batch_report(verdicts, failures, report_path)
    print(f"[*] Word report saved to: {report_path}")
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
        verdict = asyncio.run(run_pipeline(ticker_input, pipeline))
        if verdict is not None:
            write_ticker_report(verdict)