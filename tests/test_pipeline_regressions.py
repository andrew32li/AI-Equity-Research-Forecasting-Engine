import asyncio
import tempfile
import unittest
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

import pandas as pd
from pydantic import ValidationError

from council import (
    CouncilVerdict,
    DebateArguments,
    HorizonForecast,
    PrePromptValidator,
    SEAT_CONFIGS,
    SeatOpinion,
    Signal,
    TickerCouncil,
)
from forecast_store import ForecastStore
from ingestion import DataIngestionPipeline
from main import prompt_ticker, run_all_tickers, write_ticker_report


def _write_capital_iq_workbook(path: Path, headers, rows):
    spreadsheet_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", spreadsheet_ns)

    def column_name(index):
        name = ""
        while index:
            index, remainder = divmod(index - 1, 26)
            name = chr(ord("A") + remainder) + name
        return name

    worksheet = ET.Element(f"{{{spreadsheet_ns}}}worksheet")
    sheet_data = ET.SubElement(worksheet, f"{{{spreadsheet_ns}}}sheetData")
    for row_number, values in enumerate([headers, *rows], start=1):
        row_element = ET.SubElement(
            sheet_data, f"{{{spreadsheet_ns}}}row", {"r": str(row_number)}
        )
        for column_number, value in enumerate(values, start=1):
            if value is None:
                continue
            cell = ET.SubElement(
                row_element,
                f"{{{spreadsheet_ns}}}c",
                {"r": f"{column_name(column_number)}{row_number}"},
            )
            if isinstance(value, tuple):
                formula, cached_value = value
                if isinstance(cached_value, str):
                    cell.set("t", "str")
                ET.SubElement(cell, f"{{{spreadsheet_ns}}}f").text = formula
                ET.SubElement(cell, f"{{{spreadsheet_ns}}}v").text = str(cached_value)
            elif isinstance(value, str):
                cell.set("t", "inlineStr")
                inline_string = ET.SubElement(cell, f"{{{spreadsheet_ns}}}is")
                ET.SubElement(inline_string, f"{{{spreadsheet_ns}}}t").text = value
            else:
                cell.set("t", "n")
                ET.SubElement(cell, f"{{{spreadsheet_ns}}}v").text = str(value)

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    workbook_relationships = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/></Relationships>'
    )
    package_relationships = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>'
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types)
        workbook.writestr("_rels/.rels", package_relationships)
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_relationships)
        workbook.writestr(
            "xl/worksheets/sheet1.xml",
            ET.tostring(worksheet, encoding="utf-8", xml_declaration=True),
        )


class PriceChartValidationTests(unittest.TestCase):
    def test_fourteen_closes_are_not_enough_for_fourteen_rsi_changes(self):
        opinion, payload = PrePromptValidator.validate_price_chart(
            {"closes": [100.0] * 14}
        )

        self.assertEqual(opinion.signal, Signal.NO_READ)
        self.assertIsNone(payload)

    def test_fifteen_flat_closes_produce_neutral_rsi_and_trend(self):
        opinion, payload = PrePromptValidator.validate_price_chart(
            {"closes": [100.0] * 15}
        )

        self.assertIsNone(opinion)
        self.assertEqual(payload["rsi_14"], 50.0)
        self.assertEqual(payload["trend_5d"], "EVEN")


class InsiderValidationTests(unittest.TestCase):
    def test_unverified_yahoo_classification_is_no_read(self):
        opinion, payload = PrePromptValidator.validate_insider_trades(
            {
                "transactions_verified": False,
                "filings": [{"reported_type": "PURCHASE", "value_usd": 500_000}],
            }
        )

        self.assertEqual(opinion.signal, Signal.NO_READ)
        self.assertIsNone(payload)

    def test_form4_parser_only_counts_open_market_purchase_and_sale_codes(self):
        xml = b"""<ownershipDocument>
        <reportingOwner><reportingOwnerId><rptOwnerName>Jane Doe</rptOwnerName></reportingOwnerId></reportingOwner>
        <nonDerivativeTable>
          <nonDerivativeTransaction>
            <transactionDate><value>2026-10-01</value></transactionDate>
            <transactionCoding><transactionCode>P</transactionCode></transactionCoding>
            <transactionAmounts>
              <transactionShares><value>10</value></transactionShares>
              <transactionPricePerShare><value>100</value></transactionPricePerShare>
            </transactionAmounts>
          </nonDerivativeTransaction>
          <nonDerivativeTransaction>
            <transactionDate><value>2026-10-02</value></transactionDate>
            <transactionCoding><transactionCode>S</transactionCode></transactionCoding>
            <transactionAmounts>
              <transactionShares><value>5</value></transactionShares>
              <transactionPricePerShare><value>120</value></transactionPricePerShare>
            </transactionAmounts>
          </nonDerivativeTransaction>
          <nonDerivativeTransaction>
            <transactionCoding><transactionCode>A</transactionCode></transactionCoding>
          </nonDerivativeTransaction>
        </nonDerivativeTable></ownershipDocument>"""

        transactions = DataIngestionPipeline._parse_form4_xml(
            xml, "2026-10-03", "https://www.sec.gov/example.xml"
        )

        self.assertEqual([row["transaction_code"] for row in transactions], ["P", "S"])
        self.assertEqual([row["value_usd"] for row in transactions], [1000, 600])
        self.assertEqual(transactions[0]["insider"], "Jane Doe")


class CongressionalSourceTests(unittest.IsolatedAsyncioTestCase):
    async def test_house_matching_returns_source_link_and_marks_best_effort(self):
        pipeline = DataIngestionPipeline()

        async def fake_reports(_client):
            return (
                [
                    {
                        "member_name": "Representative Example",
                        "filing_date": date.today().isoformat(),
                        "state_district": "CA01",
                        "source_url": "https://disclosures-clerk.house.gov/example.pdf",
                        "report_text": "Apple Inc. purchase amount $1,001-$15,000",
                    }
                ],
                None,
                False,
            )

        pipeline._load_house_ptr_reports = fake_reports
        pipeline._load_senate_congress_rows = lambda _ticker: ([], None)

        payload, error, truncated = await pipeline._fetch_congress_trades(
            "AAPL", ["Apple Inc."], object()
        )

        self.assertIsNone(error)
        self.assertFalse(truncated)
        self.assertEqual(len(payload["house_events"]), 1)
        self.assertEqual(
            payload["house_events"][0]["source_url"],
            "https://disclosures-clerk.house.gov/example.pdf",
        )
        self.assertIn("best effort", payload["limitations"])

class SenateCsvSourceTests(unittest.TestCase):
    def test_loads_recent_senate_records_from_local_csv(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            csv_path = Path(temp_dir) / "congress_trades.csv"
            today = datetime.now(timezone.utc).date()
            csv_path.write_text(
                "Ticker,Chamber,Transaction_Date,Disclosure_Date,Asset_Description,"
                "Transaction_Type,Amount_Range,Source_Url\n"
                f"AAPL,Senate,{today.isoformat()},"
                f"{today.isoformat()},Apple Inc.,Purchase,"
                '"$1,001-$15,000",https://efdsearch.senate.gov/example\n'
                f"AAPL,Senate,{(today + timedelta(days=1)).isoformat()},"
                f"{today.isoformat()},Apple Inc.,Purchase,"
                '"$1,001-$15,000",https://efdsearch.senate.gov/future\n',
                encoding="utf-8",
            )
            pipeline = DataIngestionPipeline(
                congress_trades_csv_path=str(csv_path)
            )

            rows, error = pipeline._load_senate_congress_rows("AAPL")

        self.assertIsNone(error)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["amount_range"], "$1,001-$15,000")


class HouseDisclosureMatchingTests(unittest.TestCase):
    def test_single_character_ticker_requires_issuer_name_match(self):
        self.assertIsNone(
            DataIngestionPipeline._match_house_report_asset(
                "V", ["Visa Inc."], "A member disclosed a transaction in V."
            )
        )
        self.assertEqual(
            DataIngestionPipeline._match_house_report_asset(
                "V", ["Visa Inc."], "Transaction involved Visa Inc."
            ),
            "visa inc.",
        )

    def test_multi_character_ticker_matches_as_a_standalone_symbol(self):
        self.assertEqual(
            DataIngestionPipeline._match_house_report_asset(
                "AAPL", ["Apple Inc."], "AAPL purchase was reported."
            ),
            "AAPL",
        )
        self.assertEqual(
            DataIngestionPipeline._match_house_report_asset(
                "AAPL", ["Apple Inc."], "Transaction involved Pineapple Inc."
            ),
            None,
        )


class OptionsPayloadTests(unittest.TestCase):
    def test_handles_missing_yahoo_option_chain_columns(self):
        class FakeTicker:
            options = ["2027-01-15"]

            def option_chain(self, _expiration):
                return SimpleNamespace(
                    calls=pd.DataFrame({"volume": [5]}),
                    puts=pd.DataFrame({"openInterest": [10]}),
                )

        payload = DataIngestionPipeline._build_options_payload(FakeTicker())

        self.assertEqual(payload["call_volume"], 5)
        self.assertEqual(payload["put_volume"], 0)
        self.assertEqual(payload["call_open_interest"], 0)
        self.assertEqual(payload["put_open_interest"], 10)
        self.assertNotIn("put_call_open_interest_ratio", payload)
        self.assertIsNone(payload["call_implied_volatility_mean_pct"])


class SourceStatusTests(unittest.TestCase):
    def test_unavailable_source_skips_directional_opinion(self):
        opinion = TickerCouncil._unavailable_source_opinion(
            "News_Events", "missing"
        )

        self.assertEqual(opinion.signal, Signal.NO_READ)
        self.assertEqual(opinion.confidence, 0.0)

    def test_partial_source_remains_eligible_for_analysis(self):
        opinion = TickerCouncil._unavailable_source_opinion(
            "Macro_Economy", "partial"
        )

        self.assertIsNone(opinion)

    def test_specialist_prompt_limits_claims_and_treats_macro_as_context(self):
        prompt = TickerCouncil._build_seat_prompt(
            {"name": "Macro_Economy", "desc": "macro context"},
            {"yield_curve_spread_10y_2y": 0.31},
        )

        self.assertIn("Mention only metrics explicitly present", prompt)
        self.assertIn("contextual macro information", prompt)
        self.assertIn('"yield_curve_spread_10y_2y": 0.31', prompt)

    def test_congress_prompt_warns_against_overclaiming_matches(self):
        prompt = TickerCouncil._build_seat_prompt(
            {"name": "Congress_Trades", "desc": "disclosure data"},
            {"house_events": []},
        )

        self.assertIn("best effort", prompt)
        self.assertIn("do not prove that no trade occurred", prompt)

    def test_institutional_prompt_distinguishes_ownership_from_flow(self):
        prompt = TickerCouncil._build_seat_prompt(
            {"name": "Institutional_Flows", "desc": "ownership and short interest"},
            {"held_percent_institutions_pct": 70.0},
        )

        self.assertIn("not evidence of recent buying or selling", prompt)
        self.assertIn("Do not claim access to 13F accumulations", prompt)


class FinancialPayloadTests(unittest.TestCase):
    def test_computes_free_cash_flow_yield_only_with_valid_inputs(self):
        payload = DataIngestionPipeline._build_financials_payload(
            {"freeCashflow": 10, "marketCap": 200}
        )

        self.assertEqual(payload["free_cash_flow_yield_pct"], 5.0)

    def test_missing_market_cap_does_not_create_free_cash_flow_yield(self):
        payload = DataIngestionPipeline._build_financials_payload(
            {"freeCashflow": 10, "marketCap": None}
        )

        self.assertIsNone(payload["free_cash_flow_yield_pct"])

    def test_converts_yahoo_ownership_and_short_interest_fractions_to_percent(self):
        payload = DataIngestionPipeline._build_institutional_payload(
            {
                "heldPercentInstitutions": 0.7,
                "heldPercentInsiders": 0.02,
                "shortPercentOfFloat": 0.04,
                "sharesShort": 1000,
            }
        )

        self.assertAlmostEqual(payload["held_percent_institutions_pct"], 70.0)
        self.assertAlmostEqual(payload["held_percent_insiders_pct"], 2.0)
        self.assertAlmostEqual(payload["short_percent_of_float_pct"], 4.0)
        self.assertEqual(payload["shares_short"], 1000)

    def test_loads_capital_iq_workbook_row_and_formula_values(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "report.xlsx"
            headers = [
                "ticker", "as_of_date", "target_mean", "target_high",
                "target_low", "number_of_analysts", "recommendation_key",
                "forward_eps", "free_cash_flow", "market_cap",
                "ebitda_margin_pct", "held_percent_institutions_pct",
            ]
            _write_capital_iq_workbook(
                workbook_path,
                headers,
                [
                    [
                        "AAPL",
                        ("=TEXT(TODAY(),\"yyyy-mm-dd\")", "2026-10-25"),
                        ("=CIQ(A2,\"IQ_PRICE_TARGET\")", 250),
                        280,
                        220,
                        30,
                        "BUY",
                        ("=CIQ(A2,\"IQ_EPS_EST\",IQ_NTM)", 9.25),
                        100_000_000,
                        ("=CIQ(A2,\"IQ_MARKETCAP\")", 2_000_000_000),
                        31.5,
                        64.2,
                    ],
                    [
                        "MSFT", "2026-10-25", 500, 550, 450, 40, "BUY",
                        20, 200_000_000, 3_000_000_000, 45, 70,
                    ],
                ],
            )
            pipeline = DataIngestionPipeline(capital_iq_excel_path=str(workbook_path))

            data, error, as_of_date = pipeline._load_capital_iq_excel("aapl")

            self.assertIsNone(error)
            self.assertEqual(as_of_date, "2026-10-25")
            self.assertEqual(data["Analyst_Targets"]["target_mean"], 250.0)
            self.assertEqual(data["Analyst_Targets"]["target_high"], 280.0)
            self.assertEqual(data["Earnings_Estimates"]["forward_eps"], 9.25)
            self.assertEqual(data["Financials"]["free_cash_flow_yield_pct"], 5.0)
            self.assertEqual(data["Financials"]["ebitda_margin_pct"], 31.5)
            self.assertEqual(
                data["Institutional_Flows"]["held_percent_institutions_pct"],
                64.2,
            )
            self.assertEqual(pipeline.capital_iq_tickers(), ["AAPL", "MSFT"])

    def test_reports_capital_iq_workbook_ticker_mismatch(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "report.xlsx"
            _write_capital_iq_workbook(
                workbook_path,
                ["ticker", "as_of_date", "target_mean"],
                [["MSFT", "2026-10-25", 500]],
            )
            pipeline = DataIngestionPipeline(capital_iq_excel_path=str(workbook_path))

            data, error, as_of_date = pipeline._load_capital_iq_excel("AAPL")

            self.assertEqual(data, {})
            self.assertIn("No Capital IQ workbook row", error)
            self.assertIsNone(as_of_date)

    def test_capital_iq_blank_fields_preserve_yahoo_values(self):
        async def run_test():
            with tempfile.TemporaryDirectory() as temp_dir:
                workbook_path = Path(temp_dir) / "report.xlsx"
                _write_capital_iq_workbook(
                    workbook_path,
                    [
                        "ticker",
                        "as_of_date",
                        "pe_ratio",
                        "ebitda_margin_pct",
                        "target_mean",
                    ],
                    [["AAPL", "2026-10-25", "", 31.5, 250]],
                )
                pipeline = DataIngestionPipeline(
                    capital_iq_excel_path=str(workbook_path)
                )
                yahoo_data = {
                    "Financials": {"pe_ratio": 20.0, "market_cap": 1000.0},
                    "Earnings_Estimates": {"forward_eps": 5.0},
                    "News_Events": {
                        "recent_headlines": [
                            {"title": "Test headline", "publisher": "Test", "link": ""}
                        ]
                    },
                }
                pipeline._fetch_yahoo_bundle = lambda ticker: (yahoo_data, {})

                async def fetch_macro(client):
                    return {}, None

                pipeline._fetch_fred_macro = fetch_macro

                async def fetch_congress(_ticker, _company_names, _client):
                    return (
                        {
                            "house_reports_scanned": 1,
                            "house_events": [],
                            "senate_events": [],
                        },
                        None,
                        False,
                    )

                pipeline._fetch_congress_trades = fetch_congress
                data = await pipeline.fetch_all("AAPL")

                self.assertEqual(data["Financials"]["pe_ratio"], 20.0)
                self.assertEqual(data["Financials"]["ebitda_margin_pct"], 31.5)
                self.assertEqual(data["Analyst_Targets"]["target_mean"], 250.0)
                self.assertEqual(
                    data["_metadata"]["sources"]["Financials"]["as_of_date"],
                    "2026-10-25",
                )

        asyncio.run(run_test())


class CouncilSourceRoutingTests(unittest.IsolatedAsyncioTestCase):
    async def test_unavailable_seats_skip_model_call_and_partial_seat_runs(self):
        council = TickerCouncil(api_key="test-key")
        unavailable_seats = {
            seat["name"]: {"status": "not_configured"}
            for seat in SEAT_CONFIGS
        }
        unavailable_seats["Earnings_Estimates"] = {"status": "partial"}
        payload = {
            "_metadata": {"sources": unavailable_seats},
            "Earnings_Estimates": {"forward_eps": 3.0},
        }
        called_seats = []

        async def execute_seat(seat, raw_data):
            called_seats.append(seat["name"])
            return SeatOpinion(
                seat_name=seat["name"],
                signal=Signal.EVEN,
                confidence=0.5,
                summary="Partial data analyzed.",
            )

        async def execute_debate(opinions):
            return DebateArguments(
                bull_case="bull",
                bear_case="bear",
                challenger_critique="challenge",
            )

        async def execute_grand_master(ticker, opinions, debate):
            return CouncilVerdict(
                ticker=ticker,
                forecasts=[
                    HorizonForecast(
                        horizon="1_week",
                        direction=Signal.UP,
                        probability_up=0.6,
                        reasoning="week",
                    ),
                    HorizonForecast(
                        horizon="3_months",
                        direction=Signal.EVEN,
                        probability_up=0.5,
                        reasoning="quarter",
                    ),
                    HorizonForecast(
                        horizon="1_year",
                        direction=Signal.DOWN,
                        probability_up=0.4,
                        reasoning="year",
                    ),
                ],
                seat_breakdown=opinions,
                debate=debate,
            )

        council._execute_seat = execute_seat
        council._execute_debate = execute_debate
        council._execute_grand_master = execute_grand_master
        verdict = await council.run("AAPL", payload)

        self.assertEqual(called_seats, ["Earnings_Estimates"])
        self.assertEqual(len(verdict.seat_breakdown), len(SEAT_CONFIGS))
        self.assertEqual(
            next(
                opinion.signal
                for opinion in verdict.seat_breakdown
                if opinion.seat_name == "News_Events"
            ),
            Signal.NO_READ,
        )


class ForecastDirectionTests(unittest.TestCase):
    def test_direction_threshold_boundaries(self):
        for probability, expected in (
            (0.449, Signal.DOWN),
            (0.45, Signal.EVEN),
            (0.55, Signal.EVEN),
            (0.551, Signal.UP),
        ):
            with self.subTest(probability=probability):
                forecast = HorizonForecast(
                    horizon="1_week",
                    direction=expected,
                    probability_up=probability,
                    reasoning="boundary test",
                )
                self.assertEqual(forecast.direction, expected)

    def test_rejects_direction_inconsistent_with_probability(self):
        with self.assertRaises(ValidationError):
            HorizonForecast(
                horizon="1_year",
                direction=Signal.EVEN,
                probability_up=0.56,
                reasoning="inconsistent test",
            )

    def test_verdict_requires_one_forecast_for_each_horizon(self):
        with self.assertRaises(ValidationError):
            CouncilVerdict(
                ticker="AAPL",
                forecasts=[
                    HorizonForecast(
                        horizon="1_week",
                        direction=Signal.UP,
                        probability_up=0.8,
                        reasoning="weekly",
                    )
                ],
                seat_breakdown=[],
                debate=DebateArguments(
                    bull_case="bull",
                    bear_case="bear",
                    challenger_critique="challenge",
                ),
            )


class FredPartialFailureTests(unittest.IsolatedAsyncioTestCase):
    async def test_one_failed_series_does_not_discard_successful_series(self):
        class Response:
            def __init__(self, value):
                self.value = value

            def raise_for_status(self):
                return None

            def json(self):
                return {"observations": [{"value": self.value}]}

        class Client:
            async def get(self, url, params, timeout):
                if params["series_id"] == "DGS2":
                    raise TimeoutError("simulated timeout")
                return Response("4.25")

        pipeline = DataIngestionPipeline(fred_api_key="test-key")
        data, error = await pipeline._fetch_fred_macro(Client())

        self.assertEqual(data["yield_10y"], 4.25)
        self.assertNotIn("yield_2y", data)
        self.assertIn("DGS2", error)


class ForecastStoreTests(unittest.TestCase):
    def test_records_and_scores_a_forecast_with_brier_score(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            store = ForecastStore(Path(temp_dir) / "forecasts.sqlite3")
            verdict = CouncilVerdict(
                ticker="AAPL",
                forecasts=[
                    HorizonForecast(
                        horizon="1_week",
                        direction=Signal.UP,
                        probability_up=0.8,
                        reasoning="test forecast",
                    ),
                    HorizonForecast(
                        horizon="3_months",
                        direction=Signal.EVEN,
                        probability_up=0.5,
                        reasoning="test forecast",
                    ),
                    HorizonForecast(
                        horizon="1_year",
                        direction=Signal.DOWN,
                        probability_up=0.4,
                        reasoning="test forecast",
                    ),
                ],
                seat_breakdown=[],
                debate=DebateArguments(
                    bull_case="bull",
                    bear_case="bear",
                    challenger_critique="challenge",
                ),
            )
            created_at = datetime(2025, 1, 1, tzinfo=timezone.utc)
            inserted = store.record_verdict(verdict, 100.0, created_at)

            self.assertEqual(inserted, 3)
            summary = store.score_due_forecasts(
                lambda ticker, start, end: [(date(2025, 1, 8), 110.0)],
                as_of_date=date(2025, 1, 8),
            )

            self.assertEqual(summary["due"], 1)
            self.assertEqual(summary["scored"], 1)
            self.assertAlmostEqual(summary["average_brier_score"], 0.04)

    def test_does_not_score_without_a_close_in_allowed_window(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            store = ForecastStore(Path(temp_dir) / "forecasts.sqlite3")
            verdict = CouncilVerdict(
                ticker="AAPL",
                forecasts=[
                    HorizonForecast(
                        horizon="1_week",
                        direction=Signal.DOWN,
                        probability_up=0.2,
                        reasoning="test forecast",
                    ),
                    HorizonForecast(
                        horizon="3_months",
                        direction=Signal.EVEN,
                        probability_up=0.5,
                        reasoning="test forecast",
                    ),
                    HorizonForecast(
                        horizon="1_year",
                        direction=Signal.UP,
                        probability_up=0.6,
                        reasoning="test forecast",
                    ),
                ],
                seat_breakdown=[],
                debate=DebateArguments(
                    bull_case="bull",
                    bear_case="bear",
                    challenger_critique="challenge",
                ),
            )
            store.record_verdict(
                verdict, 100.0, datetime(2025, 1, 1, tzinfo=timezone.utc)
            )

            summary = store.score_due_forecasts(
                lambda ticker, start, end: [(date(2025, 1, 20), 90.0)],
                as_of_date=date(2025, 1, 20),
            )

            self.assertEqual(summary["due"], 1)
            self.assertEqual(summary["scored"], 0)
            self.assertIsNone(summary["average_brier_score"])


class BatchTickerTests(unittest.IsolatedAsyncioTestCase):
    def test_prompt_accepts_all_tickers_option(self):
        with patch("builtins.input", return_value="all"):
            self.assertEqual(prompt_ticker(["AAPL", "MSFT"]), "ALL")

    async def test_batch_continues_after_ticker_failure(self):
        calls = []

        def make_verdict(ticker):
            return CouncilVerdict(
                ticker=ticker,
                forecasts=[
                    HorizonForecast(
                        horizon="1_week",
                        direction=Signal.UP,
                        probability_up=0.6,
                        reasoning="Positive short-term momentum.",
                    ),
                    HorizonForecast(
                        horizon="3_months",
                        direction=Signal.EVEN,
                        probability_up=0.5,
                        reasoning="Mixed medium-term factors.",
                    ),
                    HorizonForecast(
                        horizon="1_year",
                        direction=Signal.DOWN,
                        probability_up=0.4,
                        reasoning="Long-term risks remain.",
                    ),
                ],
                seat_breakdown=[
                    SeatOpinion(
                        seat_name="Financials",
                        signal=Signal.EVEN,
                        confidence=0.5,
                        summary="Financial metrics are mixed.",
                    )
                ],
                debate=DebateArguments(
                    bull_case="bull case details",
                    bear_case="bear case details",
                    challenger_critique="challenger critique details",
                ),
            )

        async def fake_run_pipeline(ticker, ingestor):
            calls.append(ticker)
            if ticker == "MSFT":
                raise RuntimeError("simulated failure")
            return make_verdict(ticker)

        with tempfile.TemporaryDirectory() as temp_dir:
            with patch("main.run_pipeline", side_effect=fake_run_pipeline):
                failures = await run_all_tickers(
                    ["AAPL", "MSFT", "NVDA"],
                    report_directory=Path(temp_dir),
                )

            self.assertEqual(calls, ["AAPL", "MSFT", "NVDA"])
            self.assertEqual(len(failures), 1)
            self.assertEqual(failures[0][0], "MSFT")
            self.assertIn("simulated failure", failures[0][1])

            report_paths = list(Path(temp_dir).glob("equity_research_report_*.docx"))
            self.assertEqual(len(report_paths), 1)
            with ZipFile(report_paths[0]) as report:
                document_xml = ET.fromstring(report.read("word/document.xml"))
            report_text = " ".join(
                text.text or ""
                for text in document_xml.iter(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
                )
            )

            self.assertIn("Forecast Summary", report_text)
            self.assertIn("bull case details", report_text)
            self.assertIn("MSFT", report_text)
            self.assertIn("simulated failure", report_text)


class SingleTickerReportTests(unittest.TestCase):
    def test_writes_single_ticker_word_report_to_requested_directory(self):
        verdict = CouncilVerdict(
            ticker="AAPL",
            forecasts=[
                HorizonForecast(
                    horizon="1_week",
                    direction=Signal.UP,
                    probability_up=0.6,
                    reasoning="Positive short-term momentum.",
                ),
                HorizonForecast(
                    horizon="3_months",
                    direction=Signal.EVEN,
                    probability_up=0.5,
                    reasoning="Mixed medium-term factors.",
                ),
                HorizonForecast(
                    horizon="1_year",
                    direction=Signal.DOWN,
                    probability_up=0.4,
                    reasoning="Long-term risks remain.",
                ),
            ],
            seat_breakdown=[
                SeatOpinion(
                    seat_name="Financials",
                    signal=Signal.EVEN,
                    confidence=0.5,
                    summary="Financial metrics are mixed.",
                )
            ],
            debate=DebateArguments(
                bull_case="bull case details",
                bear_case="bear case details",
                challenger_critique="challenger critique details",
            ),
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            report_path = write_ticker_report(verdict, Path(temp_dir))

            self.assertTrue(report_path.is_file())
            self.assertEqual(report_path.parent, Path(temp_dir))
            self.assertIn("AAPL", report_path.name)
            with ZipFile(report_path) as report:
                document_xml = ET.fromstring(report.read("word/document.xml"))
            report_text = " ".join(
                text.text or ""
                for text in document_xml.iter(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
                )
            )

        self.assertIn("Equity Research Report", report_text)
        self.assertIn("AAPL", report_text)
        self.assertIn("Positive short-term momentum.", report_text)
        self.assertIn("bull case details", report_text)


if __name__ == "__main__":
    unittest.main()
