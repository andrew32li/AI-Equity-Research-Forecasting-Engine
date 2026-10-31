"""
ingestion.py - Robust Asynchronous Data Ingestion for Ticker Council
"""

import asyncio
import csv
from datetime import datetime, timedelta, timezone
from html.parser import HTMLParser
from io import BytesIO
import math
import os
import posixpath
from pathlib import Path
import re
import time
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote_plus
from zipfile import BadZipFile, ZipFile
import xml.etree.ElementTree as ET

import httpx
import pandas as pd
import yfinance as yf
from pypdf import PdfReader
from pypdf.errors import PdfReadError


SEC_TICKER_MAP_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"
SEC_ARCHIVES_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{accession}/{document}"
SEC_LOOKBACK_DAYS = 90
SEC_FILINGS_LOOKBACK_DAYS = 365
HOUSE_PTR_MAX_REPORTS = 75
SECTOR_ETFS = {
    "communication services": "XLC",
    "consumer cyclical": "XLY",
    "consumer defensive": "XLP",
    "energy": "XLE",
    "financial services": "XLF",
    "healthcare": "XLV",
    "industrials": "XLI",
    "real estate": "XLRE",
    "technology": "XLK",
    "basic materials": "XLB",
    "utilities": "XLU",
}


class _VisibleTextParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self._hidden_depth = 0
        self.parts: List[str] = []

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        if tag.lower() in {"script", "style", "noscript"}:
            self._hidden_depth += 1

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"script", "style", "noscript"} and self._hidden_depth:
            self._hidden_depth -= 1

    def handle_data(self, data: str) -> None:
        if not self._hidden_depth:
            cleaned = " ".join(data.split())
            if cleaned:
                self.parts.append(cleaned)


class DataIngestionPipeline:
    def __init__(
        self,
        fred_api_key: Optional[str] = None,
        sec_user_agent: Optional[str] = None,
        capital_iq_excel_path: Optional[str] = None,
        congress_trades_csv_path: Optional[str] = None,
    ):
        self.fred_api_key = fred_api_key or os.environ.get("FRED_API_KEY")
        contact_email = os.environ.get("SEC_CONTACT_EMAIL", "").strip()
        configured_user_agent = (
            sec_user_agent
            or os.environ.get("SEC_USER_AGENT")
            or (
                f"StockEvaluatorEngine contact {contact_email}"
                if contact_email
                else None
            )
        )
        self.sec_headers = (
            {
                "User-Agent": configured_user_agent,
                "Accept-Encoding": "gzip, deflate",
            }
            if configured_user_agent
            else {}
        )
        configured_excel_path = (
            capital_iq_excel_path
            or os.environ.get("CAPITAL_IQ_XLSX_PATH")
            or Path(__file__).resolve().with_name("report.xlsx")
        )
        self.capital_iq_excel_path = Path(configured_excel_path)
        configured_congress_path = (
            congress_trades_csv_path
            or os.environ.get("CONGRESS_TRADES_CSV_PATH")
            or Path(__file__).resolve().with_name("congress_trades.csv")
        )
        self.congress_trades_csv_path = Path(configured_congress_path)
        self._sec_ticker_map: Optional[Dict[str, Dict[str, Any]]] = None
        self._sec_request_lock = asyncio.Lock()
        self._sec_last_request = 0.0
        self._house_ptr_cache: Optional[
            Tuple[List[Dict[str, Any]], Optional[str], bool]
        ] = None

    @staticmethod
    def _parse_excel_number(value: Any, field_name: str) -> Optional[float]:
        if value is None or (isinstance(value, str) and not value.strip()):
            return None
        normalized = str(value).strip().replace(",", "").replace("$", "").replace("%", "")
        try:
            return float(normalized)
        except ValueError as exc:
            raise ValueError(f"Invalid numeric value for {field_name}.") from exc

    def _read_capital_iq_excel_rows(self) -> List[Dict[str, Any]]:
        if not self.capital_iq_excel_path.is_file():
            raise FileNotFoundError(
                f"Capital IQ workbook not found: {self.capital_iq_excel_path}"
            )

        spreadsheet_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        document_rel_ns = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        )
        package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        namespaces = {"m": spreadsheet_ns, "r": document_rel_ns, "p": package_rel_ns}

        with ZipFile(self.capital_iq_excel_path) as workbook:
            workbook_xml = ET.fromstring(workbook.read("xl/workbook.xml"))
            relationships_xml = ET.fromstring(
                workbook.read("xl/_rels/workbook.xml.rels")
            )
            relationships = {
                relationship.attrib["Id"]: relationship.attrib["Target"]
                for relationship in relationships_xml.findall("p:Relationship", namespaces)
            }

            report_sheet = next(
                (
                    sheet
                    for sheet in workbook_xml.findall("m:sheets/m:sheet", namespaces)
                    if sheet.attrib.get("name", "").strip().casefold() == "report"
                ),
                None,
            )
            if report_sheet is None:
                raise ValueError("Capital IQ workbook must contain a 'Report' worksheet.")

            relationship_id = report_sheet.attrib.get(f"{{{document_rel_ns}}}id")
            worksheet_target = relationships.get(relationship_id or "")
            if worksheet_target is None:
                raise ValueError("Could not locate the 'Report' worksheet in the workbook.")
            worksheet_path = (
                worksheet_target.lstrip("/")
                if worksheet_target.startswith("/")
                else posixpath.normpath(posixpath.join("xl", worksheet_target))
            )
            worksheet_xml = ET.fromstring(workbook.read(worksheet_path))

            shared_strings: List[str] = []
            if "xl/sharedStrings.xml" in workbook.namelist():
                shared_strings_xml = ET.fromstring(
                    workbook.read("xl/sharedStrings.xml")
                )
                shared_strings = [
                    "".join(text.text or "" for text in item.findall(".//m:t", namespaces))
                    for item in shared_strings_xml.findall("m:si", namespaces)
                ]

            def column_index(cell_reference: str) -> int:
                match = re.match(r"([A-Z]+)", cell_reference.upper())
                if match is None:
                    raise ValueError(f"Invalid Excel cell reference: {cell_reference}")
                index = 0
                for character in match.group(1):
                    index = index * 26 + ord(character) - ord("A") + 1
                return index - 1

            def cell_value(cell: ET.Element) -> Any:
                cell_type = cell.attrib.get("t")
                if cell_type == "e":
                    return (
                        f"#EXCEL_ERROR:{cell.attrib.get('r')}:"
                        f"{cell.findtext('m:v', default='', namespaces=namespaces)}"
                    )
                if cell_type == "inlineStr":
                    return "".join(
                        text.text or ""
                        for text in cell.findall(".//m:t", namespaces)
                    ).strip()

                value = cell.findtext("m:v", default="", namespaces=namespaces)
                if cell_type == "s":
                    if not value:
                        return ""
                    try:
                        return shared_strings[int(value)].strip()
                    except (IndexError, ValueError) as exc:
                        raise ValueError(
                            f"Invalid shared-string index in cell "
                            f"{cell.attrib.get('r')}."
                        ) from exc
                return value.strip() if isinstance(value, str) else value

            rows_by_number: Dict[int, Dict[int, Any]] = {}
            for row in worksheet_xml.findall(".//m:sheetData/m:row", namespaces):
                row_number = int(row.attrib["r"])
                cells: Dict[int, Any] = {}
                for cell in row.findall("m:c", namespaces):
                    cells[column_index(cell.attrib["r"])] = cell_value(cell)
                rows_by_number[row_number] = cells

        header_cells = rows_by_number.get(1, {})
        headers = {
            str(value).strip().lower(): index
            for index, value in header_cells.items()
            if value not in (None, "")
        }
        if "ticker" not in headers:
            raise ValueError("The 'Report' worksheet must have a ticker header in row 1.")

        rows: List[Dict[str, Any]] = []
        for row_number, cells in rows_by_number.items():
            if row_number == 1:
                continue
            row = {
                header: cells.get(column_index_value, "")
                for header, column_index_value in headers.items()
            }
            if str(row.get("ticker", "")).strip():
                rows.append(row)
        return rows

    def capital_iq_tickers(self) -> List[str]:
        """Return non-empty tickers listed in the workbook's Report worksheet."""
        try:
            return [
                str(row["ticker"]).strip().upper()
                for row in self._read_capital_iq_excel_rows()
                if str(row.get("ticker", "")).strip()
            ]
        except (OSError, BadZipFile, ET.ParseError, KeyError, ValueError) as exc:
            raise ValueError(f"Could not read Capital IQ ticker list: {exc}") from exc

    def _load_capital_iq_excel(
        self, ticker: str
    ) -> Tuple[Dict[str, Dict[str, Any]], Optional[str], Optional[str]]:
        try:
            matching_rows = [
                row
                for row in self._read_capital_iq_excel_rows()
                if str(row.get("ticker", "")).strip().upper() == ticker.upper()
            ]

            if not matching_rows:
                return {}, f"No Capital IQ workbook row found for {ticker.upper()}.", None
            if len(matching_rows) > 1:
                return {}, f"Multiple Capital IQ workbook rows found for {ticker.upper()}.", None

            row = matching_rows[0]
            as_of_date = str(row.get("as_of_date", "")).strip()
            if not as_of_date:
                return {}, "Capital IQ workbook row is missing as_of_date.", None
            try:
                datetime.strptime(as_of_date, "%Y-%m-%d")
            except ValueError:
                return {}, "Capital IQ as_of_date must use YYYY-MM-DD format.", None

            numeric_fields = {
                "target_mean": ("Analyst_Targets", "target_mean"),
                "target_high": ("Analyst_Targets", "target_high"),
                "target_low": ("Analyst_Targets", "target_low"),
                "number_of_analysts": ("Analyst_Targets", "number_of_analysts"),
                "forward_eps": ("Earnings_Estimates", "forward_eps"),
                "forward_revenue": ("Earnings_Estimates", "forward_revenue"),
                "eps_growth_pct": ("Earnings_Estimates", "eps_growth_pct"),
                "revenue_growth_pct": ("Earnings_Estimates", "revenue_growth_pct"),
                "pe_ratio": ("Financials", "pe_ratio"),
                "price_to_book": ("Financials", "price_to_book"),
                "free_cash_flow": ("Financials", "free_cash_flow"),
                "market_cap": ("Financials", "market_cap"),
                "total_debt": ("Financials", "total_debt"),
                "cash_and_equivalents": ("Financials", "cash_and_equivalents"),
                "operating_margins_pct": ("Financials", "operating_margins_pct"),
                "ebitda_margin_pct": ("Financials", "ebitda_margin_pct"),
                "institutional_ownership_pct": (
                    "Institutional_Flows",
                    "held_percent_institutions_pct",
                ),
                "held_percent_institutions_pct": (
                    "Institutional_Flows",
                    "held_percent_institutions_pct",
                ),
            }
            result: Dict[str, Dict[str, Any]] = {}
            for excel_field, (seat_name, payload_field) in numeric_fields.items():
                raw_value = row.get(excel_field)
                if isinstance(raw_value, str) and raw_value.startswith("#EXCEL_ERROR:"):
                    raise ValueError(
                        f"Capital IQ formula error in {raw_value.split(':', 2)[1]}: "
                        f"{raw_value.split(':', 2)[2]}"
                    )
                value = self._parse_excel_number(raw_value, excel_field)
                if value is not None:
                    if excel_field == "number_of_analysts":
                        if not value.is_integer() or value < 0:
                            raise ValueError("number_of_analysts must be a non-negative integer.")
                        value = int(value)
                    result.setdefault(seat_name, {})[payload_field] = value

            recommendation = row.get("recommendation_key")
            if recommendation and str(recommendation).strip():
                result.setdefault("Analyst_Targets", {})[
                    "recommendation_key"
                ] = str(recommendation).strip()

            for seat_data in result.values():
                seat_data["as_of_date"] = as_of_date

            financials = result.get("Financials", {})
            free_cash_flow = financials.get("free_cash_flow")
            market_cap = financials.get("market_cap")
            if free_cash_flow is not None and market_cap is not None and market_cap > 0:
                financials["free_cash_flow_yield_pct"] = round(
                    (free_cash_flow / market_cap) * 100, 2
                )

            return result, None, as_of_date
        except (OSError, BadZipFile, ET.ParseError, KeyError, ValueError) as exc:
            return {}, f"Capital IQ workbook could not be read: {exc}", None

    @staticmethod
    def _calculate_series_metrics(closes: List[float], volumes: List[int]) -> Dict[str, Any]:
        if not closes or len(closes) < 14:
            return {}

        current = closes[-1]
        sma_20 = sum(closes[-20:]) / min(len(closes), 20)
        sma_50 = sum(closes[-50:]) / min(len(closes), 50)
        sma_200 = sum(closes[-200:]) / min(len(closes), 200) if len(closes) >= 200 else None

        returns = [
            math.log(closes[i] / closes[i - 1])
            for i in range(max(1, len(closes) - 30), len(closes))
        ]
        if len(returns) > 1:
            mean_ret = sum(returns) / len(returns)
            variance = sum((r - mean_ret) ** 2 for r in returns) / (len(returns) - 1)
            vol_30d = math.sqrt(variance * 252) * 100
        else:
            vol_30d = 0.0

        return {
            "current_price": round(current, 2),
            "sma_20": round(sma_20, 2),
            "sma_50": round(sma_50, 2),
            "sma_200": round(sma_200, 2) if sma_200 else None,
            "pct_above_sma50": round(((current - sma_50) / sma_50) * 100, 2),
            "realized_volatility_30d_pct": round(vol_30d, 2),
            "avg_volume_20d": int(sum(volumes[-20:]) / min(len(volumes), 20)) if volumes else 0,
        }

    @staticmethod
    def _build_financials_payload(info: Dict[str, Any]) -> Dict[str, Any]:
        free_cash_flow = info.get("freeCashflow")
        market_cap = info.get("marketCap")
        has_fcf_yield_inputs = (
            isinstance(free_cash_flow, (int, float))
            and isinstance(market_cap, (int, float))
            and market_cap > 0
        )
        return {
            "pe_ratio": info.get("trailingPE") or info.get("forwardPE"),
            "price_to_book": info.get("priceToBook"),
            "free_cash_flow": free_cash_flow,
            "market_cap": market_cap,
            "free_cash_flow_yield_pct": (
                round((free_cash_flow / market_cap) * 100, 2)
                if has_fcf_yield_inputs
                else None
            ),
            "total_debt": info.get("totalDebt"),
            "cash_and_equivalents": info.get("totalCash"),
            "operating_margins": info.get("operatingMargins"),
        }

    @staticmethod
    def _build_institutional_payload(info: Dict[str, Any]) -> Dict[str, Any]:
        fields = {
            "held_percent_institutions_pct": (
                info.get("heldPercentInstitutions") * 100
                if isinstance(info.get("heldPercentInstitutions"), (int, float))
                else None
            ),
            "held_percent_insiders_pct": (
                info.get("heldPercentInsiders") * 100
                if isinstance(info.get("heldPercentInsiders"), (int, float))
                else None
            ),
            "short_percent_of_float_pct": (
                info.get("shortPercentOfFloat") * 100
                if isinstance(info.get("shortPercentOfFloat"), (int, float))
                else None
            ),
            "shares_short": info.get("sharesShort"),
            "shares_short_prior_month": info.get("sharesShortPriorMonth"),
            "short_ratio_days_to_cover": info.get("shortRatio"),
            "short_interest_as_of": info.get("dateShortInterest"),
        }
        return {key: value for key, value in fields.items() if value is not None}

    @staticmethod
    def _return_over_period(series: pd.Series, trading_days: int) -> Optional[float]:
        prices = series.dropna()
        if len(prices) <= trading_days or prices.iloc[-trading_days - 1] <= 0:
            return None
        return round(
            (float(prices.iloc[-1]) / float(prices.iloc[-trading_days - 1]) - 1) * 100,
            2,
        )

    @staticmethod
    def _visible_html_text(content: str) -> str:
        parser = _VisibleTextParser()
        parser.feed(content)
        return " ".join(parser.parts)

    async def _sec_get(self, client: httpx.AsyncClient, url: str) -> httpx.Response:
        if not self.sec_headers:
            raise RuntimeError(
                "Set SEC_CONTACT_EMAIL or SEC_USER_AGENT before requesting SEC data."
            )
        async with self._sec_request_lock:
            wait_seconds = 0.12 - (time.monotonic() - self._sec_last_request)
            if wait_seconds > 0:
                await asyncio.sleep(wait_seconds)
            response = await client.get(
                url,
                headers=self.sec_headers,
                timeout=20.0,
            )
            self._sec_last_request = time.monotonic()
        response.raise_for_status()
        return response

    async def _load_sec_ticker_map(
        self, client: httpx.AsyncClient
    ) -> Dict[str, Dict[str, Any]]:
        if self._sec_ticker_map is None:
            response = await self._sec_get(client, SEC_TICKER_MAP_URL)
            raw_mapping = response.json()
            self._sec_ticker_map = {
                str(record.get("ticker", "")).strip().upper(): record
                for record in raw_mapping.values()
                if record.get("ticker")
            }
        return self._sec_ticker_map

    @staticmethod
    def _xml_text(element: ET.Element, *path: str) -> Optional[str]:
        current = element
        for segment in path:
            match = next(
                (
                    child for child in current.iter()
                    if child is not current and child.tag.rsplit("}", 1)[-1] == segment
                ),
                None,
            )
            if match is None:
                return None
            current = match
        return current.text.strip() if current.text and current.text.strip() else None

    @classmethod
    def _parse_form4_xml(
        cls, content: bytes, filed_date: str, filing_url: str
    ) -> List[Dict[str, Any]]:
        root = ET.fromstring(content)
        owner_names = [
            value.text.strip()
            for value in root.iter()
            if value.tag.rsplit("}", 1)[-1] == "rptOwnerName"
            and value.text
            and value.text.strip()
        ]
        transactions = []
        for transaction in root.iter():
            if transaction.tag.rsplit("}", 1)[-1] != "nonDerivativeTransaction":
                continue
            code = cls._xml_text(transaction, "transactionCoding", "transactionCode")
            if code not in {"P", "S"}:
                continue
            transaction_date = cls._xml_text(
                transaction, "transactionDate", "value"
            )
            shares_text = cls._xml_text(
                transaction, "transactionAmounts", "transactionShares", "value"
            )
            price_text = cls._xml_text(
                transaction, "transactionAmounts",
                "transactionPricePerShare", "value"
            )
            try:
                shares = float(shares_text) if shares_text else None
                price = float(price_text) if price_text else None
            except ValueError:
                shares = price = None
            transactions.append(
                {
                    "type": "PURCHASE" if code == "P" else "SALE",
                    "transaction_code": code,
                    "date": transaction_date,
                    "shares": shares,
                    "price_per_share": price,
                    "value_usd": round(shares * price, 2)
                    if shares is not None and price is not None
                    else None,
                    "insider": ", ".join(owner_names) or None,
                    "filed_date": filed_date,
                    "source_url": filing_url,
                }
            )
        return transactions

    @staticmethod
    def _parse_sec_excerpt(form: str, html: str) -> str:
        text = DataIngestionPipeline._visible_html_text(html)
        if form in {"10-K", "10-Q"}:
            for heading in ("ITEM 1A. RISK FACTORS", "ITEM 1A RISK FACTORS"):
                match = re.search(re.escape(heading), text, flags=re.IGNORECASE)
                if match:
                    return text[match.start():match.start() + 9000]
        return text[:9000]

    @staticmethod
    def _build_options_payload(ticker: yf.Ticker) -> Dict[str, Any]:
        expirations = list(ticker.options or [])
        if not expirations:
            raise ValueError("Yahoo Finance returned no listed options expirations.")

        calls_volume = puts_volume = calls_open_interest = puts_open_interest = 0.0
        call_implied_volatility: List[float] = []
        put_implied_volatility: List[float] = []
        used_expirations = []
        for expiration in expirations[:3]:
            chain = ticker.option_chain(expiration)
            used_expirations.append(expiration)
            for frame, is_call, iv_values in (
                (chain.calls, True, call_implied_volatility),
                (chain.puts, False, put_implied_volatility),
            ):
                if frame.empty:
                    continue
                volume = (
                    pd.to_numeric(frame["volume"], errors="coerce").fillna(0)
                    if "volume" in frame
                    else pd.Series(0, index=frame.index, dtype=float)
                )
                interest = (
                    pd.to_numeric(frame["openInterest"], errors="coerce").fillna(0)
                    if "openInterest" in frame
                    else pd.Series(0, index=frame.index, dtype=float)
                )
                implied_volatility = (
                    pd.to_numeric(
                        frame["impliedVolatility"], errors="coerce"
                    ).dropna()
                    if "impliedVolatility" in frame
                    else pd.Series(dtype=float)
                )
                if is_call:
                    calls_volume += float(volume.sum())
                    calls_open_interest += float(interest.sum())
                else:
                    puts_volume += float(volume.sum())
                    puts_open_interest += float(interest.sum())
                iv_values.extend(
                    float(value) for value in implied_volatility if value > 0
                )

        nearest_expiration = datetime.strptime(used_expirations[0], "%Y-%m-%d").date()
        payload: Dict[str, Any] = {
            "source": "Yahoo Finance options-chain snapshot (unofficial)",
            "retrieved_at": datetime.now(timezone.utc).isoformat(),
            "expiration_dates": used_expirations,
            "nearest_expiration_days": (
                nearest_expiration - datetime.now(timezone.utc).date()
            ).days,
            "call_volume": int(calls_volume),
            "put_volume": int(puts_volume),
            "call_open_interest": int(calls_open_interest),
            "put_open_interest": int(puts_open_interest),
            "call_implied_volatility_mean_pct": (
                round(sum(call_implied_volatility) / len(call_implied_volatility) * 100, 2)
                if call_implied_volatility
                else None
            ),
            "put_implied_volatility_mean_pct": (
                round(sum(put_implied_volatility) / len(put_implied_volatility) * 100, 2)
                if put_implied_volatility
                else None
            ),
            "interpretation_limit": (
                "Delayed chain snapshot; volume/open interest do not identify trade direction "
                "or prove unusual options flow."
            ),
        }
        if calls_volume > 0:
            payload["put_call_volume_ratio"] = round(puts_volume / calls_volume, 3)
        if calls_open_interest > 0:
            payload["put_call_open_interest_ratio"] = round(
                puts_open_interest / calls_open_interest, 3
            )
        return payload

    @staticmethod
    def _build_related_markets_payload(
        ticker: str,
        info: Dict[str, Any],
        stock_history: pd.DataFrame,
        spy_history: pd.DataFrame,
    ) -> Tuple[Dict[str, Any], Optional[str]]:
        sector = str(info.get("sector") or "").strip()
        benchmark = SECTOR_ETFS.get(sector.casefold())
        payload: Dict[str, Any] = {
            "classification": sector or None,
            "company_name": (
                info.get("longName") or info.get("shortName") or None
            ),
            "classification_source": "Yahoo Finance company profile",
            "market_benchmark": "SPY",
            "lookback": "1 year of daily prices",
            "source": "Yahoo Finance; sector ETF is a proxy, not an official index return.",
        }
        if not stock_history.empty and not spy_history.empty:
            aligned = pd.concat(
                [
                    stock_history["Close"].rename("stock"),
                    spy_history["Close"].rename("market"),
                ],
                axis=1,
            ).dropna()
            for period_name, trading_days in (("1m", 21), ("3m", 63), ("1y", 252)):
                stock_return = DataIngestionPipeline._return_over_period(
                    aligned["stock"], trading_days
                )
                market_return = DataIngestionPipeline._return_over_period(
                    aligned["market"], trading_days
                )
                if stock_return is not None:
                    payload[f"stock_return_{period_name}_pct"] = stock_return
                if market_return is not None:
                    payload[f"spy_return_{period_name}_pct"] = market_return
                if stock_return is not None and market_return is not None:
                    payload[f"excess_return_vs_spy_{period_name}_pct"] = round(
                        stock_return - market_return, 2
                    )

        error = None
        if benchmark:
            try:
                sector_history = yf.Ticker(benchmark).history(period="1y")
                if not sector_history.empty:
                    payload["sector_etf_proxy"] = benchmark
                    for period_name, trading_days in (("1m", 21), ("3m", 63), ("1y", 252)):
                        sector_return = DataIngestionPipeline._return_over_period(
                            sector_history["Close"], trading_days
                        )
                        if sector_return is not None:
                            payload[f"sector_return_{period_name}_pct"] = sector_return
                    if not stock_history.empty:
                        aligned = pd.concat(
                            [
                                stock_history["Close"].rename("stock"),
                                sector_history["Close"].rename("sector"),
                            ],
                            axis=1,
                        ).dropna()
                        for period_name, trading_days in (("1m", 21), ("3m", 63)):
                            stock_return = DataIngestionPipeline._return_over_period(
                                aligned["stock"], trading_days
                            )
                            sector_return = DataIngestionPipeline._return_over_period(
                                aligned["sector"], trading_days
                            )
                            if stock_return is not None and sector_return is not None:
                                payload[
                                    f"excess_return_vs_sector_{period_name}_pct"
                                ] = round(stock_return - sector_return, 2)
                else:
                    error = f"No daily history returned for sector ETF proxy {benchmark}."
            except Exception as exc:
                error = f"{benchmark}: {type(exc).__name__}: {exc}"
        elif not sector:
            error = "Yahoo Finance did not provide a sector classification."

        if not any(key.startswith("stock_return_") for key in payload):
            error = error or f"No aligned market history available for {ticker.upper()}."
        return payload, error

    async def _fetch_sec_company_data(
        self, ticker: str, client: httpx.AsyncClient
    ) -> Tuple[Dict[str, Any], Dict[str, Any], Optional[str]]:
        ticker_record = (await self._load_sec_ticker_map(client)).get(ticker.upper())
        if not ticker_record:
            raise ValueError(f"Ticker {ticker.upper()} is not in the SEC ticker mapping.")

        cik = f"{int(ticker_record['cik_str']):010d}"
        company_name = str(ticker_record.get("title", "")).strip()
        submission_response = await self._sec_get(
            client, SEC_SUBMISSIONS_URL.format(cik=cik)
        )
        recent = submission_response.json().get("filings", {}).get("recent", {})
        forms = recent.get("form", [])
        accession_numbers = recent.get("accessionNumber", [])
        filing_dates = recent.get("filingDate", [])
        primary_documents = recent.get("primaryDocument", [])
        today = datetime.now(timezone.utc).date()
        insider_cutoff = today - timedelta(days=SEC_LOOKBACK_DAYS)
        filings_cutoff = today - timedelta(days=SEC_FILINGS_LOOKBACK_DAYS)
        filing_rows = []
        for index, form in enumerate(forms):
            try:
                filed_date = datetime.strptime(
                    filing_dates[index], "%Y-%m-%d"
                ).date()
            except (IndexError, ValueError):
                continue
            if filed_date < filings_cutoff:
                continue
            if form not in {"10-K", "10-Q", "8-K", "4", "4/A"}:
                continue
            accession = accession_numbers[index]
            document = primary_documents[index]
            filing_rows.append(
                {
                    "form": form,
                    "filed_date": filed_date,
                    "accession": accession,
                    "document": document,
                    "url": SEC_ARCHIVES_URL.format(
                        cik=str(int(cik)),
                        accession=accession.replace("-", ""),
                        document=document,
                    ),
                }
            )

        form4_rows = [
            row for row in filing_rows
            if row["form"] in {"4", "4/A"} and row["filed_date"] >= insider_cutoff
        ][:8]
        insider_transactions: List[Dict[str, Any]] = []
        errors: List[str] = []
        parsed_form4_count = 0
        for row in form4_rows:
            try:
                response = await self._sec_get(client, row["url"])
                transactions = self._parse_form4_xml(
                    response.content, row["filed_date"].isoformat(), row["url"]
                )
                insider_transactions.extend(transactions)
                parsed_form4_count += 1
            except (httpx.HTTPError, ET.ParseError, ValueError) as exc:
                errors.append(f"Form 4 {row['accession']}: {type(exc).__name__}")

        company_filings = [
            row for row in filing_rows
            if row["form"] in {"10-K", "10-Q", "8-K"}
        ]
        selected_filings: List[Dict[str, Any]] = []
        for form in ("10-K", "10-Q"):
            candidates = [row for row in company_filings if row["form"] == form]
            if candidates:
                selected_filings.append(candidates[0])
        selected_filings.extend(
            [row for row in company_filings if row["form"] == "8-K"][:3]
        )
        filing_payload = []
        for row in selected_filings:
            try:
                response = await self._sec_get(client, row["url"])
                filing_payload.append(
                    {
                        "form": row["form"],
                        "filed_date": row["filed_date"].isoformat(),
                        "accession": row["accession"],
                        "url": row["url"],
                        "excerpt": self._parse_sec_excerpt(
                            row["form"], response.text
                        ),
                    }
                )
            except (httpx.HTTPError, ValueError) as exc:
                errors.append(
                    f"{row['form']} {row['accession']}: {type(exc).__name__}"
                )

        insider_payload = {
            "company_name": company_name,
            "lookback_days": SEC_LOOKBACK_DAYS,
            "filing_count": len(form4_rows),
            "parsed_filing_count": parsed_form4_count,
            "transactions_verified": parsed_form4_count > 0,
            "filings": insider_transactions,
            "open_market_buys_usd": sum(
                item["value_usd"]
                for item in insider_transactions
                if item["type"] == "PURCHASE" and item["value_usd"] is not None
            ),
            "open_market_sales_usd": sum(
                item["value_usd"]
                for item in insider_transactions
                if item["type"] == "SALE" and item["value_usd"] is not None
            ),
            "source": "SEC EDGAR Form 4",
        }
        filing_result = {
            "company_name": company_name,
            "lookback_days": SEC_FILINGS_LOOKBACK_DAYS,
            "filings": filing_payload,
            "source": "SEC EDGAR",
        }
        return insider_payload, filing_result, "; ".join(errors) or None

    def _fetch_yahoo_bundle(self, ticker_str: str) -> Tuple[Dict[str, Any], Dict[str, str]]:
        errors: Dict[str, str] = {}
        try:
            t = yf.Ticker(ticker_str)
            spy = yf.Ticker("SPY")
        except Exception as exc:
            msg = f"{type(exc).__name__}: {exc}"
            return {}, {"Global": msg}

        try:
            info = t.info or {}
        except Exception as exc:
            info = {}
            errors["Financials"] = f"Info fetch failed: {type(exc).__name__}"

        # 1. 1-Year History + Strict Date-Aligned Benchmark
        hist_1y = pd.DataFrame()
        spy_hist_1y = pd.DataFrame()
        try:
            hist_1y = t.history(period="1y")
            spy_hist_1y = spy.history(period="1y")

            closes = hist_1y["Close"].tolist() if not hist_1y.empty else []
            volumes = hist_1y["Volume"].tolist() if not hist_1y.empty else []
            dates = [d.strftime("%Y-%m-%d") for d in hist_1y.index]
            tech_metrics = self._calculate_series_metrics(closes, volumes)

            rel_perf = {}
            if not hist_1y.empty and not spy_hist_1y.empty:
                df = pd.DataFrame({"stock": hist_1y["Close"], "spy": spy_hist_1y["Close"]}).dropna()
                if len(df) >= 20:
                    stock_ret_1m = (df["stock"].iloc[-1] - df["stock"].iloc[-20]) / df["stock"].iloc[-20]
                    spy_ret_1m = (df["spy"].iloc[-1] - df["spy"].iloc[-20]) / df["spy"].iloc[-20]
                    rel_perf["excess_return_vs_spy_1m_pct"] = round((stock_ret_1m - spy_ret_1m) * 100, 2)

            price_payload = {
                "closes": closes[-60:],
                "dates": dates[-60:],
                "metrics": tech_metrics,
                "benchmark_comparison": rel_perf,
                "currency": info.get("currency", "USD"),
            }
        except Exception as exc:
            price_payload = {"closes": [], "dates": [], "metrics": {}}
            errors["Price_Chart"] = f"{type(exc).__name__}: {exc}"

        # 2. Earnings History
        earnings_payload = {
            "forward_eps": info.get("forwardEps"),
            "trailing_eps": info.get("trailingEps"),
            "surprises": [],
        }
        try:
            surprises_df = t.earnings_dates
            if surprises_df is not None and not surprises_df.empty:
                recent = []
                for idx, row in surprises_df.head(4).iterrows():
                    rep = row.get("Reported EPS")
                    est = row.get("EPS Estimate")
                    pct = row.get("Surprise(%)")
                    if rep is not None and not math.isnan(rep):
                        recent.append({
                            "date": idx.strftime("%Y-%m-%d") if hasattr(idx, "strftime") else str(idx),
                            "reported_eps": rep,
                            "estimated_eps": est,
                            "surprise_pct": round(pct, 2) if pct and not math.isnan(pct) else None
                        })
                earnings_payload["surprises"] = recent
        except Exception as exc:
            errors["Earnings_Estimates"] = f"{type(exc).__name__}: {exc}"

        insider_data = {
            "filings": [],
            "transactions_verified": False,
            "source": "SEC EDGAR Form 4",
        }
        try:
            options_payload = self._build_options_payload(t)
        except Exception as exc:
            options_payload = {}
            errors["Options_Market"] = (
                f"{type(exc).__name__}: {exc}"
            )

        # 4. Standard Fundamental Metrics
        financials_payload = self._build_financials_payload(info)
        institutional_payload = self._build_institutional_payload(info)

        # 5. Yahoo Finance headlines; an RSS fallback is attempted asynchronously.
        news_payload = {"recent_headlines": []}
        try:
            for item in (t.news or [])[:10]:
                content = item.get("content") or item
                title = content.get("title")
                if not title:
                    continue
                publisher = content.get("publisher") or {}
                canonical_url = content.get("canonicalUrl") or {}
                news_payload["recent_headlines"].append(
                    {
                        "title": title,
                        "publisher": (
                            publisher.get("displayName")
                            if isinstance(publisher, dict)
                            else publisher
                        ),
                        "link": (
                            canonical_url.get("url")
                            if isinstance(canonical_url, dict)
                            else content.get("link")
                        ),
                        "published_at": content.get("pubDate")
                        or content.get("providerPublishTime"),
                    }
                )
        except Exception as exc:
            errors["News_Events"] = f"{type(exc).__name__}: {exc}"

        try:
            related_payload, related_error = self._build_related_markets_payload(
                ticker_str, info, hist_1y, spy_hist_1y
            )
            if related_error:
                errors["Related_Markets"] = related_error
        except Exception as exc:
            related_payload = {}
            errors["Related_Markets"] = f"{type(exc).__name__}: {exc}"

        return {
            "Price_Chart": price_payload,
            "Financials": financials_payload,
            "Earnings_Estimates": earnings_payload,
            "Insider_Trades": insider_data,
            "News_Events": news_payload,
            "Institutional_Flows": institutional_payload,
            "Options_Market": options_payload,
            "Related_Markets": related_payload,
            "_ticker_company_name": (
                info.get("longName") or info.get("shortName") or ""
            ),
        }, errors

    async def _fetch_fred_macro(self, client: httpx.AsyncClient) -> Tuple[Dict[str, Any], Optional[str]]:
        if not self.fred_api_key:
            return {}, "FRED_API_KEY not configured."
        macro_data = {}
        errors = []
        for key, series_id in {"yield_10y": "DGS10", "yield_2y": "DGS2"}.items():
            url = "https://api.stlouisfed.org/fred/series/observations"
            params = {
                "series_id": series_id,
                "api_key": self.fred_api_key,
                "file_type": "json",
                "sort_order": "desc",
                "limit": 1,
            }
            try:
                response = await client.get(url, params=params, timeout=5.0)
                response.raise_for_status()
                observations = response.json().get("observations", [])
                if observations and observations[0]["value"] != ".":
                    macro_data[key] = float(observations[0]["value"])
                else:
                    errors.append(f"{series_id}: no current observation")
            except Exception as exc:
                errors.append(f"{series_id}: {type(exc).__name__}")

        if "yield_10y" in macro_data and "yield_2y" in macro_data:
            macro_data["yield_curve_spread_10y_2y"] = round(
                macro_data["yield_10y"] - macro_data["yield_2y"], 2
            )
        return macro_data, "; ".join(errors) or None

    async def _fetch_news_rss(
        self, ticker: str, company_name: Optional[str], client: httpx.AsyncClient
    ) -> List[Dict[str, Any]]:
        query = company_name or ticker
        url = (
            "https://news.google.com/rss/search?q="
            f"{quote_plus(f'\"{query}\" stock')}&hl=en-US&gl=US&ceid=US:en"
        )
        response = await client.get(
            url,
            headers={"User-Agent": "Mozilla/5.0 StockEvaluatorEngine/1.0"},
            timeout=12.0,
        )
        response.raise_for_status()
        root = ET.fromstring(response.content)
        headlines = []
        for item in root.findall(".//item")[:8]:
            title = item.findtext("title")
            if not title:
                continue
            headlines.append(
                {
                    "title": title,
                    "publisher": "Google News RSS",
                    "link": item.findtext("link"),
                    "published_at": item.findtext("pubDate"),
                }
            )
        return headlines

    async def _load_house_ptr_reports(
        self, client: httpx.AsyncClient
    ) -> Tuple[List[Dict[str, Any]], Optional[str], bool]:
        if self._house_ptr_cache is not None:
            return self._house_ptr_cache

        cutoff = datetime.now(timezone.utc).date() - timedelta(days=SEC_LOOKBACK_DAYS)
        rows: List[Dict[str, Any]] = []
        errors: List[str] = []
        for year in sorted({cutoff.year, datetime.now(timezone.utc).year}):
            url = (
                "https://disclosures-clerk.house.gov/public_disc/financial-pdfs/"
                f"{year}FD.zip"
            )
            try:
                response = await client.get(url, timeout=20.0)
                response.raise_for_status()
                with ZipFile(BytesIO(response.content)) as archive:
                    text_file = next(
                        name for name in archive.namelist()
                        if name.lower().endswith(".txt")
                    )
                    reader = csv.DictReader(
                        archive.read(text_file).decode("utf-8-sig").splitlines(),
                        delimiter="\t",
                    )
                    for row in reader:
                        if row.get("FilingType", "").upper() != "P":
                            continue
                        try:
                            filed_date = datetime.strptime(
                                row["FilingDate"], "%m/%d/%Y"
                            ).date()
                        except (KeyError, ValueError):
                            continue
                        if filed_date >= cutoff:
                            rows.append({**row, "filing_date": filed_date})
            except (httpx.HTTPError, BadZipFile, OSError, StopIteration) as exc:
                errors.append(f"{year} House PTR index: {type(exc).__name__}")

        rows.sort(key=lambda row: row["filing_date"], reverse=True)
        truncated = len(rows) > HOUSE_PTR_MAX_REPORTS
        selected = rows[:HOUSE_PTR_MAX_REPORTS]
        semaphore = asyncio.Semaphore(3)

        async def fetch_report(row: Dict[str, Any]) -> Optional[Dict[str, Any]]:
            report_url = (
                "https://disclosures-clerk.house.gov/public_disc/ptr-pdfs/"
                f"{row['Year']}/{row['DocID']}.pdf"
            )
            try:
                async with semaphore:
                    response = await client.get(report_url, timeout=20.0)
                    response.raise_for_status()
                reader = PdfReader(BytesIO(response.content))
                report_text = "\n".join(
                    page.extract_text() or "" for page in reader.pages
                ).replace("\x00", "")
                if not report_text.strip():
                    return None
                return {
                    "member_name": " ".join(
                        part for part in (
                            row.get("Prefix", ""),
                            row.get("First", ""),
                            row.get("Last", ""),
                            row.get("Suffix", ""),
                        )
                        if part
                    ),
                    "filing_date": row["filing_date"].isoformat(),
                    "state_district": row.get("StateDst"),
                    "source_url": report_url,
                    "report_text": report_text,
                }
            except (httpx.HTTPError, OSError, ValueError, PdfReadError) as exc:
                errors.append(f"House PTR {row.get('DocID')}: {type(exc).__name__}")
                return None

        fetched = await asyncio.gather(*(fetch_report(row) for row in selected))
        reports = [report for report in fetched if report is not None]
        if truncated:
            errors.append(
                f"House PTR scan capped at {HOUSE_PTR_MAX_REPORTS} reports."
            )
        self._house_ptr_cache = (reports, "; ".join(errors) or None, truncated)
        return self._house_ptr_cache

    def _load_senate_congress_rows(
        self, ticker: str
    ) -> Tuple[List[Dict[str, Any]], Optional[str]]:
        if not self.congress_trades_csv_path.is_file():
            return [], None
        today = datetime.now(timezone.utc).date()
        cutoff = today - timedelta(days=SEC_LOOKBACK_DAYS)
        required_headers = {
            "ticker", "chamber", "transaction_date", "disclosure_date",
            "asset_description", "transaction_type", "amount_range", "source_url",
        }
        try:
            with self.congress_trades_csv_path.open(
                "r", encoding="utf-8-sig", newline=""
            ) as csv_file:
                reader = csv.DictReader(csv_file)
                header_lookup: Dict[str, str] = {}
                for header in reader.fieldnames or []:
                    normalized_header = header.strip().casefold()
                    if normalized_header in header_lookup:
                        return [], (
                            "Congress trade CSV contains duplicate column names "
                            f"after normalization: {normalized_header}"
                        )
                    header_lookup[normalized_header] = header
                missing_headers = required_headers - header_lookup.keys()
                if missing_headers:
                    missing = ", ".join(sorted(missing_headers))
                    return [], f"Congress trade CSV missing columns: {missing}"

                rows = []
                for raw_row in reader:
                    row = {
                        field: raw_row.get(original_header) or ""
                        for field, original_header in header_lookup.items()
                    }
                    if row["ticker"].strip().upper() != ticker.upper():
                        continue
                    chamber = row["chamber"].strip().casefold()
                    if chamber not in {"senate", "senator"}:
                        continue
                    transaction_date = datetime.strptime(
                        row["transaction_date"].strip(), "%Y-%m-%d"
                    ).date()
                    if transaction_date < cutoff or transaction_date > today:
                        continue
                    if not row["source_url"].strip():
                        continue
                    rows.append(
                        {
                            "chamber": "Senate",
                            "ticker": ticker.upper(),
                            "transaction_date": transaction_date.isoformat(),
                            "disclosure_date": row["disclosure_date"].strip(),
                            "asset_description": row["asset_description"].strip(),
                            "transaction_type": row["transaction_type"].strip(),
                            "amount_range": row["amount_range"].strip(),
                            "source_url": row["source_url"].strip(),
                        }
                    )
            return rows, None
        except (OSError, UnicodeError, csv.Error, KeyError, ValueError) as exc:
            return [], f"Congress trade CSV could not be read: {type(exc).__name__}: {exc}"

    @staticmethod
    def _match_house_report_asset(
        ticker: str, company_names: List[str], report_text: str
    ) -> Optional[str]:
        symbol_pattern = None
        symbol = ticker.strip().upper()
        if len(symbol) > 1:
            symbol_pattern = re.compile(
                rf"(?<![A-Z0-9]){re.escape(symbol)}(?![A-Z0-9])"
            )
        ticker_match = (
            symbol_pattern.search(report_text.upper().replace("\x00", ""))
            if symbol_pattern
            else None
        )
        if ticker_match:
            return ticker_match.group(0)

        normalized_text = " ".join(
            report_text.casefold().replace("\x00", "").split()
        )
        for name in company_names:
            normalized_name = " ".join(
                name.casefold().replace("\x00", "").split()
            )
            if len(normalized_name) >= 4 and re.search(
                rf"(?<![a-z0-9]){re.escape(normalized_name)}(?![a-z0-9])",
                normalized_text,
            ):
                return normalized_name
        return None

    async def _fetch_congress_trades(
        self,
        ticker: str,
        company_names: List[str],
        client: httpx.AsyncClient,
    ) -> Tuple[Dict[str, Any], Optional[str], bool]:
        reports, house_error, truncated = await self._load_house_ptr_reports(client)
        house_events = []
        for report in reports:
            text = report["report_text"]
            matched_text = self._match_house_report_asset(
                ticker, company_names, text
            )
            if not matched_text:
                continue
            house_events.append(
                {
                    "chamber": "House",
                    "member_name": report["member_name"],
                    "filing_date": report["filing_date"],
                    "state_district": report["state_district"],
                    "asset_match": matched_text,
                    "match_method": "ticker text or issuer-name substring; verify the report",
                    "report_excerpt": text[:3000],
                    "source_url": report["source_url"],
                    "amounts_are_ranges": True,
                }
            )

        senate_events, senate_error = self._load_senate_congress_rows(ticker)
        errors = [error for error in (house_error, senate_error) if error]
        result = {
            "lookback_days": SEC_LOOKBACK_DAYS,
            "house_reports_scanned": len(reports),
            "house_scan_truncated": truncated,
            "house_events": house_events,
            "senate_events": senate_events,
            "source": "House Clerk public PTR PDFs; Senate manually imported records",
            "limitations": (
                "Reports can be delayed; disclosed trade values are ranges. House "
                "ticker/company-name matching is best effort and must be checked "
                "against the linked filing. No matched record is not proof of no trade."
            ),
        }
        return result, "; ".join(errors) or None, truncated

    def _check_completeness(self, data: Any) -> Tuple[bool, bool]:
        """Recursively inspects data to return (has_data, has_missing)."""
        if data is None or (isinstance(data, (list, tuple, str, dict)) and len(data) == 0):
            return False, True
            
        if isinstance(data, dict):
            has_d, has_m = False, False
            for v in data.values():
                d, m = self._check_completeness(v)
                has_d = has_d or d
                has_m = has_m or m
            return has_d, has_m
            
        if isinstance(data, list):
            has_d, has_m = False, False
            for v in data:
                d, m = self._check_completeness(v)
                has_d = has_d or d
                has_m = has_m or m
            return has_d, has_m
            
        return True, False

    def _get_status(self, seat_data: Any, error: Optional[str]) -> str:
        """Determines if the payload is healthy, partially missing, completely null, or failed."""
        has_data, has_missing = self._check_completeness(seat_data)
        if not has_data:
            return "error" if error else "missing"
        if error:
            return "partial"
        if has_missing:
            return "partial"
        return "ok"

    async def fetch_all(self, ticker: str) -> Dict[str, Any]:
        loop = asyncio.get_running_loop()
        yahoo_task = loop.run_in_executor(None, self._fetch_yahoo_bundle, ticker)
        yahoo_res = await yahoo_task
        if isinstance(yahoo_res, Exception):
            yahoo_data, yahoo_errors = {}, {"Global": str(yahoo_res)}
        else:
            yahoo_data, yahoo_errors = yahoo_res

        if "Global" in yahoo_errors:
            global_error = yahoo_errors.pop("Global")
            for seat in (
                "Price_Chart",
                "Financials",
                "Earnings_Estimates",
                "News_Events",
                "Institutional_Flows",
                "Options_Market",
                "Related_Markets",
            ):
                yahoo_errors[seat] = global_error

        async with httpx.AsyncClient() as client:
            company_name = str(
                yahoo_data.get("_ticker_company_name", "")
            ).strip()
            fred_task = asyncio.create_task(self._fetch_fred_macro(client))
            house_task = asyncio.create_task(
                self._fetch_congress_trades(
                    ticker,
                    [company_name] if company_name else [],
                    client,
                )
            )
            news_task = None
            if not yahoo_data.get("News_Events", {}).get("recent_headlines"):
                news_task = asyncio.create_task(
                    self._fetch_news_rss(ticker, company_name or None, client)
                )
            sec_task = (
                asyncio.create_task(self._fetch_sec_company_data(ticker, client))
                if self.sec_headers
                else None
            )
            task_list = [fred_task, house_task]
            if news_task:
                task_list.append(news_task)
            if sec_task:
                task_list.append(sec_task)
            results = await asyncio.gather(*task_list, return_exceptions=True)

        result_index = 0
        fred_res = results[result_index]
        result_index += 1
        house_res = results[result_index]
        result_index += 1
        if isinstance(fred_res, Exception):
            fred_data, fred_error = {}, f"{type(fred_res).__name__}: {fred_res}"
        else:
            fred_data, fred_error = fred_res

        if isinstance(house_res, Exception):
            congress_data = {
                "lookback_days": SEC_LOOKBACK_DAYS,
                "house_reports_scanned": 0,
                "house_events": [],
                "senate_events": [],
                "limitations": "No House PTR reports were processed.",
            }
            congress_error = f"{type(house_res).__name__}: {house_res}"
            house_truncated = False
        else:
            congress_data, congress_error, house_truncated = house_res

        news_rss_error = None
        if news_task:
            news_res = results[result_index]
            result_index += 1
            if isinstance(news_res, Exception):
                news_rss_error = (
                    f"Google News RSS: {type(news_res).__name__}: {news_res}"
                )
            elif news_res:
                yahoo_data.setdefault("News_Events", {})["recent_headlines"] = news_res
                yahoo_data["News_Events"]["fallback_source"] = "Google News RSS"
        else:
            news_rss_error = None

        sec_error = None
        sec_configured = sec_task is not None
        if sec_task:
            sec_res = results[result_index]
            if isinstance(sec_res, Exception):
                sec_insider_data = {
                    "filings": [],
                    "filing_count": 0,
                    "transactions_verified": False,
                    "source": "SEC EDGAR Form 4",
                }
                sec_filings_data = {"filings": [], "source": "SEC EDGAR"}
                sec_error = f"{type(sec_res).__name__}: {sec_res}"
            else:
                sec_insider_data, sec_filings_data, sec_error = sec_res
        else:
            sec_insider_data = {
                "filings": [],
                "filing_count": 0,
                "transactions_verified": False,
                "source": "SEC EDGAR Form 4",
            }
            sec_filings_data = {"filings": [], "source": "SEC EDGAR"}

        payload = yahoo_data
        payload.pop("_ticker_company_name", None)
        payload["Insider_Trades"] = sec_insider_data
        payload["SEC_Filings"] = sec_filings_data
        payload["Congress_Trades"] = congress_data
        payload["Macro_Economy"] = fred_data

        metadata = {
            "collected_at": datetime.now(timezone.utc).isoformat(),
            "ticker": ticker.upper(),
            "sources": {},
        }

        for seat in (
            "Price_Chart",
            "Financials",
            "Earnings_Estimates",
            "News_Events",
            "Institutional_Flows",
            "Options_Market",
            "Related_Markets",
        ):
            metadata["sources"][seat] = {
                "source": (
                    "Yahoo Finance + Google News RSS"
                    if seat == "News_Events" and "fallback_source" in payload.get(seat, {})
                    else "Yahoo Finance"
                ),
                "status": self._get_status(payload.get(seat), yahoo_errors.get(seat)),
                "error": (
                    yahoo_errors.get(seat)
                    or (news_rss_error if seat == "News_Events" else None)
                ),
            }

        metadata["sources"]["Insider_Trades"] = {
            "source": "SEC EDGAR Form 4",
            "status": (
                "not_configured"
                if not sec_configured
                else (
                    "partial"
                    if sec_error and sec_insider_data.get("parsed_filing_count")
                    else (
                        "error"
                        if sec_error
                        else (
                            "missing"
                            if not sec_insider_data.get("filing_count")
                            or not sec_insider_data.get("filings")
                            else "ok"
                        )
                    )
                )
            ),
            "error": sec_error or (
                "Set SEC_CONTACT_EMAIL or SEC_USER_AGENT to enable SEC requests."
                if not sec_configured
                else None
            ),
        }
        metadata["sources"]["SEC_Filings"] = {
            "source": "SEC EDGAR",
            "status": (
                "not_configured"
                if not sec_configured
                else (
                    "partial"
                    if sec_error and sec_filings_data.get("filings")
                    else (
                        "error"
                        if sec_error
                        else (
                            "missing"
                            if not sec_filings_data.get("filings")
                            else "ok"
                        )
                    )
                )
            ),
            "error": sec_error or (
                "Set SEC_CONTACT_EMAIL or SEC_USER_AGENT to enable SEC requests."
                if not sec_configured
                else None
            ),
        }
        congress_status = (
            "partial"
            if congress_data.get("house_reports_scanned")
            or congress_data.get("senate_events")
            else "error" if congress_error else "missing"
        )
        metadata["sources"]["Congress_Trades"] = {
            "source": "House Clerk public PTR PDFs + Senate CSV import",
            "status": congress_status,
            "error": congress_error or (
                "Senate trades require a local CSV import; House PTR matches are best effort."
            ),
            "house_scan_truncated": house_truncated,
        }
        metadata["sources"]["Macro_Economy"] = {
            "source": "FRED",
            "status": self._get_status(payload.get("Macro_Economy"), fred_error),
            "error": fred_error
        }

        capiq_data, capiq_error, capiq_as_of_date = self._load_capital_iq_excel(ticker)
        for seat_name in ("Analyst_Targets", "Earnings_Estimates", "Financials", "Institutional_Flows"):
            imported_values = capiq_data.get(seat_name, {})
            if imported_values:
                existing = payload.get(seat_name, {})
                payload[seat_name] = {
                    **(existing if isinstance(existing, dict) else {}),
                    **imported_values,
                }

        for seat_name in ("Analyst_Targets", "Earnings_Estimates", "Financials", "Institutional_Flows"):
            has_capiq_data = bool(capiq_data.get(seat_name))
            if has_capiq_data:
                prior_source = metadata["sources"].get(seat_name, {}).get("source")
                metadata["sources"][seat_name] = {
                    "source": (
                        f"{prior_source} + Capital IQ Excel workbook"
                        if prior_source
                        else "Capital IQ Excel workbook"
                    ),
                    "as_of_date": capiq_as_of_date,
                    "status": self._get_status(payload[seat_name], capiq_error),
                    "error": capiq_error,
                }
            elif self.capital_iq_excel_path:
                prior_metadata = metadata["sources"].get(seat_name, {})
                prior_status = prior_metadata.get("status", "missing")
                if prior_status in {"ok", "partial"}:
                    prior_metadata["status"] = "partial"
                    prior_metadata["error"] = capiq_error
                    prior_metadata["capiq_status"] = "missing"
                else:
                    metadata["sources"][seat_name] = {
                        "source": "Capital IQ Excel workbook",
                        "as_of_date": None,
                        "status": "error" if capiq_error else "missing",
                        "error": capiq_error or f"No Capital IQ data found for {ticker.upper()}.",
                    }
            elif seat_name not in metadata["sources"]:
                metadata["sources"][seat_name] = {
                    "source": None,
                    "as_of_date": None,
                    "status": "not_configured",
                    "error": None,
                }

        payload["_metadata"] = metadata
        return payload