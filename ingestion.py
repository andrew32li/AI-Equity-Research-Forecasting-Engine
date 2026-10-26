"""
ingestion.py - Robust Asynchronous Data Ingestion for Ticker Council
"""

import asyncio
from datetime import datetime, timezone
import math
import os
import posixpath
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Tuple
from zipfile import BadZipFile, ZipFile
import xml.etree.ElementTree as ET

import httpx
import pandas as pd
import yfinance as yf


class DataIngestionPipeline:
    def __init__(
        self,
        fred_api_key: Optional[str] = None,
        sec_user_agent: str = "InvestmentCouncil ResearchBot@example.com",
        capital_iq_excel_path: Optional[str] = None,
    ):
        self.fred_api_key = fred_api_key or os.environ.get("FRED_API_KEY")
        configured_excel_path = (
            capital_iq_excel_path
            or os.environ.get("CAPITAL_IQ_XLSX_PATH")
            or Path(__file__).resolve().with_name("report.xlsx")
        )
        self.capital_iq_excel_path = Path(configured_excel_path)
        self.headers = {"User-Agent": sec_user_agent}

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

        # Yahoo's transaction descriptions are not authoritative Form 4 codes.
        insider_data = {"filings": [], "transactions_verified": False}
        try:
            insiders_df = t.insider_transactions
            if insiders_df is not None and not insiders_df.empty:
                parsed = []
                for _, row in insiders_df.head(15).iterrows():
                    text = str(row.get("Text", "")).lower()
                    shares = row.get("Shares", 0) or 0
                    val = row.get("Value", 0) or 0
                    reported_type = "UNKNOWN"
                    if "buy" in text or "purchase" in text:
                        reported_type = "PURCHASE"
                    elif "sale" in text or "sold" in text:
                        reported_type = "SALE"

                    parsed.append({
                        "insider": row.get("Insider"),
                        "reported_type": reported_type,
                        "shares": int(shares),
                        "value_usd": float(val),
                        "date": str(row.get("Start Date", ""))[:10],
                        "source_text": str(row.get("Text", "")),
                    })
                insider_data["filings"] = parsed
        except Exception as exc:
            errors["Insider_Trades"] = f"{type(exc).__name__}: {exc}"

        # 4. Standard Fundamental Metrics
        financials_payload = self._build_financials_payload(info)

        # 5. Isolated News Fetch
        news_payload = {"recent_headlines": []}
        try:
            if t.news:
                news_payload["recent_headlines"] = [
                    {"title": item.get("title"), "publisher": item.get("publisher"), "link": item.get("link")}
                    for item in t.news[:5]
                ]
        except Exception as exc:
            errors["News_Events"] = f"{type(exc).__name__}: {exc}"

        return {
            "Price_Chart": price_payload,
            "Financials": financials_payload,
            "Earnings_Estimates": earnings_payload,
            "Insider_Trades": insider_data,
            "News_Events": news_payload,
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

        async with httpx.AsyncClient() as client:
            fred_task = self._fetch_fred_macro(client)
            results = await asyncio.gather(yahoo_task, fred_task, return_exceptions=True)

        yahoo_res = results[0]
        if isinstance(yahoo_res, Exception):
            yahoo_data, yahoo_errors = {}, {"Global": str(yahoo_res)}
        else:
            yahoo_data, yahoo_errors = yahoo_res

        # Distribute Global Yahoo errors to all Yahoo-dependent seats
        if "Global" in yahoo_errors:
            global_err = yahoo_errors.pop("Global")
            for seat in ["Price_Chart", "Financials", "Earnings_Estimates", "Insider_Trades", "News_Events"]:
                yahoo_errors[seat] = global_err

        fred_res = results[1]
        if isinstance(fred_res, Exception):
            fred_data, fred_error = {}, str(fred_res)
        else:
            fred_data, fred_error = fred_res

        payload = yahoo_data
        payload["Macro_Economy"] = fred_data

        metadata = {"collected_at": datetime.now(timezone.utc).isoformat(), "ticker": ticker.upper(), "sources": {}}
        
        for seat in ["Price_Chart", "Financials", "Earnings_Estimates", "Insider_Trades", "News_Events"]:
            metadata["sources"][seat] = {
                "source": "Yahoo Finance",
                "status": self._get_status(payload.get(seat), yahoo_errors.get(seat)),
                "error": yahoo_errors.get(seat)
            }
            
        metadata["sources"]["Macro_Economy"] = {
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

        for unconfigured in ("Congress_Trades", "SEC_Filings", "Related_Markets", "Options_Market"):
            payload[unconfigured] = {}
            metadata["sources"][unconfigured] = {"status": "not_configured", "error": None}

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