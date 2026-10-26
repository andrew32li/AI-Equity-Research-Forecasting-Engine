"""
council.py - Institutional Multi-Agent Investment Council Engine
Engineered for OpenAI with Pydantic Structured Outputs and Tiered Architecture:
  - Phase 1: 12 Specialist Seats (gpt-4o-mini)
  - Phase 2: Adversarial Bull/Bear/Challenger Debate (gpt-4o)
  - Phase 3: Calibrated Multi-Horizon Grand Master Judge (gpt-4o)
"""

import asyncio
import json
import os
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple

from openai import AsyncOpenAI
from pydantic import BaseModel, Field, model_validator

# =====================================================================
# 1. SCHEMAS & DATA STRUCTURES
# =====================================================================

class Signal(str, Enum):
    UP = "UP"
    DOWN = "DOWN"
    EVEN = "EVEN"
    NO_READ = "NO_READ"

class SeatOpinion(BaseModel):
    seat_name: str
    signal: Signal
    confidence: float = Field(ge=0.0, le=1.0, description="Confidence score from 0.0 to 1.0")
    summary: str = Field(description="Concise rationale strictly covering domain findings.")

class DebateArguments(BaseModel):
    bull_case: str = Field(description="Steelman positive argument based strictly on supportive seats.")
    bear_case: str = Field(description="Steelman negative risk assessment based on critical seats.")
    challenger_critique: str = Field(description="Adversarial red-team attack identifying consensus bias and tail risks.")

class HorizonForecast(BaseModel):
    horizon: str = Field(description="Target period: '1_week', '3_months', or '1_year'")
    direction: Signal = Field(description="UP above 55%, DOWN below 45%, otherwise EVEN.")
    probability_up: float = Field(ge=0.0, le=1.0, description="Estimated probability of positive return")
    reasoning: str

    @model_validator(mode="after")
    def validate_direction_matches_probability(self) -> "HorizonForecast":
        if self.direction == Signal.UP and self.probability_up <= 0.55:
            raise ValueError("UP requires probability_up > 0.55")
        if self.direction == Signal.DOWN and self.probability_up >= 0.45:
            raise ValueError("DOWN requires probability_up < 0.45")
        if self.direction == Signal.EVEN and not 0.45 <= self.probability_up <= 0.55:
            raise ValueError("EVEN requires probability_up between 0.45 and 0.55 inclusive")
        if self.direction == Signal.NO_READ:
            raise ValueError("Forecast direction cannot be NO_READ")
        return self


REQUIRED_FORECAST_HORIZONS = {"1_week", "3_months", "1_year"}


def _validate_forecast_horizons(forecasts: List[HorizonForecast]) -> List[HorizonForecast]:
    horizons = [forecast.horizon for forecast in forecasts]
    if set(horizons) != REQUIRED_FORECAST_HORIZONS or len(horizons) != len(REQUIRED_FORECAST_HORIZONS):
        raise ValueError("Forecasts must contain exactly one entry for each required horizon.")
    return forecasts


class CouncilVerdict(BaseModel):
    ticker: str
    forecasts: List[HorizonForecast]
    seat_breakdown: List[SeatOpinion]
    debate: DebateArguments

    @model_validator(mode="after")
    def validate_forecast_horizons(self) -> "CouncilVerdict":
        _validate_forecast_horizons(self.forecasts)
        return self


class _ForecastWrapper(BaseModel):
    forecasts: List[HorizonForecast]

    @model_validator(mode="after")
    def validate_forecast_horizons(self) -> "_ForecastWrapper":
        _validate_forecast_horizons(self.forecasts)
        return self


# =====================================================================
# 2. SEAT DEFINITIONS
# =====================================================================

SEAT_CONFIGS = [
    {"name": "Price_Chart", "desc": "Technical momentum, support/resistance, RSI, moving averages, and volume."},
    {"name": "Financials", "desc": "Provided valuation ratios, free cash flow and free-cash-flow yield when available, debt, cash, and margin metrics."},
    {"name": "Analyst_Targets", "desc": "Consensus price targets, upgrades/downgrades, estimate dispersion."},
    {"name": "Earnings_Estimates", "desc": "EPS consensus revisions, revenue drift, historical earnings surprise rates."},
    {"name": "Insider_Trades", "desc": "Form 4 purchases/disposals by C-suite executives and board directors."},
    {"name": "Congress_Trades", "desc": "Capitol Hill stock disclosure activity and committee cross-holdings."},
    {"name": "News_Events", "desc": "Media sentiment catalysts, product announcements, management turnover."},
    {"name": "SEC_Filings", "desc": "10-Q/10-K notes, auditor warnings, pending litigation, dilution risks."},
    {"name": "Institutional_Flows", "desc": "13F accumulations, hedge fund concentration, short interest % of float."},
    {"name": "Options_Market", "desc": "Put/Call ratio, implied volatility skew, gamma walls, unusual options sweeps."},
    {"name": "Macro_Economy", "desc": "Treasury yield curve and macro context; a yield spread alone is not directional evidence for an individual stock."},
    {"name": "Related_Markets", "desc": "Sector ETF relative strength, supply chain peer shifts, currency/commodity beta."}
]


# =====================================================================
# 3. DETERMINISTIC PRE-PROMPTING FILTER (ZERO-TOKEN LOCAL BYPASS)
# =====================================================================

# =====================================================================
# 3. DETERMINISTIC PRE-PROMPTING FILTER (UPDATED DATA FLOW)
# =====================================================================

class PrePromptValidator:
    @staticmethod
    def validate_price_chart(data: Optional[Dict[str, Any]]) -> Tuple[Optional[SeatOpinion], Optional[Dict[str, Any]]]:
        if not data or not data.get("closes") or len(data["closes"]) < 15:
            return SeatOpinion(
                seat_name="Price_Chart",
                signal=Signal.NO_READ,
                confidence=0.0,
                summary="Insufficient price bars (14 RSI changes require at least 15 closes)."
            ), None

        closes = data["closes"]
        
        # Fixed RSI calculation covering flatline division edge cases
        gains, losses = [], []
        for i in range(len(closes) - 14, len(closes)):
            diff = closes[i] - closes[i - 1]
            gains.append(max(diff, 0.0))
            losses.append(max(-diff, 0.0))

        avg_gain = sum(gains) / 14
        avg_loss = sum(losses) / 14
        
        if avg_loss == 0 and avg_gain == 0:
            rsi = 50.0
        elif avg_loss == 0:
            rsi = 100.0
        else:
            rs = avg_gain / avg_loss
            rsi = 100 - (100 / (1 + rs))

        payload = {
            "rsi_14": round(rsi, 2),
            # Compare current close to the close 5 trading days ago (index -6)
            "trend_5d": (
                "UP" if closes[-1] > closes[-6]
                else "DOWN" if closes[-1] < closes[-6]
                else "EVEN"
            )
        }
        
        if "metrics" in data:
            payload.update(data["metrics"])
        if "benchmark_comparison" in data:
            payload["benchmark_comparison"] = data["benchmark_comparison"]

        return None, payload

    @staticmethod
    def validate_insider_trades(data: Optional[Dict[str, Any]]) -> Tuple[Optional[SeatOpinion], Optional[Dict[str, Any]]]:
        if not data or not data.get("filings"):
            return SeatOpinion(
                seat_name="Insider_Trades",
                signal=Signal.NO_READ,
                confidence=0.0,
                summary="No Form 4 insider filings recorded in the query window."
            ), None

        if not data.get("transactions_verified", False):
            return SeatOpinion(
                seat_name="Insider_Trades",
                signal=Signal.NO_READ,
                confidence=0.0,
                summary="Yahoo transaction labels are unverified; no directional insider signal is assigned."
            ), None

        buys = data.get("open_market_buys_usd", 0)
        sales = data.get("open_market_sales_usd", 0)

        # Do not default to EVEN if transactions occurred but their types or dollar values couldn't be parsed
        if buys == 0 and sales == 0:
            has_known = any(f.get("type") in ["PURCHASE", "SALE"] for f in data["filings"])
            if not has_known:
                return SeatOpinion(
                    seat_name="Insider_Trades",
                    signal=Signal.NO_READ,
                    confidence=0.0,
                    summary="Filings present, but transaction directions (open-market buy/sell) could not be verified."
                ), None
                
            return SeatOpinion(
                seat_name="Insider_Trades",
                signal=Signal.NO_READ,
                confidence=0.0,
                summary="Verified filings present, but transaction dollar amounts are missing or zero."
            ), None

        if buys > 2_000_000 and sales == 0:
            return SeatOpinion(
                seat_name="Insider_Trades",
                signal=Signal.UP,
                confidence=0.85,
                summary=f"Automated resolution: Strong cluster open-market insider buying (${buys:,.0f}) with zero sales."
            ), None

        return None, {
            "net_flow_usd": buys - sales,
            "buy_volume_usd": buys,
            "sale_volume_usd": sales,
            "total_transactions": len(data["filings"])
        }

    @staticmethod
    def validate_financials(data: Optional[Dict[str, Any]]) -> Tuple[Optional[SeatOpinion], Optional[Dict[str, Any]]]:
        if not data or data.get("pe_ratio") is None:
            return SeatOpinion(
                seat_name="Financials",
                signal=Signal.NO_READ,
                confidence=0.0,
                summary="Missing primary valuation metrics."
            ), None

        compact_payload = {k: v for k, v in data.items() if v is not None}
        return None, compact_payload

    @classmethod
    def evaluate(cls, seat_name: str, raw_data: Any) -> Tuple[Optional[SeatOpinion], Optional[Dict[str, Any]]]:
        dispatch = {
            "Price_Chart": cls.validate_price_chart,
            "Insider_Trades": cls.validate_insider_trades,
            "Financials": cls.validate_financials,
        }

        if seat_name in dispatch:
            return dispatch[seat_name](raw_data)

        if not raw_data:
            return SeatOpinion(
                seat_name=seat_name,
                signal=Signal.NO_READ,
                confidence=0.0,
                summary=f"No verified data provided for {seat_name}."
            ), None

        return None, raw_data

# =====================================================================
# 4. TICKER COUNCIL ENGINE
# =====================================================================

class TickerCouncil:
    def __init__(
        self,
        api_key: Optional[str] = None,
        seat_model: str = "gpt-4o-mini",
        judge_model: str = "gpt-4o",
        max_concurrent_calls: int = 5
    ):
        """
        Initializes OpenAI client with tiered models:
          - seat_model: Low-latency, high-efficiency model for mechanical extraction.
          - judge_model: High-reasoning frontier model for debate synthesis and calibrated probabilities.
        """
        self.client = AsyncOpenAI(api_key=api_key or os.environ.get("OPENAI_API_KEY"))
        self.seat_model = seat_model
        self.judge_model = judge_model
        self.semaphore = asyncio.Semaphore(max_concurrent_calls)

    async def _execute_seat(self, seat: Dict[str, str], raw_domain_data: Any) -> SeatOpinion:
        # Pre-prompt local gate check
        precomputed_opinion, compacted_payload = PrePromptValidator.evaluate(seat["name"], raw_domain_data)
        if precomputed_opinion is not None:
            return precomputed_opinion

        prompt = self._build_seat_prompt(seat, compacted_payload)

        async with self.semaphore:
            response = await self.client.beta.chat.completions.parse(
                model=self.seat_model,
                messages=[
                    {"role": "system", "content": "You are a quantitative specialist assessing directional market bias."},
                    {"role": "user", "content": prompt}
                ],
                response_format=SeatOpinion,
                temperature=0.1,
            )
            return response.choices[0].message.parsed

    async def _execute_debate(self, opinions: List[SeatOpinion]) -> DebateArguments:
        opinions_text = "\n".join(
            [f"- {o.seat_name}: {o.signal.value} (Confidence: {o.confidence:.2f}) -> {o.summary}" for o in opinions]
        )

        prompt = f"""
Specialist Seat Votes:
{opinions_text}

Debate Assignment:
1. Bull Advocate: Steelman the strongest positive case using purely the supportive seats.
2. Bear Advocate: Steelman the strongest downside risks using purely the critical seats.
3. Challenger (Red Team): Detect consensus bias among seats and aggressively identify latent tail risks, binary traps, or contradictory evidence.
"""

        async with self.semaphore:
            response = await self.client.beta.chat.completions.parse(
                model=self.judge_model,
                messages=[
                    {"role": "system", "content": "You are an adversarial debate moderator for an institutional investment council."},
                    {"role": "user", "content": prompt}
                ],
                response_format=DebateArguments,
                temperature=0.3,
            )
            return response.choices[0].message.parsed

    async def _execute_grand_master(
        self, ticker: str, opinions: List[SeatOpinion], debate: DebateArguments
    ) -> CouncilVerdict:
        context = {
            "ticker": ticker,
            "seat_opinions": [o.model_dump() for o in opinions],
            "adversarial_debate": debate.model_dump()
        }

        prompt = f"""
Comprehensive Council Briefing:
{json.dumps(context, indent=2)}

Task:
Synthesize the specialist votes and adversarial debate arguments into calibrated probability forecasts 
for 1_week, 3_months, and 1_year horizons. 
Do NOT emit binary buy/sell recommendations; output calibrated probabilities and explicit reasoning.
Use the same direction thresholds for each horizon: UP when probability_up > 0.55,
DOWN when probability_up < 0.45, and EVEN from 0.45 through 0.55 inclusive.
Treat these as model estimates, not empirically calibrated probabilities.
"""

        async with self.semaphore:
            response = await self.client.beta.chat.completions.parse(
                model=self.judge_model,
                messages=[
                    {"role": "system", "content": "You are the Grand Master Judge producing calibrated probabilistic outlooks."},
                    {"role": "user", "content": prompt}
                ],
                response_format=_ForecastWrapper,
                temperature=0.1,
            )
            parsed: _ForecastWrapper = response.choices[0].message.parsed

        return CouncilVerdict(
            ticker=ticker,
            forecasts=parsed.forecasts,
            seat_breakdown=opinions,
            debate=debate
        )

    async def run(self, ticker: str, raw_ingested_data: Dict[str, Any]) -> CouncilVerdict:
        """
        Orchestrates Phase 1 (Concurrent Seats), Phase 2 (Debate), and Phase 3 (Grand Master).
        """
        # Phase 1: Fan-out execution across all seats concurrently
        sources = raw_ingested_data.get("_metadata", {}).get("sources", {})
        seat_tasks = []
        for seat in SEAT_CONFIGS:
            source_status = sources.get(seat["name"], {}).get("status")
            unavailable = self._unavailable_source_opinion(
                seat["name"],
                source_status,
                sources.get(seat["name"], {}).get("error"),
            )
            if unavailable is not None:
                seat_tasks.append(asyncio.sleep(0, result=unavailable))
            else:
                seat_tasks.append(
                    self._execute_seat(seat, raw_ingested_data.get(seat["name"]))
                )
        opinions: List[SeatOpinion] = await asyncio.gather(*seat_tasks)

        # Phase 2: Adversarial debate
        debate: DebateArguments = await self._execute_debate(opinions)

        # Phase 3: Calibrated probability verdict
        verdict: CouncilVerdict = await self._execute_grand_master(ticker, opinions, debate)
        return verdict

    @staticmethod
    def _unavailable_source_opinion(
        seat_name: str, status: Optional[str], error: Optional[str] = None
    ) -> Optional[SeatOpinion]:
        if status not in {"missing", "error", "not_configured"}:
            return None
        detail = f" Source status: {status}."
        if error:
            detail += f" {error}"
        return SeatOpinion(
            seat_name=seat_name,
            signal=Signal.NO_READ,
            confidence=0.0,
            summary=f"No directional opinion because the source is unavailable.{detail}",
        )

    @staticmethod
    def _build_seat_prompt(seat: Dict[str, str], compacted_payload: Any) -> str:
        domain_specific_rule = (
            "- A positive or inverted yield curve is contextual macro information, not proof that this ticker will rise or fall.\n"
            if seat["name"] == "Macro_Economy"
            else ""
        )
        return (
            f"You are the {seat['name']} Specialist on an investment council.\n"
            f"Domain Focus: {seat['desc']}.\n\n"
            "Rules:\n"
            "- Analyze ONLY the provided domain data below.\n"
            "- Mention only metrics explicitly present in the supplied data; do not invent or imply missing metrics.\n"
            "- Do not extrapolate beyond your domain.\n"
            "- If data is conflicting or neutral, return EVEN.\n"
            f"{domain_specific_rule}\n"
            f"Domain Data:\n{json.dumps(compacted_payload, indent=2)}"
        )