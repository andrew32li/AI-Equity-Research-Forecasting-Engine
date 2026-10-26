"""Local SQLite persistence and outcome scoring for council forecasts."""

import sqlite3
from contextlib import closing
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from council import CouncilVerdict


HORIZON_DAYS = {
    "1_week": 7,
    "3_months": 90,
    "1_year": 365,
}

CloseLookup = Callable[[str, date, date], List[Tuple[date, float]]]


class ForecastStore:
    def __init__(self, database_path: Optional[Path] = None):
        self.database_path = database_path or Path(__file__).with_name("forecasts.sqlite3")
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database_path)
        connection.row_factory = sqlite3.Row
        return connection

    def _initialize(self) -> None:
        with closing(self._connect()) as connection, connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS forecasts (
                    id INTEGER PRIMARY KEY,
                    ticker TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    horizon TEXT NOT NULL,
                    target_date TEXT NOT NULL,
                    direction TEXT NOT NULL,
                    probability_up REAL NOT NULL CHECK (
                        probability_up >= 0 AND probability_up <= 1
                    ),
                    reasoning TEXT NOT NULL,
                    baseline_price REAL,
                    evaluated_at TEXT,
                    realized_price REAL,
                    realized_return REAL,
                    actual_up INTEGER,
                    brier_score REAL
                )
                """
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_forecasts_due "
                "ON forecasts(evaluated_at, target_date, ticker)"
            )

    def record_verdict(
        self,
        verdict: CouncilVerdict,
        baseline_price: Optional[float],
        created_at: Optional[datetime] = None,
    ) -> int:
        timestamp = created_at or datetime.now(timezone.utc)
        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(tzinfo=timezone.utc)
        created_at_text = timestamp.astimezone(timezone.utc).isoformat()
        forecast_date = timestamp.date()
        rows = []
        for forecast in verdict.forecasts:
            days = HORIZON_DAYS.get(forecast.horizon)
            if days is None:
                raise ValueError(f"Unsupported forecast horizon: {forecast.horizon}")
            rows.append(
                (
                    verdict.ticker.upper(),
                    created_at_text,
                    forecast.horizon,
                    (forecast_date + timedelta(days=days)).isoformat(),
                    forecast.direction.value,
                    forecast.probability_up,
                    forecast.reasoning,
                    baseline_price,
                )
            )

        with closing(self._connect()) as connection, connection:
            cursor = connection.executemany(
                """
                INSERT INTO forecasts (
                    ticker, created_at, horizon, target_date, direction,
                    probability_up, reasoning, baseline_price
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                rows,
            )
            return cursor.rowcount

    def score_due_forecasts(
        self,
        close_lookup: CloseLookup,
        as_of_date: Optional[date] = None,
        max_calendar_days_late: int = 7,
    ) -> Dict[str, Optional[float]]:
        today = as_of_date or date.today()
        with closing(self._connect()) as connection, connection:
            due = connection.execute(
                """
                SELECT id, ticker, target_date, baseline_price, probability_up
                FROM forecasts
                WHERE evaluated_at IS NULL
                  AND baseline_price IS NOT NULL
                  AND target_date <= ?
                ORDER BY target_date
                """,
                (today.isoformat(),),
            ).fetchall()

            scored = 0
            for row in due:
                target_date = date.fromisoformat(row["target_date"])
                last_allowed_date = target_date + timedelta(days=max_calendar_days_late)
                closes = close_lookup(row["ticker"], target_date, last_allowed_date)
                valid_closes = sorted(
                    (close_date, close)
                    for close_date, close in closes
                    if target_date <= close_date <= last_allowed_date and close > 0
                )
                if not valid_closes:
                    continue

                _, realized_price = valid_closes[0]
                baseline_price = row["baseline_price"]
                realized_return = realized_price / baseline_price - 1
                actual_up = int(realized_return > 0)
                brier_score = (row["probability_up"] - actual_up) ** 2
                connection.execute(
                    """
                    UPDATE forecasts
                    SET evaluated_at = ?, realized_price = ?, realized_return = ?,
                        actual_up = ?, brier_score = ?
                    WHERE id = ?
                    """,
                    (
                        datetime.now(timezone.utc).isoformat(),
                        realized_price,
                        realized_return,
                        actual_up,
                        brier_score,
                        row["id"],
                    ),
                )
                scored += 1

            average = connection.execute(
                "SELECT AVG(brier_score) FROM forecasts WHERE brier_score IS NOT NULL"
            ).fetchone()[0]

        return {
            "due": len(due),
            "scored": scored,
            "average_brier_score": average,
        }
