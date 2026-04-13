"""
Fetch the Bank of Israel base interest rate (ריבית בנק ישראל) from the
official BOI SDMX API for a given billing period.

API endpoint
------------
https://edge.boi.gov.il/FusionEdgeServer/sdmx/v2/data/dataflow/BOI.STATISTICS/BR/1.0

The series ``MNT_RIB_BOI_D`` (daily) contains the official BOI base rate in
percentage-point units (e.g. 4.5 = 4.5 %).  This module converts to decimal
fractions (0.045) before returning.

Returns
-------
(rate1, rate2, change_date)
  rate1       – rate in effect at period start, as a decimal fraction
  rate2       – rate in effect at period end if the rate changed mid-period,
                else None
  change_date – "DD.MM" string of the first date the new rate applied,
                e.g. "27.11", or None if no change occurred
"""

from __future__ import annotations

import csv
import io
import logging
from datetime import datetime, date, timedelta
from typing import Optional, Tuple

import requests

logger = logging.getLogger(__name__)

_API_URL     = (
    "https://edge.boi.gov.il/FusionEdgeServer/sdmx/v2"
    "/data/dataflow/BOI.STATISTICS/BR/1.0"
)
_SERIES_CODE = "MNT_RIB_BOI_D"
_TIMEOUT     = 20   # seconds


def _parse_period_date(date_str) -> Optional[date]:
    """Parse "DD/MM/YY" or "DD/MM/YYYY" → date object. Also accepts a date/datetime directly."""
    if isinstance(date_str, date):
        return date_str if not isinstance(date_str, datetime) else date_str.date()
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            pass
    return None


def fetch_boi_rates(
    period_from,
    period_to,
) -> Tuple[float, Optional[float], Optional[str]]:
    """
    Fetch BOI base interest rate(s) for the given billing period.

    Parameters
    ----------
    period_from : str
        Start of the billing period, "DD/MM/YY" or "DD/MM/YYYY".
    period_to : str
        End of the billing period, "DD/MM/YY" or "DD/MM/YYYY".

    Returns
    -------
    rate1 : float
        BOI rate in effect at the start of the period (decimal, e.g. 0.045).
    rate2 : float or None
        BOI rate at the end of the period if the rate changed mid-period,
        else None.
    change_date : str or None
        "DD.MM" of the first day the new rate applied (e.g. "27.11"),
        or None if the rate did not change during the period.

    Raises
    ------
    RuntimeError
        If the API is unreachable, times out, or returns no data.
        The caller should print the message and abort — there is no fallback.
    """
    from_date = _parse_period_date(period_from)
    to_date   = _parse_period_date(period_to)

    if from_date is None or to_date is None:
        raise RuntimeError(
            f"Cannot parse billing period dates: {period_from!r} / {period_to!r}"
        )

    # Extend lookback 120 days to capture the previous quarter's starting rate.
    # The bank uses that earlier rate as P1 when there was a pre-period rate change.
    fetch_from = from_date - timedelta(days=120)
    start_str  = fetch_from.strftime("%Y-%m-%d")
    end_str    = to_date.strftime("%Y-%m-%d")

    logger.info("Fetching BOI rate from API for %s – %s …", period_from, period_to)

    try:
        resp = requests.get(
            _API_URL,
            params={"startperiod": start_str, "endperiod": end_str, "format": "csv"},
            timeout=_TIMEOUT,
        )
    except requests.exceptions.ConnectionError as exc:
        raise RuntimeError(
            "Cannot reach the Bank of Israel API (connection error). "
            "Please check your internet connection and try again."
        ) from exc
    except requests.exceptions.Timeout as exc:
        raise RuntimeError(
            f"Bank of Israel API did not respond within {_TIMEOUT} s. "
            "Please check your internet connection and try again."
        ) from exc
    except Exception as exc:
        raise RuntimeError(
            f"Unexpected error while fetching BOI rate: {exc}"
        ) from exc

    if resp.status_code != 200:
        raise RuntimeError(
            f"Bank of Israel API returned HTTP {resp.status_code}. "
            "Cannot fetch the interest rate — check your internet connection."
        )

    if not resp.text.strip():
        raise RuntimeError(
            f"Bank of Israel API returned an empty response for period "
            f"{period_from} – {period_to}."
        )

    # ── Parse CSV ─────────────────────────────────────────────────────────
    rows: list[tuple[date, float]] = []
    reader = csv.DictReader(io.StringIO(resp.text))
    for row in reader:
        if row.get("SERIES_CODE") != _SERIES_CODE:
            continue
        val_str = row.get("OBS_VALUE", "").strip()
        if not val_str:
            continue
        try:
            dt  = datetime.strptime(row["TIME_PERIOD"], "%Y-%m-%d").date()
            val = float(val_str) / 100.0   # percentage → decimal
        except (ValueError, KeyError):
            continue
        rows.append((dt, val))

    if not rows:
        raise RuntimeError(
            f"No BOI base-rate data found for {period_from} – {period_to}. "
            "The Bank of Israel API returned no observations for series "
            f"'{_SERIES_CODE}'. "
            "Please check your internet connection and try again."
        )

    rows.sort(key=lambda x: x[0])

    # ── Identify rates ────────────────────────────────────────────────────
    # prev_quarter_start: first day of the quarter BEFORE the billing period.
    # The bank uses the BOI rate in effect at that date as P1 when the rate
    # changed between then and period_from.
    q_month   = from_date.month
    pq_month  = q_month - 3 if q_month > 3 else q_month + 9
    pq_year   = from_date.year if q_month > 3 else from_date.year - 1
    prev_quarter_start = date(pq_year, pq_month, 1)

    # r_pq:    BOI rate in effect at the start of the previous quarter
    # r_start: BOI rate in effect at the start of the billing period
    r_pq:    Optional[float] = None
    r_start: Optional[float] = None
    for dt, val in rows:
        if dt <= prev_quarter_start:
            r_pq = val          # last known rate on/before prev_quarter_start
        if dt <= from_date:
            r_start = val       # last known rate on/before period start

    rate1:       float         = r_start or rows[0][1]  # type: ignore[assignment]
    rate2:       Optional[float] = None
    change_date: Optional[str]   = None

    if (r_pq is not None and r_start is not None
            and abs(r_pq - r_start) > 1e-9):
        # Pre-period rate change detected.
        # Bank convention: P1 = previous-quarter starting rate,
        #                  P2 = current-period starting rate,
        #                  change_date = first within-period API rate change.
        rate1 = r_pq
        prev_val = r_start
        for dt, val in rows:
            if dt < from_date:
                continue
            if abs(val - prev_val) > 1e-9:
                rate2       = r_start   # P2 = period-start rate, not the new API value
                change_date = dt.strftime("%d.%m")
                break
            prev_val = val
        if rate2 is None:
            # No within-period API change: report single-rate period
            rate1 = r_start
    else:
        # No pre-period change: scan within-period for a rate change (original logic)
        for dt, val in rows:
            if dt < from_date:
                continue
            if abs(val - rate1) > 1e-9:
                rate2       = val
                change_date = dt.strftime("%d.%m")
                break

    if rate2 is not None:
        logger.info(
            "BOI rate: %.2f%% -> %.2f%% (change on %s)",
            rate1 * 100, rate2 * 100, change_date,
        )
    else:
        logger.info("BOI rate: %.2f%% (no change during period)", rate1 * 100)

    return rate1, rate2, change_date
