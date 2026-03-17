"""
Parser for Bank Hapoalim interest-charge PDFs (פירוט חיוב הריבית).

PDF layout — two column layouts coexist in the same document:
────────────────────────────────────────────────────────────────
2-tier periods (מדרגה 1 + חריגה):
  x >= T1_X (260) : מדרגה 1 column
  x <  T1_X (260) : חריגה  column

3-tier periods (מדרגה 1 + מדרגה 2 + חריגה):
  x >= T1_X (260)               : מדרגה 1 column
  T2_X (165) <= x < T1_X (260) : מדרגה 2 column
  x <  T2_X (165)               : חריגה  column

Framework column (both layouts): x >= FW_X (340)

For each billing period, rows appear in order:
  1. Period-dates row  – two DD/MM/YY dates
  2. Framework row     – one or two large integers (fw1 at x>=FW_X;
                         fw2 at T1_X<=x<FW_X for 3-tier)
  3. Tier-amount row   – one or two large integers (no number at x>=FW_X)
  4. Rates row         – two or three "N.NN%" tokens
  5. Debit-numbers row – two or three large numbers
  6. Interest row      – two or three decimal numbers → emit TierRow objects

Summary block appears after the last period:
  • Totals row   – two decimal numbers (tier1 total, excess total)
  • Charge row   – one decimal + one date (total charged, charge date)
"""

from __future__ import annotations

import re
import logging
from enum import IntEnum, auto
from typing import List, Optional, Dict

import pdfplumber
from pdfminer.layout import LAParams

from parsers.base import BaseParser
from models.statement import Statement, TierRow

logger = logging.getLogger(__name__)

# ── X-position thresholds (calibrated from sample PDFs) ──────────────────────
FW_X = 340.0   # framework column:              x >= FW_X
T1_X = 260.0   # tier-1 / (tier-2 or excess) boundary
T2_X = 165.0   # tier-2 / excess boundary (3-tier periods only)

# ── Minimal LAParams — disables expensive layout analysis pdfminer does by
#    default (column detection, text-flow ordering).  This alone can cut
#    pdfminer's working-set by 30–50 % on multi-column PDFs. ──────────────────
_LAPARAMS = LAParams(boxes_flow=None, detect_vertical=False)

# ── Regex tokens ──────────────────────────────────────────────────────────────
_DATE_RE    = re.compile(r"^(\d{2}/\d{2}/\d{2,4})$")
_RATE_RE    = re.compile(r"^(\d+\.\d+)%$")
_BIG_NUM_RE = re.compile(r"^(\d+(?:,\d{3})+(?:\.\d+)?)$")
_DEC_RE     = re.compile(r"^(\d+\.\d{2})$")


def _is_cid(text: str) -> bool:
    return "(cid:" in text or (len(text) > 0 and ord(text[0]) > 127)


def _to_float(text: str) -> Optional[float]:
    try:
        return float(text.replace(",", "").replace("%", ""))
    except ValueError:
        return None


def _group_lines(words: list, y_tol: float = 4.0) -> List[List[dict]]:
    """Group pdfplumber word dicts by vertical position into lines."""
    if not words:
        return []
    by_y: Dict[float, List[dict]] = {}
    for w in words:
        key = round(w["top"] / y_tol) * y_tol
        by_y.setdefault(key, []).append(w)
    lines = []
    for y in sorted(by_y):
        line = sorted(by_y[y], key=lambda w: w["x0"])
        lines.append(line)
    return lines


def _readable(words: List[dict]) -> List[dict]:
    return [w for w in words if not _is_cid(w["text"]) and w["text"].strip("-").strip()]


def _col(x: float, has_t2: bool) -> int:
    """
    Return column index for a data token at horizontal position x.

    2 = מדרגה 1  (tier-1)
    3 = מדרגה 2  (tier-2, only in 3-tier layout)
    4 = חריגה    (excess)
    """
    if x >= T1_X:                return 2   # tier-1
    if has_t2 and x >= T2_X:    return 3   # tier-2 (3-tier layout)
    return 4                                # excess


# ── State machine states ──────────────────────────────────────────────────────

class _S(IntEnum):
    HEADER       = auto()   # scanning document header
    PERIOD_DATES = auto()   # scanning for period data rows
    FRAMEWORK    = auto()   # found framework row, expect tier-amounts next
    TIER_AMOUNT  = auto()   # found tier amounts, expect rates next
    RATES        = auto()   # found rates, expect debit numbers next
    DEBIT        = auto()   # found debit numbers, expect interest next


class HapoalimParser(BaseParser):
    """Parse a Bank Hapoalim interest-charge PDF -> :class:`~models.statement.Statement`."""

    def parse(self, file_path: str) -> Statement:
        stmt = Statement()
        self._init_sm(stmt)

        # Process one page at a time so only one page's words are in memory at
        # once.  State persists across pages via instance variables (_init_sm /
        # _feed).  flush_cache() drops pdfplumber's per-page cached objects
        # (layout tree, char list, etc.) immediately after we're done with each
        # page.
        with pdfplumber.open(file_path, laparams=_LAPARAMS) as pdf:
            for page in pdf.pages:
                words = page.extract_words(
                    x_tolerance=1, y_tolerance=3, keep_blank_chars=False
                )
                lines = _group_lines(words, y_tol=4.0)
                del words           # release word-dict list before next page
                self._feed(lines)
                del lines
                page.flush_cache()  # drop pdfplumber's page-level cache

        self._calculate_totals(stmt)
        logger.info(
            "Parsed %d tier rows; total_tier1=%.2f excess=%.2f total_charged=%.2f",
            len(stmt.tier_rows), stmt.total_tier1, stmt.total_excess, stmt.total_charged,
        )
        return stmt

    # ── State-machine initialisation ──────────────────────────────────────────

    def _init_sm(self, stmt: Statement) -> None:
        """Initialise all state-machine variables before processing page 1."""
        self._stmt      = stmt
        self._state     = _S.HEADER
        self._pf        = ""
        self._pt        = ""
        self._fw1       = ""
        self._fw2       = ""
        self._tier_t1: Optional[float] = None
        self._tier_t2: Optional[float] = None
        self._rate_t1   = 0.0
        self._rate_t2   = 0.0
        self._rate_ex   = 0.0
        self._debit_t1  = 0.0
        self._debit_t2  = 0.0
        self._debit_ex  = 0.0
        self._has_t2    = False

    # ── Incremental page feed (state persists across calls) ───────────────────

    def _feed(self, lines: List[List[dict]]) -> None:
        """
        Run the state machine over *lines* (one page's worth).
        Instance variables carry state from the previous page so periods that
        span a page boundary are handled correctly.
        """
        stmt = self._stmt

        # Copy instance state to locals for the inner loop (faster attribute
        # access and cleaner code); save back at the end.
        state    = self._state
        curr_pf  = self._pf
        curr_pt  = self._pt
        curr_fw1 = self._fw1
        curr_fw2 = self._fw2
        curr_tier_t1 = self._tier_t1
        curr_tier_t2 = self._tier_t2
        curr_rate_t1 = self._rate_t1
        curr_rate_t2 = self._rate_t2
        curr_rate_ex = self._rate_ex
        curr_debit_t1 = self._debit_t1
        curr_debit_t2 = self._debit_t2
        curr_debit_ex = self._debit_ex
        has_t2   = self._has_t2

        for line in lines:
            rw = _readable(line)
            if not rw:
                continue

            dates   = [w for w in rw if _DATE_RE.match(w["text"])]
            rates   = [w for w in rw if _RATE_RE.match(w["text"])]
            numbers = [w for w in rw if _BIG_NUM_RE.match(w["text"]) or _DEC_RE.match(w["text"])]

            # ── HEADER: account number, print date, subject-line (3 dates) ────
            if state == _S.HEADER:
                # Account number: 6-digit token optionally followed by branch code
                if not stmt.account_number:
                    texts = [w["text"] for w in rw]
                    for i, t in enumerate(texts):
                        if re.match(r"^\d{6}$", t):
                            stmt.account_number = t
                            branch = (
                                texts[i + 1]
                                if i + 1 < len(texts) and re.match(r"^\d+$", texts[i + 1])
                                else ""
                            )
                            stmt.account_full = f"{t} {branch}".strip()
                            break

                # Subject line: 3 date tokens (period_from, period_to, charge_date)
                if len(dates) == 3:
                    sorted_d = sorted(dates, key=lambda w: w["x0"])
                    stmt.period_from = sorted_d[0]["text"]
                    stmt.period_to   = sorted_d[1]["text"]
                    stmt.charge_date = sorted_d[2]["text"]
                    state = _S.PERIOD_DATES
                    continue

                # Single date before subject line = print date
                if not stmt.print_date and len(dates) == 1:
                    stmt.print_date = dates[0]["text"]

            # ── New period: exactly 2 date tokens ─────────────────────────────
            if state in (_S.PERIOD_DATES, _S.HEADER) and len(dates) == 2:
                sorted_d = sorted(dates, key=lambda w: w["x0"])
                curr_pf = sorted_d[0]["text"]
                curr_pt = sorted_d[1]["text"]
                curr_fw1 = curr_fw2 = ""
                curr_tier_t1 = curr_tier_t2 = None
                has_t2 = False
                state = _S.PERIOD_DATES
                continue

            # ── FRAMEWORK row: must contain a number at x >= FW_X ─────────────
            if state == _S.PERIOD_DATES:
                fw_nums = [n for n in numbers if n["x0"] >= FW_X]
                if fw_nums:
                    # Main framework (מסגרת חח"ד): rightmost number (highest x)
                    fw1_w = max(fw_nums, key=lambda w: w["x0"])
                    v1 = _to_float(fw1_w["text"])
                    curr_fw1 = f"{int(v1):,}" if v1 and v1 == int(v1) else fw1_w["text"]

                    # Additional framework (מסגרת נוספת, 3-tier only):
                    # number in range T1_X <= x < FW_X on the same line
                    fw2_candidates = [n for n in numbers if T1_X <= n["x0"] < FW_X]
                    if fw2_candidates:
                        fw2_w = fw2_candidates[0]
                        v2 = _to_float(fw2_w["text"])
                        curr_fw2 = f"{int(v2):,}" if v2 and v2 == int(v2) else fw2_w["text"]
                        has_t2 = True
                    else:
                        curr_fw2 = ""
                        has_t2 = False

                    state = _S.FRAMEWORK
                    continue

            # ── TIER_AMOUNT row: numbers in T1/T2 range, no rates ─────────────
            if state == _S.FRAMEWORK and numbers and not rates:
                t1_nums = [n for n in numbers if T1_X <= n["x0"] < FW_X]
                t2_nums = [n for n in numbers if T2_X <= n["x0"] < T1_X]
                if t1_nums:
                    curr_tier_t1 = _to_float(t1_nums[0]["text"])
                if t2_nums:
                    curr_tier_t2 = _to_float(t2_nums[0]["text"])
                state = _S.TIER_AMOUNT
                continue

            # ── RATES row ──────────────────────────────────────────────────────
            if state == _S.TIER_AMOUNT and rates:
                curr_rate_t1 = curr_rate_t2 = curr_rate_ex = 0.0
                for r in rates:
                    v = _to_float(r["text"])
                    if v is None:
                        continue
                    v /= 100.0
                    col = _col(r["x0"], has_t2)
                    if col == 2:   curr_rate_t1 = v
                    elif col == 3: curr_rate_t2 = v
                    else:          curr_rate_ex  = v
                # Fill missing rates (fall back to tier-1 rate when absent)
                if curr_rate_ex == 0.0 and curr_rate_t1 != 0.0:
                    curr_rate_ex = curr_rate_t1
                if has_t2 and curr_rate_t2 == 0.0 and curr_rate_t1 != 0.0:
                    curr_rate_t2 = curr_rate_t1
                state = _S.RATES
                continue

            # ── DEBIT-NUMBERS row ─────────────────────────────────────────────
            if state == _S.RATES and numbers:
                curr_debit_t1 = curr_debit_t2 = curr_debit_ex = 0.0
                for n in numbers:
                    v = _to_float(n["text"])
                    if v is None:
                        continue
                    col = _col(n["x0"], has_t2)
                    if col == 2:   curr_debit_t1 = v
                    elif col == 3: curr_debit_t2 = v
                    else:          curr_debit_ex  = v
                state = _S.DEBIT
                continue

            # ── INTEREST row: emit TierRow objects ────────────────────────────
            if state == _S.DEBIT and numbers:
                int_t1 = int_t2 = int_ex = 0.0
                for n in numbers:
                    v = _to_float(n["text"])
                    if v is None:
                        continue
                    col = _col(n["x0"], has_t2)
                    if col == 2:   int_t1 = v
                    elif col == 3: int_t2 = v
                    else:          int_ex  = v

                stmt.tier_rows.append(TierRow(
                    period_from=curr_pf, period_to=curr_pt,
                    framework=curr_fw1, framework2=curr_fw2,
                    tier_type="מדרגה 1",
                    tier_amount=curr_tier_t1,
                    rate=curr_rate_t1,
                    debit_numbers=curr_debit_t1,
                    interest=int_t1,
                ))
                if has_t2:
                    stmt.tier_rows.append(TierRow(
                        period_from=curr_pf, period_to=curr_pt,
                        framework=curr_fw1, framework2=curr_fw2,
                        tier_type="מדרגה 2",
                        tier_amount=curr_tier_t2,
                        rate=curr_rate_t2,
                        debit_numbers=curr_debit_t2,
                        interest=int_t2,
                    ))
                stmt.tier_rows.append(TierRow(
                    period_from=curr_pf, period_to=curr_pt,
                    framework=curr_fw1, framework2=curr_fw2,
                    tier_type="חריגה",
                    tier_amount=None,
                    rate=curr_rate_ex,
                    debit_numbers=curr_debit_ex,
                    interest=int_ex,
                ))

                # Reset for next period
                curr_pf = curr_pt = ""
                curr_fw1 = curr_fw2 = ""
                curr_tier_t1 = curr_tier_t2 = None
                has_t2 = False
                state = _S.PERIOD_DATES
                continue

            # ── Post-table: totals row (2 numbers, no dates) ──────────────────
            if state == _S.PERIOD_DATES and len(numbers) == 2 and len(dates) == 0:
                sorted_n = sorted(numbers, key=lambda w: w["x0"])
                if not stmt.total_tier1:
                    stmt.total_tier1  = _to_float(sorted_n[1]["text"]) or 0.0
                    stmt.total_excess = _to_float(sorted_n[0]["text"]) or 0.0
                continue

            # ── Post-table: charge row (1 number + 1 date) ────────────────────
            if state == _S.PERIOD_DATES and len(numbers) == 1 and len(dates) == 1:
                if not stmt.total_charged:
                    stmt.total_charged = _to_float(numbers[0]["text"]) or 0.0
                if not stmt.charge_date:
                    stmt.charge_date = dates[0]["text"]
                continue

        # ── Save state back to instance for the next page ─────────────────────
        self._state    = state
        self._pf       = curr_pf
        self._pt       = curr_pt
        self._fw1      = curr_fw1
        self._fw2      = curr_fw2
        self._tier_t1  = curr_tier_t1
        self._tier_t2  = curr_tier_t2
        self._rate_t1  = curr_rate_t1
        self._rate_t2  = curr_rate_t2
        self._rate_ex  = curr_rate_ex
        self._debit_t1 = curr_debit_t1
        self._debit_t2 = curr_debit_t2
        self._debit_ex = curr_debit_ex
        self._has_t2   = has_t2

    # ─────────────────────────────────────────────────────────────────────────

    def _calculate_totals(self, stmt: Statement) -> None:
        """Fill missing totals by summing tier_rows."""
        if not stmt.total_tier1:
            stmt.total_tier1 = round(
                sum(r.interest for r in stmt.tier_rows if r.tier_type == "מדרגה 1"), 2
            )
        if not stmt.total_excess:
            stmt.total_excess = round(
                sum(r.interest for r in stmt.tier_rows if r.tier_type == "חריגה"), 2
            )
        if not stmt.total_charged:
            total_t2 = round(
                sum(r.interest for r in stmt.tier_rows if r.tier_type == "מדרגה 2"), 2
            )
            stmt.total_charged = round(stmt.total_tier1 + total_t2 + stmt.total_excess, 2)
