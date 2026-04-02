"""
Bank Hapoalim Interest Charge Automation
=========================================

Usage
-----
  python main.py [PDF_PATH] [OUTPUT_XLSX]
                 --debit-spread DEBIT_SPREAD
                 --credit-spread CREDIT_SPREAD
                 [--transactions TX_FILE]

Arguments
---------
  PDF_PATH          Path to the Bank Hapoalim interest-charge PDF.
                    Default: "463560 גלובל.pdf"
  OUTPUT_XLSX       Path for the generated Excel workbook.
                    Default: "output.xlsx"
  --debit-spread    מרווח ריבית חובה in percent, e.g. 0.35 for 0.35 %.
  --credit-spread   מרווח ריבית זכות in percent, e.g. -0.60 for -0.60 %.
  --transactions    Path to a Bank Hapoalim transaction export file
                    (Excel .xlsx/.xls or CSV).  When supplied, Sheets 2 and 3
                    are populated automatically from this file.
                    When omitted, Sheets 2 and 3 are generated as empty
                    templates (headers and formulas only).

The script:
  1. Parses the interest-charge PDF produced by Bank Hapoalim.
  2. Fetches the Bank of Israel base interest rate from the BOI API for the
     statement period (internet connection required — no fallback).
  3. Optionally parses a transaction export file for Sheets 2/3.
  4. Writes a three-sheet Excel workbook:
       Sheet 1 – פירוט ריבית חובה לפי PDF       (populated from PDF)
       Sheet 2 – בנק                             (all transactions)
       Sheet 3 – חישוב החזר לפי יתרות בבנק      (transactions within period)
"""

from __future__ import annotations

import argparse
import logging
import sys
import os
from parsers.boi_rates import fetch_boi_rates
from parsers.hapoalim import HapoalimParser
from parsers.transactions import parse_transactions
from writers.excel_writer import ExcelWriter
from models.statement import Statement

# ── Logging setup ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract Bank Hapoalim interest-charge data from PDF to Excel"
    )
    parser.add_argument(
        "pdf_path",
        nargs="?",
        default="463560 גלובל.pdf",
        help="Path to the Bank Hapoalim interest-charge PDF",
    )
    parser.add_argument(
        "output_xlsx",
        nargs="?",
        default="output.xlsx",
        help="Path for the generated Excel workbook",
    )
    parser.add_argument(
        "--debit-spread", "-d",
        type=float,
        required=True,
        metavar="PCT",
        help=(
            "מרווח ריבית חובה in percent, e.g. 0.35 for 0.35 %%. "
            "Stored per account; added to the fetched BOI rate."
        ),
    )
    parser.add_argument(
        "--credit-spread", "-c",
        type=float,
        required=True,
        metavar="PCT",
        help=(
            "מרווח ריבית זכות in percent, e.g. -0.60 for -0.60 %%. "
            "Stored per account; added to the fetched BOI rate."
        ),
    )
    parser.add_argument(
        "--transactions", "-t",
        default=None,
        metavar="TX_FILE",
        help=(
            "Bank Hapoalim transaction export file (Excel or CSV). "
            "When omitted, Sheets 2/3 are written as empty templates."
        ),
    )
    return parser.parse_args()


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> int:
    args = parse_args()

    pdf_path     = args.pdf_path
    output_path  = args.output_xlsx
    tx_path      = args.transactions
    debit_spread = args.debit_spread / 100.0    # e.g. 0.35 → 0.0035
    credit_spread = args.credit_spread / 100.0  # e.g. -0.60 → -0.006

    # ── Validate inputs ───────────────────────────────────────────────────
    if not os.path.isfile(pdf_path):
        logger.error("PDF not found: %s", pdf_path)
        return 1

    if tx_path and not os.path.isfile(tx_path):
        logger.error("Transactions file not found: %s", tx_path)
        return 1

    # ── Parse interest-charge PDF ─────────────────────────────────────────
    logger.info("Parsing PDF: %s", pdf_path)
    hapoalim_parser = HapoalimParser()
    try:
        stmt: Statement = hapoalim_parser.parse(pdf_path)
    except Exception as exc:
        logger.error("Failed to parse PDF: %s", exc, exc_info=True)
        return 1

    # ── Fill in known / fallback values ───────────────────────────────────
    if not stmt.account_full and stmt.account_number:
        stmt.account_full = stmt.account_number
    if not stmt.company_name:
        stmt.company_name = "ארביטראז' גלובל אל פי"
    if not stmt.company_address:
        stmt.company_address = "דרך בגין 144"
    if not stmt.company_city:
        stmt.company_city = "תל אביב - יפו"
    if not stmt.company_id:
        stmt.company_id = "6492102"

    # ── Fetch BOI interest rate from Bank of Israel API ───────────────────
    if not stmt.period_from or not stmt.period_to:
        logger.error(
            "Could not determine billing period from PDF — "
            "cannot fetch BOI rate."
        )
        return 1

    try:
        rate1, rate2, change_date = fetch_boi_rates(
            stmt.period_from, stmt.period_to
        )
    except RuntimeError as exc:
        logger.error("BOI rate fetch failed: %s", exc)
        return 1

    stmt.bank_rate_p1    = rate1
    stmt.bank_rate_p2    = rate2        # None if rate did not change
    stmt.rate_change_date = change_date  # None if rate did not change

    # ── Apply user-supplied spreads ───────────────────────────────────────
    stmt.debit_margin  = debit_spread
    stmt.credit_margin = credit_spread

    # ── Parse transactions ────────────────────────────────────────────────
    if tx_path:
        logger.info("Parsing transactions: %s", tx_path)
        try:
            all_tx = parse_transactions(tx_path, stmt.period_from, stmt.period_to)
        except Exception as exc:
            logger.error("Failed to parse transactions file: %s", exc, exc_info=True)
            return 1

        if not all_tx:
            logger.warning(
                "No transactions found in %s — check that the file has recognised "
                "Hebrew column headers (תאריך, חובה, זכות, יתרה, …). "
                "Sheets 2/3 will be empty templates.", tx_path
            )

        logger.info("Transactions loaded: %d rows", len(all_tx))
        stmt.transactions = all_tx
    else:
        logger.info(
            "No transactions file supplied — Sheets 2/3 will be empty templates. "
            "Run with --transactions TX_FILE to populate them automatically."
        )
        stmt.transactions = []

    # ── Log summary ───────────────────────────────────────────────────────
    logger.info("Account      : %s", stmt.account_full)
    logger.info("Period       : %s – %s", stmt.period_from, stmt.period_to)
    logger.info("Charge date  : %s", stmt.charge_date)
    logger.info("Tier rows    : %d", len(stmt.tier_rows))
    logger.info("Tier 1 total : %.2f", stmt.total_tier1)
    logger.info("Excess total : %.2f", stmt.total_excess)
    logger.info("Total charged: %.2f", stmt.total_charged)

    # ── Write Excel ────────────────────────────────────────────────────────
    logger.info("Writing Excel: %s", output_path)
    try:
        writer = ExcelWriter(stmt)
        writer.write(output_path)
    except Exception as exc:
        logger.error("Failed to write Excel: %s", exc, exc_info=True)
        return 1

    logger.info("Done. Output saved to: %s", output_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
