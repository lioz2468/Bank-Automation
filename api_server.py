"""
FastAPI backend for Bank Hapoalim Interest Analyzer.

Deploy this on Render (or any Python host).
The frontend (Next.js on Vercel) calls POST /process.
"""

from __future__ import annotations

import gc
import io
import logging
import os
import tempfile

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from models.statement import Statement
from parsers.boi_rates import fetch_boi_rates
from parsers.hapoalim import HapoalimParser
from parsers.transactions import parse_transactions
from writers.excel_writer import ExcelWriter

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-8s %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(title="Bank Automation API")

# Allow requests from any origin (lock this down to your Vercel URL in production)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/process")
async def process(
    pdf_file: UploadFile = File(..., description="Bank Hapoalim interest-charge PDF"),
    tx_file: UploadFile | None = File(None, description="Transactions Excel/CSV (optional)"),
    debit_spread: float = Form(0.35, description="Debit spread in percent, e.g. 0.35"),
    credit_spread: float = Form(-0.60, description="Credit spread in percent, e.g. -0.60"),
) -> StreamingResponse:
    """
    Parse the interest-charge PDF, fetch BOI rates, optionally parse transactions,
    and return a generated Excel workbook as a download.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        # ── Save uploaded files ────────────────────────────────────────────
        pdf_path = os.path.join(tmpdir, "statement.pdf")
        pdf_bytes = await pdf_file.read()
        if not pdf_bytes:
            raise HTTPException(400, "Uploaded PDF is empty")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        tx_path: str | None = None
        if tx_file and tx_file.filename:
            ext = os.path.splitext(tx_file.filename)[1].lower()
            if ext not in (".xlsx", ".xls", ".csv"):
                raise HTTPException(400, f"Unsupported transactions file type: {ext!r}")
            tx_path = os.path.join(tmpdir, f"transactions{ext}")
            with open(tx_path, "wb") as f:
                f.write(await tx_file.read())

        # ── Parse interest-charge PDF ──────────────────────────────────────
        logger.info("Parsing PDF: %s", pdf_file.filename)
        try:
            stmt: Statement = HapoalimParser().parse(pdf_path)
        except Exception as exc:
            logger.error("PDF parse failed: %s", exc)
            raise HTTPException(400, f"Failed to parse PDF: {exc}") from exc

        # ── Fill default company info (fallbacks from existing main.py) ────
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

        # Release pdfminer/pdfplumber memory before the next heavy step.
        # CPython's cycle collector won't necessarily run between the parse
        # and the BOI HTTP call, so trigger it explicitly here.
        gc.collect()

        # ── Validate period dates ──────────────────────────────────────────
        if not stmt.period_from or not stmt.period_to:
            raise HTTPException(400, "Could not determine billing period from PDF")

        # ── Fetch BOI interest rate ────────────────────────────────────────
        logger.info("Fetching BOI rate for %s – %s", stmt.period_from, stmt.period_to)
        try:
            rate1, rate2, change_date = fetch_boi_rates(stmt.period_from, stmt.period_to)
        except RuntimeError as exc:
            logger.error("BOI rate fetch failed: %s", exc)
            raise HTTPException(502, f"BOI rate fetch failed: {exc}") from exc

        stmt.bank_rate_p1 = rate1
        stmt.bank_rate_p2 = rate2
        stmt.rate_change_date = change_date
        stmt.debit_margin = debit_spread / 100.0
        stmt.credit_margin = credit_spread / 100.0

        # ── Parse transactions (optional) ──────────────────────────────────
        if tx_path:
            logger.info("Parsing transactions: %s", tx_file.filename if tx_file else "")
            try:
                stmt.transactions = parse_transactions(tx_path)
            except Exception as exc:
                logger.error("Transactions parse failed: %s", exc)
                raise HTTPException(400, f"Failed to parse transactions file: {exc}") from exc
            logger.info("Loaded %d transaction rows", len(stmt.transactions))
        else:
            stmt.transactions = []

        # ── Write Excel ────────────────────────────────────────────────────
        output_path = os.path.join(tmpdir, "output.xlsx")
        try:
            ExcelWriter(stmt).write(output_path)
        except Exception as exc:
            logger.error("Excel write failed: %s", exc)
            raise HTTPException(500, f"Failed to write Excel: {exc}") from exc

        # Read into memory before tempdir is cleaned up
        with open(output_path, "rb") as f:
            content = f.read()

    logger.info("Done — returning %.1f KB Excel file", len(content) / 1024)
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="bank_interest_output.xlsx"'},
    )
