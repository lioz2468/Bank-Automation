"""
Data models for Bank Hapoalim interest charge statements.
"""
from dataclasses import dataclass, field
from typing import List, Optional



@dataclass
class TierRow:
    """
    One tier (step) row from the interest charge table.
    Each billing period has two or three tier rows:
      - "מדרגה 1"  (first tier, within overdraft limit)
      - "מדרגה 2"  (second tier, optional – within additional framework)
      - "חריגה"    (excess, above overdraft limit)
    """
    period_from: str          # "01/10/25"
    period_to: str            # "31/10/25"
    framework: str            # Display string e.g. "1,000,000" (מסגרת חח"ד)
    framework2: str           # Display string e.g. "265,000,000" (מסגרת נוספת); "" if none
    tier_type: str            # "מדרגה 1", "מדרגה 2", or "חריגה"
    tier_amount: Optional[float]  # Numeric cap for מדרגה 1/2; None for חריגה
    rate: float               # Actual rate charged, e.g. 0.049
    debit_numbers: float      # מספרי חובה (sum of daily debit balances)
    interest: float           # ריבית חובה in ₪


@dataclass
class BankTransaction:
    """One row from the bank transaction register."""
    date: str                 # "31.12.2025"
    code: str                 # Operation code e.g. "469"
    operation: str            # Hebrew operation name
    details: str              # Free-text details
    reference: str            # Reference / אסמכתא
    batch: str                # Batch / צרור
    debit: Optional[float]    # Debit amount (positive = money leaving account)
    credit: Optional[float]   # Credit amount
    balance: Optional[float]  # Running balance


@dataclass
class Statement:
    """
    Complete representation of a Bank Hapoalim interest charge document.

    Sheet 1 data (from interest-charge PDF):
      tier_rows, totals, rate parameters.

    Sheet 2/3 data (from full account-statement PDF – separate source):
      transactions.
    """
    # ── Bank / account identification ──────────────────────────────────────
    bank_name: str = "בנק הפועלים"
    branch: str = "סניף מרכז תפעולי עסקי 600"
    phone: str = "03-6532407"
    account_number: str = ""          # e.g. "463560"
    account_full: str = ""            # e.g. "463560 600"
    company_name: str = ""            # e.g. "ארביטראז' גלובל אל פי"
    company_address: str = ""         # e.g. "דרך בגין 144"
    company_city: str = ""            # e.g. "תל אביב - יפו"
    company_id: str = ""              # e.g. "6492102"
    print_date: str = ""              # "02/01/26"

    # ── Billing period ──────────────────────────────────────────────────────
    period_from: str = ""             # "01/10/25"
    period_to: str = ""               # "31/12/25"
    charge_date: str = ""             # "01/01/26"

    # ── Interest-rate parameters (auto-fetched from BOI API) ────────────────
    bank_rate_p1: float = 0.0              # BoI rate at period start (auto-fetched)
    bank_rate_p2: Optional[float] = None   # BoI rate after mid-period change, or None
    rate_change_date: Optional[str] = None # "DD.MM" of rate change, or None
    debit_margin: float = 0.0              # מרווח ריבית חובה (user-supplied)
    credit_margin: float = 0.0             # מרווח ריבית זכות (user-supplied)

    # ── Tier rows parsed from the interest-charge PDF ───────────────────────
    tier_rows: List[TierRow] = field(default_factory=list)

    # ── Transaction rows from the full account statement ────────────────────
    transactions: List[BankTransaction] = field(default_factory=list)

    # ── Totals ──────────────────────────────────────────────────────────────
    total_tier1: float = 0.0          # Sum of מדרגה 1 interest
    total_excess: float = 0.0         # Sum of חריגה interest
    total_charged: float = 0.0        # Amount actually debited from account
