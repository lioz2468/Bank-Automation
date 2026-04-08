
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from parsers.transactions import parse_transactions
txs = parse_transactions("transactions.xlsx", "01/10/25", "31/12/25")
RIBBIT = chr(1512)+chr(1489)+chr(1497)+chr(1514)
for t in txs:
    if t.value_date is not None:
        op = t.operation or ""
        print(repr(op[:30]))
        print("  contains:", RIBBIT in op)
        print("  debit:", t.debit, "bal:", t.balance)
