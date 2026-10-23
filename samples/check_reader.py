"""Check invoice_reader.py against the test invoices: every PDF in samples/test-invoices/ must read
exactly as its .expected.json says.

    uv run --with pdfplumber python samples/check_reader.py
"""
import json, sys
from pathlib import Path
HERE = Path(__file__).parent
sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent.parent))
from invoice_reader import read_invoice
ok = True
for name in ["01-harbour-pigments", "02-gumleaf-easels", "03-paperbark-mills", "04-kookaburra-brushes"]:
    got = read_invoice(open(f"{HERE}/test-invoices/{name}.pdf", "rb").read(), json.load(open(f"{HERE}/test-invoices/suppliers.json")))
    exp = json.load(open(f"{HERE}/test-invoices/{name}.expected.json"))
    print(f"== {name}: {len(got['lines'])}/{len(exp['lines'])} lines; supplier={got['supplier_name']!r} inv={got['invoice_number']!r} date={got['invoice_date']!r} subtotal={got['subtotal_ex_tax']} (exp {exp['subtotal_ex_tax']})")
    for k in ("supplier_name", "invoice_number", "invoice_date", "subtotal_ex_tax"):
        if got[k] != exp[k]: print("   header mismatch", k, got[k], "!=", exp[k]); ok = False
    for g, e in zip(got["lines"], exp["lines"]):
        for k in ("supplier_sku", "barcode", "description", "qty", "unit_cost_ex_tax"):
            if (abs(g[k]-e[k]) > 0.0001 if k == "unit_cost_ex_tax" else g[k] != e[k]):
                print(f"   {e['supplier_sku']}: {k} got {g[k]!r} expected {e[k]!r}"); ok = False
    if len(got["lines"]) != len(exp["lines"]): ok = False
print("ALL MATCH" if ok else "DIFFERENCES"); sys.exit(0 if ok else 1)
