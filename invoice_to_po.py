# /// script
# requires-python = ">=3.10"
# dependencies = ["anthropic>=0.70", "openpyxl>=3.1"]
# ///
"""Turn supplier PDF invoices into a filled-in Retail Express
"Purchase Orders - Purchase Order Details" import sheet.

    uv run invoice_to_po.py INVOICE.pdf [MORE.pdf ...] --outlet "The Rocks" --supplier FG

Each invoice becomes its own .xlsx (one PO per supplier invoice), written next to
the PDF unless --out-dir is given. Pass --template with your own copy of the
Retail Express template to fill that exact file; otherwise a replica is generated.
"""

from __future__ import annotations

import argparse
import base64
import json
import sys
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

MODEL = "claude-opus-5-5"

# Template columns, in order, as Retail Express names them.
COLUMNS = [
    "ProductId",
    "ManufacturerSKU",
    "SupplierSKU",
    "ShortDescription",
    "SupplierSKU2",
    "Qty",
    "SupplierBuy",
]

EXTRACTION_PROMPT = """\
You are reading a supplier invoice so its lines can be loaded into a Retail Express
purchase order. Return every product line on the invoice.

For each line:
- supplier_sku: the supplier's own product code / item number / SKU exactly as printed
  (keep leading zeros, dashes and letters). Empty string if the line has none.
- barcode: the EAN/UPC/APN barcode if printed, else empty string.
- description: the product description as printed (trim to 50 characters).
- qty: the number of SELLABLE UNITS supplied, as a whole number. If the invoice bills in
  cartons/packs and states the units per carton, multiply it out (e.g. 2 ctn x 12 = 24)
  and say so in note.
- unit_cost_ex_tax: the cost of ONE sellable unit EXCLUDING GST/VAT, after any line
  discount, as a decimal number. If the invoice only shows tax-inclusive prices, remove
  the tax at the rate the invoice states. If qty was multiplied out from cartons, divide
  the carton price the same way.
- note: anything a buyer should double-check (unclear quantity, free goods, backorder,
  pack conversion). Empty string if nothing.

Skip lines that are not stock: freight, delivery, surcharges, deposits, rounding,
subtotals, payment terms. Do include free/bonus stock lines at cost 0 with a note.

Also return the header fields: supplier_name, invoice_number, invoice_date (YYYY-MM-DD
or as printed), and subtotal_ex_tax for the stock lines if the invoice shows one
(null if not shown).
"""

EXTRACTION_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["supplier_name", "invoice_number", "invoice_date", "subtotal_ex_tax", "lines"],
    "properties": {
        "supplier_name": {"type": "string"},
        "invoice_number": {"type": "string"},
        "invoice_date": {"type": "string"},
        "subtotal_ex_tax": {"type": ["number", "null"]},
        "lines": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["supplier_sku", "barcode", "description", "qty", "unit_cost_ex_tax", "note"],
                "properties": {
                    "supplier_sku": {"type": "string"},
                    "barcode": {"type": "string"},
                    "description": {"type": "string"},
                    "qty": {"type": "integer"},
                    "unit_cost_ex_tax": {"type": "number"},
                    "note": {"type": "string"},
                },
            },
        },
    },
}


@dataclass
class Line:
    supplier_sku: str
    barcode: str
    description: str
    qty: int
    unit_cost: Decimal
    note: str


@dataclass
class Invoice:
    supplier_name: str
    invoice_number: str
    invoice_date: str
    subtotal_ex_tax: Decimal | None
    lines: list[Line]


# --------------------------------------------------------------------------- extraction


def extract_invoice(pdf: Path) -> dict:
    """Ask Claude to read the PDF and return the schema above as a dict."""
    import anthropic

    client = anthropic.Anthropic()
    pdf_b64 = base64.standard_b64encode(pdf.read_bytes()).decode()
    with client.beta.messages.stream(
        model=MODEL,
        max_tokens=64000,
        betas=["server-side-fallback-2026-07-01"],
        extra_body={"fallbacks": "default"},
        output_config={"effort": "medium", "format": {"type": "json_schema", "schema": EXTRACTION_SCHEMA}},
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}},
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }
        ],
    ) as stream:
        response = stream.get_final_message()

    if response.stop_reason == "refusal":
        raise RuntimeError(f"{pdf.name}: the model declined to read this file")
    if response.stop_reason == "max_tokens":
        raise RuntimeError(f"{pdf.name}: invoice too long for one pass — split the PDF")
    text = "".join(b.text for b in response.content if b.type == "text")
    return json.loads(text)


def to_invoice(data: dict) -> Invoice:
    def money(v) -> Decimal:
        return Decimal(str(v)).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    subtotal = data.get("subtotal_ex_tax")
    return Invoice(
        supplier_name=data.get("supplier_name", ""),
        invoice_number=data.get("invoice_number", ""),
        invoice_date=data.get("invoice_date", ""),
        subtotal_ex_tax=None if subtotal is None else money(subtotal),
        lines=[
            Line(
                supplier_sku=str(l.get("supplier_sku", "")).strip(),
                barcode=str(l.get("barcode", "")).strip(),
                description=str(l.get("description", "")).strip()[:50],
                qty=int(l["qty"]),
                unit_cost=money(l["unit_cost_ex_tax"]),
                note=str(l.get("note", "")).strip(),
            )
            for l in data["lines"]
        ],
    )


# --------------------------------------------------------------------------- sheet


def find_cell(ws: Worksheet, label: str):
    want = label.replace(" ", "").lower()
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.replace(" ", "").lower() == want:
                return cell
    return None


def build_template() -> Workbook:
    """A stand-in for the Retail Express template, laid out like the original."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Purchase Order Details"
    navy = PatternFill("solid", fgColor="000080")
    red = PatternFill("solid", fgColor="D9341F")
    dark = PatternFill("solid", fgColor="333333")
    white_bold = Font(bold=True, color="FFFFFF")

    ws["A1"] = "Purchase Orders - Purchase Order Details"
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    for col in "ABCD":
        ws[f"{col}1"].fill = navy
    ws["A3"], ws["A4"] = "Outlet", "Supplier"
    for r in (3, 4):
        ws[f"A{r}"].font = Font(bold=True)
        ws[f"B{r}"].fill = red
        ws[f"B{r}"].font = Font(color="FFFFFF")

    types = [
        ("Integer", "Whole Number"),
        ("Reference Only", "See Master Sheet"),
        ("Text (lookup)", "50 chars"),
        ("Reference Only", "See Master Sheet"),
        ("Text (lookup)", "50 chars"),
        ("Integer", "Whole Number"),
        ("Money", "$123.45"),
    ]
    required = {"ProductId", "SupplierSKU", "SupplierSKU2", "Qty"}
    for i, (name, (t1, t2)) in enumerate(zip(COLUMNS, types), start=1):
        ws.cell(6, i, t1)
        ws.cell(7, i, t2)
        if t1 == "Reference Only":
            for r in (6, 7):
                ws.cell(r, i).fill = dark
                ws.cell(r, i).font = Font(color="FFFFFF")
        h = ws.cell(9, i, name)
        h.font = white_bold if name in required else Font(bold=True)
        if name in required:
            h.fill = red
        ws.column_dimensions[h.column_letter].width = 22
    ws["F1"] = (
        "Enter the Product ID and/or SupplierSKU and/or Supplier SKU 2 along with a desired "
        "Quantity for each product. To the left, please ensure you have entered at minimum an "
        "Outlet and a Supplier Code into the boxes provided to create your purchase order."
    )
    ws["F1"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells("F1:J4")
    return wb


def fill(wb: Workbook, outlet: str, supplier_code: str, inv: Invoice, sku2: str) -> Worksheet:
    ws = None
    for sheet in wb.worksheets:
        if find_cell(sheet, "SupplierSKU"):
            ws = sheet
            break
    if ws is None:
        raise RuntimeError("template has no 'SupplierSKU' header cell")

    for label, value in (("Outlet", outlet), ("Supplier", supplier_code)):
        cell = find_cell(ws, label)
        if cell is None:
            raise RuntimeError(f"template has no '{label}' label cell")
        ws.cell(cell.row, cell.column + 1, value)

    header = find_cell(ws, "SupplierSKU")
    cols = {}
    for cell in ws[header.row]:
        if isinstance(cell.value, str):
            cols[cell.value.replace(" ", "")] = cell.column

    first = header.row + 1
    # Clear anything left in the template's data area.
    for r in range(first, ws.max_row + 1):
        for c in cols.values():
            ws.cell(r, c).value = None

    for i, line in enumerate(inv.lines):
        r = first + i
        ws.cell(r, cols["SupplierSKU"], line.supplier_sku or None)
        if sku2 == "barcode" and "SupplierSKU2" in cols:
            ws.cell(r, cols["SupplierSKU2"], line.barcode or None)
        if "ShortDescription" in cols:
            ws.cell(r, cols["ShortDescription"], line.description)
        ws.cell(r, cols["Qty"], line.qty)
        if "SupplierBuy" in cols:
            c = ws.cell(r, cols["SupplierBuy"], float(line.unit_cost))
            c.number_format = '"$"#,##0.00##'
    return ws


def add_check_sheet(wb: Workbook, inv: Invoice, pdf: Path) -> Decimal:
    ws = wb.create_sheet("Invoice Check")
    ws.append(["Source", pdf.name])
    ws.append(["Supplier on invoice", inv.supplier_name])
    ws.append(["Invoice #", inv.invoice_number])
    ws.append(["Invoice date", inv.invoice_date])
    total = sum((l.unit_cost * l.qty for l in inv.lines), Decimal(0)).quantize(Decimal("0.01"))
    ws.append(["Lines total ex tax", float(total)])
    ws.append(["Invoice subtotal ex tax", None if inv.subtotal_ex_tax is None else float(inv.subtotal_ex_tax)])
    ws.append([])
    ws.append(["SupplierSKU", "Barcode", "Description", "Qty", "Unit ex tax", "Line ex tax", "Note"])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    for l in inv.lines:
        ws.append([l.supplier_sku, l.barcode, l.description, l.qty, float(l.unit_cost),
                   float((l.unit_cost * l.qty).quantize(Decimal("0.01"))), l.note])
    for col, w in zip("ABCDEFG", (18, 16, 40, 8, 12, 12, 50)):
        ws.column_dimensions[col].width = w
    return total


# --------------------------------------------------------------------------- cli


def resolve_supplier(arg: str | None, inv: Invoice, suppliers_file: Path) -> str:
    if arg:
        return arg
    if suppliers_file.exists():
        mapping = json.loads(suppliers_file.read_text())
        name = inv.supplier_name.lower()
        for key, code in mapping.items():
            if key.lower() in name:
                return code
    print(f"  ! no supplier code for '{inv.supplier_name}' — pass --supplier or add it to "
          f"{suppliers_file.name}; leaving the Supplier box blank", file=sys.stderr)
    return ""


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("pdfs", nargs="+", type=Path, help="invoice PDFs (or .json extractions from --save-json)")
    p.add_argument("--outlet", required=True, help='Retail Express outlet name, e.g. "The Rocks"')
    p.add_argument("--supplier", help="Retail Express supplier code, e.g. FG (else looked up in suppliers.json)")
    p.add_argument("--template", type=Path, help="your Retail Express template .xlsx to fill (else a replica)")
    p.add_argument("--out-dir", type=Path, help="where to write the .xlsx files (default: next to each PDF)")
    p.add_argument("--sku2", choices=["barcode", "none"], default="barcode",
                   help="what to put in SupplierSKU2 (default: the invoice barcode)")
    p.add_argument("--save-json", action="store_true", help="also save the raw extraction as .json for review/re-use")
    p.add_argument("--suppliers", type=Path, default=Path(__file__).with_name("suppliers.json"),
                   help="JSON map of supplier-name fragment → supplier code")
    args = p.parse_args()

    failures = 0
    for src in args.pdfs:
        print(f"{src.name}")
        try:
            if src.suffix.lower() == ".json":
                data = json.loads(src.read_text())
            else:
                data = extract_invoice(src)
            out_dir = args.out_dir or src.parent
            out_dir.mkdir(parents=True, exist_ok=True)
            if args.save_json and src.suffix.lower() != ".json":
                (out_dir / f"{src.stem}.json").write_text(json.dumps(data, indent=2))

            inv = to_invoice(data)
            supplier = resolve_supplier(args.supplier, inv, args.suppliers)
            wb = load_workbook(args.template) if args.template else build_template()
            fill(wb, args.outlet, supplier, inv, args.sku2)
            total = add_check_sheet(wb, inv, src)
            out = out_dir / f"PO {supplier or 'UNKNOWN'} {inv.invoice_number or src.stem}.xlsx".replace("/", "-")
            wb.save(out)

            print(f"  {inv.supplier_name} · invoice {inv.invoice_number} · {len(inv.lines)} lines · "
                  f"${total} ex tax → {out}")
            if inv.subtotal_ex_tax is not None and abs(inv.subtotal_ex_tax - total) > Decimal("0.05"):
                print(f"  ! lines add to ${total} but the invoice subtotal is ${inv.subtotal_ex_tax:.2f} — check it",
                      file=sys.stderr)
            for l in inv.lines:
                if not l.supplier_sku:
                    print(f"  ! no supplier SKU: {l.description}", file=sys.stderr)
                if l.note:
                    print(f"  · {l.supplier_sku or l.description}: {l.note}")
        except Exception as e:  # keep going with the other invoices
            failures += 1
            print(f"  ✗ {e}", file=sys.stderr)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
