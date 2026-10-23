# /// script
# requires-python = ">=3.10"
# dependencies = ["reportlab>=4", "openpyxl>=3.1"]
# ///
"""Make three fake art-supply invoices to test invoice_to_po.py with.

    uv run samples/make_test_invoices.py

Writes into samples/test-invoices/:
  <name>.pdf             the invoice to feed the tool
  <name>.expected.json   what a correct read of that PDF looks like (run the tool on it to
                         check the sheet without an API key, or diff against --save-json)
  templates/<code>.xlsx  a mock Retail Express template for that supplier + The Rocks,
                         with a master sheet that leaves one invoiced product out
  suppliers.json         invoice name → supplier code for the three of them

Every supplier, product, ABN and barcode here is made up. Each invoice is built to
exercise something different:
  1. Harbour Pigments  — ex-GST prices, cartons of 6/12, a 10% line discount, freight,
                         one product missing from Retail Express
  2. Gumleaf Easels    — GST-INCLUSIVE prices only, SKUs with leading zeros, a free bonus item
  3. Paperbark Mills   — two pages, a backordered line (0 supplied), SKUs formatted
                         differently from Retail Express (dots vs dashes), a minimum-order fee
"""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

HERE = Path(__file__).parent
OUT = HERE / "test-invoices"
GST = Decimal("0.10")
C2 = Decimal("0.01")


@dataclass
class Item:
    sku: str
    barcode: str
    desc: str
    qty: int              # as invoiced (cartons, packs or units)
    uom: str              # "EA", "CTN6", "PK10" ...
    price: Decimal        # per invoiced uom, as printed (ex or inc GST per the invoice)
    disc: int = 0         # line discount %
    supplied: int | None = None   # when backordered: units actually supplied
    re_sku: str | None = None     # how Retail Express spells the SKU, if different
    in_re: bool = True            # is this product in the Retail Express master?
    note: str = ""

    @property
    def pack(self) -> int:
        digits = "".join(ch for ch in self.uom if ch.isdigit())
        return int(digits) if digits else 1


@dataclass
class Invoice:
    file: str
    code: str
    name: str
    address: list[str]
    abn: str
    number: str
    date: str
    tax_inclusive: bool
    items: list[Item]
    charges: list[tuple[str, Decimal]] = field(default_factory=list)  # freight etc, same tax basis
    first_product_id: int = 20000


INVOICES = [
    Invoice(
        file="01-harbour-pigments",
        code="HP",
        name="Harbour Pigments Pty Ltd",
        address=["Unit 4, 18 Wharf Lane", "Pyrmont NSW 2009"],
        abn="11 222 333 444",
        number="HP-INV-58213",
        date="2026-10-14",
        tax_inclusive=False,
        first_product_id=30100,
        items=[
            Item("HP-AC75-TW", "9399001000017", "Studio Acrylic 75ml Titanium White", 2, "CTN6", Decimal("33.00")),
            Item("HP-AC75-UB", "9399001000024", "Studio Acrylic 75ml Ultramarine Blue", 1, "CTN6", Decimal("36.60")),
            Item("HP-AC75-CY", "9399001000031", "Studio Acrylic 75ml Cadmium Yellow Hue", 1, "CTN6", Decimal("36.60")),
            Item("HP-OIL37-BS", "9399001000109", "Artist Oil 37ml Burnt Sienna", 6, "EA", Decimal("9.85")),
            Item("HP-OIL37-PB", "9399001000116", "Artist Oil 37ml Phthalo Blue", 6, "EA", Decimal("11.40")),
            Item("HP-GES500", "9399001000208", "Gesso Primer 500ml", 4, "EA", Decimal("14.20"), disc=10),
            Item("HP-MED-GL250", "9399001000215", "Gloss Medium 250ml", 3, "EA", Decimal("12.75")),
            Item("HP-BR-SET12", "9399001000307", "Synthetic Brush Set 12pc", 1, "CTN12", Decimal("102.00"),
                 in_re=False),
        ],
        charges=[("Freight - Metro", Decimal("18.50"))],
    ),
    Invoice(
        file="02-gumleaf-easels",
        code="GE",
        name="Gumleaf Easels & Canvas Co.",
        address=["PO Box 712", "Marrickville NSW 2204"],
        abn="55 666 777 888",
        number="102947",
        date="2026-10-16",
        tax_inclusive=True,
        first_product_id=41200,
        items=[
            Item("000412", "9399002000014", "Stretched Canvas 30x40cm Cotton", 10, "EA", Decimal("12.10")),
            Item("000413", "9399002000021", "Stretched Canvas 40x50cm Cotton", 10, "EA", Decimal("16.50")),
            Item("000415", "9399002000038", "Stretched Canvas 60x90cm Deep Edge", 4, "EA", Decimal("44.00")),
            Item("000088", "9399002000106", "Canvas Panel 20x25cm", 2, "PK10", Decimal("27.50")),
            Item("001730", "9399002000205", "Studio H-Frame Easel Beech", 1, "EA", Decimal("231.00")),
            Item("001731", "9399002000212", "Table Easel Pine", 3, "EA", Decimal("38.50")),
            Item("000999", "9399002000304", "Canvas Pliers (promo bonus)", 1, "EA", Decimal("0.00"),
                 note="free bonus item"),
        ],
        charges=[("Delivery", Decimal("27.50"))],
    ),
    Invoice(
        file="03-paperbark-mills",
        code="PB",
        name="Paperbark Paper Mills",
        address=["212 Fibre Road", "Wetherill Park NSW 2164"],
        abn="99 888 777 666",
        number="PBM/26/04411",
        date="2026-10-20",
        tax_inclusive=False,
        first_product_id=52000,
        items=[
            *[Item(f"PB.SK.A{s}.{g}", f"93990030{i:05d}", f"Sketch Pad A{s} {g}gsm 50 sheets", q, "EA", p,
                   re_sku=f"PB-SK-A{s}-{g}")
              for i, (s, g, q, p) in enumerate([(5, 110, 12, Decimal("4.20")), (4, 110, 12, Decimal("6.10")),
                                                 (3, 110, 6, Decimal("10.90")), (5, 150, 12, Decimal("5.30")),
                                                 (4, 150, 12, Decimal("7.80")), (3, 150, 6, Decimal("13.40"))])],
            *[Item(f"PB.WC.{n}", f"93990031{i:05d}", f"Watercolour Pad {d} 300gsm 12 sheets", 6, "EA", p,
                   re_sku=f"PB-WC-{n}")
              for i, (n, d, p) in enumerate([("A5C", "A5 Cold Press", Decimal("8.90")),
                                             ("A4C", "A4 Cold Press", Decimal("14.50")),
                                             ("A3C", "A3 Cold Press", Decimal("26.80")),
                                             ("A4H", "A4 Hot Press", Decimal("15.20")),
                                             ("A3H", "A3 Hot Press", Decimal("27.90"))])],
            *[Item(f"PB.CS.{c}", f"93990032{i:05d}", f"Coloured Card A4 220gsm {c.title()} 25pk", 4, "EA",
                   Decimal("6.75"), re_sku=f"PB-CS-{c}")
              for i, c in enumerate(["BLACK", "WHITE", "CREAM", "RED", "NAVY", "FOREST", "SKY", "LEMON"])],
            Item("PB.MX.A4", "9399003300011", "Mixed Media Pad A4 250gsm", 6, "EA", Decimal("12.60"),
                 re_sku="PB-MX-A4"),
            Item("PB.MX.A3", "9399003300028", "Mixed Media Pad A3 250gsm", 6, "EA", Decimal("21.40"),
                 re_sku="PB-MX-A3", supplied=0, note="backordered, none supplied"),
            Item("PB.TR.A4", "9399003300103", "Tracing Pad A4 90gsm", 10, "EA", Decimal("5.95"),
                 re_sku="PB-TR-A4"),
            Item("PB.KR.ROLL", "9399003300202", "Kraft Paper Roll 1m x 10m", 2, "EA", Decimal("18.00"),
                 re_sku="PB-KR-ROLL"),
        ],
        charges=[("Small order surcharge", Decimal("0.00")), ("Freight", Decimal("32.00"))],
    ),
]


def money(d: Decimal) -> str:
    return f"${d.quantize(C2, ROUND_HALF_UP):,.2f}"


def line_total(it: Item) -> Decimal:
    qty = it.qty if it.supplied is None else it.supplied
    return (it.price * qty * (100 - it.disc) / 100).quantize(C2, ROUND_HALF_UP)


def render_pdf(inv: Invoice, path: Path) -> None:
    styles = getSampleStyleSheet()
    small = styles["Normal"].clone("small", fontSize=8, leading=10)
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=15 * mm, bottomMargin=18 * mm, title=f"Tax Invoice {inv.number}")
    story = []
    head = Table([[
        Paragraph(f"<b><font size=15>{inv.name}</font></b><br/>" + "<br/>".join(inv.address)
                  + f"<br/>ABN {inv.abn}", styles["Normal"]),
        Paragraph(f"<b><font size=15>TAX INVOICE</font></b><br/>Invoice No: <b>{inv.number}</b><br/>"
                  f"Date: {inv.date}<br/>Terms: 30 days EOM", styles["Normal"]),
    ]], colWidths=[105 * mm, 75 * mm])
    story += [head, Spacer(1, 6 * mm)]
    story.append(Paragraph("<b>Bill to / Ship to:</b> Sample Art Supplies — The Rocks store, "
                           "1 Test Street, The Rocks NSW 2000 · Cust. A/C TR-0042 · Your PO: 7781", small))
    story.append(Spacer(1, 5 * mm))

    price_hdr = "Price inc GST" if inv.tax_inclusive else "Unit Price"
    rows = [["Item Code", "Barcode", "Description", "UOM", "Ord", "Sup", price_hdr, "Disc", "Amount"]]
    for it in inv.items:
        sup = it.qty if it.supplied is None else it.supplied
        rows.append([it.sku, it.barcode, Paragraph(it.desc, small), it.uom, str(it.qty), str(sup),
                     money(it.price), f"{it.disc}%" if it.disc else "", money(line_total(it))])
    for label, amt in inv.charges:
        rows.append(["", "", Paragraph(label, small), "", "", "", "", "", money(amt)])
    t = Table(rows, repeatRows=1,
              colWidths=[26 * mm, 26 * mm, 52 * mm, 12 * mm, 9 * mm, 9 * mm, 18 * mm, 10 * mm, 18 * mm])
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 7.5),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 7.5),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2f4858")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (4, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#cccccc")),
    ]))
    story += [t, Spacer(1, 5 * mm)]

    goods = sum((line_total(i) for i in inv.items), Decimal(0))
    charges = sum((a for _, a in inv.charges), Decimal(0))
    if inv.tax_inclusive:
        total = goods + charges
        gst = (total / 11).quantize(C2, ROUND_HALF_UP)
        totals = [["Total (inc GST)", money(total)], ["Includes GST of", money(gst)]]
        story.append(Paragraph("<i>All prices shown include GST.</i>", small))
    else:
        sub = goods + charges
        gst = (sub * GST).quantize(C2, ROUND_HALF_UP)
        totals = [["Goods (ex GST)", money(goods)], ["Freight & charges (ex GST)", money(charges)],
                  ["Subtotal (ex GST)", money(sub)], ["GST 10%", money(gst)], ["TOTAL (inc GST)", money(sub + gst)]]
    tt = Table(totals, colWidths=[50 * mm, 25 * mm], hAlign="RIGHT")
    tt.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "Helvetica", 9), ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                            ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 9),
                            ("LINEABOVE", (0, -1), (-1, -1), 0.5, colors.black)]))
    story += [tt, Spacer(1, 8 * mm),
              Paragraph("Pay to BSB 000-000 Acc 00000000. TEST INVOICE — made-up data for software testing, "
                        "not a real document.", small)]

    def footer(canvas, doc_):
        canvas.setFont("Helvetica", 7)
        canvas.drawRightString(A4[0] - 15 * mm, 10 * mm, f"{inv.name} · Invoice {inv.number} · Page {doc_.page}")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)


def expected(inv: Invoice) -> dict:
    """What a correct extraction looks like, in invoice_to_po.py's schema."""
    lines = []
    for it in inv.items:
        sup = it.qty if it.supplied is None else it.supplied
        unit = it.price * (100 - it.disc) / 100 / it.pack
        if inv.tax_inclusive:
            unit = unit / (1 + GST)
        notes = [n for n in (f"{it.qty} x {it.uom} of {it.pack} units" if it.pack > 1 else "",
                             f"{it.disc}% line discount" if it.disc else "", it.note) if n]
        lines.append({"supplier_sku": it.sku, "barcode": it.barcode, "description": it.desc[:50],
                      "qty": sup * it.pack, "unit_cost_ex_tax": float(unit.quantize(Decimal("0.0001"), ROUND_HALF_UP)),
                      "note": "; ".join(notes)})
    goods = sum((line_total(i) for i in inv.items), Decimal(0))
    subtotal = None if inv.tax_inclusive else float(goods)
    return {"supplier_name": inv.name, "invoice_number": inv.number, "invoice_date": inv.date,
            "subtotal_ex_tax": subtotal, "lines": lines}


def make_template(inv: Invoice, path: Path, tool) -> None:
    wb = tool.build_template()
    ws = wb.active
    for label, value in (("Outlet", "The Rocks"), ("Supplier", inv.code)):
        cell = tool.find_cell(ws, label)
        ws.cell(cell.row, cell.column + 1, value)
    m = wb.create_sheet("Master Sheet")
    m.append(["ProductId", "ManufacturerSKU", "SupplierSKU", "ShortDescription", "SupplierSKU2"])
    pid = inv.first_product_id
    for it in inv.items:
        pid += 1
        if it.in_re:
            m.append([pid, it.barcode, it.re_sku or it.sku, it.desc[:50], ""])
    m.append([pid + 1, "", f"{inv.code}-DISCONTINUED", "Old line no longer invoiced", ""])
    wb.save(path)


def main() -> int:
    import importlib.util
    spec = importlib.util.spec_from_file_location("invoice_to_po", HERE.parent / "invoice_to_po.py")
    tool = importlib.util.module_from_spec(spec)
    sys.modules["invoice_to_po"] = tool
    spec.loader.exec_module(tool)

    (OUT / "templates").mkdir(parents=True, exist_ok=True)
    suppliers = {}
    for inv in INVOICES:
        render_pdf(inv, OUT / f"{inv.file}.pdf")
        (OUT / f"{inv.file}.expected.json").write_text(json.dumps(expected(inv), indent=2))
        make_template(inv, OUT / "templates" / f"{inv.code}.xlsx", tool)
        suppliers[inv.name.split(" ")[0] + " " + inv.name.split(" ")[1]] = inv.code
        print(f"{inv.file}.pdf  ({inv.code}, {len(inv.items)} lines)")
    (OUT / "suppliers.json").write_text(json.dumps(suppliers, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
