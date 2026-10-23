"""Read supplier invoice PDFs without AI.

Works on PDFs that contain real text (what accounting systems email out), not scans or photos.
It finds the line-item table by its header row ("Item Code", "Description", "Qty", "Unit Price",
"Amount", ...), works out each column's position from that header, and reads the rows under it.
Header words it knows are in HEADER_WORDS. When a supplier names a column differently, add the
word there.

    uv run --with pdfplumber python invoice_reader.py invoice.pdf     # prints what it read
"""

from __future__ import annotations

import csv
import io
import json
import re
import sys
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

import pdfplumber

# Column roles and the header text that names them, after lower-casing and removing
# anything that isn't a letter or digit ("Unit Price (ex)" → "unitpriceex").
HEADER_WORDS: dict[str, list[str]] = {
    "code": ["itemcode", "code", "item", "itemno", "sku", "productcode", "product", "prodcode", "partno",
             "part", "catno", "catalogue", "stockcode", "article", "articleno", "ref", "itemnumber",
             "productno", "productid", "supplierCode".lower(), "yourcode"],
    "barcode": ["barcode", "ean", "apn", "upc", "gtin", "ean13"],
    "desc": ["description", "desc", "itemdescription", "productdescription", "details", "product name".replace(" ", ""),
             "name", "particulars"],
    "uom": ["uom", "unit", "units", "pack", "packsize", "um", "per"],
    "qty_supplied": ["sup", "supplied", "shipped", "ship", "qtysupplied", "qtyshipped", "despatched", "delivered",
                     "qtydel", "invoiced", "qtyinvoiced", "inv"],
    "qty": ["qty", "quantity", "ord", "ordered", "qtyordered", "qtyord"],
    "price": ["unitprice", "price", "unitcost", "cost", "rate", "each", "unitex", "priceex", "unitpriceex",
              "unitpriceexgst", "priceexgst", "priceincgst", "unitpriceincgst", "priceinc", "unitpriceinc",
              "nett", "netprice", "netunit", "sell", "buyprice", "wholesale"],
    "disc": ["disc", "discount", "disc%", "discpercent", "less"],
    "amount": ["amount", "total", "linetotal", "extended", "ext", "value", "nettamount", "amountex",
               "amountexgst", "totalex", "totalexgst", "lineamount", "extprice", "amountincgst", "totalinc"],
    "gst": ["gst", "tax", "gstamount", "taxamount"],
}
ROLE_OF = {w: role for role, words in HEADER_WORDS.items() for w in words}

CHARGE_WORDS = re.compile(r"freight|delivery|postage|shipping|courier|surcharge|handling|fuel levy|"
                          r"min(imum)?\.? order|admin fee|small order|rounding|deposit|carriage", re.I)
END_OF_TABLE = re.compile(r"^(sub\s*-?total|total|goods|balance|amount due|gst\b|freight\s*&)", re.I)
TAX_INCLUSIVE = re.compile(r"(prices?|amounts?)\s+(shown\s+|are\s+)?(include|incl?\.?)\s*(of\s+)?gst|"
                           r"gst\s+inclusive|inc(l|lusive)?\.?\s+gst\s+prices", re.I)
INVOICE_NO = re.compile(r"(?:tax\s+)?invoice\s*(?:no\.?|number|num|#)\s*[:#.]?\s*([A-Z0-9][A-Z0-9/\-_.]*)", re.I)
INVOICE_DATE = re.compile(r"(?:invoice\s+)?date\s*[:.]?\s*(\d{4}-\d{2}-\d{2}|\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}|"
                          r"\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{4})", re.I)
GOODS_EX = re.compile(r"(?:goods|merchandise|products?)\s*(?:total)?\s*\(?\s*(?:ex|excl?)\.?\s*gst\s*\)?\s*:?\s*"
                      r"\$?\s*([\d,]+\.\d{2})", re.I)
SUBTOTAL_EX = re.compile(r"sub\s*-?\s*total\s*\(?\s*(?:ex|excl?)?\.?\s*(?:gst)?\s*\)?\s*:?\s*\$?\s*([\d,]+\.\d{2})", re.I)
PACK = re.compile(r"(?:ctn|carton|box|bx|pk|pack|pkt|case|cs|inner|outer|bdl|bundle|doz)\s*(?:of|x)?\s*(\d+)|"
                  r"(\d+)\s*(?:per|/)\s*(?:ctn|carton|box|pk|pack|case)|^(?:x\s*)?(\d+)\s*(?:pk|pack|pc)?$", re.I)
GAP = 5.0   # points between words that still belong to the same phrase


class ReadError(Exception):
    pass


@dataclass
class Phrase:
    text: str
    x0: float
    x1: float
    top: float
    words: list[tuple[str, float, float]] = field(default_factory=list)

    @property
    def mid(self) -> float:
        return (self.x0 + self.x1) / 2


@dataclass
class Column:
    role: str
    label: str
    x0: float
    x1: float


@dataclass
class Row:
    cells: dict[str, str] = field(default_factory=dict)
    before: list[str] = field(default_factory=list)   # wrapped description lines above the numbers
    after: list[str] = field(default_factory=list)    # … and below
    top: float = 0.0


def place_wrapped(pending: list[tuple[float, str]], prev: Row | None, nxt: Row | None, nxt_top: float) -> None:
    """Give each wrapped description line to whichever item row it sits closer to."""
    for top, text in pending:
        if prev is not None and (nxt is None or top - prev.top <= nxt_top - top):
            prev.after.append(text)
        elif nxt is not None:
            nxt.before.append(text)


def key(text: str) -> str:
    return re.sub(r"[^a-z0-9%]", "", text.lower())


def number(text: str | None) -> Decimal | None:
    if not text:
        return None
    t = text.replace("$", "").replace(",", "").replace(" ", "").replace("%", "")
    neg = t.startswith("(") and t.endswith(")") or t.endswith("-") or t.startswith("-")
    t = t.strip("()-")
    try:
        d = Decimal(t)
    except InvalidOperation:
        return None
    return -d if neg else d


def lines_of(page) -> list[list[Phrase]]:
    """Words on a page → lines (top to bottom) → phrases (words close together)."""
    words = page.extract_words(x_tolerance=1.5, y_tolerance=2, keep_blank_chars=False)
    words.sort(key=lambda w: (round(w["top"]), w["x0"]))
    rows: list[list[dict]] = []
    for w in words:
        if rows and abs(rows[-1][0]["top"] - w["top"]) <= 2.5:
            rows[-1].append(w)
        else:
            rows.append([w])
    out = []
    for ws in rows:
        ws.sort(key=lambda w: w["x0"])
        phrases: list[Phrase] = []
        for w in ws:
            if phrases and w["x0"] - phrases[-1].x1 <= GAP:
                p = phrases[-1]
                p.text += " " + w["text"]
                p.x1 = w["x1"]
                p.words.append((w["text"], w["x0"], w["x1"]))
            else:
                phrases.append(Phrase(w["text"], w["x0"], w["x1"], w["top"], [(w["text"], w["x0"], w["x1"])]))
        out.append(phrases)
    return out


def split_header(p: Phrase) -> list[Column]:
    """One header phrase → its columns. Usually one ("Unit Price (ex GST)"), but narrow columns
    can sit so close that two names read as one phrase ("Sup Price inc GST")."""
    words, cols, i = p.words, [], 0
    while i < len(words):
        for j in range(len(words), i, -1):   # longest run of words that names a column
            role = ROLE_OF.get(key("".join(w[0] for w in words[i:j])))
            if role:
                cols.append(Column(role, " ".join(w[0] for w in words[i:j]), words[i][1], words[j - 1][2]))
                i = j
                break
        else:
            if cols:   # trailing words like "(ex GST)" belong to the column before
                c = cols[-1]
                c.label += " " + words[i][0]
                c.x1 = words[i][2]
            i += 1
    return cols


def header_columns(line: list[Phrase]) -> list[Column] | None:
    cols = []
    for p in line:
        cols += split_header(p)
    roles = {c.role for c in cols}
    has_qty = roles & {"qty", "qty_supplied"}
    if has_qty and roles & {"price", "amount"} and roles & {"code", "desc"} and len(cols) >= 3:
        return cols
    return None


def assign(line: list[Phrase], cols: list[Column]) -> Row:
    """Put each phrase under the header column it overlaps most (or is nearest to)."""
    row = Row()
    for p in line:
        best, score = None, None
        for c in cols:
            overlap = min(p.x1, c.x1) - max(p.x0, c.x0)
            s = overlap if overlap > 0 else -min(abs(p.x0 - c.x1), abs(c.x0 - p.x1), abs(p.mid - (c.x0 + c.x1) / 2))
            if score is None or s > score:
                best, score = c, s
        if best.role in row.cells:
            # Two phrases under one header (e.g. a long description split by a gap).
            row.cells[best.role] += " " + p.text
        else:
            row.cells[best.role] = p.text
    return row


def pack_size(uom: str | None, desc: str) -> int:
    for text in (uom or "",):
        m = PACK.search(text.strip())
        if m:
            n = next(g for g in m.groups() if g)
            if int(n) > 1:
                return int(n)
    return 1


def to_lines(rows: list[Row], cols: list[Column], tax_inclusive: bool) -> tuple[list[dict], Decimal]:
    """Table rows → invoice lines (units, unit cost ex GST), plus the total of freight-type charges."""
    price_label = next((c.label for c in cols if c.role == "price"), "")
    if re.search(r"\binc", price_label, re.I):
        tax_inclusive = True
    if re.search(r"\bex", price_label, re.I):
        tax_inclusive = False

    lines, charges = [], Decimal(0)
    for r in rows:
        desc = " ".join(t for t in r.before + [r.cells.get("desc", "")] + r.after if t).strip()
        code = (r.cells.get("code") or "").strip()
        everything = " ".join(r.cells.values())
        amount = number(r.cells.get("amount"))
        # Freight & co. — whether or not the supplier gives it an item code like "FRT".
        if CHARGE_WORDS.search(desc if desc else everything):
            charges += amount or Decimal(0)
            continue
        qty_ord = number(r.cells.get("qty"))
        qty = number(r.cells.get("qty_supplied"))
        if qty is None:
            qty = qty_ord
        price = number(r.cells.get("price"))
        disc = number(r.cells.get("disc")) or Decimal(0)
        pack = pack_size(r.cells.get("uom"), desc)
        notes = []
        if price is None and amount is not None and qty:
            price = amount / qty
            disc = Decimal(0)
        if price is None:
            continue
        unit = price * (100 - disc) / 100 / pack
        if tax_inclusive:
            unit = unit / Decimal("1.1")
        if pack > 1:
            notes.append(f"{qty} x {r.cells.get('uom')} of {pack} units")
        if disc:
            notes.append(f"{disc.normalize():f}% line discount")
        if qty == 0 and qty_ord:
            notes.append(f"{qty_ord.normalize():f} ordered, none supplied (backorder)")
        elif qty_ord is not None and r.cells.get("qty_supplied") and qty != qty_ord:
            notes.append(f"{qty_ord.normalize():f} ordered, {qty.normalize():f} supplied")
        if price == 0:
            notes.append("no charge")
        if tax_inclusive:
            notes.append("GST removed from inc-GST price")
        barcode = (r.cells.get("barcode") or "").strip()
        lines.append({
            "supplier_sku": code,
            "barcode": barcode,
            "description": desc[:50],
            "qty": int(qty * pack),
            "unit_cost_ex_tax": float(unit.quantize(Decimal("0.0001"), ROUND_HALF_UP)),
            "note": "; ".join(n for n in notes if n),
        })
    if not lines:
        raise ReadError("Found the table header but no product lines under it.")
    return lines, charges


def read_invoice(pdf_bytes: bytes, suppliers: dict[str, str] | None = None) -> dict:
    """PDF bytes → the same dict shape invoice_to_po.to_invoice() takes."""
    try:
        pdf = pdfplumber.open(io.BytesIO(pdf_bytes))
    except Exception as e:
        raise ReadError(f"Couldn't open the PDF: {e}")
    with pdf:
        pages = [lines_of(p) for p in pdf.pages]
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    if not full_text.strip():
        raise ReadError("This PDF has no text in it — it's probably a scan or photo. This tool needs the PDF "
                        "the supplier emailed (or ask them for one).")

    cols: list[Column] | None = None
    rows: list[Row] = []
    tax_inclusive = bool(TAX_INCLUSIVE.search(full_text))
    pending: list[tuple[float, str]] = []
    for page in pages:
        in_table = False
        for line in page:
            found = header_columns(line)
            if found:
                cols, in_table = found, True
                continue
            if not cols:
                continue
            if not in_table and page is not pages[0]:
                # A later page without a repeated header: keep using the last one.
                in_table = True
            if not in_table:
                continue
            first = line[0].text if line else ""
            if END_OF_TABLE.match(first):
                place_wrapped(pending, rows[-1] if rows else None, None, 1e9)
                pending = []
                in_table = False
                continue
            row = assign(line, cols)
            top = line[0].top
            qty_text = row.cells.get("qty_supplied") or row.cells.get("qty")
            is_item = number(qty_text) is not None and (row.cells.get("price") or row.cells.get("amount"))
            if is_item or CHARGE_WORDS.search(" ".join(p.text for p in line)):
                # Charges are kept so their amount is known; they're filtered out below.
                place_wrapped(pending, rows[-1] if rows else None, row, top)
                pending = []
                row.top = top
                rows.append(row)
            elif set(row.cells) <= {"desc", "code"}:
                # A wrapped description line. With vertically-centred cells it can sit above
                # or below the line holding the numbers; decide once the next row turns up.
                pending.append((top, " ".join(p.text for p in line)))
        place_wrapped(pending, rows[-1] if rows else None, None, 1e9)
        pending = []

    if not cols:
        raise ReadError("Couldn't find the line-item table — no header row with a code/description, quantity and "
                        "price column. Add the supplier's column names to HEADER_WORDS in invoice_reader.py.")
    lines, charges = to_lines(rows, cols, tax_inclusive)

    goods = GOODS_EX.search(full_text)
    sub = SUBTOTAL_EX.search(full_text)
    subtotal = None
    if not tax_inclusive:
        if goods:
            subtotal = float(number(goods.group(1)))
        elif sub:
            subtotal = float(number(sub.group(1)) - charges)

    inv_no = INVOICE_NO.search(full_text)
    inv_date = INVOICE_DATE.search(full_text)
    return {
        "supplier_name": supplier_name(pages[0], full_text, suppliers or {}),
        "invoice_number": inv_no.group(1) if inv_no else "",
        "invoice_date": inv_date.group(1) if inv_date else "",
        "subtotal_ex_tax": subtotal,
        "lines": lines,
    }


def role_for(label: str) -> str | None:
    words = label.split()
    for n in range(len(words), 0, -1):
        role = ROLE_OF.get(key("".join(words[:n])))
        if role:
            return role
    return None


def read_csv(raw: bytes, suppliers: dict[str, str] | None = None, name: str = "") -> dict:
    """A CSV of invoice lines (e.g. exported by an email/PDF-to-sheets tool) → same dict as read_invoice.

    The header row is found by the same column names as for PDFs. Columns for the supplier,
    invoice number and date are used when present; otherwise the file name stands in."""
    text = raw.decode("utf-8-sig", errors="replace")
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    table = [r for r in csv.reader(io.StringIO(text), dialect) if any(c.strip() for c in r)]

    meta_words = {"supplier": ["supplier", "suppliername", "vendor", "from", "sender", "company"],
                  "invoice_number": ["invoiceno", "invoicenumber", "invoice", "invoice#", "invno", "docno"],
                  "invoice_date": ["invoicedate", "date", "docdate"]}
    for h, header in enumerate(table):
        roles = [role_for(c) for c in header]
        found = {r for r in roles if r}
        if (found & {"qty", "qty_supplied"}) and (found & {"price", "amount"}) and (found & {"code", "desc"}):
            break
    else:
        raise ReadError("Couldn't find a header row with a code/description, quantity and price column in the CSV.")

    cols = [Column(r, header[i], i, i) for i, r in enumerate(roles) if r]
    meta_idx = {m: next((i for i, c in enumerate(header) if key(c) in words and not roles[i]), None)
                for m, words in meta_words.items()}
    rows = []
    for rec in table[h + 1:]:
        row = Row()
        for i, r in enumerate(roles):
            if r and i < len(rec) and rec[i].strip():
                row.cells.setdefault(r, rec[i].strip())
        if row.cells:
            rows.append(row)
    tax_inclusive = bool(TAX_INCLUSIVE.search(text))
    lines, _ = to_lines(rows, cols, tax_inclusive)

    def meta(m: str) -> str:
        i = meta_idx[m]
        return next((rec[i].strip() for rec in table[h + 1:] if i is not None and i < len(rec) and rec[i].strip()), "")

    supplier = meta("supplier")
    for frag in (suppliers or {}):
        if not supplier and frag.lower() in text.lower():
            supplier = frag
    return {"supplier_name": supplier or Path(name).stem, "invoice_number": meta("invoice_number"),
            "invoice_date": meta("invoice_date"), "subtotal_ex_tax": None, "lines": lines}


def supplier_name(first_page: list[list[Phrase]], text: str, suppliers: dict[str, str]) -> str:
    """A known supplier from suppliers.json if its name appears; else the first phrase on page 1."""
    phrases = [p.text for line in first_page for p in line]
    for name in suppliers:
        if name.lower() in text.lower():
            return next((t for t in phrases if name.lower() in t.lower()), name)
    for t in phrases:
        if not re.search(r"tax invoice|^invoice|page \d|^abn|^date", t, re.I):
            return t
    return ""

if __name__ == "__main__":
    for path in sys.argv[1:]:
        with open(path, "rb") as f:
            raw = f.read()
        print(json.dumps(read_csv(raw, name=path) if path.lower().endswith(".csv") else read_invoice(raw), indent=2))
