// Read supplier invoice PDFs (via pdf.js) and CSV exports into invoice lines — no AI.
//
// Works on PDFs that contain real text (what accounting systems email out), not scans or photos.
// It finds the line-item table by its header row ("Item Code", "Description", "Qty", "Unit Price",
// "Amount", ...), works out each column's position from that header, and reads the rows under it.
// Header words it knows are in HEADER_WORDS; when a supplier names a column differently, add it there.
//
// Output shape (also what po.js takes):
//   {supplier_name, invoice_number, invoice_date, subtotal_ex_tax, lines: [
//     {supplier_sku, barcode, description, qty, unit_cost_ex_tax, note}]}
(function () {
  "use strict";

  // Column roles and the header text that names them, after lower-casing and removing anything
  // that isn't a letter, digit or % ("Unit Price (ex)" → "unitpriceex").
  const HEADER_WORDS = {
    code: ["itemcode", "code", "item", "itemno", "sku", "productcode", "product", "prodcode", "partno", "part",
           "catno", "catalogue", "stockcode", "article", "articleno", "ref", "itemnumber", "productno",
           "productid", "suppliercode", "yourcode"],
    barcode: ["barcode", "ean", "apn", "upc", "gtin", "ean13"],
    desc: ["description", "desc", "itemdescription", "productdescription", "details", "productname", "name",
           "particulars"],
    uom: ["uom", "unit", "units", "pack", "packsize", "um", "per"],
    qty_supplied: ["sup", "supplied", "shipped", "ship", "qtysupplied", "qtyshipped", "despatched", "delivered",
                   "qtydel", "invoiced", "qtyinvoiced", "inv"],
    qty: ["qty", "quantity", "ord", "ordered", "qtyordered", "qtyord"],
    price: ["unitprice", "price", "unitcost", "cost", "rate", "each", "unitex", "priceex", "unitpriceex",
            "unitpriceexgst", "priceexgst", "priceincgst", "unitpriceincgst", "priceinc", "unitpriceinc", "nett",
            "netprice", "netunit", "sell", "buyprice", "wholesale"],
    disc: ["disc", "discount", "disc%", "discpercent", "less"],
    amount: ["amount", "total", "linetotal", "extended", "ext", "value", "nettamount", "amountex", "amountexgst",
             "totalex", "totalexgst", "lineamount", "extprice", "amountincgst", "totalinc"],
    gst: ["gst", "tax", "gstamount", "taxamount"],
  };
  const ROLE_OF = {};
  for (const [role, words] of Object.entries(HEADER_WORDS)) for (const w of words) ROLE_OF[w] = role;

  const CHARGE_WORDS = /freight|delivery|postage|shipping|courier|surcharge|handling|fuel levy|min(imum)?\.? order|admin fee|small order|rounding|deposit|carriage/i;
  const END_OF_TABLE = /^(sub\s*-?total|total|goods|balance|amount due|gst\b|freight\s*&)/i;
  const TAX_INCLUSIVE = /(prices?|amounts?)\s+(shown\s+|are\s+)?(include|incl?\.?)\s*(of\s+)?gst|gst\s+inclusive|inc(l|lusive)?\.?\s+gst\s+prices/i;
  const INVOICE_NO = /(?:tax\s+)?invoice\s*(?:no\.?|number|num|#)\s*[:#.]?\s*([A-Z0-9][A-Z0-9/\-_.]*)/i;
  const INVOICE_DATE = /(?:invoice\s+)?date\s*[:.]?\s*(\d{4}-\d{2}-\d{2}|\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}|\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{4})/i;
  const GOODS_EX = /(?:goods|merchandise|products?)\s*(?:total)?\s*\(?\s*(?:ex|excl?)\.?\s*gst\s*\)?\s*:?\s*\$?\s*([\d,]+\.\d{2})/i;
  const SUBTOTAL_EX = /sub\s*-?\s*total\s*\(?\s*(?:ex|excl?)?\.?\s*(?:gst)?\s*\)?\s*:?\s*\$?\s*([\d,]+\.\d{2})/i;
  const PACK = /(?:ctn|carton|box|bx|pk|pack|pkt|case|cs|inner|outer|bdl|bundle|doz)\s*(?:of|x)?\s*(\d+)|(\d+)\s*(?:per|\/)\s*(?:ctn|carton|box|pk|pack|case)|^(?:x\s*)?(\d+)\s*(?:pk|pack|pc)?$/i;
  const GAP = 5.0;       // points between text runs that still belong to the same phrase
  const LINE_TOL = 2.5;  // points of vertical wobble allowed within one line

  class ReadError extends Error {}

  const key = t => String(t).toLowerCase().replace(/[^a-z0-9%]/g, "");
  const round = (n, dp) => { const f = 10 ** dp; return Math.round((n + Math.sign(n) * Number.EPSILON) * f) / f; };

  function number(text) {
    if (text == null || text === "") return null;
    let t = String(text).replace(/[$,\s%]/g, "");
    const neg = (t.startsWith("(") && t.endsWith(")")) || t.endsWith("-") || t.startsWith("-");
    t = t.replace(/^[(\-]+|[)\-]+$/g, "");
    if (!/^\d*\.?\d+$/.test(t)) return null;
    const n = parseFloat(t);
    return neg ? -n : n;
  }

  // ------------------------------------------------------------------ PDF → lines of phrases

  // One pdf.js text item → words with (approximate) x positions. Characters are spread evenly
  // across the item's width, which is close enough to tell neighbouring columns apart.
  function itemWords(item, pageHeight) {
    const str = item.str;
    if (!str || !str.trim()) return null;
    const x = item.transform[4], top = pageHeight - item.transform[5];
    const cw = str.length ? item.width / str.length : 0;
    const words = [];
    const re = /\S+/g;
    let m;
    while ((m = re.exec(str))) words.push([m[0], x + m.index * cw, x + (m.index + m[0].length) * cw]);
    return { text: words.map(w => w[0]).join(" "), x0: words[0][1], x1: words[words.length - 1][2], top, words };
  }

  async function pageLines(page) {
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    const runs = content.items.map(i => itemWords(i, viewport.height)).filter(Boolean);
    runs.sort((a, b) => a.top - b.top || a.x0 - b.x0);
    const rows = [];
    for (const r of runs) {
      const last = rows[rows.length - 1];
      if (last && Math.abs(last[0].top - r.top) <= LINE_TOL) last.push(r); else rows.push([r]);
    }
    return rows.map(row => {
      row.sort((a, b) => a.x0 - b.x0);
      const phrases = [];
      for (const r of row) {
        const p = phrases[phrases.length - 1];
        if (p && r.x0 - p.x1 <= GAP) {
          p.text += " " + r.text; p.x1 = r.x1; p.words.push(...r.words);
        } else {
          phrases.push({ text: r.text, x0: r.x0, x1: r.x1, top: r.top, words: [...r.words] });
        }
      }
      return phrases;
    });
  }

  // ------------------------------------------------------------------ the table

  function roleFor(label) {
    const words = String(label).split(/\s+/).filter(Boolean);
    for (let n = words.length; n > 0; n--) {
      const role = ROLE_OF[key(words.slice(0, n).join(""))];
      if (role) return role;
    }
    return null;
  }

  // One header phrase → its columns. Usually one ("Unit Price (ex GST)"), but narrow columns can
  // sit so close that two names read as one phrase ("Sup Price inc GST").
  function splitHeader(p) {
    const words = p.words, cols = [];
    let i = 0;
    while (i < words.length) {
      let matched = false;
      for (let j = words.length; j > i; j--) {
        const role = ROLE_OF[key(words.slice(i, j).map(w => w[0]).join(""))];
        if (role) {
          cols.push({ role, label: words.slice(i, j).map(w => w[0]).join(" "), x0: words[i][1], x1: words[j - 1][2] });
          i = j; matched = true; break;
        }
      }
      if (!matched) {
        if (cols.length) {   // trailing words like "(ex GST)" belong to the column before
          const c = cols[cols.length - 1];
          c.label += " " + words[i][0]; c.x1 = words[i][2];
        }
        i++;
      }
    }
    return cols;
  }

  function isHeader(cols) {
    const roles = new Set(cols.map(c => c.role));
    return (roles.has("qty") || roles.has("qty_supplied")) && (roles.has("price") || roles.has("amount"))
      && (roles.has("code") || roles.has("desc")) && cols.length >= 3;
  }

  function headerColumns(line) {
    const cols = line.flatMap(splitHeader);
    return isHeader(cols) ? cols : null;
  }

  // Put each phrase under the header column it overlaps most (or is nearest to).
  function assign(line, cols) {
    const cells = {};
    for (const p of line) {
      let best = null, score = -Infinity;
      for (const c of cols) {
        const overlap = Math.min(p.x1, c.x1) - Math.max(p.x0, c.x0);
        const mid = (p.x0 + p.x1) / 2, cmid = (c.x0 + c.x1) / 2;
        const s = overlap > 0 ? overlap : -Math.min(Math.abs(p.x0 - c.x1), Math.abs(c.x0 - p.x1), Math.abs(mid - cmid));
        if (s > score) { best = c; score = s; }
      }
      cells[best.role] = cells[best.role] ? cells[best.role] + " " + p.text : p.text;
    }
    return { cells, before: [], after: [], top: 0 };
  }

  // Give each wrapped description line to whichever item row it sits closer to.
  function placeWrapped(pending, prev, next, nextTop) {
    for (const [top, text] of pending) {
      if (prev && (!next || top - prev.top <= nextTop - top)) prev.after.push(text);
      else if (next) next.before.push(text);
    }
  }

  function packSize(uom) {
    const m = PACK.exec(String(uom || "").trim());
    if (m) {
      const n = parseInt(m[1] || m[2] || m[3], 10);
      if (n > 1) return n;
    }
    return 1;
  }

  // Table rows → invoice lines (units, unit cost ex GST), plus the total of freight-type charges.
  function toLines(rows, cols, taxInclusive) {
    const priceLabel = (cols.find(c => c.role === "price") || {}).label || "";
    if (/\binc/i.test(priceLabel)) taxInclusive = true;
    if (/\bex/i.test(priceLabel)) taxInclusive = false;

    const lines = [];
    let charges = 0;
    for (const r of rows) {
      const c = r.cells;
      const desc = [...r.before, c.desc || "", ...r.after].filter(Boolean).join(" ").trim();
      const code = (c.code || "").trim();
      const amount = number(c.amount);
      // Freight & co. — whether or not the supplier gives it an item code like "FRT".
      if (CHARGE_WORDS.test(desc || Object.values(c).join(" "))) { charges += amount || 0; continue; }
      const qtyOrd = number(c.qty);
      let qty = number(c.qty_supplied);
      if (qty == null) qty = qtyOrd;
      if (qty == null) continue;
      let price = number(c.price);
      let disc = number(c.disc) || 0;
      const pack = packSize(c.uom);
      if (price == null && amount != null && qty) { price = amount / qty; disc = 0; }
      if (price == null) continue;
      let unit = price * (100 - disc) / 100 / pack;
      if (taxInclusive) unit /= 1.1;
      const notes = [];
      if (pack > 1) notes.push(`${qty} x ${c.uom} of ${pack} units`);
      if (disc) notes.push(`${disc}% line discount`);
      if (qty === 0 && qtyOrd) notes.push(`${qtyOrd} ordered, none supplied (backorder)`);
      else if (qtyOrd != null && c.qty_supplied && qty !== qtyOrd) notes.push(`${qtyOrd} ordered, ${qty} supplied`);
      if (price === 0) notes.push("no charge");
      if (taxInclusive) notes.push("GST removed from inc-GST price");
      lines.push({
        supplier_sku: code,
        barcode: (c.barcode || "").trim(),
        description: desc.slice(0, 50),
        qty: Math.round(qty * pack),
        unit_cost_ex_tax: round(unit, 4),
        note: notes.join("; "),
      });
    }
    if (!lines.length) throw new ReadError("Found the table header but no product lines under it.");
    return [lines, charges];
  }

  // A known supplier from the saved supplier names if one appears; else the first phrase on page 1.
  function supplierName(firstPage, text, suppliers) {
    const phrases = firstPage.flat().map(p => p.text);
    for (const name of Object.keys(suppliers || {})) {
      if (text.toLowerCase().includes(name.toLowerCase()))
        return phrases.find(t => t.toLowerCase().includes(name.toLowerCase())) || name;
    }
    return phrases.find(t => !/tax invoice|^invoice|page \d|^abn|^date/i.test(t)) || "";
  }

  async function readPdf(arrayBuffer, suppliers) {
    let pdf;
    try {
      pdf = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer), isEvalSupported: false }).promise;
    } catch (e) {
      throw new ReadError(`Couldn't open the PDF: ${e.message}`);
    }
    const pages = [];
    for (let n = 1; n <= pdf.numPages; n++) pages.push(await pageLines(await pdf.getPage(n)));
    await pdf.destroy();
    const fullText = pages.map(lines => lines.map(l => l.map(p => p.text).join(" ")).join("\n")).join("\n");
    if (!fullText.trim())
      throw new ReadError("This PDF has no text in it — it's probably a scan or photo. This tool needs the PDF " +
                          "the supplier emailed (or ask them for one).");

    let cols = null, pending = [];
    const rows = [];
    let taxInclusive = TAX_INCLUSIVE.test(fullText);
    pages.forEach((page, pageNo) => {
      let inTable = false;
      for (const line of page) {
        const found = headerColumns(line);
        if (found) { cols = found; inTable = true; continue; }
        if (!cols) continue;
        if (!inTable && pageNo > 0) inTable = true;   // a later page without a repeated header
        if (!inTable) continue;
        if (END_OF_TABLE.test(line[0] ? line[0].text : "")) {
          placeWrapped(pending, rows[rows.length - 1], null, Infinity); pending = [];
          inTable = false; continue;
        }
        const row = assign(line, cols);
        const top = line[0].top;
        const qtyText = row.cells.qty_supplied || row.cells.qty;
        const isItem = number(qtyText) != null && (row.cells.price || row.cells.amount);
        if (isItem || CHARGE_WORDS.test(line.map(p => p.text).join(" "))) {
          // Charges are kept so their amount is known; they're filtered out in toLines.
          placeWrapped(pending, rows[rows.length - 1], row, top); pending = [];
          row.top = top;
          rows.push(row);
        } else if (Object.keys(row.cells).every(k => k === "desc" || k === "code")) {
          // A wrapped description line. With vertically-centred cells it can sit above or below
          // the line holding the numbers; decide once the next row turns up.
          pending.push([top, line.map(p => p.text).join(" ")]);
        }
      }
      placeWrapped(pending, rows[rows.length - 1], null, Infinity); pending = [];
    });

    if (!cols)
      throw new ReadError("Couldn't find the line-item table — no header row with a code/description, quantity " +
                          "and price column. Add the supplier's column names to HEADER_WORDS in reader.js.");
    const [lines, charges] = toLines(rows, cols, taxInclusive);
    const priceLabel = (cols.find(c => c.role === "price") || {}).label || "";
    if (/\binc/i.test(priceLabel)) taxInclusive = true;
    if (/\bex/i.test(priceLabel)) taxInclusive = false;

    let subtotal = null;
    if (!taxInclusive) {
      const goods = GOODS_EX.exec(fullText), sub = SUBTOTAL_EX.exec(fullText);
      if (goods) subtotal = number(goods[1]);
      else if (sub) subtotal = round(number(sub[1]) - charges, 2);
    }
    const invNo = INVOICE_NO.exec(fullText), invDate = INVOICE_DATE.exec(fullText);
    return {
      supplier_name: supplierName(pages[0], fullText, suppliers),
      invoice_number: invNo ? invNo[1] : "",
      invoice_date: invDate ? invDate[1] : "",
      subtotal_ex_tax: subtotal,
      lines,
    };
  }

  // ------------------------------------------------------------------ CSV

  function parseCsv(text) {
    const first = text.split(/\r?\n/, 1)[0];
    const delim = [",", ";", "\t", "|"].reduce((a, b) => (first.split(b).length > first.split(a).length ? b : a));
    const rows = [];
    let row = [], field = "", quoted = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
        else if (ch === '"') quoted = false;
        else field += ch;
      } else if (ch === '"') quoted = true;
      else if (ch === delim) { row.push(field); field = ""; }
      else if (ch === "\n" || ch === "\r") {
        if (ch === "\r" && text[i + 1] === "\n") i++;
        row.push(field); rows.push(row); row = []; field = "";
      } else field += ch;
    }
    if (field || row.length) { row.push(field); rows.push(row); }
    return rows.filter(r => r.some(c => c.trim()));
  }

  // A CSV of invoice lines (e.g. from an email-export extension) → same shape as readPdf.
  // The header row is found by the same column names as for PDFs. Supplier, invoice number and
  // date columns are used when present; otherwise the file name stands in for the supplier.
  function readCsv(text, suppliers, name) {
    text = text.replace(/^﻿/, "");
    const table = parseCsv(text);
    const META = {
      supplier: ["supplier", "suppliername", "vendor", "from", "sender", "company"],
      invoice_number: ["invoiceno", "invoicenumber", "invoice", "invoice#", "invno", "docno"],
      invoice_date: ["invoicedate", "date", "docdate"],
    };
    let h = -1, roles = [];
    for (let i = 0; i < table.length; i++) {
      roles = table[i].map(roleFor);
      const cols = roles.map((r, j) => r && { role: r, label: table[i][j] }).filter(Boolean);
      if (isHeader(cols)) { h = i; break; }
    }
    if (h < 0) throw new ReadError("Couldn't find a header row with a code/description, quantity and price column in the CSV.");
    const header = table[h];
    const cols = roles.map((r, j) => r && { role: r, label: header[j] }).filter(Boolean);
    const rows = table.slice(h + 1).map(rec => {
      const cells = {};
      roles.forEach((r, j) => { if (r && rec[j] && rec[j].trim() && !(r in cells)) cells[r] = rec[j].trim(); });
      return { cells, before: [], after: [], top: 0 };
    }).filter(r => Object.keys(r.cells).length);
    const [lines] = toLines(rows, cols, TAX_INCLUSIVE.test(text));
    const meta = m => {
      const j = header.findIndex((c, k) => META[m].includes(key(c)) && !roles[k]);
      if (j < 0) return "";
      const rec = table.slice(h + 1).find(r => r[j] && r[j].trim());
      return rec ? rec[j].trim() : "";
    };
    let supplier = meta("supplier");
    if (!supplier) supplier = Object.keys(suppliers || {}).find(f => text.toLowerCase().includes(f.toLowerCase())) || "";
    return {
      supplier_name: supplier || String(name || "").replace(/\.[^.]+$/, ""),
      invoice_number: meta("invoice_number"),
      invoice_date: meta("invoice_date"),
      subtotal_ex_tax: null,
      lines,
    };
  }

  window.InvoiceReader = { readPdf, readCsv, ReadError, HEADER_WORDS };
})();
