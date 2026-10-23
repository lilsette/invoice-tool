// Fill a Retail Express "Purchase Order Details" import sheet from one read invoice (via ExcelJS).
//
// Retail Express makes a different template per supplier + outlet (Get Excel Template →
// 1 Supplier, for 1 Outlet). Its Outlet and Supplier boxes come filled in, and a master sheet
// lists the supplier's products; invoice lines are matched against it to get ProductIds.
(function () {
  "use strict";

  const COLUMNS = ["ProductId", "ManufacturerSKU", "SupplierSKU", "ShortDescription", "SupplierSKU2", "Qty", "SupplierBuy"];
  const SKU_HEADERS = ["suppliersku", "suppliersku2", "manufacturersku", "sku"];
  const BARCODE_HEADERS = ["barcode", "ean", "apn", "upc", "gtin", "productbarcode"];

  // Compare codes loosely: case, spaces, dashes and dots don't matter.
  const norm = v => (v == null ? "" : String(v).toUpperCase().replace(/[^A-Z0-9]/g, ""));
  const headerKey = v => norm(v).toLowerCase();
  const round2 = n => Math.round((n + Math.sign(n) * Number.EPSILON) * 100) / 100;
  const text = cell => (cell && cell.value != null ? String(cell.text ?? "").trim() : "");

  function findCell(ws, label) {
    const want = label.replace(/\s/g, "").toLowerCase();
    let found = null;
    ws.eachRow({ includeEmpty: false }, row => {
      if (found) return;
      row.eachCell({ includeEmpty: false }, cell => {
        if (!found && typeof cell.value === "string" && cell.value.replace(/\s/g, "").toLowerCase() === want) found = cell;
      });
    });
    return found;
  }

  function poSheet(wb) {
    const ws = wb.worksheets.find(s => findCell(s, "SupplierSKU") && findCell(s, "Qty"));
    if (!ws) throw new Error("This isn't a Retail Express PO template — no sheet with 'SupplierSKU' and 'Qty' headers.");
    return ws;
  }

  function boxValue(ws, label) {
    const c = findCell(ws, label);
    return c ? text(ws.getCell(c.row, c.col + 1)) : "";
  }

  async function load(buf) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    return wb;
  }

  // Supplier code and outlet a downloaded template is for.
  async function describeTemplate(buf) {
    const ws = poSheet(await load(buf));
    return { code: boxValue(ws, "Supplier"), outlet: boxValue(ws, "Outlet") };
  }

  // A stand-in laid out like the Retail Express template, for suppliers without one.
  function standIn() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Purchase Order Details");
    const fill = argb => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
    ws.getCell("A1").value = "Purchase Orders - Purchase Order Details";
    ws.getCell("A1").font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
    for (const c of "ABCD") ws.getCell(`${c}1`).fill = fill("FF000080");
    ws.getCell("A3").value = "Outlet"; ws.getCell("A4").value = "Supplier";
    for (const r of [3, 4]) { ws.getCell(`A${r}`).font = { bold: true }; ws.getCell(`B${r}`).fill = fill("FFD9341F"); }
    const types = [["Integer", "Whole Number"], ["Reference Only", "See Master Sheet"], ["Text (lookup)", "50 chars"],
                   ["Reference Only", "See Master Sheet"], ["Text (lookup)", "50 chars"], ["Integer", "Whole Number"],
                   ["Money", "$123.45"]];
    const required = new Set(["ProductId", "SupplierSKU", "SupplierSKU2", "Qty"]);
    COLUMNS.forEach((name, i) => {
      ws.getCell(6, i + 1).value = types[i][0];
      ws.getCell(7, i + 1).value = types[i][1];
      const h = ws.getCell(9, i + 1);
      h.value = name;
      h.font = { bold: true, color: { argb: required.has(name) ? "FFFFFFFF" : "FF000000" } };
      if (required.has(name)) h.fill = fill("FFD9341F");
      ws.getColumn(i + 1).width = 22;
    });
    return wb;
  }

  // The product list Retail Express puts in a supplier's template ("See Master Sheet").
  function readMaster(wb, skip) {
    for (const ws of wb.worksheets) {
      if (ws === skip) continue;
      for (let r = 1; r <= Math.min(ws.rowCount, 30); r++) {
        const heads = {};
        ws.getRow(r).eachCell({ includeEmpty: false }, (cell, col) => { heads[headerKey(text(cell))] = col; });
        if (!("productid" in heads) || ![...SKU_HEADERS, ...BARCODE_HEADERS].some(h => h in heads)) continue;
        const index = {};
        let count = 0;
        for (let i = r + 1; i <= ws.rowCount; i++) {
          const row = {};
          for (const [h, col] of Object.entries(heads)) row[h] = text(ws.getRow(i).getCell(col));
          if (!row.productid) continue;
          count++;
          for (const h of [...SKU_HEADERS, ...BARCODE_HEADERS]) {
            const k = norm(row[h]);
            if (!k) continue;
            index[h] = index[h] || {};
            index[h][k] = index[h][k] || row;
            const stripped = k.replace(/^0+/, "");
            index[h][stripped] = index[h][stripped] || row;
          }
        }
        if (count) return index;
      }
    }
    return null;
  }

  function match(index, line) {
    const tries = [[line.supplier_sku, [...SKU_HEADERS, ...BARCODE_HEADERS]],
                   [line.barcode, [...BARCODE_HEADERS, ...SKU_HEADERS]]];
    for (const [value, cols] of tries) {
      const k = norm(value);
      if (!k) continue;
      for (const col of cols) {
        const idx = index[col] || {};
        const hit = idx[k] || idx[k.replace(/^0+/, "")];
        if (hit) return [hit, col];
      }
    }
    return null;
  }

  // One invoice → one filled workbook plus what the page shows.
  // `code`/`outlet` override the template's boxes (and are required without a template).
  async function buildPo(data, { template = null, code = null, outlet = null, sku2 = "none", source = "" } = {}) {
    const warnings = [];
    const wb = template ? await load(template) : standIn();
    if (!template) warnings.push("No Retail Express template for this supplier — used a stand-in layout, so lines could not be matched to ProductIds");
    const ws = poSheet(wb);
    const master = readMaster(wb, ws);

    const lines = data.lines.map(l => ({ ...l, product_id: "", matched_on: "" }));
    if (master) for (const l of lines) {
      const hit = match(master, l);
      if (hit) {
        l.product_id = hit[0].productid; l.matched_on = hit[1];
        if (hit[0].suppliersku) l.supplier_sku = hit[0].suppliersku;   // as Retail Express spells it
      }
    }

    // The template arrives with Outlet and Supplier filled in; only overwrite when told to.
    for (const [label, value] of [["Outlet", outlet], ["Supplier", code]]) {
      const cell = findCell(ws, label);
      if (!cell) throw new Error(`The template has no '${label}' box`);
      const box = ws.getCell(cell.row, cell.col + 1);
      if (value) box.value = value;
      else if (!text(box)) throw new Error(`The template's ${label} box is empty — set the ${label.toLowerCase()}`);
    }

    const header = findCell(ws, "SupplierSKU");
    const cols = {};
    ws.getRow(header.row).eachCell({ includeEmpty: false }, (cell, col) => { cols[headerKey(text(cell))] = col; });
    const first = header.row + 1;
    for (let r = first; r <= ws.rowCount; r++) for (const c of Object.values(cols)) ws.getCell(r, c).value = null;
    let r = first;
    for (const l of lines) {
      if (l.qty <= 0) continue;   // backordered / nil-supplied: on the check sheet, not the PO
      if (l.product_id && cols.productid) ws.getCell(r, cols.productid).value = /^\d+$/.test(l.product_id) ? Number(l.product_id) : l.product_id;
      ws.getCell(r, cols.suppliersku).value = l.supplier_sku || null;
      if (sku2 === "barcode" && cols.suppliersku2) ws.getCell(r, cols.suppliersku2).value = l.barcode || null;
      if (cols.shortdescription) ws.getCell(r, cols.shortdescription).value = l.description;
      ws.getCell(r, cols.qty).value = l.qty;
      if (cols.supplierbuy) {
        const c = ws.getCell(r, cols.supplierbuy);
        c.value = l.unit_cost_ex_tax;
        c.numFmt = '"$"#,##0.00##';
      }
      r++;
    }

    // The check sheet: every line, what it matched on, and anything to look at.
    const total = round2(lines.reduce((s, l) => s + Math.max(l.qty, 0) * l.unit_cost_ex_tax, 0));
    const chk = wb.addWorksheet("Invoice Check");
    chk.addRow(["Source", source]);
    chk.addRow(["Supplier on invoice", data.supplier_name]);
    chk.addRow(["Invoice #", data.invoice_number]);
    chk.addRow(["Invoice date", data.invoice_date]);
    chk.addRow(["Lines total ex tax", total]);
    chk.addRow(["Invoice subtotal ex tax", data.subtotal_ex_tax]);
    chk.addRow([]);
    chk.addRow(["ProductId", "SupplierSKU", "Barcode", "Description", "Qty", "Unit ex tax", "Line ex tax", "Matched on", "Note"]).font = { bold: true };
    for (const l of lines) {
      const row = chk.addRow([l.product_id, l.supplier_sku, l.barcode, l.description, l.qty, l.unit_cost_ex_tax,
        round2(l.qty * l.unit_cost_ex_tax), l.matched_on || (master ? "NOT IN RETAIL EXPRESS" : ""), l.note]);
      if ((master && !l.product_id) || l.qty <= 0)
        row.eachCell({ includeEmpty: true }, c => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE699" } }; });
    }
    [10, 18, 16, 40, 8, 12, 12, 22, 50].forEach((w, i) => { chk.getColumn(i + 1).width = w; });

    if (data.subtotal_ex_tax != null && Math.abs(data.subtotal_ex_tax - total) > 0.05)
      warnings.push(`Lines add to $${total.toFixed(2)} but the invoice subtotal is $${Number(data.subtotal_ex_tax).toFixed(2)}`);
    for (const l of lines) {
      const label = `${l.supplier_sku || "-"} ${l.description}`;
      if (l.qty <= 0) warnings.push(`Left off the PO (qty ${l.qty}): ${label}`);
      else if (master && !l.product_id) warnings.push(`Not in Retail Express for this supplier: ${label}`);
      else if (!l.supplier_sku && !l.product_id) warnings.push(`No supplier SKU: ${l.description}`);
    }
    const supplierCode = boxValue(ws, "Supplier") || "UNKNOWN";
    const stem = String(source).replace(/\.[^.]+$/, "");
    const filename = `PO ${supplierCode} ${data.invoice_number || stem}.xlsx`.replace(/[/\\]/g, "-");
    return {
      workbook: wb, filename, supplierCode, hasMaster: !!master, total, warnings,
      lines: lines.map(l => ({ product_id: l.product_id, matched_on: l.matched_on, supplier_sku: l.supplier_sku })),
    };
  }

  window.PO = { buildPo, describeTemplate, norm };
})();
