// The page: supplier templates (kept in this browser), dropped invoices, editable lines, downloads.
(function () {
  "use strict";

  const $ = (sel, el = document) => el.querySelector(sel);
  const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const cost = v => (+v).toFixed(4).replace(/0{1,2}$/, "");   // 5.5 → 5.50, 8.3333 stays
  const money = n => n == null ? "—" : "$" + Number(n).toLocaleString("en-AU", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // ---------------------------------------------------------------- storage (this browser only)

  // Templates are binary .xlsx files, so they live in IndexedDB; the supplier-name → code map is
  // small and lives in localStorage. Both survive closing the page, per browser and computer.
  const DB = "invoice-to-po", STORE = "templates";
  function db() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: "key" });
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function tx(mode, fn) {
    const d = await db();
    return new Promise((resolve, reject) => {
      const t = d.transaction(STORE, mode);
      const out = fn(t.objectStore(STORE));
      t.oncomplete = () => resolve(out && out.result);
      t.onerror = () => reject(t.error);
    });
  }
  const storage = {
    templates: () => tx("readonly", s => s.getAll()).catch(() => []),
    putTemplate: rec => tx("readwrite", s => s.put(rec)),
    deleteTemplate: key => tx("readwrite", s => s.delete(key)),
    suppliers() { try { return JSON.parse(localStorage.getItem("suppliers") || "{}"); } catch { return {}; } },
    saveSuppliers(map) { try { localStorage.setItem("suppliers", JSON.stringify(map)); } catch {} },
  };

  let templates = [];          // [{key, code, outlet, file, buf}]
  let suppliers = storage.suppliers();
  const invoices = [];         // {id, file, status, error, data, code, rawCode, outlet, preview}
  let nextId = 1;
  const templateFor = code => templates.find(t => t.key === PO.norm(code));

  // ---------------------------------------------------------------- templates

  async function loadTemplates() {
    templates = (await storage.templates()).sort((a, b) => a.code.localeCompare(b.code));
    renderTemplates();
  }

  function renderTemplates() {
    $("#templateCount").textContent = templates.length ? `${templates.length} supplier${templates.length > 1 ? "s" : ""} · saved in this browser` : "";
    $("#templateTable").innerHTML = templates.length
      ? `<tr><th>Supplier code</th><th>Outlet</th><th>File</th><th></th></tr>` +
        templates.map(x => `<tr><td><b>${esc(x.code)}</b></td><td>${esc(x.outlet)}</td><td class="muted">${esc(x.file)}</td>
          <td class="num"><button class="link" data-del="${esc(x.key)}" title="Remove">Remove</button></td></tr>`).join("")
      : `<tr><td class="muted empty">No templates yet — without one, POs use a stand-in layout and can't be matched to ProductIds.</td></tr>`;
  }

  $("#templateTable").addEventListener("click", async e => {
    const key = e.target.dataset.del;
    const t = key && templates.find(x => x.key === key);
    if (!t || !confirm(`Remove the ${t.code} template?`)) return;
    await storage.deleteTemplate(key);
    await loadTemplates();
    invoices.forEach(i => { if (i.code && !templateFor(i.code)) i.code = null; refresh(i); });
  });

  $("#templateInput").addEventListener("change", async e => {
    for (const f of e.target.files) {
      try {
        const buf = await f.arrayBuffer();
        const { code, outlet } = await PO.describeTemplate(buf);
        if (!code) throw new Error("its Supplier box is empty — download it with a supplier selected");
        await storage.putTemplate({ key: PO.norm(code), code, outlet, file: f.name, buf });
      } catch (err) { alert(`${f.name}: ${err.message}`); }
    }
    e.target.value = "";
    await loadTemplates();
    invoices.filter(i => i.status === "ready").forEach(i => { if (!i.code) i.code = guessCode(i.data); refresh(i); });
  });

  function guessCode(data) {
    const name = (data.supplier_name || "").toLowerCase();
    for (const [frag, code] of Object.entries(suppliers))
      if (name.includes(frag.toLowerCase()) && templateFor(code)) return code;
    return null;
  }

  // "Fila Group Pty Ltd" → "Fila Group": the part of the name that stays the same across invoices.
  function rememberKey(name) {
    return (name || "").replace(/\b(pty\.?|ltd\.?|limited|inc\.?|co\.?|&)\s*/gi, " ").trim().split(/\s+/).slice(0, 2).join(" ");
  }

  // ---------------------------------------------------------------- invoices

  const drop = $("#drop");
  ["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", e => addFiles(e.dataTransfer.files));
  $("#fileInput").addEventListener("change", e => { addFiles(e.target.files); e.target.value = ""; });

  function addFiles(files) {
    for (const file of files) {
      const inv = { id: nextId++, file, status: "reading", data: null, code: null, rawCode: "", outlet: "", preview: null };
      invoices.unshift(inv);
      const el = document.createElement("section");
      el.className = "panel"; el.id = `inv-${inv.id}`;
      $("#invoices").prepend(el);
      render(inv);
      read(inv);
    }
  }

  async function read(inv) {
    try {
      const name = inv.file.name.toLowerCase();
      if (name.endsWith(".csv") || name.endsWith(".tsv") || name.endsWith(".txt"))
        inv.data = InvoiceReader.readCsv(await inv.file.text(), suppliers, inv.file.name);
      else if (name.endsWith(".json"))
        inv.data = JSON.parse(await inv.file.text());
      else
        inv.data = await InvoiceReader.readPdf(await inv.file.arrayBuffer(), suppliers);
      inv.code = guessCode(inv.data);
      inv.status = "ready";
      await refresh(inv);
    } catch (err) {
      inv.status = "error"; inv.error = err.message; render(inv);
    }
  }

  async function build(inv) {
    const t = inv.code ? templateFor(inv.code) : null;
    return PO.buildPo(inv.data, {
      template: t ? t.buf : null,
      code: t ? null : inv.rawCode || null,
      outlet: t ? null : inv.outlet || null,
      source: inv.file.name,
    });
  }

  async function refresh(inv) {
    if (inv.status !== "ready") return;
    // Without a template the PO needs an outlet and a supplier code typed in.
    if (!inv.code && !(inv.outlet && inv.rawCode)) { inv.preview = null; inv.previewError = null; render(inv); return; }
    try {
      inv.preview = await build(inv);
      inv.previewError = null;
    } catch (err) { inv.preview = null; inv.previewError = err.message; }
    render(inv);
  }

  let editTimer;
  function edited(inv) { clearTimeout(editTimer); editTimer = setTimeout(() => refresh(inv), 300); }

  async function download(inv) {
    const po = await build(inv);   // fresh, so it carries every edit
    const blob = new Blob([await po.workbook.xlsx.writeBuffer()],
                          { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = po.filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  function render(inv) {
    const el = $(`#inv-${inv.id}`);
    if (!el) return;
    if (inv.status === "reading") {
      el.innerHTML = `<div class="card-head"><span class="title">${esc(inv.file.name)}</span>
        <span class="muted"><span class="spinner"></span>&nbsp; Reading invoice…</span></div>`;
      return;
    }
    if (inv.status === "error") {
      el.innerHTML = `<div class="card-head"><span class="title">${esc(inv.file.name)}</span><span class="pill bad">Couldn't read it</span>
        <span class="spacer" style="flex:1"></span><button class="link" data-act="close">Dismiss</button></div>
        <div class="panel-body" style="color:var(--bad)">${esc(inv.error)}</div>`;
      el.querySelector("[data-act=close]").onclick = () => removeInv(inv);
      return;
    }

    const d = inv.data, p = inv.preview;
    const onPo = d.lines.filter(l => l.qty > 0).length;
    const matched = p ? p.lines.filter((l, i) => l.product_id && d.lines[i] && d.lines[i].qty > 0).length : 0;
    const options = [`<option value="">— choose supplier template —</option>`]
      .concat(templates.map(t => `<option value="${esc(t.code)}" ${inv.code && t.key === PO.norm(inv.code) ? "selected" : ""}>${esc(t.code)} · ${esc(t.outlet)}</option>`));
    const noTemplate = !inv.code;

    const rows = d.lines.map((l, i) => {
      const pl = (p && p.lines[i]) || {};
      const skip = l.qty <= 0;
      const unmatched = p && p.hasMaster && !pl.product_id && !skip;
      const badge = skip ? `<span class="pill muted">not on PO</span>`
        : pl.product_id ? `<span class="pill ok" title="matched on ${esc(pl.matched_on)}">${esc(pl.product_id)}</span>`
        : p && p.hasMaster ? `<span class="pill warn">not in Retail Express</span>` : `<span class="muted">—</span>`;
      return `<tr class="${skip ? "skip" : unmatched ? "unmatched" : ""}" data-i="${i}">
        <td>${badge}</td>
        <td><input class="sku" data-f="supplier_sku" value="${esc(l.supplier_sku)}"></td>
        <td>${esc(l.description)}${l.barcode ? `<div class="muted small">${esc(l.barcode)}</div>` : ""}</td>
        <td class="num"><input class="qty" type="number" step="1" data-f="qty" value="${l.qty}"></td>
        <td class="num"><input class="cost" type="number" step="0.01" data-f="unit_cost_ex_tax" value="${cost(l.unit_cost_ex_tax)}"></td>
        <td class="num">${money(l.qty * l.unit_cost_ex_tax)}</td>
        <td class="note">${esc(l.note)}</td>
        <td class="num"><button class="link" data-act="del" title="Remove line">✕</button></td></tr>`;
    }).join("");

    const total = d.lines.reduce((s, l) => s + Math.max(l.qty, 0) * l.unit_cost_ex_tax, 0);
    const sub = d.subtotal_ex_tax;
    const diff = sub != null && Math.abs(sub - total) > 0.05;
    const warnings = ((p && p.warnings) || []).concat(inv.previewError ? [inv.previewError] : []);

    el.innerHTML = `
      <div class="card-head">
        <span class="title">${esc(d.supplier_name || "Unknown supplier")}</span>
        <span class="muted">Invoice ${esc(d.invoice_number)} · ${esc(d.invoice_date)} · ${esc(inv.file.name)}</span>
        <span class="spacer" style="flex:1"></span>
        ${p ? (p.hasMaster ? `<span class="pill ${matched === onPo ? "ok" : "warn"}">${matched} of ${onPo} matched</span>` : `<span class="pill warn">no product matching</span>`) : ""}
        <button class="link" data-act="close" title="Remove this invoice">Remove</button>
      </div>
      <div class="controls">
        <label>Supplier <select data-act="code">${options.join("")}</select></label>
        ${noTemplate ? `<label>Outlet <input data-act="outlet" placeholder="e.g. The Rocks" value="${esc(inv.outlet)}"></label>
                        <label>Code <input data-act="rawcode" placeholder="e.g. FG" size="5" value="${esc(inv.rawCode)}"></label>` : ""}
        ${noTemplate && templates.length ? `<span class="muted small">Pick the supplier's template to match products — it's remembered for next time.</span>` : ""}
      </div>
      <div class="table-wrap"><table>
        <tr><th>ProductId</th><th>Supplier SKU</th><th>Description</th><th class="num">Qty</th><th class="num">Unit ex GST</th>
            <th class="num">Line</th><th>Note</th><th></th></tr>
        ${rows}
      </table></div>
      <div class="foot">
        ${warnings.length ? `<ul>${warnings.map(w => `<li>${esc(w)}</li>`).join("")}</ul>`
          : p ? `<span class="pill ok">No problems found</span>`
          : `<span class="muted">Choose the supplier's template — or, without one, type the outlet and supplier code.</span>`}
        <div class="totals">
          <div>Lines total ex GST <b>${money(total)}</b></div>
          <div class="${diff ? "" : "muted"}" style="${diff ? "color:var(--bad)" : ""}">Invoice subtotal ${money(sub)}</div>
          <div style="margin-top:8px"><button class="primary" data-act="download" ${p ? "" : "disabled"}>Download ${esc(p ? p.filename : "PO")}</button></div>
        </div>
      </div>`;

    el.querySelector("[data-act=close]").onclick = () => removeInv(inv);
    el.querySelector("[data-act=download]").onclick = () => download(inv).catch(e => alert(e.message));
    el.querySelector("[data-act=code]").onchange = e => {
      inv.code = e.target.value || null;
      if (inv.code && d.supplier_name) {
        suppliers[rememberKey(d.supplier_name)] = inv.code;
        storage.saveSuppliers(suppliers);
      }
      refresh(inv);
    };
    const outlet = el.querySelector("[data-act=outlet]");
    if (outlet) { outlet.oninput = e => { inv.outlet = e.target.value; }; outlet.onchange = () => refresh(inv); }
    const raw = el.querySelector("[data-act=rawcode]");
    if (raw) raw.onchange = e => { inv.rawCode = e.target.value.trim(); refresh(inv); };
    el.querySelectorAll("tr[data-i]").forEach(tr => {
      const line = d.lines[+tr.dataset.i];
      tr.querySelectorAll("input[data-f]").forEach(input => input.onchange = () => {
        const f = input.dataset.f;
        line[f] = f === "qty" ? parseInt(input.value || "0", 10) : f === "unit_cost_ex_tax" ? parseFloat(input.value || "0") : input.value;
        edited(inv);
      });
      tr.querySelector("[data-act=del]").onclick = () => { d.lines.splice(+tr.dataset.i, 1); refresh(inv); };
    });
  }

  function removeInv(inv) {
    invoices.splice(invoices.indexOf(inv), 1);
    const el = $(`#inv-${inv.id}`);
    if (el) el.remove();
  }

  // For the self-test and for poking at it from the console.
  window.App = { addFiles, invoices, loadTemplates };
  loadTemplates();
})();
