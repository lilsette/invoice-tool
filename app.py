# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "pdfplumber>=0.11", "fastapi>=0.110", "uvicorn>=0.29", "python-multipart>=0.0.9"]
# ///
"""Local web front-end for invoice_to_po.

    uv run app.py                      # http://127.0.0.1:8765
    uv run app.py --templates samples/test-invoices/templates --suppliers samples/test-invoices/suppliers.json

Runs on this computer only (127.0.0.1), with no internet or AI: invoices are read by
invoice_reader.py, and templates, supplier codes and the finished sheets stay on disk here.
"""

from __future__ import annotations

import argparse
import io
import json
import re
import threading
import webbrowser
from pathlib import Path

import uvicorn
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel

import invoice_to_po as core
from invoice_reader import ReadError

HERE = Path(__file__).parent
app = FastAPI(title="Invoice → PO")
CONFIG = {"templates": HERE / "templates", "suppliers": HERE / "suppliers.json"}


def templates() -> dict[str, dict]:
    return core.template_index(CONFIG["templates"])


def template_list() -> list[dict]:
    return [{"code": t["supplier"], "outlet": t["outlet"], "file": t["path"].name} for t in templates().values()]


@app.get("/")
def index():
    return FileResponse(HERE / "web" / "index.html")


@app.get("/api/status")
def status():
    return {
        "templates": template_list(),
        "suppliers": core.load_suppliers(CONFIG["suppliers"]),
    }


@app.post("/api/templates")
async def upload_template(file: UploadFile = File(...)):
    raw = await file.read()
    try:
        wb = core.load_workbook(io.BytesIO(raw))
        ws = core.po_sheet(wb)
    except Exception as e:
        raise HTTPException(400, f"That doesn't look like a Retail Express PO template: {e}")
    code = core.box_value(ws, "Supplier")
    if not code:
        raise HTTPException(400, "The template's Supplier box is empty — download it with a supplier selected")
    CONFIG["templates"].mkdir(parents=True, exist_ok=True)
    # One file per supplier code: a fresh download replaces the old one.
    existing = templates().get(core.norm(code))
    if existing:
        existing["path"].unlink()
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", str(code))
    (CONFIG["templates"] / f"{safe}.xlsx").write_bytes(raw)
    return {"templates": template_list()}


@app.delete("/api/templates/{code}")
def delete_template(code: str):
    t = templates().get(core.norm(code))
    if not t:
        raise HTTPException(404, "no such template")
    t["path"].unlink()
    return {"templates": template_list()}


@app.post("/api/extract")
async def extract(file: UploadFile = File(...)):
    raw = await file.read()
    name = file.filename or "invoice"
    suppliers = core.load_suppliers(CONFIG["suppliers"])
    if name.lower().endswith(".pdf") and not raw.startswith(b"%PDF"):
        raise HTTPException(400, "Not a PDF")
    try:
        data = core.read_any(name, raw, suppliers)
    except ReadError as e:
        raise HTTPException(422, str(e))
    code = core.code_for(core.to_invoice(data), suppliers)
    return {"data": data, "code": code if code and core.norm(code) in templates() else None}


class BuildRequest(BaseModel):
    data: dict
    code: str | None = None       # supplier code → template
    outlet: str | None = None     # only needed without a template
    sku2: str = "none"
    source: str = ""
    remember_as: str | None = None  # invoice-name fragment to remember for this code


def build(req: BuildRequest) -> core.PoResult:
    t = templates().get(core.norm(req.code)) if req.code else None
    if req.code and req.remember_as:
        suppliers = core.load_suppliers(CONFIG["suppliers"])
        if suppliers.get(req.remember_as) != req.code:
            suppliers[req.remember_as] = req.code
            CONFIG["suppliers"].write_text(json.dumps(suppliers, indent=2))
    try:
        return core.build_po(req.data, template=t["path"] if t else None,
                             outlet=None if t else req.outlet, supplier=None if t else req.code,
                             sku2=req.sku2, source=req.source)
    except (RuntimeError, KeyError, ValueError) as e:
        raise HTTPException(400, str(e))


@app.post("/api/preview")
def preview(req: BuildRequest):
    r = build(req)
    return {
        "filename": r.filename,
        "supplier_code": r.supplier_code,
        "template": r.template.name if r.template else None,
        "has_master": r.has_master,
        "total": float(r.total),
        "subtotal": None if r.invoice.subtotal_ex_tax is None else float(r.invoice.subtotal_ex_tax),
        "warnings": r.warnings,
        "lines": [{"product_id": l.product_id, "matched_on": l.matched_on, "supplier_sku": l.supplier_sku}
                  for l in r.invoice.lines],
    }


@app.post("/api/download")
def download(req: BuildRequest):
    r = build(req)
    buf = io.BytesIO()
    r.workbook.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{r.filename}"'})


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--port", type=int, default=8765)
    p.add_argument("--templates", type=Path, default=CONFIG["templates"])
    p.add_argument("--suppliers", type=Path, default=CONFIG["suppliers"])
    p.add_argument("--no-browser", action="store_true", help="don't open the page on start")
    args = p.parse_args()
    CONFIG["templates"], CONFIG["suppliers"] = args.templates.resolve(), args.suppliers.resolve()
    url = f"http://127.0.0.1:{args.port}"
    print(f"Invoice → PO on {url}  (templates: {CONFIG['templates']})")
    if not args.no_browser:
        threading.Timer(1.0, webbrowser.open, [url]).start()
    uvicorn.run(app, host="127.0.0.1", port=args.port, log_level="warning")


if __name__ == "__main__":
    main()
