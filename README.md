# invoice-to-po

Supplier PDF invoice → filled-in Retail Express "Purchase Order Details" import sheet.

## Setup (once)

Needs [uv](https://docs.astral.sh/uv/) and an Anthropic API key:

    export ANTHROPIC_API_KEY=sk-ant-...

## Use

    uv run invoice_to_po.py invoice.pdf --outlet "The Rocks" --supplier FG

- Several PDFs at once: `uv run invoice_to_po.py ~/Downloads/inv*.pdf --outlet "The Rocks"`.
  You get one .xlsx per invoice, named `PO <supplier> <invoice#>.xlsx`.
- `--template RetailExpress.xlsx`: fill your real template file. The tool finds the
  Outlet/Supplier boxes and the header row by their labels.
- `suppliers.json` (copy `suppliers.example.json`) maps part of a supplier's name to its
  Retail Express supplier code, so you can leave out `--supplier`.
- `--save-json` keeps the raw extraction. Edit the .json and run the tool on it
  to rebuild the sheet without reading the PDF again.

## What goes where

| Column | Filled with |
|---|---|
| SupplierSKU | supplier's item code from the invoice |
| SupplierSKU2 | barcode (EAN/APN), when printed (`--sku2 none` to leave blank) |
| ShortDescription | invoice description (reference only) |
| Qty | units: cartons are converted to units when the pack size is printed |
| SupplierBuy | unit cost ex GST, after line discounts |
| ProductId, ManufacturerSKU | left blank: Retail Express looks these up |

Freight, surcharges and similar lines are skipped. A second sheet, **Invoice Check**, lists every
line with notes and compares the line total against the invoice subtotal. The tool also prints a
warning when those two numbers disagree, or when a line has no supplier SKU.
