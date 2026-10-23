# invoice-to-po

Supplier PDF invoice → filled-in Retail Express purchase-order import sheet.

## Setup (once)

1. Install [uv](https://docs.astral.sh/uv/) (Mac: `brew install uv`; Windows: see the uv site).
2. Get an Anthropic API key at https://console.anthropic.com → API Keys. Reading invoices uses it
   (a few cents per invoice).
3. Copy `.env.example` to `.env` and paste the key in.

## The web app

Double-click `start.command` (Mac) or `start.bat` (Windows), or run `uv run app.py`.
It opens http://127.0.0.1:8765, which only this computer can reach.

- **Supplier templates**: add each template downloaded from Retail Express (see below) once.
- **Drop invoices** onto the page. Each one is read, matched to its supplier's template and shown
  as a table you can correct (SKU, qty, unit cost, remove lines) before downloading the PO sheet.
- Lines not in Retail Express are highlighted; the lines total is checked against the invoice subtotal.
- Choosing a template for an unknown supplier remembers that supplier for next time.

Try it with the made-up invoices in `samples/test-invoices/` (regenerate with
`uv run samples/make_test_invoices.py`):

    uv run app.py --templates samples/test-invoices/templates --suppliers samples/test-invoices/suppliers.json

The `.expected.json` files there can be dropped on the page too: they show the right result
without needing an API key.

## Command line

## Get each supplier's template (once per supplier, refresh when their range changes)

In Retail Express: **Get Excel Template** → *1 Supplier, for 1 Outlet* → pick the outlet
(e.g. The Rocks) and supplier (e.g. FG - Fila Group) → File Format **.xlsx** → **Get Excel Template**.
Save the file into `templates/` here. Any filename works: the tool reads the Supplier box inside the file.

Then add the supplier's invoice name to `suppliers.json` (copy `suppliers.example.json`),
so the tool knows which template an invoice belongs to:

    {"Fila Group": "FG", "Faber-Castell": "FC"}

    uv run invoice_to_po.py ~/Downloads/invoice.pdf

- Several invoices at once, even from different suppliers: `uv run invoice_to_po.py ~/Downloads/*.pdf`.
  You get one .xlsx per invoice, named `PO <code> <invoice#>.xlsx`, ready for **Upload Excel**.
- `--supplier FG` picks the template directly; `--template file.xlsx` uses one file for everything.
- `--save-json` keeps what was read from the PDF. Edit the .json and run the tool on it to rebuild the
  sheet without reading the PDF again.

## What goes where

The template already has Outlet and Supplier filled in; the tool leaves them alone (`--outlet` overrides).

| Column | Filled with |
|---|---|
| ProductId | the matching product from the template's master sheet, found by supplier SKU, then barcode |
| SupplierSKU | the supplier's code as Retail Express has it (from the invoice when there's no match) |
| ShortDescription | invoice description (reference only) |
| Qty | units: cartons are converted to units when the pack size is printed |
| SupplierBuy | unit cost ex GST, after line discounts |
| SupplierSKU2 | blank (`--sku2 barcode` puts the invoice barcode there) |

Codes match loosely: case, spaces, dashes, dots and leading zeros are ignored (`FG-0012` = `fg0012`).

Freight, surcharges and similar lines are skipped. A second sheet, **Invoice Check**, lists every line,
the field it matched on, and notes, and highlights lines that **aren't in Retail Express** for that
supplier (new products to create first). The tool also warns when the lines don't add up to the
invoice subtotal.
