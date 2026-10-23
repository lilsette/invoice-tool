# invoice-to-po

Supplier PDF invoice → filled-in Retail Express purchase-order import sheet.

## Setup (once)

Needs [uv](https://docs.astral.sh/uv/) and an Anthropic API key:

    export ANTHROPIC_API_KEY=sk-ant-...

## Get each supplier's template (once per supplier, refresh when their range changes)

In Retail Express: **Get Excel Template** → *1 Supplier, for 1 Outlet* → pick the outlet
(e.g. The Rocks) and supplier (e.g. FG - Fila Group) → File Format **.xlsx** → **Get Excel Template**.
Save the file into `templates/` here. Any filename works: the tool reads the Supplier box inside the file.

Then add the supplier's invoice name to `suppliers.json` (copy `suppliers.example.json`),
so the tool knows which template an invoice belongs to:

    {"Fila Group": "FG", "Faber-Castell": "FC"}

## Use

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
