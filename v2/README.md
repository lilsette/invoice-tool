# invoice-to-po v2: one HTML file, no Python, no server

**`dist/invoice-to-po.html` is the whole tool.** Double-click it and it opens in your browser.
It runs offline with no AI, and nothing leaves the computer. You can also put it anywhere that
serves files: a shared drive, GitHub Pages, Netlify, the shop's website.

It does what v1 does: reads supplier invoice PDFs (or CSV exports), matches lines to your Retail
Express products using each supplier's downloaded template, lets you correct the table, and
downloads the filled PO sheet.

## Using it

1. **Add each supplier's template once.** In Retail Express: **Get Excel Template** →
   *1 Supplier, for 1 Outlet* → outlet + supplier → .xlsx. Then click **Add template…** on the page.
2. **Drop invoices** onto the page.
3. The first time a supplier's invoice comes in, pick their template from the **Supplier** menu.
   From then on, the page picks it for you.
4. Check the table, fix anything, then click **Download**. Upload the file in Retail Express with **Upload Excel**.

Templates and remembered suppliers are saved **in this browser on this computer**. Another
computer, another browser, or clearing browser data starts empty, so add the templates again.

## Changing it

The source is in `src/`: `reader.js` reads PDFs and CSVs, `po.js` fills the sheet, `app.js` runs
the page, and `index.html` holds the layout and styles. `vendor/` holds pdf.js 3.11.174 and
ExcelJS 4.4.0.

    node v2/build.mjs                 # rebuild dist/invoice-to-po.html (needs Node, only to build)

**Self-test.** Serve the repo folder over http (`python3 -m http.server`, `npx serve`, …) and open
`/v2/test/selftest.html`. All four test invoices and the CSV must read exactly right. To try the
page itself the same way, open `/v2/src/index.html`.

**When a supplier's column names aren't recognised**, add them to `HEADER_WORDS` at the top of
`src/reader.js`, then rebuild.
