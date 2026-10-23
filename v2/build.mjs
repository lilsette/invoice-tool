// Bundle v2/src/index.html and every script it loads into ONE self-contained HTML file:
//
//     node v2/build.mjs        →  v2/dist/invoice-to-po.html
//                                 docs/index.html + docs/examples/   (the GitHub Pages site)
//
// That file is the whole tool: double-click it, email it, or put it on any web host.
// It needs no server and no internet (pdf.js and ExcelJS are inlined from v2/vendor/).
import { cpSync, readFileSync, writeFileSync, mkdirSync, rmSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const src = join(here, "src", "index.html");
let html = readFileSync(src, "utf8");

html = html.replace(/<script src="([^"]+)"><\/script>/g, (_, path) => {
  const code = readFileSync(resolve(dirname(src), path), "utf8")
    .replace(/<\/script/gi, "<\\/script");   // a literal </script> inside code would end the tag early
  return `<script>/* ${path.split("/").pop()} */\n${code}\n</script>`;
});

const out = join(here, "dist", "invoice-to-po.html");
mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, html);
console.log(`${out}  (${(Buffer.byteLength(html) / 1024 / 1024).toFixed(2)} MB)`);

// The GitHub Pages site: the same page as docs/index.html, plus the example invoices and
// templates it offers to load (served beside it, so only a hosted copy shows them).
const repo = resolve(here, "..");
const docs = join(repo, "docs");
const samples = join(repo, "samples", "test-invoices");
rmSync(join(docs, "examples"), { recursive: true, force: true });
mkdirSync(join(docs, "examples"), { recursive: true });
writeFileSync(join(docs, "index.html"), html);
writeFileSync(join(docs, ".nojekyll"), "");
const SUPPLIERS = [
  ["HP", "Harbour Pigments", "01-harbour-pigments", "prices ex GST, cartons, a line discount, freight, one product not in Retail Express"],
  ["GE", "Gumleaf Easels", "02-gumleaf-easels", "prices include GST, codes with leading zeros, a free bonus item"],
  ["PB", "Paperbark Paper Mills", "03-paperbark-mills", "two pages, a backordered line, codes spelt differently from Retail Express"],
  ["KB", "Kookaburra Brush Company", "04-kookaburra-brushes", "a different layout: qty first, no $ signs, table runs onto page 2"],
];
for (const [code, , file] of SUPPLIERS) {
  cpSync(join(samples, `${file}.pdf`), join(docs, "examples", `${file}.pdf`));
  cpSync(join(samples, "templates", `${code}.xlsx`), join(docs, "examples", `template-${code}.xlsx`));
}
cpSync(join(samples, "05-harbour-export.csv"), join(docs, "examples", "05-harbour-export.csv"));
writeFileSync(join(docs, "examples", "examples.json"), JSON.stringify({
  suppliers: SUPPLIERS.map(([code, name, file, about]) => ({ code, name, about, invoice: `${file}.pdf`, template: `template-${code}.xlsx` })),
  csv: "05-harbour-export.csv",
}, null, 2));
console.log(`${join(docs, "index.html")}  + examples/`);
