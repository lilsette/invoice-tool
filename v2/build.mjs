// Bundle v2/src/index.html and every script it loads into ONE self-contained HTML file:
//
//     node v2/build.mjs        →  v2/dist/invoice-to-po.html
//
// That file is the whole tool: double-click it, email it, or put it on any web host.
// It needs no server and no internet (pdf.js and ExcelJS are inlined from v2/vendor/).
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
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
