// Genera el PDF del informe a partir de informe.html con Chromium (Playwright).
//
// Uso:
//   node informe/build.mjs                 -> entregables/*.pdf
//   PREVIEW_DIR=/ruta node informe/build.mjs -> además, un PNG por página
//
// Falla si el contenido de alguna página desborda su área útil, para que el
// PDF nunca salga con texto recortado.
import { chromium } from 'playwright';
import { fileURLToPath, pathToFileURL } from 'node:url';
import path from 'node:path';
import fs from 'node:fs';

const here = path.dirname(fileURLToPath(import.meta.url));
const src = path.join(here, 'informe.html');
const outDir = path.resolve(here, '..', 'entregables');
const out = path.join(outDir, 'Autos_Ochoa_Reto_Final_Desarrollo_Organizacional.pdf');
const previewDir = process.env.PREVIEW_DIR;

const launchOptions = {};
if (process.env.CHROMIUM_PATH) launchOptions.executablePath = process.env.CHROMIUM_PATH;

const browser = await chromium.launch(launchOptions);
const page = await browser.newPage({
  viewport: { width: 816, height: 1056 },
  deviceScaleFactor: 2,
});
await page.goto(pathToFileURL(src).href, { waitUntil: 'networkidle' });
await page.evaluate(() => document.fonts.ready);

const overflow = await page.$$eval('.page', (pages) =>
  pages
    .map((p, i) => {
      const body = p.querySelector('.content');
      return { page: i + 1, px: body ? body.scrollHeight - body.clientHeight : 0 };
    })
    .filter((x) => x.px > 1),
);

if (previewDir) {
  fs.mkdirSync(previewDir, { recursive: true });
  const pages = await page.$$('.page');
  for (let i = 0; i < pages.length; i++) {
    await pages[i].screenshot({ path: path.join(previewDir, `p${String(i + 1).padStart(2, '0')}.png`) });
  }
}

fs.mkdirSync(outDir, { recursive: true });
await page.emulateMedia({ media: 'print' });
await page.pdf({ path: out, preferCSSPageSize: true, printBackground: true });
const count = await page.$$eval('.page', (p) => p.length);
await browser.close();

console.log(`PDF: ${path.relative(process.cwd(), out)} (${count} páginas)`);
if (overflow.length) {
  for (const o of overflow) console.error(`Desborde en página ${o.page}: ${o.px}px`);
  process.exit(1);
}
