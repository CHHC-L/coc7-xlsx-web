// Bundle the web app into one self-contained HTML file (works offline from file://).
import { build } from 'esbuild';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join, basename } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
// built-in blank cards, in the order offered when creating a card; the first is the default
const templates = [
  { id: 'CY23.2', file: join(root, 'COC7空白卡CY23.2.xlsx') },
  { id: 'CY23Final', file: join(root, 'COC7空白卡CY23Final.xlsx') },
];
const out = process.argv[2] || join(root, 'dist', 'COC7车卡.html');

const res = await build({
  entryPoints: [join(root, 'src/ui/main.js')],
  bundle: true,
  format: 'iife',
  target: ['safari15', 'chrome100'],
  minify: true,
  write: false,
  legalComments: 'none',
});
const js = res.outputFiles[0].text.replace(/<\/script/gi, '<\\/script');
const css = readFileSync(join(root, 'src/ui/style.css'), 'utf8');
const tpls = templates.map((t) =>
  `<script class="xlsx-template" type="application/octet-stream" data-id="${t.id}" data-name="${basename(t.file).replace(/"/g, '')}">${readFileSync(t.file).toString('base64')}</script>`).join('\n');
const html = readFileSync(join(root, 'src/index.html'), 'utf8')
  .replace('/*__CSS__*/', () => css)
  .replace('<!--__TEMPLATES__-->', () => tpls)
  .replace('/*__JS__*/', () => js);
mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, html);
console.log(`wrote ${out} (${(html.length / 1024).toFixed(0)} KB; js ${(js.length / 1024).toFixed(0)} KB)`);
