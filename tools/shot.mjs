// Screenshot one sheet of the built page: node tools/shot.mjs <out.png> <sheet> [scrollX] [scrollY] [engine] [width] [height]
import { webkit, chromium } from 'playwright';
import { join } from 'node:path';

const [out, sheet, sx = '0', sy = '0', engine = 'webkit', w = '1440', h = '900'] = process.argv.slice(2);
const browser = await (engine === 'chrome' ? chromium.launch({ channel: 'chrome' }) : webkit.launch());
const page = await browser.newPage({ viewport: { width: +w, height: +h }, deviceScaleFactor: 2 });
const errs = [];
page.on('pageerror', (e) => errs.push(e.message));
page.on('console', (m) => { if (m.type() === 'error') errs.push(m.text()); });
await page.goto('file://' + join(process.cwd(), 'dist', 'COC7车卡.html'));
await page.waitForSelector('.tab.on', { timeout: 20000 });
if (sheet) await page.click(`.tab:has-text("${sheet}")`);
await page.evaluate(([x, y]) => { const s = document.querySelector('.scroller'); s.scrollLeft = x; s.scrollTop = y; }, [+sx, +sy]);
await page.waitForTimeout(300);
await page.screenshot({ path: out });
if (errs.length) console.log('errors:', errs.join('\n'));
await browser.close();
