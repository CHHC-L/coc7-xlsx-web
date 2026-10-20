// Evaluate a JS expression in the built page (WebKit by default):
//   node tools/peval.mjs '<expression>' [sheet] [engine]
import { webkit, chromium } from 'playwright';
import { join } from 'node:path';

const [expr, sheet, engine = 'webkit'] = process.argv.slice(2);
const browser = await (engine === 'chrome' ? chromium.launch({ channel: 'chrome' }) : webkit.launch());
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
page.on('pageerror', (e) => console.log('pageerror', e.message));
await page.goto('file://' + join(process.cwd(), 'dist', 'COC7车卡.html'));
await page.waitForSelector('.tab.on');
if (sheet) await page.click(`.tab:has-text("${sheet}")`);
const r = await page.evaluate(`(async () => (${expr}))()`);
console.log(typeof r === 'string' ? r : JSON.stringify(r, null, 1));
await browser.close();
