// Version flows in WebKit: new card per version, settings shows the source, convert CY23.2 <-> CY23Final.
import { webkit } from 'playwright';
import { join } from 'node:path';

const outDir = process.argv[2];
const url = 'file://' + join(process.cwd(), 'dist', 'COC7车卡.html');
const browser = await webkit.launch();
const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 }, acceptDownloads: true });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
const ok = (cond, msg) => { console.log((cond ? 'PASS ' : 'FAIL ') + msg); if (!cond) process.exitCode = 1; };
const rc = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return [+m[2], c]; };
const val = (a) => page.evaluate(([r, c]) => window.cocApp.doc.eng.getValue(0, r, c), rc(a));
const typeAt = async (a, text) => {
  await page.evaluate(([r, c]) => { window.cocApp.grid.select(r, c, true); window.cocApp.grid.focus(); }, rc(a));
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
};
const dlgText = () => page.evaluate(() => { const d = document.querySelector('.dlg'); return d ? d.textContent.replace(/\s+/g, ' ') : null; });

await page.goto(url);
await page.waitForSelector('.tab.on');
ok(await page.evaluate(() => window.cocApp.version) === 'CY23.2', 'first start opens a CY23.2 blank card');
ok(!(await page.$('.brand')) && !(await page.$('.fx')), 'toolbar title/file name and fx label are gone');
await page.screenshot({ path: join(outDir, 'top.png'), clip: { x: 0, y: 0, width: 1280, height: 90 } });

// new CY23Final card from the cards dialog
await page.click('[data-act="cards"]');
await page.waitForSelector('.dlg-form .i-tpl');
await page.selectOption('.dlg-form .i-tpl', 'CY23Final');
await page.screenshot({ path: join(outDir, 'cards.png') });
await page.click('.dlg-form [data-op="new"]');
await page.waitForTimeout(500);
ok(await page.evaluate(() => window.cocApp.version) === 'CY23Final', 'new card uses CY23Final');
ok(await page.evaluate(() => window.cocApp.doc.wb.sheets[0].cells.get(40 * 16384 + 27).value) === '动物驯养', 'CY23Final content (AB41 动物驯养)');

// fill something on the CY23.2 card and convert it
await page.click('[data-act="cards"]');
await page.waitForSelector('.dlg-form .card');
const ids = await page.evaluate(() => [...document.querySelectorAll('.dlg-form .card')].map((c) => [c.dataset.id, c.textContent.replace(/\s+/g, ' ')]));
ok(ids.some(([, t]) => /CY23\.2/.test(t)) && ids.some(([, t]) => /CY23Final/.test(t)), 'card list shows each card\'s version');
await page.click('.dlg-form .card:not(.cur) [data-op="open"]');
await page.waitForTimeout(500);
ok(await page.evaluate(() => window.cocApp.version) === 'CY23.2', 'reopened the CY23.2 card');
await typeAt('E3', '转换测试');
await typeAt('AG5', '70');
await typeAt('M5', '106');
await typeAt('G54', '燧发枪');
const j50 = await val('J50');
await page.waitForTimeout(600);
const status = await page.evaluate(() => document.querySelector('.status').textContent);
ok(/^已保存 \d\d:\d\d$/.test(status), `status is short: "${status}"`);
await page.screenshot({ path: join(outDir, 'bottom.png'), clip: { x: 0, y: 760, width: 1280, height: 40 } });

await page.click('[data-act="settings"]');
await page.waitForSelector('.dlg-form .ver');
const st = await dlgText();
ok(/版本\s*CY23\.2\s*内置空白卡/.test(st) && /转换为 CY23Final/.test(st), 'settings shows version and convert button');
await page.screenshot({ path: join(outDir, 'settings.png') });
await page.click('.dlg-form [data-op="convert"]');
await page.waitForTimeout(200);
ok(/从 CY23\.2 转换为 CY23Final/.test(await dlgText()), 'conversion asks for confirmation');
await page.click('.dlg-btns button.primary');
await page.waitForFunction(() => window.cocApp.version === 'CY23Final' && document.querySelector('.dlg'), null, { timeout: 10000 });
const note = await dlgText();
ok(/G54/.test(note) && /燧发枪/.test(note), 'renamed weapon is reported after conversion');
await page.screenshot({ path: join(outDir, 'converted.png') });
await page.click('.dlg-btns button');
ok(await val('E3') === '转换测试' && await val('M5') === 106 && await val('J50') === j50, 'converted card keeps the inputs and results');
const n = await page.evaluate(async () => (await window.cocApp.store.list()).length);
ok(n === 3, `original kept, converted card added (${n} cards)`);

// and back
await page.click('[data-act="settings"]');
await page.waitForSelector('.dlg-form [data-op="convert"]');
ok(/转换为 CY23\.2/.test(await dlgText()), 'Final card offers conversion back to CY23.2');
await page.click('.dlg-form [data-op="convert"]');
await page.click('.dlg-btns button.primary');
await page.waitForFunction(() => window.cocApp.version === 'CY23.2', null, { timeout: 10000 });
await page.waitForTimeout(300);
ok(await val('G54') === '燧发枪' && await val('J50') === j50, 'converted back to CY23.2 with the same values');
ok(!(await page.$('.dlg')), 'no notes when going back');

// reload keeps the version
await page.waitForTimeout(600);
await page.reload();
await page.waitForSelector('.tab.on');
ok(await page.evaluate(() => window.cocApp.version) === 'CY23.2' && await val('E3') === '转换测试', 'reload restores the card and its version');

// an imported CY23.2 file: settings names the file, conversion diffs it against the blank card
await page.click('[data-act="export"]');
await page.waitForSelector('.dlg-form');
const [dl] = await Promise.all([page.waitForEvent('download'), page.click('.dlg-form button.primary')]);
const file = join(outDir, 'imported-cy232.xlsx');
await dl.saveAs(file);
await page.setInputFiles('.filepick', file);
await page.waitForFunction(() => window.cocApp.rec && window.cocApp.rec.base, null, { timeout: 10000 });
await page.click('[data-act="settings"]');
await page.waitForSelector('.dlg-form .ver');
ok(/版本\s*CY23\.2\s*导入自 imported-cy232\.xlsx/.test(await dlgText()), 'imported card: settings shows version and file');
await page.click('.dlg-form [data-op="convert"]');
await page.click('.dlg-btns button.primary');
await page.waitForFunction(() => window.cocApp.version === 'CY23Final' && document.querySelector('.dlg'), null, { timeout: 10000 });
ok(/G54/.test(await dlgText()), 'imported card converted, weapon reported');
await page.click('.dlg-btns button');
ok(await val('E3') === '转换测试' && await val('J50') === j50, 'imported card keeps inputs after conversion');
ok(errors.length === 0, 'no page errors ' + errors.join(' | '));
await browser.close();
