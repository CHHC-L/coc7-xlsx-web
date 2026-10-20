// Persistence / undo / re-import / unprotect flows in WebKit with a persistent profile.
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
const val = (a, si = 0) => page.evaluate(([a, si]) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  return window.cocApp.doc.eng.getValue(si, +m[2], c);
}, [a, si]);
// protection lives in the settings dialog
const toggleProtection = async () => {
  await page.click('[data-act="settings"]');
  await page.click('.dlg-form [data-op="protect"]');
  await page.waitForTimeout(100);
};
const typeAt = async (a, text) => {
  await page.evaluate((a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; window.cocApp.grid.select(+m[2], c, true); window.cocApp.grid.focus(); }, a);
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
};

await page.goto(url);
await page.waitForSelector('.tab.on');
await typeAt('E3', '持久化测试');
await typeAt('U3', '60');
await page.waitForTimeout(700); // autosave debounce
await page.reload();
await page.waitForSelector('.tab.on');
ok(await val('E3') === '持久化测试', 'name survives reload');
ok(await val('W3') === 30, 'formula recomputed after reload (W3 = 30)');

// undo / redo
await typeAt('U3', '70');
ok(await val('W3') === 35, 'edit applied');
await page.keyboard.press('Meta+z');
ok(await val('U3') === 60, 'undo restores 60');
await page.keyboard.press('Meta+Shift+z');
ok(await val('U3') === 70, 'redo re-applies 70');

// IME-like input: insertText without key events
await page.evaluate(() => { window.cocApp.grid.select(4, 5, true); window.cocApp.grid.focus(); });
await page.keyboard.insertText('输入法文字');
await page.keyboard.press('Enter');
ok(await val('E4') === '输入法文字', 'insertText (IME path) commits');

// delete key clears
await page.evaluate(() => { window.cocApp.grid.select(4, 5, true); window.cocApp.grid.focus(); });
await page.keyboard.press('Delete');
ok(await val('E4') === null, 'Delete clears a cell');

// locked cell edit is refused
await page.evaluate(() => { window.cocApp.grid.sel = { r: 2, c: 33 }; window.cocApp.grid.focus(); });
const locked = await page.evaluate(() => window.cocApp.doc.canEdit(0, 2, 33));
ok(locked === false, 'AG2 is locked while the sheet is protected');
ok(await page.evaluate(() => window.cocApp.doc.canEdit(0, 3, 23)), 'W3 is editable (author allowed range W3:Y8)');

// unprotect 人物卡 (no password) and edit a formula cell
await toggleProtection();
ok(await page.evaluate(() => window.cocApp.doc.canEdit(0, 2, 33)), 'after unprotect AG2 is editable');
await typeAt('AG2', '99');
ok(await val('AG2') === 99, 'locked formula cell overwritten after unprotect');
await toggleProtection();
ok(!(await page.evaluate(() => window.cocApp.doc.canEdit(0, 2, 33))), 're-protect works');

// export and re-import the exported file
await page.click('[data-act="export"]');
await page.waitForSelector('.dlg-form');
const [dl] = await Promise.all([page.waitForEvent('download'), page.click('.dlg-form button.primary')]);
const file = join(outDir, 'flow-export.xlsx');
await dl.saveAs(file);
await page.setInputFiles('.filepick', file);
await page.waitForTimeout(800);
ok(await val('E3') === '持久化测试', 're-imported file shows the name');
ok(await val('AG2') === 99, 're-imported file keeps the overwritten value');
const cards = await page.evaluate(async () => (await window.cocApp.store.list()).length);
ok(cards === 2, 'two cards stored (original + imported)');
ok(errors.length === 0, 'no page errors ' + errors.join(' | '));
await browser.close();
