// Drive the built page in WebKit (Safari's engine): render, type a few cells,
// export, and save screenshots + the exported file for inspection.
import { webkit, chromium } from 'playwright';
import { writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';

const outDir = process.argv[2] || 'scratch';
const engine = process.argv[3] || 'webkit';
mkdirSync(outDir, { recursive: true });
const html = 'file://' + join(process.cwd(), 'dist', 'COC7车卡.html');
const browser = await (engine === 'chrome' ? chromium.launch({ channel: 'chrome' }) : webkit.launch());
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2, acceptDownloads: true });
const page = await ctx.newPage();
const logs = [];
page.on('console', (m) => logs.push(`[${m.type()}] ${m.text()}`));
page.on('pageerror', (e) => logs.push(`[pageerror] ${e.message}\n${e.stack}`));
const t0 = Date.now();
await page.goto(html);
await page.waitForSelector('.tab.on', { timeout: 20000 });
console.log('boot ms', Date.now() - t0);
await page.screenshot({ path: join(outDir, '01-initial.png') });

// type into a few cells using the keyboard like a user would
async function typeAt(a, text) {
  await page.evaluate((a) => {
    const app = window.cocApp;
    const m = /^([A-Z]+)(\d+)$/.exec(a);
    let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
    app.grid.select(+m[2], c, true);
    app.grid.focus();
  }, a);
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
  await page.waitForTimeout(30);
}
const steps = [['E3', '约翰·史密斯'], ['E4', '测试玩家'], ['M5', '2'], ['E6', '25'], ['U3', '50'], ['U5', '60'], ['U7', '55'], ['AA3', '70'], ['AA5', '45'], ['AA7', '80'], ['AG3', '65'], ['AG5', '75'], ['AG7', '50'], ['N16', '40']];
for (const [a, t] of steps) await typeAt(a, t);
await page.screenshot({ path: join(outDir, '02-filled.png') });

// invalid value -> validation dialog (STR must be 1..99, warning style)
await typeAt('U3', '150');
await page.waitForTimeout(100);
const dlg = await page.$('.dlg');
console.log('validation dialog shown:', !!dlg, dlg ? await dlg.innerText() : '');
await page.screenshot({ path: join(outDir, '03-validation.png') });
if (dlg) await page.click('.dlg button:has-text("否")');
await page.keyboard.press('Escape');

// dropdown on B16 (☐/☑)
await page.evaluate(() => { const g = window.cocApp.grid; g.select(16, 2, true); g.focus(); });
await page.waitForTimeout(50);
await page.screenshot({ path: join(outDir, '04-dropdown-btn.png') });
await page.keyboard.press('Alt+ArrowDown');
await page.waitForTimeout(50);
await page.screenshot({ path: join(outDir, '05-dropdown-open.png') });
await page.keyboard.press('ArrowDown');
await page.keyboard.press('Enter');

const vals = await page.evaluate(() => {
  const d = window.cocApp.doc;
  const v = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return d.eng.getValue(0, +m[2], c); };
  return { E3: v('E3'), U3: v('U3'), W3: v('W3'), AG2: v('AG2'), B13: v('B13'), B16: v('B16'), R16: v('R16') };
});
console.log(JSON.stringify(vals));

// export via the dialog
await page.click('[data-act="export"]');
await page.waitForSelector('.dlg-form');
const [dl] = await Promise.all([page.waitForEvent('download'), page.click('.dlg-form button.primary')]);
const p = join(outDir, 'exported.xlsx');
await dl.saveAs(p);
console.log('exported to', p, 'suggested name', dl.suggestedFilename());

for (const name of ['简化卡 骰娘导入', '职业列表', '成长表（测试）', '武器列表 战斗', '更新说明']) {
  await page.click(`.tab:has-text("${name}")`);
  await page.waitForTimeout(150);
  await page.screenshot({ path: join(outDir, `10-${name}.png`) });
}
writeFileSync(join(outDir, 'console.log'), logs.join('\n'));
console.log('console lines:', logs.length, logs.slice(0, 10).join('\n'));
await browser.close();
