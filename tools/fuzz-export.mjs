// Generate randomised but realistic character sheets, export them, so another
// engine can recompute and compare. Usage: node tools/fuzz-export.mjs <outdir> <count> [seed]
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { Doc } from '../src/core/doc.js';
import { exportXlsx } from '../src/core/export.js';
import { interpretInput } from '../src/core/input.js';
import { findDV, listItems } from '../src/core/rules.js';
import { parseAddr, addr } from '../src/core/refs.js';

const [outDir, countS = '5', seedS = '1'] = process.argv.slice(2);
mkdirSync(outDir, { recursive: true });
let seed = +seedS;
const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
const ri = (a, b) => a + Math.floor(rnd() * (b - a + 1));
const pick = (arr) => arr[Math.floor(rnd() * arr.length)];

const SKILL_ROWS = [];
for (let r = 16; r <= 49; r++) SKILL_ROWS.push(r);

for (let n = 0; n < +countS; n++) {
  const doc = new Doc(readFileSync('COC7空白卡CY23.2.xlsx'));
  const si = 0;
  const log = [];
  const set = (a, text) => {
    const { r, c } = parseAddr(a);
    if (!doc.canEdit(si, r, c)) return;
    const content = interpretInput(String(text), doc.numFmtCode(doc.styleOf(si, r, c)));
    doc.setContent(si, r, c, content);
    log.push([a, text]);
  };
  set('E3', pick(['张三', 'John Smith', '玛丽']));
  set('M5', String(ri(0, 230)));
  set('E6', String(ri(15, 89)));
  for (const a of ['U3', 'U5', 'U7', 'AA3', 'AA5', 'AA7', 'AG3', 'AG5', 'AG7']) set(a, String(ri(3, 18) * 5));
  // list-validated cells on the sheet: choose options at random
  const sh = doc.wb.sheets[si];
  for (const dv of sh.dvs) {
    if (dv.type !== 'list') continue;
    for (const g of dv.ranges) {
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        const ar = sh.mergeAt.get((r - 1) * 16384 + c - 1);
        if (ar && (ar.r1 !== r || ar.c1 !== c)) continue;
        if (!doc.canEdit(si, r, c) || rnd() < 0.5) continue;
        const items = (listItems(doc, si, dv, r, c) || []).filter((x) => x.text !== '');
        if (items.length) set(addr(r, c), pick(items).text);
      }
    }
  }
  doc.recalc();
  // skill points: occupation / interest / growth columns
  for (const r of SKILL_ROWS) {
    for (const col of ['L', 'N', 'P', 'AH', 'AJ', 'AL']) if (rnd() < 0.25) set(col + r, String(ri(1, 40)));
  }
  set('I12', String(ri(0, 5)));
  set('N12', String(ri(0, 8)));
  doc.recalc();
  const res = exportXlsx(doc, { modified: new Date('2026-10-19T00:00:00Z') });
  const file = join(outDir, `fuzz-${n}.xlsx`);
  writeFileSync(file, res.bytes);
  writeFileSync(file + '.inputs.json', JSON.stringify(log));
  console.log(file, 'inputs', log.length, 'changed cells', res.changedCells.length, 'circular', doc.eng.circular.length);
}
