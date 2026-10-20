// Recalculate a workbook and compare every formula result with the cached
// value Excel stored in the file.
import { readFileSync } from 'node:fs';
import { loadWorkbook } from '../src/core/workbook.js';
import { Engine } from '../src/core/calc.js';
import { XErr } from '../src/core/values.js';
import { addr } from '../src/core/refs.js';

const file = process.argv[2] || 'COC7空白卡CY23.2.xlsx';
const t0 = performance.now();
const wb = loadWorkbook(readFileSync(file));
const t1 = performance.now();
const cached = new Map();
let nf = 0;
wb.sheets.forEach((sh, si) => { for (const c of sh.cells.values()) if (c.ast) { cached.set(c, c.value); nf++; } });
const eng = new Engine(wb);
eng.recalc();
const t2 = performance.now();
eng.recalc();
const t3 = performance.now();
const show = (v) => v instanceof XErr ? v.code : JSON.stringify(v);
let bad = 0;
for (const [c, old] of cached) {
  const nv = c.value;
  const same = (old instanceof XErr && nv instanceof XErr) ? old.code === nv.code
    : (typeof old === 'number' && typeof nv === 'number') ? Math.abs(old - nv) <= 1e-9 * Math.max(1, Math.abs(old))
    : old === nv;
  if (!same && !c.volatile) {
    bad++;
    const sh = wb.sheets.find((s) => s.cells.get((c.r - 1) * 16384 + c.c - 1) === c);
    if (bad <= 40) console.log(`${sh.name}!${addr(c.r, c.c)}  excel=${show(old)}  ours=${show(nv)}  f=${(c.formula || '(shared)').slice(0, 120)}${c.evalError ? ' ERR ' + c.evalError.message : ''}`);
  }
}
console.log(`formulas=${nf} mismatches=${bad} circular=${eng.circular.length} load=${(t1 - t0).toFixed(0)}ms recalc1=${(t2 - t1).toFixed(0)}ms recalc2=${(t3 - t2).toFixed(0)}ms`);
const parseErrs = [];
wb.sheets.forEach((sh) => { for (const c of sh.cells.values()) if (c.ast && c.ast.parseError) parseErrs.push(sh.name + '!' + addr(c.r, c.c) + ' ' + c.ast.parseError); });
if (parseErrs.length) console.log('parse errors:', parseErrs.slice(0, 10));
