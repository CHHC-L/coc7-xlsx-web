// Calculation engine: owns cell values of a parsed workbook and recalculates
// formulas on demand, ordered by Excel's calcChain to keep recursion shallow.

import { evaluate, scalar, MISSING } from './evaluator.js';
import { Ref, Arr, ERR, isErr } from './values.js';
import { key, parseAddr } from './refs.js';
import { parseXml } from './xml.js';
import { formatValue } from './numfmt.js';

export class Engine {
  constructor(wb) {
    this.wb = wb;
    this.pass = 0;
    this.calculating = false;
    this.circular = [];
    this.byName = new Map(wb.sheets.map((s, i) => [s.name.toLowerCase(), i]));
    this.names = new Map();
    for (const n of wb.names) {
      const k = n.name.toLowerCase();
      if (!this.names.has(k)) this.names.set(k, []);
      this.names.get(k).push(n);
    }
    this.buildOrder();
  }

  buildOrder() {
    // calcChain order first, then any remaining formula cells
    const order = [];
    const seen = new Set();
    const wb = this.wb;
    if (wb.calcChainPath) {
      const txt = wb.zip.text(wb.calcChainPath);
      if (txt) {
        const root = parseXml(txt);
        let si = -1;
        for (const c of root.all('c')) {
          if (c.attrs.i !== undefined) {
            const sid = +c.attrs.i;
            si = wb.sheets.findIndex((s) => s.sheetId === sid);
          }
          if (si < 0) continue;
          const a = parseAddr(c.attrs.r);
          if (!a) continue;
          const cell = wb.sheets[si].cells.get(key(a.r, a.c));
          if (cell && cell.ast && !seen.has(cell)) { seen.add(cell); order.push([si, cell]); }
        }
      }
    }
    wb.sheets.forEach((sh, si) => {
      for (const cell of sh.cells.values()) {
        if (cell.ast && !seen.has(cell)) { seen.add(cell); order.push([si, cell]); }
      }
    });
    this.order = order;
  }

  sheet(i) { return this.wb.sheets[i]; }

  sheetIndex(name) {
    const i = this.byName.get(name.toLowerCase());
    return i === undefined ? -1 : i;
  }

  getCell(si, r, c) {
    const sh = this.wb.sheets[si];
    return sh ? sh.cells.get(key(r, c)) || null : null;
  }

  getValue(si, r, c) {
    const sh = this.wb.sheets[si];
    if (!sh) return ERR.REF;
    const cell = sh.cells.get(key(r, c));
    if (!cell) return null;
    if (cell.ast && this.calculating && cell.pass !== this.pass) this.compute(si, cell);
    return cell.value;
  }

  evalName(name, sheetName, ctx) {
    const list = this.names.get(name.toLowerCase());
    if (!list) return ERR.NAME;
    let si = ctx.sheet;
    if (sheetName) si = this.sheetIndex(sheetName);
    const n = list.find((x) => x.localSheetId === si) || list.find((x) => x.localSheetId === null);
    if (!n || !n.ast) return ERR.NAME;
    return evaluate(n.ast, { ...ctx, sheet: si });
  }

  forEachInRef(ref, cb) {
    const sh = this.wb.sheets[ref.sheet];
    if (!sh) return;
    const r2 = Math.min(ref.r2, sh.maxRow), c2 = Math.min(ref.c2, sh.maxCol);
    const area = (r2 - ref.r1 + 1) * (c2 - ref.c1 + 1);
    if (area <= 0) return;
    if (area > sh.cells.size * 2) {
      const cells = [];
      for (const cell of sh.cells.values()) {
        if (cell.r >= ref.r1 && cell.r <= r2 && cell.c >= ref.c1 && cell.c <= c2) cells.push(cell);
      }
      cells.sort((a, b) => a.r - b.r || a.c - b.c);
      for (const cell of cells) if (cb(this.getValue(ref.sheet, cell.r, cell.c))) return;
      return;
    }
    for (let r = ref.r1; r <= r2; r++) {
      for (let c = ref.c1; c <= c2; c++) {
        const v = this.getValue(ref.sheet, r, c);
        if (v !== null && cb(v)) return;
      }
    }
  }

  // Final value stored in a formula cell.
  finalize(v, ctx) {
    if (v === MISSING) return 0;
    if (v instanceof Ref || v instanceof Arr) v = scalar(v, ctx);
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number' && (!isFinite(v) || Number.isNaN(v))) return ERR.NUM;
    return v;
  }

  compute(si, cell) {
    if (cell.computing) {
      this.circular.push([si, cell.r, cell.c]);
      return;
    }
    cell.computing = true;
    const ctx = { eng: this, sheet: si, r: cell.r, c: cell.c };
    let v;
    try {
      v = this.finalize(evaluate(cell.ast, ctx), ctx);
    } catch (e) {
      if (e instanceof RangeError) throw e;
      v = ERR.VALUE;
      cell.evalError = e;
    }
    cell.computing = false;
    cell.pass = this.pass;
    cell.value = v;
  }

  recalc() {
    this.pass++;
    this.circular = [];
    this.calculating = true;
    try {
      for (const [si, cell] of this.order) {
        if (cell.ast && cell.pass !== this.pass) this.compute(si, cell);
      }
    } finally {
      this.calculating = false;
    }
  }

  // Evaluate an arbitrary formula AST (data validation, conditional formats).
  evalAt(ast, si, r, c, { keepRef = false } = {}) {
    const was = this.calculating;
    this.calculating = true;
    const ctx = { eng: this, sheet: si, r, c };
    try {
      const v = evaluate(ast, ctx);
      if (keepRef) return v;
      return v instanceof Ref || v instanceof Arr ? scalar(v, ctx) : v === MISSING ? null : v;
    } finally {
      this.calculating = was;
    }
  }

  formatText(v, code) {
    return formatValue(v, code, this.wb.styles.indexed).text;
  }
}

export { isErr };
