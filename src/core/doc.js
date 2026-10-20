// Document = parsed workbook + calculation engine + the user's edits.

import { loadWorkbook, Cell } from './workbook.js';
import { Engine } from './calc.js';
import { key, addr, parseAddr, inRange } from './refs.js';
import { parseFormula } from './formula.js';
import { XErr, errFromCode } from './values.js';

export class Doc {
  constructor(bytes) {
    this.bytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    this.wb = loadWorkbook(this.bytes);
    this.eng = new Engine(this.wb);
    this.seq = 0;
    this.styleAdds = []; // xfs appended for auto number formats: {base, numFmtId, quotePrefix}
    this.unprotected = new Set(); // sheet indices whose protection the user removed
    this.eng.recalc();
  }

  get sheets() { return this.wb.sheets; }

  cell(si, r, c) {
    return this.wb.sheets[si].cells.get(key(r, c)) || null;
  }

  ensureCell(si, r, c) {
    const sh = this.wb.sheets[si];
    let cell = sh.cells.get(key(r, c));
    if (!cell) {
      cell = new Cell(r, c);
      cell.s = this.defaultStyle(si, r, c);
      cell.isNew = true;
      sh.cells.set(key(r, c), cell);
      if (r > sh.maxRow) sh.maxRow = r;
      if (c > sh.maxCol) sh.maxCol = c;
    }
    return cell;
  }

  // style of a cell that has no <c> element: row style (customFormat) or column style
  defaultStyle(si, r, c) {
    const sh = this.wb.sheets[si];
    const row = sh.rows.get(r);
    if (row && row.customFormat && row.s !== null) return row.s;
    for (const col of sh.cols) if (c >= col.min && c <= col.max) return col.style || 0;
    return 0;
  }

  styleOf(si, r, c) {
    const cell = this.cell(si, r, c);
    return cell ? cell.s : this.defaultStyle(si, r, c);
  }

  xf(s) { return this.wb.styles.cellXfs[s] || this.wb.styles.cellXfs[0]; }

  numFmtCode(s) {
    const xf = this.xf(s);
    return this.wb.styles.numFmtCode(xf ? xf.numFmtId : 0);
  }

  // Excel protection semantics
  isProtected(si) {
    const p = this.wb.sheets[si].protection;
    return !!(p && p.sheet) && !this.unprotected.has(si);
  }

  hasProtection(si) {
    const p = this.wb.sheets[si].protection;
    return !!(p && p.sheet);
  }

  setUnprotected(si, off) {
    if (off) this.unprotected.add(si); else this.unprotected.delete(si);
  }

  protectionState() {
    return [...this.unprotected].map((si) => this.wb.sheets[si].name);
  }

  applyProtectionState(names) {
    for (const n of names || []) {
      const si = this.wb.sheets.findIndex((s) => s.name === n);
      if (si >= 0 && this.hasProtection(si)) this.unprotected.add(si);
    }
  }

  isLocked(si, r, c) {
    const xf = this.xf(this.styleOf(si, r, c));
    return xf ? xf.locked : true;
  }

  inAllowedRange(si, r, c) {
    for (const pr of this.wb.sheets[si].protectedRanges) {
      if (pr.hasPassword) continue;
      for (const g of pr.ranges) if (inRange(g, r, c)) return true;
    }
    return false;
  }

  canEdit(si, r, c) {
    if (!this.isProtected(si)) return true;
    return !this.isLocked(si, r, c) || this.inAllowedRange(si, r, c);
  }

  canSelect(si, r, c) {
    const sh = this.wb.sheets[si];
    if (!this.isProtected(si)) return true;
    const p = sh.protection;
    const locked = this.isLocked(si, r, c);
    if (locked && p.selectLockedCells) return this.inAllowedRange(si, r, c) ? true : false;
    if (!locked && p.selectUnlockedCells) return false;
    return true;
  }

  // --- snapshots for undo ---------------------------------------------------------
  snapshot(si, r, c) {
    const cell = this.cell(si, r, c);
    if (!cell) return { si, r, c, exists: false };
    const keys = ['s', 't', 'value', 'formula', 'ast', 'fattrs', 'sharedFrom', 'sst', 'runs', 'edited', 'strSeq', 'volatile', 'isNew', 'quotePrefix', 'arrayRef'];
    const o = { si, r, c, exists: true };
    for (const k of keys) o[k] = cell[k];
    return o;
  }

  restore(snap) {
    const sh = this.wb.sheets[snap.si];
    if (!snap.exists) {
      const cell = sh.cells.get(key(snap.r, snap.c));
      if (cell && cell.isNew) sh.cells.delete(key(snap.r, snap.c));
      return;
    }
    const cell = this.ensureCell(snap.si, snap.r, snap.c);
    for (const k of Object.keys(snap)) if (!['si', 'r', 'c', 'exists'].includes(k)) cell[k] = snap[k];
    if (cell.ast) this.eng.order.push([snap.si, cell]);
  }

  // --- edits ---------------------------------------------------------------------
  // content: {kind:'clear'} | {kind:'value', value, numFmtId?, quotePrefix?} | {kind:'formula', formula}
  setContent(si, r, c, content) {
    const cell = this.ensureCell(si, r, c);
    cell.edited = true;
    if (content.kind === 'formula') {
      let ast;
      try { ast = parseFormula(content.formula); } catch (e) { return { ok: false, error: e.message }; }
      cell.formula = content.formula;
      cell.ast = ast;
      cell.fattrs = null;
      cell.sharedFrom = null;
      cell.sst = -1;
      cell.runs = null;
      cell.volatile = /RAND|NOW|TODAY|OFFSET|INDIRECT/i.test(content.formula);
      this.eng.order.push([si, cell]);
    } else {
      cell.formula = null;
      cell.ast = null;
      cell.fattrs = null;
      cell.sharedFrom = null;
      cell.runs = null;
      cell.sst = -1;
      cell.volatile = false;
      if (content.kind === 'clear') {
        cell.value = null;
      } else {
        cell.value = content.value;
        if (typeof content.value === 'string') cell.strSeq = ++this.seq;
        if (content.numFmtId !== undefined || content.quotePrefix) {
          cell.s = this.styleWith(cell.s, content.numFmtId, content.quotePrefix);
        }
      }
    }
    return { ok: true };
  }

  // find or register an xf identical to `s` but with another number format / quotePrefix
  styleWith(s, numFmtId, quotePrefix) {
    const xfs = this.wb.styles.cellXfs;
    const base = xfs[s];
    const want = { ...base, numFmtId: numFmtId !== undefined ? numFmtId : base.numFmtId, quotePrefix: !!quotePrefix || base.quotePrefix };
    const same = (a, b) => a.numFmtId === b.numFmtId && a.fontId === b.fontId && a.fillId === b.fillId && a.borderId === b.borderId &&
      a.xfId === b.xfId && a.quotePrefix === b.quotePrefix && a.locked === b.locked && a.hidden === b.hidden &&
      JSON.stringify(a.alignment) === JSON.stringify(b.alignment);
    for (let i = 0; i < xfs.length; i++) if (same(xfs[i], want)) return i;
    xfs.push(want);
    this.styleAdds.push({ base: s, index: xfs.length - 1, numFmtId: want.numFmtId, quotePrefix: want.quotePrefix });
    return xfs.length - 1;
  }

  recalc() { this.eng.recalc(); }

  // --- persistence of edits ---------------------------------------------------------
  editList() {
    const out = [];
    this.wb.sheets.forEach((sh, si) => {
      for (const cell of sh.cells.values()) {
        if (!cell.edited) continue;
        const e = { sheet: sh.name, a: addr(cell.r, cell.c), seq: cell.strSeq || 0 };
        if (cell.ast && cell.formula !== null && !cell.fattrs) e.formula = cell.formula;
        else if (cell.value === null) e.clear = true;
        else if (cell.value instanceof XErr) e.err = cell.value.code;
        else e.value = cell.value;
        const xf = this.wb.styles.cellXfs[cell.s];
        if (xf) { e.numFmtId = xf.numFmtId; e.quotePrefix = xf.quotePrefix; }
        out.push(e);
      }
    });
    out.sort((a, b) => a.seq - b.seq);
    return out;
  }

  applyEditList(list) {
    for (const e of list) {
      const si = this.wb.sheets.findIndex((s) => s.name === e.sheet);
      if (si < 0) continue;
      const a = parseAddr(e.a);
      if (!a) continue;
      let content;
      if (e.formula !== undefined) content = { kind: 'formula', formula: e.formula };
      else if (e.clear) content = { kind: 'clear' };
      else if (e.err) content = { kind: 'value', value: errFromCode(e.err) };
      else {
        content = { kind: 'value', value: e.value };
        const cur = this.styleOf(si, a.r, a.c);
        const xf = this.wb.styles.cellXfs[cur];
        if (xf && e.numFmtId !== undefined && e.numFmtId !== xf.numFmtId) content.numFmtId = e.numFmtId;
        if (e.quotePrefix && xf && !xf.quotePrefix) content.quotePrefix = true;
      }
      this.setContent(si, a.r, a.c, content);
    }
    this.recalc();
  }
}
