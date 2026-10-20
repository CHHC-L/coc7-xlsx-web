// Data validation and conditional formatting, evaluated like Excel does.

import { parseFormula, shiftAst } from './formula.js';
import { inRange } from './refs.js';
import { Ref, XErr, isErr, compareValues, toNumber, toText } from './values.js';
import { formatValue } from './numfmt.js';

const astCache = new WeakMap();

function compiled(owner, idx, text) {
  let m = astCache.get(owner);
  if (!m) { m = new Map(); astCache.set(owner, m); }
  if (!m.has(idx)) {
    let ast = null;
    try { ast = text === null || text === undefined ? null : parseFormula(text); } catch (e) { ast = null; }
    m.set(idx, ast);
  }
  return m.get(idx);
}

// Relative references in DV/CF formulas are relative to the top-left cell of
// the rule's first range.
function evalRel(doc, ast, si, anchor, r, c, opts) {
  const shifted = anchor && (r !== anchor.r1 || c !== anchor.c1) ? shiftAst(ast, r - anchor.r1, c - anchor.c1) : ast;
  return doc.eng.evalAt(shifted, si, r, c, opts);
}

// ---- data validation ----------------------------------------------------------------

export function findDV(doc, si, r, c) {
  const sh = doc.wb.sheets[si];
  for (const dv of sh.dvs) {
    for (const g of dv.ranges) if (inRange(g, r, c)) return dv;
  }
  return null;
}

export function displayText(doc, v, s) {
  if (v === null || v === undefined) return '';
  const code = doc.numFmtCode(s);
  return formatValue(v, code, doc.wb.styles.indexed).text;
}

// List items for a list validation: [{text, value}] or null when the source is invalid.
export function listItems(doc, si, dv, r, c) {
  if (dv.type !== 'list' || dv.formula1 === null) return null;
  const f = dv.formula1.trim();
  if (f.startsWith('"') && f.endsWith('"')) {
    return f.slice(1, -1).replace(/""/g, '"').split(',').map((t) => ({ text: t, value: t }));
  }
  const ast = compiled(dv, 1, f);
  if (!ast) return null;
  const v = evalRel(doc, ast, si, dv.ranges[0], r, c, { keepRef: true });
  if (v instanceof Ref) {
    if (v.rows > 1 && v.cols > 1) return null;
    const out = [];
    for (let rr = v.r1; rr <= v.r2; rr++) {
      for (let cc = v.c1; cc <= v.c2; cc++) {
        const val = doc.eng.getValue(v.sheet, rr, cc);
        const s = doc.styleOf(v.sheet, rr, cc);
        out.push({ text: isErr(val) ? val.code : displayText(doc, val, s), value: val });
      }
    }
    return out;
  }
  if (isErr(v)) return null;
  if (typeof v === 'string') return v.split(',').map((t) => ({ text: t, value: t }));
  return [{ text: toText(v), value: v }];
}

function bound(doc, si, dv, idx, text, r, c) {
  if (text === null || text === undefined) return null;
  const ast = compiled(dv, idx, text);
  if (!ast) return null;
  const v = evalRel(doc, ast, si, dv.ranges[0], r, c);
  if (isErr(v)) return v;
  const n = toNumber(v);
  return isErr(n) ? n : n;
}

function opOk(op, x, a, b) {
  switch (op) {
    case 'between': return x >= a && x <= b;
    case 'notBetween': return x < a || x > b;
    case 'equal': return x === a;
    case 'notEqual': return x !== a;
    case 'greaterThan': return x > a;
    case 'lessThan': return x < a;
    case 'greaterThanOrEqual': return x >= a;
    case 'lessThanOrEqual': return x <= a;
    default: return true;
  }
}

// Validate a candidate value for a cell. Returns true if allowed.
export function validateValue(doc, si, r, c, dv, value) {
  if (!dv || !dv.showErrorMessage) return true;
  if (value === null || value === '') return dv.allowBlank || true;
  switch (dv.type) {
    case 'none': return true;
    case 'list': {
      const items = listItems(doc, si, dv, r, c);
      if (!items) return false;
      const s = doc.styleOf(si, r, c);
      const asText = typeof value === 'string' ? value : displayText(doc, value, s);
      for (const it of items) {
        if (typeof it.value === 'number' && typeof value === 'number' && it.value === value) return true;
        if (it.text === asText) return true;
        if (typeof value === 'string' && typeof it.value === 'string' && it.value === value) return true;
        if (typeof value === 'number' && it.text === toText(value)) return true;
      }
      return false;
    }
    case 'whole': case 'decimal': case 'date': case 'time': {
      if (typeof value !== 'number') return false;
      if (dv.type === 'whole' && !Number.isInteger(value)) return false;
      const a = bound(doc, si, dv, 1, dv.formula1, r, c);
      const b = bound(doc, si, dv, 2, dv.formula2, r, c);
      if (isErr(a) || isErr(b)) return false;
      if (a === null) return true;
      return opOk(dv.operator, value, a, b === null ? a : b);
    }
    case 'textLength': {
      const len = (typeof value === 'string' ? value : toText(value)).length;
      const a = bound(doc, si, dv, 1, dv.formula1, r, c);
      const b = bound(doc, si, dv, 2, dv.formula2, r, c);
      if (isErr(a) || isErr(b) || a === null) return false;
      return opOk(dv.operator, len, a, b === null ? a : b);
    }
    case 'custom': {
      const ast = compiled(dv, 1, dv.formula1);
      if (!ast) return false;
      const cell = doc.cell(si, r, c);
      const saved = cell ? { v: cell.value, ast: cell.ast } : null;
      const tmp = doc.ensureCell(si, r, c);
      const prev = { v: tmp.value, ast: tmp.ast };
      tmp.value = value; tmp.ast = null;
      try {
        const res = evalRel(doc, ast, si, dv.ranges[0], r, c);
        return res === true || (typeof res === 'number' && res !== 0);
      } finally {
        tmp.value = prev.v; tmp.ast = prev.ast;
        if (!saved && tmp.isNew && tmp.value === null) doc.wb.sheets[si].cells.delete((r - 1) * 16384 + (c - 1));
      }
    }
    default: return true;
  }
}

// ---- conditional formatting ---------------------------------------------------------

function ruleMatches(doc, si, rule, r, c, value) {
  const anchor = rule.ranges[0];
  const f = (i) => {
    const ast = rule.asts[i];
    if (!ast) return null;
    return evalRel(doc, ast, si, anchor, r, c);
  };
  switch (rule.type) {
    case 'cellIs': {
      if (isErr(value)) return false;
      const a = f(0);
      if (isErr(a)) return false;
      const cmp = compareValues(value, a);
      switch (rule.operator) {
        case 'equal': return cmp === 0;
        case 'notEqual': return cmp !== 0;
        case 'greaterThan': return cmp > 0;
        case 'lessThan': return cmp < 0;
        case 'greaterThanOrEqual': return cmp >= 0;
        case 'lessThanOrEqual': return cmp <= 0;
        case 'between': case 'notBetween': {
          const b = f(1);
          if (isErr(b)) return false;
          const lo = compareValues(a, b) <= 0 ? a : b, hi = lo === a ? b : a;
          const inside = compareValues(value, lo) >= 0 && compareValues(value, hi) <= 0;
          return rule.operator === 'between' ? inside : !inside;
        }
        default: return false;
      }
    }
    case 'expression': {
      const v = f(0);
      return v === true || (typeof v === 'number' && v !== 0);
    }
    case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
      const v = f(0);
      return v === true || (typeof v === 'number' && v !== 0);
    }
    case 'containsBlanks': return value === null || (typeof value === 'string' && value.trim() === '');
    case 'notContainsBlanks': return !(value === null || (typeof value === 'string' && value.trim() === ''));
    case 'containsErrors': return value instanceof XErr;
    case 'notContainsErrors': return !(value instanceof XErr);
    default: return false;
  }
}

// Merged differential format for a cell, or null.
export function cfFormat(doc, si, r, c, value) {
  const sh = doc.wb.sheets[si];
  if (!sh.cfs.length) return null;
  let out = null;
  for (const rule of sh.cfs) {
    let hit = false;
    for (const g of rule.ranges) if (inRange(g, r, c)) { hit = true; break; }
    if (!hit) continue;
    if (!ruleMatches(doc, si, rule, r, c, value)) continue;
    const dxf = rule.dxfId !== null ? doc.wb.styles.dxfs[rule.dxfId] : null;
    if (dxf) {
      if (!out) out = {};
      // higher priority (earlier) wins per property
      if (dxf.font) {
        out.font = out.font || {};
        for (const [k, v] of Object.entries(dxf.font)) if (out.font[k] === undefined) out.font[k] = v;
      }
      if (dxf.fill && !out.fill) out.fill = dxf.fill;
      if (dxf.border && !out.border) out.border = dxf.border;
      if (dxf.numFmt && !out.numFmt) out.numFmt = dxf.numFmt;
    }
    if (rule.stopIfTrue) break;
  }
  return out;
}

export function cfCells(doc, si) {
  // all cells covered by some CF rule (for refresh after recalculation)
  const sh = doc.wb.sheets[si];
  const out = [];
  const seen = new Set();
  for (const rule of sh.cfs) {
    for (const g of rule.ranges) {
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        const k = r * 20000 + c;
        if (!seen.has(k)) { seen.add(k); out.push([r, c]); }
      }
    }
  }
  return out;
}
