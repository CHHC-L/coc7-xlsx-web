// Moving a card's own content between template versions (CY23.2 <-> CY23Final).

import { Doc } from './doc.js';
import { addr, parseAddr } from './refs.js';
import { astToString } from './formula.js';
import { XErr } from './values.js';
import { findDV, validateValue } from './rules.js';

// The version the author wrote at the top of 更新说明, e.g. "2023.4.05更新 版本号：CY23.2".
export function detectVersion(wb) {
  const sh = wb.sheets.find((s) => s.name === '更新说明');
  if (!sh) return null;
  let best = null;
  for (const c of sh.cells.values()) {
    if (typeof c.value !== 'string') continue;
    const m = /版本号[：:]?\s*([A-Za-z0-9.]+)/.exec(c.value);
    if (m && (!best || c.r < best.r || (c.r === best.r && c.c < best.c))) best = { r: c.r, c: c.c, v: m[1] };
  }
  return best ? best.v : null;
}

function contentOf(cell) {
  if (!cell) return { clear: true };
  if (cell.ast) return { formula: astToString(cell.ast) };
  if (cell.value === null) return { clear: true };
  if (cell.value instanceof XErr) return { err: cell.value.code };
  return { value: cell.value };
}

function sameContent(a, b) {
  if (a.formula !== undefined || b.formula !== undefined) return a.formula === b.formula;
  if (a.clear || b.clear) return !!a.clear === !!b.clear;
  if (a.err !== undefined || b.err !== undefined) return a.err === b.err;
  return a.value === b.value;
}

// Cells whose content differs from the blank template, in the shape of Doc.editList().
// Used for imported cards, whose edits are baked into the file rather than recorded.
export function contentEdits(doc, blank) {
  const out = [];
  for (const sh of doc.wb.sheets) {
    const bs = blank.wb.sheets.find((s) => s.name === sh.name);
    if (!bs) continue;
    const keys = new Set([...sh.cells.keys(), ...bs.cells.keys()]);
    for (const k of keys) {
      const a = sh.cells.get(k), b = bs.cells.get(k);
      const ca = contentOf(a);
      if (sameContent(ca, contentOf(b))) continue;
      const c = a || b;
      const e = { sheet: sh.name, a: addr(c.r, c.c), seq: 0, ...ca };
      const xa = a && doc.wb.styles.cellXfs[a.s], xb = b && blank.wb.styles.cellXfs[b.s];
      if (xa && xb && xa.numFmtId !== xb.numFmtId) e.numFmtId = xa.numFmtId;
      if (xa && xa.quotePrefix && !(xb && xb.quotePrefix)) e.quotePrefix = true;
      out.push(e);
    }
  }
  return out;
}

// Replays `edits` (made on `srcBlank`'s version) onto a fresh copy of `targetBytes`.
// `srcDoc` is the card as it is now; it tells which values were valid before the move.
// Returns the converted Doc and the places the player should look at again.
export function convertCard(edits, srcDoc, srcBlank, targetBytes) {
  const target = new Doc(targetBytes);
  // custom number formats (id >= 164) are numbered per workbook; keep the target's own
  const list = edits.map((e) => (e.numFmtId >= 164 ? { ...e, numFmtId: undefined } : e));
  const at = (d, e) => {
    const si = d.wb.sheets.findIndex((s) => s.name === e.sheet);
    const p = parseAddr(e.a);
    return si < 0 || !p ? null : { si, r: p.r, c: p.c };
  };
  const notes = [];
  for (const e of list) {
    const t = at(target, e), s = at(srcBlank, e);
    if (!t) { notes.push({ sheet: e.sheet, a: e.a, kind: 'missing' }); continue; }
    const tc = target.cell(t.si, t.r, t.c), sc = s && srcBlank.cell(s.si, s.r, s.c);
    if (e.formula === undefined && tc && tc.ast && !(sc && sc.ast)) notes.push({ sheet: e.sheet, a: e.a, kind: 'formula' });
  }
  target.applyEditList(list);
  for (const e of list) {
    if (e.formula !== undefined || e.clear) continue;
    const t = at(target, e), s = at(srcDoc, e);
    if (!t || !s) continue;
    const dv = findDV(target, t.si, t.r, t.c);
    if (!dv) continue;
    const v = target.cell(t.si, t.r, t.c).value;
    if (validateValue(target, t.si, t.r, t.c, dv, v)) continue;
    const sdv = findDV(srcDoc, s.si, s.r, s.c);
    if (sdv && !validateValue(srcDoc, s.si, s.r, s.c, sdv, srcDoc.cell(s.si, s.r, s.c).value)) continue;
    notes.push({ sheet: e.sheet, a: e.a, kind: 'invalid', value: v });
  }
  return { doc: target, notes };
}
