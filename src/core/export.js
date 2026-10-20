// Write the document back into its own .xlsx package, touching only what an
// Excel save of the same edits would touch.

import { writeZip } from './zip.js';
import { escapeText, escapeAttr, encodeXstring, parseXml } from './xml.js';
import { addr, key, parseAddr, parseRange, rangeToStr } from './refs.js';
import { XErr, Ref } from './values.js';
import { astToString, parseFormula } from './formula.js';

const te = new TextEncoder();

// How numbers are written into <v>. Excel's output is the shortest text that
// reads back to the same double in the large majority of cases (7.03, not
// 7.0300000000000002); it only differs for a few computed values.
export function excelNum(x) {
  if (Object.is(x, -0)) x = 0;
  const s = String(x);
  if (s.indexOf('e') < 0) return s;
  const m = /^(-?[\d.]+)e([+-])(\d+)$/.exec(s);
  return m[1] + 'E' + m[2] + m[3].padStart(2, '0');
}

function needsPreserve(s) {
  return /^[\s]|[\s]$/.test(s);
}

// Text nodes: Excel writes line breaks as CRLF.
export function textEl(tag, s) {
  if (s === '') return `<${tag}/>`;
  const body = escapeText(encodeXstring(s)).replace(/\n/g, '\r\n');
  return needsPreserve(s) ? `<${tag} xml:space="preserve">${body}</${tag}>` : `<${tag}>${body}</${tag}>`;
}

function sameTyped(a, b) {
  if (a === null || b === null) return a === b;
  if (typeof a !== typeof b && !(a instanceof XErr && b instanceof XErr)) return false;
  return sameValue(a, b);
}

function sameValue(a, b) {
  if (a instanceof XErr || b instanceof XErr) return a instanceof XErr && b instanceof XErr && a.code === b.code;
  if (typeof a === 'number' && typeof b === 'number') return a === b || excelNum(a) === excelNum(b);
  return a === b;
}

function nowW3C(d) {
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

// ---- shared strings -------------------------------------------------------------

function planSST(doc) {
  const wb = doc.wb;
  const base = wb.sst.items;
  const plainIndex = new Map();
  base.forEach((it, i) => {
    if (!it.runs && !it.hasPhonetic && !plainIndex.has(it.text)) plainIndex.set(it.text, i);
  });
  const refs = []; // {cell, si, old: index|-1, text}
  const used = new Set();
  const fresh = new Map(); // text -> min seq
  wb.sheets.forEach((sh, si) => {
    for (const cell of sh.cells.values()) {
      if (cell.ast || typeof cell.value !== 'string') continue;
      let idx = -1;
      if (!cell.edited && cell.sst >= 0) idx = cell.sst;
      else if (plainIndex.has(cell.value)) idx = plainIndex.get(cell.value);
      if (idx >= 0) used.add(idx);
      else {
        const seq = cell.strSeq || 0;
        if (!fresh.has(cell.value) || fresh.get(cell.value) > seq) fresh.set(cell.value, seq);
      }
      refs.push({ cell, si, idx });
    }
  });
  const newStrings = [...fresh.entries()].sort((a, b) => a[1] - b[1]).map((e) => e[0]);
  const remap = new Map();
  let n = 0;
  for (let i = 0; i < base.length; i++) if (used.has(i)) remap.set(i, n++);
  const freshIndex = new Map();
  for (const s of newStrings) freshIndex.set(s, n++);
  const total = refs.length;
  const changed = used.size !== base.length || newStrings.length > 0 || total !== +(wb.sst.root ? wb.sst.root.attrs.count || -1 : -1);
  for (const r of refs) r.newIdx = r.idx >= 0 ? remap.get(r.idx) : freshIndex.get(r.cell.value);
  return { refs, used, newStrings, remap, total, unique: n, changed };
}

function buildSSTXml(doc, plan) {
  const wb = doc.wb;
  const xml = wb.sst.xml;
  if (!xml) {
    let out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + plan.total + '" uniqueCount="' + plan.unique + '">';
    for (const s of plan.newStrings) out += '<si>' + textEl('t', s) + '</si>';
    return out + '</sst>';
  }
  const root = wb.sst.root;
  let head = xml.slice(0, root.openEnd);
  head = head.replace(/\bcount="\d+"/, `count="${plan.total}"`).replace(/\buniqueCount="\d+"/, `uniqueCount="${plan.unique}"`);
  const parts = [head];
  wb.sst.items.forEach((it, i) => { if (plan.used.has(i)) parts.push(it.raw); });
  for (const s of plan.newStrings) parts.push('<si>' + textEl('t', s) + '</si>');
  const sis = root.all('si');
  const tailStart = sis.length ? sis[sis.length - 1].end : root.openEnd;
  parts.push(xml.slice(tailStart));
  return parts.join('');
}

// ---- cells ------------------------------------------------------------------------

function cellXml(sh, cell, sstIndex, fXml) {
  const a = cell.node ? cell.node.attrs : {};
  let attrs = ` r="${addr(cell.r, cell.c)}"`;
  if (cell.s) attrs += ` s="${cell.s}"`;
  const v = cell.value;
  let t = null, vXml = null;
  if (cell.ast) {
    if (typeof v === 'string') { t = 'str'; vXml = textEl('v', v); }
    else if (typeof v === 'boolean') { t = 'b'; vXml = `<v>${v ? 1 : 0}</v>`; }
    else if (v instanceof XErr) { t = 'e'; vXml = `<v>${escapeText(v.code)}</v>`; }
    else if (typeof v === 'number') vXml = `<v>${excelNum(v)}</v>`;
    else vXml = '<v>0</v>';
  } else if (typeof v === 'string') { t = 's'; vXml = `<v>${sstIndex}</v>`; }
  else if (typeof v === 'boolean') { t = 'b'; vXml = `<v>${v ? 1 : 0}</v>`; }
  else if (v instanceof XErr) { t = 'e'; vXml = `<v>${escapeText(v.code)}</v>`; }
  else if (typeof v === 'number') vXml = `<v>${excelNum(v)}</v>`;
  if (t) attrs += ` t="${t}"`;
  for (const extra of ['cm', 'vm', 'ph']) if (a[extra] !== undefined) attrs += ` ${extra}="${escapeAttr(a[extra])}"`;
  const inner = (fXml || '') + (vXml || '');
  return inner ? `<c${attrs}>${inner}</c>` : `<c${attrs}/>`;
}

function formulaXml(sh, cell, masters) {
  if (!cell.ast) return '';
  const fn = cell.base && cell.base.fnode;
  if (!cell.edited && fn) {
    const promo = masters.get(cell);
    if (promo) return promo;
    // master overwritten? (handled through `masters`) — otherwise reuse the original element
    return sh.xml.slice(fn.start, fn.end);
  }
  return `<f>${escapeText(cell.formula)}</f>`;
}

// When the master cell of a shared formula is overwritten, the first remaining
// cell of the group becomes the new master (with the full formula text).
function planSharedMasters(sh) {
  const out = new Map();
  const groups = new Map();
  for (const cell of sh.cells.values()) {
    const fn = cell.base && cell.base.fnode;
    if (!fn || fn.attrs.t !== 'shared') continue;
    const si = fn.attrs.si;
    if (!groups.has(si)) groups.set(si, { master: null, members: [] });
    const g = groups.get(si);
    if (fn.text) g.master = cell;
    g.members.push(cell);
  }
  for (const [si, g] of groups) {
    if (!g.master) continue;
    const m = g.master;
    if (m.ast && !m.edited) continue; // master still intact
    const alive = g.members.filter((c) => c !== m && c.ast && !c.edited).sort((a, b) => a.r - b.r || a.c - b.c);
    if (!alive.length) continue;
    const nm = alive[0];
    let r1 = Infinity, c1 = Infinity, r2 = -Infinity, c2 = -Infinity;
    for (const c of alive) { r1 = Math.min(r1, c.r); c1 = Math.min(c1, c.c); r2 = Math.max(r2, c.r); c2 = Math.max(c2, c.c); }
    const ref = rangeToStr({ r1, c1, r2, c2 });
    out.set(nm, `<f t="shared" ref="${ref}" si="${si}">${escapeText(astToString(nm.ast))}</f>`);
  }
  return out;
}

function rowOpenTag(row, spans) {
  // rebuild a <row> start tag with updated spans, preserving the other attributes
  const raw = row.node ? row.node.attrs : { r: String(row.r) };
  let s = '<row';
  for (const [k, v] of Object.entries(raw)) {
    if (k === 'spans') s += ` spans="${spans}"`;
    else s += ` ${k}="${escapeAttr(v)}"`;
  }
  if (!('spans' in raw) && spans) s = s.replace('<row', '<row').replace(/ r="(\d+)"/, ` r="$1" spans="${spans}"`);
  return s + '>';
}

function patchSheet(doc, si, sstPlanByCell, opts, changedOut) {
  const sh = doc.wb.sheets[si];
  const xml = sh.xml;
  const splices = []; // [start, end, text]
  const masters = planSharedMasters(sh);
  const anyEdits = opts.anyEdits;
  const inserts = new Map(); // row -> [cells]
  let changed = false;
  for (const cell of sh.cells.values()) {
    const b = cell.base;
    const newIdx = sstPlanByCell.get(cell);
    let need;
    if (!b) need = cell.value !== null || !!cell.ast;
    else {
      need = cell.s !== b.s || !!cell.ast !== b.hasF || masters.has(cell);
      if (!need && cell.ast) {
        if (cell.edited) need = !b.fnode || b.fnode.attrs.t === 'shared' || cell.formula !== b.fnode.text;
        if (!need && !(cell.volatile && !anyEdits)) need = !sameTyped(cell.value, b.value);
      } else if (!need) {
        need = !sameTyped(cell.value, b.value) || (typeof cell.value === 'string' && newIdx !== b.sst);
      }
    }
    if (opts.verifyAll && cell.node) {
      const gen = cellXml(sh, cell, newIdx, formulaXml(sh, cell, masters));
      if (gen !== xml.slice(cell.node.start, cell.node.end)) opts.verifyAll.push([sh.name, addr(cell.r, cell.c), xml.slice(cell.node.start, cell.node.end), gen]);
    }
    if (!need) continue;
    const fXml = formulaXml(sh, cell, masters);
    const text = cellXml(sh, cell, newIdx, fXml);
    changed = true;
    changedOut.push(`${sh.name}!${addr(cell.r, cell.c)}`);
    if (cell.node) splices.push([cell.node.start, cell.node.end, text]);
    else {
      if (!inserts.has(cell.r)) inserts.set(cell.r, []);
      inserts.get(cell.r).push({ c: cell.c, text });
    }
  }
  // an unprotected sheet loses its <sheetProtection> element, as when saved by Excel
  const protNode = sh.root.child('sheetProtection');
  if (protNode && doc.unprotected && doc.unprotected.has(si)) {
    splices.push([protNode.start, protNode.end, '']);
    changed = true;
    changedOut.push(`${sh.name}!(取消保护)`);
  }
  if (!changed) return null;

  // insert new cells into existing rows / new rows
  const sd = sh.sheetDataNode;
  const rowNums = [...inserts.keys()].sort((a, b) => a - b);
  const touchedBlocks = new Set();
  for (const r of rowNums) {
    const list = inserts.get(r).sort((a, b) => a.c - b.c);
    touchedBlocks.add(Math.floor((r - 1) / 16));
    const row = sh.rows.get(r);
    if (row && row.node) {
      const cellsInRow = row.node.all('c');
      for (const ins of list) {
        const nextCell = cellsInRow.find((cn) => parseAddr(cn.attrs.r).c > ins.c);
        if (nextCell) splices.push([nextCell.start, nextCell.start, ins.text]);
        else if (row.node.selfClosing) {
          // <row .../> -> <row ...>cells</row>
          splices.push([row.node.start, row.node.end, xml.slice(row.node.start, row.node.end - 2) + '>' + ins.text + '</row>']);
        } else splices.push([row.node.end - 6, row.node.end - 6, ins.text]);
      }
    } else {
      const text = `<row r="${r}" spans="__SPANS${r}__">` + list.map((x) => x.text).join('') + '</row>';
      const nextRow = sd ? sd.all('row').find((rn) => +rn.attrs.r > r) : null;
      if (nextRow) splices.push([nextRow.start, nextRow.start, text]);
      else if (sd && !sd.selfClosing) splices.push([sd.end - 12, sd.end - 12, text]);
      else if (sd) splices.push([sd.start, sd.end, '<sheetData>' + text + '</sheetData>']);
      sh.rows.set(r, { r, ht: null, customHeight: false, hidden: false, s: null, customFormat: false, node: null, pending: true });
    }
  }
  // recompute spans for touched 16-row blocks
  for (const blk of touchedBlocks) {
    let mn = Infinity, mx = -Infinity;
    for (let r = blk * 16 + 1; r <= blk * 16 + 16; r++) {
      for (const cell of sh.cells.values()) if (cell.r === r) { mn = Math.min(mn, cell.c); mx = Math.max(mx, cell.c); }
    }
    const spans = `${mn}:${mx}`;
    for (let r = blk * 16 + 1; r <= blk * 16 + 16; r++) {
      const row = sh.rows.get(r);
      if (!row) continue;
      if (row.node) {
        if (row.node.attrs.spans !== spans) {
          const openEnd = row.node.selfClosing ? row.node.end : row.node.openEnd;
          const tag = rowOpenTag(row, spans);
          // keep self-closing form if it was
          splices.push([row.node.start, openEnd, row.node.selfClosing && !inserts.has(r) ? tag.replace(/>$/, '/>') : tag, 'rowtag']);
        }
      } else if (row.pending) {
        for (const sp of splices) if (typeof sp[2] === 'string') sp[2] = sp[2].replace(`__SPANS${r}__`, spans);
        row.pending = false;
      }
    }
  }
  // dimension
  const dimNode = sh.root.child('dimension');
  if (dimNode) {
    let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
    for (const cell of sh.cells.values()) { r1 = Math.min(r1, cell.r); c1 = Math.min(c1, cell.c); r2 = Math.max(r2, cell.r); c2 = Math.max(c2, cell.c); }
    const old = parseRange(dimNode.attrs.ref);
    if (old && r2 > 0) {
      const g = { r1: Math.min(old.r1, r1), c1: Math.min(old.c1, c1), r2: Math.max(old.r2, r2), c2: Math.max(old.c2, c2) };
      const ref = rangeToStr(g);
      if (ref !== dimNode.attrs.ref) splices.push([dimNode.start, dimNode.end, `<dimension ref="${ref}"/>`]);
    }
  }
  return applySplices(xml, splices, sh);
}

function applySplices(xml, splices, sh) {
  // a row-tag splice may overlap an insertion at the row end for self-closing rows; resolve by merging
  splices.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  let out = '';
  let pos = 0;
  for (const [s, e, text] of splices) {
    if (s < pos) {
      // overlapping (self-closing row converted + retagged): skip the retag
      continue;
    }
    out += xml.slice(pos, s) + text;
    pos = e;
  }
  return out + xml.slice(pos);
}

// ---- calcChain -----------------------------------------------------------------------

function patchCalcChain(doc) {
  const wb = doc.wb;
  if (!wb.calcChainPath) return null;
  const xml = wb.zip.text(wb.calcChainPath);
  if (!xml) return null;
  const root = parseXml(xml);
  const entries = root.all('c');
  const keep = [];
  const present = new Set();
  let curSheet = null;
  let changed = false;
  for (const e of entries) {
    if (e.attrs.i !== undefined) curSheet = +e.attrs.i;
    const si = wb.sheets.findIndex((s) => s.sheetId === curSheet);
    const a = parseAddr(e.attrs.r);
    const cell = si >= 0 && a ? wb.sheets[si].cells.get(key(a.r, a.c)) : null;
    if (!cell || !cell.ast) { changed = true; continue; }
    present.add(cell);
    keep.push({ sheetId: curSheet, attrs: e.attrs });
  }
  wb.sheets.forEach((sh) => {
    for (const cell of sh.cells.values()) {
      if (cell.ast && !present.has(cell)) {
        keep.push({ sheetId: sh.sheetId, attrs: { r: addr(cell.r, cell.c), i: String(sh.sheetId) } });
        changed = true;
      }
    }
  });
  if (!changed) return null;
  let body = '';
  let prev = null;
  for (const k of keep) {
    const at = { ...k.attrs };
    if (at.i === undefined && k.sheetId !== prev) at.i = String(k.sheetId);
    let s = `<c r="${at.r}"`;
    if (at.i !== undefined) s += ` i="${at.i}"`;
    for (const [kk, vv] of Object.entries(at)) if (kk !== 'r' && kk !== 'i') s += ` ${kk}="${vv}"`;
    body += s + '/>';
    prev = k.sheetId;
  }
  return xml.slice(0, root.openEnd) + body + xml.slice(entries.length ? entries[entries.length - 1].end : root.openEnd);
}

// ---- charts ----------------------------------------------------------------------------

function patchChart(doc, path) {
  const wb = doc.wb;
  const xml = wb.zip.text(path);
  if (!xml) return null;
  const root = parseXml(xml);
  const splices = [];
  for (const kind of ['numRef', 'strRef']) {
    for (const ref of root.find(kind)) {
      const f = ref.child('f');
      const cache = ref.child(kind === 'numRef' ? 'numCache' : 'strCache');
      if (!f || !cache) continue;
      let ast;
      try { ast = parseFormula(f.text); } catch (e) { continue; }
      const v = doc.eng.evalAt(ast, 0, 1, 1, { keepRef: true });
      if (!(v instanceof Ref)) continue;
      const vals = [];
      for (let r = v.r1; r <= v.r2; r++) for (let c = v.c1; c <= v.c2; c++) vals.push(doc.eng.getValue(v.sheet, r, c));
      let pts = '';
      vals.forEach((x, i) => {
        if (x === null || x === undefined) return;
        if (kind === 'numRef') {
          // errors and text are cached as 0 (as in the original file's #DIV/0! points)
          pts += `<c:pt idx="${i}"><c:v>${typeof x === 'number' ? excelNum(x) : typeof x === 'boolean' ? (x ? 1 : 0) : 0}</c:v></c:pt>`;
        } else {
          const s = typeof x === 'number' ? String(x) : x instanceof XErr ? x.code : typeof x === 'boolean' ? (x ? 'TRUE' : 'FALSE') : x;
          pts += `<c:pt idx="${i}"><c:v>${escapeText(encodeXstring(s)).replace(/\n/g, '\r\n')}</c:v></c:pt>`;
        }
      });
      const fc = cache.child('formatCode');
      const head = (fc ? xml.slice(fc.start, fc.end) : '') + `<c:ptCount val="${vals.length}"/>`;
      const ext = cache.child('extLst');
      const tail = ext ? xml.slice(ext.start, ext.end) : '';
      const inner = head + pts + tail;
      const oldInner = xml.slice(cache.openEnd, cache.end - (`</${cache.name}>`).length);
      if (inner !== oldInner) splices.push([cache.openEnd, cache.end - (`</${cache.name}>`).length, inner]);
    }
  }
  if (!splices.length) return null;
  return applySplices(xml, splices);
}

// ---- styles -----------------------------------------------------------------------------

function patchStyles(doc) {
  if (!doc.styleAdds.length) return null;
  const st = doc.wb.styles;
  const xml = st.xml;
  const cx = st.root.child('cellXfs');
  const xfNodes = cx.all('xf');
  let added = '';
  for (const a of doc.styleAdds) {
    const bn = xfNodes[a.base];
    let raw = xml.slice(bn.start, bn.end);
    // numFmtId
    raw = raw.replace(/numFmtId="\d+"/, `numFmtId="${a.numFmtId}"`);
    if (a.quotePrefix && !/quotePrefix=/.test(raw)) raw = raw.replace(/(xfId="\d+")/, '$1 quotePrefix="1"');
    if (!/applyNumberFormat=/.test(raw) && a.numFmtId !== doc.wb.styles.cellXfs[a.base].numFmtId) {
      raw = /quotePrefix="1"/.test(raw) ? raw.replace(/(quotePrefix="1")/, '$1 applyNumberFormat="1"') : raw.replace(/(xfId="\d+")/, '$1 applyNumberFormat="1"');
    }
    added += raw;
  }
  const total = xfNodes.length + doc.styleAdds.length;
  const insertAt = cx.end - '</cellXfs>'.length;
  let out = xml.slice(0, insertAt) + added + xml.slice(insertAt);
  out = out.slice(0, cx.start) + out.slice(cx.start).replace(/<cellXfs count="\d+"/, `<cellXfs count="${total}"`);
  return out;
}

// ---- docProps / workbook view -----------------------------------------------------------

function patchCore(doc, opts) {
  const p = 'docProps/core.xml';
  const xml = doc.wb.zip.text(p);
  if (!xml) return null;
  let out = xml;
  if (opts.modifiedBy !== null && opts.modifiedBy !== undefined) {
    if (/<cp:lastModifiedBy>[\s\S]*?<\/cp:lastModifiedBy>|<cp:lastModifiedBy\/>/.test(out)) {
      out = out.replace(/<cp:lastModifiedBy>[\s\S]*?<\/cp:lastModifiedBy>|<cp:lastModifiedBy\/>/, `<cp:lastModifiedBy>${escapeText(opts.modifiedBy)}</cp:lastModifiedBy>`);
    } else {
      out = out.replace(/(<dc:creator>[\s\S]*?<\/dc:creator>)/, `$1<cp:lastModifiedBy>${escapeText(opts.modifiedBy)}</cp:lastModifiedBy>`);
    }
  }
  if (opts.modified) {
    out = out.replace(/(<dcterms:modified[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${nowW3C(opts.modified)}$2`);
  }
  return out === xml ? null : out;
}

function patchViews(doc, opts, sheetXml) {
  // active tab + selection of the active sheet
  const wb = doc.wb;
  const res = { workbook: null };
  if (opts.activeSheet === undefined || opts.activeSheet === null) return res;
  const act = opts.activeSheet;
  if (act !== wb.activeTab) {
    let x = wb.xml;
    const wv = wb.root.child('bookViews') && wb.root.child('bookViews').child('workbookView');
    if (wv) {
      let tag = x.slice(wv.start, wv.openEnd);
      if (/activeTab="\d+"/.test(tag)) tag = act ? tag.replace(/activeTab="\d+"/, `activeTab="${act}"`) : tag.replace(/ activeTab="\d+"/, '');
      else if (act) tag = /\sxr2:uid=/.test(tag) ? tag.replace(/(\sxr2:uid=)/, ` activeTab="${act}"$1`) : tag.replace(/(\/?>)$/, ` activeTab="${act}"$1`);
      x = x.slice(0, wv.start) + tag + x.slice(wv.openEnd);
      res.workbook = x;
    }
  }
  wb.sheets.forEach((sh, si) => {
    if (!sh.view || !sh.view.node) return;
    const cur = sheetXml.has(si) ? sheetXml.get(si) : sh.xml;
    // re-parse only the sheetView start tag region from the current text
    const m = /<sheetView\b[^>]*>/.exec(cur);
    if (!m) return;
    let tag = m[0];
    const was = /tabSelected="1"/.test(tag);
    const should = si === act;
    let next = cur;
    if (was !== should) {
      tag = should ? tag.replace(/<sheetView\b/, '<sheetView tabSelected="1"').replace(/<sheetView tabSelected="1"((?:\s+(?:showGridLines|showRowColHeaders|showZeros|rightToLeft)="[^"]*")*)/, '<sheetView$1 tabSelected="1"') : tag.replace(/\s+tabSelected="1"/, '');
      next = cur.slice(0, m.index) + tag + cur.slice(m.index + m[0].length);
    }
    if (should && opts.selection && opts.selection.si === si) {
      const sel = opts.selection;
      const sq = rangeToStr(sel);
      const ac = addr(sel.ar, sel.ac);
      next = next.replace(/(<sheetView\b[\s\S]*?)(<selection\b[^>]*?)\/>((?![\s\S]*<selection\b)[\s\S]*?<\/sheetView>)/, (all, pre, s, post) => {
        let t = s;
        t = /activeCell="/.test(t) ? t.replace(/activeCell="[^"]*"/, `activeCell="${ac}"`) : t + ` activeCell="${ac}"`;
        t = /sqref="/.test(t) ? t.replace(/sqref="[^"]*"/, `sqref="${sq}"`) : t + ` sqref="${sq}"`;
        return pre + t + '/>' + post;
      });
    }
    if (next !== cur) sheetXml.set(si, next);
  });
  return res;
}

// ---- main --------------------------------------------------------------------------------

export function exportXlsx(doc, opts = {}) {
  const wb = doc.wb;
  const replacements = new Map();
  const changedCells = [];
  const plan = planSST(doc);
  const sstPlanByCell = new Map(plan.refs.map((r) => [r.cell, r.newIdx]));
  const anyEdits = wb.sheets.some((sh) => [...sh.cells.values()].some((c) => c.edited));
  const sheetXml = new Map();
  wb.sheets.forEach((sh, si) => {
    if (!sh.xml) return;
    const out = patchSheet(doc, si, sstPlanByCell, { anyEdits, verifyAll: opts.verifyAll }, changedCells);
    if (out !== null) sheetXml.set(si, out);
  });
  const views = opts.saveView ? patchViews(doc, opts, sheetXml) : { workbook: null };
  for (const [si, x] of sheetXml) replacements.set(wb.sheets[si].path, x);
  if (views.workbook) replacements.set(wb.path, views.workbook);
  if (plan.changed && wb.sstPath) replacements.set(wb.sstPath, buildSSTXml(doc, plan));
  const cc = patchCalcChain(doc);
  if (cc) replacements.set(wb.calcChainPath, cc);
  const st = patchStyles(doc);
  if (st) replacements.set(wb.stylesPath, st);
  for (const sh of wb.sheets) {
    for (const d of sh.drawings || []) {
      if (d.kind !== 'chart' || replacements.has(d.path)) continue;
      const x = patchChart(doc, d.path);
      if (x) replacements.set(d.path, x);
    }
  }
  if (replacements.size || opts.forceProps) {
    const core = patchCore(doc, opts);
    if (core) replacements.set('docProps/core.xml', core);
  }
  const bytes = replacements.size ? writeZip(wb.zip, replacements) : doc.bytes;
  return { bytes, changedCells, parts: [...replacements.keys()] };
}
