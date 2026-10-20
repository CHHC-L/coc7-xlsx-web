// Parse an .xlsx package into the in-memory model used by the engine, the
// renderer and the exporter.

import { readZip } from './zip.js';
import { parseXml, decodeXstring } from './xml.js';
import { parseStyles, parseTheme, parseFont } from './styles.js';
import { parseRange, parseSqref, key, colToNum, parseAddr } from './refs.js';
import { parseFormula, shiftAst } from './formula.js';
import { errFromCode } from './values.js';

const REL_NS_OFFDOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

function dirOf(p) { const i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i + 1); }

export function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dirOf(base) + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

function relsPathFor(part) {
  const d = dirOf(part);
  return d + '_rels/' + part.slice(d.length) + '.rels';
}

function readRels(zip, part) {
  const txt = zip.text(relsPathFor(part));
  const rels = new Map();
  if (!txt) return rels;
  for (const r of parseXml(txt).all('Relationship')) {
    rels.set(r.attrs.Id, {
      id: r.attrs.Id,
      type: r.attrs.Type,
      target: r.attrs.Target,
      external: r.attrs.TargetMode === 'External',
      path: r.attrs.TargetMode === 'External' ? null : resolvePath(part, r.attrs.Target),
    });
  }
  return rels;
}

function parseRich(si) {
  // returns {text, runs|null}
  const runs = si.all('r');
  if (!runs.length) {
    const t = si.child('t');
    return { text: decodeXstring(t ? t.text : ''), runs: null };
  }
  let text = '';
  const out = [];
  // a plain <t> may precede runs
  for (const ch of si.children) {
    if (ch.local === 't') {
      const s = decodeXstring(ch.text);
      text += s;
      out.push({ text: s, font: null });
    } else if (ch.local === 'r') {
      const t = ch.child('t');
      const s = decodeXstring(t ? t.text : '');
      text += s;
      out.push({ text: s, font: ch.child('rPr') ? parseFont(ch.child('rPr')) : null });
    }
  }
  return { text, runs: out };
}

function parseSST(xml) {
  if (!xml) return { items: [], root: null };
  const root = parseXml(xml);
  const items = root.all('si').map((si) => {
    const { text, runs } = parseRich(si);
    return { text, runs, raw: xml.slice(si.start, si.end), hasPhonetic: !!si.child('rPh') };
  });
  return { items, root, xml };
}

function attrBool(v, dflt = false) {
  if (v === undefined) return dflt;
  return v === '1' || v === 'true';
}

function parseDV(n, isX14) {
  const a = n.attrs;
  let f1 = null, f2 = null, sq = a.sqref || '';
  if (isX14) {
    const x1 = n.child('formula1'), x2 = n.child('formula2');
    f1 = x1 ? x1.deepText() : null;
    f2 = x2 ? x2.deepText() : null;
    const s = n.child('sqref');
    if (s) sq = s.text;
  } else {
    const x1 = n.child('formula1'), x2 = n.child('formula2');
    f1 = x1 ? x1.text : null;
    f2 = x2 ? x2.text : null;
  }
  return {
    ranges: parseSqref(sq),
    sqref: sq,
    type: a.type || 'none',
    operator: a.operator || 'between',
    allowBlank: attrBool(a.allowBlank),
    hideDropDown: attrBool(a.showDropDown),
    showInputMessage: attrBool(a.showInputMessage),
    showErrorMessage: attrBool(a.showErrorMessage),
    errorStyle: a.errorStyle || 'stop',
    errorTitle: a.errorTitle ? decodeXstring(a.errorTitle) : '',
    error: a.error ? decodeXstring(a.error) : '',
    promptTitle: a.promptTitle ? decodeXstring(a.promptTitle) : '',
    prompt: a.prompt ? decodeXstring(a.prompt) : '',
    formula1: f1,
    formula2: f2,
  };
}

function parseAnchorPoint(n) {
  if (!n) return null;
  const v = (k) => +(n.child(k) ? n.child(k).text : 0);
  return { col: v('col'), colOff: v('colOff'), row: v('row'), rowOff: v('rowOff') };
}

function parseChart(xml) {
  const root = parseXml(xml);
  const plot = root.find('plotArea')[0];
  if (!plot) return null;
  const kinds = ['radarChart', 'barChart', 'lineChart', 'pieChart', 'areaChart', 'scatterChart', 'doughnutChart'];
  let ch = null;
  for (const k of kinds) { const f = plot.child(k); if (f) { ch = f; break; } }
  if (!ch) return null;
  const series = ch.all('ser').map((s) => {
    const cat = s.child('cat'), val = s.child('val');
    const refOf = (n) => {
      if (!n) return null;
      const r = n.find('f')[0];
      return r ? r.text : null;
    };
    const fillClr = (() => {
      const sp = s.child('spPr');
      if (!sp) return null;
      const sf = sp.child('solidFill');
      if (!sf || !sf.children[0]) return null;
      const c = sf.children[0];
      return { kind: c.local, val: c.attrs.val, mods: c.children.map((m) => ({ k: m.local, v: +m.attrs.val })) };
    })();
    return { cat: refOf(cat), val: refOf(val), fill: fillClr };
  });
  const layout = plot.child('layout');
  const ml = layout && layout.child('manualLayout');
  const lay = ml ? {
    x: +(ml.child('x') ? ml.child('x').attrs.val : 0), y: +(ml.child('y') ? ml.child('y').attrs.val : 0),
    w: +(ml.child('w') ? ml.child('w').attrs.val : 1), h: +(ml.child('h') ? ml.child('h').attrs.val : 1),
  } : null;
  return {
    type: ch.local,
    radarStyle: ch.child('radarStyle') ? ch.child('radarStyle').attrs.val : null,
    series,
    layout: lay,
    hasMinorGrid: !!root.find('minorGridlines').length,
  };
}

function parseVmlNotes(vml) {
  // returns Map(key -> {visible, anchor:[c1,dx1,r1,dy1,c2,dx2,r2,dy2]})
  const out = new Map();
  if (!vml) return out;
  const shapeRe = /<v:shape\b[\s\S]*?<\/v:shape>/g;
  let m;
  while ((m = shapeRe.exec(vml))) {
    const s = m[0];
    if (!/ObjectType="Note"/.test(s)) continue;
    const row = /<x:Row>(\d+)<\/x:Row>/.exec(s);
    const col = /<x:Column>(\d+)<\/x:Column>/.exec(s);
    if (!row || !col) continue;
    const anchor = /<x:Anchor>([^<]*)<\/x:Anchor>/.exec(s);
    const style = /style=['"]([^'"]*)['"]/.exec(s);
    out.set(key(+row[1] + 1, +col[1] + 1), {
      visible: /<x:Visible\s*\/>|<x:Visible>/.test(s) && !/visibility:\s*hidden/.test(style ? style[1] : ''),
      anchor: anchor ? anchor[1].split(',').map((x) => +x.trim()) : null,
    });
  }
  return out;
}

export class Cell {
  constructor(r, c) {
    this.r = r; this.c = c;
    this.s = 0;
    this.t = null; // original type attribute
    this.value = null;
    this.formula = null; // formula text without "=" (null if constant)
    this.ast = null;
    this.fattrs = null; // attributes of the original <f>
    this.sst = -1;
    this.runs = null; // rich text runs for display
    this.node = null; // original XNode (for export)
    this.volatile = false;
  }
}

function parseSheet(wb, sh) {
  const xml = sh.xml;
  const root = parseXml(xml);
  sh.root = root;
  const sp = root.child('sheetPr');
  sh.tabColor = sp && sp.child('tabColor') ? sp.child('tabColor').attrs : null;
  const dim = root.child('dimension');
  sh.dimension = dim ? dim.attrs.ref : null;
  const views = root.child('sheetViews');
  const v0 = views && views.child('sheetView');
  sh.view = {
    showGridLines: v0 ? attrBool(v0.attrs.showGridLines, true) : true,
    showRowColHeaders: v0 ? attrBool(v0.attrs.showRowColHeaders, true) : true,
    zoomScale: v0 && v0.attrs.zoomScale ? +v0.attrs.zoomScale : 100,
    topLeftCell: v0 ? v0.attrs.topLeftCell || null : null,
    pane: null,
    selection: null,
    node: v0,
  };
  if (v0) {
    const pane = v0.child('pane');
    if (pane && (pane.attrs.state === 'frozen' || pane.attrs.state === 'frozenSplit')) {
      sh.view.pane = { xSplit: +(pane.attrs.xSplit || 0), ySplit: +(pane.attrs.ySplit || 0), topLeftCell: pane.attrs.topLeftCell };
    }
    const sel = v0.all('selection');
    if (sel.length) {
      const s = sel[sel.length - 1];
      sh.view.selection = { activeCell: s.attrs.activeCell || null, sqref: s.attrs.sqref || null };
    }
  }
  const fp = root.child('sheetFormatPr');
  sh.defaultColWidth = fp && fp.attrs.defaultColWidth ? +fp.attrs.defaultColWidth : null;
  sh.baseColWidth = fp && fp.attrs.baseColWidth ? +fp.attrs.baseColWidth : 8;
  sh.defaultRowHeight = fp && fp.attrs.defaultRowHeight ? +fp.attrs.defaultRowHeight : 15;
  sh.zeroHeight = fp ? attrBool(fp.attrs.zeroHeight) : false;
  sh.cols = [];
  const cols = root.child('cols');
  if (cols) {
    for (const c of cols.all('col')) {
      sh.cols.push({
        min: +c.attrs.min, max: +c.attrs.max,
        width: c.attrs.width !== undefined ? +c.attrs.width : null,
        style: +(c.attrs.style || 0),
        hidden: attrBool(c.attrs.hidden),
        customWidth: attrBool(c.attrs.customWidth),
      });
    }
  }
  sh.rows = new Map();
  sh.cells = new Map();
  sh.sheetDataNode = root.child('sheetData');
  const shared = new Map();
  let maxR = 0, maxC = 0;
  if (sh.sheetDataNode) {
    for (const rn of sh.sheetDataNode.all('row')) {
      const r = +rn.attrs.r;
      sh.rows.set(r, {
        r,
        ht: rn.attrs.ht !== undefined ? +rn.attrs.ht : null,
        customHeight: attrBool(rn.attrs.customHeight),
        hidden: attrBool(rn.attrs.hidden),
        s: rn.attrs.s !== undefined ? +rn.attrs.s : null,
        customFormat: attrBool(rn.attrs.customFormat),
        node: rn,
      });
      let lastC = 0;
      for (const cn of rn.all('c')) {
        let rr = r, cc;
        if (cn.attrs.r) {
          const a = parseAddr(cn.attrs.r);
          rr = a.r; cc = a.c;
        } else cc = lastC + 1;
        lastC = cc;
        const cell = new Cell(rr, cc);
        cell.node = cn;
        cell.s = +(cn.attrs.s || 0);
        const t = cn.attrs.t || 'n';
        cell.t = t;
        const vn = cn.child('v');
        const fn = cn.child('f');
        const vtext = vn ? vn.text : null;
        if (t === 's') {
          const idx = +vtext;
          cell.sst = idx;
          const item = wb.sst.items[idx];
          cell.value = item ? item.text : '';
          cell.runs = item && item.runs ? item.runs : null;
        } else if (t === 'inlineStr') {
          const is = cn.child('is');
          const rich = is ? parseRich(is) : { text: '', runs: null };
          cell.value = rich.text;
          cell.runs = rich.runs;
        } else if (t === 'str') {
          cell.value = vtext === null ? '' : decodeXstring(vtext);
        } else if (t === 'b') {
          cell.value = vtext === '1' || vtext === 'true';
        } else if (t === 'e') {
          cell.value = vtext ? errFromCode(vtext) : null;
        } else if (t === 'd') {
          cell.value = vtext; // ISO date text (rare)
        } else {
          cell.value = vtext === null || vtext === '' ? null : +vtext;
        }
        if (fn) {
          cell.fattrs = { ...fn.attrs };
          const ft = fn.attrs.t || 'normal';
          if (ft === 'shared') {
            const si = fn.attrs.si;
            if (fn.text) {
              cell.formula = fn.text;
              try { cell.ast = parseFormula(fn.text); } catch (e) { cell.ast = { type: 'err', v: errFromCode('#NAME?'), parseError: e.message }; }
              shared.set(si, { r: rr, c: cc, ast: cell.ast, text: fn.text });
            } else {
              const m = shared.get(si);
              if (m) {
                cell.ast = shiftAst(m.ast, rr - m.r, cc - m.c);
                cell.sharedFrom = m;
                cell.formula = null; // printed lazily
              }
            }
          } else if (ft === 'normal' || ft === 'array') {
            if (fn.text) {
              cell.formula = fn.text;
              try { cell.ast = parseFormula(fn.text); } catch (e) { cell.ast = { type: 'err', v: errFromCode('#NAME?'), parseError: e.message }; }
              if (ft === 'array') cell.arrayRef = fn.attrs.ref || null;
            }
          }
          if (cell.ast) cell.volatile = /RAND|NOW|TODAY|OFFSET|INDIRECT/i.test(fn.text || (cell.sharedFrom && cell.sharedFrom.text) || '');
        }
        cell.base = { value: cell.value, s: cell.s, t, hasF: !!fn, sst: cell.sst, fnode: fn };
        sh.cells.set(key(rr, cc), cell);
        if (rr > maxR) maxR = rr;
        if (cc > maxC) maxC = cc;
      }
    }
  }
  sh.maxRow = maxR; sh.maxCol = maxC;

  // merges
  sh.merges = [];
  sh.mergeAt = new Map(); // key of every covered cell -> merge range
  const mc = root.child('mergeCells');
  if (mc) {
    for (const m of mc.all('mergeCell')) {
      const g = parseRange(m.attrs.ref);
      if (!g) continue;
      sh.merges.push(g);
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) sh.mergeAt.set(key(r, c), g);
    }
  }

  // protection
  const prot = root.child('sheetProtection');
  sh.protection = prot ? {
    sheet: attrBool(prot.attrs.sheet),
    selectLockedCells: attrBool(prot.attrs.selectLockedCells),
    selectUnlockedCells: attrBool(prot.attrs.selectUnlockedCells),
    hasPassword: !!(prot.attrs.password || prot.attrs.hashValue),
    attrs: prot.attrs,
  } : null;
  sh.protectedRanges = [];
  const pr = root.child('protectedRanges');
  if (pr) {
    for (const p of pr.all('protectedRange')) {
      sh.protectedRanges.push({
        name: p.attrs.name,
        ranges: parseSqref(p.attrs.sqref || ''),
        hasPassword: !!(p.attrs.password || p.attrs.hashValue || p.attrs.securityDescriptor),
      });
    }
  }

  // data validation (standard + x14 extension)
  sh.dvs = [];
  const dvs = root.child('dataValidations');
  if (dvs) for (const d of dvs.all('dataValidation')) sh.dvs.push(parseDV(d, false));
  for (const x of root.find('dataValidation')) {
    if (x.name.startsWith('x14:')) sh.dvs.push(parseDV(x, true));
  }

  // conditional formatting
  sh.cfs = [];
  for (const cf of root.all('conditionalFormatting')) {
    const ranges = parseSqref(cf.attrs.sqref || '');
    for (const rule of cf.all('cfRule')) {
      const formulas = rule.all('formula').map((f) => f.text);
      sh.cfs.push({
        ranges,
        type: rule.attrs.type,
        operator: rule.attrs.operator || null,
        dxfId: rule.attrs.dxfId !== undefined ? +rule.attrs.dxfId : null,
        priority: +(rule.attrs.priority || 0),
        stopIfTrue: attrBool(rule.attrs.stopIfTrue),
        text: rule.attrs.text || null,
        formulas,
        asts: formulas.map((f) => { try { return parseFormula(f); } catch (e) { return null; } }),
      });
    }
  }
  sh.cfs.sort((a, b) => a.priority - b.priority);

  // hyperlinks
  sh.hyperlinks = [];
  const hl = root.child('hyperlinks');
  if (hl) {
    for (const h of hl.all('hyperlink')) {
      const rel = h.attrs['r:id'] ? sh.rels.get(h.attrs['r:id']) : null;
      sh.hyperlinks.push({
        ranges: parseSqref(h.attrs.ref || ''),
        url: rel ? rel.target : null,
        location: h.attrs.location || null,
        tooltip: h.attrs.tooltip || null,
      });
    }
  }

  // ignored errors (only to know where Excel suppresses green triangles; not drawn)
  sh.drawings = [];
  sh.comments = new Map();
  for (const rel of sh.rels.values()) {
    if (rel.type === REL_NS_OFFDOC + 'drawing' && rel.path) {
      const dx = wb.zip.text(rel.path);
      if (!dx) continue;
      const droot = parseXml(dx);
      const drels = readRels(wb.zip, rel.path);
      for (const anc of droot.children) {
        if (anc.local !== 'twoCellAnchor' && anc.local !== 'oneCellAnchor') continue;
        const from = parseAnchorPoint(anc.child('from'));
        const to = parseAnchorPoint(anc.child('to'));
        const gf = anc.child('graphicFrame');
        if (gf) {
          const cref = gf.find('chart')[0];
          const rid = cref && (cref.attrs['r:id']);
          const crel = rid && drels.get(rid);
          if (crel && crel.path) {
            const cx = wb.zip.text(crel.path);
            sh.drawings.push({ kind: 'chart', from, to, path: crel.path, chart: cx ? parseChart(cx) : null });
          }
          continue;
        }
        const spn = anc.child('sp');
        if (spn) {
          const txt = spn.find('t').map((t) => t.text).join('');
          const fill = spn.find('srgbClr')[0];
          sh.drawings.push({ kind: 'shape', from, to, text: txt, fill: fill ? fill.attrs.val : null });
        }
      }
    }
    if (rel.type === REL_NS_OFFDOC + 'comments' && rel.path) {
      const cx = wb.zip.text(rel.path);
      if (!cx) continue;
      const croot = parseXml(cx);
      const authors = (croot.child('authors') || { all: () => [] }).all('author').map((a) => a.text);
      const list = croot.child('commentList');
      if (!list) continue;
      for (const c of list.all('comment')) {
        const a = parseAddr(c.attrs.ref);
        if (!a) continue;
        const rich = c.child('text') ? parseRich(c.child('text')) : { text: '', runs: null };
        sh.comments.set(key(a.r, a.c), { author: authors[+c.attrs.authorId] || '', text: rich.text, runs: rich.runs, visible: false, anchor: null });
      }
    }
  }
  for (const rel of sh.rels.values()) {
    if (rel.type === REL_NS_OFFDOC + 'vmlDrawing' && rel.path) {
      const notes = parseVmlNotes(wb.zip.text(rel.path));
      for (const [k, info] of notes) {
        const cm = sh.comments.get(k);
        if (cm) { cm.visible = info.visible; cm.anchor = info.anchor; }
      }
    }
  }
}

export function loadWorkbook(bytes) {
  const zip = readZip(bytes);
  const wb = { zip, sheets: [], names: [], bytes };
  // locate workbook part via root rels
  const rootRels = readRels(zip, '');
  let wbPath = 'xl/workbook.xml';
  for (const r of rootRels.values()) {
    if (r.type.endsWith('/officeDocument')) wbPath = r.path;
  }
  wb.path = wbPath;
  wb.xml = zip.text(wbPath);
  const wroot = parseXml(wb.xml);
  wb.root = wroot;
  wb.rels = readRels(zip, wbPath);
  const findRel = (suffix) => {
    for (const r of wb.rels.values()) if (r.type.endsWith(suffix)) return r;
    return null;
  };
  const sstRel = findRel('/sharedStrings');
  const stRel = findRel('/styles');
  const thRel = findRel('/theme');
  const ccRel = findRel('/calcChain');
  wb.sstPath = sstRel ? sstRel.path : null;
  wb.stylesPath = stRel ? stRel.path : null;
  wb.calcChainPath = ccRel ? ccRel.path : null;
  wb.sst = parseSST(wb.sstPath ? zip.text(wb.sstPath) : null);
  wb.styles = parseStyles(zip.text(wb.stylesPath));
  wb.theme = parseTheme(thRel ? zip.text(thRel.path) : null);
  const wbPr = wroot.child('workbookPr');
  wb.date1904 = wbPr ? attrBool(wbPr.attrs.date1904) : false;
  const bv = wroot.child('bookViews');
  const wv = bv && bv.child('workbookView');
  wb.activeTab = wv && wv.attrs.activeTab ? +wv.attrs.activeTab : 0;

  const sheetsNode = wroot.child('sheets');
  let idx = 0;
  for (const s of sheetsNode.all('sheet')) {
    const rel = wb.rels.get(s.attrs['r:id']);
    const sh = {
      index: idx++,
      name: s.attrs.name,
      sheetId: +s.attrs.sheetId,
      rId: s.attrs['r:id'],
      state: s.attrs.state || 'visible',
      path: rel ? rel.path : null,
    };
    sh.isChartsheet = rel && rel.type.endsWith('/chartsheet');
    wb.sheets.push(sh);
  }
  for (const sh of wb.sheets) {
    if (!sh.path || sh.isChartsheet) { sh.cells = new Map(); sh.rows = new Map(); sh.cols = []; sh.merges = []; sh.mergeAt = new Map(); sh.dvs = []; sh.cfs = []; sh.drawings = []; sh.comments = new Map(); sh.hyperlinks = []; sh.protectedRanges = []; sh.view = { showGridLines: true, showRowColHeaders: true, zoomScale: 100 }; continue; }
    sh.xml = zip.text(sh.path);
    sh.rels = readRels(zip, sh.path);
    parseSheet(wb, sh);
  }
  const dn = wroot.child('definedNames');
  if (dn) {
    for (const n of dn.all('definedName')) {
      const text = n.text;
      let ast = null;
      try { ast = parseFormula(text); } catch (e) { ast = null; }
      wb.names.push({
        name: n.attrs.name,
        localSheetId: n.attrs.localSheetId !== undefined ? +n.attrs.localSheetId : null,
        hidden: attrBool(n.attrs.hidden),
        text, ast,
      });
    }
  }
  wb.sheetByName = new Map(wb.sheets.map((s) => [s.name.toLowerCase(), s]));
  return wb;
}

export { colToNum };
