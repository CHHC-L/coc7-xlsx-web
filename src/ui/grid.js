// Worksheet view: Excel-like rendering + interaction for one sheet at a time.

import { key, addr, inRange } from '../core/refs.js';
import { formatValue } from '../core/numfmt.js';
import { XErr } from '../core/values.js';
import { resolveColor } from '../core/styles.js';
import { findDV, listItems, cfFormat } from '../core/rules.js';
import { fontCss, fillColor, borderSpec, strongerBorder, fontFamily } from './cellstyle.js';
import { renderChart } from './chart.js';

const MDW = 8; // max digit width of the default font (宋体 11pt) at 96 dpi
const EMU = 9525;

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function colPx(w) {
  return Math.trunc(((256 * w + Math.trunc(128 / MDW)) / 256) * MDW);
}

function bsearch(arr, v, n) {
  // largest i in [1,n] with arr[i] <= v
  let lo = 1, hi = n;
  if (v < arr[1]) return 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (arr[mid] <= v) lo = mid; else hi = mid - 1;
  }
  return lo;
}

export class Grid {
  constructor(host, app) {
    this.app = app;
    this.host = host;
    host.innerHTML = `
      <div class="scroller">
       <div class="zwrap">
        <div class="hdr hdr-top"></div>
        <div class="hdr hdr-left"></div>
        <div class="hdr hdr-corner"></div>
        <div class="canvas">
          <svg class="static" xmlns="http://www.w3.org/2000/svg"></svg>
          <div class="dyn"></div>
          <div class="draw"></div>
          <div class="ovl">
            <div class="sel"></div>
            <div class="dd-btn" title="下拉列表"><svg viewBox="0 0 10 10"><path d="M2 3.5 L5 6.5 L8 3.5" fill="none" stroke="currentColor" stroke-width="1.4"/></svg></div>
            <textarea class="editor" spellcheck="false" autocomplete="off" autocapitalize="off"></textarea>
            <div class="prompt"></div>
            <div class="ddlist" role="listbox"></div>
          </div>
        </div>
       </div>
      </div>
      <div class="note"></div>`;
    this.scroller = host.querySelector('.scroller');
    this.canvas = host.querySelector('.canvas');
    this.zwrap = host.querySelector('.zwrap');
    this.hdrTop = host.querySelector('.hdr-top');
    this.hdrLeft = host.querySelector('.hdr-left');
    this.hdrCorner = host.querySelector('.hdr-corner');
    this.hw = 0; this.hh = 0;
    this.svg = host.querySelector('svg.static');
    this.dyn = host.querySelector('.dyn');
    this.drawLayer = host.querySelector('.draw');
    this.selEl = host.querySelector('.sel');
    this.ddBtn = host.querySelector('.dd-btn');
    this.editor = host.querySelector('.editor');
    this.promptEl = host.querySelector('.prompt');
    this.ddList = host.querySelector('.ddlist');
    this.noteEl = host.querySelector('.note');
    this.si = -1;
    this.sel = null; // {r, c}
    this.editing = false;
    this.editMode = 'enter'; // 'enter' (typed over) | 'edit' (F2 / double click)
    this.composing = false;
    this.viewState = new Map(); // si -> {sel, scrollLeft, scrollTop}
    this.zoom = 1;
    this.styleCache = new Map();
    this.bind();
  }

  get doc() { return this.app.doc; }
  get sh() { return this.doc.wb.sheets[this.si]; }

  // ---- styles ----------------------------------------------------------------------
  info(s) {
    let inf = this.styleCache.get(s);
    if (inf) return inf;
    const st = this.doc.wb.styles;
    const theme = this.doc.wb.theme;
    const xf = st.cellXfs[s] || st.cellXfs[0];
    const font = st.fonts[xf.fontId] || st.fonts[0];
    const al = xf.alignment || {};
    const b = st.borders[xf.borderId] || {};
    inf = {
      xf,
      font,
      fontCss: fontCss(font, theme, st.indexed),
      fill: fillColor(st.fills[xf.fillId], theme, st.indexed),
      border: {
        left: borderSpec(b.left, theme, st.indexed), right: borderSpec(b.right, theme, st.indexed),
        top: borderSpec(b.top, theme, st.indexed), bottom: borderSpec(b.bottom, theme, st.indexed),
        diagonal: b.diagonal && (b.diagonalUp || b.diagonalDown) ? { ...borderSpec(b.diagonal, theme, st.indexed), up: b.diagonalUp, down: b.diagonalDown } : null,
      },
      h: al.horizontal || 'general',
      v: al.vertical || 'bottom',
      wrap: !!al.wrapText,
      shrink: !!al.shrinkToFit,
      indent: al.indent || 0,
      rot: al.textRotation || 0,
      numFmt: st.numFmtCode(xf.numFmtId),
    };
    inf.visual = !!(inf.fill || inf.border.left || inf.border.right || inf.border.top || inf.border.bottom || inf.border.diagonal);
    this.styleCache.set(s, inf);
    return inf;
  }

  resetStyles() { this.styleCache.clear(); }

  // ---- geometry --------------------------------------------------------------------
  buildGeometry() {
    const sh = this.sh;
    const doc = this.doc;
    let defW;
    if (sh.defaultColWidth !== null && sh.defaultColWidth !== undefined) defW = colPx(sh.defaultColWidth);
    else defW = colPx(Math.trunc(((sh.baseColWidth * MDW + 5) / MDW) * 256) / 256);
    const defH = Math.round((sh.defaultRowHeight || 15) * 4 / 3);
    // extents: content + visual formats + merges + drawings
    let maxR = 1, maxC = 1;
    for (const cell of sh.cells.values()) {
      if ((cell.value !== null && cell.value !== '') || this.info(cell.s).visual) {
        if (cell.r > maxR) maxR = cell.r;
        if (cell.c > maxC) maxC = cell.c;
      }
    }
    for (const g of sh.merges) { maxR = Math.max(maxR, g.r2); maxC = Math.max(maxC, g.c2); }
    for (const d of sh.drawings) { if (d.to) { maxR = Math.max(maxR, d.to.row + 1); maxC = Math.max(maxC, d.to.col + 1); } }
    for (const [k] of sh.comments) { const r = Math.floor(k / 16384) + 1, c = (k % 16384) + 1; maxR = Math.max(maxR, r + 4); maxC = Math.max(maxC, c + 4); }
    for (const [, row] of sh.rows) if (row.customFormat && row.s !== null && this.info(row.s).visual) maxR = Math.max(maxR, row.r);
    maxR = Math.min(maxR + 6, 1048576);
    maxC = Math.min(maxC + 3, 16384);
    this.maxR = maxR; this.maxC = maxC;
    const cx = new Float64Array(maxC + 2);
    for (let c = 1; c <= maxC; c++) {
      let w = defW, hidden = false;
      for (const col of sh.cols) {
        if (c >= col.min && c <= col.max) { if (col.width !== null) w = colPx(col.width); hidden = col.hidden; break; }
      }
      cx[c + 1] = cx[c] + (hidden ? 0 : w);
    }
    const ry = new Float64Array(maxR + 2);
    for (let r = 1; r <= maxR; r++) {
      const row = sh.rows.get(r);
      let h = defH;
      let hidden = sh.zeroHeight && !(row && row.ht !== null);
      if (row) {
        if (row.ht !== null) h = Math.round(row.ht * 4 / 3);
        if (row.hidden) hidden = true;
      }
      ry[r + 1] = ry[r] + (hidden ? 0 : h);
    }
    this.colX = cx; this.rowY = ry;
    this.W = cx[maxC + 1]; this.H = ry[maxR + 1];
  }

  x(c) { return this.colX[Math.min(c, this.maxC + 1)]; }
  y(r) { return this.rowY[Math.min(r, this.maxR + 1)]; }

  // merged area containing (r,c) or the cell itself
  area(r, c) {
    const g = this.sh.mergeAt.get(key(r, c));
    return g || { r1: r, c1: c, r2: r, c2: c };
  }

  rect(g) {
    return { x: this.x(g.c1), y: this.y(g.r1), w: this.x(g.c2 + 1) - this.x(g.c1), h: this.y(g.r2 + 1) - this.y(g.r1) };
  }

  hit(px, py) {
    const c = bsearch(this.colX, px, this.maxC);
    const r = bsearch(this.rowY, py, this.maxR);
    return { r, c };
  }

  // ---- sheet switching -------------------------------------------------------------
  setSheet(si) {
    if (this.si >= 0) this.saveViewState();
    this.closeDropdown();
    this.cancelEdit();
    this.si = si;
    this.buildGeometry();
    const z = this.app.zoomFor(si);
    this.zoom = z;
    this.canvas.style.width = this.W + 'px';
    this.canvas.style.height = this.H + 'px';
    this.layoutHeaders();
    this.zwrap.style.zoom = z;
    this.renderStatic();
    this.renderDynamic();
    const vs = this.viewState.get(si);
    if (vs) {
      this.sel = vs.sel;
      this.scroller.scrollLeft = vs.scrollLeft;
      this.scroller.scrollTop = vs.scrollTop;
    } else {
      this.sel = this.initialSelection();
      this.scroller.scrollLeft = 0;
      this.scroller.scrollTop = 0;
      const tl = this.sh.view.topLeftCell;
      const m = tl && /^([A-Z]+)(\d+)$/.exec(tl);
      if (m) {
        let c = 0;
        for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
        const r = +m[2];
        if (r <= this.maxR && c <= this.maxC) {
          const f0 = this.frozenInfo();
          this.scroller.scrollLeft = (f0 && f0.cols ? 0 : this.x(c)) * z;
          this.scroller.scrollTop = (f0 && f0.rows ? 0 : this.y(r)) * z;
        }
      }
    }
    this.updateSelection(false);
  }

  setZoom(z) {
    this.saveViewState();
    this.zoom = z;
    this.zwrap.style.zoom = z;
    this.updateSelection(false);
  }

  saveViewState() {
    this.viewState.set(this.si, { sel: this.sel, scrollLeft: this.scroller.scrollLeft, scrollTop: this.scroller.scrollTop });
  }

  initialSelection() {
    const sh = this.sh;
    const doc = this.doc;
    const sel = sh.view.selection && sh.view.selection.activeCell;
    if (sel) {
      const m = /^([A-Z]+)(\d+)$/.exec(sel);
      if (m) {
        let c = 0;
        for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
        const r = +m[2];
        if (doc.canSelect(this.si, r, c)) { const g = this.area(r, c); return { r: g.r1, c: g.c1 }; }
      }
    }
    const first = this.nextSelectable(1, 0, 1);
    return first || null;
  }

  // ---- row / column headers -----------------------------------------------------------
  layoutHeaders() {
    const show = this.sh.view.showRowColHeaders;
    this.hh = show ? 20 : 0;
    this.hw = show ? Math.max(28, String(this.maxR).length * 8 + 12) : 0;
    const { hw, hh, W, H } = this;
    Object.assign(this.zwrap.style, { width: W + hw + 'px', height: H + hh + 'px' });
    Object.assign(this.canvas.style, { left: hw + 'px', top: hh + 'px' });
    for (const el of [this.hdrTop, this.hdrLeft, this.hdrCorner]) el.style.display = show ? 'block' : 'none';
    if (!show) { this.hdrTop.innerHTML = this.hdrLeft.innerHTML = ''; return; }
    Object.assign(this.hdrTop.style, { width: W + 'px', height: hh + 'px', marginLeft: hw + 'px', marginBottom: -hh + 'px' });
    Object.assign(this.hdrLeft.style, { width: hw + 'px', height: H + 'px', marginTop: hh + 'px', marginBottom: -(H + hh) + 'px' });
    Object.assign(this.hdrCorner.style, { width: hw + 'px', height: hh + 'px', marginBottom: -hh + 'px' });
    let t = '';
    for (let c = 1; c <= this.maxC; c++) {
      const w = this.x(c + 1) - this.x(c);
      if (w <= 0) continue;
      t += `<div class="hc" data-c="${c}" style="left:${this.x(c)}px;width:${w}px">${numToColLabel(c)}</div>`;
    }
    const f = this.frozenInfo();
    if (f && f.cols) {
      // header cells of frozen columns stay put while scrolling horizontally
      let fz = '';
      for (let c = f.cols.c1; c <= f.cols.c2; c++) {
        const w = this.x(c + 1) - this.x(c);
        if (w > 0) fz += `<div class="hc" data-c="${c}" style="left:${this.x(c) - this.x(f.c0)}px;width:${w}px">${numToColLabel(c)}</div>`;
      }
      t += `<div class="hfrz" style="left:${hw}px;width:${f.Wf}px;height:${hh}px;margin-left:${this.x(f.c0)}px">${fz}</div>`;
    }
    this.hdrTop.innerHTML = t;
    let l = '';
    for (let r = 1; r <= this.maxR; r++) {
      const h = this.y(r + 1) - this.y(r);
      if (h <= 0) continue;
      l += `<div class="hr" data-r="${r}" style="top:${this.y(r)}px;height:${h}px;line-height:${h}px">${r}</div>`;
    }
    if (f && f.rows) {
      let fz = '';
      for (let r = f.rows.r1; r <= f.rows.r2; r++) {
        const h = this.y(r + 1) - this.y(r);
        if (h > 0) fz += `<div class="hr" data-r="${r}" style="top:${this.y(r) - this.y(f.r0)}px;height:${h}px;line-height:${h}px">${r}</div>`;
      }
      l = `<div class="hfrz" style="top:${hh}px;width:${hw}px;height:${f.Hf}px;margin-top:${this.y(f.r0)}px;margin-bottom:${-f.Hf - this.y(f.r0)}px">${fz}</div>` + l;
    }
    this.hdrLeft.innerHTML = l;
  }

  markHeaders(g) {
    if (!this.hh) return;
    for (const el of this.hdrTop.querySelectorAll('.on')) el.classList.remove('on');
    for (const el of this.hdrLeft.querySelectorAll('.on')) el.classList.remove('on');
    if (!g) return;
    for (let c = g.c1; c <= g.c2; c++) { const el = this.hdrTop.querySelector(`[data-c="${c}"]`); if (el) el.classList.add('on'); }
    for (let r = g.r1; r <= g.r2; r++) { const el = this.hdrLeft.querySelector(`[data-r="${r}"]`); if (el) el.classList.add('on'); }
  }

  // ---- static layer (gridlines, fills, borders) -----------------------------------
  renderStatic() {
    const sh = this.sh, doc = this.doc;
    const { maxR, maxC } = this;
    const parts = [];
    // gridlines
    if (sh.view.showGridLines) {
      let d = '';
      for (let c = 1; c <= maxC + 1; c++) {
        const x = this.x(c) - 0.5;
        let run = -1;
        for (let r = 1; r <= maxR + 1; r++) {
          const inside = r <= maxR && c > 1 && c <= maxC && (() => { const g = sh.mergeAt.get(key(r, c)); return g && g.c1 < c; })();
          if (!inside && r <= maxR) { if (run < 0) run = r; }
          else if (run >= 0) { d += `M${x} ${this.y(run)}V${this.y(r)}`; run = -1; }
        }
        if (run >= 0) d += `M${x} ${this.y(run)}V${this.y(maxR + 1)}`;
      }
      for (let r = 1; r <= maxR + 1; r++) {
        const y = this.y(r) - 0.5;
        let run = -1;
        for (let c = 1; c <= maxC + 1; c++) {
          const inside = c <= maxC && r > 1 && r <= maxR && (() => { const g = sh.mergeAt.get(key(r, c)); return g && g.r1 < r; })();
          if (!inside && c <= maxC) { if (run < 0) run = c; }
          else if (run >= 0) { d += `M${this.x(run)} ${y}H${this.x(c)}`; run = -1; }
        }
        if (run >= 0) d += `M${this.x(run)} ${y}H${this.x(maxC + 1)}`;
      }
      parts.push(`<path d="${d}" stroke="#d9d9d9" stroke-width="1" fill="none" shape-rendering="crispEdges"/>`);
    }
    // fills
    const fills = new Map();
    const styleAt = (r, c) => doc.styleOf(this.si, r, c);
    for (let r = 1; r <= maxR; r++) {
      if (this.rowY[r + 1] === this.rowY[r]) continue;
      for (let c = 1; c <= maxC; c++) {
        const g = sh.mergeAt.get(key(r, c));
        if (g && (g.r1 !== r || g.c1 !== c)) continue;
        const inf = this.info(styleAt(r, c));
        if (!inf.fill) continue;
        const rc = this.rect(g || { r1: r, c1: c, r2: r, c2: c });
        if (rc.w <= 0 || rc.h <= 0) continue;
        const k = inf.fill.gradient || inf.fill.solid;
        if (inf.fill.gradient) {
          parts.push(`<foreignObject x="${rc.x}" y="${rc.y}" width="${rc.w}" height="${rc.h}"><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;background:${inf.fill.gradient}"></div></foreignObject>`);
          continue;
        }
        if (!fills.has(k)) fills.set(k, '');
        fills.set(k, fills.get(k) + `M${rc.x} ${rc.y}h${rc.w}v${rc.h}h${-rc.w}z`);
      }
    }
    for (const [col, d] of fills) parts.push(`<path d="${d}" fill="${col}" shape-rendering="crispEdges"/>`);
    // borders: resolve shared edges
    const H = new Map(); // key(r,c) -> spec : top edge of row r at column c
    const V = new Map(); // key(r,c) -> spec : left edge of column c at row r
    const put = (map, k, spec) => { if (spec) map.set(k, strongerBorder(map.get(k), spec)); };
    const diags = [];
    for (let r = 1; r <= maxR; r++) {
      for (let c = 1; c <= maxC; c++) {
        const inf = this.info(styleAt(r, c));
        const b = inf.border;
        if (!b.left && !b.right && !b.top && !b.bottom && !b.diagonal) continue;
        const g = sh.mergeAt.get(key(r, c));
        if (g) {
          if (r === g.r1) put(H, key(r, c), b.top);
          if (r === g.r2) put(H, key(r + 1, c), b.bottom);
          if (c === g.c1) put(V, key(r, c), b.left);
          if (c === g.c2) put(V, key(r, c + 1), b.right);
          if (b.diagonal && r === g.r1 && c === g.c1) diags.push([g, b.diagonal]);
        } else {
          put(H, key(r, c), b.top);
          put(H, key(r + 1, c), b.bottom);
          put(V, key(r, c), b.left);
          put(V, key(r, c + 1), b.right);
          if (b.diagonal) diags.push([{ r1: r, c1: c, r2: r, c2: c }, b.diagonal]);
        }
      }
    }
    const groups = new Map();
    const addSeg = (spec, seg) => {
      const k = spec.style + spec.color;
      if (!groups.has(k)) groups.set(k, { spec, d: '' });
      groups.get(k).d += seg;
    };
    for (const [k, spec] of H) {
      const r = Math.floor(k / 16384) + 1, c = (k % 16384) + 1;
      if (c > maxC || r > maxR + 1) continue;
      const x1 = this.x(c), x2 = this.x(c + 1);
      if (x2 <= x1) continue;
      const y = this.y(r);
      if (spec.double) { addSeg(spec, `M${x1 - 1} ${y - 1.5}H${x2 + 1}M${x1 - 1} ${y + 0.5}H${x2 + 1}`); continue; }
      const off = spec.w % 2 ? -0.5 : 0;
      addSeg(spec, `M${x1 - (spec.w > 1 ? 1 : 0)} ${y + off}H${x2 + (spec.w > 1 ? 1 : 0)}`);
    }
    for (const [k, spec] of V) {
      const r = Math.floor(k / 16384) + 1, c = (k % 16384) + 1;
      if (r > maxR || c > maxC + 1) continue;
      const y1 = this.y(r), y2 = this.y(r + 1);
      if (y2 <= y1) continue;
      const x = this.x(c);
      if (spec.double) { addSeg(spec, `M${x - 1.5} ${y1 - 1}V${y2 + 1}M${x + 0.5} ${y1 - 1}V${y2 + 1}`); continue; }
      const off = spec.w % 2 ? -0.5 : 0;
      addSeg(spec, `M${x + off} ${y1 - (spec.w > 1 ? 1 : 0)}V${y2 + (spec.w > 1 ? 1 : 0)}`);
    }
    for (const [g, spec] of diags) {
      const rc = this.rect(g);
      let d = '';
      if (spec.down) d += `M${rc.x} ${rc.y}L${rc.x + rc.w} ${rc.y + rc.h}`;
      if (spec.up) d += `M${rc.x} ${rc.y + rc.h}L${rc.x + rc.w} ${rc.y}`;
      addSeg({ ...spec, style: spec.style + 'diag' }, d);
    }
    for (const { spec, d } of groups.values()) {
      const w = spec.double ? 1 : spec.w;
      parts.push(`<path d="${d}" stroke="${spec.color}" stroke-width="${w}" fill="none"${spec.dash ? ` stroke-dasharray="${spec.dash}"` : ''} shape-rendering="crispEdges"/>`);
    }
    this.svg.setAttribute('width', this.W);
    this.svg.setAttribute('height', this.H);
    this.svg.setAttribute('viewBox', `0 0 ${this.W} ${this.H}`);
    this.svg.innerHTML = parts.join('');
  }

  // ---- dynamic layer (text, CF, comments, charts) --------------------------------
  cellDisplay(cell, inf, cf) {
    const v = this.app.shownValue(cell);
    if (v === null || v === undefined) return null;
    if (v === 0 && this.sh.view.node && this.sh.view.node.attrs.showZeros === '0') return null;
    const code = cf && cf.numFmt ? cf.numFmt : inf.numFmt;
    if (cell.runs && typeof v === 'string' && !cell.ast) {
      return { rich: cell.runs, color: null, kind: 's' };
    }
    const f = formatValue(v, code, this.doc.wb.styles.indexed);
    const kind = typeof v === 'number' ? 'n' : typeof v === 'string' ? 's' : 'b';
    return { parts: f.parts, color: f.color, kind, text: f.text };
  }

  renderDynamic() {
    const sh = this.sh, doc = this.doc;
    const st = doc.wb.styles, theme = doc.wb.theme;
    const out = [];
    const occupied = new Set();
    const items = [];
    for (const cell of sh.cells.values()) {
      if (cell.r > this.maxR || cell.c > this.maxC) continue;
      const g = sh.mergeAt.get(key(cell.r, cell.c));
      if (g && (g.r1 !== cell.r || g.c1 !== cell.c)) continue;
      const inf = this.info(cell.s);
      const cf = sh.cfs.length ? cfFormat(doc, this.si, cell.r, cell.c, cell.value) : null;
      const disp = this.cellDisplay(cell, inf, cf);
      if (cf && (cf.fill || cf.border)) {
        const rc = this.rect(g || { r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c });
        const fc = cf.fill ? fillColor(cf.fill, theme, st.indexed, true) : null;
        let bcss = '';
        if (cf.border) {
          for (const side of ['left', 'right', 'top', 'bottom']) {
            const b = borderSpec(cf.border[side], theme, st.indexed);
            if (b) bcss += `border-${side}:${b.w}px ${b.double ? 'double' : b.dash ? (/^1,/.test(b.dash) || b.style === 'dotted' ? 'dotted' : 'dashed') : 'solid'} ${b.color};`;
          }
        }
        if (fc || bcss) out.push(`<div class="cf" style="left:${rc.x - (bcss ? 1 : 0)}px;top:${rc.y - (bcss ? 1 : 0)}px;width:${rc.w + (bcss ? 1 : 0)}px;height:${rc.h + (bcss ? 1 : 0)}px;${fc ? `background:${fc.gradient || fc.solid};` : ''}${bcss}"></div>`);
      }
      if (!disp) continue;
      if (disp.text === '' && !disp.rich) continue;
      occupied.add(key(cell.r, cell.c));
      items.push({ cell, g, inf, cf, disp });
    }
    for (const it of items) out.push(this.cellHtml(it, occupied));
    // comment indicators
    for (const [k, cm] of sh.comments) {
      const r = Math.floor(k / 16384) + 1, c = (k % 16384) + 1;
      if (r > this.maxR || c > this.maxC) continue;
      const g = this.area(r, c);
      const rc = this.rect(g);
      out.push(`<div class="cmt" data-k="${k}" style="left:${rc.x + rc.w - 7}px;top:${rc.y}px"></div>`);
      if (cm.visible) out.push(this.noteHtml(cm, rc, true));
    }
    // hyperlinks: make the whole area clickable
    for (const h of sh.hyperlinks) {
      for (const g0 of h.ranges) {
        const g = this.area(g0.r1, g0.c1);
        const rc = this.rect({ r1: g.r1, c1: g.c1, r2: Math.max(g.r2, g0.r2), c2: Math.max(g.c2, g0.c2) });
        out.push(`<a class="hl" href="${esc(h.url || '#')}" target="_blank" rel="noopener" title="${esc(h.tooltip || h.url || '')}" style="left:${rc.x}px;top:${rc.y}px;width:${rc.w}px;height:${rc.h}px"></a>`);
      }
    }
    this.dyn.innerHTML = out.join('');
    this.renderDrawings();
    this.postLayout();
    this.renderFrozen();
  }

  // Frozen panes: sticky bands holding a copy of the frozen rows / columns.
  frozenInfo() {
    const p = this.sh.view.pane;
    if (!p || (!p.xSplit && !p.ySplit)) return null;
    let r0 = 1, c0 = 1;
    const m = this.sh.view.topLeftCell && /^([A-Z]+)(\d+)$/.exec(this.sh.view.topLeftCell);
    if (m) { c0 = 0; for (const ch of m[1]) c0 = c0 * 26 + (ch.charCodeAt(0) - 64); r0 = +m[2]; }
    const rows = p.ySplit ? { r1: r0, r2: r0 + p.ySplit - 1 } : null;
    const cols = p.xSplit ? { c1: c0, c2: c0 + p.xSplit - 1 } : null;
    return {
      r0, c0, rows, cols,
      Hf: rows ? this.y(rows.r2 + 1) - this.y(r0) : 0,
      Wf: cols ? this.x(cols.c2 + 1) - this.x(c0) : 0,
    };
  }

  renderFrozen() {
    for (const el of this.canvas.querySelectorAll('.frz')) el.remove();
    const f = this.frozenInfo();
    this.frozen = f;
    if (!f) return;
    const inner = this.svg.outerHTML + this.dyn.outerHTML + this.drawLayer.outerHTML;
    const band = (cls, w, h, dx, dy) =>
      `<div class="frz ${cls}" style="width:${w}px;height:${h}px;margin-bottom:${-h}px;${cls !== 'frz-left' ? `top:${this.hh}px;` : ''}${cls !== 'frz-top' ? `left:${this.hw}px;` : ''}"><div class="frz-in" style="left:${-dx}px;top:${-dy}px;width:${this.W}px;height:${this.H}px">${inner}</div></div>`;
    let html = '';
    if (f.rows) html += band('frz-top', this.W, f.Hf, 0, this.y(f.r0));
    if (f.cols) html += band('frz-left', f.Wf, this.H, this.x(f.c0), 0);
    if (f.rows && f.cols) html += band('frz-corner', f.Wf, f.Hf, this.x(f.c0), this.y(f.r0));
    this.canvas.insertAdjacentHTML('afterbegin', html);
  }

  cellHtml({ cell, g, inf, cf, disp }, occupied) {
    const sh = this.sh;
    const st = this.doc.wb.styles, theme = this.doc.wb.theme;
    const area = g || { r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c };
    let rc = this.rect(area);
    if (rc.w <= 0 || rc.h <= 0) return '';
    let h = inf.h;
    if (h === 'general') h = disp.kind === 'n' ? 'right' : disp.kind === 'b' ? 'center' : 'left';
    const wrap = inf.wrap || h === 'justify' || h === 'distributed' || inf.v === 'justify' || inf.v === 'distributed';
    // text overflow into empty neighbours (not for merged, wrapped or numeric cells)
    let ox = rc.x, ow = rc.w;
    if (!g && !wrap && !inf.shrink && disp.kind === 's' && inf.rot === 0) {
      const free = (c) => c >= 1 && c <= this.maxC && !occupied.has(key(cell.r, c)) && !sh.mergeAt.get(key(cell.r, c));
      if (h === 'left' || h === 'fill' || h === 'centerContinuous' || h === 'center') {
        let c = cell.c + 1;
        const lim = h === 'center' || h === 'centerContinuous' ? 6 : 40;
        while (c <= this.maxC && free(c) && c - cell.c <= lim) c++;
        ow = this.x(c) - rc.x;
      }
      if (h === 'right' || h === 'center' || h === 'centerContinuous') {
        let c = cell.c - 1;
        const lim = h === 'right' ? 40 : 6;
        while (c >= 1 && free(c) && cell.c - c <= lim) c--;
        const nx = this.x(c + 1);
        ow += rc.x - nx;
        ox = nx;
      }
      if (h === 'center' || h === 'centerContinuous') {
        // keep the text centred on its own cell
        const left = rc.x - ox, right = ox + ow - (rc.x + rc.w);
        const m = Math.min(left, right);
        ox = rc.x - m; ow = rc.w + 2 * m;
      }
    }
    const font = cf && cf.font ? { ...inf.font, ...cf.font } : inf.font;
    let css = cf && cf.font ? fontCss(font, theme, st.indexed) : inf.fontCss;
    if (disp.color) css += `color:#${disp.color};`;
    const va = inf.v === 'top' ? 'flex-start' : inf.v === 'center' || inf.v === 'justify' || inf.v === 'distributed' ? 'center' : 'flex-end';
    const ta = h === 'center' || h === 'centerContinuous' ? 'center' : h === 'right' ? 'right' : h === 'justify' ? 'justify' : h === 'distributed' ? 'justify' : 'left';
    let cls = 'c' + (wrap ? ' w' : '') + (inf.shrink ? ' shrink' : '');
    let inner;
    if (disp.rich) {
      inner = disp.rich.map((run) => {
        const f = run.font ? fontCss({ ...run.font }, theme, st.indexed) : '';
        return `<span style="${f}">${esc(wrap ? run.text : run.text.replace(/\n/g, ''))}</span>`;
      }).join('');
    } else {
      inner = disp.parts.map((p) => (p.t === 'pad' ? `<span class="pad">${esc(p.s)}</span>` : esc(wrap ? p.s : p.s.replace(/\n/g, '')))).join('');
    }
    let tstyle = `text-align:${ta};`;
    if (h === 'distributed') tstyle += 'text-align-last:justify;';
    if (inf.indent) tstyle += h === 'right' ? `padding-right:${inf.indent * 3 * MDW}px;` : `padding-left:${inf.indent * 3 * MDW}px;`;
    if (inf.rot === 255) cls += ' vert';
    else if (inf.rot) {
      const deg = inf.rot <= 90 ? -inf.rot : inf.rot - 90;
      tstyle += `transform:rotate(${deg}deg);`;
      cls += ' rot';
    }
    const just = ta === 'center' ? 'center' : ta === 'right' ? 'flex-end' : 'flex-start';
    return `<div class="${cls}" data-k="${key(cell.r, cell.c)}" style="left:${ox}px;top:${rc.y}px;width:${ow}px;height:${rc.h}px;justify-content:${va};align-items:${wrap ? 'stretch' : just};${css}"><div class="t" style="${tstyle}">${inner}</div></div>`;
  }

  noteHtml(cm, rc, visible) {
    const runs = cm.runs || [{ text: cm.text, font: null }];
    const st = this.doc.wb.styles, theme = this.doc.wb.theme;
    const body = runs.map((r) => `<span style="${r.font ? fontCss(r.font, theme, st.indexed) : ''}">${esc(r.text)}</span>`).join('');
    let w = 160, h = 80;
    if (cm.anchor && cm.anchor.length === 8) {
      const [c1, dx1, r1, dy1, c2, dx2, r2, dy2] = cm.anchor;
      const x1 = this.x(c1 + 1) + dx1, x2 = this.x(c2 + 1) + dx2;
      const y1 = this.y(r1 + 1) + dy1, y2 = this.y(r2 + 1) + dy2;
      w = Math.max(60, x2 - x1); h = Math.max(30, y2 - y1);
      if (visible) return `<div class="notebox vis" style="left:${x1}px;top:${y1}px;width:${w}px;height:${h}px">${body}</div>`;
    }
    return `<div class="notebox vis" style="left:${rc.x + rc.w + 12}px;top:${rc.y}px;width:${w}px;min-height:${h}px">${body}</div>`;
  }

  postLayout() {
    // shrink-to-fit: scale text down until it fits its box
    for (const el of this.dyn.querySelectorAll('.c.shrink')) {
      const t = el.firstChild;
      const avail = el.clientWidth - 4;
      if (t.scrollWidth > avail && avail > 0) {
        const k = Math.max(0.3, avail / t.scrollWidth);
        const fs = parseFloat(getComputedStyle(el).fontSize);
        t.style.fontSize = (fs * k) + 'px';
      }
    }
  }

  renderDrawings() {
    const sh = this.sh;
    const out = [];
    const pos = (p) => ({ x: this.x(p.col + 1) + p.colOff / EMU, y: this.y(p.row + 1) + p.rowOff / EMU });
    for (const d of sh.drawings) {
      if (!d.from || !d.to) continue;
      const a = pos(d.from), b = pos(d.to);
      const w = Math.max(1, b.x - a.x), h = Math.max(1, b.y - a.y);
      if (d.kind === 'chart' && d.chart) {
        out.push(`<div class="chart" style="left:${a.x}px;top:${a.y}px;width:${w}px;height:${h}px">${renderChart(this.doc, d.chart, w, h)}</div>`);
      } else if (d.kind === 'shape') {
        out.push(`<div class="shape" style="left:${a.x}px;top:${a.y}px;width:${w}px;height:${h}px;background:${d.fill ? '#' + d.fill : 'transparent'}">${esc(d.text || '')}</div>`);
      }
    }
    this.drawLayer.innerHTML = out.join('');
  }

  refresh() {
    this.renderDynamic();
    this.updateSelection(false);
  }

  // ---- selection & navigation ---------------------------------------------------------
  selectable(r, c) {
    if (r < 1 || c < 1 || r > this.maxR || c > this.maxC) return false;
    if (this.rowY[r + 1] === this.rowY[r] || this.colX[c + 1] === this.colX[c]) return false;
    return this.doc.canSelect(this.si, r, c);
  }

  nextSelectable(r, c, dir) {
    // reading order scan (dir = +1 / -1), wrapping once
    const total = this.maxR * this.maxC;
    let idx = (r - 1) * this.maxC + (c - 1);
    for (let n = 0; n < total; n++) {
      idx = (idx + dir + total) % total;
      const rr = Math.floor(idx / this.maxC) + 1, cc = (idx % this.maxC) + 1;
      const g = this.area(rr, cc);
      if (g.r1 !== rr || g.c1 !== cc) continue;
      if (this.selectable(rr, cc)) return { r: rr, c: cc };
    }
    return null;
  }

  move(dr, dc) {
    if (!this.sel) { this.sel = this.nextSelectable(1, 0, 1); this.updateSelection(true); return; }
    const g = this.area(this.sel.r, this.sel.c);
    let r = dr > 0 ? g.r2 : dr < 0 ? g.r1 : this.sel.r;
    let c = dc > 0 ? g.c2 : dc < 0 ? g.c1 : this.sel.c;
    for (let i = 0; i < 2000; i++) {
      r += dr; c += dc;
      if (r < 1 || c < 1 || r > this.maxR || c > this.maxC) return;
      const a = this.area(r, c);
      if (this.selectable(a.r1, a.c1) && !(a.r1 === g.r1 && a.c1 === g.c1)) {
        this.sel = { r: a.r1, c: a.c1 };
        this.updateSelection(true);
        return;
      }
    }
  }

  tabMove(dir) {
    if (!this.sel) return this.move(0, 1);
    const sh = this.sh;
    const prot = this.doc.isProtected(this.si) && sh.protection.selectLockedCells;
    if (prot) {
      const n = this.nextSelectable(this.sel.r, this.sel.c, dir);
      if (n) { this.sel = n; this.updateSelection(true); }
    } else this.move(0, dir);
  }

  enterMove(dir) {
    const sh = this.sh;
    const prot = this.doc.isProtected(this.si) && sh.protection.selectLockedCells;
    if (prot) return this.tabMove(dir);
    this.move(dir, 0);
  }

  select(r, c, scroll = true) {
    const g = this.area(r, c);
    if (!this.selectable(g.r1, g.c1)) return false;
    this.sel = { r: g.r1, c: g.c1 };
    this.updateSelection(scroll);
    return true;
  }

  updateSelection(scrollIntoView) {
    this.closeDropdown();
    const sel = this.sel;
    if (!sel) {
      this.selEl.style.display = 'none';
      this.ddBtn.style.display = 'none';
      this.promptEl.style.display = 'none';
      this.app.onSelection(null);
      return;
    }
    const g = this.area(sel.r, sel.c);
    const rc = this.rect(g);
    Object.assign(this.selEl.style, { display: 'block', left: rc.x - 1 + 'px', top: rc.y - 1 + 'px', width: rc.w + 1 + 'px', height: rc.h + 1 + 'px' });
    if (!this.editing) { this.placeEditor(rc, false); this.syncHidden(); }
    // dropdown & prompt
    const dv = findDV(this.doc, this.si, sel.r, sel.c);
    this.curDV = dv;
    const canEdit = this.doc.canEdit(this.si, sel.r, sel.c);
    if (dv && dv.type === 'list' && !dv.hideDropDown && canEdit) {
      const bh = Math.min(rc.h, 20);
      Object.assign(this.ddBtn.style, { display: 'flex', left: rc.x + rc.w + 1 + 'px', top: rc.y + rc.h - bh + 'px', height: bh + 'px' });
    } else this.ddBtn.style.display = 'none';
    if (dv && dv.showInputMessage && (dv.prompt || dv.promptTitle)) {
      this.promptEl.innerHTML = (dv.promptTitle ? `<b>${esc(dv.promptTitle)}</b>` : '') + `<div>${esc(dv.prompt)}</div>`;
      Object.assign(this.promptEl.style, { display: 'block', left: rc.x + Math.min(rc.w, 48) + 'px', top: rc.y + rc.h + 6 + 'px' });
    } else this.promptEl.style.display = 'none';
    if (scrollIntoView) this.scrollTo(rc);
    this.markHeaders(g);
    this.app.onSelection(sel);
  }

  scrollTo(rc) {
    const z = this.zoom;
    const s = this.scroller;
    const pad = 8;
    const f = this.frozen;
    const offX = (this.hw + (f && f.cols ? f.Wf : 0)) * z, offY = (this.hh + (f && f.rows ? f.Hf : 0)) * z;
    const inFrozenR = f && f.rows && rc.y < this.y(f.rows.r2 + 1) && rc.y >= this.y(f.r0);
    const inFrozenC = f && f.cols && rc.x < this.x(f.cols.c2 + 1) && rc.x >= this.x(f.c0);
    const left = (rc.x + this.hw) * z, top = (rc.y + this.hh) * z, right = (rc.x + rc.w + this.hw) * z, bottom = (rc.y + rc.h + this.hh) * z;
    if (!inFrozenC) {
      if (left - offX + this.hw * z < s.scrollLeft + this.hw * z) s.scrollLeft = Math.max(0, left - offX - pad);
      else if (right > s.scrollLeft + s.clientWidth) s.scrollLeft = right - s.clientWidth + pad;
    }
    if (!inFrozenR) {
      if (top - offY + this.hh * z < s.scrollTop + this.hh * z) s.scrollTop = Math.max(0, top - offY - pad);
      else if (bottom > s.scrollTop + s.clientHeight) s.scrollTop = bottom - s.clientHeight + pad;
    }
  }

  // ---- editor ------------------------------------------------------------------------
  placeEditor(rc, visible) {
    const ed = this.editor;
    const sel = this.sel;
    const inf = this.info(this.doc.styleOf(this.si, sel.r, sel.c));
    ed.style.left = rc.x + 'px';
    ed.style.top = rc.y + 'px';
    ed.style.minWidth = rc.w + 'px';
    ed.style.width = rc.w + 'px';
    ed.style.minHeight = rc.h + 'px';
    ed.style.height = rc.h + 'px';
    const f = inf.font || {};
    ed.style.fontFamily = fontFamily(f.name, this.doc.wb.theme, f.scheme);
    ed.style.fontSize = (f.sz || 11) + 'pt';
    ed.style.fontWeight = f.b ? 'bold' : 'normal';
    ed.style.fontStyle = f.i ? 'italic' : 'normal';
    ed.classList.toggle('on', visible);
    ed.dataset.wrap = inf.wrap ? '1' : '';
  }

  // While not editing, the (invisible) editor holds the cell's text fully
  // selected: native copy works everywhere and typing / IME replaces it.
  syncHidden() {
    if (this.editing || !this.sel) return;
    const ed = this.editor;
    ed.value = this.app.copyText(this.si, this.sel.r, this.sel.c);
    if (document.activeElement === ed) ed.select();
  }

  focus() {
    try { this.editor.focus({ preventScroll: true }); } catch (e) { this.editor.focus(); }
    if (!this.editing) this.editor.select();
  }

  rawText(cell) {
    if (!cell) return '';
    if (cell.ast) {
      const xf = this.doc.xf(cell.s);
      if (xf && xf.hidden && this.doc.isProtected(this.si)) return '';
      return '=' + this.app.formulaText(this.si, cell);
    }
    const v = cell.value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (v instanceof XErr) return v.code;
    if (typeof v === 'number') {
      const inf = this.info(cell.s);
      if (/[yd]|h:|上午/i.test(inf.numFmt) && !/^\[/.test(inf.numFmt)) return formatValue(v, inf.numFmt).text;
      if (/%/.test(inf.numFmt)) return String(+(v * 100).toPrecision(15)) + '%';
      return String(+v.toPrecision(15));
    }
    if (cell.xfQuote) return "'" + v;
    const xf = this.doc.xf(cell.s);
    return (xf && xf.quotePrefix ? "'" : '') + v;
  }

  beginEdit(mode, initial) {
    if (!this.sel) return false;
    const { r, c } = this.sel;
    if (!this.doc.canEdit(this.si, r, c)) {
      this.app.protectedAlert();
      return false;
    }
    this.editing = true;
    this.editMode = mode;
    this.closeDropdown();
    const g = this.area(r, c);
    const rc = this.rect(g);
    this.placeEditor(rc, true);
    const ed = this.editor;
    if (initial !== undefined) ed.value = initial;
    this.originalText = this.rawText(this.doc.cell(this.si, r, c));
    this.autosize();
    try { ed.focus({ preventScroll: true }); } catch (e) { ed.focus(); }
    if (mode === 'edit' || initial !== undefined) ed.setSelectionRange(ed.value.length, ed.value.length);
    this.app.onEditing(true, ed.value);
    return true;
  }

  autosize() {
    const ed = this.editor;
    if (!this.editing) return;
    const wrap = ed.dataset.wrap === '1';
    ed.style.height = 'auto';
    if (!wrap) {
      ed.style.whiteSpace = 'pre';
      ed.style.width = 'auto';
      const minW = parseFloat(ed.style.minWidth) || 0;
      ed.style.width = Math.max(minW, ed.scrollWidth + 4) + 'px';
    } else ed.style.whiteSpace = 'pre-wrap';
    const minH = parseFloat(ed.style.minHeight) || 0;
    ed.style.height = Math.max(minH, ed.scrollHeight) + 'px';
  }

  cancelEdit() {
    if (!this.editing) return;
    this.editing = false;
    this.editor.value = '';
    this.editor.classList.remove('on');
    this.app.onEditing(false);
    if (this.sel) this.updateSelection(false);
  }

  async commitEdit(after) {
    if (!this.editing) return true;
    const text = this.editor.value;
    const { r, c } = this.sel;
    const res = await this.app.commitInput(this.si, r, c, text, { fromEditor: true });
    if (res === 'retry') {
      this.focus();
      this.editor.select();
      return false;
    }
    this.editing = false;
    this.editor.value = '';
    this.editor.classList.remove('on');
    this.app.onEditing(false);
    if (after) after();
    else this.updateSelection(false);
    this.syncHidden();
    this.focus();
    return true;
  }

  // ---- dropdown ------------------------------------------------------------------------
  openDropdown() {
    if (!this.sel || !this.curDV) return;
    const { r, c } = this.sel;
    if (!this.doc.canEdit(this.si, r, c)) return;
    const items = listItems(this.doc, this.si, this.curDV, r, c) || [];
    const cell = this.doc.cell(this.si, r, c);
    const cur = cell ? this.rawText(cell) : '';
    const g = this.area(r, c);
    const rc = this.rect(g);
    const list = this.ddList;
    list.innerHTML = items.map((it, i) => `<div class="opt${it.text === cur ? ' cur' : ''}" data-i="${i}">${esc(it.text) || '&nbsp;'}</div>`).join('');
    list._items = items;
    Object.assign(list.style, { display: 'block', left: rc.x + 'px', top: rc.y + rc.h + 2 + 'px', minWidth: Math.max(rc.w + 18, 60) + 'px' });
    this.ddIndex = Math.max(0, items.findIndex((it) => it.text === cur));
    this.highlightDD();
  }

  highlightDD() {
    const opts = this.ddList.querySelectorAll('.opt');
    opts.forEach((o, i) => o.classList.toggle('hi', i === this.ddIndex));
    const el = opts[this.ddIndex];
    if (el) el.scrollIntoView({ block: 'nearest' });
  }

  closeDropdown() {
    if (this.ddList.style.display === 'block') this.ddList.style.display = 'none';
  }

  async chooseDD(i) {
    const items = this.ddList._items || [];
    const it = items[i];
    this.closeDropdown();
    if (!it || !this.sel) return;
    const { r, c } = this.sel;
    await this.app.commitInput(this.si, r, c, it.text, { fromList: true });
    this.updateSelection(false);
    this.focus();
  }

  // ---- notes ----------------------------------------------------------------------------
  showNote(k, anchorEl) {
    const cm = this.sh.comments.get(k);
    if (!cm || cm.visible) return;
    const r = Math.floor(k / 16384) + 1, c = (k % 16384) + 1;
    const rc = this.rect(this.area(r, c));
    const html = this.noteHtml(cm, rc, false);
    const tmp = document.createElement('div');
    tmp.innerHTML = html;
    const box = tmp.firstChild;
    const z = this.zoom;
    const cr = this.canvas.getBoundingClientRect();
    const hr = this.host.getBoundingClientRect();
    this.noteEl.innerHTML = box.innerHTML;
    Object.assign(this.noteEl.style, {
      display: 'block',
      left: (cr.left - hr.left + (rc.x + rc.w + 12) * z) + 'px',
      top: (cr.top - hr.top + rc.y * z) + 'px',
      width: box.style.width,
      minHeight: box.style.minHeight || '',
    });
    void anchorEl;
  }

  hideNote() { this.noteEl.style.display = 'none'; }

  // ---- events --------------------------------------------------------------------------
  pointToCell(ev) {
    const cr = this.canvas.getBoundingClientRect();
    const sr = this.scroller.getBoundingClientRect();
    const z = this.zoom;
    let px = (ev.clientX - cr.left) / z, py = (ev.clientY - cr.top) / z;
    const f = this.frozen;
    if (f) {
      const vx = (ev.clientX - sr.left) / z - this.hw, vy = (ev.clientY - sr.top) / z - this.hh;
      if (f.rows && vy < f.Hf) py = this.y(f.r0) + vy;
      if (f.cols && vx < f.Wf) px = this.x(f.c0) + vx;
    }
    if (px < 0 || py < 0 || px > this.W || py > this.H) return null;
    return this.hit(px, py);
  }

  bind() {
    const ed = this.editor;
    this.canvas.addEventListener('mousedown', async (ev) => {
      if (ev.target.closest('.ddlist') || ev.target.closest('.dd-btn') || ev.target.closest('.editor.on') || ev.target.closest('a.hl')) return;
      if (ev.button !== 0) return;
      const p = this.pointToCell(ev);
      if (!p) return;
      if (this.editing) {
        const ok = await this.commitEdit(() => {});
        if (!ok) { ev.preventDefault(); return; }
      }
      const g = this.area(p.r, p.c);
      if (this.select(g.r1, g.c1, false)) {
        ev.preventDefault();
        this.focus();
      }
    });
    this.canvas.addEventListener('dblclick', (ev) => {
      if (ev.target.closest('.ddlist') || ev.target.closest('.dd-btn') || ev.target.closest('.editor.on')) return;
      const p = this.pointToCell(ev);
      if (!p || !this.sel) return;
      const g = this.area(p.r, p.c);
      if (g.r1 !== this.sel.r || g.c1 !== this.sel.c) return;
      ev.preventDefault();
      this.beginEdit('edit', this.rawText(this.doc.cell(this.si, this.sel.r, this.sel.c)));
    });
    this.ddBtn.addEventListener('mousedown', (ev) => {
      ev.preventDefault();
      if (this.ddList.style.display === 'block') this.closeDropdown();
      else this.openDropdown();
    });
    this.ddList.addEventListener('mousedown', (ev) => {
      ev.preventDefault();
      const o = ev.target.closest('.opt');
      if (o) this.chooseDD(+o.dataset.i);
    });
    this.canvas.addEventListener('mouseover', (ev) => {
      const m = ev.target.closest('.cmt');
      if (m) this.showNote(+m.dataset.k, m);
    });
    this.canvas.addEventListener('mouseout', (ev) => {
      if (ev.target.closest('.cmt')) this.hideNote();
    });
    this.scroller.addEventListener('scroll', () => this.hideNote());

    ed.addEventListener('compositionstart', () => {
      this.composing = true;
      // the composition replaces the selected text; do not touch the value here (Safari IME)
      if (!this.editing) this.beginEdit('enter', undefined);
    });
    ed.addEventListener('compositionend', () => { this.composing = false; this.autosize(); this.app.onEditing(true, ed.value); });
    ed.addEventListener('input', () => {
      if (!this.editing) {
        const v = ed.value;
        if (!this.beginEdit('enter', v)) { this.syncHidden(); return; }
      }
      this.autosize();
      this.app.onEditing(true, ed.value);
    });
    ed.addEventListener('paste', (ev) => {
      if (this.editing) return;
      ev.preventDefault();
      const text = (ev.clipboardData || window.clipboardData).getData('text/plain');
      if (this.sel) this.app.pasteText(this.si, this.sel.r, this.sel.c, text);
    });
    ed.addEventListener('copy', (ev) => {
      if (this.editing || !this.sel) return;
      ev.preventDefault();
      ev.clipboardData.setData('text/plain', this.app.copyText(this.si, this.sel.r, this.sel.c));
    });
    ed.addEventListener('cut', (ev) => {
      if (this.editing || !this.sel) return;
      ev.preventDefault();
      ev.clipboardData.setData('text/plain', this.app.copyText(this.si, this.sel.r, this.sel.c));
      if (this.doc.canEdit(this.si, this.sel.r, this.sel.c)) this.app.commitInput(this.si, this.sel.r, this.sel.c, '', { clear: true });
    });
    ed.addEventListener('blur', () => {
      // keep editing state; commit happens on explicit actions or clicks
    });
    ed.addEventListener('keydown', (ev) => this.onKey(ev));
  }

  onKey(ev) {
    const ed = this.editor;
    if (ev.isComposing || this.composing || ev.keyCode === 229) return;
    const k = ev.key;
    const mod = ev.metaKey || ev.ctrlKey;
    // dropdown navigation
    if (this.ddList.style.display === 'block') {
      const n = (this.ddList._items || []).length;
      if (k === 'ArrowDown') { this.ddIndex = Math.min(n - 1, this.ddIndex + 1); this.highlightDD(); ev.preventDefault(); return; }
      if (k === 'ArrowUp') { this.ddIndex = Math.max(0, this.ddIndex - 1); this.highlightDD(); ev.preventDefault(); return; }
      if (k === 'Enter') { this.chooseDD(this.ddIndex); ev.preventDefault(); return; }
      if (k === 'Escape') { this.closeDropdown(); ev.preventDefault(); return; }
    }
    if (this.editing) {
      if (k === 'Enter') {
        if (ev.altKey || (ev.ctrlKey && ev.altKey)) {
          ev.preventDefault();
          const s = ed.selectionStart, e = ed.selectionEnd;
          ed.setRangeText('\n', s, e, 'end');
          ed.dataset.wrap = '1';
          this.autosize();
          this.app.onEditing(true, ed.value);
          return;
        }
        ev.preventDefault();
        this.commitEdit(() => (ev.ctrlKey || ev.metaKey ? this.updateSelection(false) : this.enterMove(ev.shiftKey ? -1 : 1)));
        return;
      }
      if (k === 'Tab') { ev.preventDefault(); this.commitEdit(() => this.tabMove(ev.shiftKey ? -1 : 1)); return; }
      if (k === 'Escape') { ev.preventDefault(); this.cancelEdit(); return; }
      if (this.editMode === 'enter' && (k === 'ArrowUp' || k === 'ArrowDown' || k === 'ArrowLeft' || k === 'ArrowRight') && !ev.shiftKey && !mod) {
        ev.preventDefault();
        const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
        this.commitEdit(() => this.move(d[0], d[1]));
        return;
      }
      if (k === 'F2') { ev.preventDefault(); this.editMode = this.editMode === 'enter' ? 'edit' : 'enter'; return; }
      return;
    }
    // not editing
    if (mod && (k === 'z' || k === 'Z')) { ev.preventDefault(); if (ev.shiftKey) this.app.redo(); else this.app.undo(); return; }
    if (mod && (k === 'y' || k === 'Y')) { ev.preventDefault(); this.app.redo(); return; }
    if (mod && (k === 's' || k === 'S')) { ev.preventDefault(); this.app.exportDialog(); return; }
    if (mod && (k === 'c' || k === 'x' || k === 'v' || k === 'a')) return; // native clipboard events
    if (mod) return;
    switch (k) {
      case 'ArrowUp': ev.preventDefault(); if (ev.altKey) this.openDropdown(); else this.move(-1, 0); return;
      case 'ArrowDown': ev.preventDefault(); if (ev.altKey) this.openDropdown(); else this.move(1, 0); return;
      case 'ArrowLeft': ev.preventDefault(); this.move(0, -1); return;
      case 'ArrowRight': ev.preventDefault(); this.move(0, 1); return;
      case 'Tab': ev.preventDefault(); this.tabMove(ev.shiftKey ? -1 : 1); return;
      case 'Enter': ev.preventDefault(); this.enterMove(ev.shiftKey ? -1 : 1); return;
      case 'F2': ev.preventDefault(); if (this.sel) this.beginEdit('edit', this.rawText(this.doc.cell(this.si, this.sel.r, this.sel.c))); return;
      case 'Delete':
        ev.preventDefault();
        if (this.sel) {
          if (!this.doc.canEdit(this.si, this.sel.r, this.sel.c)) { this.app.protectedAlert(); return; }
          this.app.commitInput(this.si, this.sel.r, this.sel.c, '', { clear: true });
        }
        return;
      case 'Backspace':
        ev.preventDefault();
        if (this.sel) this.beginEdit('enter', '');
        return;
      case 'Escape': return;
      case ' ':
        break;
      default:
        if (k.length === 1 && this.sel && !this.doc.canEdit(this.si, this.sel.r, this.sel.c)) {
          ev.preventDefault();
          this.app.protectedAlert();
        }
    }
  }
}

function numToColLabel(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

export { colPx, MDW, fontFamily, resolveColor, addr, inRange };
