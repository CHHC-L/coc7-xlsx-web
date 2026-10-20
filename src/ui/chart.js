// Minimal renderer for the charts used by the card (filled radar charts).

import { parseFormula } from '../core/formula.js';
import { Ref, XErr } from '../core/values.js';

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function rangeValues(doc, f) {
  if (!f) return [];
  let ast;
  try { ast = parseFormula(f); } catch (e) { return []; }
  const v = doc.eng.evalAt(ast, 0, 1, 1, { keepRef: true });
  if (!(v instanceof Ref)) return [];
  const out = [];
  for (let r = v.r1; r <= v.r2; r++) for (let c = v.c1; c <= v.c2; c++) out.push(doc.eng.getValue(v.sheet, r, c));
  return out;
}

function niceStep(span) {
  if (span <= 0) return 0.2;
  const raw = span / 5;
  const p = Math.pow(10, Math.floor(Math.log10(raw)));
  const m = raw / p;
  return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p;
}

function schemeColor(doc, fill) {
  const theme = doc.wb.theme;
  if (!fill) return '#' + (theme.scheme ? theme.scheme.accent1 : '4472C4');
  if (fill.kind === 'srgbClr') return '#' + fill.val;
  if (fill.kind === 'schemeClr' && theme.scheme) {
    const map = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
    return '#' + (theme.scheme[map[fill.val] || fill.val] || '4472C4');
  }
  return '#4472C4';
}

export function renderChart(doc, chart, W, H) {
  if (chart.type !== 'radarChart') {
    return `<svg width="${W}" height="${H}"><rect width="${W}" height="${H}" fill="#fff" stroke="#d9d9d9"/></svg>`;
  }
  const ser = chart.series[0];
  if (!ser) return '';
  const cats = rangeValues(doc, ser.cat).map((v) => (v === null ? '' : v instanceof XErr ? v.code : String(v)));
  const vals = rangeValues(doc, ser.val).map((v) => (typeof v === 'number' ? v : 0));
  const n = Math.max(cats.length, vals.length);
  const lay = chart.layout || { x: 0.15, y: 0.15, w: 0.7, h: 0.7 };
  const px = lay.x * W, py = lay.y * H, pw = lay.w * W, ph = lay.h * H;
  const cx = px + pw / 2, cy = py + ph / 2;
  const R = Math.max(4, Math.min(pw, ph) / 2);
  let lo = Math.min(0, ...vals), hi = Math.max(0, ...vals);
  if (hi === lo) hi = lo + 1;
  const step = niceStep(hi - lo);
  hi = Math.ceil(hi / step) * step;
  lo = Math.floor(lo / step) * step;
  const ang = (i) => -Math.PI / 2 + (2 * Math.PI * i) / n;
  const pt = (i, v) => {
    const t = (v - lo) / (hi - lo);
    return [cx + Math.cos(ang(i)) * R * t, cy + Math.sin(ang(i)) * R * t];
  };
  let g = '';
  const rings = Math.round((hi - lo) / step);
  if (chart.hasMinorGrid) {
    for (let k = 1; k <= rings * 5; k++) {
      if (k % 5 === 0) continue;
      const v = lo + (k * step) / 5;
      g += `<polygon points="${[...Array(n).keys()].map((i) => pt(i, v).join(',')).join(' ')}" fill="none" stroke="#f2f2f2" stroke-width="0.75"/>`;
    }
  }
  for (let k = 1; k <= rings; k++) {
    const v = lo + k * step;
    g += `<polygon points="${[...Array(n).keys()].map((i) => pt(i, v).join(',')).join(' ')}" fill="none" stroke="#d9d9d9" stroke-width="0.75"/>`;
  }
  for (let i = 0; i < n; i++) {
    const [x, y] = pt(i, hi);
    g += `<line x1="${cx}" y1="${cy}" x2="${x}" y2="${y}" stroke="#d9d9d9" stroke-width="0.75"/>`;
  }
  const poly = [...Array(n).keys()].map((i) => pt(i, vals[i] || 0).join(',')).join(' ');
  g += `<polygon points="${poly}" fill="${schemeColor(doc, ser.fill)}" stroke="none"/>`;
  for (let i = 0; i < n; i++) {
    const a = ang(i);
    const [x, y] = [cx + Math.cos(a) * (R + 6), cy + Math.sin(a) * (R + 6)];
    const anchor = Math.abs(Math.cos(a)) < 0.2 ? 'middle' : Math.cos(a) > 0 ? 'start' : 'end';
    const lines = (cats[i] || '').split(/\r?\n/);
    const dy0 = Math.sin(a) < -0.5 ? -(lines.length - 1) * 11 : Math.sin(a) > 0.5 ? 9 : 4 - ((lines.length - 1) * 11) / 2;
    g += `<text x="${x}" y="${y + dy0}" text-anchor="${anchor}" font-size="9pt" fill="#595959">${lines.map((l, j) => `<tspan x="${x}" dy="${j ? 11 : 0}">${esc(l)}</tspan>`).join('')}</text>`;
  }
  return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="font-family:Calibri,'PingFang SC',sans-serif"><rect x="0.5" y="0.5" width="${W - 1}" height="${H - 1}" fill="#fff" stroke="#d9d9d9"/>${g}</svg>`;
}
