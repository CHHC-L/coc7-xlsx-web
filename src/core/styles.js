// styles.xml + theme parsing and colour resolution.

import { parseXml } from './xml.js';

export const BUILTIN_NUMFMTS = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'yyyy/m/d', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy/m/d h:mm',
  27: 'yyyy"年"m"月"', 28: 'm"月"d"日"', 29: 'm"月"d"日"', 30: 'm-d-yy', 31: 'yyyy"年"m"月"d"日"',
  32: 'h"时"mm"分"', 33: 'h"时"mm"分"ss"秒"', 34: '上午/下午h"时"mm"分"', 35: '上午/下午h"时"mm"分"ss"秒"',
  36: 'yyyy"年"m"月"',
  37: '#,##0_);(#,##0)', 38: '#,##0_);[Red](#,##0)', 39: '#,##0.00_);(#,##0.00)', 40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
  50: 'yyyy"年"m"月"', 51: 'm"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"', 54: 'm"月"d"日"',
  55: '上午/下午h"时"mm"分"', 56: '上午/下午h"时"mm"分"ss"秒"', 57: 'yyyy"年"m"月"', 58: 'm"月"d"日"',
};

const DEFAULT_INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

function parseColor(n) {
  if (!n) return null;
  const c = {};
  if (n.attrs.auto === '1' || n.attrs.auto === 'true') c.auto = true;
  if (n.attrs.rgb !== undefined) c.rgb = n.attrs.rgb;
  if (n.attrs.theme !== undefined) c.theme = +n.attrs.theme;
  if (n.attrs.indexed !== undefined) c.indexed = +n.attrs.indexed;
  if (n.attrs.tint !== undefined) c.tint = +n.attrs.tint;
  return c;
}

export function parseFont(n) {
  if (!n) return null;
  const f = {};
  for (const ch of n.children) {
    const v = ch.attrs.val;
    switch (ch.local) {
      case 'name': case 'rFont': f.name = v; break;
      case 'sz': f.sz = +v; break;
      case 'b': f.b = v === undefined || v === '1' || v === 'true'; break;
      case 'i': f.i = v === undefined || v === '1' || v === 'true'; break;
      case 'strike': f.strike = v === undefined || v === '1' || v === 'true'; break;
      case 'u': f.u = v === undefined ? 'single' : v; break;
      case 'vertAlign': f.vertAlign = v; break;
      case 'color': f.color = parseColor(ch); break;
      case 'family': f.family = +v; break;
      case 'scheme': f.scheme = v; break;
      case 'charset': f.charset = +v; break;
      case 'outline': f.outline = true; break;
      case 'shadow': f.shadow = true; break;
      default:
    }
  }
  return f;
}

function parseFill(n) {
  if (!n) return null;
  const p = n.child('patternFill');
  if (p) {
    return {
      type: 'pattern',
      pattern: p.attrs.patternType || (p.child('fgColor') || p.child('bgColor') ? 'solid' : 'none'),
      patternExplicit: p.attrs.patternType !== undefined,
      fg: parseColor(p.child('fgColor')),
      bg: parseColor(p.child('bgColor')),
    };
  }
  const g = n.child('gradientFill');
  if (g) {
    return {
      type: 'gradient',
      gtype: g.attrs.type || 'linear',
      degree: +(g.attrs.degree || 0),
      stops: g.all('stop').map((s) => ({ pos: +s.attrs.position, color: parseColor(s.child('color')) })),
    };
  }
  return { type: 'pattern', pattern: 'none' };
}

function parseBorder(n) {
  if (!n) return null;
  const b = { diagonalUp: n.attrs.diagonalUp === '1', diagonalDown: n.attrs.diagonalDown === '1' };
  for (const side of ['left', 'right', 'top', 'bottom', 'diagonal', 'start', 'end']) {
    const s = n.child(side);
    if (s && s.attrs.style && s.attrs.style !== 'none') {
      b[side === 'start' ? 'left' : side === 'end' ? 'right' : side] = { style: s.attrs.style, color: parseColor(s.child('color')) };
    }
  }
  return b;
}

function parseXf(n) {
  const a = n.attrs;
  const xf = {
    numFmtId: +(a.numFmtId || 0),
    fontId: +(a.fontId || 0),
    fillId: +(a.fillId || 0),
    borderId: +(a.borderId || 0),
    xfId: a.xfId !== undefined ? +a.xfId : 0,
    quotePrefix: a.quotePrefix === '1',
    alignment: null,
    locked: true,
    hidden: false,
  };
  const al = n.child('alignment');
  if (al) {
    xf.alignment = {
      horizontal: al.attrs.horizontal || 'general',
      vertical: al.attrs.vertical || 'bottom',
      wrapText: al.attrs.wrapText === '1' || al.attrs.wrapText === 'true',
      shrinkToFit: al.attrs.shrinkToFit === '1',
      indent: +(al.attrs.indent || 0),
      textRotation: +(al.attrs.textRotation || 0),
      justifyLastLine: al.attrs.justifyLastLine === '1',
    };
  }
  const pr = n.child('protection');
  if (pr) {
    if (pr.attrs.locked !== undefined) xf.locked = !(pr.attrs.locked === '0' || pr.attrs.locked === 'false');
    if (pr.attrs.hidden !== undefined) xf.hidden = pr.attrs.hidden === '1' || pr.attrs.hidden === 'true';
  }
  return xf;
}

export function parseStyles(xml) {
  const root = parseXml(xml);
  const st = {
    numFmts: new Map(),
    fonts: [], fills: [], borders: [], cellXfs: [], cellStyleXfs: [], dxfs: [],
    indexed: DEFAULT_INDEXED.slice(),
    root, xml,
  };
  const nf = root.child('numFmts');
  if (nf) for (const n of nf.all('numFmt')) st.numFmts.set(+n.attrs.numFmtId, n.attrs.formatCode);
  const fonts = root.child('fonts');
  if (fonts) st.fonts = fonts.all('font').map(parseFont);
  const fills = root.child('fills');
  if (fills) st.fills = fills.all('fill').map(parseFill);
  const borders = root.child('borders');
  if (borders) st.borders = borders.all('border').map(parseBorder);
  const csx = root.child('cellStyleXfs');
  if (csx) st.cellStyleXfs = csx.all('xf').map(parseXf);
  const cx = root.child('cellXfs');
  if (cx) st.cellXfs = cx.all('xf').map(parseXf);
  const dxfs = root.child('dxfs');
  if (dxfs) {
    st.dxfs = dxfs.all('dxf').map((d) => {
      const o = {};
      if (d.child('font')) o.font = parseFont(d.child('font'));
      if (d.child('fill')) o.fill = parseFill(d.child('fill'));
      if (d.child('border')) o.border = parseBorder(d.child('border'));
      if (d.child('numFmt')) o.numFmt = d.child('numFmt').attrs.formatCode;
      if (d.child('alignment')) o.alignment = d.child('alignment').attrs;
      return o;
    });
  }
  const colors = root.child('colors');
  if (colors && colors.child('indexedColors')) {
    st.indexed = colors.child('indexedColors').all('rgbColor').map((c) => (c.attrs.rgb || 'FF000000').slice(-6));
  }
  st.numFmtCode = (id) => (st.numFmts.has(id) ? st.numFmts.get(id) : BUILTIN_NUMFMTS[id] || 'General');
  return st;
}

export function parseTheme(xml) {
  const theme = { colors: [], minorFont: null, majorFont: null };
  if (!xml) {
    theme.colors = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
    return theme;
  }
  const root = parseXml(xml);
  const scheme = root.find('clrScheme')[0];
  const get = (name) => {
    const n = scheme && scheme.child(name);
    if (!n) return '000000';
    const c = n.children[0];
    if (!c) return '000000';
    if (c.local === 'srgbClr') return c.attrs.val;
    if (c.local === 'sysClr') return c.attrs.lastClr || (c.attrs.val === 'window' ? 'FFFFFF' : '000000');
    return '000000';
  };
  // SpreadsheetML theme index order swaps dk/lt pairs
  theme.colors = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(get);
  theme.scheme = {};
  for (const k of ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) theme.scheme[k] = get(k);
  const fontFace = (tag) => {
    const f = root.find(tag)[0];
    if (!f) return null;
    const latin = f.child('latin');
    const hans = f.all('font').find((x) => x.attrs.script === 'Hans');
    const ea = f.child('ea');
    return { latin: latin && latin.attrs.typeface, ea: (hans && hans.attrs.typeface) || (ea && ea.attrs.typeface) || '' };
  };
  theme.minorFont = fontFace('minorFont');
  theme.majorFont = fontFace('majorFont');
  return theme;
}

// ---- colour maths (Excel's HLS tint) -------------------------------------------

function rgbToHls(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  let h = 0, s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  return [h, l, s];
}

function hlsToRgb(h, l, s) {
  if (s === 0) return [l, l, l].map((x) => Math.round(x * 255));
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3), f(h), f(h - 1 / 3)].map((x) => Math.round(x * 255));
}

export function applyTint(hex, tint) {
  if (!tint) return hex;
  const r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
  let [h, l, s] = rgbToHls(r, g, b);
  if (tint < 0) l = l * (1 + tint);
  else l = l * (1 - tint) + tint;
  const [R, G, B] = hlsToRgb(h, l, s);
  return [R, G, B].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
}

// Resolve a colour spec to "RRGGBB" (or null for automatic / none).
export function resolveColor(c, theme, indexed, autoDefault = null) {
  if (!c) return null;
  let hex = null;
  if (c.rgb) hex = c.rgb.length === 8 ? c.rgb.slice(2) : c.rgb;
  else if (c.theme !== undefined) hex = theme.colors[c.theme] || '000000';
  else if (c.indexed !== undefined) {
    if (c.indexed === 64) return autoDefault || '000000';
    if (c.indexed === 65) return autoDefault || 'FFFFFF';
    hex = indexed[c.indexed] || '000000';
  } else if (c.auto) return autoDefault;
  else return null;
  return applyTint(hex.toUpperCase(), c.tint);
}
