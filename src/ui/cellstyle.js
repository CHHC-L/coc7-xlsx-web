// Excel cell formats -> CSS / SVG drawing parameters.

import { resolveColor } from '../core/styles.js';

const FONT_MAP = {
  '宋体': "'SimSun','Songti SC','STSong',serif",
  '新宋体': "'NSimSun','SimSun','Songti SC','STSong',serif",
  '微软雅黑': "'Microsoft YaHei','PingFang SC','Hiragino Sans GB',sans-serif",
  '微软雅黑 light': "'Microsoft YaHei Light','Microsoft YaHei','PingFang SC','Hiragino Sans GB',sans-serif",
  '等线': "'DengXian','PingFang SC','Hiragino Sans GB',sans-serif",
  '等线 light': "'DengXian Light','DengXian','PingFang SC',sans-serif",
  '黑体': "'SimHei','Heiti SC','PingFang SC',sans-serif",
  '楷体': "'KaiTi','Kaiti SC','STKaiti',serif",
  '仿宋': "'FangSong','STFangsong',serif",
  'fzbiaoysk--gbk1-0': "'FZBiaoYaSongS-R-GB','FZBiaoYaSong-GBK','Songti SC','STSong',serif",
  'calibri': "Calibri,Carlito,'Helvetica Neue',Arial,sans-serif",
  'calibri light': "'Calibri Light',Calibri,'Helvetica Neue',sans-serif",
  'arial unicode ms': "'Arial Unicode MS',Arial,'PingFang SC',sans-serif",
  'arial': "Arial,'Helvetica Neue','PingFang SC',sans-serif",
};

export function fontFamily(name, theme, scheme) {
  if (scheme === 'minor' && theme.minorFont) name = name || theme.minorFont.ea || theme.minorFont.latin;
  if (scheme === 'major' && theme.majorFont) name = name || theme.majorFont.ea || theme.majorFont.latin;
  if (!name) return FONT_MAP['宋体'];
  const k = name.toLowerCase();
  if (FONT_MAP[k]) return FONT_MAP[k];
  if (FONT_MAP[name]) return FONT_MAP[name];
  return `'${name.replace(/['"]/g, '')}','PingFang SC',sans-serif`;
}

export function fontCss(font, theme, indexed) {
  if (!font) return '';
  let css = `font-family:${fontFamily(font.name, theme, font.scheme)};`;
  if (font.sz) css += `font-size:${font.sz}pt;`;
  if (font.b) css += 'font-weight:bold;';
  else if (font.name && /light/i.test(font.name)) css += 'font-weight:300;';
  if (font.i) css += 'font-style:italic;';
  const deco = [];
  if (font.u && font.u !== 'none') deco.push('underline');
  if (font.strike) deco.push('line-through');
  if (deco.length) css += `text-decoration:${deco.join(' ')};`;
  if (font.u === 'double' || font.u === 'doubleAccounting') css += 'text-decoration-style:double;';
  const col = resolveColor(font.color, theme, indexed, '000000');
  if (col) css += `color:#${col};`;
  if (font.vertAlign === 'superscript') css += 'vertical-align:super;font-size:smaller;';
  else if (font.vertAlign === 'subscript') css += 'vertical-align:sub;font-size:smaller;';
  return css;
}

// isDxf: in differential formats (conditional formatting) a solid fill's
// colour is stored in bgColor, and patternType may be omitted.
export function fillColor(fill, theme, indexed, isDxf = false) {
  if (!fill) return null;
  if (isDxf && fill.type === 'pattern') {
    const p = fill.patternExplicit ? fill.pattern : 'solid';
    if (p === 'none') return null;
    if (p === 'solid') {
      const c = resolveColor(fill.bg, theme, indexed, 'FFFFFF') || resolveColor(fill.fg, theme, indexed, 'FFFFFF');
      return c ? { solid: '#' + c } : null;
    }
  }
  if (fill.type === 'gradient') {
    const stops = fill.stops.map((s) => `#${resolveColor(s.color, theme, indexed, 'FFFFFF') || 'FFFFFF'} ${Math.round(s.pos * 100)}%`);
    return { gradient: `linear-gradient(${90 + fill.degree}deg,${stops.join(',')})`, solid: stops.length ? stops[0].split(' ')[0] : null };
  }
  const p = fill.pattern;
  if (!p || p === 'none') return null;
  const fg = resolveColor(fill.fg, theme, indexed, '000000') || '000000';
  const bg = resolveColor(fill.bg, theme, indexed, 'FFFFFF') || 'FFFFFF';
  if (p === 'solid') return { solid: '#' + fg };
  // approximate pattern fills by blending foreground over background
  const density = {
    gray125: 0.125, gray0625: 0.0625, lightGray: 0.25, mediumGray: 0.5, darkGray: 0.75,
    lightHorizontal: 0.25, lightVertical: 0.25, lightDown: 0.25, lightUp: 0.25, lightGrid: 0.4, lightTrellis: 0.4,
    darkHorizontal: 0.5, darkVertical: 0.5, darkDown: 0.5, darkUp: 0.5, darkGrid: 0.6, darkTrellis: 0.7,
  }[p] || 0.5;
  const mix = (a, b) => Math.round(parseInt(a, 16) * density + parseInt(b, 16) * (1 - density)).toString(16).padStart(2, '0');
  const c = mix(fg.slice(0, 2), bg.slice(0, 2)) + mix(fg.slice(2, 4), bg.slice(2, 4)) + mix(fg.slice(4, 6), bg.slice(4, 6));
  return { solid: '#' + c };
}

// border style -> {w: line width px, dash: svg dasharray|null, double: bool, rank}
const BORDER = {
  hair: { w: 1, dash: '1,1', rank: 1 },
  dotted: { w: 1, dash: '1,2', rank: 2 },
  dashDotDot: { w: 1, dash: '6,2,2,2,2,2', rank: 3 },
  dashDot: { w: 1, dash: '6,2,2,2', rank: 4 },
  dashed: { w: 1, dash: '3,1', rank: 5 },
  thin: { w: 1, dash: null, rank: 6 },
  mediumDashDotDot: { w: 2, dash: '6,2,2,2,2,2', rank: 7 },
  slantDashDot: { w: 2, dash: '7,2,3,2', rank: 8 },
  mediumDashDot: { w: 2, dash: '7,2,3,2', rank: 9 },
  mediumDashed: { w: 2, dash: '6,2', rank: 10 },
  medium: { w: 2, dash: null, rank: 11 },
  double: { w: 3, dash: null, double: true, rank: 12 },
  thick: { w: 3, dash: null, rank: 13 },
};

export function borderSpec(side, theme, indexed) {
  if (!side || !side.style) return null;
  const b = BORDER[side.style] || BORDER.thin;
  const col = resolveColor(side.color, theme, indexed, '000000') || '000000';
  return { ...b, style: side.style, color: '#' + col };
}

export function strongerBorder(a, b) {
  if (!a) return b;
  if (!b) return a;
  if (b.rank > a.rank) return b;
  if (b.rank < a.rank) return a;
  // same weight: darker colour wins
  const lum = (h) => { const n = parseInt(h.slice(1), 16); return ((n >> 16) & 255) * 0.3 + ((n >> 8) & 255) * 0.59 + (n & 255) * 0.11; };
  return lum(b.color) < lum(a.color) ? b : a;
}
