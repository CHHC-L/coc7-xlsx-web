// Excel number-format rendering.
// format(value, code) -> { parts: [{t:'s'|'pad', s}], color: 'RRGGBB'|null, text }
//   'pad' parts come from "_x" (a blank as wide as x) and are rendered hidden.

import { XErr, numToText, serialToDate } from './values.js';

const NAMED_COLORS = {
  black: '000000', blue: '0000FF', cyan: '00FFFF', green: '00FF00',
  magenta: 'FF00FF', red: 'FF0000', white: 'FFFFFF', yellow: 'FFFF00',
};

const cache = new Map();

function splitSections(code) {
  const out = [];
  let cur = '', q = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (q) { cur += ch; if (ch === '"') q = false; continue; }
    if (ch === '"') { q = true; cur += ch; continue; }
    if (ch === '\\' || ch === '_' || ch === '*') { cur += ch + (code[i + 1] || ''); i++; continue; }
    if (ch === ';') { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

// tokenise a section
function parseSection(src, indexed) {
  const sec = { tokens: [], color: null, cond: null, isDate: false, isText: false, isGeneral: false, hasDigits: false };
  let i = 0;
  const T = sec.tokens;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '"') {
      const j = src.indexOf('"', i + 1);
      T.push({ k: 'lit', s: src.slice(i + 1, j < 0 ? src.length : j) });
      i = j < 0 ? src.length : j + 1;
      continue;
    }
    if (ch === '\\') { T.push({ k: 'lit', s: src[i + 1] || '' }); i += 2; continue; }
    if (ch === '_') { T.push({ k: 'pad', s: src[i + 1] || ' ' }); i += 2; continue; }
    if (ch === '*') { T.push({ k: 'fill', s: src[i + 1] || ' ' }); i += 2; continue; }
    if (ch === '[') {
      const j = src.indexOf(']', i);
      const inner = src.slice(i + 1, j);
      i = j + 1;
      let m;
      if ((m = /^(=|<>|<=|>=|<|>)\s*(-?[\d.]+(?:[eE][+-]?\d+)?)$/.exec(inner))) {
        sec.cond = { op: m[1], v: parseFloat(m[2]) };
      } else if (NAMED_COLORS[inner.toLowerCase()]) {
        sec.color = NAMED_COLORS[inner.toLowerCase()];
      } else if ((m = /^color\s*(\d+)$/i.exec(inner))) {
        sec.color = indexed[+m[1] + 7] || null;
      } else if (/^(h+|m+|s+)$/i.test(inner)) {
        T.push({ k: 'elapsed', s: inner.toLowerCase() });
        sec.isDate = true;
      } else if (/^\$/.test(inner)) {
        // locale/currency: [$-804] or [$¥-804]
        const cur = /^\$([^-]*)/.exec(inner);
        if (cur && cur[1]) T.push({ k: 'lit', s: cur[1] });
      }
      continue;
    }
    if (/^general/i.test(src.slice(i, i + 7))) { T.push({ k: 'general' }); sec.isGeneral = true; i += 7; continue; }
    if (/^(AM\/PM|A\/P)/i.test(src.slice(i))) {
      const m = /^(AM\/PM|A\/P)/i.exec(src.slice(i));
      T.push({ k: 'ampm', s: m[0] });
      sec.isDate = true;
      i += m[0].length;
      continue;
    }
    if (src.startsWith('上午/下午', i)) { T.push({ k: 'ampm', s: '上午/下午' }); sec.isDate = true; i += 5; continue; }
    const lower = ch.toLowerCase();
    if ('ymdhse'.includes(lower) || (lower === 'a' && /^aaa/i.test(src.slice(i)))) {
      let j = i;
      while (j < src.length && src[j].toLowerCase() === lower) j++;
      T.push({ k: 'dt', s: src.slice(i, j).toLowerCase() });
      sec.isDate = true;
      i = j;
      continue;
    }
    if (ch === '0' || ch === '#' || ch === '?') { T.push({ k: 'dig', s: ch }); sec.hasDigits = true; i++; continue; }
    if (ch === '.') { T.push({ k: 'dot' }); i++; continue; }
    if (ch === ',') { T.push({ k: 'comma' }); i++; continue; }
    if (ch === '%') { T.push({ k: 'pct' }); i++; continue; }
    if ((ch === 'E' || ch === 'e') && (src[i + 1] === '+' || src[i + 1] === '-')) { T.push({ k: 'exp', s: src[i + 1] }); i += 2; continue; }
    if (ch === '@') { T.push({ k: 'text' }); sec.isText = true; i++; continue; }
    T.push({ k: 'lit', s: ch });
    i++;
  }
  // "e" alone is a year token only in date formats; "s"/"h"/"d" in non-date context are literals
  if (!sec.isDate) {
    for (const t of T) if (t.k === 'dt') { t.k = 'lit'; }
  }
  // 'm' after h or before s means minutes
  if (sec.isDate) {
    const dts = T.filter((t) => t.k === 'dt' || t.k === 'elapsed');
    for (let k = 0; k < dts.length; k++) {
      const t = dts[k];
      if (t.k === 'dt' && t.s[0] === 'm' && t.s.length <= 2) {
        const prev = dts[k - 1], next = dts[k + 1];
        if ((prev && prev.s[0] === 'h') || (next && next.s[0] === 's')) t.min = true;
      }
    }
  }
  return sec;
}

function compile(code, indexed) {
  const ck = code;
  if (cache.has(ck)) return cache.get(ck);
  const sections = splitSections(code).map((s) => parseSection(s, indexed));
  const r = { sections, hasCond: sections.some((s) => s.cond) };
  cache.set(ck, r);
  return r;
}

function condOk(c, v) {
  switch (c.op) {
    case '=': return v === c.v;
    case '<>': return v !== c.v;
    case '<': return v < c.v;
    case '>': return v > c.v;
    case '<=': return v <= c.v;
    case '>=': return v >= c.v;
    default: return false;
  }
}

function roundHalfAway(x, d) {
  const f = Math.pow(10, d);
  const y = Number((Math.abs(x) * f).toPrecision(15));
  return (Math.sign(x) * Math.round(y)) / f;
}

// Excel "General" for a cell of ordinary width (at most ~11 characters).
export function generalText(v) {
  if (v === 0) return '0';
  let s = numToText(v);
  const maxLen = v < 0 ? 12 : 11;
  if (s.length <= maxLen) return s;
  const a = Math.abs(v);
  if (a >= 1e11 || a < 1e-9) {
    for (let p = 6; p >= 1; p--) {
      let e = a.toExponential(p - 1).replace(/\.?0+e/, 'e');
      const m = /^([\d.]+)e([+-])(\d+)$/.exec(e);
      e = m[1] + 'E' + m[2] + m[3].padStart(2, '0');
      e = (v < 0 ? '-' : '') + e;
      if (e.length <= maxLen) return e;
    }
    return s;
  }
  const intLen = Math.floor(a).toString().length;
  for (let d = maxLen - intLen - 1 - (v < 0 ? 1 : 0); d >= 0; d--) {
    const t = Number(v.toFixed(d)).toString();
    if (t.length <= maxLen) return t;
  }
  return Math.round(v).toString();
}

function fmtNumberSection(sec, value, showMinus) {
  const T = sec.tokens;
  let v = value;
  const neg = v < 0;
  v = Math.abs(v);
  const parts = [];
  if (sec.isGeneral && !sec.hasDigits) {
    let pre = showMinus && neg ? '-' : '';
    for (const t of T) {
      if (t.k === 'general') parts.push({ t: 's', s: pre + generalText(v) }), (pre = '');
      else if (t.k === 'lit') parts.push({ t: 's', s: t.s });
      else if (t.k === 'pad') parts.push({ t: 'pad', s: t.s });
      else if (t.k === 'pct') parts.push({ t: 's', s: '%' });
    }
    return parts;
  }
  if (!sec.hasDigits) {
    // only literals (e.g. [=0]"" or "☆ 二选一")
    let first = true;
    for (const t of T) {
      if (t.k === 'lit') { parts.push({ t: 's', s: (first && showMinus && neg ? '-' : '') + t.s }); first = false; }
      else if (t.k === 'pad') parts.push({ t: 'pad', s: t.s });
      else if (t.k === 'pct') parts.push({ t: 's', s: '%' });
    }
    if (first && showMinus && neg) parts.unshift({ t: 's', s: '-' });
    return parts;
  }
  // percent scaling
  const pctCount = T.filter((t) => t.k === 'pct').length;
  v *= Math.pow(100, pctCount);
  // locate exponent
  const expIdx = T.findIndex((t) => t.k === 'exp');
  const dotIdx = T.findIndex((t, i) => t.k === 'dot' && (expIdx < 0 || i < expIdx));
  const mantEnd = expIdx >= 0 ? expIdx : T.length;
  // digits in integer / fraction part
  let lastDigit = -1;
  for (let i = 0; i < mantEnd; i++) if (T[i].k === 'dig') lastDigit = i;
  // scaling commas: commas directly after the last integer digit placeholder
  const intEnd = dotIdx >= 0 ? dotIdx : mantEnd;
  let lastIntDigit = -1;
  for (let i = 0; i < intEnd; i++) if (T[i].k === 'dig') lastIntDigit = i;
  let scaleCommas = 0;
  for (let i = lastIntDigit + 1; i < intEnd && T[i].k === 'comma'; i++) scaleCommas++;
  if (dotIdx < 0) for (let i = lastDigit + 1; i < T.length && T[i].k === 'comma'; i++) scaleCommas = Math.max(scaleCommas, i - lastDigit);
  v /= Math.pow(1000, scaleCommas);
  let thousands = false;
  let firstIntDigit = T.findIndex((t) => t.k === 'dig');
  for (let i = firstIntDigit; i >= 0 && i < lastIntDigit; i++) if (T[i].k === 'comma') thousands = true;

  const fracDigits = [];
  if (dotIdx >= 0) for (let i = dotIdx + 1; i < mantEnd; i++) if (T[i].k === 'dig') fracDigits.push(T[i].s);
  let expVal = 0;
  if (expIdx >= 0 && v !== 0) {
    const intPlaces = T.slice(0, intEnd).filter((t) => t.k === 'dig').length || 1;
    expVal = Math.floor(Math.log10(v));
    // engineering-style grouping when several integer placeholders (##0.0E+0)
    if (intPlaces > 1) expVal = Math.floor(expVal / intPlaces) * intPlaces;
    v = v / Math.pow(10, expVal);
  }
  v = roundHalfAway(v, fracDigits.length);
  if (expIdx >= 0 && v >= 10 && T.slice(0, intEnd).filter((t) => t.k === 'dig').length <= 1) { v /= 10; expVal += 1; }
  let intStr = Math.floor(v).toString();
  let fracStr = fracDigits.length ? Math.round((v - Math.floor(v)) * Math.pow(10, fracDigits.length)).toString().padStart(fracDigits.length, '0') : '';
  if (fracStr.length > fracDigits.length) { intStr = (Math.floor(v) + 1).toString(); fracStr = fracStr.slice(1); }
  if (intStr === '0') intStr = '';
  if (thousands && intStr) intStr = intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

  // integer part: fill placeholders from the right
  const intTokens = T.slice(0, intEnd);
  const intOut = new Array(intTokens.length).fill('');
  const digitsArr = intStr.split('');
  let di = digitsArr.length - 1;
  let firstPlaceholder = -1;
  for (let i = 0; i < intTokens.length; i++) if (intTokens[i].k === 'dig') { firstPlaceholder = i; break; }
  for (let i = intTokens.length - 1; i >= 0; i--) {
    const t = intTokens[i];
    if (t.k !== 'dig') continue;
    if (di >= 0) {
      if (i === firstPlaceholder) { intOut[i] = digitsArr.slice(0, di + 1).join(''); di = -1; }
      else { intOut[i] = digitsArr[di--]; if (thousands && intOut[i] === ',' ) { intOut[i] = digitsArr[di--] + ','; } }
    } else intOut[i] = t.s === '0' ? '0' : t.s === '?' ? ' ' : '';
  }
  const out = [];
  let signDone = !(showMinus && neg && (v !== 0 || fracStr.replace(/0/g, '') !== ''));
  const pushS = (s) => {
    if (!signDone && s !== '') { out.push({ t: 's', s: '-' }); signDone = true; }
    out.push({ t: 's', s });
  };
  for (let i = 0; i < intTokens.length; i++) {
    const t = intTokens[i];
    if (t.k === 'dig') { if (intOut[i] !== '') pushS(intOut[i]); }
    else if (t.k === 'lit') pushS(t.s);
    else if (t.k === 'pad') out.push({ t: 'pad', s: t.s });
    else if (t.k === 'pct') pushS('%');
    else if (t.k === 'text') pushS('');
  }
  if (dotIdx >= 0) {
    // fraction
    const fr = fracStr.split('');
    // trailing optional digits ('#') drop zeros; '?' become spaces
    let cut = fr.length;
    for (let k = fracDigits.length - 1; k >= 0; k--) {
      if (fracDigits[k] === '0') break;
      if (fr[k] === '0') cut = k; else break;
    }
    pushS('.');
    let fk = 0;
    for (let i = dotIdx + 1; i < mantEnd; i++) {
      const t = T[i];
      if (t.k === 'dig') {
        if (fk < cut) pushS(fr[fk]);
        else if (fracDigits[fk] === '?') pushS(' ');
        fk++;
      } else if (t.k === 'lit') pushS(t.s);
      else if (t.k === 'pad') out.push({ t: 'pad', s: t.s });
      else if (t.k === 'pct') pushS('%');
    }
  }
  if (expIdx >= 0) {
    const e = T[expIdx];
    const expDigits = T.slice(expIdx + 1).filter((t) => t.k === 'dig').length;
    const sign = expVal < 0 ? '-' : e.s === '+' ? '+' : '';
    pushS('E' + sign + Math.abs(expVal).toString().padStart(expDigits, '0'));
    for (let i = expIdx + 1; i < T.length; i++) {
      const t = T[i];
      if (t.k === 'lit') pushS(t.s);
      else if (t.k === 'pad') out.push({ t: 'pad', s: t.s });
    }
  }
  if (!signDone) out.unshift({ t: 's', s: '-' });
  return out;
}

const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const DAY_ZH = ['日', '一', '二', '三', '四', '五', '六'];
const MON_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

function fmtDateSection(sec, value) {
  if (value < 0) return [{ t: 's', s: '#'.repeat(8) }];
  const T = sec.tokens;
  const hasAmPm = T.some((t) => t.k === 'ampm');
  const fracDigitsAfterS = 0;
  let total = value;
  const dayInt = Math.floor(total);
  let secs = Math.round((total - dayInt) * 86400 * Math.pow(10, fracDigitsAfterS)) / Math.pow(10, fracDigitsAfterS);
  let dInt = dayInt;
  if (secs >= 86400) { secs -= 86400; dInt += 1; }
  const { y, m, d } = serialToDate(dInt);
  const H = Math.floor(secs / 3600), M = Math.floor((secs % 3600) / 60), S = Math.floor(secs % 60);
  const dow = (dInt + 6) % 7; // serial 1 = 1900-01-01 (Sunday per Excel's bug-compatible calendar)
  const out = [];
  for (const t of T) {
    if (t.k === 'lit') { out.push({ t: 's', s: t.s }); continue; }
    if (t.k === 'pad') { out.push({ t: 'pad', s: t.s }); continue; }
    if (t.k === 'ampm') {
      if (t.s === '上午/下午') out.push({ t: 's', s: H < 12 ? '上午' : '下午' });
      else if (/^a\/p$/i.test(t.s)) out.push({ t: 's', s: H < 12 ? t.s[0] : t.s[2] });
      else out.push({ t: 's', s: H < 12 ? 'AM' : 'PM' });
      continue;
    }
    if (t.k === 'elapsed') {
      const totalSecs = value * 86400;
      let n = t.s[0] === 'h' ? Math.floor(totalSecs / 3600) : t.s[0] === 'm' ? Math.floor(totalSecs / 60) : Math.floor(totalSecs);
      out.push({ t: 's', s: String(n).padStart(t.s.length, '0') });
      continue;
    }
    if (t.k === 'dt') {
      const s = t.s;
      let r = '';
      switch (s[0]) {
        case 'y': case 'e': r = s.length <= 2 && s[0] === 'y' ? String(y % 100).padStart(2, '0') : String(y); break;
        case 'm':
          if (t.min) r = s.length === 1 ? String(M) : String(M).padStart(2, '0');
          else if (s.length === 1) r = String(m);
          else if (s.length === 2) r = String(m).padStart(2, '0');
          else if (s.length === 3) r = MON_NAMES[m - 1].slice(0, 3);
          else if (s.length === 5) r = MON_NAMES[m - 1][0];
          else r = MON_NAMES[m - 1];
          break;
        case 'd':
          if (s.length === 1) r = String(d);
          else if (s.length === 2) r = String(d).padStart(2, '0');
          else if (s.length === 3) r = DAY_NAMES[dow].slice(0, 3);
          else r = DAY_NAMES[dow];
          break;
        case 'a': r = s.length >= 4 ? '星期' + DAY_ZH[dow] : DAY_ZH[dow]; break;
        case 'h': {
          let hh = H;
          if (hasAmPm) { hh = H % 12; if (hh === 0) hh = 12; }
          r = s.length === 1 ? String(hh) : String(hh).padStart(2, '0');
          break;
        }
        case 's': r = s.length === 1 ? String(S) : String(S).padStart(2, '0'); break;
        default: r = s;
      }
      out.push({ t: 's', s: r });
      continue;
    }
    if (t.k === 'dig' || t.k === 'dot') out.push({ t: 's', s: t.k === 'dot' ? '.' : '0' });
  }
  return out;
}

function fmtTextSection(sec, text) {
  const out = [];
  for (const t of sec.tokens) {
    if (t.k === 'text') out.push({ t: 's', s: text });
    else if (t.k === 'lit') out.push({ t: 's', s: t.s });
    else if (t.k === 'pad') out.push({ t: 'pad', s: t.s });
  }
  return out;
}

function finish(parts, color) {
  return { parts, color, text: parts.map((p) => (p.t === 'pad' ? ' ' : p.s)).join('') };
}

export function formatValue(value, code, indexed = []) {
  if (value === null || value === undefined) return finish([], null);
  if (value instanceof XErr) return finish([{ t: 's', s: value.code }], null);
  if (typeof value === 'boolean') return finish([{ t: 's', s: value ? 'TRUE' : 'FALSE' }], null);
  if (!code) code = 'General';
  const f = compile(code, indexed);
  const S = f.sections;
  if (typeof value === 'string') {
    let sec = null;
    if (S.length >= 4) sec = S[3];
    else if (S.length === 1 && S[0].isText) sec = S[0];
    else sec = S.find((s) => s.isText && !s.hasDigits && !s.isDate) || null;
    if (!sec) return finish([{ t: 's', s: value }], S.length === 1 && S[0].color ? S[0].color : null);
    return finish(fmtTextSection(sec, value), sec.color);
  }
  if (typeof value !== 'number') return finish([{ t: 's', s: String(value) }], null);
  let sec, showMinus = true;
  if (f.hasCond) {
    const c1 = S[0], c2 = S[1], c3 = S[2];
    if (c1 && c1.cond && condOk(c1.cond, value)) sec = c1;
    else if (c2 && c2.cond && condOk(c2.cond, value)) sec = c2;
    else if (c1 && !c1.cond && (!c2 || !c2.cond)) sec = c1; // unusual
    else if (c2 && !c2.cond && S.filter((s) => s.cond).length === 1 && S.length === 2) sec = c2;
    else if (c3 && !c3.cond) sec = c3;
    else if (c2 && !c2.cond && c1.cond && S.length >= 3) sec = c2;
    else return finish([{ t: 's', s: '#'.repeat(6) }], null);
    if (sec.cond && (sec.cond.op === '<' || sec.cond.op === '<=') && sec.cond.v <= 0) showMinus = false;
  } else if (S.length === 1 || (S.length >= 1 && S[0].isText && S.length === 1)) {
    sec = S[0];
  } else if (value > 0 || (value === 0 && S.length < 3)) {
    sec = S[0];
  } else if (value < 0) {
    sec = S[1] || S[0];
    if (S[1]) showMinus = false;
  } else {
    sec = S[2] || S[0];
  }
  if (sec.isText && !sec.hasDigits && !sec.isGeneral && !sec.isDate) {
    // text-only section applied to a number: numbers show as General
    return finish([{ t: 's', s: generalText(value) }], sec.color);
  }
  if (sec.isDate) return finish(fmtDateSection(sec, value), sec.color);
  return finish(fmtNumberSection(sec, value, showMinus), sec.color);
}

export function isDateFormat(code) {
  if (!code) return false;
  const f = compile(code, []);
  return f.sections[0] && f.sections[0].isDate;
}
