// Value model shared by the formula engine, formatter and exporter.
//   number | string | boolean | XErr | null (blank) | Ref | Arr

export class XErr {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}

export const ERR = {
  NULL: new XErr('#NULL!'),
  DIV0: new XErr('#DIV/0!'),
  VALUE: new XErr('#VALUE!'),
  REF: new XErr('#REF!'),
  NAME: new XErr('#NAME?'),
  NUM: new XErr('#NUM!'),
  NA: new XErr('#N/A'),
  CALC: new XErr('#CALC!'),
  SPILL: new XErr('#SPILL!'),
  GETTING_DATA: new XErr('#GETTING_DATA'),
};

const ERR_BY_CODE = new Map(Object.values(ERR).map((e) => [e.code, e]));
export function errFromCode(code) {
  return ERR_BY_CODE.get(code.toUpperCase()) || new XErr(code);
}

export const isErr = (v) => v instanceof XErr;

// A rectangular reference into a worksheet (by sheet index).
export class Ref {
  constructor(sheet, r1, c1, r2, c2) {
    this.sheet = sheet;
    this.r1 = r1; this.c1 = c1; this.r2 = r2; this.c2 = c2;
  }
  get rows() { return this.r2 - this.r1 + 1; }
  get cols() { return this.c2 - this.c1 + 1; }
  get single() { return this.r1 === this.r2 && this.c1 === this.c2; }
}

// An in-memory array (array constants, intermediate results).
export class Arr {
  constructor(rows) { this.data = rows; }
  get rows() { return this.data.length; }
  get cols() { return this.data[0] ? this.data[0].length : 0; }
  get(r, c) { return this.data[r][c]; }
}

// ---- numbers <-> text -------------------------------------------------------

// Excel's 15-significant-digit text form used by "&", TEXT-less conversion
// and the General format when width is unlimited.
export function numToText(n) {
  if (!isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  const a = Math.abs(n);
  // 15 significant digits: mantissa digits + decimal exponent
  const [mant, expS] = a.toExponential(14).split('e');
  const exp = +expS;
  const digits = mant.replace('.', '').replace(/0+$/, '') || '0';
  const sign = n < 0 ? '-' : '';
  if (exp >= 15 || exp < -9) {
    const m = digits.length > 1 ? digits[0] + '.' + digits.slice(1) : digits;
    return sign + m + 'E' + (exp < 0 ? '-' : '+') + String(Math.abs(exp)).padStart(2, '0');
  }
  if (exp >= 0) {
    if (digits.length <= exp + 1) return sign + digits + '0'.repeat(exp + 1 - digits.length);
    return sign + digits.slice(0, exp + 1) + '.' + digits.slice(exp + 1);
  }
  return sign + '0.' + '0'.repeat(-exp - 1) + digits;
}

// Full-width -> half-width for digits and the usual numeric punctuation.
export function normalizeWidth(s) {
  return s.replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0)).replace(/　/g, ' ');
}

const DAYS_BEFORE_MONTH = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];
function isLeap(y) { return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0; }

// Excel serial date (1900 system, with the fictitious 1900-02-29).
export function dateToSerial(y, m, d) {
  if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1) return null;
  const mdays = [31, isLeap(y) || y === 1900 ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (d > mdays[m - 1]) return null;
  // days since 1899-12-31
  let days = 0;
  for (let yy = 1900; yy < y; yy++) days += isLeap(yy) ? 366 : 365;
  days += DAYS_BEFORE_MONTH[m - 1] + (m > 2 && isLeap(y) ? 1 : 0) + d;
  if (y > 1900 || m > 2) days += 1; // Lotus 1-2-3 leap-year bug
  return days;
}

export function serialToDate(serial) {
  let n = Math.floor(serial);
  if (n === 60) return { y: 1900, m: 2, d: 29 };
  if (n > 60) n -= 1;
  let y = 1900;
  for (;;) {
    const len = isLeap(y) ? 366 : 365;
    if (n <= len) break;
    n -= len; y++;
  }
  const mdays = [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  let m = 0;
  while (n > mdays[m]) { n -= mdays[m]; m++; }
  return { y, m: m + 1, d: n };
}

// Parse text the way Excel coerces strings to numbers (and the way typed
// input becomes a number). Returns {value, fmt} or null.
//   fmt is a suggested built-in number format id when Excel would apply one.
export function parseNumericText(text, { allowDate = true } = {}) {
  let s = normalizeWidth(String(text)).trim();
  if (!s) return null;
  // plain / scientific number with optional thousands separators
  let neg = false;
  let m;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1).trim(); }
  let sign = 1;
  if (s[0] === '-' || s[0] === '+') { if (s[0] === '-') sign = -1; s = s.slice(1).trim(); }
  let currency = false;
  if (s[0] === '$' || s[0] === '¥' || s[0] === '￥') { currency = true; s = s.slice(1).trim(); if (s[0] === '-') { sign = -sign; s = s.slice(1); } }
  let pct = false;
  if (s.endsWith('%')) { pct = true; s = s.slice(0, -1).trim(); }
  if ((m = /^(\d{1,3}(,\d{3})+|\d+)?(\.\d*)?([eE][+-]?\d+)?$/.exec(s)) && (m[1] || (m[3] && m[3].length > 1))) {
    const thousands = !!(m[1] && m[1].indexOf(',') >= 0);
    let v = parseFloat(s.replace(/,/g, '')) * sign * (neg ? -1 : 1);
    if (pct) v /= 100;
    if (!isFinite(v)) return null;
    let fmt = null;
    if (pct) fmt = (m[3] && m[3].length > 1) ? 10 : 9;
    else if (currency) fmt = (m[3] && m[3].length > 1) ? 8 : 6; // approximate currency formats
    else if (thousands) fmt = (m[3] && m[3].length > 1) ? 4 : 3;
    else if (m[4]) fmt = 11;
    return { value: v, fmt };
  }
  if (!allowDate || neg || pct || currency || sign < 0) return null;
  return parseDateTimeText(s);
}

function parseTime(s) {
  const m = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?\s*([AaPp][Mm])?$/.exec(s);
  if (!m) return null;
  let h = +m[1];
  const mi = +m[2];
  const se = m[3] ? +m[3] : 0;
  if (mi > 59 || se >= 60) return null;
  if (m[4]) {
    if (h > 12 || h === 0) return null;
    const pm = /p/i.test(m[4]);
    if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12;
  }
  if (h > 9999) return null;
  return { value: (h * 3600 + mi * 60 + se) / 86400, fmt: m[3] ? (m[4] ? 19 : 21) : (m[4] ? 18 : 20) };
}

function parseDateTimeText(s) {
  const t = parseTime(s);
  if (t) return t;
  let m;
  const now = new Date();
  let y, mo, d, fmt = 14;
  if ((m = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:\s+(.+))?$/.exec(s))) {
    y = +m[1]; mo = +m[2]; d = +m[3];
    const ser = dateToSerial(y, mo, d);
    if (ser === null) return null;
    if (m[4]) { const tt = parseTime(m[4]); if (!tt) return null; return { value: ser + tt.value, fmt: 22 }; }
    return { value: ser, fmt: 14 };
  }
  if ((m = /^(\d{4})年(\d{1,2})月(?:(\d{1,2})日)?$/.exec(s))) {
    y = +m[1]; mo = +m[2]; d = m[3] ? +m[3] : 1;
    const ser = dateToSerial(y, mo, d);
    if (ser === null) return null;
    return { value: ser, fmt: m[3] ? 31 : 57 };
  }
  if ((m = /^(\d{1,2})月(\d{1,2})日$/.exec(s))) {
    mo = +m[1]; d = +m[2];
    const ser = dateToSerial(now.getFullYear(), mo, d);
    if (ser === null) return null;
    return { value: ser, fmt: 58 };
  }
  if ((m = /^(\d{1,2})[/-](\d{1,2})$/.exec(s))) {
    // m/d in the current year (zh-CN short date order)
    mo = +m[1]; d = +m[2];
    const ser = dateToSerial(now.getFullYear(), mo, d);
    if (ser === null) {
      // maybe year/month like 2023/1 handled above; try y/m with 2-digit year? Excel treats "1/40" as Jan 1940
      if (mo >= 1 && mo <= 12 && d >= 0 && d <= 99) {
        const yy = d < 30 ? 2000 + d : 1900 + d;
        const s2 = dateToSerial(yy, mo, 1);
        if (s2 !== null) return { value: s2, fmt: 17 };
      }
      return null;
    }
    return { value: ser, fmt: 58 };
  }
  return null;
}

// ---- coercions ---------------------------------------------------------------

export function toNumber(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof XErr) return v;
  if (typeof v === 'string') {
    const p = parseNumericText(v);
    return p ? p.value : ERR.VALUE;
  }
  return ERR.VALUE;
}

export function toBool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === undefined) return false;
  if (v instanceof XErr) return v;
  if (typeof v === 'string') {
    const u = v.toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    return ERR.VALUE;
  }
  return ERR.VALUE;
}

export function toText(v) {
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numToText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v === null || v === undefined) return '';
  if (v instanceof XErr) return v;
  return ERR.VALUE;
}

// Type rank used when comparing values of different types.
function rank(v) {
  if (typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 3;
  return 0;
}

const collator = typeof Intl !== 'undefined' ? new Intl.Collator('zh-CN', { sensitivity: 'base' }) : null;

export function compareText(a, b) {
  const x = a.toLowerCase(), y = b.toLowerCase();
  if (x === y) return 0;
  if (collator) {
    const c = collator.compare(x, y);
    if (c !== 0) return c < 0 ? -1 : 1;
  }
  return x < y ? -1 : 1;
}

// Excel treats numbers as equal when they agree to 15 significant digits
// (so =0.1+0.2=0.3 is TRUE).
export function numEq(a, b) {
  if (a === b) return true;
  return Number(a.toPrecision(15)) === Number(b.toPrecision(15));
}

// Excel comparison for =, <, > ... (blank is coerced to the other side's type)
export function compareValues(a, b) {
  if (a === null || a === undefined) {
    if (b === null || b === undefined) return 0;
    a = typeof b === 'number' ? 0 : typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  } else if (b === null || b === undefined) {
    b = typeof a === 'number' ? 0 : typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  }
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (ra === 1) return numEq(a, b) ? 0 : a < b ? -1 : 1;
  if (ra === 2) return compareText(a, b);
  if (ra === 3) return a === b ? 0 : a ? 1 : -1;
  return 0;
}
