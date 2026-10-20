// Formula evaluation with Excel semantics (non-array / implicit-intersection
// mode, which is how every formula in a classic workbook behaves).

import {
  ERR, XErr, Ref, Arr, isErr,
  toNumber, toBool, toText, compareValues, compareText, numEq, numToText,
} from './values.js';
import { MAXROW, MAXCOL } from './refs.js';

const MISSING = { missing: true };

// ctx: { eng, sheet, r, c }
export function evaluate(node, ctx) {
  switch (node.type) {
    case 'num': case 'str': case 'bool': return node.v;
    case 'err': return node.v;
    case 'missing': return MISSING;
    case 'paren': return evaluate(node.a, ctx);
    case 'ref': {
      let si = ctx.sheet;
      if (node.sheet) {
        si = ctx.eng.sheetIndex(node.sheet);
        if (si < 0) return ERR.REF;
      }
      return new Ref(si, node.r1, node.c1, node.r2, node.c2);
    }
    case 'name': return ctx.eng.evalName(node.name, node.sheet, ctx);
    case 'arr':
      return new Arr(node.rows.map((row) => row.map((x) => (x.type === 'err' ? x.v : x.v))));
    case 'un': {
      const v = toNumber(scalar(evaluate(node.a, ctx), ctx));
      if (isErr(v)) return v;
      return node.op === '-' ? -v : v;
    }
    case 'pct': {
      const v = toNumber(scalar(evaluate(node.a, ctx), ctx));
      if (isErr(v)) return v;
      return v / 100;
    }
    case 'bin': return binop(node, ctx);
    case 'fn': {
      const f = FUNCS[node.name] || FUNCS[node.name.replace(/^_XLFN\.|^_XLWS\./, '')];
      if (!f) return ERR.NAME;
      return f(node.args, ctx);
    }
    default: return ERR.VALUE;
  }
}

export function cellValue(ctx, si, r, c) {
  return ctx.eng.getValue(si, r, c);
}

// Implicit intersection / dereference to a single value.
export function scalar(v, ctx) {
  if (v instanceof Ref) {
    if (v.single) return cellValue(ctx, v.sheet, v.r1, v.c1);
    if (v.c1 === v.c2 && ctx.r >= v.r1 && ctx.r <= v.r2) return cellValue(ctx, v.sheet, ctx.r, v.c1);
    if (v.r1 === v.r2 && ctx.c >= v.c1 && ctx.c <= v.c2) return cellValue(ctx, v.sheet, v.r1, ctx.c);
    return ERR.VALUE;
  }
  if (v instanceof Arr) return v.rows ? v.get(0, 0) : ERR.VALUE;
  if (v === MISSING) return null;
  return v;
}

function binop(node, ctx) {
  const op = node.op;
  if (op === ':') {
    const a = evaluate(node.a, ctx), b = evaluate(node.b, ctx);
    if (!(a instanceof Ref) || !(b instanceof Ref) || a.sheet !== b.sheet) return ERR.VALUE;
    return new Ref(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
  }
  const a = scalar(evaluate(node.a, ctx), ctx);
  const b = scalar(evaluate(node.b, ctx), ctx);
  if (op === '&') {
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    return toText(a) + toText(b);
  }
  if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    const c = compareValues(a, b);
    switch (op) {
      case '=': return c === 0;
      case '<>': return c !== 0;
      case '<': return c < 0;
      case '>': return c > 0;
      case '<=': return c <= 0;
      default: return c >= 0;
    }
  }
  const x = toNumber(a);
  if (isErr(x)) return x;
  const y = toNumber(b);
  if (isErr(y)) return y;
  let r;
  switch (op) {
    case '+': r = x + y; break;
    case '-': r = x - y; break;
    case '*': r = x * y; break;
    case '/': if (y === 0) return ERR.DIV0; r = x / y; break;
    case '^':
      if (x === 0 && y === 0) return ERR.NUM;
      if (x === 0 && y < 0) return ERR.DIV0;
      r = Math.pow(x, y);
      break;
    default: return ERR.VALUE;
  }
  if (!isFinite(r) || Number.isNaN(r)) return ERR.NUM;
  return r;
}

// ---- helpers --------------------------------------------------------------------

function ev(args, i, ctx) {
  return i < args.length ? evaluate(args[i], ctx) : MISSING;
}

function sv(args, i, ctx) {
  // scalar value of argument i (MISSING -> null)
  if (i >= args.length) return MISSING;
  const v = evaluate(args[i], ctx);
  if (v === MISSING) return MISSING;
  return scalar(v, ctx);
}

function num(args, i, ctx, dflt) {
  const v = sv(args, i, ctx);
  if (v === MISSING) return dflt !== undefined ? dflt : 0;
  return toNumber(v);
}

// Iterate the values of a range/array argument. cb(value) returns true to stop.
function eachValue(v, ctx, cb) {
  if (v instanceof Ref) {
    ctx.eng.forEachInRef(v, cb);
  } else if (v instanceof Arr) {
    for (const row of v.data) for (const x of row) if (cb(x)) return;
  } else cb(v);
}

// Collect numbers like SUM/MAX/AVERAGE do. Returns XErr or array of numbers.
function collectNumbers(args, ctx) {
  const out = [];
  for (let i = 0; i < args.length; i++) {
    const raw = evaluate(args[i], ctx);
    if (raw === MISSING) { out.push(0); continue; }
    if (raw instanceof Ref || raw instanceof Arr) {
      let err = null;
      eachValue(raw, ctx, (x) => {
        if (typeof x === 'number') out.push(x);
        else if (isErr(x)) { err = x; return true; }
        return false;
      });
      if (err) return err;
    } else {
      const n = toNumber(raw);
      if (isErr(n)) return n;
      out.push(n);
    }
  }
  return out;
}

function collectLogicals(args, ctx) {
  const out = [];
  for (let i = 0; i < args.length; i++) {
    const raw = evaluate(args[i], ctx);
    if (raw === MISSING) continue;
    if (raw instanceof Ref || raw instanceof Arr) {
      let err = null;
      eachValue(raw, ctx, (x) => {
        if (typeof x === 'boolean') out.push(x);
        else if (typeof x === 'number') out.push(x !== 0);
        else if (isErr(x)) { err = x; return true; }
        return false;
      });
      if (err) return err;
    } else {
      const b = toBool(raw);
      if (isErr(b)) return b;
      out.push(b);
    }
  }
  return out;
}

function roundSig(x) {
  return Number(x.toPrecision(15));
}

function wildcardRegex(pattern) {
  if (!/[*?~]/.test(pattern)) return null;
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) {
      re += '\\' + pattern[++i];
    } else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp('^' + re + '$', 'i');
}

function exactMatcher(target) {
  if (typeof target === 'string') {
    const re = wildcardRegex(target);
    const low = target.toLowerCase();
    if (re) return (v) => typeof v === 'string' && re.test(v);
    return (v) => typeof v === 'string' && v.toLowerCase() === low;
  }
  if (typeof target === 'number') return (v) => typeof v === 'number' && numEq(v, target);
  if (typeof target === 'boolean') return (v) => v === target;
  return () => false;
}

function typeClass(v) {
  if (typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 3;
  return 0; // blank / error
}

// Excel-style approximate binary search. get(i) for i in [0,len).
// asc: largest value <= target; desc: smallest value >= target.
function approxSearch(len, get, target, desc = false) {
  const tc = typeClass(target);
  let lo = 0, hi = len - 1, best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let m = mid;
    let v = get(m);
    while (m >= lo && typeClass(v) !== tc) { m--; if (m >= lo) v = get(m); }
    if (m < lo) { lo = mid + 1; continue; }
    const c = compareValues(v, target);
    if (desc ? c >= 0 : c <= 0) { best = m; lo = mid + 1; } else hi = m - 1;
  }
  return best;
}

// 1-D accessor for a range or array; null if 2-D.
function vectorOf(v, ctx, allow2D = false) {
  if (v instanceof Ref) {
    if (v.rows !== 1 && v.cols !== 1 && !allow2D) return null;
    const vertical = v.cols === 1 || (allow2D && v.rows >= v.cols);
    const sh = ctx.eng.sheet(v.sheet);
    const len0 = vertical ? v.rows : v.cols;
    // clamp very long (whole column) ranges to the used area for speed
    const limit = vertical ? Math.max(0, Math.min(len0, (sh ? sh.maxRow : 0) - v.r1 + 1)) : Math.max(0, Math.min(len0, (sh ? sh.maxCol : 0) - v.c1 + 1));
    return {
      len: len0,
      scanLen: limit,
      get: (i) => (vertical ? cellValue(ctx, v.sheet, v.r1 + i, v.c1) : cellValue(ctx, v.sheet, v.r1, v.c1 + i)),
      vertical,
    };
  }
  if (v instanceof Arr) {
    if (v.rows !== 1 && v.cols !== 1 && !allow2D) return null;
    const vertical = v.cols === 1 || (allow2D && v.rows >= v.cols);
    const len = vertical ? v.rows : v.cols;
    return { len, scanLen: len, get: (i) => (vertical ? v.get(i, 0) : v.get(0, i)), vertical };
  }
  return { len: 1, scanLen: 1, get: () => v, vertical: true };
}

function lookupTarget(v) {
  // empty lookup value behaves like 0
  if (v === null || v === MISSING) return 0;
  return v;
}

// ---- functions -------------------------------------------------------------------

const FUNCS = {};

FUNCS.IF = (args, ctx) => {
  if (args.length < 1) return ERR.VALUE;
  const c = sv(args, 0, ctx);
  if (isErr(c)) return c;
  const b = toBool(c === MISSING ? null : c);
  if (isErr(b)) return b;
  if (b) {
    if (args.length < 2) return true;
    const v = evaluate(args[1], ctx);
    return v === MISSING ? 0 : v;
  }
  if (args.length < 3) return false;
  const v = evaluate(args[2], ctx);
  return v === MISSING ? 0 : v;
};

FUNCS.IFERROR = (args, ctx) => {
  const v = sv(args, 0, ctx);
  if (isErr(v)) {
    const alt = ev(args, 1, ctx);
    if (alt === MISSING) return 0;
    return alt instanceof Ref ? scalar(alt, ctx) : alt;
  }
  return v === MISSING ? 0 : v;
};

FUNCS.IFNA = (args, ctx) => {
  const v = sv(args, 0, ctx);
  if (v === ERR.NA) {
    const alt = ev(args, 1, ctx);
    if (alt === MISSING) return 0;
    return alt instanceof Ref ? scalar(alt, ctx) : alt;
  }
  return v === MISSING ? 0 : v;
};

FUNCS.AND = (args, ctx) => {
  const l = collectLogicals(args, ctx);
  if (isErr(l)) return l;
  if (!l.length) return ERR.VALUE;
  return l.every(Boolean);
};

FUNCS.OR = (args, ctx) => {
  const l = collectLogicals(args, ctx);
  if (isErr(l)) return l;
  if (!l.length) return ERR.VALUE;
  return l.some(Boolean);
};

FUNCS.NOT = (args, ctx) => {
  const v = sv(args, 0, ctx);
  if (isErr(v)) return v;
  const b = toBool(v === MISSING ? null : v);
  return isErr(b) ? b : !b;
};

FUNCS.TRUE = () => true;
FUNCS.FALSE = () => false;

FUNCS.ISTEXT = (args, ctx) => typeof sv(args, 0, ctx) === 'string';
FUNCS.ISNUMBER = (args, ctx) => typeof sv(args, 0, ctx) === 'number';
FUNCS.ISLOGICAL = (args, ctx) => typeof sv(args, 0, ctx) === 'boolean';
FUNCS.ISERROR = (args, ctx) => isErr(sv(args, 0, ctx));
FUNCS.ISNA = (args, ctx) => sv(args, 0, ctx) === ERR.NA;
FUNCS.ISERR = (args, ctx) => { const v = sv(args, 0, ctx); return isErr(v) && v !== ERR.NA; };
FUNCS.ISNONTEXT = (args, ctx) => typeof sv(args, 0, ctx) !== 'string';
FUNCS.ISBLANK = (args, ctx) => {
  const v = ev(args, 0, ctx);
  if (!(v instanceof Ref)) return false;
  return scalar(v, ctx) === null;
};

FUNCS.SUM = (args, ctx) => {
  const ns = collectNumbers(args, ctx);
  if (isErr(ns)) return ns;
  let s = 0;
  for (const n of ns) s += n;
  return s;
};

FUNCS.MAX = (args, ctx) => {
  const ns = collectNumbers(args, ctx);
  if (isErr(ns)) return ns;
  return ns.length ? Math.max(...ns) : 0;
};

FUNCS.MIN = (args, ctx) => {
  const ns = collectNumbers(args, ctx);
  if (isErr(ns)) return ns;
  return ns.length ? Math.min(...ns) : 0;
};

FUNCS.AVERAGE = (args, ctx) => {
  const ns = collectNumbers(args, ctx);
  if (isErr(ns)) return ns;
  if (!ns.length) return ERR.DIV0;
  let s = 0;
  for (const n of ns) s += n;
  return s / ns.length;
};

FUNCS.COUNT = (args, ctx) => {
  let n = 0;
  for (let i = 0; i < args.length; i++) {
    const raw = evaluate(args[i], ctx);
    if (raw instanceof Ref || raw instanceof Arr) eachValue(raw, ctx, (x) => { if (typeof x === 'number') n++; return false; });
    else if (typeof raw === 'number' || typeof raw === 'boolean' || (typeof raw === 'string' && !isErr(toNumber(raw)))) n++;
  }
  return n;
};

FUNCS.COUNTA = (args, ctx) => {
  let n = 0;
  for (let i = 0; i < args.length; i++) {
    const raw = evaluate(args[i], ctx);
    if (raw instanceof Ref || raw instanceof Arr) eachValue(raw, ctx, (x) => { if (x !== null) n++; return false; });
    else if (raw !== MISSING) n++;
  }
  return n;
};

FUNCS.COUNTBLANK = (args, ctx) => {
  const raw = ev(args, 0, ctx);
  if (!(raw instanceof Ref)) return ERR.VALUE;
  let n = 0;
  for (let r = raw.r1; r <= raw.r2; r++) for (let c = raw.c1; c <= raw.c2; c++) {
    const x = cellValue(ctx, raw.sheet, r, c);
    if (x === null || x === '') n++;
  }
  return n;
};

function criteriaMatcher(crit) {
  if (typeof crit === 'number' || typeof crit === 'boolean') return exactMatcher(crit);
  if (crit === null) return (v) => v === null || v === '';
  const s = String(crit);
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s);
  const op = m[1] || '=';
  const rest = m[2];
  let target = rest;
  const n = toNumber(rest);
  if (rest !== '' && !isErr(n)) target = n;
  else if (/^true$/i.test(rest)) target = true;
  else if (/^false$/i.test(rest)) target = false;
  if (op === '=') {
    if (rest === '') return (v) => v === null || v === '';
    return exactMatcher(target);
  }
  if (op === '<>') {
    if (rest === '') return (v) => v !== null && v !== '';
    const eq = exactMatcher(target);
    return (v) => !eq(v);
  }
  return (v) => {
    if (typeClass(v) !== typeClass(target)) return false;
    const c = typeof v === 'string' ? compareText(v, target) : compareValues(v, target);
    return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
  };
}

FUNCS.COUNTIF = (args, ctx) => {
  const rng = ev(args, 0, ctx);
  if (!(rng instanceof Ref) && !(rng instanceof Arr)) return ERR.VALUE;
  const crit = sv(args, 1, ctx);
  if (isErr(crit)) return crit;
  const m = criteriaMatcher(crit === MISSING ? null : crit);
  let n = 0;
  if (rng instanceof Ref) {
    for (let r = rng.r1; r <= Math.min(rng.r2, ctx.eng.sheet(rng.sheet).maxRow || 0); r++) for (let c = rng.c1; c <= rng.c2; c++) if (m(cellValue(ctx, rng.sheet, r, c))) n++;
  } else eachValue(rng, ctx, (x) => { if (m(x)) n++; return false; });
  return n;
};

FUNCS.SUMIF = (args, ctx) => {
  const rng = ev(args, 0, ctx);
  if (!(rng instanceof Ref)) return ERR.VALUE;
  const crit = sv(args, 1, ctx);
  if (isErr(crit)) return crit;
  const sumR = args.length > 2 ? ev(args, 2, ctx) : rng;
  if (!(sumR instanceof Ref)) return ERR.VALUE;
  const m = criteriaMatcher(crit === MISSING ? null : crit);
  let s = 0;
  for (let r = rng.r1; r <= Math.min(rng.r2, ctx.eng.sheet(rng.sheet).maxRow || 0); r++) {
    for (let c = rng.c1; c <= rng.c2; c++) {
      if (m(cellValue(ctx, rng.sheet, r, c))) {
        const x = cellValue(ctx, sumR.sheet, sumR.r1 + (r - rng.r1), sumR.c1 + (c - rng.c1));
        if (typeof x === 'number') s += x;
        else if (isErr(x)) return x;
      }
    }
  }
  return s;
};

FUNCS.INT = (args, ctx) => { const v = num(args, 0, ctx); return isErr(v) ? v : Math.floor(v); };
FUNCS.ABS = (args, ctx) => { const v = num(args, 0, ctx); return isErr(v) ? v : Math.abs(v); };
FUNCS.SIGN = (args, ctx) => { const v = num(args, 0, ctx); return isErr(v) ? v : Math.sign(v); };
FUNCS.SQRT = (args, ctx) => { const v = num(args, 0, ctx); if (isErr(v)) return v; return v < 0 ? ERR.NUM : Math.sqrt(v); };
FUNCS.PI = () => Math.PI;

FUNCS.MOD = (args, ctx) => {
  const a = num(args, 0, ctx); if (isErr(a)) return a;
  const b = num(args, 1, ctx); if (isErr(b)) return b;
  if (b === 0) return ERR.DIV0;
  const r = a - b * Math.floor(a / b);
  return roundSig(r);
};

function roundTo(x, d, mode) {
  const f = Math.pow(10, d);
  const y = roundSig(x * f);
  let z;
  if (mode === 'round') z = Math.sign(y) * Math.round(Math.abs(y));
  else if (mode === 'up') z = Math.sign(y) * Math.ceil(Math.abs(y));
  else z = Math.sign(y) * Math.floor(Math.abs(y));
  return roundSig(z / f);
}

FUNCS.ROUND = (args, ctx) => { const a = num(args, 0, ctx); if (isErr(a)) return a; const d = num(args, 1, ctx); if (isErr(d)) return d; return roundTo(a, Math.trunc(d), 'round'); };
FUNCS.ROUNDUP = (args, ctx) => { const a = num(args, 0, ctx); if (isErr(a)) return a; const d = num(args, 1, ctx); if (isErr(d)) return d; return roundTo(a, Math.trunc(d), 'up'); };
FUNCS.ROUNDDOWN = (args, ctx) => { const a = num(args, 0, ctx); if (isErr(a)) return a; const d = num(args, 1, ctx); if (isErr(d)) return d; return roundTo(a, Math.trunc(d), 'down'); };
FUNCS.TRUNC = (args, ctx) => { const a = num(args, 0, ctx); if (isErr(a)) return a; const d = num(args, 1, ctx, 0); if (isErr(d)) return d; return roundTo(a, Math.trunc(d), 'down'); };

FUNCS.CEILING = (args, ctx) => {
  const x = num(args, 0, ctx); if (isErr(x)) return x;
  if (args.length < 2) return ERR.NA; // Excel requires the significance argument
  const s = num(args, 1, ctx); if (isErr(s)) return s;
  if (s === 0) return 0;
  if (x > 0 && s < 0) return ERR.NUM;
  const q = roundSig(x / s);
  return roundSig(Math.ceil(q) * s);
};

FUNCS.FLOOR = (args, ctx) => {
  const x = num(args, 0, ctx); if (isErr(x)) return x;
  if (args.length < 2) return ERR.NA;
  const s = num(args, 1, ctx); if (isErr(s)) return s;
  if (s === 0) return x === 0 ? 0 : ERR.DIV0;
  if (x > 0 && s < 0) return ERR.NUM;
  const q = roundSig(x / s);
  return roundSig(Math.floor(q) * s);
};

FUNCS.CHAR = (args, ctx) => {
  const n = num(args, 0, ctx); if (isErr(n)) return n;
  const k = Math.trunc(n);
  if (k < 1 || k > 255) return ERR.VALUE;
  return String.fromCharCode(k);
};

FUNCS.CODE = (args, ctx) => {
  const v = sv(args, 0, ctx); if (isErr(v)) return v;
  const s = toText(v === MISSING ? null : v);
  if (!s.length) return ERR.VALUE;
  return s.charCodeAt(0);
};

FUNCS.LEN = (args, ctx) => { const v = sv(args, 0, ctx); if (isErr(v)) return v; return toText(v === MISSING ? null : v).length; };

function textArg(args, i, ctx) {
  const v = sv(args, i, ctx);
  if (isErr(v)) return v;
  return toText(v === MISSING ? null : v);
}

FUNCS.LEFT = (args, ctx) => { const s = textArg(args, 0, ctx); if (isErr(s)) return s; const n = num(args, 1, ctx, 1); if (isErr(n)) return n; if (n < 0) return ERR.VALUE; return s.slice(0, Math.trunc(n)); };
FUNCS.RIGHT = (args, ctx) => { const s = textArg(args, 0, ctx); if (isErr(s)) return s; const n = num(args, 1, ctx, 1); if (isErr(n)) return n; if (n < 0) return ERR.VALUE; const k = Math.trunc(n); return k === 0 ? '' : s.slice(-k); };
FUNCS.MID = (args, ctx) => { const s = textArg(args, 0, ctx); if (isErr(s)) return s; const a = num(args, 1, ctx); if (isErr(a)) return a; const n = num(args, 2, ctx); if (isErr(n)) return n; if (a < 1 || n < 0) return ERR.VALUE; return s.substr(Math.trunc(a) - 1, Math.trunc(n)); };
FUNCS.CONCATENATE = (args, ctx) => { let out = ''; for (let i = 0; i < args.length; i++) { const s = textArg(args, i, ctx); if (isErr(s)) return s; out += s; } return out; };
FUNCS.TRIM = (args, ctx) => { const s = textArg(args, 0, ctx); if (isErr(s)) return s; return s.replace(/ +/g, ' ').trim(); };
FUNCS.UPPER = (args, ctx) => { const s = textArg(args, 0, ctx); return isErr(s) ? s : s.toUpperCase(); };
FUNCS.LOWER = (args, ctx) => { const s = textArg(args, 0, ctx); return isErr(s) ? s : s.toLowerCase(); };
FUNCS.REPT = (args, ctx) => { const s = textArg(args, 0, ctx); if (isErr(s)) return s; const n = num(args, 1, ctx); if (isErr(n)) return n; if (n < 0) return ERR.VALUE; return s.repeat(Math.trunc(n)); };
FUNCS.VALUE = (args, ctx) => { const v = sv(args, 0, ctx); if (isErr(v)) return v; if (typeof v === 'number') return v; return toNumber(toText(v === MISSING ? null : v)); };
FUNCS.N = (args, ctx) => { const v = sv(args, 0, ctx); if (isErr(v)) return v; if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? 1 : 0; return 0; };
FUNCS.T = (args, ctx) => { const v = sv(args, 0, ctx); if (isErr(v)) return v; return typeof v === 'string' ? v : ''; };

FUNCS.RAND = () => Math.random();
FUNCS.RANDBETWEEN = (args, ctx) => {
  const a = num(args, 0, ctx); if (isErr(a)) return a;
  const b = num(args, 1, ctx); if (isErr(b)) return b;
  const lo = Math.ceil(a), hi = Math.floor(b);
  if (lo > hi) return ERR.NUM;
  return lo + Math.floor(Math.random() * (hi - lo + 1));
};

FUNCS.CHOOSE = (args, ctx) => {
  const i = num(args, 0, ctx); if (isErr(i)) return i;
  const k = Math.trunc(i);
  if (k < 1 || k >= args.length) return ERR.VALUE;
  const v = evaluate(args[k], ctx);
  return v === MISSING ? 0 : v;
};

FUNCS.ROW = (args, ctx) => { if (!args.length) return ctx.r; const v = ev(args, 0, ctx); return v instanceof Ref ? v.r1 : ERR.VALUE; };
FUNCS.COLUMN = (args, ctx) => { if (!args.length) return ctx.c; const v = ev(args, 0, ctx); return v instanceof Ref ? v.c1 : ERR.VALUE; };
FUNCS.ROWS = (args, ctx) => { const v = ev(args, 0, ctx); return v instanceof Ref || v instanceof Arr ? v.rows : 1; };
FUNCS.COLUMNS = (args, ctx) => { const v = ev(args, 0, ctx); return v instanceof Ref || v instanceof Arr ? v.cols : 1; };

FUNCS.VLOOKUP = (args, ctx) => lookupHV(args, ctx, true);
FUNCS.HLOOKUP = (args, ctx) => lookupHV(args, ctx, false);

function lookupHV(args, ctx, vertical) {
  if (args.length < 3) return ERR.VALUE;
  let target = sv(args, 0, ctx);
  if (isErr(target)) return target;
  target = lookupTarget(target);
  const table = ev(args, 1, ctx);
  if (isErr(table)) return table;
  const idx = num(args, 2, ctx);
  if (isErr(idx)) return idx;
  let approx = true;
  if (args.length > 3) {
    const a = sv(args, 3, ctx);
    if (isErr(a)) return a;
    if (a !== MISSING) {
      const b = toBool(a);
      if (isErr(b)) return b;
      approx = b;
    } else approx = false; // VLOOKUP(x,t,2,) -> exact
  }
  const k = Math.trunc(idx);
  if (k < 1) return ERR.VALUE;
  let len, get, result, scan;
  if (table instanceof Ref) {
    const sh = ctx.eng.sheet(table.sheet);
    if (vertical) {
      if (k > table.cols) return ERR.REF;
      len = table.rows;
      get = (i) => cellValue(ctx, table.sheet, table.r1 + i, table.c1);
      result = (i) => cellValue(ctx, table.sheet, table.r1 + i, table.c1 + k - 1);
      scan = Math.max(0, Math.min(len, (sh ? sh.maxRow : 0) - table.r1 + 1));
    } else {
      if (k > table.rows) return ERR.REF;
      len = table.cols;
      get = (i) => cellValue(ctx, table.sheet, table.r1, table.c1 + i);
      result = (i) => cellValue(ctx, table.sheet, table.r1 + k - 1, table.c1 + i);
      scan = Math.max(0, Math.min(len, (sh ? sh.maxCol : 0) - table.c1 + 1));
    }
  } else if (table instanceof Arr) {
    if (vertical) {
      if (k > table.cols) return ERR.REF;
      len = table.rows; get = (i) => table.get(i, 0); result = (i) => table.get(i, k - 1);
    } else {
      if (k > table.rows) return ERR.REF;
      len = table.cols; get = (i) => table.get(0, i); result = (i) => table.get(k - 1, i);
    }
    scan = len;
  } else return ERR.NA;
  let pos = -1;
  if (approx) pos = approxSearch(len, get, target);
  else {
    const eq = exactMatcher(target);
    for (let i = 0; i < scan; i++) if (eq(get(i))) { pos = i; break; }
  }
  if (pos < 0) return ERR.NA;
  return result(pos);
}

FUNCS.MATCH = (args, ctx) => {
  if (args.length < 2) return ERR.VALUE;
  let target = sv(args, 0, ctx);
  if (isErr(target)) return target;
  target = lookupTarget(target);
  const arr = ev(args, 1, ctx);
  if (isErr(arr)) return arr;
  let type = 1;
  if (args.length > 2) {
    const t = num(args, 2, ctx, 1);
    if (isErr(t)) return t;
    type = t > 0 ? 1 : t < 0 ? -1 : 0;
  }
  const vec = vectorOf(arr, ctx);
  if (!vec) return ERR.NA;
  let pos = -1;
  if (type === 0) {
    const eq = exactMatcher(target);
    for (let i = 0; i < vec.scanLen; i++) if (eq(vec.get(i))) { pos = i; break; }
  } else pos = approxSearch(vec.len, vec.get, target, type < 0);
  return pos < 0 ? ERR.NA : pos + 1;
};

FUNCS.LOOKUP = (args, ctx) => {
  if (args.length < 2) return ERR.VALUE;
  let target = sv(args, 0, ctx);
  if (isErr(target)) return target;
  target = lookupTarget(target);
  const lv = ev(args, 1, ctx);
  if (isErr(lv)) return lv;
  if (args.length >= 3) {
    const vec = vectorOf(lv, ctx, true);
    const rv = ev(args, 2, ctx);
    if (isErr(rv)) return rv;
    const res = vectorOf(rv, ctx, true);
    if (!vec || !res) return ERR.NA;
    const pos = approxSearch(vec.len, vec.get, target);
    if (pos < 0) return ERR.NA;
    if (pos >= res.len) {
      if (rv instanceof Ref) {
        // Excel reads past the end of a short result range
        return res.vertical ? cellValue(ctx, rv.sheet, rv.r1 + pos, rv.c1) : cellValue(ctx, rv.sheet, rv.r1, rv.c1 + pos);
      }
      return ERR.NA;
    }
    return res.get(pos);
  }
  // array form
  let rows, cols, get;
  if (lv instanceof Ref) {
    rows = lv.rows; cols = lv.cols;
    get = (r, c) => cellValue(ctx, lv.sheet, lv.r1 + r, lv.c1 + c);
  } else if (lv instanceof Arr) {
    rows = lv.rows; cols = lv.cols;
    get = (r, c) => lv.get(r, c);
  } else {
    return compareValues(lv, target) <= 0 && typeClass(lv) === typeClass(target) ? lv : ERR.NA;
  }
  if (cols > rows) {
    const pos = approxSearch(cols, (i) => get(0, i), target);
    return pos < 0 ? ERR.NA : get(rows - 1, pos);
  }
  const pos = approxSearch(rows, (i) => get(i, 0), target);
  return pos < 0 ? ERR.NA : get(pos, cols - 1);
};

FUNCS.INDEX = (args, ctx) => {
  const a = ev(args, 0, ctx);
  if (isErr(a)) return a;
  let rn = args.length > 1 ? num(args, 1, ctx, 0) : 0;
  if (isErr(rn)) return rn;
  const colGiven = args.length > 2 && args[2].type !== 'missing';
  let cn = colGiven ? num(args, 2, ctx, 0) : 0;
  if (isErr(cn)) return cn;
  rn = Math.trunc(rn); cn = Math.trunc(cn);
  if (rn < 0 || cn < 0) return ERR.VALUE;
  const rows = a instanceof Ref || a instanceof Arr ? a.rows : 1;
  const cols = a instanceof Ref || a instanceof Arr ? a.cols : 1;
  let R = rn, C = cn;
  if (!colGiven) {
    if (rows === 1 && cols > 1) { C = rn; R = 0; }
    else if (cols === 1) C = 0;
  }
  if (R > rows || C > cols) return ERR.REF;
  if (a instanceof Ref) {
    const r1 = R ? a.r1 + R - 1 : a.r1, r2 = R ? a.r1 + R - 1 : a.r2;
    const c1 = C ? a.c1 + C - 1 : a.c1, c2 = C ? a.c1 + C - 1 : a.c2;
    return new Ref(a.sheet, r1, c1, r2, c2);
  }
  if (a instanceof Arr) {
    const rr = R ? R - 1 : 0, cc = C ? C - 1 : 0;
    if ((!R && rows > 1) || (!C && cols > 1)) {
      const data = [];
      for (let i = R ? R - 1 : 0; i <= (R ? R - 1 : rows - 1); i++) {
        const row = [];
        for (let j = C ? C - 1 : 0; j <= (C ? C - 1 : cols - 1); j++) row.push(a.get(i, j));
        data.push(row);
      }
      return new Arr(data);
    }
    return a.get(rr, cc);
  }
  return a;
};

FUNCS.TEXT = (args, ctx) => {
  const v = sv(args, 0, ctx); if (isErr(v)) return v;
  const f = textArg(args, 1, ctx); if (isErr(f)) return f;
  return ctx.eng.formatText(v === MISSING ? null : v, f);
};

FUNCS.EXACT = (args, ctx) => {
  const a = textArg(args, 0, ctx); if (isErr(a)) return a;
  const b = textArg(args, 1, ctx); if (isErr(b)) return b;
  return a === b;
};

FUNCS.FIND = (args, ctx) => {
  const f = textArg(args, 0, ctx); if (isErr(f)) return f;
  const s = textArg(args, 1, ctx); if (isErr(s)) return s;
  const st = num(args, 2, ctx, 1); if (isErr(st)) return st;
  const i = s.indexOf(f, Math.trunc(st) - 1);
  return i < 0 ? ERR.VALUE : i + 1;
};

FUNCS.SEARCH = (args, ctx) => {
  const f = textArg(args, 0, ctx); if (isErr(f)) return f;
  const s = textArg(args, 1, ctx); if (isErr(s)) return s;
  const st = num(args, 2, ctx, 1); if (isErr(st)) return st;
  const re = wildcardRegex(f);
  const from = Math.trunc(st) - 1;
  if (!re) { const i = s.toLowerCase().indexOf(f.toLowerCase(), from); return i < 0 ? ERR.VALUE : i + 1; }
  for (let i = from; i <= s.length; i++) {
    for (let j = s.length; j >= i; j--) if (re.test(s.slice(i, j))) return i + 1;
  }
  return ERR.VALUE;
};

FUNCS.SUBSTITUTE = (args, ctx) => {
  const s = textArg(args, 0, ctx); if (isErr(s)) return s;
  const o = textArg(args, 1, ctx); if (isErr(o)) return o;
  const n = textArg(args, 2, ctx); if (isErr(n)) return n;
  if (!o) return s;
  if (args.length > 3) {
    const k = num(args, 3, ctx); if (isErr(k)) return k;
    let idx = -1;
    for (let i = 0; i < k; i++) { idx = s.indexOf(o, idx + 1); if (idx < 0) return s; }
    return s.slice(0, idx) + n + s.slice(idx + o.length);
  }
  return s.split(o).join(n);
};

export { FUNCS, MISSING, numToText };
