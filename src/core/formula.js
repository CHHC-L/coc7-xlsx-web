// Excel formula tokenizer / parser producing a small AST, plus helpers to
// translate shared formulas and print an AST back to text.

import { colToNum, numToCol, MAXCOL, MAXROW, quoteSheet } from './refs.js';
import { errFromCode } from './values.js';

const ERR_RE = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|CALC!|SPILL!)/i;
const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;
const COL_RE = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_RE = /^(\$?)(\d{1,7})$/;

function isWordChar(ch) {
  const c = ch.charCodeAt(0);
  return (c >= 48 && c <= 57) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c === 95 || c === 46 || c === 36 || c === 92 || c === 63 || c > 127;
}

function cellPart(word) {
  const m = CELL_RE.exec(word);
  if (!m) return null;
  const c = colToNum(m[2].toUpperCase());
  const r = +m[4];
  if (c > MAXCOL || r < 1 || r > MAXROW) return null;
  return { r, c, ar: !!m[3], ac: !!m[1] };
}

export class FormulaError extends Error {}

export function tokenize(src) {
  const toks = [];
  const n = src.length;
  let i = 0;

  const readWord = (j) => {
    let k = j;
    while (k < n && isWordChar(src[k])) k++;
    return k;
  };

  // after an optional sheet prefix, read a reference (cell, area, cols, rows) or a name
  const readRefOrName = (j, sheet) => {
    if (src[j] === '#') {
      const m = ERR_RE.exec(src.slice(j));
      if (m) return { tok: { t: 'err', v: m[0].toUpperCase(), sheet }, next: j + m[0].length };
    }
    const k = readWord(j);
    const w = src.slice(j, k);
    if (!w) throw new FormulaError('引用解析失败: ' + src);
    // area "A1:B2"
    const a = cellPart(w);
    if (a) {
      if (src[k] === ':') {
        const k2 = readWord(k + 1);
        const b = cellPart(src.slice(k + 1, k2));
        if (b) return { tok: { t: 'ref', sheet, kind: 'area', a, b }, next: k2 };
      }
      return { tok: { t: 'ref', sheet, kind: 'cell', a }, next: k };
    }
    let m;
    if ((m = COL_RE.exec(w)) && src[k] === ':') {
      const k2 = readWord(k + 1);
      const m2 = COL_RE.exec(src.slice(k + 1, k2));
      if (m2) {
        return {
          tok: { t: 'ref', sheet, kind: 'cols',
            a: { r: 1, c: colToNum(m[2].toUpperCase()), ar: true, ac: !!m[1] },
            b: { r: MAXROW, c: colToNum(m2[2].toUpperCase()), ar: true, ac: !!m2[1] } },
          next: k2,
        };
      }
    }
    if ((m = ROW_RE.exec(w)) && src[k] === ':') {
      const k2 = readWord(k + 1);
      const m2 = ROW_RE.exec(src.slice(k + 1, k2));
      if (m2) {
        return {
          tok: { t: 'ref', sheet, kind: 'rows',
            a: { r: +m[2], c: 1, ar: !!m[1], ac: true },
            b: { r: +m2[2], c: MAXCOL, ar: !!m2[1], ac: true } },
          next: k2,
        };
      }
    }
    return { tok: { t: 'name', v: w, sheet }, next: k };
  };

  while (i < n) {
    const ch = src[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') {
      let j = i;
      while (j < n && /\s/.test(src[j])) j++;
      toks.push({ t: 'ws' });
      i = j;
      continue;
    }
    if (ch === '"') {
      let j = i + 1, s = '';
      for (;;) {
        if (j >= n) throw new FormulaError('字符串未闭合: ' + src);
        if (src[j] === '"') {
          if (src[j + 1] === '"') { s += '"'; j += 2; continue; }
          break;
        }
        s += src[j++];
      }
      toks.push({ t: 'str', v: s });
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const m = ERR_RE.exec(src.slice(i));
      if (!m) throw new FormulaError('未知错误值: ' + src);
      toks.push({ t: 'err', v: m[0].toUpperCase() });
      i += m[0].length;
      continue;
    }
    if (ch === "'") {
      let j = i + 1, s = '';
      for (;;) {
        if (j >= n) throw new FormulaError('工作表名未闭合: ' + src);
        if (src[j] === "'") {
          if (src[j + 1] === "'") { s += "'"; j += 2; continue; }
          break;
        }
        s += src[j++];
      }
      if (src[j + 1] !== '!') throw new FormulaError('缺少 ! : ' + src);
      const { tok, next } = readRefOrName(j + 2, s);
      toks.push(tok);
      i = next;
      continue;
    }
    if ((ch >= '0' && ch <= '9') || (ch === '.' && /[0-9]/.test(src[i + 1] || ''))) {
      // row range like 1:3 ?
      const rm = /^(\d+):(\$?\d+)/.exec(src.slice(i));
      if (rm && (i === 0 || !isWordChar(src[i - 1]))) {
        const { tok, next } = readRefOrName(i, null);
        if (tok.t === 'ref') { toks.push(tok); i = next; continue; }
      }
      const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i));
      toks.push({ t: 'num', v: parseFloat(m[0]) });
      i += m[0].length;
      continue;
    }
    if (isWordChar(ch)) {
      const k = readWord(i);
      const w = src.slice(i, k);
      if (src[k] === '!') {
        const { tok, next } = readRefOrName(k + 1, w);
        toks.push(tok);
        i = next;
        continue;
      }
      if (src[k] === '(') {
        toks.push({ t: 'fn', v: w.toUpperCase() });
        i = k + 1;
        continue;
      }
      const up = w.toUpperCase();
      if ((up === 'TRUE' || up === 'FALSE') && !cellPart(w)) {
        toks.push({ t: 'bool', v: up === 'TRUE' });
        i = k;
        continue;
      }
      const { tok, next } = readRefOrName(i, null);
      toks.push(tok);
      i = next;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      toks.push({ t: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%:(),;{}'.indexOf(ch) >= 0) {
      toks.push({ t: 'op', v: ch });
      i++;
      continue;
    }
    throw new FormulaError(`无法识别的字符 "${ch}": ${src}`);
  }
  // drop whitespace tokens; an intersection operator (ref ws ref) is not supported
  const out = [];
  for (let k = 0; k < toks.length; k++) {
    const t = toks[k];
    if (t.t !== 'ws') { out.push(t); continue; }
    const prev = out[out.length - 1];
    const next = toks[k + 1];
    if (prev && next && (prev.t === 'ref' || (prev.t === 'op' && prev.v === ')')) && (next.t === 'ref' || next.t === 'name' || next.t === 'fn')) {
      throw new FormulaError('不支持交集运算符: ' + src);
    }
  }
  return out;
}

const BIN_PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5, ':': 8 };

export function parseFormula(src) {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const isOp = (v) => toks[p] && toks[p].t === 'op' && toks[p].v === v;
  const expectOp = (v) => {
    if (!isOp(v)) throw new FormulaError(`缺少 "${v}": ${src}`);
    p++;
  };

  function refNode(t) {
    const a = t.a, b = t.b || t.a;
    return {
      type: 'ref', sheet: t.sheet || null, kind: t.kind,
      r1: a.r, c1: a.c, ar1: a.ar, ac1: a.ac,
      r2: b.r, c2: b.c, ar2: b.ar, ac2: b.ac,
    };
  }

  function constant() {
    const t = toks[p];
    if (!t) throw new FormulaError('数组常量不完整: ' + src);
    if (t.t === 'op' && (t.v === '-' || t.v === '+')) {
      p++;
      const n = toks[p++];
      if (!n || n.t !== 'num') throw new FormulaError('数组常量错误: ' + src);
      return { type: 'num', v: t.v === '-' ? -n.v : n.v };
    }
    p++;
    if (t.t === 'num') return { type: 'num', v: t.v };
    if (t.t === 'str') return { type: 'str', v: t.v };
    if (t.t === 'bool') return { type: 'bool', v: t.v };
    if (t.t === 'err') return { type: 'err', v: errFromCode(t.v) };
    throw new FormulaError('数组常量错误: ' + src);
  }

  function primary() {
    const t = toks[p];
    if (!t) throw new FormulaError('公式不完整: ' + src);
    p++;
    switch (t.t) {
      case 'num': return { type: 'num', v: t.v };
      case 'str': return { type: 'str', v: t.v };
      case 'bool': return { type: 'bool', v: t.v };
      case 'err':
        if (t.sheet) return { type: 'err', v: errFromCode(t.v) };
        return { type: 'err', v: errFromCode(t.v) };
      case 'ref': return refNode(t);
      case 'name': return { type: 'name', name: t.v, sheet: t.sheet || null };
      case 'fn': {
        const args = [];
        if (isOp(')')) { p++; return { type: 'fn', name: t.v, args }; }
        for (;;) {
          if (isOp(',') || isOp(')')) args.push({ type: 'missing' });
          else args.push(expr(1));
          if (isOp(',')) { p++; continue; }
          expectOp(')');
          break;
        }
        return { type: 'fn', name: t.v, args };
      }
      case 'op':
        if (t.v === '(') {
          const e = expr(1);
          expectOp(')');
          return { type: 'paren', a: e };
        }
        if (t.v === '{') {
          const rows = [[]];
          for (;;) {
            rows[rows.length - 1].push(constant());
            if (isOp(',')) { p++; continue; }
            if (isOp(';')) { p++; rows.push([]); continue; }
            expectOp('}');
            break;
          }
          return { type: 'arr', rows };
        }
        break;
      default:
    }
    throw new FormulaError(`意外的符号 "${t.v}": ${src}`);
  }

  function unary() {
    if (isOp('-') || isOp('+')) {
      const op = toks[p++].v;
      return { type: 'un', op, a: unary() };
    }
    let e = primary();
    while (isOp('%')) { p++; e = { type: 'pct', a: e }; }
    return e;
  }

  function expr(minPrec) {
    let left = unary();
    for (;;) {
      const t = peek();
      if (!t || t.t !== 'op') break;
      const prec = BIN_PREC[t.v];
      if (prec === undefined || prec < minPrec) break;
      p++;
      const right = prec === 8 ? unary() : expr(prec + 1);
      left = { type: 'bin', op: t.v, a: left, b: right };
      if (prec !== 8) {
        // postfix % may follow a parenthesised group etc. already handled in unary()
      }
    }
    return left;
  }

  const ast = expr(1);
  if (p !== toks.length) throw new FormulaError('公式末尾有多余内容: ' + src);
  return ast;
}

// Translate relative references by (dr, dc) — used for shared formulas.
export function shiftAst(node, dr, dc) {
  switch (node.type) {
    case 'ref': {
      const n = { ...node };
      if (n.kind !== 'cols') {
        if (!n.ar1) n.r1 += dr;
        if (!n.ar2) n.r2 += dr;
      }
      if (n.kind !== 'rows') {
        if (!n.ac1) n.c1 += dc;
        if (!n.ac2) n.c2 += dc;
      }
      if (n.r1 < 1 || n.c1 < 1 || n.r2 > MAXROW || n.c2 > MAXCOL || n.r2 < 1 || n.c2 < 1) {
        return { type: 'err', v: errFromCode('#REF!') };
      }
      return n;
    }
    case 'fn': return { ...node, args: node.args.map((a) => shiftAst(a, dr, dc)) };
    case 'un': case 'pct': case 'paren': return { ...node, a: shiftAst(node.a, dr, dc) };
    case 'bin': return { ...node, a: shiftAst(node.a, dr, dc), b: shiftAst(node.b, dr, dc) };
    default: return node;
  }
}

function fmtNum(v) {
  let s = String(v);
  if (s.indexOf('e') >= 0) s = s.replace('e', 'E');
  return s;
}

function refStr(n) {
  const pre = n.sheet ? quoteSheet(n.sheet) + '!' : '';
  const cell = (r, c, ar, ac) => (ac ? '$' : '') + numToCol(c) + (ar ? '$' : '') + r;
  if (n.kind === 'cols') return pre + (n.ac1 ? '$' : '') + numToCol(n.c1) + ':' + (n.ac2 ? '$' : '') + numToCol(n.c2);
  if (n.kind === 'rows') return pre + (n.ar1 ? '$' : '') + n.r1 + ':' + (n.ar2 ? '$' : '') + n.r2;
  const a = cell(n.r1, n.c1, n.ar1, n.ac1);
  if (n.kind === 'cell') return pre + a;
  return pre + a + ':' + cell(n.r2, n.c2, n.ar2, n.ac2);
}

export function astToString(node) {
  switch (node.type) {
    case 'num': return fmtNum(node.v);
    case 'str': return '"' + node.v.replace(/"/g, '""') + '"';
    case 'bool': return node.v ? 'TRUE' : 'FALSE';
    case 'err': return node.v.code;
    case 'ref': return refStr(node);
    case 'name': return (node.sheet ? quoteSheet(node.sheet) + '!' : '') + node.name;
    case 'missing': return '';
    case 'fn': return node.name + '(' + node.args.map(astToString).join(',') + ')';
    case 'paren': return '(' + astToString(node.a) + ')';
    case 'un': return node.op + astToString(node.a);
    case 'pct': return astToString(node.a) + '%';
    case 'bin': return astToString(node.a) + node.op + astToString(node.b);
    case 'arr': return '{' + node.rows.map((r) => r.map(astToString).join(',')).join(';') + '}';
    default: return '?';
  }
}
