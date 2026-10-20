// A1-style address helpers. Rows and columns are 1-based everywhere.

export function colToNum(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) & 0x1f);
  return n;
}

export function numToCol(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}

export function addr(r, c) {
  return numToCol(c) + r;
}

export function parseAddr(a) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a);
  if (!m) return null;
  return { r: +m[2], c: colToNum(m[1].toUpperCase()) };
}

// "A1:B3" -> {r1,c1,r2,c2}; single cell allowed.
export function parseRange(s) {
  const parts = s.split(':');
  const a = parseAddr(parts[0]);
  if (!a) return null;
  const b = parts.length > 1 ? parseAddr(parts[1]) : a;
  if (!b) return null;
  return {
    r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c),
  };
}

// Space separated list of ranges (sqref).
export function parseSqref(s) {
  return s.trim().split(/\s+/).map(parseRange).filter(Boolean);
}

export function rangeToStr(g) {
  const a = addr(g.r1, g.c1);
  if (g.r1 === g.r2 && g.c1 === g.c2) return a;
  return a + ':' + addr(g.r2, g.c2);
}

export const MAXCOL = 16384;
export const MAXROW = 1048576;

export function key(r, c) {
  return (r - 1) * MAXCOL + (c - 1);
}

export function keyRC(k) {
  return { r: Math.floor(k / MAXCOL) + 1, c: (k % MAXCOL) + 1 };
}

export function inRange(g, r, c) {
  return r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;
}

export function quoteSheet(name) {
  return /^[A-Za-z_\u0080-￿][\w.\u0080-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name
    : "'" + name.replace(/'/g, "''") + "'";
}
