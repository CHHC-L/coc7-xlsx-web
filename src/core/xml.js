// Minimal, fast XML parser for OOXML parts.
// Keeps source offsets on every element so the exporter can splice the
// original text instead of re-serialising it (that is how we keep the file
// byte-identical outside the cells we actually change).

const ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

// XML end-of-line normalisation (CRLF / CR -> LF), as every XML reader does.
function normEol(s) {
  return s.indexOf('\r') < 0 ? s : s.replace(/\r\n?/g, '\n');
}

export function decodeEntities(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, e) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return String.fromCodePoint(cp);
    }
    return ENT[e] !== undefined ? ENT[e] : m;
  });
}

export function escapeText(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function escapeAttr(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// OOXML ST_Xstring escaping of characters XML cannot carry (_xHHHH_).
export function decodeXstring(s) {
  if (s.indexOf('_x') < 0) return s;
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
}

export function encodeXstring(s) {
  // A literal "_xHHHH_" must have its underscore escaped first.
  let out = s.replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1');
  // eslint-disable-next-line no-control-regex
  out = out.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/g, (ch) =>
    '_x' + ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') + '_');
  return out;
}

export class XNode {
  constructor(name, start) {
    this.name = name; // qualified name as written, e.g. "x14:dataValidation"
    this.attrs = {};
    this.children = [];
    this.text = '';
    this.start = start; // offset of '<'
    this.openEnd = 0; // offset just after the start tag's '>'
    this.end = 0; // offset just after the end tag (or self-closing tag)
    this.selfClosing = false;
  }
  get local() {
    const i = this.name.indexOf(':');
    return i < 0 ? this.name : this.name.slice(i + 1);
  }
  child(local) {
    for (const c of this.children) if (c.local === local) return c;
    return null;
  }
  all(local) {
    return this.children.filter((c) => c.local === local);
  }
  // Depth-first search for descendants with the given local name.
  find(local, out = []) {
    for (const c of this.children) {
      if (c.local === local) out.push(c);
      c.find(local, out);
    }
    return out;
  }
  attr(name, dflt = undefined) {
    const v = this.attrs[name];
    return v === undefined ? dflt : v;
  }
  // Concatenated text of this node and its descendants.
  deepText() {
    if (!this.children.length) return this.text;
    let s = this.text;
    for (const c of this.children) s += c.deepText();
    return s;
  }
}

const ATTR_RE = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

export function parseXml(src) {
  const root = new XNode('#document', 0);
  const stack = [root];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt < 0) break;
    if (lt > i) {
      const top = stack[stack.length - 1];
      if (top !== root) top.text += decodeEntities(normEol(src.slice(i, lt)));
    }
    const c1 = src.charCodeAt(lt + 1);
    if (c1 === 63 /* ? */) {
      i = src.indexOf('?>', lt) + 2;
      continue;
    }
    if (c1 === 33 /* ! */) {
      if (src.startsWith('<!--', lt)) {
        i = src.indexOf('-->', lt) + 3;
      } else if (src.startsWith('<![CDATA[', lt)) {
        const e = src.indexOf(']]>', lt);
        stack[stack.length - 1].text += normEol(src.slice(lt + 9, e));
        i = e + 3;
      } else {
        i = src.indexOf('>', lt) + 1;
      }
      continue;
    }
    if (c1 === 47 /* / */) {
      const gt = src.indexOf('>', lt);
      const node = stack.pop();
      node.end = gt + 1;
      i = gt + 1;
      continue;
    }
    // start tag; find the closing '>' while respecting quoted attributes
    let j = lt + 1;
    let q = 0;
    for (; j < n; j++) {
      const ch = src.charCodeAt(j);
      if (q) {
        if (ch === q) q = 0;
      } else if (ch === 34 || ch === 39) q = ch;
      else if (ch === 62) break;
    }
    const selfClosing = src.charCodeAt(j - 1) === 47;
    const inner = src.slice(lt + 1, selfClosing ? j - 1 : j);
    let k = 0;
    while (k < inner.length && !/\s/.test(inner[k])) k++;
    const node = new XNode(inner.slice(0, k), lt);
    if (k < inner.length) {
      ATTR_RE.lastIndex = k;
      let m;
      while ((m = ATTR_RE.exec(inner))) {
        node.attrs[m[1]] = decodeEntities(normEol(m[3] !== undefined ? m[3] : m[4]));
      }
    }
    node.openEnd = j + 1;
    stack[stack.length - 1].children.push(node);
    if (selfClosing) {
      node.selfClosing = true;
      node.end = j + 1;
    } else {
      stack.push(node);
    }
    i = j + 1;
  }
  return root.children.find((c) => true) || root;
}
