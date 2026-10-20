// Worksheet protection password verification (ECMA-376 legacy hash and the
// agile SHA-family hash with salt + spin count).

export function legacyHash(pw) {
  // Excel hashes the password's ANSI bytes; non-ASCII characters are rare here.
  const bytes = [];
  for (const ch of pw) {
    const c = ch.charCodeAt(0);
    bytes.push(c < 256 ? c : (c & 0xff) || (c >> 8));
  }
  let h = 0;
  for (let i = bytes.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
    h ^= bytes[i];
  }
  h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
  h ^= bytes.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
}

function b64ToBytes(s) {
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function bytesToB64(b) {
  let s = '';
  for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]);
  return btoa(s);
}

const ALG = { 'SHA-512': 'SHA-512', 'SHA-384': 'SHA-384', 'SHA-256': 'SHA-256', 'SHA-1': 'SHA-1' };

// attrs: the <sheetProtection> attributes. Resolves true if pw unlocks it.
export async function checkPassword(attrs, pw, subtle = globalThis.crypto && globalThis.crypto.subtle) {
  if (attrs.password) return legacyHash(pw) === attrs.password.toUpperCase();
  if (attrs.hashValue && attrs.saltValue && attrs.algorithmName) {
    const alg = ALG[attrs.algorithmName.toUpperCase()] || null;
    if (!alg || !subtle) return false;
    const salt = b64ToBytes(attrs.saltValue);
    const pwBytes = new Uint8Array(pw.length * 2);
    for (let i = 0; i < pw.length; i++) { const c = pw.charCodeAt(i); pwBytes[2 * i] = c & 0xff; pwBytes[2 * i + 1] = c >> 8; }
    let buf = new Uint8Array(salt.length + pwBytes.length);
    buf.set(salt); buf.set(pwBytes, salt.length);
    let h = new Uint8Array(await subtle.digest(alg, buf));
    const spin = +(attrs.spinCount || 0);
    const tmp = new Uint8Array(h.length + 4);
    for (let i = 0; i < spin; i++) {
      tmp.set(h, 0);
      tmp[h.length] = i & 0xff; tmp[h.length + 1] = (i >> 8) & 0xff; tmp[h.length + 2] = (i >> 16) & 0xff; tmp[h.length + 3] = (i >>> 24) & 0xff;
      h = new Uint8Array(await subtle.digest(alg, tmp));
    }
    return bytesToB64(h) === attrs.hashValue;
  }
  return pw === '';
}
