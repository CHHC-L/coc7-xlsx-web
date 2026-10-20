// Zip container handling for .xlsx.
// Reading keeps every entry's raw (compressed) bytes and header fields, so
// that writing can copy untouched parts verbatim and give changed parts the
// same header shape Excel produces (deflate, flags 6, version 45/20,
// 1980-01-01 timestamps, entry order preserved).

import { inflateSync, deflateSync } from 'fflate';

const td = new TextDecoder('utf-8');
const te = new TextEncoder();

function u16(b, o) { return b[o] | (b[o + 1] << 8); }
function u32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }

let CRC_TABLE = null;
export function crc32(data) {
  if (!CRC_TABLE) {
    CRC_TABLE = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      CRC_TABLE[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) crc = CRC_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

export function readZip(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  // locate End Of Central Directory
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) {
    if (u32(b, i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('不是有效的 xlsx（zip）文件');
  const count = u16(b, eocd + 10);
  let p = u32(b, eocd + 16);
  const entries = [];
  for (let k = 0; k < count; k++) {
    if (u32(b, p) !== 0x02014b50) throw new Error('zip 中央目录损坏');
    const e = {
      versionMadeBy: u16(b, p + 4),
      versionNeeded: u16(b, p + 6),
      flags: u16(b, p + 8),
      method: u16(b, p + 10),
      time: u16(b, p + 12),
      date: u16(b, p + 14),
      crc: u32(b, p + 16),
      csize: u32(b, p + 20),
      usize: u32(b, p + 24),
      internalAttr: u16(b, p + 36),
      externalAttr: u32(b, p + 38),
      localOffset: u32(b, p + 42),
    };
    const nlen = u16(b, p + 28), xlen = u16(b, p + 30), clen = u16(b, p + 32);
    e.nameBytes = b.slice(p + 46, p + 46 + nlen);
    e.name = td.decode(e.nameBytes);
    e.cdExtra = b.slice(p + 46 + nlen, p + 46 + nlen + xlen);
    e.cdComment = b.slice(p + 46 + nlen + xlen, p + 46 + nlen + xlen + clen);
    p += 46 + nlen + xlen + clen;
    // local header
    const lo = e.localOffset;
    const lnlen = u16(b, lo + 26), lxlen = u16(b, lo + 28);
    const dataStart = lo + 30 + lnlen + lxlen;
    e.localHeader = b.slice(lo, dataStart);
    e.raw = b.subarray(dataStart, dataStart + e.csize);
    // a data descriptor may follow when flag bit 3 is set; we keep it too
    e.trailer = new Uint8Array(0);
    if (e.flags & 8) {
      const after = dataStart + e.csize;
      const hasSig = u32(b, after) === 0x08074b50;
      e.trailer = b.slice(after, after + (hasSig ? 16 : 12));
    }
    entries.push(e);
  }
  const map = new Map(entries.map((e) => [e.name, e]));
  return {
    entries,
    has(name) { return map.has(name); },
    bytes(name) {
      const e = map.get(name);
      if (!e) return null;
      if (!e._data) e._data = e.method === 0 ? e.raw.slice() : inflateSync(e.raw);
      return e._data;
    },
    text(name) {
      const d = this.bytes(name);
      return d ? td.decode(d) : null;
    },
  };
}

// replacements: Map(name -> Uint8Array|string|null). null deletes an entry.
// additions: [{name, data}] appended (after existing entries, before nothing).
export function writeZip(zip, replacements = new Map(), additions = []) {
  const parts = [];
  const central = [];
  let offset = 0;
  const push = (arr) => { parts.push(arr); offset += arr.length; };

  const emit = (e, data) => {
    const start = offset;
    if (data === undefined) {
      // untouched: copy local header + raw data verbatim
      push(e.localHeader);
      push(e.raw);
      if (e.trailer.length) push(e.trailer);
      central.push({ e, start, crc: e.crc, csize: e.csize, usize: e.usize, flags: e.flags, method: e.method });
      return;
    }
    const raw = data instanceof Uint8Array ? data : te.encode(data);
    const comp = deflateSync(raw, { level: 6 });
    const crc = crc32(raw);
    const flags = e.flags & ~8; // never write a data descriptor
    const nameBytes = e.nameBytes;
    const h = new Uint8Array(30 + nameBytes.length);
    const dv = new DataView(h.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, e.versionNeeded, true);
    dv.setUint16(6, flags, true);
    dv.setUint16(8, 8, true);
    dv.setUint16(10, e.time, true);
    dv.setUint16(12, e.date, true);
    dv.setUint32(14, crc, true);
    dv.setUint32(18, comp.length, true);
    dv.setUint32(22, raw.length, true);
    dv.setUint16(26, nameBytes.length, true);
    dv.setUint16(28, 0, true);
    h.set(nameBytes, 30);
    push(h);
    push(comp);
    central.push({ e, start, crc, csize: comp.length, usize: raw.length, flags, method: 8, noExtra: true });
  };

  for (const e of zip.entries) {
    if (replacements.has(e.name)) {
      const d = replacements.get(e.name);
      if (d === null) continue;
      emit(e, d);
    } else emit(e, undefined);
  }
  for (const a of additions) {
    const tmpl = zip.entries[0];
    const nameBytes = te.encode(a.name);
    emit({ ...tmpl, name: a.name, nameBytes, cdExtra: new Uint8Array(0), cdComment: new Uint8Array(0), flags: tmpl.flags, versionNeeded: tmpl.versionNeeded, versionMadeBy: tmpl.versionMadeBy, time: tmpl.time, date: tmpl.date, internalAttr: 0, externalAttr: 0 }, a.data);
  }

  const cdStart = offset;
  for (const c of central) {
    const e = c.e;
    const extra = c.noExtra ? new Uint8Array(0) : e.cdExtra;
    const h = new Uint8Array(46 + e.nameBytes.length + extra.length + e.cdComment.length);
    const dv = new DataView(h.buffer);
    dv.setUint32(0, 0x02014b50, true);
    dv.setUint16(4, e.versionMadeBy, true);
    dv.setUint16(6, e.versionNeeded, true);
    dv.setUint16(8, c.flags, true);
    dv.setUint16(10, c.method, true);
    dv.setUint16(12, e.time, true);
    dv.setUint16(14, e.date, true);
    dv.setUint32(16, c.crc, true);
    dv.setUint32(20, c.csize, true);
    dv.setUint32(24, c.usize, true);
    dv.setUint16(28, e.nameBytes.length, true);
    dv.setUint16(30, extra.length, true);
    dv.setUint16(32, e.cdComment.length, true);
    dv.setUint16(34, 0, true);
    dv.setUint16(36, e.internalAttr, true);
    dv.setUint32(38, e.externalAttr, true);
    dv.setUint32(42, c.start, true);
    h.set(e.nameBytes, 46);
    h.set(extra, 46 + e.nameBytes.length);
    h.set(e.cdComment, 46 + e.nameBytes.length + extra.length);
    push(h);
  }
  const cdSize = offset - cdStart;
  const end = new Uint8Array(22);
  const dv = new DataView(end.buffer);
  dv.setUint32(0, 0x06054b50, true);
  dv.setUint16(8, central.length, true);
  dv.setUint16(10, central.length, true);
  dv.setUint32(12, cdSize, true);
  dv.setUint32(16, cdStart, true);
  push(end);

  const out = new Uint8Array(offset);
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
