import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { Doc } from '../src/core/doc.js';
import { loadWorkbook } from '../src/core/workbook.js';
import { Engine } from '../src/core/calc.js';
import { exportXlsx, excelNum } from '../src/core/export.js';
import { readZip } from '../src/core/zip.js';
import { parseXml } from '../src/core/xml.js';
import { interpretInput } from '../src/core/input.js';
import { formatValue } from '../src/core/numfmt.js';
import { numToText, XErr } from '../src/core/values.js';
import { parseFormula, astToString } from '../src/core/formula.js';
import { addr, key, parseAddr } from '../src/core/refs.js';
import { detectVersion, contentEdits, convertCard } from '../src/core/convert.js';
import { applyScenario } from './scenario.mjs';

const bytes = readFileSync(new URL('../COC7空白卡CY23.2.xlsx', import.meta.url));
const finalBytes = readFileSync(new URL('../COC7空白卡CY23Final.xlsx', import.meta.url));
const TEMPLATES = [['CY23.2', bytes], ['CY23Final', finalBytes]];

function sameCached(a, b) {
  if (a instanceof XErr || b instanceof XErr) return a instanceof XErr && b instanceof XErr && a.code === b.code;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a));
  return a === b;
}

for (const [ver, bytes] of TEMPLATES) {
test(`${ver}: every formula reproduces the value Excel cached in the template`, () => {
  const wb = loadWorkbook(bytes);
  const cached = [];
  wb.sheets.forEach((sh) => { for (const c of sh.cells.values()) if (c.ast && !c.volatile) cached.push([sh.name, c, c.value]); });
  new Engine(wb).recalc();
  const bad = cached.filter(([, c, v]) => !sameCached(c.value, v)).map(([s, c, v]) => `${s}!${addr(c.r, c.c)} excel=${v} ours=${c.value}`);
  assert.equal(cached.length > 2000, true);
  assert.deepEqual(bad, []);
});

test(`${ver}: export without edits is byte-identical to the input file`, () => {
  const doc = new Doc(bytes);
  const out = exportXlsx(doc, { modifiedBy: 'x', modified: new Date() });
  assert.equal(Buffer.compare(Buffer.from(out.bytes), bytes), 0);
});

const exportScenario = () => {
  const doc = new Doc(bytes);
  applyScenario(doc);
  const out = exportXlsx(doc, { modifiedBy: '玩家甲', modified: new Date('2026-10-19T10:00:00Z') });
  return { doc, out };
};

test(`${ver}: exported file re-parses to the same values and is self-consistent`, () => {
  const { doc, out } = exportScenario();
  const wb2 = loadWorkbook(out.bytes);
  const diffs = [];
  wb2.sheets.forEach((sh2, si) => {
    const sh1 = doc.wb.sheets[si];
    for (const c1 of sh1.cells.values()) {
      if (c1.volatile) continue;
      const c2 = sh2.cells.get(key(c1.r, c1.c));
      const v2 = c2 ? c2.value : null;
      if (!sameCached(c1.value, v2) && !(c1.value === null && v2 === null)) diffs.push(`${sh1.name}!${addr(c1.r, c1.c)} ${c1.value} vs ${v2}`);
    }
  });
  assert.deepEqual(diffs.slice(0, 10), []);
  // recalculating the exported file changes nothing
  const before = [];
  wb2.sheets.forEach((sh) => { for (const c of sh.cells.values()) if (c.ast && !c.volatile) before.push([c, c.value]); });
  new Engine(wb2).recalc();
  const changed = before.filter(([c, v]) => !sameCached(c.value, v));
  assert.equal(changed.length, 0);
});

test(`${ver}: shared strings, calcChain and parts of the export are consistent`, () => {
  const { out } = exportScenario();
  const z = readZip(out.bytes);
  const sst = parseXml(z.text('xl/sharedStrings.xml'));
  const unique = sst.all('si').length;
  assert.equal(+sst.attrs.uniqueCount, unique);
  let refs = 0;
  const used = new Set();
  for (const e of z.entries) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(e.name)) continue;
    const root = parseXml(z.text(e.name));
    for (const c of root.find('c')) {
      if (c.attrs.t === 's') {
        const i = +c.child('v').text;
        assert.ok(i >= 0 && i < unique, `${e.name} ${c.attrs.r} index ${i}`);
        used.add(i);
        refs++;
      }
    }
  }
  assert.equal(+sst.attrs.count, refs);
  assert.equal(used.size, unique, 'no unreferenced shared strings');
  // calcChain only points at formula cells
  const wb = loadWorkbook(out.bytes);
  const cc = parseXml(z.text('xl/calcChain.xml'));
  let sid = null;
  for (const c of cc.all('c')) {
    if (c.attrs.i) sid = +c.attrs.i;
    const sh = wb.sheets.find((s) => s.sheetId === sid);
    const a = parseAddr(c.attrs.r);
    const cell = sh.cells.get(key(a.r, a.c));
    assert.ok(cell && cell.ast, `calcChain ${sh.name}!${c.attrs.r}`);
  }
  // all xml parts parse
  for (const e of z.entries) if (/\.(xml|rels|vml)$/.test(e.name)) parseXml(z.text(e.name));
  assert.match(z.text('docProps/core.xml'), /<cp:lastModifiedBy>玩家甲<\/cp:lastModifiedBy>/);
  assert.match(z.text('docProps/core.xml'), /2026-10-19T10:00:00Z/);
});
}

test('typed text is interpreted like Excel', () => {
  assert.deepEqual(interpretInput('50', 'General'), { kind: 'value', value: 50, numFmtId: undefined });
  assert.equal(interpretInput('５０', 'General').value, 50);
  assert.equal(interpretInput('50%', 'General').value, 0.5);
  assert.equal(interpretInput('50%', 'General').numFmtId, 9);
  assert.deepEqual(interpretInput("'007", 'General'), { kind: 'value', value: '007', quotePrefix: true });
  assert.deepEqual(interpretInput('=1+2', 'General'), { kind: 'formula', formula: '1+2' });
  assert.equal(interpretInput('true', 'General').value, true);
  assert.equal(interpretInput('123', '@').value, '123');
  assert.equal(interpretInput('张三', 'General').value, '张三');
  assert.equal(interpretInput('1,234', 'General').value, 1234);
});

test('number formats used by the card render like Excel', () => {
  const f = (v, code) => formatValue(v, code).text;
  assert.equal(f(0, '[=0]"";General'), '');
  assert.equal(f(5, '[=0]"";General'), '5');
  assert.equal(f(12, '[=0]"";"/"0'), '/12');
  assert.equal(f(0, '[=0]"";"/"0'), '');
  assert.equal(f(1, '[=1]"☆ 二选一";[=0]""'), '☆ 二选一');
  assert.equal(f(0, '[=1]"☆ 二选一";[=0]""'), '');
  assert.equal(f(3, '\\+0;\\-0;"±"0'), '+3');
  assert.equal(f(-3, '\\+0;\\-0;"±"0'), '-3');
  assert.equal(f(0, '\\+0;\\-0;"±"0'), '±0');
  assert.equal(f(2023, '0"年"'), '2023年');
  assert.equal(f(30, '\\$0'), '$30');
  assert.equal(f(-5, '0_);[Red]\\(0\\)').trim(), '(5)');
  assert.equal(formatValue(-5, '0_);[Red]\\(0\\)').color, 'FF0000');
  assert.equal(f(1234.5, '\\$#,##0_);[Red]\\(\\$#,##0\\)').trim(), '$1,235');
  assert.equal(f(0.5, '0.00'), '0.50');
  assert.equal(f(1 / 3, 'General'), '0.333333333');
  assert.equal(f(32767, '[=32767]"";General'), '');
  assert.equal(f('abc', '0'), 'abc');
});

test('number <-> text conversions follow Excel', () => {
  assert.equal(numToText(5e-9), '0.000000005');
  assert.equal(numToText(1 / 3), '0.333333333333333');
  assert.equal(numToText(0.1 + 0.2), '0.3');
  assert.equal(numToText(1e15), '1E+15');
  assert.equal(excelNum(7.03), '7.03');
  assert.equal(excelNum(2023), '2023');
});

test('formula parser round-trips the formulas of the workbook', () => {
  const wb = loadWorkbook(bytes);
  let n = 0;
  for (const sh of wb.sheets) {
    for (const c of sh.cells.values()) {
      if (!c.formula) continue;
      const a = parseFormula(c.formula);
      const b = parseFormula(astToString(a));
      assert.equal(astToString(b), astToString(a));
      n++;
    }
  }
  assert.ok(n > 1000);
});

test('removing sheet protection drops <sheetProtection> on export', () => {
  const doc = new Doc(bytes);
  doc.setUnprotected(0, true);
  assert.equal(doc.canEdit(0, 3, 21), true); // W3 is a locked formula cell
  const out = exportXlsx(doc, {});
  const z = readZip(out.bytes);
  assert.doesNotMatch(z.text('xl/worksheets/sheet1.xml'), /<sheetProtection/);
  assert.match(z.text('xl/worksheets/sheet2.xml'), /<sheetProtection/);
});

test('overwriting the master of a shared formula promotes the next cell', () => {
  const doc = new Doc(bytes);
  const si = doc.wb.sheets.findIndex((s) => s.name === '资产及物价参考');
  const { r, c } = parseAddr('K93');
  assert.ok(doc.canEdit(si, r, c));
  doc.setContent(si, r, c, { kind: 'value', value: 1 });
  doc.recalc();
  const out = exportXlsx(doc, {});
  const wb2 = loadWorkbook(out.bytes);
  const sh = wb2.sheets[si];
  const k94 = sh.cells.get(key(94, c));
  assert.equal(k94.fattrs.t, 'shared');
  assert.ok(k94.formula, 'K94 is the new master with full formula text');
  const e2 = new Engine(wb2);
  const before = [...sh.cells.values()].filter((x) => x.ast).map((x) => [x, x.value]);
  e2.recalc();
  assert.equal(before.filter(([x, v]) => !sameCached(x.value, v)).length, 0);
});

test('export records the active sheet and selection like an Excel save', () => {
  const doc = new Doc(bytes);
  doc.setContent(0, 3, 5, { kind: 'value', value: '某人' });
  doc.recalc();
  const si = doc.wb.sheets.findIndex((s) => s.name === '职业列表');
  const out = exportXlsx(doc, { saveView: true, activeSheet: si, selection: { si, r1: 3, c1: 3, r2: 3, c2: 3, ar: 3, ac: 3 } });
  const z = readZip(out.bytes);
  const wbx = z.text('xl/workbook.xml');
  assert.match(wbx, new RegExp(`activeTab="${si}" xr2:uid=`));
  const s1 = z.text('xl/worksheets/sheet1.xml');
  const s3 = z.text('xl/worksheets/sheet3.xml');
  assert.doesNotMatch(s1, /tabSelected=/);
  assert.match(s3, /<sheetView showGridLines="0" showRowColHeaders="0" tabSelected="1" topLeftCell="B1"/);
  assert.match(s3, /<selection pane="bottomLeft" activeCell="C3" sqref="C3"\/>/);
});

test('template versions are read from 更新说明', () => {
  assert.equal(detectVersion(loadWorkbook(bytes)), 'CY23.2');
  assert.equal(detectVersion(loadWorkbook(finalBytes)), 'CY23Final');
});

const inputsOf = (doc, edits) => edits.map((e) => {
  const si = doc.wb.sheets.findIndex((s) => s.name === e.sheet);
  const { r, c } = parseAddr(e.a);
  const v = doc.cell(si, r, c)?.value ?? null;
  return `${e.sheet}!${e.a}=${v instanceof XErr ? v.code : v}`;
});

test('converting CY23.2 -> CY23Final -> CY23.2 keeps every input and reports renamed weapons', () => {
  const src = new Doc(bytes);
  applyScenario(src);
  const g54 = parseAddr('G54');
  src.setContent(0, g54.r, g54.c, { kind: 'value', value: '燧发枪' }); // not in CY23Final's weapon list
  src.recalc();
  const edits = src.editList();
  const { doc: fin, notes } = convertCard(edits, src, new Doc(bytes), finalBytes);
  assert.deepEqual(inputsOf(fin, edits), inputsOf(src, edits));
  assert.deepEqual(notes, [{ sheet: '人物卡', a: 'G54', kind: 'invalid', value: '燧发枪' }]);
  for (const a of ['B13', 'J50', 'B50', 'G10', 'R26']) {
    const { r, c } = parseAddr(a);
    assert.deepEqual(fin.eng.getValue(0, r, c), src.eng.getValue(0, r, c), a);
  }
  const { doc: back, notes: notes2 } = convertCard(fin.editList(), fin, new Doc(finalBytes), bytes);
  assert.deepEqual(notes2, []);
  assert.deepEqual(inputsOf(back, edits), inputsOf(src, edits));
});

test('an imported card yields the same edits as the card it was exported from', () => {
  const src = new Doc(bytes);
  applyScenario(src);
  const out = exportXlsx(src, {});
  const imported = new Doc(out.bytes);
  const keys = (l) => l.map((e) => `${e.sheet}!${e.a}`).sort();
  const edits = contentEdits(imported, new Doc(bytes));
  assert.deepEqual(keys(edits), keys(src.editList()));
  const { doc: fin } = convertCard(edits, imported, new Doc(bytes), finalBytes);
  assert.deepEqual(inputsOf(fin, edits), inputsOf(src, edits));
});
