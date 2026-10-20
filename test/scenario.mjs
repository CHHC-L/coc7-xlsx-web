// A typical character-creation edit sequence, used by several tests.
import { interpretInput } from '../src/core/input.js';
import { parseAddr } from '../src/core/refs.js';

export const SCENARIO = [
  ['人物卡', 'E3', '测试调查员'],
  ['人物卡', 'E4', '玩家甲'],
  ['人物卡', 'E5', '会计师'],
  ['人物卡', 'M5', '2'],
  ['人物卡', 'E6', '25'],
  ['人物卡', 'M6', '男'],
  ['人物卡', 'E7', '阿卡姆'],
  ['人物卡', 'U3', '50'],
  ['人物卡', 'U5', '60'],
  ['人物卡', 'U7', '55'],
  ['人物卡', 'AA3', '70'],
  ['人物卡', 'AA5', '45'],
  ['人物卡', 'AA7', '80'],
  ['人物卡', 'AG3', '65'],
  ['人物卡', 'AG5', '75'],
  ['人物卡', 'AG7', '50'],
  ['人物卡', 'N16', '40'],
  ['人物卡', 'N17', '20'],
  ['人物卡', 'P18', '10'],
  ['人物卡', 'AJ17', '30'],
  ['人物卡', 'B16', '☑'],
  ['人物卡', 'W77', '第一行背景\n第二行背景'],
  ['人物卡', 'I12', '3'],
  ['人物卡', 'D12', '7'], // overrides a formula cell (重伤值)
];

export function applyScenario(doc, steps = SCENARIO) {
  for (const [sheet, a, text] of steps) {
    const si = doc.wb.sheets.findIndex((s) => s.name === sheet);
    const { r, c } = parseAddr(a);
    if (!doc.canEdit(si, r, c)) throw new Error(`${sheet}!${a} is not editable`);
    const content = interpretInput(text, doc.numFmtCode(doc.styleOf(si, r, c)));
    const res = doc.setContent(si, r, c, content);
    if (!res.ok) throw new Error(res.error);
  }
  doc.recalc();
}
