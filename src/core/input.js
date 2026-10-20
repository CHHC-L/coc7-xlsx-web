// Interpret text typed into a cell the way Excel does.

import { parseNumericText, errFromCode, normalizeWidth } from './values.js';

const ERR_LIT = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)$/i;

// fmtCode: the cell's current number format code.
// Returns { kind: 'clear' } | { kind: 'formula', formula } |
//         { kind: 'value', value, numFmtId?, quotePrefix? }
export function interpretInput(text, fmtCode) {
  if (text === '' || text === null || text === undefined) return { kind: 'clear' };
  const isTextFmt = fmtCode === '@';
  if (isTextFmt) return { kind: 'value', value: text };
  if (text[0] === "'") return { kind: 'value', value: text.slice(1), quotePrefix: true };
  if (text[0] === '=' && text.length > 1) return { kind: 'formula', formula: text.slice(1) };
  const t = normalizeWidth(text).trim();
  if (/^true$/i.test(t)) return { kind: 'value', value: true };
  if (/^false$/i.test(t)) return { kind: 'value', value: false };
  if (ERR_LIT.test(t)) return { kind: 'value', value: errFromCode(t) };
  const p = parseNumericText(text);
  if (p) {
    const general = !fmtCode || fmtCode === 'General';
    return { kind: 'value', value: p.value, numFmtId: general && p.fmt !== null ? p.fmt : undefined };
  }
  return { kind: 'value', value: text };
}
