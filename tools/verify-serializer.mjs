// Regenerate every <c> element of the template with our serializer and
// compare with Excel's original bytes.
import { readFileSync } from 'node:fs';
import { Doc } from '../src/core/doc.js';
import { exportXlsx } from '../src/core/export.js';
const doc = new Doc(readFileSync(process.argv[2] || 'COC7空白卡CY23.2.xlsx'));
const verifyAll = [];
const res = exportXlsx(doc, { verifyAll });
console.log('parts changed with no edits:', res.parts, 'cells:', res.changedCells.length, res.changedCells.slice(0, 10));
console.log('serializer mismatches:', verifyAll.length);
for (const m of verifyAll.slice(0, 15)) console.log(m[0] + '!' + m[1] + '\n  orig: ' + m[2].slice(0, 200) + '\n  ours: ' + m[3].slice(0, 200));
