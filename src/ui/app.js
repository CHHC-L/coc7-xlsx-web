// Application shell: toolbar, formula bar, sheet tabs, persistence, import/export.

import { Doc } from '../core/doc.js';
import { exportXlsx } from '../core/export.js';
import { interpretInput } from '../core/input.js';
import { findDV, validateValue, displayText } from '../core/rules.js';
import { detectVersion, contentEdits, convertCard } from '../core/convert.js';
import { astToString } from '../core/formula.js';
import { addr, key } from '../core/refs.js';
import { XErr } from '../core/values.js';
import { formatValue } from '../core/numfmt.js';
import { Grid } from './grid.js';
import { messageBox, formDialog } from './dialog.js';
import { openStore, unb64 } from './storage.js';

const DEFAULT_ERR = '此值与此单元格定义的数据验证限制不匹配。';

// Page-only wording for the sheet's own hints; the exported xlsx keeps the original text.
const SHOWN_TEXT = new Map([
  ['电脑推荐使用excel2010及以后版本或WPS打开本卡，使用横轴滚动条可以看见更多东西', '使用横轴滚动条或调整缩放以看见更多东西'],
]);

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

const p2 = (n) => String(n).padStart(2, '0');

function fmtTime(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}`;
}

function fmtClock(ts) {
  const d = new Date(ts);
  return `${p2(d.getHours())}:${p2(d.getMinutes())}`;
}

class App {
  // templates: built-in blank cards [{id, name, bytes}], the first being the default
  constructor(root, templates) {
    this.root = root;
    this.templates = templates;
    this.doc = null;
    this.rec = null;
    this.version = null;
    this.undoStack = [];
    this.redoStack = [];
    this.uiZoom = 1;
    this.showHidden = false;
    this.settings = { modifiedBy: '', saveView: true, tpl: templates[0].id, compact: false };
    this.buildShell();
  }

  buildShell() {
    this.root.innerHTML = `
      <header class="bar">
        <div class="tools">
          <button class="btn" data-act="cards" title="角色卡列表">列表</button>
          <button class="btn" data-act="import" title="打开一个 xlsx 文件">打开 xlsx</button>
          <button class="btn primary" data-act="export" title="导出为 xlsx（⌘S）">导出 xlsx</button>
          <span class="sep"></span>
          <button class="btn icon" data-act="undo" title="撤销（⌘Z）">↶</button>
          <button class="btn icon" data-act="redo" title="重做（⇧⌘Z）">↷</button>
          <span class="sep"></span>
          <button class="btn icon" data-act="zoomout" title="缩小">−</button>
          <span class="zoom">100%</span>
          <button class="btn icon" data-act="zoomin" title="放大">+</button>
          <span class="sep"></span>
          <button class="btn" data-act="settings" title="设置">设置</button>
        </div>
        <input type="file" class="filepick" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden>
      </header>
      <div class="fbar">
        <div class="namebox">—</div>
        <div class="fcontent" tabindex="-1"></div>
      </div>
      <main class="gridhost"></main>
      <footer class="tabs-bar">
        <div class="tabs"></div>
        <div class="status"></div>
      </footer>`;
    this.$ = (s) => this.root.querySelector(s);
    this.grid = new Grid(this.$('.gridhost'), this);
    this.root.querySelector('.tools').addEventListener('click', (ev) => {
      const b = ev.target.closest('[data-act]');
      if (!b) return;
      this.action(b.dataset.act);
    });
    this.$('.filepick').addEventListener('change', (ev) => {
      const f = ev.target.files && ev.target.files[0];
      ev.target.value = '';
      if (f) this.importFile(f);
    });
    this.$('.tabs').addEventListener('click', (ev) => {
      const t = ev.target.closest('.tab');
      if (t) this.showSheet(+t.dataset.si);
    });
    this.$('.fcontent').addEventListener('mousedown', (ev) => {
      ev.preventDefault();
      const g = this.grid;
      if (!g.sel || g.editing) { g.focus(); return; }
      g.beginEdit('edit', g.rawText(this.doc.cell(g.si, g.sel.r, g.sel.c)));
    });
    // drop a file anywhere to open it
    this.root.addEventListener('dragover', (ev) => { ev.preventDefault(); });
    this.root.addEventListener('drop', (ev) => {
      ev.preventDefault();
      const f = ev.dataTransfer.files && ev.dataTransfer.files[0];
      if (f) this.importFile(f);
    });
    window.addEventListener('beforeunload', () => { if (this.saveTimer) this.saveNow(); });
  }

  async start() {
    this.store = await openStore();
    const s = await this.store.getMeta('settings');
    if (s) Object.assign(this.settings, s);
    this.applyCompact(this.settings.compact);
    const z = await this.store.getMeta('uiZoom');
    if (z) this.uiZoom = z;
    const last = await this.store.getMeta('last');
    let rec = last ? await this.store.get(last) : null;
    if (!rec) {
      const all = await this.store.list();
      if (all.length) rec = await this.store.get(all[0].id);
    }
    if (rec) {
      try { await this.openRecord(rec); return; } catch (e) { console.error(e); }
    }
    await this.newBlank(this.settings.tpl);
  }

  // built-in template by id; records from before versions existed are CY23.2
  tpl(id) {
    return this.templates.find((t) => t.id === id) || this.templates[0];
  }

  // ---- documents ----------------------------------------------------------------------
  async openRecord(rec) {
    const bytes = rec.base || this.tpl(rec.tpl).bytes;
    const doc = new Doc(bytes);
    this.version = rec.base ? detectVersion(doc.wb) : this.tpl(rec.tpl).id;
    if (!rec.ver && this.version) { rec.ver = this.version; this.store.put(rec); }
    if (rec.unprotected) doc.applyProtectionState(rec.unprotected);
    if (rec.edits && rec.edits.length) doc.applyEditList(rec.edits);
    this.doc = doc;
    this.rec = rec;
    this.undoStack = [];
    this.redoStack = [];
    this.grid.viewState.clear();
    this.grid.resetStyles();
    this.grid.si = -1;
    if (rec.view && rec.view.sel) {
      for (const [si, v] of Object.entries(rec.view.sel)) this.grid.viewState.set(+si, { sel: v, scrollLeft: 0, scrollTop: 0 });
    }
    this.renderTabs();
    const visible = doc.wb.sheets.findIndex((s) => s.state === 'visible');
    let si = rec.view && rec.view.sheet !== undefined ? rec.view.sheet : doc.wb.activeTab;
    if (!doc.wb.sheets[si] || (doc.wb.sheets[si].state !== 'visible' && !this.showHidden)) si = visible;
    this.showSheet(si);
    this.updateTitle();
    await this.store.setMeta('last', rec.id);
  }

  async newBlank(id) {
    const t = this.tpl(id);
    const rec = { id: uid(), tpl: t.id, ver: t.id, baseName: t.name, base: null, edits: [], updated: Date.now(), title: '' };
    await this.store.put(rec);
    await this.openRecord(rec);
    if (this.settings.tpl !== t.id) {
      this.settings.tpl = t.id;
      await this.store.setMeta('settings', this.settings);
    }
    this.toast(`已新建 ${t.id} 空白卡`);
  }

  async importFile(file) {
    let bytes, ver;
    try {
      bytes = new Uint8Array(await file.arrayBuffer());
      ver = detectVersion(new Doc(bytes).wb);
    } catch (e) {
      await messageBox({ title: '无法打开', icon: 'stop', message: `“${file.name}”不是可识别的 xlsx 文件。\n${e.message || e}` });
      return;
    }
    if (this.saveTimer) await this.saveNow();
    const rec = { id: uid(), ver, baseName: file.name, base: bytes, edits: [], updated: Date.now(), title: '' };
    await this.store.put(rec);
    await this.openRecord(rec);
    this.toast(`已打开 ${file.name}`);
  }

  docTitle() {
    const wb = this.doc.wb;
    const si = wb.sheets.findIndex((s) => s.name === '人物卡');
    if (si >= 0) {
      const v = this.doc.eng.getValue(si, 3, 5);
      if (typeof v === 'string' && v.trim()) return v.trim();
    }
    return '';
  }

  updateTitle() {
    document.title = this.docTitle() || '未命名调查员';
  }

  // compact: the toolbar moves to the right end of the formula bar
  applyCompact(on) {
    const tools = this.$('.tools');
    this.root.classList.toggle('compact', !!on);
    if (on) this.$('.fbar').appendChild(tools);
    else this.$('.bar').prepend(tools);
  }

  // what the grid shows for a cell's value
  shownValue(cell) {
    const v = cell.value;
    return typeof v === 'string' && !cell.ast && SHOWN_TEXT.has(v) ? SHOWN_TEXT.get(v) : v;
  }

  scheduleSave() {
    clearTimeout(this.saveTimer);
    this.setStatus('保存中…');
    this.saveTimer = setTimeout(() => this.saveNow(), 400);
  }

  async saveNow() {
    clearTimeout(this.saveTimer);
    this.saveTimer = null;
    if (!this.rec || !this.doc) return;
    this.grid.saveViewState();
    const sel = {};
    for (const [si, v] of this.grid.viewState) if (v.sel) sel[si] = v.sel;
    this.rec.edits = this.doc.editList();
    this.rec.unprotected = this.doc.protectionState();
    this.rec.updated = Date.now();
    this.rec.title = this.docTitle();
    this.rec.view = { sheet: this.grid.si, sel };
    try {
      await this.store.put(this.rec);
      this.setStatus(this.store.kind === 'memory' ? '浏览器不允许保存，关闭前请导出' : `已保存 ${fmtClock(this.rec.updated)}`);
    } catch (e) {
      this.setStatus('保存失败：' + (e.message || e));
    }
  }

  setStatus(s) { this.$('.status').textContent = s; }

  toast(msg) {
    let t = document.querySelector('.toast');
    if (!t) { t = document.createElement('div'); t.className = 'toast'; document.body.appendChild(t); }
    t.textContent = msg;
    t.classList.add('on');
    clearTimeout(this.toastTimer);
    this.toastTimer = setTimeout(() => t.classList.remove('on'), 2200);
  }

  // ---- sheets --------------------------------------------------------------------------
  renderTabs() {
    const wb = this.doc.wb;
    const html = wb.sheets.map((s, i) => {
      if (s.state !== 'visible' && !this.showHidden) return '';
      const color = s.tabColor && s.tabColor.rgb ? `style="--tab:#${s.tabColor.rgb.slice(-6)}"` : '';
      return `<button class="tab${s.state !== 'visible' ? ' hidden-sheet' : ''}" data-si="${i}" ${color}>${esc(s.name)}</button>`;
    }).join('');
    this.$('.tabs').innerHTML = html;
  }

  showSheet(si) {
    this.grid.setSheet(si);
    for (const t of this.root.querySelectorAll('.tab')) t.classList.toggle('on', +t.dataset.si === si);
    const cur = this.root.querySelector('.tab.on');
    if (cur) cur.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    this.updateZoomLabel();
    this.grid.focus();
  }

  async toggleProtection() {
    const si = this.grid.si;
    const doc = this.doc;
    const sh = doc.wb.sheets[si];
    if (!doc.hasProtection(si)) return;
    if (doc.isProtected(si)) {
      if (sh.protection.hasPassword) {
        const pw = await formDialog({
          title: '撤销工作表保护',
          buttons: [{ label: '取消', value: null }, { label: '确定', value: 'collect', primary: true }],
          build: (body) => {
            body.innerHTML = `<label class="fld"><span>密码</span><input type="password" class="i-pw" autocomplete="off"></label><p class="hint">“${esc(sh.name)}”设有密码保护。</p>`;
            return () => body.querySelector('.i-pw').value;
          },
        });
        if (pw === null) return;
        this.setStatus('正在验证密码…');
        const { checkPassword } = await import('../core/protect.js');
        const ok = await checkPassword(sh.protection.attrs, pw);
        if (!ok) {
          await messageBox({ title: 'Microsoft Excel', icon: 'warn', message: '您提供的密码不正确。请验证 CAPS LOCK 键是否关闭，并确保使用正确的大小写。' });
          this.setStatus('');
          return;
        }
      }
      doc.setUnprotected(si, true);
      this.toast(`已撤销“${sh.name}”的工作表保护：锁定的单元格和公式现在也可以修改`);
    } else {
      doc.setUnprotected(si, false);
      this.toast(`已恢复“${sh.name}”原有的工作表保护`);
    }
    this.grid.updateSelection(false);
    this.scheduleSave();
  }

  zoomFor(si) {
    const sh = this.doc.wb.sheets[si];
    return ((sh.view && sh.view.zoomScale) || 100) / 100 * this.uiZoom;
  }

  updateZoomLabel() {
    this.$('.zoom').textContent = Math.round(this.zoomFor(this.grid.si) * 100) + '%';
  }

  setUiZoom(z) {
    this.uiZoom = Math.max(0.4, Math.min(3, Math.round(z * 100) / 100));
    this.grid.setZoom(this.zoomFor(this.grid.si));
    this.updateZoomLabel();
    this.store.setMeta('uiZoom', this.uiZoom);
  }

  // ---- grid callbacks ----------------------------------------------------------------
  formulaText(si, cell) {
    if (cell.formula) return cell.formula;
    return cell.ast ? astToString(cell.ast) : '';
  }

  onSelection(sel) {
    const nb = this.$('.namebox');
    const fc = this.$('.fcontent');
    if (!sel) { nb.textContent = '—'; fc.textContent = ''; return; }
    const g = this.grid.area(sel.r, sel.c);
    nb.textContent = addr(sel.r, sel.c);
    nb.title = g.r2 > g.r1 || g.c2 > g.c1 ? `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}` : '';
    fc.textContent = this.grid.rawText(this.doc.cell(this.grid.si, sel.r, sel.c));
    fc.classList.toggle('locked', !this.doc.canEdit(this.grid.si, sel.r, sel.c));
  }

  onEditing(on, text) {
    const fc = this.$('.fcontent');
    fc.classList.toggle('editing', !!on);
    if (on) fc.textContent = text;
  }

  async protectedAlert() {
    await messageBox({ title: 'Microsoft Excel', icon: 'warn', message: '您试图更改的单元格或图表位于受保护的工作表中。若要进行更改，请取消工作表保护。您可能需要输入密码。' });
    this.grid.focus();
  }

  async validationAlert(dv) {
    const msg = dv.error || DEFAULT_ERR;
    const title = dv.errorTitle || 'Microsoft Excel';
    if (dv.errorStyle === 'warning') {
      const v = await messageBox({ title, icon: 'warn', message: msg + '\n\n是否继续?', buttons: [{ label: '是', value: 'accept' }, { label: '否', value: 'retry', primary: true }, { label: '取消', value: 'cancel', cancel: true }] });
      return v;
    }
    if (dv.errorStyle === 'information') {
      return messageBox({ title, icon: 'info', message: msg, buttons: [{ label: '确定', value: 'accept', primary: true }, { label: '取消', value: 'cancel', cancel: true }] });
    }
    return messageBox({ title, icon: 'stop', message: msg, buttons: [{ label: '重试', value: 'retry', primary: true }, { label: '取消', value: 'cancel', cancel: true }] });
  }

  // returns 'ok' | 'cancel' | 'retry'
  async commitInput(si, r, c, text, opts = {}) {
    const doc = this.doc;
    if (!doc.canEdit(si, r, c)) { await this.protectedAlert(); return 'cancel'; }
    const s = doc.styleOf(si, r, c);
    const content = opts.clear || text === '' ? { kind: 'clear' } : interpretInput(text, doc.numFmtCode(s));
    const before = doc.snapshot(si, r, c);
    // unchanged?
    const cur = doc.cell(si, r, c);
    if (content.kind === 'clear' && (!cur || (cur.value === null && !cur.ast))) return 'ok';
    const dv = findDV(doc, si, r, c);
    // like Excel, pasting bypasses data validation
    if (content.kind === 'value' && dv && dv.showErrorMessage && !opts.fromList && !opts.noValidate) {
      if (!validateValue(doc, si, r, c, dv, content.value)) {
        const choice = await this.validationAlert(dv);
        if (choice !== 'accept') return choice === 'retry' && opts.fromEditor ? 'retry' : 'cancel';
      }
    }
    const res = doc.setContent(si, r, c, content);
    if (!res.ok) {
      doc.restore(before);
      await messageBox({ title: 'Microsoft Excel', icon: 'stop', message: '此公式有问题。\n\n' + res.error });
      return opts.fromEditor ? 'retry' : 'cancel';
    }
    doc.recalc();
    if (content.kind === 'formula' && dv && dv.showErrorMessage && !opts.noValidate) {
      const cell = doc.cell(si, r, c);
      if (!validateValue(doc, si, r, c, dv, cell.value)) {
        const choice = await this.validationAlert(dv);
        if (choice !== 'accept') {
          doc.restore(before);
          doc.recalc();
          return choice === 'retry' && opts.fromEditor ? 'retry' : 'cancel';
        }
      }
    }
    const after = doc.snapshot(si, r, c);
    this.pushUndo([{ before, after }], { si, r, c });
    this.afterChange();
    return 'ok';
  }

  pushUndo(changes, where) {
    this.undoStack.push({ changes, where });
    if (this.undoStack.length > 300) this.undoStack.shift();
    this.redoStack = [];
  }

  afterChange() {
    this.grid.refresh();
    this.updateTitle();
    this.scheduleSave();
  }

  applyTransaction(tr, dir) {
    const doc = this.doc;
    const list = dir < 0 ? [...tr.changes].reverse() : tr.changes;
    for (const ch of list) doc.restore(dir < 0 ? ch.before : ch.after);
    doc.recalc();
    if (tr.where) {
      if (this.grid.si !== tr.where.si) this.showSheet(tr.where.si);
      this.grid.sel = { r: tr.where.r, c: tr.where.c };
    }
    this.afterChange();
  }

  undo() {
    if (this.grid.editing) return;
    const tr = this.undoStack.pop();
    if (!tr) return;
    this.applyTransaction(tr, -1);
    this.redoStack.push(tr);
  }

  redo() {
    if (this.grid.editing) return;
    const tr = this.redoStack.pop();
    if (!tr) return;
    this.applyTransaction(tr, 1);
    this.undoStack.push(tr);
  }

  copyText(si, r, c) {
    const cell = this.doc.cell(si, r, c);
    if (!cell || cell.value === null) return '';
    if (cell.value instanceof XErr) return cell.value.code;
    return formatValue(cell.value, this.doc.numFmtCode(cell.s), this.doc.wb.styles.indexed).text;
  }

  async pasteText(si, r, c, text) {
    text = text.replace(/\r\n?/g, '\n');
    if (text.endsWith('\n')) text = text.slice(0, -1);
    const rows = text.split('\n').map((l) => l.split('\t'));
    if (rows.length === 1 && rows[0].length === 1) {
      await this.commitInput(si, r, c, rows[0][0], { noValidate: true });
      return;
    }
    const doc = this.doc;
    const targets = [];
    for (let i = 0; i < rows.length; i++) {
      for (let j = 0; j < rows[i].length; j++) {
        const rr = r + i, cc = c + j;
        if (!doc.canEdit(si, rr, cc)) { await this.protectedAlert(); return; }
        targets.push([rr, cc, rows[i][j]]);
      }
    }
    const changes = [];
    for (const [rr, cc, t] of targets) {
      const before = doc.snapshot(si, rr, cc);
      const content = t === '' ? { kind: 'clear' } : interpretInput(t, doc.numFmtCode(doc.styleOf(si, rr, cc)));
      const res = doc.setContent(si, rr, cc, content);
      if (!res.ok) { doc.restore(before); continue; }
      changes.push({ before, after: doc.snapshot(si, rr, cc) });
    }
    doc.recalc();
    this.pushUndo(changes, { si, r, c });
    this.afterChange();
  }

  // ---- actions ----------------------------------------------------------------------
  async action(a) {
    if (this.grid.editing) await this.grid.commitEdit();
    switch (a) {
      case 'import': this.$('.filepick').click(); break;
      case 'export': await this.exportDialog(); break;
      case 'undo': this.undo(); break;
      case 'redo': this.redo(); break;
      case 'zoomin': this.setUiZoom(this.uiZoom + 0.1); break;
      case 'zoomout': this.setUiZoom(this.uiZoom - 0.1); break;
      case 'cards': await this.cardsDialog(); break;
      case 'settings': await this.settingsDialog(); break;
      default:
    }
    this.grid.focus();
  }

  currentModifiedBy() {
    const core = this.doc.wb.zip.text('docProps/core.xml') || '';
    const m = /<cp:lastModifiedBy>([^<]*)<\/cp:lastModifiedBy>/.exec(core);
    return m ? m[1] : '';
  }

  async exportDialog() {
    if (this.grid.editing) { const ok = await this.grid.commitEdit(); if (!ok) return; }
    const defName = this.rec.baseName || this.tpl(this.rec.tpl).name;
    const res = await formDialog({
      title: '导出 xlsx',
      buttons: [{ label: '取消', value: null }, { label: '导出', value: 'collect', primary: true }],
      build: (body) => {
        body.innerHTML = `
          <label class="fld"><span>文件名</span><input class="i-name" value="${esc(defName)}"></label>
          <label class="fld"><span>上次修改者</span><input class="i-by" value="${esc(this.settings.modifiedBy || this.currentModifiedBy())}" placeholder="写入文件属性，Excel 的“上次修改者”"></label>
          <label class="chk"><input type="checkbox" class="i-view" ${this.settings.saveView ? 'checked' : ''}> 记录当前工作表和选中单元格（与在 Excel 中保存一致）</label>
          <p class="hint">导出的文件以${this.rec.base ? '你打开的 xlsx' : ` ${esc(this.version)} 空白卡`}为底稿，只改动了你填写和变动的单元格，以及由此重新计算的结果。</p>`;
        return () => ({
          name: body.querySelector('.i-name').value.trim() || defName,
          by: body.querySelector('.i-by').value,
          view: body.querySelector('.i-view').checked,
        });
      },
    });
    if (!res) return;
    this.settings.modifiedBy = res.by;
    this.settings.saveView = res.view;
    await this.store.setMeta('settings', this.settings);
    let name = res.name;
    if (!/\.xlsx$/i.test(name)) name += '.xlsx';
    const g = this.grid;
    let selection = null;
    if (g.sel) {
      const a = g.area(g.sel.r, g.sel.c);
      selection = { si: g.si, r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2, ar: g.sel.r, ac: g.sel.c };
    }
    let out;
    try {
      out = exportXlsx(this.doc, {
        modifiedBy: res.by,
        modified: new Date(),
        saveView: res.view,
        activeSheet: g.si,
        selection,
      });
    } catch (e) {
      console.error(e);
      await messageBox({ title: '导出失败', icon: 'stop', message: String(e.message || e) });
      return;
    }
    const blob = new Blob([out.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    this.toast(`已导出 ${name}（改动 ${out.changedCells.length} 个单元格）`);
  }

  async cardsDialog() {
    await this.saveNow();
    const list = await this.store.list();
    const res = await formDialog({
      title: '角色卡列表',
      buttons: [{ label: '关闭', value: null }],
      build: (body, close) => {
        body.innerHTML = `
          <div class="cards">${list.map((r) => `
            <div class="card${r.id === this.rec.id ? ' cur' : ''}" data-id="${r.id}">
              <div class="ct"><b>${esc(r.title || '（未命名调查员）')}</b><small>${esc(r.ver || (r.hasBase ? '导入的卡' : this.tpl(r.tpl).id))} · ${fmtTime(r.updated)}</small></div>
              <button class="btn sm" data-op="open">打开</button>
              <button class="btn sm danger" data-op="del">删除</button>
            </div>`).join('')}</div>
          <div class="cards-actions"><select class="i-tpl" title="新建空白卡使用的版本">${this.templates.map((t) => `<option value="${esc(t.id)}"${t.id === this.settings.tpl ? ' selected' : ''}>${esc(t.id)}</option>`).join('')}</select><button class="btn" data-op="new">新建空白卡</button><button class="btn" data-op="import">打开 xlsx…</button></div>
          <p class="hint">角色卡保存在本机这个浏览器里（${this.store.kind === 'indexedDB' ? 'IndexedDB' : this.store.kind === 'localStorage' ? 'localStorage' : '仅内存'}）。换浏览器或清除网站数据会丢失，重要的卡请导出 xlsx 备份。</p>`;
        body.addEventListener('click', async (ev) => {
          const b = ev.target.closest('[data-op]');
          if (!b) return;
          const op = b.dataset.op;
          const card = b.closest('.card');
          if (op === 'open') close({ op, id: card.dataset.id });
          else if (op === 'del') close({ op, id: card.dataset.id });
          else if (op === 'new') close({ op, tpl: body.querySelector('.i-tpl').value });
          else close({ op });
        });
        return () => null;
      },
    });
    if (!res) return;
    if (res.op === 'open') {
      const rec = await this.store.get(res.id);
      if (rec) await this.openRecord(rec);
    } else if (res.op === 'del') {
      const r = list.find((x) => x.id === res.id);
      const ok = await messageBox({ title: '删除角色卡', icon: 'warn', message: `确定删除“${r && (r.title || r.baseName)}”？此操作不能撤销。`, buttons: [{ label: '删除', value: true }, { label: '取消', value: false, primary: true, cancel: true }] });
      if (ok) {
        await this.store.del(res.id);
        if (res.id === this.rec.id) {
          const rest = await this.store.list();
          if (rest.length) await this.openRecord(await this.store.get(rest[0].id));
          else await this.newBlank(this.settings.tpl);
        }
      }
    } else if (res.op === 'new') await this.newBlank(res.tpl);
    else if (res.op === 'import') this.$('.filepick').click();
  }

  // the other built-in version this card can be converted to, if any
  convertTarget() {
    if (!this.templates.some((t) => t.id === this.version)) return null;
    return this.templates.find((t) => t.id !== this.version) || null;
  }

  async settingsDialog() {
    const to = this.convertTarget();
    const source = this.rec.base ? `导入自 ${this.rec.baseName}` : '内置空白卡';
    const si = this.grid.si;
    const sheet = this.doc.wb.sheets[si];
    const prot = !this.doc.hasProtection(si) ? { state: '未设保护' }
      : this.doc.isProtected(si) ? { state: '已保护', op: sheet.protection.hasPassword ? '撤销保护…' : '撤销保护' }
        : { state: '已取消保护', op: '恢复保护' };
    const res = await formDialog({
      title: '设置',
      buttons: [{ label: '取消', value: null }, { label: '确定', value: 'collect', primary: true }],
      build: (body, close) => {
        body.innerHTML = `
          <div class="fld"><span>版本</span><div class="ver"><b>${esc(this.version || '未知')}</b><small>${esc(source)}</small>${to ? `<button class="btn sm" data-op="convert">转换为 ${esc(to.id)}…</button>` : ''}</div></div>
          ${to || !this.rec.base ? '' : `<p class="hint">只有 ${this.templates.map((t) => t.id).join(' 和 ')} 之间可以转换版本。</p>`}
          <div class="fld"><span>工作表保护</span><div class="ver"><b>${esc(sheet.name)}</b><small>${prot.state}</small>${prot.op ? `<button class="btn sm" data-op="protect">${prot.op}</button>` : ''}</div></div>
          <label class="fld"><span>上次修改者</span><input class="i-by" value="${esc(this.settings.modifiedBy || this.currentModifiedBy())}"></label>
          <label class="chk"><input type="checkbox" class="i-hidden" ${this.showHidden ? 'checked' : ''}> 显示隐藏的工作表（本职技能、附表，仅查看）</label>
          <label class="chk"><input type="checkbox" class="i-view" ${this.settings.saveView ? 'checked' : ''}> 导出时记录当前工作表和选中单元格</label>
          <label class="chk"><input type="checkbox" class="i-compact" ${this.settings.compact ? 'checked' : ''}> 紧凑模式（工具栏并入编辑栏）</label>`;
        for (const op of ['convert', 'protect']) {
          const b = body.querySelector(`[data-op="${op}"]`);
          if (b) b.addEventListener('click', () => close({ op }));
        }
        return () => ({
          by: body.querySelector('.i-by').value,
          hidden: body.querySelector('.i-hidden').checked,
          view: body.querySelector('.i-view').checked,
          compact: body.querySelector('.i-compact').checked,
        });
      },
    });
    if (!res) return;
    if (res.op === 'convert') { await this.convertVersion(to); return; }
    if (res.op === 'protect') { await this.toggleProtection(); return; }
    this.settings.modifiedBy = res.by;
    this.settings.saveView = res.view;
    if (res.compact !== this.settings.compact) {
      this.settings.compact = res.compact;
      this.applyCompact(res.compact);
    }
    await this.store.setMeta('settings', this.settings);
    if (res.hidden !== this.showHidden) {
      this.showHidden = res.hidden;
      this.renderTabs();
      const si = this.grid.si;
      if (this.doc.wb.sheets[si].state !== 'visible' && !this.showHidden) this.showSheet(this.doc.wb.sheets.findIndex((s) => s.state === 'visible'));
      else this.showSheet(si);
    }
  }

  // Copy this card onto the other version's blank card; the original card is kept.
  async convertVersion(to) {
    const from = this.version;
    const ok = await messageBox({
      title: '转换版本',
      icon: 'info',
      message: `把这张卡从 ${from} 转换为 ${to.id}。\n\n你填写和修改过的格子会原样搬到 ${to.id} 空白卡上，其余内容以 ${to.id} 为准。转换结果另存为一张新卡，原卡保留在“列表”里。`,
      buttons: [{ label: '转换', value: true, primary: true }, { label: '取消', value: false, cancel: true }],
    });
    if (!ok) return;
    await this.saveNow();
    this.setStatus('正在转换…');
    await new Promise((r) => setTimeout(r, 30));
    let result;
    try {
      const srcBlank = new Doc(this.tpl(from).bytes);
      const edits = this.rec.base ? contentEdits(this.doc, srcBlank) : this.doc.editList();
      result = convertCard(edits, this.doc, srcBlank, to.bytes);
      result.doc.applyProtectionState(this.doc.protectionState());
    } catch (e) {
      console.error(e);
      this.setStatus('');
      await messageBox({ title: '转换失败', icon: 'stop', message: String(e.message || e) });
      return;
    }
    const { doc, notes } = result;
    const rec = {
      id: uid(), tpl: to.id, ver: to.id, baseName: to.name, base: null,
      edits: doc.editList(), unprotected: doc.protectionState(), view: this.rec.view,
      title: this.rec.title, updated: Date.now(),
    };
    await this.store.put(rec);
    await this.openRecord(rec);
    this.setStatus(`已保存 ${fmtClock(rec.updated)}`);
    if (!notes.length) { this.toast(`已转换为 ${to.id}，原卡仍在“列表”里`); return; }
    const line = (n) => {
      const at = `${n.sheet}!${n.a}`;
      if (n.kind === 'invalid') return `${at}：“${n.value}”不在 ${to.id} 的可选项里，请重新选择`;
      if (n.kind === 'formula') return `${at}：${to.id} 在这里是公式，已被你填写的内容覆盖`;
      return `${at}：${to.id} 没有这个工作表，这一格没有搬过来`;
    };
    await messageBox({ title: '转换完成', icon: 'warn', message: `已转换为 ${to.id}，原卡仍在“列表”里。以下地方需要你看一下：\n\n${notes.map(line).join('\n')}` });
  }
}

export async function boot() {
  const templates = [...document.querySelectorAll('script.xlsx-template')].map((el) => ({
    id: el.dataset.id, name: el.dataset.name || `${el.dataset.id}.xlsx`, bytes: unb64(el.textContent.trim()),
  }));
  const app = new App(document.getElementById('app'), templates);
  window.cocApp = app;
  try {
    await app.start();
  } catch (e) {
    console.error(e);
    document.getElementById('app').insertAdjacentHTML('afterbegin', `<div class="fatal">启动失败：${esc(e.message || e)}</div>`);
  }
}

export { App, displayText, key };
