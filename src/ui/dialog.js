// Small modal dialogs styled after Excel's message boxes.

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const ICONS = {
  stop: '<svg viewBox="0 0 32 32" width="32" height="32"><circle cx="16" cy="16" r="15" fill="#d13438"/><path d="M11 11l10 10M21 11L11 21" stroke="#fff" stroke-width="3" stroke-linecap="round"/></svg>',
  warn: '<svg viewBox="0 0 32 32" width="32" height="32"><path d="M16 2L31 29H1z" fill="#f2c811"/><path d="M16 11v9" stroke="#000" stroke-width="3" stroke-linecap="round"/><circle cx="16" cy="24.5" r="1.8"/></svg>',
  info: '<svg viewBox="0 0 32 32" width="32" height="32"><circle cx="16" cy="16" r="15" fill="#0f6cbd"/><path d="M16 14v9" stroke="#fff" stroke-width="3" stroke-linecap="round"/><circle cx="16" cy="9.5" r="1.9" fill="#fff"/></svg>',
};

let layer = null;

function ensureLayer() {
  if (!layer) {
    layer = document.createElement('div');
    layer.className = 'dlg-layer';
    document.body.appendChild(layer);
  }
  return layer;
}

// buttons: [{label, value, primary?, cancel?}]
export function messageBox({ title = 'Microsoft Excel', message = '', icon = null, buttons = [{ label: '确定', value: 'ok', primary: true }] }) {
  return new Promise((resolve) => {
    const L = ensureLayer();
    const prevFocus = document.activeElement;
    const box = document.createElement('div');
    box.className = 'dlg';
    box.setAttribute('role', 'alertdialog');
    box.innerHTML = `
      <div class="dlg-title">${esc(title)}</div>
      <div class="dlg-body">${icon ? `<div class="dlg-icon">${ICONS[icon] || ''}</div>` : ''}<div class="dlg-msg">${esc(message).replace(/\n/g, '<br>')}</div></div>
      <div class="dlg-btns">${buttons.map((b, i) => `<button data-i="${i}" class="${b.primary ? 'primary' : ''}">${esc(b.label)}</button>`).join('')}</div>`;
    L.appendChild(box);
    L.classList.add('on');
    const done = (v) => {
      box.remove();
      if (!L.children.length) L.classList.remove('on');
      document.removeEventListener('keydown', onKey, true);
      if (prevFocus && prevFocus.focus) try { prevFocus.focus({ preventScroll: true }); } catch (e) { /* ignore */ }
      resolve(v);
    };
    const onKey = (ev) => {
      if (ev.key === 'Escape') {
        ev.preventDefault(); ev.stopPropagation();
        const c = buttons.find((b) => b.cancel) || buttons[buttons.length - 1];
        done(c.value);
      } else if (ev.key === 'Enter') {
        ev.preventDefault(); ev.stopPropagation();
        const f = document.activeElement && box.contains(document.activeElement) && document.activeElement.dataset.i;
        const b = f !== undefined && f !== null && f !== false ? buttons[+f] : buttons.find((x) => x.primary) || buttons[0];
        done(b.value);
      }
    };
    document.addEventListener('keydown', onKey, true);
    box.addEventListener('click', (ev) => {
      const b = ev.target.closest('button');
      if (b) done(buttons[+b.dataset.i].value);
    });
    const p = box.querySelector('button.primary') || box.querySelector('button');
    if (p) p.focus();
  });
}

// A dialog with arbitrary form content. build(body) fills the body and returns
// a function collecting the result.
export function formDialog({ title, build, buttons }) {
  return new Promise((resolve) => {
    const L = ensureLayer();
    const box = document.createElement('div');
    box.className = 'dlg dlg-form';
    box.innerHTML = `<div class="dlg-title">${esc(title)}</div><div class="dlg-body"></div><div class="dlg-btns">${buttons.map((b, i) => `<button data-i="${i}" class="${b.primary ? 'primary' : ''}">${esc(b.label)}</button>`).join('')}</div>`;
    const collect = build(box.querySelector('.dlg-body'), (v) => done(v));
    L.appendChild(box);
    L.classList.add('on');
    const done = (v) => {
      box.remove();
      if (!L.children.length) L.classList.remove('on');
      document.removeEventListener('keydown', onKey, true);
      resolve(v);
    };
    const onKey = (ev) => {
      if (ev.key === 'Escape') { ev.preventDefault(); ev.stopPropagation(); done(null); }
      if (ev.key === 'Enter' && ev.target.tagName !== 'TEXTAREA' && ev.target.tagName !== 'BUTTON') {
        ev.preventDefault(); ev.stopPropagation();
        const b = buttons.find((x) => x.primary);
        if (b) done(b.value === 'collect' ? collect() : b.value);
      }
    };
    document.addEventListener('keydown', onKey, true);
    box.addEventListener('click', (ev) => {
      const b = ev.target.closest('.dlg-btns button');
      if (!b) return;
      const bt = buttons[+b.dataset.i];
      done(bt.value === 'collect' ? collect() : bt.value);
    });
    const first = box.querySelector('input,textarea,select');
    if (first) first.focus();
  });
}
