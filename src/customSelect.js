/**
 * Themed custom dropdowns (native <select> kept for logic; UI matches calendar popover).
 */

const CHEVRON =
  '<svg width="12" height="12" viewBox="0 0 10 6" aria-hidden="true"><path d="M1 1l4 4 4-4" stroke="currentColor" fill="none" stroke-width="1.5"/></svg>';

const instances = new Map();

function positionPanel(btn, panel) {
  const r = btn.getBoundingClientRect();
  const pad = 8;
  const maxW = 320;
  const w = Math.min(Math.max(r.width, 200), maxW);
  panel.style.width = `${w}px`;
  let left = r.left;
  left = Math.min(left, window.innerWidth - w - pad);
  left = Math.max(pad, left);
  panel.style.left = `${left}px`;
  let top = r.bottom + 6;
  const ph = panel.offsetHeight || 240;
  if (top + ph > window.innerHeight - pad) {
    top = Math.max(pad, r.top - ph - 6);
  }
  panel.style.top = `${top}px`;
}

function closeAllExcept(keepInst) {
  instances.forEach((inst) => {
    if (inst !== keepInst && inst._isOpen) inst.close();
  });
}

/**
 * @param {HTMLSelectElement} selectEl
 */
export function mountCustomSelect(selectEl) {
  if (!selectEl || instances.has(selectEl)) return instances.get(selectEl);

  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'custom-select-trigger filter-select-trigger';
  btn.setAttribute('aria-haspopup', 'listbox');
  btn.setAttribute('aria-expanded', 'false');
  const al = selectEl.getAttribute('aria-label');
  if (al) btn.setAttribute('aria-label', al);

  const label = document.createElement('span');
  label.className = 'custom-select-label';

  const chev = document.createElement('span');
  chev.className = 'custom-select-chevron picker-trigger-chevron';
  chev.innerHTML = CHEVRON;

  btn.appendChild(label);
  btn.appendChild(chev);

  const panel = document.createElement('div');
  panel.className = 'custom-select-panel';
  panel.setAttribute('role', 'listbox');
  panel.hidden = true;
  panel.id = `custom-select-panel-${selectEl.id || 'sel'}`;
  btn.setAttribute('aria-controls', panel.id);

  const wrap = document.createElement('div');
  wrap.className = 'custom-select';
  selectEl.parentNode.insertBefore(wrap, selectEl);
  wrap.appendChild(selectEl);
  wrap.appendChild(btn);

  selectEl.classList.add('custom-select-native');
  selectEl.setAttribute('tabindex', '-1');
  selectEl.setAttribute('aria-hidden', 'true');

  document.body.appendChild(panel);

  const inst = {
    _isOpen: false,
    updateLabel() {
      const sel = selectEl.options[selectEl.selectedIndex];
      label.textContent = sel ? sel.textContent : '—';
      btn.disabled = selectEl.disabled;
      btn.classList.toggle('is-disabled', selectEl.disabled);
      if (selectEl.title) btn.title = selectEl.title;
    },
    rebuildPanel() {
      panel.replaceChildren();
      Array.from(selectEl.options).forEach((opt) => {
        const row = document.createElement('button');
        row.type = 'button';
        row.className = 'custom-select-option';
        if (opt.disabled) {
          row.disabled = true;
          row.classList.add('custom-select-option--disabled');
        }
        if (opt.value === selectEl.value) row.classList.add('is-selected');
        row.setAttribute('role', 'option');
        row.dataset.value = opt.value;
        row.textContent = opt.textContent;
        row.addEventListener('click', (e) => {
          e.stopPropagation();
          if (opt.disabled) return;
          selectEl.value = opt.value;
          selectEl.dispatchEvent(new Event('change', { bubbles: true }));
          inst.updateLabel();
          inst.close();
        });
        panel.appendChild(row);
      });
    },
    open() {
      if (selectEl.disabled) return;
      closeAllExcept(inst);
      inst._isOpen = true;
      panel.hidden = false;
      btn.setAttribute('aria-expanded', 'true');
      inst.rebuildPanel();
      requestAnimationFrame(() => {
        positionPanel(btn, panel);
        window.addEventListener('scroll', onScrollResize, true);
        window.addEventListener('resize', onScrollResize);
      });
    },
    close() {
      if (!inst._isOpen) return;
      inst._isOpen = false;
      panel.hidden = true;
      btn.setAttribute('aria-expanded', 'false');
      window.removeEventListener('scroll', onScrollResize, true);
      window.removeEventListener('resize', onScrollResize);
    }
  };

  function onScrollResize() {
    if (inst._isOpen) positionPanel(btn, panel);
  }

  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    if (selectEl.disabled) return;
    if (inst._isOpen) inst.close();
    else inst.open();
  });

  document.addEventListener(
    'pointerdown',
    (e) => {
      if (!inst._isOpen) return;
      if (btn.contains(e.target) || panel.contains(e.target)) return;
      inst.close();
    },
    true
  );

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && inst._isOpen) inst.close();
  });

  selectEl.addEventListener('change', () => inst.updateLabel());

  inst.updateLabel();
  instances.set(selectEl, inst);
  return inst;
}

export function syncCustomSelect(selectEl) {
  const inst = instances.get(selectEl);
  if (!inst) return;
  inst.updateLabel();
  if (inst._isOpen) inst.rebuildPanel();
}
