// Reusable UI component builders (returns DOM nodes)
import { attachCommaFormatter, sanitizeDecimalString, formatTwoDecimalsOnBlur } from './formatting.js?v=20260908d';

let uid = 0;
const nextId = () => `f${++uid}`;

export function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') node.className = v;
    else if (k === 'html') node.innerHTML = v;
    else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
    else if (v === true) node.setAttribute(k, '');
    else if (v !== false && v !== null && v !== undefined) node.setAttribute(k, v);
  }
  for (const child of children.flat()) {
    if (child === null || child === undefined || child === false) continue;
    node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
  }
  return node;
}

// Info-icon — click toggles a popover; closes on outside click or Escape
export function infoIcon(text) {
  const ic = el('span', { class: 'info-icon', tabindex: '0' }, 'i');
  ic.addEventListener('click', (e) => {
    e.stopPropagation();
    closeAllInfoPopovers();
    const pop = el('div', { class: 'info-popover' }, text);
    document.body.appendChild(pop);
    const r = ic.getBoundingClientRect();
    pop.style.left = (window.scrollX + r.right + 6) + 'px';
    pop.style.top = (window.scrollY + r.top - 4) + 'px';
    // If overflow right, flip to left
    const pr = pop.getBoundingClientRect();
    if (pr.right > window.innerWidth - 8) {
      pop.style.left = (window.scrollX + r.left - pr.width - 6) + 'px';
    }
    setTimeout(() => {
      const onDoc = (ev) => {
        if (!pop.contains(ev.target) && ev.target !== ic) closePopover();
      };
      const onKey = (ev) => { if (ev.key === 'Escape') closePopover(); };
      function closePopover() {
        pop.remove();
        document.removeEventListener('mousedown', onDoc);
        document.removeEventListener('keydown', onKey);
      }
      pop._close = closePopover;
      document.addEventListener('mousedown', onDoc);
      document.addEventListener('keydown', onKey);
    }, 0);
  });
  return ic;
}
function closeAllInfoPopovers() {
  document.querySelectorAll('.info-popover').forEach(p => p._close ? p._close() : p.remove());
}

// Number field
export function numberField({ label, name, placeholder = '', help = '', integerOnly = false, min = null, tooltip = '' }) {
  const id = nextId();
  const input = el('input', {
    id, type: 'text', inputmode: integerOnly ? 'numeric' : 'decimal',
    placeholder, autocomplete: 'off', 'data-name': name, class: 'numeric-input',
  });
  attachCommaFormatter(input, { integerOnly });
  if (!integerOnly) input.addEventListener('blur', () => formatTwoDecimalsOnBlur(input));
  const field = el('div', { class: 'field' });
  const lbl = labelWithTooltip(id, label, tooltip);
  field.appendChild(lbl);
  field.appendChild(input);
  if (help) field.appendChild(el('span', { class: 'help' }, help));

  field.getValue = () => {
    const v = input.value.replace(/,/g, '');
    if (v === '') return null;
    const n = Number(v);
    if (isNaN(n)) return null;
    if (min !== null && n < min) return null;
    return n;
  };
  field.setValue = (v) => {
    if (v === null || v === undefined || v === '') input.value = '';
    else if (integerOnly) input.value = Number(v).toLocaleString('en-US', { maximumFractionDigits: 0 });
    else input.value = Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };
  field.input = input;
  field.setLabel = (text) => updateLabel(field, text);
  return field;
}

// Plain free-text field (any characters). Used for the optional "Add Reference" tag.
export function textField({ label, name, placeholder = '', help = '' }) {
  const id = nextId();
  const input = el('input', { id, type: 'text', placeholder, autocomplete: 'off', 'data-name': name, class: 'text-input' });
  const field = el('div', { class: 'field' });
  field.appendChild(el('label', { for: id }, label));
  field.appendChild(input);
  if (help) field.appendChild(el('span', { class: 'help' }, help));
  field.getValue = () => input.value.trim();
  field.setValue = (v) => { input.value = (v === null || v === undefined) ? '' : String(v); };
  field.input = input;
  return field;
}

function labelWithTooltip(id, text, tooltip) {
  const lbl = el('label', { for: id }, text);
  if (tooltip) lbl.appendChild(infoIcon(tooltip));
  return lbl;
}
function updateLabel(field, text) {
  const l = field.querySelector('label');
  if (!l) return;
  l.firstChild.nodeValue = text;
  const ic = l.querySelector('.info-icon');
  if (ic) l.appendChild(ic);
}

// Percent field
export function percentField({ label, name, placeholder = '', help = '', tooltip = '' }) {
  const id = nextId();
  const wrapper = el('div', { class: 'field percent-field' });
  wrapper.appendChild(labelWithTooltip(id, label, tooltip));

  const inputWrap = el('div', { class: 'percent-input-fixed' });
  const input = el('input', {
    id, type: 'text', inputmode: 'decimal',
    placeholder, autocomplete: 'off', 'data-name': name, class: 'numeric-input',
  });
  const suffix = el('span', { class: 'percent-suffix-fixed' }, '%');
  inputWrap.appendChild(input);
  inputWrap.appendChild(suffix);
  wrapper.appendChild(inputWrap);
  if (help) wrapper.appendChild(el('span', { class: 'help' }, help));

  function refreshEmpty() { inputWrap.classList.toggle('empty', input.value === ''); }
  input.addEventListener('input', (e) => { e.target.value = sanitizeDecimalString(e.target.value); refreshEmpty(); });
  input.addEventListener('blur', () => {
    if (input.value === '' || input.value === '.') input.value = '';
    else { const n = Number(input.value); if (!isNaN(n)) input.value = n.toFixed(2); }
    refreshEmpty();
  });
  refreshEmpty();

  wrapper.getValue = () => {
    const v = input.value.trim();
    if (v === '' || v === '.') return null;
    const n = Number(v);
    return isNaN(n) ? null : n / 100;
  };
  wrapper.setValue = (v) => {
    if (v === null || v === undefined || v === '') input.value = '';
    else input.value = (Number(v) * 100).toFixed(2);
    refreshEmpty();
  };
  wrapper.input = input;
  wrapper.setLabel = (text) => updateLabel(wrapper, text);
  return wrapper;
}

// Option field (native select, centered)
export function optionField({ label, name, options, value = null, onChange = null, help = '', tooltip = '' }) {
  const id = nextId();
  const wrapper = el('div', { class: 'field' });
  wrapper.appendChild(labelWithTooltip(id, label, tooltip));
  const select = el('select', { id, 'data-name': name, class: 'centered-input' });
  function fillOptions(opts) {
    select.innerHTML = '';
    opts.forEach((opt) => {
      const text = typeof opt === 'string' ? opt : opt.label;
      const val = typeof opt === 'string' ? opt : opt.value;
      select.appendChild(el('option', { value: val }, text));
    });
  }
  fillOptions(options);
  if (value !== null) select.value = value;
  if (onChange) select.addEventListener('change', () => onChange(select.value));
  wrapper.appendChild(select);
  if (help) wrapper.appendChild(el('span', { class: 'help' }, help));

  wrapper.getValue = () => select.value;
  wrapper.setValue = (v) => { select.value = v; };
  wrapper.setOptions = (opts) => {
    const prev = select.value;
    fillOptions(opts);
    if ([...select.options].some(o => o.value === prev)) select.value = prev;
    else if (select.options.length) select.value = select.options[0].value;
  };
  wrapper.setLabel = (text) => updateLabel(wrapper, text);
  wrapper.select = select;
  return wrapper;
}

// ============================================================
// Windows-11-style date picker (custom, replaces flatpickr). Three drill-down views:
//   days  — click the "Month Year" title → months
//   months — click the "Year" title → years
//   years  — a decade grid
// The up/down arrows page the current view (month / year / decade). A "Today" footer button
// selects today. Value lives on the input as DD-Mmm-YYYY; the popup mounts on <body>.
// ============================================================
const DP_MONTHS_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DP_MONTHS_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const DP_WEEKDAYS = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];
const DP_CHEVRON_UP = '<svg viewBox="0 0 16 16" width="11" height="11"><path d="M3.5 10 8 5.5 12.5 10" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const DP_CHEVRON_DOWN = '<svg viewBox="0 0 16 16" width="11" height="11"><path d="M3.5 6 8 10.5 12.5 6" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';

function attachDatePicker(input, { onChange = null, disable = null } = {}) {
  let pop = null, view = 'days', viewYear, viewMonth, selected = null;
  const startOfDay = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const isDisabled = (d) => typeof disable === 'function' && !!disable(d);
  const sameDay = (a, b) => a && b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  const fireChange = () => { if (onChange) onChange(input.value); };

  function syncFromInput() {
    const d = parseDDMMMYYYY(input.value);
    const ok = d && !isNaN(d);
    selected = ok ? startOfDay(d) : null;
    const base = ok ? d : new Date();
    viewYear = base.getFullYear(); viewMonth = base.getMonth();
  }
  function commit(date) { selected = date; input.value = formatDDMMMYYYY(date); fireChange(); close(); }

  function open() {
    if (pop) return;
    syncFromInput(); view = 'days';
    pop = el('div', { class: 'dp-cal' });
    document.body.appendChild(pop);
    render();
    requestAnimationFrame(() => {
      document.addEventListener('mousedown', onDocDown, true);
      document.addEventListener('keydown', onKey, true);
      window.addEventListener('resize', position);
      window.addEventListener('scroll', position, true);
    });
  }
  function close() {
    if (!pop) return;
    pop.remove(); pop = null;
    document.removeEventListener('mousedown', onDocDown, true);
    document.removeEventListener('keydown', onKey, true);
    window.removeEventListener('resize', position);
    window.removeEventListener('scroll', position, true);
  }
  function onDocDown(e) { if (pop && !pop.contains(e.target) && e.target !== input) close(); }
  function onKey(e) { if (e.key === 'Escape') close(); }

  function position() {
    if (!pop) return;
    const r = input.getBoundingClientRect();
    const ph = pop.offsetHeight || 320;
    const above = r.top - ph - 4, below = r.bottom + 4;
    pop.style.left = (window.scrollX + r.left) + 'px';
    pop.style.top = (window.scrollY + ((r.bottom + ph + 8 > window.innerHeight && above > 4) ? above : below)) + 'px';
  }

  function navStep(dir) {
    if (view === 'days') { viewMonth += dir; if (viewMonth < 0) { viewMonth = 11; viewYear--; } else if (viewMonth > 11) { viewMonth = 0; viewYear++; } }
    else if (view === 'months') viewYear += dir;
    else viewYear += dir * 10;
    render(dir < 0 ? 'up' : 'down'); // slide in the same direction as the clicked arrow
  }
  function headerEl(titleText, onTitleClick) {
    const title = el('button', { class: 'dp-title', type: 'button' }, titleText);
    if (onTitleClick) title.addEventListener('click', onTitleClick); else title.disabled = true;
    const up = el('button', { class: 'dp-nav', type: 'button', 'aria-label': 'Previous', html: DP_CHEVRON_UP });
    const down = el('button', { class: 'dp-nav', type: 'button', 'aria-label': 'Next', html: DP_CHEVRON_DOWN });
    up.addEventListener('click', () => navStep(-1));
    down.addEventListener('click', () => navStep(1));
    return el('div', { class: 'dp-head' }, title, el('div', { class: 'dp-navs' }, up, down));
  }

  // anim: 'up' | 'down' (arrow paging) | 'zoomin' | 'zoomout' (drill in/out) — animates the body
  // (everything below the header) so the change is visible; the header & Today button stay put.
  function render(anim) {
    if (!pop) return;
    pop.innerHTML = '';
    const body = el('div', { class: 'dp-body' + (anim ? ' dp-anim-' + anim : '') });
    if (view === 'days') {
      pop.appendChild(headerEl(`${DP_MONTHS_FULL[viewMonth]} ${viewYear}`, () => { view = 'months'; render('zoomout'); }));
      body.appendChild(el('div', { class: 'dp-weekdays' }, ...DP_WEEKDAYS.map(w => el('div', { class: 'dp-wd' }, w))));
      body.appendChild(buildDaysGrid());
    } else if (view === 'months') {
      pop.appendChild(headerEl(`${viewYear}`, () => { view = 'years'; render('zoomout'); }));
      body.appendChild(buildMonthsGrid());
    } else {
      const decadeStart = Math.floor(viewYear / 10) * 10;
      pop.appendChild(headerEl(`${decadeStart} - ${decadeStart + 9}`, null));
      body.appendChild(buildYearsGrid(decadeStart));
    }
    pop.appendChild(body);
    const tbtn = el('button', { class: 'dp-todaybtn', type: 'button' }, 'Today');
    tbtn.addEventListener('click', () => {
      const t = startOfDay(new Date());
      if (isDisabled(t)) { viewYear = t.getFullYear(); viewMonth = t.getMonth(); view = 'days'; render(); }
      else commit(t);
    });
    pop.appendChild(tbtn);
    position();
  }
  function buildDaysGrid() {
    const grid = el('div', { class: 'dp-grid dp-days' });
    const firstDow = new Date(viewYear, viewMonth, 1).getDay();
    const t = startOfDay(new Date());
    for (let i = 0; i < 42; i++) {
      const d = new Date(viewYear, viewMonth, 1 - firstDow + i);
      const cls = ['dp-cell', 'dp-day'];
      if (d.getMonth() !== viewMonth) cls.push('dp-other');
      if (sameDay(d, selected)) cls.push('dp-selected');
      else if (sameDay(d, t)) cls.push('dp-today');
      const dis = isDisabled(d);
      if (dis) cls.push('dp-disabled');
      const cell = el('button', { class: cls.join(' '), type: 'button' }, String(d.getDate()));
      if (dis) cell.disabled = true; else cell.addEventListener('click', () => commit(startOfDay(d)));
      grid.appendChild(cell);
    }
    return grid;
  }
  function buildMonthsGrid() {
    const grid = el('div', { class: 'dp-grid dp-my' });
    const now = new Date();
    for (let m = 0; m < 12; m++) {
      const cls = ['dp-cell'];
      if (selected && selected.getFullYear() === viewYear && selected.getMonth() === m) cls.push('dp-selected');
      else if (now.getFullYear() === viewYear && now.getMonth() === m) cls.push('dp-today');
      const cell = el('button', { class: cls.join(' '), type: 'button' }, DP_MONTHS_ABBR[m]);
      cell.addEventListener('click', () => { viewMonth = m; view = 'days'; render('zoomin'); });
      grid.appendChild(cell);
    }
    return grid;
  }
  function buildYearsGrid(decadeStart) {
    const grid = el('div', { class: 'dp-grid dp-my' });
    const now = new Date();
    for (let i = -2; i <= 13; i++) {
      const y = decadeStart + i;
      const cls = ['dp-cell'];
      if (y < decadeStart || y > decadeStart + 9) cls.push('dp-other');
      if (selected && selected.getFullYear() === y) cls.push('dp-selected');
      else if (now.getFullYear() === y) cls.push('dp-today');
      const cell = el('button', { class: cls.join(' '), type: 'button' }, String(y));
      cell.addEventListener('click', () => { viewYear = y; view = 'months'; render('zoomin'); });
      grid.appendChild(cell);
    }
    return grid;
  }

  input.addEventListener('focus', open);
  input.addEventListener('mousedown', () => { if (document.activeElement === input) open(); });
  // Typed entry → reformat + notify (mirrors the previous allowInput behaviour).
  input.addEventListener('change', () => {
    const d = parseDDMMMYYYY(input.value);
    if (d && !isNaN(d)) { input.value = formatDDMMMYYYY(d); fireChange(); }
    else if (input.value.trim() === '') fireChange();
  });

  return { open, close };
}

// Date field — custom Windows-11 picker, DD-Mmm-YYYY
export function dateField({ label, name, placeholder = 'dd-Mmm-yyyy', tooltip = '', onChange = null, disableFn = null }) {
  const id = nextId();
  const input = el('input', { id, type: 'text', 'data-name': name, placeholder, autocomplete: 'off', class: 'centered-input date-input' });
  const field = el('div', { class: 'field' });
  field.appendChild(labelWithTooltip(id, label, tooltip));
  field.appendChild(input);

  attachDatePicker(input, { onChange: () => { if (onChange) onChange(input.value); }, disable: disableFn });

  field.getValue = () => {
    if (!input.value) return null;
    const d = parseDDMMMYYYY(input.value);
    return d ? d.toISOString().slice(0, 10) : null;
  };
  field.setValue = (v) => {
    // Programmatic set — write the input directly (the picker reads it on open); no onChange.
    if (!v) { input.value = ''; return; }
    input.value = formatDDMMMYYYY(isoToLocalDate(v));
  };
  field.input = input;
  return field;
}

// Parse an ISO 'YYYY-MM-DD' string to a Date at LOCAL midnight (avoids the UTC
// off-by-one that `new Date('2020-01-01')` causes in negative-offset timezones).
export function isoToLocalDate(iso) {
  if (iso instanceof Date) return iso;
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) { const d = new Date(iso); return isNaN(d) ? null : d; }
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

export function parseDDMMMYYYY(s) {
  if (!s) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) { const d = new Date(s); return isNaN(d) ? null : d; }
  const m = s.match(/^(\d{1,2})[-\s\/](\w{3})[-\s\/](\d{4})$/);
  if (!m) { const d = new Date(s); return isNaN(d) ? null : d; }
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const mi = months.findIndex(x => x.toLowerCase() === m[2].toLowerCase());
  if (mi < 0) return null;
  return new Date(Date.UTC(Number(m[3]), mi, Number(m[1])));
}
export function formatDDMMMYYYY(d) {
  if (!d || isNaN(d)) return '';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  return `${day}-${months[d.getMonth()]}-${d.getFullYear()}`;
}

// Month boxes — click each box to cycle its state: none (accrue) → paid → capitalized.
// Optional bulk-action buttons (set all to one state / clear) + a colour legend.
// Optional extras (all inert when omitted, so the moratorium usage is unchanged):
//   lockedFn(i)   -> true if month i+1 is fixed at Paid and cannot be clicked. Used by the
//                    split interest/principal type, where a principal-payment month always
//                    settles its interest.
//   defaultFn(i)  -> the state month i+1 takes when the RM has not clicked it. Lets the
//                    frequency dropdowns pre-fill the grid while manual overrides survive
//                    later edits to tenor or frequency.
//   groupByYear   -> lay the boxes out in rows of 12 under Year 01, Year 02, ... so a long
//                    tenor stays readable.
//   disabledFn(i) -> true if month i+1 is outside this grid's remit (in Customized, a month
//                    governed by a layer that is not the split type). Rendered inert and
//                    greyed, never clickable, and its state is meaningless.
export function monthBoxesField({ name, getCount, tooltip = '', label = '', selectAll = true, capitalizable = false,
                                  lockedFn = null, defaultFn = null, groupByYear = false, disabledFn = null }) {
  const wrapper = el('div', { class: 'field' });
  let states = []; // per month: 0 = none (accrue), 1 = paid, 2 = capitalized
  const touched = new Set(); // months the RM clicked — these survive a defaults refill
  const maxState = capitalizable ? 2 : 1;
  if (label) {
    const head = el('div', { class: 'label-row' });
    const lbl = el('label', {}, label);
    if (tooltip) lbl.appendChild(infoIcon(tooltip));
    head.appendChild(lbl);
    wrapper.appendChild(head);
  }

  // Bulk-action buttons (right-aligned) — set every box to one state. The button whose state
  // matches ALL the boxes stays sharp; the other two blur. A manual mix of states blurs all
  // three, signalling that none is currently active.
  const setAll = (v) => {
    states = states.map((_, i) => (lockedFn && lockedFn(i)) ? 1 : Math.min(v, maxState));
    // A bulk choice is a deliberate override, so it outranks the frequency defaults.
    states.forEach((_, i) => touched.add(i));
    render();
  };
  const bulkBtns = {};
  if (selectAll && capitalizable) {
    bulkBtns[0] = el('button', { type: 'button', class: 'mb-all mb-all-accrue', onclick: () => setAll(0) }, 'All to be Accrued');
    bulkBtns[1] = el('button', { type: 'button', class: 'mb-all mb-all-paid', onclick: () => setAll(1) }, 'All to be Paid');
    bulkBtns[2] = el('button', { type: 'button', class: 'mb-all mb-all-cap', onclick: () => setAll(2) }, 'All to be Capitalized');
    const bulk = el('div', { class: 'mora-bulk' }, bulkBtns[0], bulkBtns[1], bulkBtns[2]);
    // Only meaningful when the frequency drives a default fill — lets the RM discard their
    // manual overrides and go back to what the two frequency dropdowns imply.
    if (defaultFn) {
      bulk.appendChild(el('button', { type: 'button', class: 'mb-all mb-all-reset',
        onclick: () => { touched.clear(); render(); } }, 'Reset to Frequency'));
    }
    wrapper.appendChild(bulk);
  }

  const grid = el('div', { class: 'month-boxes', 'data-name': name });
  wrapper.appendChild(grid);

  // Highlight the bulk button matching the current uniform state; the other two fade (lower
  // vibrancy). A mix of states (or no boxes) fades all three — none is active.
  function updateBulkActive() {
    if (!bulkBtns[0]) return;
    const uniform = states.length > 0 && states.every(s => s === states[0]) ? states[0] : null;
    [0, 1, 2].forEach(s => { if (bulkBtns[s]) bulkBtns[s].classList.toggle('inactive', uniform !== s); });
  }

  const CLS = { 1: 'paid', 2: 'capitalized' };
  function render() {
    const n = Math.max(0, getCount() || 0);
    states = Array.from({ length: n }, (_, i) => {
      if (disabledFn && disabledFn(i)) return 0;
      if (lockedFn && lockedFn(i)) return 1;                       // principal month: always Paid
      if (defaultFn && !touched.has(i)) return Math.min(defaultFn(i), maxState);
      return Math.min(Number(states[i]) || 0, maxState);
    });
    grid.innerHTML = '';
    // Long tenors are chunked into years so the row stays scannable; the moratorium usage
    // (short, ungrouped) keeps its single flat row.
    const perRow = (groupByYear && n > 12) ? 12 : n;
    grid.classList.toggle('grouped', perRow < n);
    let cursor = null;
    for (let i = 0; i < n; i++) {
      if (perRow < n && i % perRow === 0) {
        const yr = el('div', { class: 'mb-year' },
          el('div', { class: 'mb-year-lbl' }, `Year ${String(i / perRow + 1).padStart(2, '0')}`));
        cursor = el('div', { class: 'month-boxes-row' });
        yr.appendChild(cursor);
        grid.appendChild(yr);
      }
      const off = !!(disabledFn && disabledFn(i));
      const locked = !off && !!(lockedFn && lockedFn(i));
      const cls = off ? '' : (CLS[states[i]] || '');
      const box = el('div', {
        class: 'month-box' + (cls ? ' ' + cls : '') + (locked ? ' locked' : '') + (off ? ' disabled' : ''),
        'data-month': i + 1,
        title: off ? 'Governed by this month’s payment layer, not by this grid'
             : locked ? 'Principal is paid this month, so its interest is always paid too' : '',
      },
        el('div', { class: 'mb-num' }, String(i + 1).padStart(2, '0')),
        el('div', { class: 'mb-lbl' }, 'Month'),
      );
      if (!locked && !off) {
        box.addEventListener('click', () => {
          touched.add(i);
          states[i] = (states[i] + 1) % (maxState + 1);
          render();
        });
      }
      (cursor || grid).appendChild(box);
    }
    updateBulkActive();
  }

  wrapper.refresh = render;
  wrapper.getValue = () => states.slice();
  wrapper.setValue = (arr) => {
    states = (arr || []).map(v => v === true ? 1 : v === false ? 0 : Math.min(Number(v) || 0, maxState));
    // Only months that DIFFER from the default were real overrides — mark just those, so a
    // draft saved while the grid sat at its defaults doesn't freeze it. Marking everything
    // would make a later frequency change silently do nothing to the restored months.
    touched.clear();
    if (defaultFn) {
      states.forEach((s, i) => {
        if (lockedFn && lockedFn(i)) return;
        if (s !== Math.min(defaultFn(i), maxState)) touched.add(i);
      });
    } else {
      states.forEach((_, i) => touched.add(i));
    }
    render();
  };
  wrapper.getPaidFlags = () => states.map(s => s === 1);
  wrapper.getCapFlags = () => states.map(s => s === 2);
  // 'paid' | 'accrued' | 'capitalized' — the shape buildSplitSchedule expects.
  wrapper.getIntFlags = () => states.map(s => s === 2 ? 'capitalized' : s === 1 ? 'paid' : 'accrued');
  render();
  return wrapper;
}

// ============================================================
// Layered field with universal logic
// Options:
//   schema, addLabel, help, onChange
//   minRows: number of initial undeletable rows (default 0)
//   maxRows: hard cap (default Infinity)
//   getMaturity(): returns {value: <max from-key>, kind: 'date'|'month'} — used for last-row To cascading
//   cascadingFromKey, cascadingToKey: keys used for cascading (e.g. 'fromDate'/'toDate' or 'fromInstallment'/'toInstallment')
//   advanceUnit: 'day' | 'month' (for cascading +1)
//   allowFromEqualTo: bool (default true)
//   validate(rows, maturity): return null or error message
// ============================================================
export function layeredField(opts) {
  const {
    label, name, schema, addLabel = '+ Add layer', help = '', initialRows = 1, onChange = null,
    minRows = 0, maxRows = Infinity,
    cascadingFromKey, cascadingToKey,
    getMaturity = null,
    getAnchor = null, // first-row cascadingFromKey source of truth (e.g. disbursement date)
    allowFromEqualTo = true,
  } = opts;
  // Hoisted once — schema is closed-over and never changes
  const fromKind = cascadingFromKey ? (schema.find(s => s.key === cascadingFromKey)?.type) : null;
  const toKind = cascadingToKey ? (schema.find(s => s.key === cascadingToKey)?.type) : null;

  const wrapper = el('div', { class: 'field' });
  wrapper.appendChild(el('label', {}, label));

  const layers = el('div', { class: 'layers', 'data-name': name });
  wrapper.appendChild(layers);
  if (help) wrapper.appendChild(el('span', { class: 'help' }, help));

  // Columns fill the full width (no reserved delete column — the × is overlaid at the
  // row's corner instead, so the boxes stretch edge-to-edge and stay equal width).
  const cols = schema.map(s => s.width || '1fr').join(' ');

  const header = el('div', { class: 'layer-header', style: `grid-template-columns: ${cols}` });
  schema.forEach(s => header.appendChild(el('div', { class: 'layer-th' }, s.label)));
  layers.appendChild(header);

  const rows = [];
  function fireChange() {
    if (onChange) onChange(wrapper.getValue(), rows);
    applyLayerRules();
  }
  function syncHeaderVisibility() {
    header.classList.toggle('hidden', rows.length === 0);
  }

  function addRow(values = {}, opts = {}) {
    const row = el('div', { class: 'layer-row', style: `grid-template-columns: ${cols}` });
    const inputs = {};
    schema.forEach((s) => buildCell(s, row, inputs, values));
    const del = el('button', { type: 'button', class: 'row-del', title: 'Remove' }, '×');
    del.addEventListener('click', () => removeRow(rowApi));
    row.appendChild(del);
    layers.appendChild(row);
    const rowApi = { row, inputs, deletable: !opts.undeletable, errors: new Set() };
    if (!rowApi.deletable) del.style.visibility = 'hidden';
    rows.push(rowApi);
    syncHeaderVisibility();
    return rowApi;
  }

  function buildCell(s, row, inputs, values) {
    let inp, cellNode = null;
    if (s.type === 'option') {
      inp = el('select', { class: 'centered-input' });
      const opts = typeof s.options === 'function' ? s.options() : s.options;
      if (s.allowEmpty) inp.appendChild(el('option', { value: '' }, s.placeholder ?? '— select —'));
      opts.forEach((o) => {
        const val = typeof o === 'string' ? o : o.value;
        const txt = typeof o === 'string' ? o : o.label;
        inp.appendChild(el('option', { value: val }, txt));
      });
      if (values[s.key] !== undefined && values[s.key] !== null) inp.value = values[s.key];
      else if (s.allowEmpty) inp.value = '';
      inp.addEventListener('change', () => { inp.dataset.userSet = '1'; fireChange(); });
      // Auto-derived, user-uneditable field (still set programmatically by applyLayerRules).
      if (s.readOnly) { inp.disabled = true; inp.classList.add('readonly-cell'); }
    } else if (s.type === 'date') {
      inp = el('input', { type: 'text', placeholder: 'dd-Mmm-yyyy', class: 'centered-input date-input' });
      if (values[s.key]) inp.value = values[s.key];
      attachDatePicker(inp, { onChange: () => { inp.dataset.userSet = '1'; fireChange(); } });
    } else if (s.type === 'percent') {
      const pwrap = el('div', { class: 'percent-input-fixed inline' });
      inp = el('input', { type: 'text', inputmode: 'decimal', class: 'numeric-input' });
      const sfx = el('span', { class: 'percent-suffix-fixed' }, '%');
      pwrap.appendChild(inp);
      pwrap.appendChild(sfx);
      if (values[s.key] !== undefined && values[s.key] !== null) inp.value = (values[s.key] * 100).toFixed(2);
      const refreshEmpty = () => pwrap.classList.toggle('empty', inp.value === '');
      inp.addEventListener('input', () => {
        inp.value = sanitizeDecimalString(inp.value);
        refreshEmpty();
        fireChange();
      });
      inp.addEventListener('blur', () => {
        if (inp.value === '' || inp.value === '.') inp.value = '';
        else inp.value = Number(inp.value).toFixed(2);
        refreshEmpty();
        fireChange();
      });
      refreshEmpty();
      cellNode = pwrap;
    } else {
      inp = el('input', { type: 'text', inputmode: 'decimal', placeholder: '0', class: 'numeric-input' });
      if (values[s.key] !== undefined && values[s.key] !== null) {
        inp.value = Number(values[s.key]).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      }
      attachCommaFormatter(inp, { integerOnly: !!s.integerOnly });
      if (!s.integerOnly) inp.addEventListener('blur', () => formatTwoDecimalsOnBlur(inp));
      inp.addEventListener('input', fireChange);
    }
    inputs[s.key] = inp;
    // Wrap each control with its column label so layers can stack on mobile.
    // `.layer-cell { display: contents }` keeps the desktop grid identical.
    row.appendChild(el('div', { class: 'layer-cell', 'data-label': s.label }, cellNode || inp));
  }

  function removeRow(rowApi) {
    if (!rowApi.deletable) return;
    const idx = rows.indexOf(rowApi);
    if (idx < 0) return;
    rows.splice(idx, 1);
    rowApi.row.remove();
    syncHeaderVisibility();
    fireChange();
  }

  const addBtn = el('button', { type: 'button', class: 'add-layer' }, addLabel);
  addBtn.addEventListener('click', () => {
    if (rows.length >= maxRows) {
      if (wrapper.onCannotAdd) wrapper.onCannotAdd('Maximum layers reached.');
      return;
    }
    if (!canAddNew()) {
      if (wrapper.onCannotAdd) wrapper.onCannotAdd('Complete the existing layer(s) and ensure the last "To" is before maturity before adding a new layer.');
      return;
    }
    addRow();
    fireChange();
  });
  wrapper.appendChild(addBtn);

  // Seed initial rows + minRows (undeletable)
  const totalInitial = Math.max(initialRows, minRows);
  for (let i = 0; i < totalInitial; i++) addRow({}, { undeletable: i < minRows });
  // Apply rules once initial rows have mounted (after rAF so flatpickr is attached)
  requestAnimationFrame(() => applyLayerRules());

  // ============ Universal layer rules ============
  function getMaturityValue() {
    if (!getMaturity) return null;
    const m = getMaturity();
    return m && m.value ? m.value : null;
  }
  function readVal(inp, kind) {
    if (!inp) return null;
    const v = inp.value;
    if (!v) return null;
    if (kind === 'date') {
      const d = parseDDMMMYYYY(v);
      return d ? d.toISOString().slice(0, 10) : null;
    }
    return Number(v);
  }
  function setVal(inp, v, kind) {
    if (!inp) return;
    if (kind === 'date') {
      // Programmatic set — write the input directly (the picker reads it on open). Setting the
      // value never fires a user 'change', so userSet stays clear (not mistaken for a user edit).
      if (!v) { inp.value = ''; return; }
      inp.value = formatDDMMMYYYY(isoToLocalDate(v));
    } else {
      inp.value = (v === null || v === undefined) ? '' : String(v);
    }
  }
  function advanceOne(v, kind) {
    if (kind === 'date') {
      const d = new Date(v);
      d.setDate(d.getDate() + 1);
      return d.toISOString().slice(0, 10);
    }
    return Number(v) + 1;
  }

  function rowsComplete() {
    if (!cascadingFromKey) return true;
    if (!cascadingToKey) return rows.every((r) => readVal(r.inputs[cascadingFromKey], fromKind) != null);
    return rows.every((r) =>
      readVal(r.inputs[cascadingFromKey], fromKind) != null &&
      readVal(r.inputs[cascadingToKey], toKind) != null
    );
  }
  function canAddNew() {
    if (rows.length === 0) return true;
    if (!cascadingFromKey) return true;
    if (!rowsComplete()) return false;
    const maturity = getMaturityValue();
    if (!cascadingToKey) {
      // from-only: can add while the last From is before maturity
      const lastFrom = readVal(rows[rows.length - 1].inputs[cascadingFromKey], fromKind);
      if (lastFrom && maturity && lastFrom >= maturity) return false;
      return true;
    }
    const lastTo = readVal(rows[rows.length - 1].inputs[cascadingToKey], toKind);
    if (lastTo && maturity && lastTo >= maturity) return false;
    return true;
  }

  function applyLayerRules() {
    if (!cascadingFromKey) return;
    const maturity = getMaturityValue();
    const anchor = getAnchor ? (getAnchor()?.value || null) : null;
    // From-only mode (no To column): a layer runs from its From until the day/month before the
    // next layer's From (the last layer extends to maturity). Used by Rate Revision layers.
    const fromOnly = !cascadingToKey;

    // Every field stays enabled (editable). Derived fields are AUTO-FILLED unless the
    // user has manually edited that specific cell (tracked via dataset.userSet, which is
    // only set by a genuine flatpickr/user change — never by programmatic setVal).
    rows.forEach((r, i) => {
      const fromInp = r.inputs[cascadingFromKey];
      const toInp = cascadingToKey ? r.inputs[cascadingToKey] : null;
      r.errors.clear();
      [fromInp, toInp].forEach(inp => inp && inp.classList.remove('field-error'));

      // FROM
      if (fromInp && fromInp.dataset.userSet !== '1') {
        if (i === 0) {
          // First row anchored to the external source (e.g. disbursement date / mora+1 month)
          setVal(fromInp, anchor, fromKind);
        } else if (!fromOnly) {
          // To-mode: cascade from the previous row's To + 1 unit (cleared if prev To empty)
          const prevTo = readVal(rows[i - 1].inputs[cascadingToKey], toKind);
          setVal(fromInp, prevTo ? advanceOne(prevTo, fromKind) : null, fromKind);
        }
        // from-only: each subsequent From is user-entered (no auto-fill).
      }

      // TO — only the LAST row auto-defaults to maturity (when not user-edited).
      if (!fromOnly) {
        const isLast = i === rows.length - 1;
        if (isLast && toInp && toInp.dataset.userSet !== '1') setVal(toInp, maturity, toKind);
      }
    });

    // Validate logical consistency
    rows.forEach((r, i) => {
      const fromInp = r.inputs[cascadingFromKey];
      const fv = readVal(fromInp, fromKind);
      if (fromOnly) {
        // Each From strictly increasing; every From (incl. the final layer) must be before maturity.
        if (fv && maturity && fv >= maturity) flagError(fromInp, r);
        if (i > 0) {
          const prevFrom = readVal(rows[i - 1].inputs[cascadingFromKey], fromKind);
          if (prevFrom && fv && fv <= prevFrom) flagError(fromInp, r);
        }
        return;
      }
      const toInp = r.inputs[cascadingToKey];
      const tv = readVal(toInp, toKind);
      const isLast = i === rows.length - 1;

      if (fv && tv && (allowFromEqualTo ? tv < fv : tv <= fv)) {
        flagError(toInp, r); flagError(fromInp, r);
      }
      if (tv && maturity && tv > maturity) flagError(toInp, r);
      if (!isLast && tv && maturity && tv === maturity) flagError(toInp, r);
      if (i > 0) {
        const prevTo = readVal(rows[i - 1].inputs[cascadingToKey], toKind);
        if (prevTo && fv && fv <= prevTo) flagError(fromInp, r);
      }
    });

    function flagError(inp, r) {
      if (!inp) return;
      inp.classList.add('field-error');
      r.errors.add(inp);
    }
  }
  function hasErrors() {
    return rows.some(r => r.errors && r.errors.size > 0);
  }

  // Public API
  wrapper.getValue = () => rows.map(({ inputs }) => {
    const row = {};
    schema.forEach((s) => {
      const v = inputs[s.key].value;
      if (s.type === 'percent') row[s.key] = v === '' ? null : Number(v) / 100;
      else if (s.type === 'number') {
        const raw = v.replace(/,/g, '');
        row[s.key] = raw === '' ? null : Number(raw);
      } else if (s.type === 'date') {
        const d = parseDDMMMYYYY(v);
        row[s.key] = d ? d.toISOString().slice(0, 10) : null;
      } else row[s.key] = v || null;
    });
    return row;
  });
  wrapper.setValue = (arr) => {
    rows.length = 0;
    layers.querySelectorAll('.layer-row').forEach(r => r.remove());
    (arr || [{}]).forEach((v, i) => addRow(v, { undeletable: i < minRows }));
    syncHeaderVisibility();
    applyLayerRules();
    fireChange();
  };
  wrapper.addRow = addRow;
  wrapper.rows = rows;
  wrapper.addBtn = addBtn;
  wrapper.refreshOptions = () => {
    rows.forEach(({ inputs }) => {
      schema.forEach((s) => {
        if (s.type === 'option' && typeof s.options === 'function') {
          const inp = inputs[s.key];
          const prev = inp.value;
          inp.innerHTML = '';
          if (s.allowEmpty) inp.appendChild(el('option', { value: '' }, s.placeholder ?? '— select —'));
          s.options().forEach((o) => {
            const val = typeof o === 'string' ? o : o.value;
            const txt = typeof o === 'string' ? o : o.label;
            inp.appendChild(el('option', { value: val }, txt));
          });
          if ([...inp.options].some(o => o.value === prev)) inp.value = prev;
        }
      });
    });
  };
  wrapper.applyLayerRules = applyLayerRules;
  wrapper.hasErrors = hasErrors;
  wrapper.canAddNew = canAddNew;
  // Allow page code to disable inheritance (e.g., for Customized Payment Layers that have non-cascading payment types)
  return wrapper;
}

// A confirmation overlay that stacks ABOVE the shared modal + date picker (z-index 400).
function confirmOverlay(message, { yesLabel = 'Yes', noLabel = 'No, go back', danger = false, onYes } = {}) {
  const back = el('div', { class: 'sec-confirm-back' });
  const yesBtn = el('button', { type: 'button', class: danger ? 'danger-btn' : 'primary-btn' }, yesLabel);
  const noBtn = el('button', { type: 'button', class: 'ghost-btn modal-ghost' }, noLabel);
  const card = el('div', { class: 'sec-confirm-card' },
    el('p', { class: 'sec-confirm-msg' }, message),
    el('div', { class: 'sec-confirm-actions' }, noBtn, yesBtn),
  );
  back.appendChild(card);
  document.body.appendChild(back);
  const close = () => back.remove();
  noBtn.addEventListener('click', close);
  yesBtn.addEventListener('click', () => { close(); if (onYes) onYes(); });
  back.addEventListener('click', (e) => { if (e.target === back) close(); });
}

// ── Loan Security layers (Rate Revision) ────────────────────────────────────────
// Read-only table — ONE row per From Date — showing Total Outstanding (Σ amounts) and the
// Applicable Interest Rate (amount-weighted average). Editing a row opens a popup that captures
// the date once and one-or-more {Outstanding Amount, Interest Rate} securities (an RM may take
// several). Save & Cancel each ask for a re-confirmation. getValue() returns each row as
// { fromDate, amount, activeRate, securities[] } — amount/activeRate feed the engine unchanged
// (combined per date); securities[] carries the breakdown for the Verify Excel.
export function securityLayersField({ label, name, help = '', getAnchor = null, getMaturity = null }) {
  let data = []; // [{ fromDate: ISO|null, securities: [{ amount:Number, rate:Number(decimal) }] }]
  const money2 = (n) => (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const rowTotal = (r) => (r.securities || []).reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const rowRate = (r) => {
    const tot = rowTotal(r);
    if (tot <= 0) return 0;
    return (r.securities || []).reduce((s, x) => s + (Number(x.amount) || 0) * (Number(x.rate) || 0), 0) / tot;
  };
  const dmy = (iso) => (iso ? formatDDMMMYYYY(isoToLocalDate(iso)) : '—');
  const closeEditor = () => { closeModal(); const mc = document.getElementById('modal-card'); if (mc) mc.classList.remove('modal-card--sec'); };

  const wrapper = el('div', { class: 'field' });
  // Popup commits happen outside the form section, so bubble a 'change' from the (in-section)
  // wrapper to wake the draft autosave — otherwise edits only persist on full page unload.
  const notifyChange = () => wrapper.dispatchEvent(new Event('change', { bubbles: true }));
  wrapper.appendChild(el('label', {}, label));
  const tableWrap = el('div', { class: 'sec-table', 'data-name': name });
  wrapper.appendChild(tableWrap);
  const addBtn = el('button', { type: 'button', class: 'sec-add-btn' }, '+ Add Security Layer');
  addBtn.addEventListener('click', () => openEditor(null));
  wrapper.appendChild(addBtn);
  if (help) wrapper.appendChild(el('span', { class: 'help' }, help));

  function render() {
    data.sort((a, b) => String(a.fromDate || '').localeCompare(String(b.fromDate || '')));
    tableWrap.innerHTML = '';
    if (!data.length) { tableWrap.appendChild(el('div', { class: 'sec-empty help' }, 'No security added yet.')); return; }
    // Header: the column labels sit over the bordered value box; a hidden actions placeholder keeps
    // the columns aligned with the rows below (whose Edit/Delete sit OUTSIDE the bordered box).
    tableWrap.appendChild(el('div', { class: 'sec-row sec-head' },
      el('div', { class: 'sec-vals' }, el('div', {}, 'From Date'), el('div', {}, 'Total Outstanding'), el('div', {}, 'Applicable Interest Rate')),
      el('div', { class: 'sec-acts sec-acts-ph' },
        el('button', { type: 'button', class: 'sec-btn', tabindex: '-1' }, 'Edit'),
        el('button', { type: 'button', class: 'sec-btn sec-del', tabindex: '-1' }, 'Delete'))));
    data.forEach((r, i) => {
      const editBtn = el('button', { type: 'button', class: 'sec-btn' }, 'Edit');
      editBtn.addEventListener('click', () => openEditor(i));
      const delBtn = el('button', { type: 'button', class: 'sec-btn sec-del' }, 'Delete');
      delBtn.addEventListener('click', () => confirmOverlay(`Remove the security row for ${dmy(r.fromDate)}?`,
        { yesLabel: 'Yes, remove', danger: true, onYes: () => { data.splice(i, 1); render(); notifyChange(); } }));
      tableWrap.appendChild(el('div', { class: 'sec-row' },
        el('div', { class: 'sec-vals' },
          el('div', {}, dmy(r.fromDate)),
          el('div', {}, money2(rowTotal(r))),
          el('div', {}, (rowRate(r) * 100).toFixed(2) + '%')),
        el('div', { class: 'sec-acts' }, editBtn, delBtn)));
    });
  }

  function openEditor(index) {
    const isNew = index == null;
    const src = isNew ? { fromDate: null, securities: [] } : data[index];
    const work = { fromDate: src.fromDate, securities: (src.securities || []).map((s) => ({ ...s })) };
    if (isNew && !work.fromDate && data.length === 0 && getAnchor) { const a = getAnchor(); if (a && a.value) work.fromDate = a.value; }
    if (!work.securities.length) work.securities.push({ amount: null, rate: null });

    const dateF = dateField({ label: 'From Date', name: 'secFromDate' });
    if (work.fromDate) dateF.setValue(work.fromDate);

    const entriesWrap = el('div', { class: 'sec-entries' });
    const entryRows = [];
    function addEntryRow(values = {}) {
      const amt = numberField({ label: 'Outstanding Amount', name: 'amount' });
      if (values.amount != null) amt.setValue(values.amount);
      const rate = percentField({ label: 'Interest Rate', name: 'rate' });
      if (values.rate != null) rate.setValue(values.rate);
      const rm = el('button', { type: 'button', class: 'row-del', title: 'Remove this security' }, '×');
      const rowEl = el('div', { class: 'sec-entry' }, amt, rate, rm);
      const api = { amt, rate, rowEl };
      rm.addEventListener('click', () => {
        if (entryRows.length <= 1) { toast('Keep at least one security, or use Cancel.', 'warn'); return; }
        const k = entryRows.indexOf(api); if (k >= 0) { entryRows.splice(k, 1); rowEl.remove(); }
      });
      entryRows.push(api);
      entriesWrap.appendChild(rowEl);
    }
    work.securities.forEach((s) => addEntryRow(s));

    const addEntry = el('button', { type: 'button', class: 'sec-add-btn sec-add-entry' }, '+ Add another security');
    addEntry.addEventListener('click', () => addEntryRow({}));
    const cancelBtn = el('button', { type: 'button', class: 'ghost-btn modal-ghost' }, 'Cancel');
    const saveBtn = el('button', { type: 'button', class: 'primary-btn' }, 'Save');

    cancelBtn.addEventListener('click', () => confirmOverlay('Discard your changes and close this window?',
      { yesLabel: 'Yes, discard', danger: true, onYes: closeEditor }));
    saveBtn.addEventListener('click', () => {
      const fromDate = dateF.getValue();
      const securities = entryRows.map((r) => ({ amount: r.amt.getValue(), rate: r.rate.getValue() }))
        .filter((s) => s.amount != null && s.amount > 0);
      if (!fromDate) return toast('Enter the From Date.', 'error');
      if (!securities.length) return toast('Add at least one security with an outstanding amount.', 'error');
      if (securities.some((s) => s.rate == null)) return toast('Enter an interest rate for every security.', 'error');
      const anchor = getAnchor && getAnchor();
      if (anchor && anchor.value && fromDate < anchor.value) return toast('From Date cannot be before the disbursement date.', 'error');
      const mat = getMaturity && getMaturity();
      if (mat && fromDate >= mat) return toast(`From Date must be earlier than loan maturity (${dmy(mat)}).`, 'error');
      if (data.some((r, k) => k !== index && r.fromDate === fromDate)) return toast('A security row for this date already exists — edit that row instead.', 'error');
      confirmOverlay('Save these security details?', { yesLabel: 'Yes, save', onYes: () => {
        const out = { fromDate, securities };
        if (isNew) data.push(out); else data[index] = out;
        closeEditor(); render(); notifyChange();
      } });
    });

    const card = el('div', { class: 'sec-edit' },
      el('div', { class: 'sec-edit-head' }, el('h3', {}, isNew ? 'Add Loan Security' : 'Edit Loan Security')),
      el('p', { class: 'help' }, 'Choose the date this security applies from, then add one or more securities (amount + rate) for that date.'),
      el('div', { class: 'sec-edit-date' }, dateF),
      el('div', { class: 'sec-entry sec-entry-head' }, el('div', {}, 'Outstanding Amount'), el('div', {}, 'Interest Rate'), el('div', {}, '')),
      entriesWrap, addEntry,
      el('div', { class: 'sec-edit-actions' }, cancelBtn, saveBtn));
    openModal(card);
    const back = document.querySelector('#modal-root .modal-backdrop'); if (back) back.onclick = null; // force Save/Cancel
    const mc = document.getElementById('modal-card'); if (mc) mc.classList.add('modal-card--sec');
  }

  wrapper.getValue = () => data
    .filter((r) => r.fromDate && (r.securities || []).some((s) => Number(s.amount) > 0))
    .map((r) => ({
      fromDate: r.fromDate, amount: rowTotal(r), activeRate: rowRate(r),
      securities: (r.securities || []).filter((s) => Number(s.amount) > 0).map((s) => ({ amount: Number(s.amount), rate: Number(s.rate) })),
    }));
  wrapper.setValue = (arr) => {
    data = (arr || []).map((item) => {
      const d = item.fromDate ? parseDDMMMYYYY(item.fromDate) : null;
      const iso = d ? d.toISOString().slice(0, 10) : null;
      let securities;
      if (Array.isArray(item.securities) && item.securities.length) securities = item.securities.map((s) => ({ amount: Number(s.amount) || 0, rate: Number(s.rate) || 0 }));
      else if (item.amount != null) securities = [{ amount: Number(item.amount) || 0, rate: Number(item.activeRate) || 0 }];
      else securities = [];
      return { fromDate: iso, securities };
    }).filter((r) => r.securities.length);
    render();
  };
  wrapper.applyLayerRules = () => render();
  render();
  return wrapper;
}

// ── Lending Rate layers (Rate Revision — Structured) ────────────────────────────
// Read-only table — ONE row per day — showing From Date | Active Rate. A single "Edit"
// button opens a popup holding EVERY day as a { From Date, Rate } row (add/remove days
// there, one rate per day); Save rewrites the whole table after a re-confirmation.
// getValue() returns each day as { fromDate, activeRate } sorted ascending — the engine
// applies the earliest layer's rate to any date before it, so no disbursement-anchored
// first row is required (the popup still seeds the first day to disbursement for convenience).
export function rateLayersField({ label, name, getAnchor = null, getMaturity = null }) {
  let data = []; // [{ fromDate: ISO, rate: Number(decimal) }]
  const dmy = (iso) => (iso ? formatDDMMMYYYY(isoToLocalDate(iso)) : '—');
  const closeEditor = () => { closeModal(); const mc = document.getElementById('modal-card'); if (mc) mc.classList.remove('modal-card--sec'); };

  const wrapper = el('div', { class: 'field' });
  // Popup commits happen outside the form section, so bubble a 'change' from the (in-section)
  // wrapper to wake the draft autosave — otherwise edits only persist on full page unload.
  const notifyChange = () => wrapper.dispatchEvent(new Event('change', { bubbles: true }));
  const headRow = el('div', { class: 'rl-head' });
  headRow.appendChild(el('label', {}, label));
  const editBtn = el('button', { type: 'button', class: 'sec-btn rl-edit-btn' }, 'Edit');
  editBtn.addEventListener('click', openEditor);
  headRow.appendChild(editBtn);
  wrapper.appendChild(headRow);
  const tableWrap = el('div', { class: 'sec-table rl-table', 'data-name': name });
  wrapper.appendChild(tableWrap);

  function render() {
    data.sort((a, b) => String(a.fromDate || '').localeCompare(String(b.fromDate || '')));
    tableWrap.innerHTML = '';
    if (!data.length) { tableWrap.appendChild(el('div', { class: 'sec-empty help' }, 'No lending rate added yet — click Edit to add.')); return; }
    tableWrap.appendChild(el('div', { class: 'sec-row rl-row sec-head' },
      el('div', { class: 'sec-vals rl-vals' }, el('div', {}, 'From Date'), el('div', {}, 'Active Rate'))));
    data.forEach((r) => {
      tableWrap.appendChild(el('div', { class: 'sec-row rl-row' },
        el('div', { class: 'sec-vals rl-vals' },
          el('div', {}, dmy(r.fromDate)),
          el('div', {}, (Number(r.rate) * 100).toFixed(2) + '%'))));
    });
  }

  function openEditor() {
    const work = data.map((r) => ({ ...r }));
    if (!work.length) {
      let seed = null;
      if (getAnchor) { const a = getAnchor(); if (a && a.value) seed = a.value; }
      work.push({ fromDate: seed, rate: null });
    }
    const entriesWrap = el('div', { class: 'sec-entries rl-entries' });
    const entryRows = [];
    function addEntryRow(values = {}) {
      const dateF = dateField({ label: 'From Date', name: 'rlFromDate' });
      if (values.fromDate) dateF.setValue(values.fromDate);
      const rate = percentField({ label: 'Rate', name: 'rate' });
      if (values.rate != null) rate.setValue(values.rate);
      const rm = el('button', { type: 'button', class: 'row-del', title: 'Remove this day' }, '×');
      const rowEl = el('div', { class: 'sec-entry rl-entry' }, dateF, rate, rm);
      const api = { dateF, rate, rowEl };
      rm.addEventListener('click', () => {
        if (entryRows.length <= 1) { toast('Keep at least one day, or use Cancel.', 'warn'); return; }
        const k = entryRows.indexOf(api); if (k >= 0) { entryRows.splice(k, 1); rowEl.remove(); }
      });
      entryRows.push(api);
      entriesWrap.appendChild(rowEl);
    }
    work.forEach((r) => addEntryRow(r));

    const addEntry = el('button', { type: 'button', class: 'sec-add-btn sec-add-entry' }, '+ Add another day');
    addEntry.addEventListener('click', () => addEntryRow({}));
    const cancelBtn = el('button', { type: 'button', class: 'ghost-btn modal-ghost' }, 'Cancel');
    const saveBtn = el('button', { type: 'button', class: 'primary-btn' }, 'Save');

    cancelBtn.addEventListener('click', () => confirmOverlay('Discard your changes and close this window?',
      { yesLabel: 'Yes, discard', danger: true, onYes: closeEditor }));
    saveBtn.addEventListener('click', () => {
      const days = entryRows.map((r) => ({ fromDate: r.dateF.getValue(), rate: r.rate.getValue() }));
      if (days.some((d) => !d.fromDate)) return toast('Enter the From Date for every day.', 'error');
      if (days.some((d) => d.rate == null)) return toast('Enter a rate for every day.', 'error');
      const anchor = getAnchor && getAnchor();
      const mat = getMaturity && getMaturity();
      for (const d of days) {
        if (anchor && anchor.value && d.fromDate < anchor.value) return toast('A rate day cannot be before the disbursement date.', 'error');
        if (mat && d.fromDate >= mat) return toast(`Each rate day must be earlier than loan maturity (${dmy(mat)}).`, 'error');
      }
      const seen = new Set();
      for (const d of days) { if (seen.has(d.fromDate)) return toast('Two days share the same date — each date can appear only once.', 'error'); seen.add(d.fromDate); }
      confirmOverlay('Save these lending rate layers?', { yesLabel: 'Yes, save', onYes: () => {
        data = days.map((d) => ({ fromDate: d.fromDate, rate: d.rate }));
        closeEditor(); render(); notifyChange();
      } });
    });

    const card = el('div', { class: 'sec-edit rl-editor' },
      el('div', { class: 'sec-edit-head' }, el('h3', {}, 'Lending Rate Layers')),
      el('p', { class: 'help' }, 'Add each date a new lending rate takes effect (one rate per day). The earliest rate also applies from disbursement up to that day.'),
      el('div', { class: 'sec-entry rl-entry sec-entry-head' }, el('div', {}, 'From Date'), el('div', {}, 'Rate'), el('div', {}, '')),
      entriesWrap, addEntry,
      el('div', { class: 'sec-edit-actions' }, cancelBtn, saveBtn));
    openModal(card);
    const back = document.querySelector('#modal-root .modal-backdrop'); if (back) back.onclick = null; // force Save/Cancel
    const mc = document.getElementById('modal-card'); if (mc) mc.classList.add('modal-card--sec');
  }

  wrapper.getValue = () => data
    .filter((r) => r.fromDate && r.rate != null)
    .sort((a, b) => String(a.fromDate).localeCompare(String(b.fromDate)))
    .map((r) => ({ fromDate: r.fromDate, activeRate: Number(r.rate) }));
  wrapper.setValue = (arr) => {
    data = (arr || []).map((item) => {
      const d = item.fromDate ? parseDDMMMYYYY(item.fromDate) : null;
      const iso = d ? d.toISOString().slice(0, 10) : null;
      const rate = item.activeRate != null ? Number(item.activeRate) : (item.rate != null ? Number(item.rate) : null);
      return { fromDate: iso, rate };
    }).filter((r) => r.fromDate && r.rate != null);
    render();
  };
  wrapper.applyLayerRules = () => render();
  render();
  return wrapper;
}

// Toast + modal
export function toast(message, type = 'info', duration) {
  const root = document.getElementById('toast-root');
  const t = el('div', { class: `toast ${type}` });
  const msg = el('span', { class: 'toast-msg' }, message);
  const close = el('button', { class: 'toast-close', type: 'button', 'aria-label': 'Dismiss' }, '×');
  let timer = null;
  const dismiss = () => {
    if (timer) { clearTimeout(timer); timer = null; }
    t.style.transition = 'opacity 0.2s ease';
    t.style.opacity = '0';
    setTimeout(() => t.remove(), 220);
  };
  close.addEventListener('click', dismiss);
  t.appendChild(msg);
  t.appendChild(close);
  root.appendChild(t);
  // Warnings & errors persist until the user dismisses them via the × button; info/success
  // auto-dismiss after a comfortable delay but can still be closed early.
  const persist = type === 'warn' || type === 'warning' || type === 'error';
  if (!persist) timer = setTimeout(dismiss, duration != null ? duration : 5000);
  return t;
}

export function openModal(node) {
  const root = document.getElementById('modal-root');
  const card = document.getElementById('modal-card');
  card.innerHTML = '';
  card.appendChild(node);
  root.classList.remove('hidden');
  root.querySelector('.modal-backdrop').onclick = closeModal;
}
export function closeModal() {
  document.getElementById('modal-root').classList.add('hidden');
}
