// App controller: tabs, theme, compare view
import { el, openModal, closeModal, toast, optionField } from './components.js?v=20260916a';
import {
  renderRegularLoan, renderCustomizedLoan,
  renderRateRevisionStructured, renderRateRevisionCustomized,
} from './pages.js?v=20260916a';
import { listSummaries, deleteSummary } from './storage.js?v=20260916a';
import { formatPercent, formatMoney, formatNumber } from './formatting.js?v=20260916a';
import { openManual } from './manual.js?v=20260916a';

const root = document.getElementById('app-root');
const compareBtn = document.getElementById('compare-btn');
const tabNav = document.getElementById('tab-nav');

const TABS = {
  regular: { label: 'Loan Facilities — Structured', render: renderRegularLoan, group: 'loan' },
  customized: { label: 'Loan Facilities — Customized', render: renderCustomizedLoan, group: 'loan' },
  revisionStructured: { label: 'Rate Revision — Structured', render: renderRateRevisionStructured, group: 'revision' },
  revisionCustomized: { label: 'Rate Revision — Customized', render: renderRateRevisionCustomized, group: 'revision' },
};

// ---------------- Theme toggle ----------------
const THEME_KEY = 'err_theme';
function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  document.getElementById('theme-icon').innerHTML = theme === 'dark' ? '&#9789;' : '&#9788;';
  localStorage.setItem(THEME_KEY, theme);
}
function initTheme() {
  const stored = localStorage.getItem(THEME_KEY);
  const prefersDark = window.matchMedia?.('(prefers-color-scheme: dark)').matches;
  applyTheme(stored || (prefersDark ? 'dark' : 'light'));
}
document.getElementById('theme-toggle').addEventListener('click', () => {
  const current = document.documentElement.getAttribute('data-theme');
  applyTheme(current === 'dark' ? 'light' : 'dark');
});

// ---------------- Brand back to first tab ----------------
document.getElementById('brand-home').addEventListener('click', () => navigate('regular'));

// ---------------- Info button → per-module user guide ----------------
document.getElementById('info-btn').addEventListener('click', () => openManual(TABS[currentTab] ? currentTab : 'regular'));

// ---------------- Compare button visibility ----------------
function refreshCompareVisibility() {
  const count = listSummaries().length;
  document.getElementById('saved-count').textContent = count;
  // Compare feature hidden for now — swap the two lines below to bring it back.
  document.getElementById('compare-bar').classList.add('hidden');
  // document.getElementById('compare-bar').classList.toggle('hidden', count === 0);
}
window.addEventListener('summary-saved', refreshCompareVisibility);
window.addEventListener('summary-deleted', refreshCompareVisibility);
compareBtn.addEventListener('click', () => navigate('compare'));

// ---------------- Router ----------------
let currentTab = null;
function navigate(tab) {
  if (!TABS[tab] && tab !== 'compare') tab = 'regular';
  currentTab = tab;
  tabNav.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tab);
  });
  window.scrollTo({ top: 0, behavior: 'instant' });
  if (tab === 'compare') {
    renderCompare();
  } else {
    TABS[tab].render(root);
  }
}
tabNav.querySelectorAll('.tab-btn').forEach((btn) => {
  btn.addEventListener('click', () => navigate(btn.dataset.tab));
});

// ---------------- Compare view ----------------
const PRETTY_KEY = {
  offeredRate: 'Offered Rate',
  loanAmount: 'Loan Amount',
  initialAmount: 'Initial Loan Amount',
  moratoriumAvail: 'Moratorium',
  moratoriumPeriod: 'Moratorium Period (Months)',
  loanTenor: 'Loan Tenor (Months)',
  tenorMonths: 'Loan Tenor (Months)',
  paymentMode: 'Payment Mode',
  paymentModality: 'Payment Modality',
  totalCof: 'Total COF',
  fundedSecurityType: 'Security Type',
  csAmount: 'Security Amount',
  csRate: 'Offered Rate on Security Amount',
  numInst: 'Number of Installments',
  disbursementDate: 'Disbursement Date',
  nimComparison: 'NIM Comparison',
  paymentLayers: 'Payment Layers',
  rateLayers: 'Lending Rate Layers',
  securityLayers: 'Security Layers',
  cofLayers: 'COF Layers',
  cofRecords: 'COF Records',
  uploadedRowsCount: 'Uploaded rows',
};
const PERCENT_KEYS = new Set(['offeredRate', 'totalCof', 'csRate']);
const MONEY_KEYS = new Set(['loanAmount', 'initialAmount', 'csAmount']);
const HIDDEN_INPUT_KEYS = new Set(['idpFlags']);

function renderCompare() {
  root.innerHTML = '';
  root.appendChild(el('div', { class: 'page-title' },
    el('h1', {}, 'Compare Saved Summaries'),
    (() => {
      const b = el('button', { class: 'back-link', type: 'button' }, '← Back');
      b.addEventListener('click', () => navigate('regular'));
      return b;
    })(),
  ));

  const summaries = listSummaries();
  if (!summaries.length) {
    root.appendChild(el('div', { class: 'section-card' },
      el('p', {}, 'No summaries saved yet. Run a calculation to auto-save and start comparing.')));
    return;
  }

  // Group filter — Loan Facilities vs Rate Revision
  const card = el('div', { class: 'section-card' });
  const head = el('div', { style: 'display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap' });
  head.appendChild(el('h2', {}, `${summaries.length} of 5 summaries saved`));

  const groupNames = { loan: 'Loan Facilities', revision: 'Rate Revision' };
  const filterField = optionField({
    label: 'Showing', name: 'compareGroup',
    options: ['Loan Facilities', 'Rate Revision'],
    value: 'Loan Facilities',
    onChange: render,
  });
  head.appendChild(filterField);
  card.appendChild(head);

  const gridWrap = el('div', { class: 'compare-grid-wrap' });
  card.appendChild(gridWrap);
  root.appendChild(card);

  function render() {
    gridWrap.innerHTML = '';
    const selectedGroup = filterField.getValue() === 'Loan Facilities' ? 'loan' : 'revision';
    const filtered = listSummaries().filter(s => TABS[s.pageType]?.group === selectedGroup);
    if (!filtered.length) {
      gridWrap.appendChild(el('p', { class: 'help' }, 'No saved summaries in this category yet.'));
      return;
    }
    const grid = el('div', { class: 'compare-grid' });
    filtered.forEach((s) => grid.appendChild(buildCompareCol(s, render)));
    for (let i = filtered.length; i < 5; i++) {
      grid.appendChild(el('div', { class: 'compare-col empty' }, el('div', {}, 'Empty slot')));
    }
    gridWrap.appendChild(grid);
  }
  render();
}

function buildCompareCol(s, refresh) {
  const col = el('div', { class: 'compare-col' });
  const head = el('h4', {}, TABS[s.pageType]?.label || s.pageTitle);
  const del = el('button', { class: 'danger-btn', type: 'button' }, 'Delete');
  del.addEventListener('click', () => {
    deleteSummary(s.id);
    window.dispatchEvent(new CustomEvent('summary-deleted'));
    refresh();
  });
  head.appendChild(del);
  col.appendChild(head);
  col.appendChild(el('div', { class: 'help' }, new Date(s.savedAt).toLocaleString()));

  const kv = (k, v) => el('div', { class: 'kv' },
    el('span', { class: 'k' }, k),
    el('span', { class: 'v' }, v),
  );
  col.appendChild(kv('Effective Rate', formatPercent(s.metrics.effectiveRate)));
  if (s.metrics.nim !== undefined) col.appendChild(kv('NIM', formatPercent(s.metrics.nim)));
  if (s.metrics.nii !== undefined) col.appendChild(kv('Net Int. Income', formatMoney(s.metrics.nii)));

  const inputsBlock = el('details', {}, el('summary', {}, 'Inputs'));
  Object.entries(s.inputs).forEach(([k, v]) => {
    if (HIDDEN_INPUT_KEYS.has(k)) return;
    if (Array.isArray(v) && v.length === 0) return;
    if (v === null || v === '' || v === undefined) return;
    let display;
    if (PERCENT_KEYS.has(k)) display = formatPercent(v);
    else if (MONEY_KEYS.has(k)) display = formatMoney(v);
    else if (Array.isArray(v)) display = `${v.length} layer${v.length !== 1 ? 's' : ''}`;
    else display = String(v);
    inputsBlock.appendChild(kv(PRETTY_KEY[k] || k, display));
  });
  col.appendChild(inputsBlock);
  return col;
}

// ---------------- Boot ----------------
initTheme();
refreshCompareVisibility();
navigate('regular');
