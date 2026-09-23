// Page builders for the four calculation flows
import {
  el, numberField, percentField, optionField, dateField, textField,
  monthBoxesField, layeredField, securityLayersField, rateLayersField, monthRateLayersField, toast, parseDDMMMYYYY, formatDDMMMYYYY,
  openModal, closeModal,
} from './components.js?v=20260923f';
import { isoToDDMMMYYYY } from './formatting.js?v=20260923f';
import {
  buildStructuredSchedule, buildCustomizedSchedule,
  buildRateRevisionStructured, computeMetrics,
  buildSplitSchedule, principalPaymentMonths, SPLIT_MODE, FREQ, FREQ_NAMES,
  computeRevisionMetrics, computeRevisionCustomizedMetrics, buildCofData,
  addMonthsDue, COMMERCIAL_RATE, REFINANCE_RATE, REFINANCE_LENDING_RATE,
} from './calculations.js?v=20260923f';
import { formatMoney, formatPercent, formatNumber, formatRateLayers } from './formatting.js?v=20260923f';
import { saveSummary, listSummaries, getMax, saveDraft, loadDraft, clearDraft } from './storage.js?v=20260923f';
import {
  downloadScheduleAsExcel, downloadSampleAmortization, readUploadedSchedule,
  downloadScheduleAsWord, downloadScheduleAsPDF, downloadVerificationExcel, downloadReportPDF,
  downloadCofSample, readUploadedCof,
  downloadCustomizedRevisionSample, readCustomizedRevisionFile,
} from './excel.js?v=20260923f';

// Cached page state by tab key (also persisted via storage saveDraft)
const tabState = {};

// Render a field's label across two lines: a primary phrase plus a secondary part.
// The secondary part stays inline on desktop (reads as one line, unchanged look) and
// drops onto its own line on mobile (CSS .lbl-line2). Forcing both paired fields to a
// matching two-line height keeps their input boxes aligned on the same row.
// Plain-English readout of when principal is actually paid. The "start from" month is where
// the schedule STARTS COUNTING, so with a quarterly frequency a 1 there means the first
// payment lands at the end of month 3 — that inference is exactly what confuses people, so
// the resulting months are spelled out instead of left to be worked out.
function principalScheduleHint(months, periodMonths) {
  if (!months || !months.length) return '';
  const n = months.length;
  const pad = (m) => String(m).padStart(2, '0');
  const plural = `${n} payment${n === 1 ? '' : 's'}`;
  if (n === 1) return `Principal is paid once, at the end of Month ${pad(months[0])}.`;
  if (periodMonths === 1) return `Principal is paid every month, Month ${pad(months[0])} to Month ${pad(months[n - 1])} — ${plural}.`;
  const shown = n <= 8 ? months.map(pad).join(', ')
    : months.slice(0, 6).map(pad).join(', ') + ', … , ' + pad(months[n - 1]);
  const every = `every ${periodMonths} months`;
  return `First principal payment at the end of Month ${pad(months[0])}, then ${every} — ${shown} (${plural}).`;
}

// Installment-sized funded security. Its stored value is always "<kind> after Moratorium" —
// validation, computeMetrics and the exports all match on that string — but on a loan with no
// moratorium that label names a period that does not exist, so only the display text changes.
function installmentSecurityOption(kind, hasMoratorium) {
  return { value: `${kind} after Moratorium`, label: hasMoratorium ? `${kind} after Moratorium` : `${kind} Installment` };
}

// What the chosen security is CALLED on screen. The stored value never changes (saved
// calculations, validation, metrics and the exports all key on it) — this is display only.
export function securityDisplayName(type, hasMoratorium) {
  if (type === 'EMI after Moratorium' || type === 'EQI after Moratorium') {
    return installmentSecurityOption(type.slice(0, 3), hasMoratorium).label;
  }
  return type || 'Funded Security';
}

// Once a security type is picked, its two fields are named after THAT type — "FDR Amount",
// not a slash-list of every option the user already chose between.
function labelSecurityFields(type, hasMoratorium, csAmount, csRate) {
  const name = securityDisplayName(type, hasMoratorium);
  if (csAmount) csAmount.setLabel(`${name} Amount`);
  if (csRate) csRate.setLabel(`${name} Rate`);
}

// ---- Interest Rate Layers (both Loan Facilities forms, when step 1 said the rate changes) ----
// A read-only table with one Edit popup, like Rate Revision's Lending Rate Layers. The popup
// enforces the layer rules on Save; a Refinance layer's rate is the scheme's fixed rate.
function interestRateLayersField(getTenor) {
  return monthRateLayersField({
    label: 'Interest Rate Layers', name: 'intRateLayers', getTenor,
    rateTypes: [COMMERCIAL_RATE, REFINANCE_RATE],
    fixedRates: { [REFINANCE_RATE]: REFINANCE_LENDING_RATE },
  });
}

// Saved layers are already contiguous from Month 01 to the tenor at the moment of Save; what can
// go stale is the tenor itself (or a draft from an older form), so each points back to Edit.
function validateRateLayers(layers, tenor) {
  const pad = (m) => String(m).padStart(2, '0');
  if (!layers.length) return 'Add the Interest Rate Layers — click Edit next to them.';
  const broken = layers.some((L, k) => L.fromMonth !== (k === 0 ? 1 : layers[k - 1].toMonth + 1)
    || L.toMonth < L.fromMonth || L.rate == null);
  if (broken) return 'The Interest Rate Layers are incomplete — click Edit next to them to fix them.';
  const last = layers[layers.length - 1];
  if (last.toMonth !== tenor) {
    return `The Interest Rate Layers run to Month ${pad(last.toMonth)}, but the loan’s final month is Month ${pad(tenor)} — click Edit next to them to update them.`;
  }
  return null;
}

// Engine shape: months + type + rate. Null when the loan has a single Offered Rate.
function rateLayerParams(inputs) {
  if (!inputs.rateLayered) return null;
  return inputs.intRateLayers.map(L => ({ from: L.fromMonth, to: L.toMonth, type: L.rateType, rate: L.rate }));
}

function setTwoLineLabel(field, line1, line2) {
  const lbl = field.querySelector('label');
  if (!lbl) return;
  const icon = lbl.querySelector('.info-icon');
  lbl.textContent = line1 + ' ';
  lbl.appendChild(el('span', { class: 'lbl-line2' }, line2));
  if (icon) lbl.appendChild(icon);
}

// Reset button + professional confirmation modal. On confirm, the tab's saved draft is
// cleared and the page re-rendered from scratch (blank fields, no uploads, no results).
// getState (optional) returns the page's serializable input state; it is snapshotted at
// render time — the form is built blank and drafts are restored only afterwards — so a
// click on a still-pristine page does nothing except a gentle note (no modal).
function resetButton(tabKey, rerender, getState = null) {
  const btn = el('button', { class: 'reset-btn', type: 'button' }, 'Reset');
  const pristine = getState ? JSON.stringify(getState()) : null;
  btn.addEventListener('click', () => {
    if (pristine !== null && JSON.stringify(getState()) === pristine) {
      toast('There is nothing to reset — no inputs have been made yet.');
      return;
    }
    const yes = el('button', { class: 'danger-btn', type: 'button' }, 'Yes, Reset');
    const back = el('button', { class: 'ghost-btn modal-ghost', type: 'button' }, 'Go Back');
    yes.addEventListener('click', () => { clearDraft(tabKey); closeModal(); rerender(); });
    back.addEventListener('click', closeModal);
    openModal(el('div', { class: 'confirm-card' },
      el('h3', {}, 'Reset all fields?'),
      el('p', {}, 'This will clear every input on this page — including any uploaded files and calculated results — so you can start a fresh calculation. This action cannot be undone.'),
      el('div', { class: 'confirm-actions' }, back, yes),
    ));
  });
  return btn;
}

// A hard block (Rate Revision modules): if the uploaded COF data does not reach back to the
// disbursement date, ERR cannot be computed — the earliest months would have no cost of fund.
// Show a big blocking popup with an X in the top-right corner; nothing proceeds until the user
// closes it. Returns true when COF covers from disbursement, false (and shows the popup) otherwise.
const COF_GAP_MSG = 'Please insert Cost of Fund data covering the period starting from the date of disbursement.';
function cofCoversDisbursement(cofData, disbursementISO) {
  return !!(cofData && cofData.length && disbursementISO && cofData[0].date <= disbursementISO);
}
function ensureCofCoversDisbursement(cofData, disbursementISO) {
  if (cofCoversDisbursement(cofData, disbursementISO)) return true;
  const x = el('button', { class: 'cof-block-x', type: 'button', title: 'Close', 'aria-label': 'Close' }, '×');
  const card = el('div', { class: 'cof-block-card' },
    x,
    el('div', { class: 'cof-block-icon' }, '⚠'),
    el('p', { class: 'cof-block-msg' },
      COF_GAP_MSG),
  );
  openModal(card);
  const mc = document.getElementById('modal-card');
  mc.classList.add('modal-card--cof');
  const dismiss = () => { mc.classList.remove('modal-card--cof'); closeModal(); };
  x.addEventListener('click', dismiss);
  const backdrop = document.querySelector('#modal-root .modal-backdrop');
  if (backdrop) backdrop.onclick = dismiss;
  return false;
}

// ============================================================
// REGULAR (STRUCTURED) LOAN FACILITY
// ============================================================
export function renderRegularLoan(root, pre = null) {
  root.innerHTML = '';
  const section = el('div', { class: 'section-card' });

  const loanAmount = numberField({ label: 'Loan Amount', name: 'loanAmount' });
  const offeredRate = percentField({ label: 'Offered Rate', name: 'offeredRate' });

  const moratoriumAvail = optionField({
    label: 'Moratorium Available?', name: 'moratoriumAvail',
    options: [{ label: 'select', value: '' }, 'No', 'Yes'], value: '',
    onChange: refresh,
  });
  const moratoriumPeriod = numberField({
    label: 'Moratorium Period (Months)', name: 'moratoriumPeriod', integerOnly: true, min: 1,
  });
  moratoriumPeriod.input.addEventListener('input', () => { refresh(); });

  const idpField = monthBoxesField({
    name: 'idpFlags', label: 'Interest During Moratorium Period',
    getCount: () => moratoriumPeriod.getValue() || 0, selectAll: true, capitalizable: true,
  });

  const loanTenor = numberField({ label: 'Loan Tenor (Months)', name: 'loanTenor', integerOnly: true, min: 1 });
  // Step 1 said the rate changes during the loan: the Interest Rate Layers table replaces the
  // single Offered Rate box.
  const layeredRates = !!(pre && pre.rateLayers === 'Yes');
  const intRateLayers = layeredRates ? interestRateLayersField(() => loanTenor.getValue()) : null;
  loanTenor.input.addEventListener('input', () => refresh());
  const paymentMode = optionField({
    label: 'Payment Mode', name: 'paymentMode',
    options: [{ label: 'select', value: '' }, 'EMI', 'EQI', 'Equal Principal + Interest (Monthly)', 'Equal Principal + Interest (Quarterly)', SPLIT_MODE],
    value: '',
    onChange: () => refresh(),
  });

  // ---- Split interest/principal fields (shown only for SPLIT_MODE) ----
  const isSplit = () => paymentMode.getValue() === SPLIT_MODE;
  const intFreq = optionField({
    label: 'Interest Payment Frequency', name: 'intFreq',
    options: FREQ_NAMES, value: 'Monthly', onChange: () => refresh(),
  });
  setTwoLineLabel(intFreq, 'Interest Payment', 'Frequency');
  const prinFreq = optionField({
    label: 'Principal Payment Frequency', name: 'prinFreq',
    options: FREQ_NAMES, value: 'Quarterly', onChange: () => refresh(),
  });
  setTwoLineLabel(prinFreq, 'Principal Payment', 'Frequency');
  const prinStart = numberField({
    label: 'Principal Payments Start From Month', name: 'prinStart', integerOnly: true, min: 1,
  });
  setTwoLineLabel(prinStart, 'Principal Payments', 'Start From Month');
  prinStart.setValue(1); // principal normally starts with the loan; a later month = principal grace
  prinStart.input.addEventListener('input', () => refresh());
  const prinHint = el('span', { class: 'help' });
  prinStart.appendChild(prinHint);
  const prinBasis = optionField({
    label: 'Principal Amount', name: 'prinBasis',
    options: ['Fixed (Equal)', 'Different per Date'], value: 'Fixed (Equal)', onChange: () => refresh(),
  });

  // Principal dates implied by the current inputs — drives the custom-amount boxes, the
  // locked months in the grid, and validation.
  function splitPrincipalMonths() {
    const tenor = loanTenor.getValue() || 0;
    const start = Math.max(1, prinStart.getValue() || 1);
    const period = FREQ[prinFreq.getValue()] || 1;
    if (!tenor || start > tenor) return [];
    return principalPaymentMonths(start, tenor, period);
  }

  // One amount box per principal date. The LAST box is read-only and shows the remainder, so
  // the schedule always amortises to zero however the earlier boxes are filled.
  const customWrap = el('div', { class: 'sub-card hidden' });
  const customBoxes = [];
  function rebuildCustomBoxes() {
    const months = splitPrincipalMonths();
    customWrap.innerHTML = '';
    customBoxes.length = 0;
    if (!months.length) return;
    customWrap.appendChild(el('label', {}, 'Principal Amount per Payment Date'));
    const gridEl = el('div', { class: 'form-row full' });
    months.forEach((m, i) => {
      const last = i === months.length - 1;
      const f = numberField({ label: `Month ${String(m).padStart(2, '0')}`, name: `cp${m}` });
      if (last) {
        f.input.readOnly = true;
        f.input.classList.add('readonly');
        f.setLabel(`Month ${String(m).padStart(2, '0')} (remainder)`);
      } else {
        f.input.addEventListener('input', updateCustomRemainder);
      }
      customBoxes.push(f);
      gridEl.appendChild(f);
    });
    customWrap.appendChild(gridEl);
    updateCustomRemainder();
  }
  function updateCustomRemainder() {
    if (!customBoxes.length) return;
    const loan = loanAmount.getValue() || 0;
    let used = 0;
    for (let i = 0; i < customBoxes.length - 1; i++) used += customBoxes[i].getValue() || 0;
    const rem = loan - used;
    const box = customBoxes[customBoxes.length - 1];
    box.input.value = formatNumber(rem, { decimals: 2 });
    box.classList.toggle('invalid', rem < 0);
  }

  // Whole-tenor interest grid. Principal months are locked to Paid; everything else is
  // pre-filled from the interest frequency and can be overridden month by month.
  const splitGrid = monthBoxesField({
    name: 'splitFlags', label: 'Interest Treatment by Month',
    getCount: () => (isSplit() ? (loanTenor.getValue() || 0) : 0),
    selectAll: true, capitalizable: true, groupByYear: true,
    lockedFn: (i) => splitPrincipalMonths().includes(i + 1),
    defaultFn: (i) => {
      const m = i + 1;
      const tenor = loanTenor.getValue() || 0;
      const ip = FREQ[intFreq.getValue()] || 1;
      return (m % ip === 0 || m === tenor) ? 1 : 0;
    },
  });

  const totalCof = percentField({ label: 'Total Cost of Fund [COF/ISC + OPEX]', name: 'totalCof' });
  setTwoLineLabel(totalCof, 'Total Cost of Fund', '(COF/ISC + OPEX)');

  function securityOptions() {
    const pm = paymentMode.getValue();
    const opts = ['No Funded Security', 'FDR', 'Cash Security'];
    // The split type has no single installment to size a security against (the legs pay on
    // different months and the amounts are uneven), so only the cash-backed options apply.
    if (pm === SPLIT_MODE) return [{ label: 'select', value: '' }, ...opts];
    // "after Moratorium" names a period the loan may not have. The VALUE never changes (saved
    // calculations, validation, metrics and the exports all key on it) — only the label does.
    if (pm === 'EMI') opts.push(installmentSecurityOption('EMI', moratoriumAvail.getValue() === 'Yes'));
    else if (pm === 'EQI') opts.push(installmentSecurityOption('EQI', moratoriumAvail.getValue() === 'Yes'));
    else if (pm) opts.push('Installment');
    return [{ label: 'select', value: '' }, ...opts];
  }
  const fundedSecurityType = optionField({
    label: 'Funded Security Type', name: 'fundedSecurityType', options: securityOptions(), value: '',
    onChange: refresh,
  });
  setTwoLineLabel(fundedSecurityType, 'Funded Security', 'Type');

  // Renamed after the chosen security type by labelSecurityFields(); these are placeholders.
  const csAmount = numberField({ label: 'Security Amount', name: 'csAmount' });
  const csRate = percentField({ label: 'Security Rate', name: 'csRate' });
  const numInst = numberField({ label: 'Number of Installments', name: 'numInst', integerOnly: true, min: 1 });

  section.appendChild(el('div', { class: 'form-row' }, loanAmount, ...(layeredRates ? [] : [offeredRate])));
  section.appendChild(el('div', { class: 'form-row' }, moratoriumAvail, moratoriumPeriod));
  const moraSection = el('div', { class: 'form-row full hidden' });
  moraSection.appendChild(el('div', { class: 'sub-card' }, idpField));
  section.appendChild(moraSection);
  section.appendChild(el('div', { class: 'form-row' }, loanTenor, paymentMode));
  // After the tenor, since the layers are counted in the loan's months.
  if (intRateLayers) section.appendChild(el('div', { class: 'sub-card' }, intRateLayers));
  // Split interest/principal block — frequencies, principal start + basis, the optional
  // per-date amount boxes, and the whole-tenor interest grid. Hidden for every other mode.
  const splitSection = el('div', { class: 'hidden' },
    el('div', { class: 'form-row' }, intFreq, prinFreq),
    el('div', { class: 'form-row' }, prinStart, prinBasis),
    customWrap,
    el('div', { class: 'form-row full' }, el('div', { class: 'sub-card' }, splitGrid)),
  );
  section.appendChild(splitSection);
  section.appendChild(el('div', { class: 'form-row' }, totalCof, fundedSecurityType));
  // Security detail row — fields depend on Funded Security Type (rebuilt in refresh()):
  //   FDR / Cash Security       -> Amount + Rate, both named after the chosen type
  //   EMI/EQI / Installment     -> Number of Installments + that type's Rate
  //   Installment               -> Number of Installments
  const secDetailRow = el('div', { class: 'form-row' });
  section.appendChild(secDetailRow);
  function rebuildSecurityRow() {
    const t = fundedSecurityType.getValue();
    secDetailRow.innerHTML = '';
    if (!t || t === 'No Funded Security') return; // No Funded Security => no detail fields, security = 0
    labelSecurityFields(t, moratoriumAvail.getValue() === 'Yes', csAmount, csRate);
    // Cash-backed securities state an amount; installment-built ones state how many instalments.
    if (t === 'FDR' || t === 'Cash Security') secDetailRow.append(csAmount, csRate);
    else secDetailRow.append(numInst, csRate);
  }

  function refresh() {
    const split = isSplit();
    // The split type carries no moratorium fields: a moratorium is expressed in the grid
    // (months set to Accrued/Capitalized) plus a later principal start, which is strictly
    // more flexible since those months need not share one treatment.
    const moraRow = moratoriumAvail.parentElement;
    if (moraRow) moraRow.classList.toggle('hidden', split);
    splitSection.classList.toggle('hidden', !split);

    const moraYes = !split && moratoriumAvail.getValue() === 'Yes';
    moratoriumPeriod.classList.toggle('hidden', !moraYes);
    const months = moraYes ? (moratoriumPeriod.getValue() || 0) : 0;
    moraSection.classList.toggle('hidden', months === 0);
    if (months > 0) idpField.refresh();

    paymentMode.setLabel(moraYes ? 'Payment Mode after Moratorium' : 'Payment Mode');
    loanTenor.setLabel(moraYes ? 'Loan Tenor including Moratorium (Months)' : 'Loan Tenor (Months)');

    if (split) {
      // Principal can never be settled more often than interest — the engine deducts
      // interest first, so every principal month must also be an interest month.
      const ok = FREQ[intFreq.getValue()] <= FREQ[prinFreq.getValue()];
      prinFreq.classList.toggle('invalid', !ok);
      customWrap.classList.toggle('hidden', prinBasis.getValue() !== 'Different per Date');
      rebuildCustomBoxes();
      prinHint.textContent = principalScheduleHint(splitPrincipalMonths(), FREQ[prinFreq.getValue()] || 1);
      splitGrid.refresh();
    }

    fundedSecurityType.setOptions(securityOptions());
    rebuildSecurityRow();
  }
  refresh();

  const fields = {
    loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
    loanTenor, paymentMode, totalCof, fundedSecurityType, csAmount, csRate, numInst,
    intFreq, prinFreq, prinStart, prinBasis, splitGrid, intRateLayers,
    getCustom: () => customBoxes.map(b => b.getValue() || 0),
  };
  section.appendChild(el('div', { class: 'action-bar' },
    // Re-render with the same step-1 answers, so a reset keeps the form the RM asked for.
    resetButton('regular', () => renderRegularLoan(root, pre), () => collectRegularInputs(fields))));
  const { left, right: resultsPanel } = workbench(root);
  left.appendChild(section);

  // Restore draft
  restoreDraft('regular', fields);
  // After the draft, never before: the entry answers are the fresher truth.
  applyEntryAnswers([
    [moratoriumAvail, pre && pre.moratorium],
    [paymentMode, pre && pre.modality],
  ]);
  refresh();
  attachDraftAutosave('regular', section, () => collectRegularInputs(fields));

  function compute() {
    const inputs = collectRegularInputs(fields);
    const msg = validateRegular(inputs);
    if (msg) return { msg };
    const split = inputs.paymentMode === SPLIT_MODE;
    const moraMonths = (!split && inputs.moratoriumAvail === 'Yes') ? inputs.moratoriumPeriod : 0;
    const isCs = inputs.fundedSecurityType === 'FDR' || inputs.fundedSecurityType === 'Cash Security';
    const rateLayers = rateLayerParams(inputs);
    const params = {
      loanAmount: inputs.loanAmount,
      ratePerYear: rateLayers ? rateLayers[0].rate : inputs.offeredRate,
      rateLayers,
      tenorMonths: inputs.loanTenor,
      paymentMode: inputs.paymentMode,
      moratoriumMonths: moraMonths,
      idpFlags: inputs.idpFlags,
      capFlags: inputs.capFlags,
      cofRate: inputs.totalCof,
      securityAmount: isCs ? (inputs.csAmount || 0) : 0,
      // FDR/Cash use their rate; EMI/EQI after Moratorium use the Funded Security Rate; Installment has none.
      securityRate: inputs.csRate || 0,
      securityKind: inputs.fundedSecurityType,
      numInst: inputs.numInst,
    };
    if (split) {
      Object.assign(params, {
        interestPeriod: FREQ[inputs.intFreq],
        principalPeriod: FREQ[inputs.prinFreq],
        principalStartMonth: inputs.prinStart,
        principalBasis: inputs.prinBasis === 'Different per Date' ? 'custom' : 'fixed',
        customPrincipals: inputs.customPrincipals,
        intFlags: inputs.splitFlags,
      });
    }
    const schedule = split ? buildSplitSchedule(params) : buildStructuredSchedule(params);
    const metrics = computeMetrics(schedule, params);
    const ctx = { pageType: 'regular', pageTitle: 'Loan Facilities — Structured', inputs, params, schedule, metrics };
    return { ctx };
  }
  attachLiveRecalc(section, resultsPanel, compute);
}

function collectRegularInputs(f) {
  return {
    offeredRate: f.offeredRate.getValue(),
    rateLayered: !!f.intRateLayers,
    intRateLayers: f.intRateLayers ? f.intRateLayers.getValue() : [],
    loanAmount: f.loanAmount.getValue(),
    moratoriumAvail: f.moratoriumAvail.getValue(),
    moratoriumPeriod: f.moratoriumPeriod.getValue() || 0,
    idpFlags: f.idpField.getPaidFlags(),
    capFlags: f.idpField.getCapFlags(),
    loanTenor: f.loanTenor.getValue(),
    paymentMode: f.paymentMode.getValue(),
    totalCof: f.totalCof.getValue(),
    fundedSecurityType: f.fundedSecurityType.getValue(),
    csAmount: f.csAmount.getValue(),
    csRate: f.csRate.getValue(),
    numInst: f.numInst.getValue(),
    // Split interest/principal type
    intFreq: f.intFreq ? f.intFreq.getValue() : null,
    prinFreq: f.prinFreq ? f.prinFreq.getValue() : null,
    prinStart: f.prinStart ? (f.prinStart.getValue() || 1) : 1,
    prinBasis: f.prinBasis ? f.prinBasis.getValue() : null,
    splitStates: f.splitGrid ? f.splitGrid.getValue() : [],
    splitFlags: f.splitGrid ? f.splitGrid.getIntFlags() : [],
    customPrincipals: f.getCustom ? f.getCustom() : [],
  };
}

const secName = (i) => securityDisplayName(i.fundedSecurityType, i.moratoriumAvail === 'Yes');

function validateRegular(i) {
  if (!i.loanAmount) return fail('Enter Loan Amount.');
  if (!i.rateLayered && i.offeredRate === null) return fail('Enter Offered Rate.');
  if (!i.loanTenor) return fail('Enter Loan Tenor.');
  if (i.rateLayered) { const m = validateRateLayers(i.intRateLayers, i.loanTenor); if (m) return fail(m); }
  if (!i.paymentMode) return fail('Select a Payment Mode.');
  if (i.paymentMode === SPLIT_MODE) {
    if (FREQ[i.intFreq] > FREQ[i.prinFreq])
      return fail(`Interest cannot be paid less often than principal — ${i.intFreq} interest with ${i.prinFreq} principal leaves a principal month with no interest settlement.`);
    if (i.prinStart < 1 || i.prinStart > i.loanTenor)
      return fail(`Principal Payments Start From Month must be between 1 and ${i.loanTenor}.`);
    if (i.prinBasis === 'Different per Date') {
      const dates = principalPaymentMonths(i.prinStart, i.loanTenor, FREQ[i.prinFreq]);
      const earlier = i.customPrincipals.slice(0, Math.max(0, dates.length - 1));
      const typed = earlier.reduce((s, v) => s + (v || 0), 0);
      if (typed > i.loanAmount)
        return fail('The principal amounts entered exceed the Loan Amount.');
      // Every box blank silently turns the loan into a bullet — the whole principal lands on
      // the final date. That is a real structure, but it should be chosen, not fallen into.
      if (dates.length > 1 && !earlier.some(v => v))
        return fail('Enter the principal amount for at least one payment date, or switch Principal Amount to "Fixed (Equal)".');
    }
    if (i.totalCof === null) return fail('Enter Total Cost of Fund.');
    if (!i.fundedSecurityType) return fail('Select a Funded Security Type.');
    if ((i.fundedSecurityType === 'FDR' || i.fundedSecurityType === 'Cash Security')
        && (i.csRate === null || i.csAmount === null))
      return fail(`Enter ${secName(i)} Amount and Rate.`);
    return null;
  }
  if (!i.moratoriumAvail) return fail('Select whether a moratorium is available.');
  if (i.moratoriumAvail === 'Yes' && !i.moratoriumPeriod) return fail('Enter Moratorium Period.');
  if (i.moratoriumAvail === 'Yes' && i.loanTenor <= i.moratoriumPeriod) return fail('Loan Tenor must exceed Moratorium Period.');
  if (i.totalCof === null) return fail('Enter Total Cost of Fund.');
  if (!i.fundedSecurityType) return fail('Select a Funded Security Type.');
  if (i.fundedSecurityType === 'FDR' || i.fundedSecurityType === 'Cash Security') {
    if (i.csRate === null) return fail(`Enter ${secName(i)} Rate.`);
    if (i.csAmount === null) return fail(`Enter ${secName(i)} Amount.`);
  }
  if (['EMI after Moratorium', 'EQI after Moratorium', 'Installment'].includes(i.fundedSecurityType) && !i.numInst)
    return fail('Enter Number of Installments.');
  return null;
}
// Returns the reason as a string; live recalc shows it in the results pane rather than
// firing a toast on every keystroke.
function fail(msg) { return msg; }

// ============================================================
// CUSTOMIZED LOAN FACILITY
// ============================================================
export function renderCustomizedLoan(root, pre = null) {
  root.innerHTML = '';
  const section = el('div', { class: 'section-card' });

  const loanAmount = numberField({ label: 'Loan Amount', name: 'loanAmount' });
  const offeredRate = percentField({ label: 'Offered Rate', name: 'offeredRate' });
  const moratoriumAvail = optionField({ label: 'Moratorium Available?', name: 'moratoriumAvail', options: [{ label: 'select', value: '' }, 'No', 'Yes'], value: '', onChange: refresh });
  const moratoriumPeriod = numberField({ label: 'Moratorium Period (Months)', name: 'moratoriumPeriod', integerOnly: true, min: 1 });
  moratoriumPeriod.input.addEventListener('input', () => { refresh(); refreshLayerOpts(); paymentLayers.applyLayerRules(); });
  const idpField = monthBoxesField({
    name: 'idpFlags', label: 'Interest During Moratorium Period',
    getCount: () => moratoriumPeriod.getValue() || 0, selectAll: true, capitalizable: true,
  });
  const loanTenor = numberField({ label: 'Loan Tenor (Months)', name: 'loanTenor', integerOnly: true, min: 1 });
  const layeredRates = !!(pre && pre.rateLayers === 'Yes');
  const intRateLayers = layeredRates ? interestRateLayersField(() => loanTenor.getValue()) : null;
  loanTenor.input.addEventListener('input', () => { refreshLayerOpts(); refresh(); paymentLayers.applyLayerRules(); refreshSplitGrid(); });

  function fromOptions() {
    const tenor = loanTenor.getValue() || 0;
    const mora = moratoriumAvail.getValue() === 'Yes' ? (moratoriumPeriod.getValue() || 0) : 0;
    const arr = [];
    for (let i = mora + 1; i <= tenor; i++) arr.push({ value: String(i), label: String(i).padStart(2, '0') });
    return arr;
  }
  // From and To share the same month list (mora+1 .. tenor) — mirrors the Lending Rate Layers
  // in Rate Revision — Structured. The last layer's To auto-defaults to the final month (the
  // universal cascade engine fills it from getMaturity) and stays editable.
  function toOptions() {
    const tenor = loanTenor.getValue() || 0;
    const mora = moratoriumAvail.getValue() === 'Yes' ? (moratoriumPeriod.getValue() || 0) : 0;
    const arr = [];
    for (let i = mora + 1; i <= tenor; i++) arr.push({ value: String(i), label: String(i).padStart(2, '0') });
    return arr;
  }

  const paymentLayers = layeredField({
    label: 'Payment Layers',
    name: 'paymentLayers',
    schema: [
      { key: 'fromInstallment', label: 'From Month', type: 'option', options: fromOptions, allowEmpty: true, placeholder: '', width: '0.8fr', readOnly: true },
      { key: 'toInstallment', label: 'To Month', type: 'option', options: toOptions, allowEmpty: true, placeholder: 'select', width: '0.8fr' },
      { key: 'paymentType', label: 'Payment Type', type: 'option', allowEmpty: true, placeholder: 'select', options: [
          'Customized Principal (Monthly)', 'Customized Principal (Quarterly)', 'EMI', 'EQI',
          'Equal Principal + Interest (Monthly)', 'Equal Principal + Interest (Quarterly)',
          SPLIT_MODE,
        ], width: '1.3fr' },
      // Only meaningful for SPLIT_MODE layers; disabled for every other type. The layer's own
      // From month anchors both legs, so there is no per-layer principal start — an
      // interest-only prefix is just an earlier layer.
      { key: 'intFreq', label: 'Interest Freq.', type: 'option', allowEmpty: true, placeholder: '—', options: FREQ_NAMES, width: '0.9fr' },
      { key: 'prinFreq', label: 'Principal Freq.', type: 'option', allowEmpty: true, placeholder: '—', options: FREQ_NAMES, width: '0.9fr' },
      // For SPLIT_MODE this is the amount paid on EACH principal date; blank divides the
      // balance equally. Varying amounts per date are expressed by splitting into layers.
      { key: 'customPrincipal', label: 'Custom Principal', type: 'number', width: '1.1fr' },
    ],
    addLabel: '+ Add Payment Layer',
    minRows: 2,
    initialRows: 2,
    cascadingFromKey: 'fromInstallment',
    cascadingToKey: 'toInstallment',
    getMaturity: () => {
      const tenor = loanTenor.getValue();
      return { value: tenor ? String(tenor) : null, kind: 'month' };
    },
    getAnchor: () => {
      const mora = moratoriumAvail.getValue() === 'Yes' ? (moratoriumPeriod.getValue() || 0) : 0;
      return { value: String(mora + 1), kind: 'month' };
    },
    allowFromEqualTo: true,
    onChange: () => {
      syncOptionalColumns();
      paymentLayers.rows.forEach((row) => {
        const ptype = row.inputs.paymentType.value;
        const split = ptype === SPLIT_MODE;
        // Custom Principal: required for "Customized Principal", optional for the split type
        // (blank = divide equally), meaningless for EMI/EQI/Equal-Principal.
        const wantsCp = split || (!!ptype && ptype.startsWith('Customized Principal'));
        setCellActive(row, 'customPrincipal', wantsCp);
        if (!wantsCp) row.inputs.customPrincipal.value = '';
        [['intFreq', 'Monthly'], ['prinFreq', 'Quarterly']].forEach(([key, fallback]) => {
          const f = row.inputs[key];
          if (!f) return;
          setCellActive(row, key, split);
          if (!split) f.value = '';
          else if (!f.value) f.value = fallback;
        });
      });
      refreshSplitGrid();
    },
  });
  // Toast on the layered field's "cannot add" callback (e.g. last layer already ends at maturity)
  paymentLayers.onCannotAdd = (msg) => toast(msg, 'error');
  function refreshLayerOpts() { paymentLayers.refreshOptions(); }
  // A column that no current layer can use is dropped from the grid rather than shown
  // greyed out — a disabled box still reads as something to fill in. Each column reappears
  // as soon as one layer needs it and stays while any layer still does, so a mixed set
  // (EMI in one layer, split in another) shows them.
  // A cell the row's own payment type can't use is left blank rather than greyed out: the
  // column exists for whichever layers need it, so an empty cell reads as "not this layer".
  function setCellActive(row, key, active) {
    const inp = row.inputs[key];
    if (inp) inp.disabled = !active;
    if (row.cells && row.cells[key]) row.cells[key].classList.toggle('cell-na', !active);
  }
  function syncOptionalColumns() {
    const types = paymentLayers.rows.map(r => r.inputs.paymentType.value || '');
    const anySplit = types.some(t => t === SPLIT_MODE);
    // Custom Principal serves the split type AND both "Customized Principal" types.
    const anyCustomPrincipal = types.some(t => t === SPLIT_MODE || t.startsWith('Customized Principal'));
    paymentLayers.setColumnHidden('intFreq', !anySplit);
    paymentLayers.setColumnHidden('prinFreq', !anySplit);
    paymentLayers.setColumnHidden('customPrincipal', !anyCustomPrincipal);
  }

  // ---- Split interest/principal support -------------------------------------------------
  // One whole-tenor grid serves every split layer. Months belonging to other layer types are
  // rendered inert, since their interest is decided by that layer's own payment type.
  function splitLayers() {
    return paymentLayers.getValue()
      .filter(r => r.paymentType === SPLIT_MODE && r.fromInstallment && r.toInstallment)
      .map(r => ({
        from: Number(r.fromInstallment),
        to: r.toInstallment === 'LAST' ? (loanTenor.getValue() || 0) : Number(r.toInstallment),
        ip: FREQ[r.intFreq] || 1,
        pp: FREQ[r.prinFreq] || 1,
      }));
  }
  const splitLayerAt = (m) => splitLayers().find(L => m >= L.from && m <= L.to) || null;
  const custSplitGrid = monthBoxesField({
    name: 'custSplitFlags', label: 'Interest Treatment by Month',
    getCount: () => (splitLayers().length ? (loanTenor.getValue() || 0) : 0),
    selectAll: true, capitalizable: true, groupByYear: true,
    disabledFn: (i) => !splitLayerAt(i + 1),
    lockedFn: (i) => {
      const L = splitLayerAt(i + 1);
      if (!L) return false;
      const m = i + 1;
      return ((m - L.from + 1) % L.pp === 0) || m === L.to; // principal month
    },
    defaultFn: (i) => {
      const L = splitLayerAt(i + 1);
      if (!L) return 0;
      const m = i + 1;
      return (((m - L.from + 1) % L.ip === 0) || m === L.to) ? 1 : 0;
    },
  });
  const custSplitWrap = el('div', { class: 'form-row full hidden' },
    el('div', { class: 'sub-card' }, custSplitGrid));
  function refreshSplitGrid() {
    const any = splitLayers().length > 0;
    custSplitWrap.classList.toggle('hidden', !any);
    if (any) custSplitGrid.refresh();
  }

  const totalCof = percentField({ label: 'Total Cost of Fund [COF/ISC + OPEX]', name: 'totalCof' });
  setTwoLineLabel(totalCof, 'Total Cost of Fund', '(COF/ISC + OPEX)');
  function customizedSecurityOptions() {
    const moraYes = moratoriumAvail.getValue() === 'Yes';
    return [{ label: 'select', value: '' }, 'No Funded Security', 'FDR', 'Cash Security',
      installmentSecurityOption('EMI', moraYes), installmentSecurityOption('EQI', moraYes)];
  }
  const fundedSecurityType = optionField({
    label: 'Funded Security Type', name: 'fundedSecurityType',
    options: customizedSecurityOptions(), value: '', onChange: refresh,
  });
  setTwoLineLabel(fundedSecurityType, 'Funded Security', 'Type');
  // Renamed after the chosen security type by labelSecurityFields(); these are placeholders.
  const csAmount = numberField({ label: 'Security Amount', name: 'csAmount' });
  const csRate = percentField({ label: 'Security Rate', name: 'csRate' });
  const numInst = numberField({ label: 'Number of Installments', name: 'numInst', integerOnly: true, min: 1 });

  section.appendChild(el('div', { class: 'form-row' }, loanAmount, ...(layeredRates ? [] : [offeredRate])));
  section.appendChild(el('div', { class: 'form-row' }, moratoriumAvail, moratoriumPeriod));
  const moraSection = el('div', { class: 'form-row full hidden' });
  moraSection.appendChild(el('div', { class: 'sub-card' }, idpField));
  section.appendChild(moraSection);
  section.appendChild(el('div', { class: 'form-row' }, loanTenor));
  if (intRateLayers) section.appendChild(el('div', { class: 'sub-card' }, intRateLayers));
  section.appendChild(el('div', { class: 'sub-card' }, paymentLayers));
  section.appendChild(custSplitWrap);
  section.appendChild(el('div', { class: 'form-row' }, totalCof, fundedSecurityType));
  // Security detail row — depends on Funded Security Type (rebuilt in refresh()).
  const secDetailRow = el('div', { class: 'form-row' });
  section.appendChild(secDetailRow);
  function rebuildSecurityRow() {
    const t = fundedSecurityType.getValue();
    secDetailRow.innerHTML = '';
    if (!t || t === 'No Funded Security') return; // No Funded Security => no detail fields, security = 0
    labelSecurityFields(t, moratoriumAvail.getValue() === 'Yes', csAmount, csRate);
    // Cash-backed securities state an amount; installment-built ones state how many instalments.
    if (t === 'FDR' || t === 'Cash Security') secDetailRow.append(csAmount, csRate);
    else secDetailRow.append(numInst, csRate);
  }

  function refresh() {
    const moraYes = moratoriumAvail.getValue() === 'Yes';
    moratoriumPeriod.classList.toggle('hidden', !moraYes);
    const months = moraYes ? (moratoriumPeriod.getValue() || 0) : 0;
    moraSection.classList.toggle('hidden', months === 0);
    if (months > 0) idpField.refresh();
    loanTenor.setLabel(moraYes ? 'Loan Tenor including Moratorium (Months)' : 'Loan Tenor (Months)');
    // Relabel the installment-sized security options when the moratorium is toggled
    // (setOptions keeps the current selection, since only the label changes).
    fundedSecurityType.setOptions(customizedSecurityOptions());
    rebuildSecurityRow();
    syncOptionalColumns();
  }
  refresh();

  const fields = {
    loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
    loanTenor, paymentLayers, totalCof, fundedSecurityType, csAmount, csRate, numInst,
    custSplitGrid, intRateLayers,
  };
  section.appendChild(el('div', { class: 'action-bar' },
    resetButton('customized', () => renderCustomizedLoan(root, pre), () => collectCustomizedInputs(fields))));
  const { left, right: resultsPanel } = workbench(root);
  left.appendChild(section);

  restoreDraft('customized', fields);
  // After the draft, never before: the entry answers are the fresher truth.
  applyEntryAnswers([
    [moratoriumAvail, pre && pre.moratorium],
  ]);
  refresh(); refreshLayerOpts();
  attachDraftAutosave('customized', section, () => collectCustomizedInputs(fields));

  function compute() {
    const inputs = collectCustomizedInputs(fields);
    const err = validateCustomized(inputs);
    if (err) return { msg: err };

    const mora = inputs.moratoriumAvail === 'Yes' ? inputs.moratoriumPeriod : 0;
    const isCs = inputs.fundedSecurityType === 'FDR' || inputs.fundedSecurityType === 'Cash Security';
    const rateLayers = rateLayerParams(inputs);
    const params = {
      loanAmount: inputs.loanAmount,
      ratePerYear: rateLayers ? rateLayers[0].rate : inputs.offeredRate,
      rateLayers,
      tenorMonths: inputs.loanTenor,
      moratoriumMonths: mora,
      idpFlags: inputs.idpFlags,
      capFlags: inputs.capFlags,
      cofRate: inputs.totalCof,
      layers: inputs.paymentLayers,
      intFlags: inputs.custSplitFlags,
    };
    const schedule = buildCustomizedSchedule(params);
    const metrics = computeMetrics(schedule, {
      ...params, paymentMode: 'EMI',
      securityAmount: isCs ? (inputs.csAmount || 0) : 0,
      // FDR/Cash use their rate; EMI/EQI after Moratorium use the Funded Security Rate; Installment has none.
      securityRate: inputs.csRate || 0,
      securityKind: inputs.fundedSecurityType,
      numInst: inputs.numInst,
    });
    const ctx = { pageType: 'customized', pageTitle: 'Loan Facilities — Customized', inputs, params, schedule, metrics };
    return { ctx };
  }
  attachLiveRecalc(section, resultsPanel, compute);
}

function collectCustomizedInputs(f) {
  const tenor = f.loanTenor.getValue();
  return {
    offeredRate: f.offeredRate.getValue(),
    rateLayered: !!f.intRateLayers,
    intRateLayers: f.intRateLayers ? f.intRateLayers.getValue() : [],
    loanAmount: f.loanAmount.getValue(),
    moratoriumAvail: f.moratoriumAvail.getValue(),
    moratoriumPeriod: f.moratoriumPeriod.getValue() || 0,
    idpFlags: f.idpField.getPaidFlags(),
    capFlags: f.idpField.getCapFlags(),
    loanTenor: tenor,
    paymentLayers: f.paymentLayers.getValue().map(r => ({
      fromInstallment: r.fromInstallment ? Number(r.fromInstallment) : null,
      // "LAST" sentinel -> last tenor month
      toInstallment: r.toInstallment === 'LAST' ? Number(tenor || 0) : (r.toInstallment ? Number(r.toInstallment) : null),
      paymentType: r.paymentType,
      customPrincipal: r.customPrincipal,
      // Split layers only — the engine recognises a layer by these two being set.
      interestPeriod: r.paymentType === SPLIT_MODE ? (FREQ[r.intFreq] || 1) : null,
      principalPeriod: r.paymentType === SPLIT_MODE ? (FREQ[r.prinFreq] || 1) : null,
      intFreq: r.intFreq, prinFreq: r.prinFreq,
    })),
    custSplitStates: f.custSplitGrid ? f.custSplitGrid.getValue() : [],
    custSplitFlags: f.custSplitGrid ? f.custSplitGrid.getIntFlags() : [],
    totalCof: f.totalCof.getValue(),
    fundedSecurityType: f.fundedSecurityType.getValue(),
    csAmount: f.csAmount.getValue(),
    csRate: f.csRate.getValue(),
    numInst: f.numInst.getValue(),
  };
}

function validateCustomized(i) {
  if (!i.rateLayered && i.offeredRate === null) return 'Enter Offered Rate.';
  if (!i.loanAmount) return 'Enter Loan Amount.';
  if (!i.loanTenor) return 'Enter Loan Tenor.';
  if (i.rateLayered) { const m = validateRateLayers(i.intRateLayers, i.loanTenor); if (m) return m; }
  if (i.totalCof === null) return 'Enter Total Cost of Fund.';
  if (!i.fundedSecurityType) return 'Select a Funded Security Type.';
  if (!i.moratoriumAvail) return 'Select whether a moratorium is available.';
  if (i.moratoriumAvail === 'Yes') {
    if (!i.moratoriumPeriod) return 'Enter Moratorium Period.';
    if (i.loanTenor <= i.moratoriumPeriod) return 'Loan Tenor must exceed Moratorium Period.';
  }
  if (!i.paymentLayers.length) return 'Add at least one Payment Layer.';
  // Layer count cap
  if (i.paymentLayers.length > i.loanTenor) return `Cannot have more than ${i.loanTenor} payment layers (tenor).`;

  const mora = i.moratoriumAvail === 'Yes' ? i.moratoriumPeriod : 0;
  const startMonth = mora + 1;
  const endMonth = i.loanTenor;

  // Per-layer validation
  for (let k = 0; k < i.paymentLayers.length; k++) {
    const L = i.paymentLayers[k];
    if (!L.paymentType) return `Layer ${k + 1}: select a Payment Type.`;
    if (!L.fromInstallment || !L.toInstallment) return `Layer ${k + 1}: select both From and To months.`;
    if (L.toInstallment < L.fromInstallment) return `Layer ${k + 1}: To must be ≥ From.`;
    if (L.fromInstallment < startMonth) return `Layer ${k + 1}: From must be ≥ Month ${String(startMonth).padStart(2, '0')}.`;
    if (L.toInstallment > endMonth) return `Layer ${k + 1}: To must be ≤ Month ${String(endMonth).padStart(2, '0')}.`;
    if (L.paymentType && L.paymentType.startsWith('Customized Principal') && !L.customPrincipal)
      return `Layer ${k + 1}: enter Custom Principal for the "${L.paymentType}" type.`;
    if (L.paymentType === SPLIT_MODE) {
      if (!L.intFreq || !L.prinFreq)
        return `Layer ${k + 1}: select both an Interest Freq. and a Principal Freq.`;
      if (FREQ[L.intFreq] > FREQ[L.prinFreq])
        return `Layer ${k + 1}: interest cannot be paid less often than principal — ${L.intFreq} interest with ${L.prinFreq} principal leaves a principal month with no interest settlement.`;
    }
  }
  // Sort and check for overlaps / gaps
  const sorted = i.paymentLayers.slice().sort((a, b) => a.fromInstallment - b.fromInstallment);
  for (let k = 0; k < sorted.length; k++) {
    if (k === 0 && sorted[k].fromInstallment !== startMonth)
      return `Payment layers must start at Month ${String(startMonth).padStart(2, '0')}. First layer starts at Month ${String(sorted[k].fromInstallment).padStart(2, '0')}.`;
    if (k > 0) {
      const prev = sorted[k - 1];
      const cur = sorted[k];
      if (cur.fromInstallment <= prev.toInstallment)
        return `Layers overlap: Layer ending at Month ${String(prev.toInstallment).padStart(2, '0')} conflicts with layer starting at Month ${String(cur.fromInstallment).padStart(2, '0')}.`;
      if (cur.fromInstallment !== prev.toInstallment + 1)
        return `Gap between layers: nothing covers Month ${String(prev.toInstallment + 1).padStart(2, '0')} to Month ${String(cur.fromInstallment - 1).padStart(2, '0')}.`;
    }
  }
  if (sorted[sorted.length - 1].toInstallment !== endMonth)
    return `Last payment layer must end at Month ${String(endMonth).padStart(2, '0')} (loan tenor). Currently ends at Month ${String(sorted[sorted.length - 1].toInstallment).padStart(2, '0')}.`;
  return null;
}

// ============================================================
// RATE REVISION — STRUCTURED
// ============================================================
export function renderRateRevisionStructured(root, pre = null) {
  root.innerHTML = '';
  const section = el('div', { class: 'section-card' });

  const initialAmount = numberField({ label: 'Initial Loan Amount', name: 'initialAmount' });
  const disbursementDate = dateField({
    label: 'Disbursement Date', name: 'disbursementDate',
    disableFn: (d) => {
      const day = d.getDay(); // 5=Fri, 6=Sat
      return day === 5 || day === 6;
    },
    // flatpickr does NOT fire a native 'change' event on date pick — route its
    // onChange here so the layer cascade (first From = disbursement, last To = maturity) updates.
    onChange: () => rerunLayerRules(),
  });

  const moratoriumAvail = optionField({ label: 'Moratorium Given at Disbursement?', name: 'moratoriumAvail', options: [{ label: 'select', value: '' }, 'No', 'Yes'], value: '', onChange: refresh });
  const moratoriumPeriod = numberField({ label: 'Moratorium Period (Months)', name: 'moratoriumPeriod', integerOnly: true, min: 1 });
  moratoriumPeriod.input.addEventListener('input', refresh);

  const idpField = monthBoxesField({
    name: 'idpFlags', label: 'Interest During Moratorium Period',
    getCount: () => moratoriumPeriod.getValue() || 0, selectAll: true, capitalizable: true,
  });

  const paymentModality = optionField({
    label: 'Payment Modality', name: 'paymentModality',
    options: [{ label: 'select', value: '' }, 'EMI', 'EQI', 'Equal Principal + Interest (Monthly)', 'Equal Principal + Interest (Quarterly)', SPLIT_MODE], value: '',
    onChange: () => refresh(),
  });
  const tenorMonths = numberField({ label: 'Loan Tenor at Disbursement (Months)', name: 'tenorMonths', integerOnly: true, min: 1 });

  // ---- Split interest/principal modality (mirrors Loan Facilities - Structured) ----
  const isSplit = () => paymentModality.getValue() === SPLIT_MODE;
  const rrIntFreq = optionField({ label: 'Interest Payment Frequency', name: 'rrIntFreq',
    options: FREQ_NAMES, value: 'Monthly', onChange: () => refresh() });
  setTwoLineLabel(rrIntFreq, 'Interest Payment', 'Frequency');
  const rrPrinFreq = optionField({ label: 'Principal Payment Frequency', name: 'rrPrinFreq',
    options: FREQ_NAMES, value: 'Quarterly', onChange: () => refresh() });
  setTwoLineLabel(rrPrinFreq, 'Principal Payment', 'Frequency');
  const rrPrinStart = numberField({ label: 'Principal Payments Start From Month', name: 'rrPrinStart', integerOnly: true, min: 1 });
  setTwoLineLabel(rrPrinStart, 'Principal Payments', 'Start From Month');
  rrPrinStart.setValue(1);
  rrPrinStart.input.addEventListener('input', () => refresh());
  const rrPrinHint = el('span', { class: 'help' });
  rrPrinStart.appendChild(rrPrinHint);
  const rrPrinBasis = optionField({ label: 'Principal Amount', name: 'rrPrinBasis',
    options: ['Fixed (Equal)', 'Different per Date'], value: 'Fixed (Equal)', onChange: () => refresh() });

  function rrPrincipalMonths() {
    const t = tenorMonths.getValue() || 0;
    const st = Math.max(1, rrPrinStart.getValue() || 1);
    const per = FREQ[rrPrinFreq.getValue()] || 1;
    if (!t || st > t) return [];
    return principalPaymentMonths(st, t, per);
  }
  const rrCustomWrap = el('div', { class: 'sub-card hidden' });
  const rrCustomBoxes = [];
  function rrRebuildCustom() {
    const months = rrPrincipalMonths();
    rrCustomWrap.innerHTML = ''; rrCustomBoxes.length = 0;
    if (!months.length) return;
    rrCustomWrap.appendChild(el('label', {}, 'Principal Amount per Payment Date'));
    const g = el('div', { class: 'form-row full' });
    months.forEach((mth, i) => {
      const last = i === months.length - 1;
      const f = numberField({ label: `Month ${String(mth).padStart(2, '0')}`, name: `rrcp${mth}` });
      if (last) { f.input.readOnly = true; f.input.classList.add('readonly'); f.setLabel(`Month ${String(mth).padStart(2, '0')} (remainder)`); }
      else f.input.addEventListener('input', rrUpdateRemainder);
      rrCustomBoxes.push(f); g.appendChild(f);
    });
    rrCustomWrap.appendChild(g);
    rrUpdateRemainder();
  }
  function rrUpdateRemainder() {
    if (!rrCustomBoxes.length) return;
    const loan = initialAmount.getValue() || 0;
    let used = 0;
    for (let i = 0; i < rrCustomBoxes.length - 1; i++) used += rrCustomBoxes[i].getValue() || 0;
    const rem = loan - used;
    const box = rrCustomBoxes[rrCustomBoxes.length - 1];
    box.input.value = formatNumber(rem, { decimals: 2 });
    box.classList.toggle('invalid', rem < 0);
  }
  const rrSplitGrid = monthBoxesField({
    name: 'rrSplitFlags', label: 'Interest Treatment by Month',
    getCount: () => (isSplit() ? (tenorMonths.getValue() || 0) : 0),
    selectAll: true, capitalizable: true, groupByYear: true,
    lockedFn: (i) => rrPrincipalMonths().includes(i + 1),
    defaultFn: (i) => {
      const mth = i + 1, t = tenorMonths.getValue() || 0;
      const ip = FREQ[rrIntFreq.getValue()] || 1;
      return (mth % ip === 0 || mth === t) ? 1 : 0;
    },
  });
  const rrSplitSection = el('div', { class: 'hidden' },
    el('div', { class: 'form-row' }, rrIntFreq, rrPrinFreq),
    el('div', { class: 'form-row' }, rrPrinStart, rrPrinBasis),
    rrCustomWrap,
    el('div', { class: 'form-row full' }, el('div', { class: 'sub-card' }, rrSplitGrid)),
  );

  // Actual loan maturity = disbursement + tenor months − 1 day.
  // e.g. 01-Jan-2020 + 60 months → 01-Jan-2025, minus 1 day → 31-Dec-2024.
  // Used only for the last layer's To Date (Lending Rate + Loan Security layers).
  function maturityISO() {
    const d = disbursementDate.getValue();
    const t = tenorMonths.getValue();
    if (!d || !t) return null;
    // Month-t due date per the due-day convention (Feb -> last day for 28th-31st
    // anchors; no rollover into March), then one day back.
    const dt = addMonthsDue(d, t);
    dt.setDate(dt.getDate() - 1);
    return dt.toISOString().slice(0, 10);
  }

  // From-only layers: each layer's rate applies from its From Date until the day before the
  // next layer's From (the earliest layer also extends back to disbursement; the final layer
  // runs to maturity). Read-only table + single "Edit" popup that captures multiple days, one
  // rate each. Output is { fromDate, activeRate }[] — unchanged for the engine + Verify Excel.
  const rateLayers = rateLayersField({
    label: 'Lending Rate Layers',
    name: 'rateLayers',
    getAnchor: () => ({ value: disbursementDate.getValue() }),
    getMaturity: () => maturityISO(),
  });

  const securityLayers = securityLayersField({
    label: 'Loan Security Layers',
    name: 'securityLayers',
    getAnchor: () => ({ value: disbursementDate.getValue() }),
    getMaturity: () => maturityISO(),
  });

  // External inputs (disbursement / tenor) feed the cascade engine — re-run on change.
  // Disbursement is wired via dateField's onChange above (flatpickr quirk); tenor is a
  // plain text input so 'input' works.
  function rerunLayerRules() {
    rateLayers.applyLayerRules();
    securityLayers.applyLayerRules();
  }
  tenorMonths.input.addEventListener('input', () => { rerunLayerRules(); refresh(); });

  // COF Data Upload — upload button on the left, "Download Sample File" link below it.
  const cofUpload = cofUploadField();
  const cofField = cofUpload.field;

  // Optional free-text reference — prefixed onto every downloaded file's name for this calc.
  const referenceField = textField({
    label: 'Add Reference (Optional)', name: 'reference', placeholder: 'e.g. loan account / proposal no.',
  });

  section.appendChild(el('div', { class: 'form-row' }, initialAmount, disbursementDate));
  section.appendChild(el('div', { class: 'form-row' }, moratoriumAvail, moratoriumPeriod));
  const moraSection = el('div', { class: 'form-row full hidden' });
  moraSection.appendChild(el('div', { class: 'sub-card' }, idpField));
  section.appendChild(moraSection);
  section.appendChild(el('div', { class: 'form-row' }, paymentModality, tenorMonths));
  section.appendChild(rrSplitSection);
  section.appendChild(el('div', { class: 'layer-panel' }, rateLayers));
  section.appendChild(el('div', { class: 'layer-panel' }, securityLayers));
  section.appendChild(cofField);
  section.appendChild(el('div', { class: 'form-row' }, referenceField));

  function refresh() {
    // This modality carries no moratorium: a moratorium is expressed in the grid plus a later
    // principal start, exactly as in Loan Facilities - Structured.
    const split = isSplit();
    const moraRow = moratoriumAvail.parentElement;
    if (moraRow) moraRow.classList.toggle('hidden', split);
    rrSplitSection.classList.toggle('hidden', !split);

    const moraYes = !split && moratoriumAvail.getValue() === 'Yes';
    moratoriumPeriod.classList.toggle('hidden', !moraYes);
    const months = moraYes ? (moratoriumPeriod.getValue() || 0) : 0;
    moraSection.classList.toggle('hidden', months === 0);
    if (months > 0) idpField.refresh();
    paymentModality.setLabel(moraYes ? 'Payment Modality after Moratorium' : 'Payment Modality');
    tenorMonths.setLabel(moraYes ? 'Loan Tenor including Moratorium at Disbursement (Months)' : 'Loan Tenor at Disbursement (Months)');
    if (split) {
      // Interest can never be settled less often than principal.
      const ok = FREQ[rrIntFreq.getValue()] <= FREQ[rrPrinFreq.getValue()];
      rrPrinFreq.classList.toggle('invalid', !ok);
      rrCustomWrap.classList.toggle('hidden', rrPrinBasis.getValue() !== 'Different per Date');
      rrRebuildCustom();
      rrPrinHint.textContent = principalScheduleHint(rrPrincipalMonths(), FREQ[rrPrinFreq.getValue()] || 1);
      rrSplitGrid.refresh();
    }
  }
  refresh();

  section.appendChild(el('div', { class: 'action-bar' },
    resetButton('revisionStructured', () => renderRateRevisionStructured(root), () => ({
      ...collectRevisionStructuredInputs({
        initialAmount, disbursementDate, moratoriumAvail, moratoriumPeriod, idpField,
        paymentModality, tenorMonths, rateLayers, securityLayers, referenceField,
        rrIntFreq, rrPrinFreq, rrPrinStart, rrPrinBasis, rrSplitGrid,
        getRrCustom: () => rrCustomBoxes.map(b => b.getValue() || 0),
      }),
      cofRows: (cofUpload.getRows() || []).length,
    }))));
  const { left, right: resultsPanel } = workbench(root);
  left.appendChild(section);

  restoreDraft('revisionStructured', {
    initialAmount, disbursementDate, moratoriumAvail, moratoriumPeriod, idpField,
    paymentModality, tenorMonths, rateLayers, securityLayers, referenceField,
    rrIntFreq, rrPrinFreq, rrPrinStart, rrPrinBasis, rrSplitGrid,
    getRrCustom: () => rrCustomBoxes.map(b => b.getValue() || 0),
  });
  // After the draft, never before: the entry answers are the fresher truth.
  applyEntryAnswers([
    [moratoriumAvail, pre && pre.moratorium],
    [paymentModality, pre && pre.modality],
  ]);
  refresh();
  setTimeout(() => { rateLayers.applyLayerRules(); securityLayers.applyLayerRules(); }, 150);
  attachDraftAutosave('revisionStructured', section, () => collectRevisionStructuredInputs({
    initialAmount, disbursementDate, moratoriumAvail, moratoriumPeriod, idpField,
    paymentModality, tenorMonths, rateLayers, securityLayers, referenceField,
    rrIntFreq, rrPrinFreq, rrPrinStart, rrPrinBasis, rrSplitGrid,
    getRrCustom: () => rrCustomBoxes.map(b => b.getValue() || 0),
  }));

  function compute() {
    // Every guard below used to be a toast fired by the Calculate button. Live, they become
    // the line shown in the results pane instead.
    const stop = (msg) => ({ msg });
    const inputs = collectRevisionStructuredInputs({
      initialAmount, disbursementDate, moratoriumAvail, moratoriumPeriod, idpField,
      paymentModality, tenorMonths, rateLayers, securityLayers, referenceField,
      rrIntFreq, rrPrinFreq, rrPrinStart, rrPrinBasis, rrSplitGrid,
      getRrCustom: () => rrCustomBoxes.map(b => b.getValue() || 0),
    });
    if (!inputs.initialAmount) return stop('Enter Initial Loan Amount.');
    if (!inputs.disbursementDate) return stop('Enter Disbursement Date.');
    const dow = new Date(inputs.disbursementDate).getDay();
    if (dow === 5 || dow === 6) return stop('Disbursement Date cannot be Friday or Saturday.');
    if (!inputs.tenorMonths) return stop('Enter Loan Tenor.');
    if (!inputs.paymentModality) return stop('Select a Payment Modality.');
    const rrSplit = inputs.paymentModality === SPLIT_MODE;
    if (!rrSplit && !inputs.moratoriumAvail) return stop('Select whether a moratorium is given at disbursement.');
    if (rrSplit) {
      if (FREQ[inputs.rrIntFreq] > FREQ[inputs.rrPrinFreq])
        return stop(`Interest cannot be paid less often than principal — ${inputs.rrIntFreq} interest with ${inputs.rrPrinFreq} principal leaves a principal month with no interest settlement.`);
      if (inputs.rrPrinStart < 1 || inputs.rrPrinStart > inputs.tenorMonths)
        return stop(`Principal Payments Start From Month must be between 1 and ${inputs.tenorMonths}.`);
      if (inputs.rrPrinBasis === 'Different per Date') {
        const dts = principalPaymentMonths(inputs.rrPrinStart, inputs.tenorMonths, FREQ[inputs.rrPrinFreq]);
        const earlier = inputs.rrCustomPrincipals.slice(0, Math.max(0, dts.length - 1));
        if (earlier.reduce((a, v) => a + (v || 0), 0) > inputs.initialAmount)
          return stop('The principal amounts entered exceed the Initial Loan Amount.');
        if (dts.length > 1 && !earlier.some(v => v))
          return stop('Enter the principal amount for at least one payment date, or switch Principal Amount to "Fixed (Equal)".');
      }
    }
    if (!inputs.rateLayers.length) return stop('Add at least one Lending Rate Layer.');

    const mat = maturityISO();
    const lastRateFrom = inputs.rateLayers[inputs.rateLayers.length - 1].fromDate;
    if (lastRateFrom && mat && lastRateFrom >= mat) return stop(`The last Lending Rate Layer's From Date (${lastRateFrom}) must be earlier than loan maturity (${mat}).`);
    if (inputs.securityLayers.length) {
      const lastSecFrom = inputs.securityLayers[inputs.securityLayers.length - 1].fromDate;
      if (lastSecFrom && mat && lastSecFrom >= mat) return stop(`The last Loan Security Layer's From Date (${lastSecFrom}) must be earlier than loan maturity (${mat}).`);
    }

    // Build COF effective-date data from the uploaded file (cut at maturity). COF must cover from
    // the disbursement date — otherwise ERR is not calculated and a blocking popup is shown.
    const { cofData } = buildCofData(cofUpload.getRows(), inputs.disbursementDate, mat);
    if (!cofCoversDisbursement(cofData, inputs.disbursementDate)) return stop(COF_GAP_MSG);

    const mora = (!rrSplit && inputs.moratoriumAvail === 'Yes') ? inputs.moratoriumPeriod : 0;
    const params = {
      initialLoanAmount: inputs.initialAmount,
      disbursementDate: inputs.disbursementDate,
      moratoriumMonths: mora,
      idpFlags: inputs.idpFlags,
      capFlags: inputs.capFlags,
      paymentModality: inputs.paymentModality,
      tenorMonths: inputs.tenorMonths,
      rateLayers: inputs.rateLayers,
      securityLayers: inputs.securityLayers,
      cofData,
      maturityDate: mat,
    };
    if (rrSplit) {
      Object.assign(params, {
        interestPeriod: FREQ[inputs.rrIntFreq],
        principalPeriod: FREQ[inputs.rrPrinFreq],
        principalStartMonth: inputs.rrPrinStart,
        principalBasis: inputs.rrPrinBasis === 'Different per Date' ? 'custom' : 'fixed',
        customPrincipals: inputs.rrCustomPrincipals,
        intFlags: inputs.rrSplitFlags,
      });
    }
    const schedule = buildRateRevisionStructured(params);
    const metrics = computeRevisionMetrics(schedule);
    const ctx = { pageType: 'revisionStructured', pageTitle: 'Rate Revision — Structured',
      inputs: { ...inputs, cofRecordCount: cofData.length }, params, schedule, metrics };
    return { ctx };
  }
  attachLiveRecalc(section, resultsPanel, compute);
}

function collectRevisionStructuredInputs(f) {
  return {
    initialAmount: f.initialAmount.getValue(),
    disbursementDate: f.disbursementDate.getValue(),
    moratoriumAvail: f.moratoriumAvail.getValue(),
    moratoriumPeriod: f.moratoriumPeriod.getValue() || 0,
    idpFlags: f.idpField.getPaidFlags(),
    capFlags: f.idpField.getCapFlags(),
    paymentModality: f.paymentModality.getValue(),
    tenorMonths: f.tenorMonths.getValue(),
    rateLayers: f.rateLayers.getValue().filter(r => r.fromDate && r.activeRate !== null),
    securityLayers: f.securityLayers.getValue().filter(r => r.fromDate && r.amount),
    reference: f.referenceField ? f.referenceField.getValue() : '',
    // Split interest/principal modality
    rrIntFreq: f.rrIntFreq ? f.rrIntFreq.getValue() : null,
    rrPrinFreq: f.rrPrinFreq ? f.rrPrinFreq.getValue() : null,
    rrPrinStart: f.rrPrinStart ? (f.rrPrinStart.getValue() || 1) : 1,
    rrPrinBasis: f.rrPrinBasis ? f.rrPrinBasis.getValue() : null,
    rrSplitStates: f.rrSplitGrid ? f.rrSplitGrid.getValue() : [],
    rrSplitFlags: f.rrSplitGrid ? f.rrSplitGrid.getIntFlags() : [],
    rrCustomPrincipals: f.getRrCustom ? f.getRrCustom() : [],
  };
}

// ============================================================
// RATE REVISION — CUSTOMIZED
// ============================================================
export function renderRateRevisionCustomized(root) {
  root.innerHTML = '';
  const section = el('div', { class: 'section-card' });

  const fileInput = el('input', { type: 'file', accept: '.xlsx,.xls', style: 'display:none' });
  const uploadBtn = el('button', { class: 'secondary-btn', type: 'button' }, '⬆ Upload Excel');
  uploadBtn.addEventListener('click', () => fileInput.click());
  const uploadedLabel = el('span', { class: 'help' }, 'No file uploaded');
  const sampleLink = el('a', { class: 'link-btn', href: '#', role: 'button' }, 'Download Sample File');
  sampleLink.addEventListener('click', (e) => { e.preventDefault(); downloadCustomizedRevisionSample(); });

  let uploadedRows = null;   // amortization schedule rows (Schedule sheet)
  let uploadedCof = null;    // COF records (COF Layers sheet of the same file)
  fileInput.addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      const parsed = await readCustomizedRevisionFile(f);
      uploadedRows = parsed.scheduleRows;
      uploadedCof = parsed.cofRows;
      uploadedLabel.textContent = `${f.name} — ${uploadedRows.length} schedule rows, ${uploadedCof.length} COF record(s)`;
      toast('File parsed: amortization schedule and COF layers loaded.', 'success');
    } catch (err) { uploadedRows = null; uploadedCof = null; toast(err.message, 'error'); }
  });

  // Distinguished upload-zone panel (dashed border, icon badge, title + hint).
  const uploadZone = el('div', { class: 'upload-zone' },
    el('div', { class: 'uz-head' },
      el('span', { class: 'uz-icon' }, '⬆'),
      el('div', { class: 'uz-titles' },
        el('div', { class: 'uz-title' }, 'Upload Amortization Schedule and COF Layers'),
        el('div', { class: 'uz-sub' }, 'One Excel file (.xlsx) with the Schedule and COF Layers sheets — use the sample as the template'))),
    el('div', { class: 'uz-actions' }, uploadBtn, fileInput, uploadedLabel),
    el('div', { class: 'uz-sample' }, sampleLink));
  section.appendChild(uploadZone);

  // From-only security layers (no To Date) — each applies from its From until the next layer's From.
  // Date defaults/bounds come from the uploaded schedule's span (first row = start, last = maturity).
  const securityLayers = securityLayersField({
    label: 'Loan Security Layers',
    name: 'securityLayers',
    getAnchor: () => ({ value: (uploadedRows && uploadedRows[0] && uploadedRows[0].date) || null }),
    getMaturity: () => (uploadedRows && uploadedRows.length ? uploadedRows[uploadedRows.length - 1].date : null),
  });

  section.appendChild(el('div', { class: 'layer-panel' }, securityLayers));

  // Optional free-text reference — prefixed onto every downloaded file's name for this calc.
  const referenceField = textField({
    label: 'Add Reference (Optional)', name: 'reference', placeholder: 'e.g. loan account / proposal no.',
  });
  section.appendChild(el('div', { class: 'form-row' }, referenceField));

  section.appendChild(el('div', { class: 'action-bar' },
    resetButton('revisionCustomized', () => renderRateRevisionCustomized(root), () => ({
      uploadedRows: uploadedRows ? uploadedRows.length : 0,
      uploadedCof: uploadedCof ? uploadedCof.length : 0,
      securityLayers: securityLayers.getValue(),
    }))));
  const { left, right: resultsPanel } = workbench(root);
  left.appendChild(section);

  function compute() {
    if (!uploadedRows) return { msg: 'Upload the amortization schedule + COF layers file to see the results.' };
    // COF effective-date list from the uploaded file's COF Layers sheet, clipped to the schedule's
    // span (first row = disbursement, last row = maturity), mirroring Rate Revision — Structured.
    const firstDate = uploadedRows[0] && uploadedRows[0].date;
    const lastDate = uploadedRows[uploadedRows.length - 1] && uploadedRows[uploadedRows.length - 1].date;
    const { cofData } = buildCofData(uploadedCof, firstDate, lastDate);
    if (!cofCoversDisbursement(cofData, firstDate)) return { msg: COF_GAP_MSG };
    const inputs = {
      securityLayers: securityLayers.getValue().filter(r => r.fromDate),
      cofRecords: (cofData || []).length,
      reference: referenceField.getValue(),
    };
    const metrics = computeRevisionCustomizedMetrics(uploadedRows, {
      securityLayers: inputs.securityLayers,
      cofData,
    });
    const schedule = {
      rows: uploadedRows.map((r, i) => ({
        sl: i, date: r.date,
        installment: r.installmentAmount,
        interest: r.interestAmount,
        principal: r.principalAmount,
        urpa: r.urpa,
        interestExpense: 0,
      })),
    };
    const ctx = {
      pageType: 'revisionCustomized', pageTitle: 'Rate Revision — Customized',
      inputs: { ...inputs, uploadedRowsCount: uploadedRows.length },
      // cofData + security layers ride along so the Verify Excel can compute the
      // per-row NIM/ERR (yield to maturity) columns with the same day-count method.
      params: { cofData, securityLayers: inputs.securityLayers }, schedule, metrics,
    };
    return { ctx };
  }
  attachLiveRecalc(section, resultsPanel, compute);
}

// ============================================================
// Shared helpers
// ============================================================
// Answers already given on the entry screen: fill the field, then drop it out of the form.
// An inline display:none survives refresh()'s .hidden toggling, so each form's own
// conditional logic keeps running untouched underneath.
export function applyEntryAnswers(pairs) {
  const rows = new Set();
  pairs.forEach(([field, value]) => {
    if (!field || value === undefined || value === null || value === '') return;
    field.setValue(value);
    field.style.display = 'none';
    const row = field.closest('.form-row');
    if (row) rows.add(row);
  });
  // A row whose every field has gone would otherwise leave a gap behind.
  rows.forEach((row) => {
    const live = [...row.querySelectorAll('.field')].some(f => f.style.display !== 'none');
    if (!live) row.style.display = 'none';
  });
}

// Inputs left, results right on a wide screen; stacked on a narrow one. The results pane
// is never empty-and-silent: until the inputs are usable it says what is still needed.
function workbench(root) {
  const left = el('div', { class: 'wb-input' });
  const right = el('div', { class: 'wb-output' });
  root.appendChild(el('div', { class: 'workbench' }, left, right));
  return { left, right };
}

function renderPending(panel, msg) {
  panel.innerHTML = '';
  panel.appendChild(el('div', { class: 'section-card pending-card' },
    el('div', { class: 'pending-title' }, 'Results'),
    el('p', { class: 'pending-msg' },
      msg || 'Fill in the details and the schedule will build itself here.')));
}

// Recalculation runs on every edit, so it must be cheap to call and must never interrupt
// with a toast or a modal — `compute` returns a ctx, or null having recorded why not.
function attachLiveRecalc(sectionEl, panel, compute) {
  let timer = null;
  const run = () => {
    if (!document.body.contains(sectionEl)) return;
    let ctx = null, msg = null;
    try { const r = compute(); ctx = r && r.ctx; msg = r && r.msg; }
    catch (e) { msg = 'Could not calculate with these inputs yet.'; }
    if (!ctx) return renderPending(panel, msg);
    autoSaveSummary(ctx);
    renderResults(panel, ctx);
  };
  const schedule = () => { if (timer) clearTimeout(timer); timer = setTimeout(() => { timer = null; run(); }, 140); };
  sectionEl.addEventListener('input', schedule);
  sectionEl.addEventListener('change', schedule);
  run();
  return run;
}

function pageTitle(text) {
  return el('div', { class: 'page-title' }, el('h1', {}, text));
}

// Header shown before a downloaded schedule (Excel/Word/PDF): only a few key facts.
// Loan Amount (money, 2dp), Interest Rate (%, 2dp), Tenor, Moratorium Period (if moratorium),
// and Payment Modality (labelled "… After Moratorium" when a moratorium exists).
function buildScheduleMeta(ctx) {
  const i = ctx.inputs || {};
  const moraYes = i.moratoriumAvail === 'Yes' || (Number(i.moratoriumPeriod) || 0) > 0;
  const header = {};
  const loanAmt = i.loanAmount != null ? i.loanAmount : i.initialAmount;
  if (loanAmt != null) header['Loan Amount'] = formatMoney(loanAmt);
  if (i.rateLayered) header['Interest Rate'] = formatRateLayers(i.intRateLayers);
  else if (i.offeredRate != null) header['Interest Rate'] = formatPercent(i.offeredRate);
  const tenor = i.loanTenor != null ? i.loanTenor : i.tenorMonths;
  if (tenor != null) header['Tenor (Months)'] = String(tenor);
  if (moraYes && i.moratoriumPeriod) header['Moratorium Period (Months)'] = String(i.moratoriumPeriod);
  const modality = i.paymentMode || i.paymentModality;
  if (modality) header[moraYes ? 'Payment Modality After Moratorium' : 'Payment Modality'] = modality;
  // A layered rate gets a per-row Interest Rate column. Never COF: these files go to clients.
  return { header, showRate: !!i.rateLayered };
}

// Shared "COF Data Upload" widget (Rate Revision — Structured & Customized).
// Upload button sits on the left; a blue "Download Sample File" link sits below it.
// Returns { field, getRows } where getRows() yields the parsed COF rows (or null).
function cofUploadField(onParsed) {
  let rows = null;
  const fileInput = el('input', { type: 'file', accept: '.xlsx,.xls', style: 'display:none' });
  const uploadBtn = el('button', { class: 'secondary-btn', type: 'button' }, '⬆ Upload COF Data');
  uploadBtn.addEventListener('click', () => fileInput.click());
  const status = el('span', { class: 'help' }, 'No COF file uploaded — interest expense will be 0.');
  const sampleLink = el('a', { class: 'link-btn', href: '#', role: 'button' }, 'Download Sample File');
  sampleLink.addEventListener('click', (e) => { e.preventDefault(); downloadCofSample(); });
  fileInput.addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      rows = await readUploadedCof(f);
      status.textContent = `${f.name} — ${rows.length} COF record(s) loaded.`;
      toast('COF data parsed successfully.', 'success');
      if (onParsed) onParsed(rows);
    } catch (err) { rows = null; status.textContent = 'Upload failed: ' + err.message; toast(err.message, 'error'); }
  });
  // Distinguished upload-zone panel (dashed border, icon badge, title + hint).
  const title = el('div', { class: 'uz-title' }, 'COF Data Upload');
  const field = el('div', { class: 'upload-zone' },
    el('div', { class: 'uz-head' }, el('div', { class: 'uz-titles' }, title)),
    el('div', { class: 'uz-actions' }, uploadBtn, fileInput, status),
    el('div', { class: 'uz-sample' }, sampleLink));
  return { field, getRows: () => rows };
}

function autoSaveSummary(ctx) {
  return; // Compare feature hidden for now — delete this line to restore auto-save + toasts.
  const result = saveSummary({
    pageType: ctx.pageType, pageTitle: ctx.pageTitle,
    inputs: JSON.parse(JSON.stringify(ctx.inputs)),
    metrics: {
      effectiveRate: ctx.metrics.effectiveRate, nim: ctx.metrics.nim, nii: ctx.metrics.nii,
      avgPortfolio: ctx.metrics.avgPortfolio, totalInterest: ctx.metrics.totalInterest, tenorYears: ctx.metrics.tenorYears,
    },
    label: ctx.pageTitle + ' @ ' + new Date().toLocaleString(),
  });
  if (result.saved) {
    toast(`Saved (${listSummaries().length}/${getMax()}).`, 'success');
    window.dispatchEvent(new CustomEvent('summary-saved'));
  } else if (result.reason === 'full') {
    toast(`Storage is full (${getMax()} max). Remove a saved summary to add this one.`, 'warn');
  }
  // duplicate: silently no-op
}

function renderResults(panel, ctx) {
  panel.innerHTML = '';
  const card = el('div', { class: 'section-card' });
  card.appendChild(el('h2', {}, 'Results'));

  const m = ctx.metrics;
  const grid = el('div', { class: 'results-grid' });
  const metric = (label, value, primary = false) => el('div', { class: 'metric-card' + (primary ? ' primary' : '') },
    el('div', { class: 'label' }, label),
    el('div', { class: 'value' }, value));
  grid.appendChild(metric('Effective Rate (ERR)', formatPercent(m.effectiveRate), true));
  grid.appendChild(metric('NIM', formatPercent(m.nim)));
  grid.appendChild(metric('Net Interest Income', formatMoney(m.nii)));
  // With rate layers some months may run on the 1% Refinance COF, so the COF half of
  // ERR = COF + NIM is no longer the typed Total COF — show the figure actually used.
  if (ctx.params && ctx.params.rateLayers) grid.appendChild(metric('Effective COF', formatPercent(m.effectiveCof)));
  // If model derived a security amount (EMI/EQI after Moratorium / Installment), show it for transparency
  const sk = String(ctx.inputs.fundedSecurityType || '');
  if ((sk.startsWith('EMI') || sk.startsWith('EQI') || sk === 'Installment') && m.derivedSecurityAmount > 0) {
    grid.appendChild(metric(
      `Security Amount (${sk} × ${ctx.inputs.numInst || 1})`,
      formatMoney(m.derivedSecurityAmount)
    ));
  }
  card.appendChild(grid);

  // Top-right actions: Download Report (left), Verify Calculation (right).
  // Optional user reference is prefixed to every file name (sanitized of OS-illegal chars).
  const refRaw = (ctx.inputs && ctx.inputs.reference) ? String(ctx.inputs.reference) : '';
  const refPrefix = refRaw.replace(/[\\/:*?"<>|]+/g, '').trim();
  const baseFname = (refPrefix ? refPrefix + '_' : '') +
    ctx.pageTitle.replace(/[^a-z0-9]+/gi, '_') + '_' + new Date().toISOString().slice(0, 10);
  const reportBtn = el('button', { class: 'secondary-btn', type: 'button' }, '📄 Download Report');
  const verifyBtn = el('button', { class: 'verify-btn', type: 'button' }, 'Verify Calculation');
  reportBtn.addEventListener('click', () => downloadReportPDF(baseFname + '_Report.pdf', ctx));
  verifyBtn.addEventListener('click', () => downloadVerificationExcel(baseFname + '_Verification.xlsx', ctx));
  card.appendChild(el('div', { class: 'results-actions' }, reportBtn, verifyBtn));

  panel.appendChild(card);

  // Amortization schedule + totals + downloads
  const tableCard = el('div', { class: 'section-card' });
  tableCard.appendChild(el('h2', {}, 'Amortization Schedule'));
  const wrap = el('div', { class: 'table-wrap' });
  const table = el('table', { class: 'schedule' });
  const hasDate = ctx.schedule.rows[0]?.date !== undefined;
  const hasIDP = ctx.schedule.rows.some(r => (r.idpReceivable || 0) > 0);
  // Rate-layered Loan Facilities: each month's own lending rate and COF.
  const hasRates = !!(ctx.params && ctx.params.rateLayers);
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  if (hasIDP) headers.push('Accrued Interest');
  if (hasRates) headers.push('Rate', 'COF');
  const thead = el('thead', {}, el('tr', {}, ...headers.map(h => el('th', {}, h))));
  const tbody = el('tbody');
  let totPay = 0, totInt = 0, totPrin = 0;
  ctx.schedule.rows.forEach((r) => {
    totPay += r.installment || 0;
    totInt += r.interest || 0;
    totPrin += r.principal || 0;
    const cells = [String(r.sl)];
    if (hasDate) cells.push(r.date || '');
    cells.push(formatMoney(r.installment), formatMoney(r.interest), formatMoney(r.principal), formatMoney(r.urpa));
    if (hasIDP) cells.push(formatMoney(r.idpReceivable || 0));
    if (hasRates) cells.push(r.sl > 0 ? formatPercent(r.rate) : '', r.sl > 0 ? formatPercent(r.cof) : '');
    tbody.appendChild(el('tr', {}, ...cells.map(c => el('td', {}, c))));
  });
  const totalCells = ['TOTAL'];
  if (hasDate) totalCells.push('');
  totalCells.push(formatMoney(totPay), formatMoney(totInt), formatMoney(totPrin), '');
  if (hasIDP) totalCells.push('');
  if (hasRates) totalCells.push('', '');
  tbody.appendChild(el('tr', { class: 'totals-row' }, ...totalCells.map(c => el('td', {}, c))));
  table.appendChild(thead);
  table.appendChild(tbody);
  wrap.appendChild(table);
  tableCard.appendChild(wrap);

  const summaryRow = el('div', { class: 'totals-summary' },
    el('div', { class: 'totals-card' },
      el('div', { class: 'label' }, 'Total Principal Paid'),
      el('div', { class: 'value' }, formatMoney(totPrin))),
    el('div', { class: 'totals-card' },
      el('div', { class: 'label' }, 'Total Interest Paid'),
      el('div', { class: 'value' }, formatMoney(totInt))),
    el('div', { class: 'totals-card' },
      el('div', { class: 'label' }, 'Total Payment'),
      el('div', { class: 'value' }, formatMoney(totPay))),
  );
  tableCard.appendChild(summaryRow);

  const meta = buildScheduleMeta(ctx);
  const dlExcel = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ Excel');
  const dlWord = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ Word');
  const dlPdf = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ PDF');
  const dlBar = el('div', { class: 'download-box' },
    el('div', { class: 'dl-title' }, 'Download the Schedule'),
    el('div', { class: 'download-buttons' }, dlExcel, dlWord, dlPdf),
  );
  tableCard.appendChild(dlBar);
  dlExcel.addEventListener('click', () => downloadScheduleAsExcel(baseFname + '.xlsx', ctx.schedule, meta));
  dlWord.addEventListener('click', () => downloadScheduleAsWord(baseFname + '.docx', ctx.schedule, meta));
  dlPdf.addEventListener('click', () => downloadScheduleAsPDF(baseFname + '.pdf', ctx.schedule, meta));

  panel.appendChild(tableCard);
}

// ============================================================
// Draft preservation
// ============================================================
function attachDraftAutosave(tabKey, sectionEl, collector) {
  let last = '';
  let timer = null;
  function doSave() {
    // Section replaced (e.g. after Reset): never save from a detached form.
    if (!document.body.contains(sectionEl)) return;
    try {
      const data = collector();
      const ser = JSON.stringify(data);
      if (ser !== last) {
        last = ser;
        saveDraft(tabKey, data);
      }
    } catch {}
  }
  function schedule() {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => { timer = null; doSave(); }, 300);
  }
  function flush() {
    if (timer) { clearTimeout(timer); timer = null; }
    doSave();
  }
  // Debounce keystrokes; force-save on navigation
  sectionEl.addEventListener('input', schedule);
  sectionEl.addEventListener('change', schedule);
  window.addEventListener('beforeunload', flush);
  setTimeout(doSave, 200);
}

function restoreDraft(tabKey, fields) {
  const data = loadDraft(tabKey);
  if (!data) return;
  // Apply each known field
  try {
    if (fields.loanAmount && data.loanAmount !== undefined) fields.loanAmount.setValue(data.loanAmount);
    if (fields.offeredRate && data.offeredRate !== undefined) fields.offeredRate.setValue(data.offeredRate);
    if (fields.initialAmount && data.initialAmount !== undefined) fields.initialAmount.setValue(data.initialAmount);
    if (fields.disbursementDate && data.disbursementDate) fields.disbursementDate.setValue(data.disbursementDate);
    if (fields.moratoriumAvail && data.moratoriumAvail) fields.moratoriumAvail.setValue(data.moratoriumAvail);
    if (fields.moratoriumPeriod && data.moratoriumPeriod) fields.moratoriumPeriod.setValue(data.moratoriumPeriod);
    if (fields.idpField && (Array.isArray(data.idpFlags) || Array.isArray(data.capFlags))) {
      const idp = data.idpFlags || [], cap = data.capFlags || [];
      const n = Math.max(idp.length, cap.length);
      fields.idpField.setValue(Array.from({ length: n }, (_, i) => cap[i] ? 2 : (idp[i] ? 1 : 0)));
    }
    if (fields.loanTenor && data.loanTenor) fields.loanTenor.setValue(data.loanTenor);
    if (fields.tenorMonths && data.tenorMonths) fields.tenorMonths.setValue(data.tenorMonths);
    if (fields.paymentMode && data.paymentMode) fields.paymentMode.setValue(data.paymentMode);
    // Split interest/principal type
    if (fields.intFreq && data.intFreq) fields.intFreq.setValue(data.intFreq);
    if (fields.prinFreq && data.prinFreq) fields.prinFreq.setValue(data.prinFreq);
    if (fields.prinStart && data.prinStart) fields.prinStart.setValue(data.prinStart);
    if (fields.prinBasis && data.prinBasis) fields.prinBasis.setValue(data.prinBasis);
    if (fields.splitGrid && Array.isArray(data.splitStates)) fields.splitGrid.setValue(data.splitStates);
    if (fields.custSplitGrid && Array.isArray(data.custSplitStates)) fields.custSplitGrid.setValue(data.custSplitStates);
    if (fields.rrIntFreq && data.rrIntFreq) fields.rrIntFreq.setValue(data.rrIntFreq);
    if (fields.rrPrinFreq && data.rrPrinFreq) fields.rrPrinFreq.setValue(data.rrPrinFreq);
    if (fields.rrPrinStart && data.rrPrinStart) fields.rrPrinStart.setValue(data.rrPrinStart);
    if (fields.rrPrinBasis && data.rrPrinBasis) fields.rrPrinBasis.setValue(data.rrPrinBasis);
    if (fields.rrSplitGrid && Array.isArray(data.rrSplitStates)) fields.rrSplitGrid.setValue(data.rrSplitStates);
    if (fields.paymentModality && data.paymentModality) fields.paymentModality.setValue(data.paymentModality);
    if (fields.totalCof && data.totalCof !== undefined) fields.totalCof.setValue(data.totalCof);
    if (fields.fundedSecurityType && data.fundedSecurityType) fields.fundedSecurityType.setValue(data.fundedSecurityType);
    if (fields.csAmount && data.csAmount !== undefined) fields.csAmount.setValue(data.csAmount);
    if (fields.csRate && data.csRate !== undefined) fields.csRate.setValue(data.csRate);
    if (fields.numInst && data.numInst) fields.numInst.setValue(data.numInst);
    if (fields.nimComparison && data.nimComparison) fields.nimComparison.setValue(data.nimComparison);
    if (fields.paymentLayers && Array.isArray(data.paymentLayers) && data.paymentLayers.length) {
      // Convert ISO->display etc. handled by layered field's number/percent setters internally
      fields.paymentLayers.setValue(data.paymentLayers.map(L => ({
        fromInstallment: L.fromInstallment != null ? String(L.fromInstallment) : '',
        toInstallment: L.toInstallment != null ? String(L.toInstallment) : '',
        paymentType: L.paymentType,
        customPrincipal: L.customPrincipal,
      })));
    }
    if (fields.intRateLayers && Array.isArray(data.intRateLayers)) fields.intRateLayers.setValue(data.intRateLayers);
    if (fields.rateLayers && Array.isArray(data.rateLayers) && data.rateLayers.length) {
      fields.rateLayers.setValue(data.rateLayers.map(L => ({
        fromDate: isoToDDMMMYYYY(L.fromDate),
        activeRate: L.activeRate,
      })));
    }
    if (fields.securityLayers && Array.isArray(data.securityLayers) && data.securityLayers.length) {
      fields.securityLayers.setValue(data.securityLayers.map(L => ({
        fromDate: isoToDDMMMYYYY(L.fromDate),
        amount: L.amount,
        activeRate: L.activeRate,
        securities: L.securities,
      })));
    }
    if (fields.cofLayers && Array.isArray(data.cofLayers) && data.cofLayers.length) {
      fields.cofLayers.setValue(data.cofLayers.map(L => ({
        fromDate: isoToDDMMMYYYY(L.fromDate),
        toDate: isoToDDMMMYYYY(L.toDate),
        cofRate: L.cofRate,
      })));
    }
    if (fields.referenceField && data.reference !== undefined) fields.referenceField.setValue(data.reference);
  } catch (err) {
    console.warn('Draft restore failed', err);
  }
}
