// ============================================================
// Entry screens
//
// The app used to open on four tabs whose names ("Structured" / "Customized") describe how
// the code is organised, not anything an RM knows about their deal. These screens replace
// that with two modules and a handful of questions the RM can actually answer; the answers
// decide which form they land on, and are carried into it so nothing is asked twice.
// ============================================================
import { el, optionField } from './components.js?v=20261109c';
import { SPLIT_MODE, CBS_PRODUCTS } from './calculations.js?v=20261109c';

// "Upload an existing schedule" is switched off until it is ready to offer (user, 01-Nov-2026).
// While it is off, Rate Revision goes straight to its questions; true brings the fork back.
export const SHOW_REVISION_UPLOAD = false;

export const MODALITIES = [
  'EMI', 'EQI',
  'Equal Principal + Interest (Monthly)',
  'Equal Principal + Interest (Quarterly)',
  SPLIT_MODE,
];

function choiceCard(title, blurb, onPick) {
  const card = el('button', { class: 'choice-card', type: 'button' },
    el('span', { class: 'cc-title' }, title),
    el('span', { class: 'cc-sub' }, blurb));
  card.addEventListener('click', onPick);
  return card;
}

function backLink(onClick) {
  const b = el('button', { class: 'back-link', type: 'button' }, '← Back');
  b.addEventListener('click', onClick);
  return b;
}

// ---------------- Module picker ----------------
export function renderLanding(root, go) {
  root.innerHTML = '';
  root.appendChild(el('div', { class: 'screen' },
    el('div', { class: 'choice-grid' },
      choiceCard('ERR - New Loan Facility',
        'Price a new facility and see what it actually earns.',
        () => go({ screen: 'questions', module: 'loan' })),
      choiceCard('ERR - Rate Revision',
        'Re-price a facility that is already running.',
        () => go(SHOW_REVISION_UPLOAD
          ? { screen: 'revisionChoice', module: 'revision' }
          : { screen: 'questions', module: 'revision', mode: 'structured' })),
    )));
}

// ---------------- Rate Revision fork ----------------
// Upload-driven revisions never need the moratorium/modality questions — the schedule
// already encodes them — so the fork comes before the questions, not after.
export function renderRevisionChoice(root, go) {
  root.innerHTML = '';
  root.appendChild(el('div', { class: 'screen' },
    backLink(() => go({ screen: 'landing' })),
    el('h1', { class: 'screen-title' }, 'ERR - Rate Revision'),
    el('div', { class: 'choice-grid' },
      choiceCard('Enter the loan details',
        'Rebuild the schedule from the facility’s original terms.',
        () => go({ screen: 'questions', module: 'revision', mode: 'structured' })),
      choiceCard('Upload an existing schedule',
        'You already have the amortisation schedule and COF layers in one Excel file.',
        // No questions precede this one, so it must carry no answers.
        () => go({ screen: 'form', module: 'revision', mode: 'upload', answers: null })),
    )));
}

// ---------------- Question card ----------------
// Order matters: "multiple layers?" is asked BEFORE modality, because a layered facility
// carries a modality per layer — so a single upfront answer would mean nothing. When the
// answer is Yes the modality question is skipped entirely.
export function renderQuestions(root, ctx, go) {
  root.innerHTML = '';
  const isLoan = ctx.module === 'loan';
  const a = ctx.answers || {};

  // Rate Revision only, asked first: the loan's product in the bank's core banking system,
  // which decides how each period's interest is counted.
  const product = optionField({
    label: 'Select the CBS Product for the Schedule', name: 'qProduct',
    options: [{ label: 'select', value: '' }, ...CBS_PRODUCTS.map(({ value, label }) => ({ value, label }))],
    value: a.cbsProduct || '',
    help: 'DAYS360 counts every month as 30 days; Actual/360 counts the days from start date to end date (End Date − Start Date). Both divide by 360.',
    onChange: () => refresh(),
  });

  const mora = optionField({
    label: 'Does the loan have a moratorium period?', name: 'qMora',
    options: [{ label: 'select', value: '' }, 'No', 'Yes'], value: a.moratorium || '',
    onChange: () => refresh(),
  });

  // Asked whether or not there is a moratorium: a rate can change inside the moratorium, after
  // it, or on a loan that has none (a Refinance Rate for part of the tenor, say).
  const rates = optionField({
    label: 'Are there Multiple Layers of Interest Rates?', name: 'qRates',
    options: [{ label: 'select', value: '' }, 'No', 'Yes'], value: a.rateLayers || '',
    help: 'Pick Yes if the interest rate changes at any point in the loan’s life, including a switch to a Refinance Rate.',
    onChange: () => refresh(),
  });

  const layers = optionField({
    label: 'Does the payment have multiple layers?', name: 'qLayers',
    options: [{ label: 'select', value: '' }, 'No', 'Yes'], value: a.multiLayers || '',
    help: 'Pick Yes if the modality changes at any point in the loan’s life.',
    onChange: () => refresh(),
  });
  const modality = optionField({
    label: 'Payment Modality', name: 'qModality',
    options: [{ label: 'select', value: '' }, ...MODALITIES], value: a.modality || '',
    onChange: () => refresh(),
  });

  const productRow = el('div', { class: 'form-row' }, product);
  const moraRow = el('div', { class: 'form-row' }, mora);
  const ratesRow = el('div', { class: 'form-row' }, rates);
  const layersRow = el('div', { class: 'form-row' }, layers);
  const modalityRow = el('div', { class: 'form-row' }, modality);

  const contBtn = el('button', { class: 'primary-btn', type: 'button' }, 'Continue');

  const card = el('div', { class: 'section-card' }, ...(isLoan ? [] : [productRow]), moraRow, ratesRow, layersRow, modalityRow,
    el('div', { class: 'action-bar' }, contBtn));

  root.appendChild(el('div', { class: 'screen' },
    backLink(() => go({ screen: isLoan || !SHOW_REVISION_UPLOAD ? 'landing' : 'revisionChoice', module: ctx.module })),
    el('h1', { class: 'screen-title' }, isLoan ? 'ERR - New Loan Facility' : 'ERR - Rate Revision'),
    card));

  // How many months is asked on the form itself; here we only need to know whether there
  // is a moratorium at all, since that is what the later labels and routing depend on.
  const productAnswered = () => isLoan || !!product.getValue();
  const moraAnswered = () => productAnswered() && !!mora.getValue();
  // Revision has no rates question (its Lending Rate Layers already carry every change).
  const ratesAnswered = () => !isLoan || !!rates.getValue();
  // Revision has no layers question, so it always needs a modality.
  const needsModality = () => !isLoan || layers.getValue() === 'No';

  function refresh() {
    const yes = mora.getValue() === 'Yes';

    // Each question appears only once the one before it is settled.
    moraRow.classList.toggle('hidden', !productAnswered());
    ratesRow.classList.toggle('hidden', !isLoan || !moraAnswered());
    layersRow.classList.toggle('hidden', !isLoan || !moraAnswered() || !ratesAnswered());
    layers.setLabel(yes
      ? 'Does the payment after the moratorium have multiple layers?'
      : 'Does the payment have multiple layers?');

    const showModality = moraAnswered() && ratesAnswered() && needsModality() && (!isLoan || !!layers.getValue());
    modalityRow.classList.toggle('hidden', !showModality);
    modality.setLabel(yes ? 'Payment Modality After Moratorium' : 'Payment Modality');

    const done = moraAnswered() && ratesAnswered()
      && (!isLoan || !!layers.getValue())
      && (!showModality || !!modality.getValue());
    contBtn.disabled = !done;
  }
  refresh();

  contBtn.addEventListener('click', () => {
    const answers = {
      moratorium: mora.getValue(),
      rateLayers: isLoan ? rates.getValue() : 'No',
      multiLayers: isLoan ? layers.getValue() : 'No',
      modality: needsModality() ? modality.getValue() : null,
      cbsProduct: isLoan ? null : product.getValue(),
    };
    go({
      screen: 'form',
      module: ctx.module,
      mode: isLoan ? (answers.multiLayers === 'Yes' ? 'customized' : 'structured') : 'structured',
      answers,
    });
  });
}

// A compact, clickable read-back of the answers, shown above the form so the RM can see
// what they said and change it without starting over.
export function answerSummary(ctx, onEdit) {
  const a = ctx.answers || {};
  const bits = [];
  const product = CBS_PRODUCTS.find(p => p.value === a.cbsProduct);
  if (ctx.module === 'revision' && product) bits.push(product.short);
  bits.push(a.moratorium === 'Yes' ? 'Moratorium' : 'No moratorium');
  if (ctx.module === 'loan') bits.push(a.rateLayers === 'Yes' ? 'Multiple interest rates' : 'Single interest rate');
  if (ctx.module === 'loan') bits.push(a.multiLayers === 'Yes' ? 'Multiple payment layers' : 'Single modality');
  if (a.modality) bits.push(a.modality);

  const row = el('div', { class: 'answer-bar' },
    ...bits.map(b => el('span', { class: 'answer-chip' }, b)));
  const edit = el('button', { class: 'link-btn', type: 'button' }, 'Change');
  edit.addEventListener('click', onEdit);
  row.appendChild(edit);
  return row;
}
