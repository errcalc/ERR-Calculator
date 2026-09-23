// Per-module user guide — opened from the info button in the header. Plain, simple language
// (explained as if to a 10-year-old). Sections are collapsed by default; click to expand.
import { el, openModal, closeModal } from './components.js?v=20260923g';

// ---- Shared sections reused across modules ----
const MODALITY_BULLETS = [
  'EMI — the same total amount every month.',
  'EQI — the same total amount every 3 months.',
  'Equal Principal + Interest (Monthly or Quarterly) — pays the same chunk of the original loan each time, plus interest on what is left.',
  'Interest & Principal (Separate Frequency) — interest and principal are paid on DIFFERENT months. For example, interest every month but principal only at the end of each quarter.',
];
const MORATORIUM_Q = {
  h: 'Does the loan have a moratorium period?',
  p: 'A moratorium is a “rest period” at the very start, when the borrower does not repay yet. Answer Yes or No.',
  bullets: [
    'How many months it lasts is asked on the next screen, with the rest of the numbers.',
    'Saying Yes also changes the wording of the questions that follow, so they talk about what happens AFTER the rest period.',
  ],
};
const NEXT_STEP_Q = {
  h: 'What happens after this',
  p: 'Answer everything and “Continue” takes you to the form, which only shows the boxes your answers made relevant.',
  bullets: [
    'Your answers appear as grey pills along the top of the form.',
    'Click “Change” beside them to come back here and answer differently — what you typed on the form is kept.',
  ],
};
const FLOW_SECTION = {
  h: 'How this page works',
  p: 'Each module has two steps, so you are never shown boxes that have nothing to do with your deal.',
  bullets: [
    'Step 1 — a few short questions: is there a moratorium, does the interest rate change during the loan, does the payment style change part-way through the loan, and if it does not, which style is it.',
    'Step 2 — the form: everything else. Only the boxes your answers made relevant appear.',
    'The grey pills along the top of the form show what you answered in step 1. Click “Change” next to them to go back and answer differently.',
  ],
};
const ANSWERED_SECTION = {
  h: 'Answers you gave in step 1',
  p: 'These were asked before you reached this form, which is why you will not find boxes for them here:',
  bullets: [
    'Whether the loan has a moratorium — a “rest period” at the start when the borrower does not repay yet. How many months it lasts is asked on this form.',
    'Whether the interest rate changes during the loan. No means one Offered Rate; Yes gives you the Interest Rate Layers table instead.',
    'Whether the payment style changes during the loan. No means one style throughout; Yes gives you the Payment Layers table instead.',
    'Which payment style it is — only asked when the style does NOT change.',
    'Click “Change” at the top of the form to revisit any of them.',
  ],
};

const IDP_SECTION = {
  h: 'Interest During Moratorium Period',
  p: 'Even during the rest period, interest keeps growing. For each month, click the box to tell the system what to do with that month’s interest. Click again to change it.',
  bullets: [
    'Light blue = Accrued: the interest waits and is collected at the next paying month.',
    'Indigo = Paid: the borrower pays that month’s interest right away.',
    'Purple = Capitalized: the interest is added on top of the loan, so later interest grows on it too.',
    'Use the three buttons on the right (“All to be …”) to set every month at once.',
  ],
};
const RATE_LAYERS_SECTION = {
  h: 'Interest Rate Layers',
  p: 'Only there when you said in step 1 that the interest rate changes. It takes the place of the Offered Rate box: one row per stretch of the loan, each with its own rate. The form only shows the saved layers — click “Edit” beside the heading to add or change them in a pop-up, then “Save” and confirm.',
  bullets: [
    'Enter the Loan Tenor first — the layers are counted in its months.',
    'Rows are counted in the loan’s own months. The first row always starts at Month 01, and each next row starts the month after the one before it ends — you only pick where each row ends (To Month). “+ Add another layer” works once the last row ends before the final month.',
    'The rate can change inside the moratorium too, because the months are the whole loan’s months, not the months after the rest period.',
    'The last row must end at the loan’s final month. Change the tenor afterwards and the last layer follows it by itself — a longer tenor stretches it, a shorter one ends the layers at the new final month. Lengthen the tenor again and the layers you saved come back.',
    'Rate Type — Commercial Rate: you type the rate. Refinance Rate: the rate is fixed at 5.00% and its box locks. Switch back to Commercial and whatever you had typed returns.',
    'In Refinance months the bank’s cost of fund is 1% flat instead of the Total Cost of Fund you typed. Commercial months keep the Total Cost of Fund. The Loan Security Benefit always uses the Total Cost of Fund.',
    'EMI and EQI are worked out again from the first payment that feels a new rate, so the loan still finishes at exactly zero. If a rate changes in the middle of a quarter, that quarter’s interest is split month by month.',
    'Because some months may cost only 1%, the results also show the Effective COF — the cost of fund actually paid across the whole loan. ERR = Effective COF + NIM. A Refinance stretch pulls ERR down, since it only earns 5%, while its cheap funding keeps the NIM healthy.',
  ],
};
const SECURITY_LAYERS_SECTION = {
  h: 'Loan Security Layers',
  p: 'This is money the borrower keeps parked with the bank as safety. Each row is the security balance in effect from its date, at its rate.',
  bullets: [
    'Type the full balance for that date (a total, not an addition). The most recent row applies until the next row’s date.',
    'The bank earns the gap between the cost of fund and the security’s rate — this is the Loan Security Benefit.',
    'If the balance changes on a day that is not a payment day, the system splits that month by the exact number of days, so the benefit is correct.',
  ],
};
const REFERENCE_SECTION = {
  h: 'Add Reference (optional)',
  p: 'Type any tag you like — a loan account number, a proposal number, anything. It is added to the front of every file you download (separated by an underscore), so the files are easy to find later. You can also leave it blank.',
};
const DOWNLOADS_SECTION = {
  h: 'Results & Downloads',
  p: 'There is no Calculate button. As soon as the form has enough to work with, the right-hand side fills in by itself and keeps up with every change you make — so you can nudge the rate and watch the return move. Until it has enough, it simply tells you what is still missing. You get ERR (the real yearly return), NIM (the bank’s interest margin) and Net Interest Income (the money earned), and below them you can download:',
  bullets: [
    'The full payment Schedule as Excel, Word, or PDF.',
    'A Report PDF that summarises the inputs and results.',
    'A Verify Calculation Excel that shows every single formula, so anyone can check the maths.',
  ],
};

const MANUALS = {
  // Shown while the user is on the landing screen or answering the step-1 questions.
  home: {
    title: 'Effective Rate of Return Calculator',
    intro: 'This tool works out what a loan really earns the bank — its Effective Rate of Return. Start by picking one of the two modules.',
    sections: [
      {
        h: 'Which module do I want?',
        p: 'Two choices, and the difference is simply whether the loan already exists:',
        bullets: [
          'ERR - New Loan Facility — you are pricing a NEW facility and want to know what it would earn.',
          'ERR - Rate Revision — the facility is already running and its interest rate has changed since it was given out.',
        ],
      },
      FLOW_SECTION,
    ],
  },

  // Step 1 of Loan Facilities: only the questions actually on screen.
  questionsLoan: {
    title: 'ERR - New Loan Facility — the questions',
    intro: 'A few short questions about the deal. They decide which boxes the next screen shows you, so you are never given fields that have nothing to do with your loan.',
    sections: [
      MORATORIUM_Q,
      {
        h: 'Are there Multiple Layers of Interest Rates?',
        p: 'This asks whether the interest rate stays the same for the whole loan, or changes at some point — with or without a moratorium.',
        bullets: [
          'No — one Offered Rate from start to finish.',
          'Yes — for example 12% for the first 9 months and 14% afterwards, or a Refinance Rate for part of the loan. The form then gives you an Interest Rate Layers table instead of the single Offered Rate box.',
        ],
      },
      {
        h: 'Does the payment have multiple layers?',
        p: 'This asks whether the borrower pays the same way for the whole loan, or differently at different points in its life.',
        bullets: [
          'No — one style from start to finish. You will be asked which one next.',
          'Yes — for example EMI for the first year and something else afterwards. You get a Payment Layers table on the form, one row per stretch of the loan, each with its own style. There is no single modality question, because each layer carries its own.',
        ],
      },
      {
        h: 'Payment Modality',
        p: 'Only appears when you answered No above — a layered loan sets its style per layer instead. The choices are:',
        bullets: MODALITY_BULLETS,
      },
      NEXT_STEP_Q,
    ],
  },

  // Step 1 of Rate Revision: no layers question here.
  questionsRevision: {
    title: 'ERR - Rate Revision — the questions',
    intro: 'Two short questions about the facility as it was originally given out. The rate changes themselves are listed on the next screen.',
    sections: [
      MORATORIUM_Q,
      {
        h: 'Payment Modality',
        p: 'How the borrower repays. Unlike ERR - New Loan Facility there is no layers question here, so this is always asked. The choices are:',
        bullets: MODALITY_BULLETS,
      },
      NEXT_STEP_Q,
    ],
  },

  // The fork shown before Rate Revision's questions.
  revisionChoice: {
    title: 'ERR - Rate Revision',
    intro: 'First, how do you want to give the system the loan? Pick whichever matches what you already have to hand.',
    sections: [
      {
        h: 'Enter the loan details',
        p: 'Choose this when you have the original terms — amount, disbursement date, tenor — and the list of rates with the dates they changed. The system rebuilds the whole payment schedule from them.',
      },
      {
        h: 'Upload an existing schedule',
        p: 'Choose this when the finished payment schedule already exists in a file, with the rate changes worked into it. You upload that file instead of answering anything.',
        bullets: [
          'The moratorium and modality questions are skipped, because a finished schedule already has them baked in.',
        ],
      },
    ],
  },

  regular: {
    title: 'ERR - New Loan Facility — one payment style',
    intro: 'You said the loan is paid back the same way from start to finish. Fill the boxes from the top down; the schedule and the return build themselves on the right as you go.',
    sections: [
      { h: 'Loan Amount', p: 'The total money the bank hands over to the borrower. Type the whole amount, for example 100,000,000.' },
      { h: 'Offered Rate', p: 'The yearly interest the bank charges, like 12%. Think of it as the price the borrower pays for using the money for one year. If you said in step 1 that the rate changes, this box is replaced by the Interest Rate Layers table further down.' },
      ANSWERED_SECTION,
      { h: 'Moratorium Period (Months)', p: 'How many months the rest period lasts. For example, 6 means the first 6 months are the rest period. It only appears when you said there IS a moratorium.' },
      IDP_SECTION,
      { h: 'Loan Tenor (Months)', p: 'The total life of the loan in months, counting the rest period too. For example, 60 means the loan lasts 5 years.' },
      RATE_LAYERS_SECTION,
      {
        h: 'Payment Mode — chosen in step 1',
        p: 'You picked this before reaching the form, so there is no box for it here; use “Change” at the top to swap it. The choices mean:',
        bullets: [
          'EMI — the same total amount every month.',
          'EQI — the same total amount every 3 months.',
          'Equal Principal + Interest (Monthly or Quarterly) — pays the same chunk of the original loan each time, plus interest on what is left.',
          'Interest & Principal (Separate Frequency) — interest and principal are paid on DIFFERENT months. For example, interest every month but principal only at the end of each quarter.',
        ],
      },
      {
        h: 'Interest & Principal (Separate Frequency) — the extra boxes',
        p: 'Pick this mode and four more boxes appear. They only show for this mode.',
        bullets: [
          'Interest Payment Frequency — how often interest is actually paid: Monthly, Quarterly, Half-yearly or Yearly.',
          'Principal Payment Frequency — how often a chunk of the loan is repaid. It can never be more often than interest, because interest is always taken first. So Monthly interest with Quarterly principal is fine, but Quarterly interest with Monthly principal is not.',
          'Principal Payments Start From Month — put 1 if principal starts right away. Put a later month to give the borrower a break from principal at the start while interest is still being paid.',
          'Principal Amount — “Fixed (Equal)” splits the loan evenly across the principal dates. “Different per Date” gives you one box per date so you can type each amount; the last box fills itself in with whatever is left, so the loan always finishes at zero.',
        ],
      },
      {
        h: 'Interest Treatment by Month',
        p: 'This mode has no separate rest-period boxes. Instead you get one row of month boxes covering the WHOLE loan, and you click a month to change what happens to its interest — the same three choices as the rest period. That is more flexible, because each month can be different.',
        bullets: [
          'Accrued — the interest is not paid this month; it waits and is collected at the next paying month.',
          'Paid — the interest is paid in cash this month, along with anything waiting from earlier months.',
          'Capitalized — the interest is added onto the loan instead of being paid, so the borrower then owes interest on it too.',
          'Months where principal is repaid are locked on Paid and show a small padlock. Interest is always settled where principal is settled.',
          'The two frequency boxes fill the row in for you. You can still click any unlocked month to change it, and “Reset to Frequency” puts it back.',
          'A rest period is simply the first few months set to Accrued or Capitalized, with principal starting later.',
        ],
      },
      { h: 'Total Cost of Fund (COF/ISC + OPEX)', p: 'What it costs the bank to get this money, plus its running costs, as a yearly rate. The bank’s real earning is the gap between the Offered Rate and this number. With Interest Rate Layers, Refinance Rate months use 1% instead.' },
      {
        h: 'Funded Security Type',
        p: 'Sometimes the borrower also keeps some money parked with the bank as safety. Pick the type, or “No Funded Security” if there is none.',
        bullets: [
          'FDR / Cash Security — one fixed safety amount.',
          'EMI / EQI after Moratorium, or Installment — the safety is built up bit by bit in installments.',
        ],
      },
      { h: 'The security’s Amount & Rate', p: 'These two boxes are named after whichever security you picked — choose FDR and they read “FDR Amount” and “FDR Rate”. Type how much safety money is kept and the rate the bank pays on it; the bank’s gain from it is added to the return.' },
      DOWNLOADS_SECTION,
    ],
  },

  customized: {
    title: 'ERR - New Loan Facility — payment layers',
    intro: 'You said the payment style changes during the loan, so instead of one modality you build “Payment Layers” — one row per stretch of the loan, each with its own style.',
    sections: [
      { h: 'Loan Amount', p: 'The total money the bank hands over to the borrower. Type the whole amount, for example 100,000,000.' },
      { h: 'Offered Rate', p: 'The yearly interest the bank charges, like 12% — the price of borrowing for one year. Replaced by the Interest Rate Layers table when you said in step 1 that the rate changes.' },
      ANSWERED_SECTION,
      { h: 'Moratorium Period (Months)', p: 'How many months the rest period lasts. It only appears when you said there IS a moratorium.' },
      IDP_SECTION,
      { h: 'Loan Tenor (Months)', p: 'The total life of the loan in months, including the rest period.' },
      RATE_LAYERS_SECTION,
      {
        h: 'Payment Layers',
        p: 'Instead of one payment style for the whole loan, you split the loan into time ranges and choose a style for each. For example: months 1–12 as EMI, then months 13–24 as EQI.',
        bullets: [
          'Each layer has a From month and a To month.',
          'Layers must line up neatly — no gaps and no overlaps. The system helps you keep them in order.',
          '“Customized Principal” lets you type exactly how much principal is paid each month in that range.',
          '“Interest & Principal (Separate Frequency)” lets that range pay interest and principal on different months — pick the two frequencies in the Interest Freq. and Principal Freq. boxes on the layer. Custom Principal then means the amount paid on EACH principal date; leave it blank to split the balance evenly.',
          'A layer only shows the boxes its own style can use — an EMI layer has no frequency or Custom Principal boxes at all, because they would mean nothing there.',
          'If any layer uses that style, month boxes appear below the table for setting each month’s interest. Only the months belonging to those layers are shown; the rest are left out, because their own payment style already decides what happens to them.',
          'On a narrow screen the table becomes one card per layer, with each box labelled down the side.',
        ],
      },
      { h: 'Total Cost of Fund (COF/ISC + OPEX)', p: 'What the money costs the bank plus running costs, as a yearly rate. The earning is the gap between the Offered Rate and this. With Interest Rate Layers, Refinance Rate months use 1% instead.' },
      {
        h: 'Funded Security Type',
        p: 'Safety money the borrower keeps with the bank. Pick the type, or “No Funded Security” if there is none.',
        bullets: [
          'FDR / Cash Security — one fixed amount.',
          'EMI / EQI after Moratorium, or Installment — built up in installments.',
        ],
      },
      { h: 'The security’s Amount & Rate', p: 'Named after whichever security you picked — pick FDR and they read “FDR Amount” and “FDR Rate”. How much safety money is kept, and its rate. The bank’s gain from it is added to the return.' },
      DOWNLOADS_SECTION,
    ],
  },

  revisionStructured: {
    title: 'ERR - Rate Revision — from the loan details',
    intro: 'You chose to give the original terms rather than upload a finished schedule. List each rate and the day it started, add any security, and upload the bank’s monthly cost-of-fund file — the system rebuilds the whole schedule and finds the ERR.',
    sections: [
      { h: 'Initial Loan Amount', p: 'The starting loan money, for example 100,000,000.' },
      { h: 'Disbursement Date', p: 'The day the loan money was handed out. Everything is counted from this day. It cannot be a Friday or Saturday.' },
      ANSWERED_SECTION,
      { h: 'Moratorium Period (Months)', p: 'How many months the rest period lasts. It only appears when you said there IS a moratorium.' },
      IDP_SECTION,
      { h: 'Payment Modality — chosen in step 1', p: 'Picked before you reached this form; use “Change” at the top to swap it. It sets how the borrower pays after the rest period — EMI (every month), EQI (every 3 months), Equal Principal + Interest, or “Interest & Principal (Separate Frequency)” when interest and principal fall on different months. Picking the last one swaps the rest-period boxes for the same frequency boxes and month grid used on the Structured page, and a rate change part-way through a month is split across that month by day count — the new rate then applies for the rest of the loan.' },
      { h: 'Loan Tenor including Moratorium (Months)', p: 'The total months of the loan, counting the rest period.' },
      {
        h: 'Lending Rate Layers',
        p: 'List each interest rate and the date it starts. The first one starts on the disbursement date. Add a new row every time the rate changes.',
        bullets: [
          'From Date — the day this rate becomes active.',
          'Active Rate — the yearly rate from that day until the next change.',
        ],
      },
      SECURITY_LAYERS_SECTION,
      {
        h: 'COF Data Upload',
        p: 'Upload the bank’s monthly Cost of Fund file. Press “Download Sample File”, fill only the input columns (COF and ISC), then upload it back. The system uses Eligible COF = the bigger of COF and ISC, plus 0.3%.',
        bullets: [
          'The data MUST start on or before the disbursement date. If it does not, the system stops and asks you to add earlier data — nothing is calculated until you do.',
        ],
      },
      REFERENCE_SECTION,
      DOWNLOADS_SECTION,
    ],
  },

  revisionCustomized: {
    title: 'ERR - Rate Revision — from an uploaded schedule',
    intro: 'You already have the full payment schedule in a file, with the rate changes worked out. Upload it, add security and cost-of-fund data, and the system finds the ERR. None of the step-1 questions apply here, which is why you were not asked them.',
    sections: [
      {
        h: 'Upload Amortization Schedule + COF Layers',
        p: 'Press “Download Sample File” to get the workbook. It has two sheets — the payment Schedule and the COF Layers. Fill them in, then upload the file back.',
        bullets: [
          'The Schedule sheet holds every payment: date, installment, interest, principal, and the outstanding balance (URPA).',
          'The COF Layers sheet holds the monthly cost-of-fund data, and must cover from the first schedule date.',
        ],
      },
      SECURITY_LAYERS_SECTION,
      REFERENCE_SECTION,
      {
        h: 'Results & Downloads',
        p: 'The ERR, NIM and Net Interest Income appear on the right as soon as the file is uploaded — there is no Calculate button. Then download the schedule, a Report PDF, or the Verify Calculation Excel that shows every formula — including a sheet with your Loan Security and COF layer tables.',
      },
    ],
  },
};

// Build and show the guide modal for one module.
export function openManual(moduleId) {
  const man = MANUALS[moduleId] || MANUALS.regular;

  const acc = el('div', { class: 'manual-acc' });
  man.sections.forEach((s) => {
    const item = el('div', { class: 'manual-item' });
    const q = el('button', { class: 'manual-q', type: 'button' },
      el('span', {}, s.h),
      el('span', { class: 'manual-chevron' }, '›'),
    );
    const bodyNodes = [];
    if (s.p) bodyNodes.push(el('p', {}, s.p));
    if (s.bullets) bodyNodes.push(el('ul', {}, ...s.bullets.map(b => el('li', {}, b))));
    const a = el('div', { class: 'manual-a' }, ...bodyNodes);
    q.addEventListener('click', () => item.classList.toggle('open'));
    item.appendChild(q);
    item.appendChild(a);
    acc.appendChild(item);
  });

  const closeBtn = el('button', { class: 'manual-close', type: 'button', title: 'Close', 'aria-label': 'Close' }, '×');
  const card = el('div', { class: 'manual-card' },
    el('div', { class: 'manual-head' },
      el('div', {}, el('div', { class: 'manual-kicker' }, 'User Guide'), el('h2', {}, man.title)),
      closeBtn,
    ),
    el('p', { class: 'manual-intro' }, man.intro),
    acc,
  );

  openModal(card);
  const mc = document.getElementById('modal-card');
  mc.classList.add('modal-card--manual');
  const dismiss = () => { mc.classList.remove('modal-card--manual'); closeModal(); };
  closeBtn.addEventListener('click', dismiss);
  const backdrop = document.querySelector('#modal-root .modal-backdrop');
  if (backdrop) backdrop.onclick = dismiss;
}
