// Per-module user guide — opened from the info button in the header. Plain, simple language
// (explained as if to a 10-year-old). Sections are collapsed by default; click to expand.
import { el, openModal, closeModal } from './components.js?v=20260908i';

// ---- Shared sections reused across modules ----
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
  p: 'Press Calculate ERR to see three numbers: ERR (the real yearly return), NIM (the bank’s interest margin), and Net Interest Income (the money earned). Below the results you can download:',
  bullets: [
    'The full payment Schedule as Excel, Word, or PDF.',
    'A Report PDF that summarises the inputs and results.',
    'A Verify Calculation Excel that shows every single formula, so anyone can check the maths.',
  ],
};

const MANUALS = {
  regular: {
    title: 'Loan Facilities — Structured',
    intro: 'This page finds the real yearly return (ERR) on a normal loan that is paid back in equal, regular installments. Fill the boxes from top to bottom, then press Calculate ERR.',
    sections: [
      { h: 'Loan Amount', p: 'The total money the bank hands over to the borrower. Type the whole amount, for example 100,000,000.' },
      { h: 'Offered Rate', p: 'The yearly interest the bank charges, like 12%. Think of it as the price the borrower pays for using the money for one year.' },
      { h: 'Moratorium Available?', p: 'A moratorium is a “rest period” at the very start, when the borrower does not pay the loan back yet. Pick Yes if there is one, otherwise No. If you pick Yes, a box appears for how many months it lasts.' },
      { h: 'Moratorium Period (Months)', p: 'How many months the rest period lasts. For example, 6 means the first 6 months are the rest period.' },
      IDP_SECTION,
      { h: 'Loan Tenor (Months)', p: 'The total life of the loan in months, counting the rest period too. For example, 60 means the loan lasts 5 years.' },
      {
        h: 'Payment Mode',
        p: 'How the borrower pays the loan back after the rest period:',
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
      { h: 'Total Cost of Fund (COF/ISC + OPEX)', p: 'What it costs the bank to get this money, plus its running costs, as a yearly rate. The bank’s real earning is the gap between the Offered Rate and this number.' },
      {
        h: 'Funded Security Type',
        p: 'Sometimes the borrower also keeps some money parked with the bank as safety. Pick the type, or “No Funded Security” if there is none.',
        bullets: [
          'FDR / Cash Security — one fixed safety amount.',
          'EMI / EQI after Moratorium, or Installment — the safety is built up bit by bit in installments.',
        ],
      },
      { h: 'Cash Security / FDR Amount & Rate', p: 'How much safety money is kept, and the rate the bank pays on it. The bank’s gain from this safety money is added to the return.' },
      DOWNLOADS_SECTION,
    ],
  },

  customized: {
    title: 'Loan Facilities — Customized',
    intro: 'Same idea as Loan Facilities — Structured, but here the borrower can pay in different ways during different parts of the loan. You build “Payment Layers” to describe that.',
    sections: [
      { h: 'Loan Amount', p: 'The total money the bank hands over to the borrower. Type the whole amount, for example 100,000,000.' },
      { h: 'Offered Rate', p: 'The yearly interest the bank charges, like 12% — the price of borrowing for one year.' },
      { h: 'Moratorium Available?', p: 'A “rest period” at the start where no repayment happens yet. Pick Yes or No; Yes reveals a box for how many months it lasts.' },
      { h: 'Moratorium Period (Months)', p: 'How many months the rest period lasts.' },
      IDP_SECTION,
      { h: 'Loan Tenor (Months)', p: 'The total life of the loan in months, including the rest period.' },
      {
        h: 'Payment Layers',
        p: 'Instead of one payment style for the whole loan, you split the loan into time ranges and choose a style for each. For example: months 1–12 as EMI, then months 13–24 as EQI.',
        bullets: [
          'Each layer has a From month and a To month.',
          'Layers must line up neatly — no gaps and no overlaps. The system helps you keep them in order.',
          '“Customized Principal” lets you type exactly how much principal is paid each month in that range.',
          '“Interest & Principal (Separate Frequency)” lets that range pay interest and principal on different months — pick the two frequencies in the Interest Freq. and Principal Freq. boxes on the layer. Custom Principal then means the amount paid on EACH principal date; leave it blank to split the balance evenly.',
          'If a layer uses that style, a row of month boxes appears below the table for setting each month’s interest. Months belonging to other layers are greyed out, because their own payment style decides what happens to them.',
        ],
      },
      { h: 'Total Cost of Fund (COF/ISC + OPEX)', p: 'What the money costs the bank plus running costs, as a yearly rate. The earning is the gap between the Offered Rate and this.' },
      {
        h: 'Funded Security Type',
        p: 'Safety money the borrower keeps with the bank. Pick the type, or “No Funded Security” if there is none.',
        bullets: [
          'FDR / Cash Security — one fixed amount.',
          'EMI / EQI after Moratorium, or Installment — built up in installments.',
        ],
      },
      { h: 'Cash Security / FDR Amount & Rate', p: 'How much safety money is kept and its rate. The bank’s gain from it is added to the return.' },
      DOWNLOADS_SECTION,
    ],
  },

  revisionStructured: {
    title: 'Rate Revision — Structured',
    intro: 'Use this when a loan’s interest rate CHANGES over time (a “rate revision”). You list each rate, add any security, and upload the bank’s monthly cost-of-fund file. The system rebuilds the whole schedule and finds the ERR.',
    sections: [
      { h: 'Initial Loan Amount', p: 'The starting loan money, for example 100,000,000.' },
      { h: 'Disbursement Date', p: 'The day the loan money was handed out. Everything is counted from this day. It cannot be a Friday or Saturday.' },
      { h: 'Moratorium Given at Disbursement?', p: 'Whether there is a “rest period” at the start, and (if Yes) how many months it lasts.' },
      IDP_SECTION,
      { h: 'Payment Modality', p: 'How the borrower pays after the rest period — EMI (every month), EQI (every 3 months), Equal Principal + Interest, or “Interest & Principal (Separate Frequency)” when interest and principal fall on different months. Picking the last one swaps the rest-period boxes for the same frequency boxes and month grid used on the Structured page, and a rate change part-way through a month is split across that month by day count — the new rate then applies for the rest of the loan.' },
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
    title: 'Rate Revision — Customized',
    intro: 'Use this when you already have the full payment schedule in a file (with all the rate changes already worked out). You upload it, add security and cost-of-fund data, and the system finds the ERR.',
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
        p: 'Press Calculate ERR to see the ERR, NIM and Net Interest Income. Then download the schedule, a Report PDF, or the Verify Calculation Excel that shows every formula — including a sheet with your Loan Security and COF layer tables.',
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
