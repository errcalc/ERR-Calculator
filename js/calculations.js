// Core financial math + amortization + ERR engine
// All rates passed as decimals (0.12 = 12%). Periodic = monthly unless stated.

export function PMT(rate, nper, pv, fv = 0, type = 0) {
  if (rate === 0) return -(pv + fv) / nper;
  const pvif = Math.pow(1 + rate, nper);
  return -rate * (pv * pvif + fv) / ((pvif - 1) * (1 + rate * type));
}

export function IRR(cashflows, guess = 0.1) {
  let rate = guess;
  for (let i = 0; i < 80; i++) {
    let npv = 0, dnpv = 0;
    for (let t = 0; t < cashflows.length; t++) {
      const denom = Math.pow(1 + rate, t);
      npv += cashflows[t] / denom;
      dnpv += -t * cashflows[t] / (denom * (1 + rate));
    }
    if (Math.abs(npv) < 1e-9) return rate;
    if (dnpv === 0) break;
    const next = rate - npv / dnpv;
    if (!isFinite(next)) break;
    if (Math.abs(next - rate) < 1e-12) return next;
    rate = next;
  }
  let lo = -0.999, hi = 10;
  const npvAt = (r) => cashflows.reduce((s, c, t) => s + c / Math.pow(1 + r, t), 0);
  const fl = npvAt(lo), fh = npvAt(hi);
  if (fl * fh > 0) return rate;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    const fm = npvAt(mid);
    if (Math.abs(fm) < 1e-9) return mid;
    if (fl * fm < 0) hi = mid; else lo = mid;
  }
  return (lo + hi) / 2;
}

function periodsPerYear(mode) {
  if (mode === 'EQI' || mode === 'Equal Principal + Interest (Quarterly)') return 4;
  return 12;
}

// Apply the spec: Int Expense at end of month N = URPA at start of month N * COF/12.
// URPA at start of month N is the post-payment URPA from row N-1.
// Row 0 (disbursement) carries no expense.
// `cofForRow(i)` lets callers vary COF per row (e.g. Rate Revision); default uses constant monthlyCof.
function applyIntExpenseAccrual(rows, monthlyCof, cofForRow = null) {
  for (let i = 0; i < rows.length; i++) {
    if (i === 0) { rows[i].interestExpense = 0; continue; }
    const cof = cofForRow ? cofForRow(i) : monthlyCof;
    rows[i].interestExpense = (rows[i - 1].urpa || 0) * cof;
  }
}

// Months per payment period for a payment type: 1 = monthly, 3 = quarterly.
function monthsPerPeriod(type) {
  const t = String(type || '');
  return (t === 'EQI' || t.includes('(Quarterly)')) ? 3 : 1;
}

// Payments needed to COVER [from..tenor] at `ppm` months each. Any leftover month brings
// one more (imaginary) period: a 25-month span at ppm=3 needs 9, not 8 — same ceil rule as
// buildStructuredSchedule / buildRateRevisionStructured.
function periodsToCover(from, tenor, ppm) {
  return Math.max(1, Math.ceil((tenor - from + 1) / ppm));
}

// ============================================================
// Schedule generation
// ============================================================
// Moratorium logic (per user spec):
//   - In months marked IDP-paid: that month's interest is paid; any accrued interest from prior unpaid months is ALSO paid
//   - In months not marked: interest accrues, no payment
//   - If accrued interest remains at the end of moratorium, it is added to the FIRST regular installment
//
// Row shape: { sl, installment, interest, principal, urpa, interestExpense, idpReceivable, date? }
// ============================================================

export function buildStructuredSchedule(p) {
  const {
    loanAmount, ratePerYear, tenorMonths,
    paymentMode,
    moratoriumMonths = 0,
    idpFlags = [],
    capFlags = [],
    cofRate = 0,
  } = p;

  const ppy = periodsPerYear(paymentMode);
  const regularTenor = tenorMonths - moratoriumMonths;
  // Payments required to COVER the regular period: a partial final quarter still needs a
  // payment, so quarterly uses ceil. When the tenor doesn't divide evenly the installment
  // is sized over the phantom final quarter and the leftover principal is settled on the
  // maturity month (1-2 month stub) with the interest accrued on it — same rule as RRS.
  const regularPeriods = (ppy === 12) ? regularTenor : Math.ceil(regularTenor / 3);
  const stubMonths = (ppy === 12) ? 0 : regularTenor % 3;

  const ratePerPeriod = ratePerYear / ppy;
  const monthlyRate = ratePerYear / 12;
  const monthlyCof = cofRate / 12;

  const rows = [];
  let urpa = loanAmount;
  // Row 0 = disbursement; no interest expense
  rows.push({ sl: 0, installment: 0, interest: 0, principal: 0, urpa, interestExpense: 0, idpReceivable: 0 });

  // Moratorium accrual
  let accruedReceivable = 0;
  for (let m = 1; m <= moratoriumMonths; m++) {
    const interest = urpa * monthlyRate;
    let installment = 0;
    accruedReceivable += interest;
    if (capFlags[m - 1]) {
      // Capitalized: fold accrued-but-unpaid interest into principal at this month-end;
      // interest then accrues on the larger (compounding) principal.
      urpa += accruedReceivable;
      accruedReceivable = 0;
    } else if (idpFlags[m - 1]) {
      // Paid: settle the accrued-but-unpaid interest in cash this month.
      installment = accruedReceivable;
      accruedReceivable = 0;
    }
    // else: keep accruing into the next Paid/Capitalized month or first post-moratorium installment
    rows.push({
      sl: m, installment, interest, principal: 0, urpa,
      interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
    });
  }
  // Post-moratorium principal (grown by any capitalized interest). The regular
  // installments below amortise this amount; funded-security sizing uses it too.
  const capitalizedPrincipal = urpa;

  // Regular installments
  if (regularPeriods <= 0) {
    applyIntExpenseAccrual(rows, monthlyCof);
    return { rows, accruedReceivable, capitalizedPrincipal };
  }

  // Compute the "base" regular installment without accrued-receivable add-on (used as the EMI/installment-size for security calc)
  let baseInstallment = 0;
  if (paymentMode === 'EMI' || paymentMode === 'EQI') {
    const pmt = PMT(ratePerPeriod, regularPeriods, -urpa);
    baseInstallment = pmt;
    let paymentCounter = 0;
    for (let m = moratoriumMonths + 1; m <= tenorMonths; m++) {
      let installment = 0, interest = 0, principal = 0, stubOut;
      const monthsSinceMora = m - moratoriumMonths;
      const isPaymentMonth = (ppy === 12) || (monthsSinceMora % 3 === 0);
      if (isPaymentMonth) {
        interest = urpa * ratePerPeriod;
        installment = pmt;
        if (paymentCounter === regularPeriods - 1) {
          principal = urpa;
          installment = principal + interest + accruedReceivable;
          accruedReceivable = 0;
        } else {
          principal = installment - interest;
        }
        if (paymentCounter === 0 && accruedReceivable > 0) {
          installment += accruedReceivable;
          accruedReceivable = 0;
        }
        urpa = Math.max(0, urpa - principal);
        paymentCounter++;
      } else if (stubMonths > 0 && m === tenorMonths) {
        // Maturity stub: pay the leftover principal + the interest accrued on it since the
        // last quarterly due date (+ any still-uncollected moratorium accrued).
        interest = urpa * monthlyRate * stubMonths;
        principal = urpa;
        installment = principal + interest + accruedReceivable;
        accruedReceivable = 0;
        urpa = 0;
        stubOut = stubMonths;
      }
      rows.push({
        sl: m, installment, interest, principal, urpa,
        interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
        paymentType: paymentMode, isPaymentMonth, stubMonths: stubOut,
      });
    }
  } else {
    // Equal Principal + Interest
    const principalPer = urpa / regularPeriods;
    let paymentCounter = 0;
    // Base installment for security size = principal-per + interest on initial URPA at first payment
    const firstInterest = urpa * ratePerPeriod;
    baseInstallment = principalPer + firstInterest;
    for (let m = moratoriumMonths + 1; m <= tenorMonths; m++) {
      let installment = 0, interest = 0, principal = 0, stubOut;
      const monthsSinceMora = m - moratoriumMonths;
      const isPaymentMonth = (ppy === 12) || (monthsSinceMora % 3 === 0);
      if (isPaymentMonth) {
        interest = urpa * ratePerPeriod;
        principal = principalPer;
        installment = principal + interest;
        if (paymentCounter === regularPeriods - 1) {
          principal = urpa;
          installment = principal + interest + accruedReceivable;
          accruedReceivable = 0;
        }
        if (paymentCounter === 0 && accruedReceivable > 0) {
          installment += accruedReceivable;
          accruedReceivable = 0;
        }
        urpa = Math.max(0, urpa - principal);
        paymentCounter++;
      } else if (stubMonths > 0 && m === tenorMonths) {
        // Maturity stub: pay the leftover principal + the interest accrued on it since the
        // last quarterly due date (+ any still-uncollected moratorium accrued).
        interest = urpa * monthlyRate * stubMonths;
        principal = urpa;
        installment = principal + interest + accruedReceivable;
        accruedReceivable = 0;
        urpa = 0;
        stubOut = stubMonths;
      }
      rows.push({
        sl: m, installment, interest, principal, urpa,
        interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
        paymentType: paymentMode, isPaymentMonth, stubMonths: stubOut,
      });
    }
  }
  applyIntExpenseAccrual(rows, monthlyCof);
  return { rows, accruedReceivable, baseInstallment, capitalizedPrincipal };
}

// Customized Loan
export function buildCustomizedSchedule(p) {
  const {
    loanAmount, ratePerYear, tenorMonths,
    moratoriumMonths = 0,
    idpFlags = [],
    capFlags = [],
    cofRate = 0,
    layers,
  } = p;

  const monthlyRate = ratePerYear / 12;
  const monthlyCof = cofRate / 12;

  const rows = [];
  let urpa = loanAmount;
  rows.push({ sl: 0, installment: 0, interest: 0, principal: 0, urpa, interestExpense: 0, idpReceivable: 0 });

  let accruedReceivable = 0;
  for (let m = 1; m <= moratoriumMonths; m++) {
    const interest = urpa * monthlyRate;
    let installment = 0;
    accruedReceivable += interest;
    if (capFlags[m - 1]) {
      // Capitalized: fold accrued-but-unpaid interest into principal at this month-end;
      // interest then accrues on the larger (compounding) principal.
      urpa += accruedReceivable;
      accruedReceivable = 0;
    } else if (idpFlags[m - 1]) {
      // Paid: settle the accrued-but-unpaid interest in cash this month.
      installment = accruedReceivable;
      accruedReceivable = 0;
    }
    // else: keep accruing into the next Paid/Capitalized month or first post-moratorium installment
    rows.push({
      sl: m, installment, interest, principal: 0, urpa,
      interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
    });
  }
  // Post-moratorium principal (grown by any capitalized interest); the payment layers
  // below amortise this amount and funded-security sizing uses it.
  const capitalizedPrincipal = urpa;

  // Layers' fromInstallment/toInstallment are ABSOLUTE month numbers (Month NN values)
  // E.g. if moratorium=6 and tenor=36, layers have from=7, to=36
  const sorted = layers.slice().sort((a, b) => (a.fromInstallment ?? 0) - (b.fromInstallment ?? 0));

  // Track per-layer base installment for security-size lookup
  const layerInstallments = {}; // { 'EMI': base_pmt, 'EQI': base_pmt, ... }

  for (let li = 0; li < sorted.length; li++) {
    const L = sorted[li];
    const from = L.fromInstallment;
    const to = L.toInstallment;
    const count = to - from + 1;
    const isLastLayer = (li === sorted.length - 1);

    let pmt = 0;
    // Payment months are LAYER-RELATIVE: the layer's own start month counts as month 1 of its
    // first period, so quarterly pays on layer months 3, 6, 9, ... The layer's FINAL month is
    // always a payment month too — if it isn't a whole period from the previous one it becomes
    // a stub (same installment, interest prorated over the months actually elapsed).
    const ppm = monthsPerPeriod(L.paymentType);
    const isPaymentMonth = (m) => ((m - from + 1) % ppm === 0) || (m === to);

    // Equal Principal + Interest layers: the per-period principal is CONSTANT and sized over
    // the periods remaining to MATURITY (not just the layer's own span). So a layer that ends
    // before maturity only partially amortizes, leaving a balance for the next layer.
    const layerStartBalance = urpa;
    let epiConstPrincipal = 0;
    if (L.paymentType && L.paymentType.startsWith('Equal Principal + Interest')) {
      const periodsToMaturity = periodsToCover(from, tenorMonths, ppm);
      epiConstPrincipal = layerStartBalance / periodsToMaturity;
    }

    if (L.paymentType === 'EMI' || L.paymentType === 'EQI') {
      // Size the annuity over the periods remaining to MATURITY (not just this layer's span),
      // so a layer that ends before maturity only partially amortises and leaves a balance for
      // the next layer (matching the Equal-Principal layer behaviour). Leftover months round
      // the period count UP — see periodsToCover.
      pmt = PMT(ratePerYear * ppm / 12, periodsToCover(from, tenorMonths, ppm), -urpa);
      if (!layerInstallments[L.paymentType]) layerInstallments[L.paymentType] = pmt;
    } else if (L.paymentType === 'Equal Principal + Interest (Monthly)' && !layerInstallments['Installment']) {
      layerInstallments['Installment'] = (urpa / count) + urpa * monthlyRate;
    } else if (L.paymentType === 'Equal Principal + Interest (Quarterly)' && !layerInstallments['Installment']) {
      const qPeriods = periodsToCover(from, to, 3);
      layerInstallments['Installment'] = (urpa / qPeriods) + urpa * (ratePerYear / 4);
    } else if (L.paymentType && L.paymentType.startsWith('Customized Principal') && !layerInstallments['Customized']) {
      layerInstallments['Customized'] = (L.customPrincipal || 0) + urpa * monthlyRate;
    }

    let paymentCounter = 0;
    // Interest is settled on payment months only, over the months actually ELAPSED since the
    // previous payment. A whole quarter gives exactly balance * rate/4; a layer-end stub
    // prorates. Seeded at from-1 so the layer's first payment covers from its start month.
    let prevPayMonth = from - 1;
    const ptype = String(L.paymentType || '');
    for (let m = from; m <= to; m++) {
      let installment = 0, interest = 0, principal = 0, stubOut;
      const isLastInstallmentOfLoan = isLastLayer && (m === to);
      const elapsed = m - prevPayMonth;

      if (isPaymentMonth(m)) {
        interest = urpa * monthlyRate * elapsed;
        if (elapsed !== ppm) stubOut = elapsed; // short final period of the layer
        if (ptype === 'EMI' || ptype === 'EQI') {
          installment = pmt;
          principal = installment - interest;
        } else if (ptype.startsWith('Equal Principal + Interest')) {
          principal = Math.min(epiConstPrincipal, urpa);
          installment = principal + interest;
        } else if (ptype.startsWith('Customized Principal')) {
          principal = Math.min(L.customPrincipal || 0, urpa);
          installment = principal + interest;
        }
        prevPayMonth = m;
      }

      // Final month of the LOAN: settle whatever principal remains, plus the interest accrued
      // on it since the previous payment (the maturity-stub rule the Structured modules use)
      // and any still-uncollected moratorium interest.
      if (isLastInstallmentOfLoan) {
        principal = urpa;
        installment = principal + interest + accruedReceivable;
        accruedReceivable = 0;
      } else if (li === 0 && paymentCounter === 0 && accruedReceivable > 0 && installment > 0) {
        // Unpaid moratorium interest rides on top of the first installment of the first
        // layer, and stays OUT of principal.
        installment += accruedReceivable;
        accruedReceivable = 0;
      }
      urpa = Math.max(0, urpa - principal);

      rows.push({
        sl: m, installment, interest, principal, urpa,
        interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
        paymentType: L.paymentType, isPaymentMonth: isPaymentMonth(m), stubMonths: stubOut,
      });

      if (installment > 0) paymentCounter++;
    }
  }
  applyIntExpenseAccrual(rows, monthlyCof);
  return { rows, accruedReceivable, layerInstallments, capitalizedPrincipal };
}

// ============================================================
// Split interest / principal schedules
// ============================================================
// The two legs pay on independent calendars. Interest accrues monthly on the outstanding
// balance; every month is Paid (settle the running accrual in cash), Accrued (carry it to
// the next Paid month) or Capitalized (fold it into principal at that month's end, so it
// compounds from there — convention 5). A month carrying a principal payment is always
// Paid, so the balance can never move inside an unsettled interest stretch.

// Principal payment months: counting from `startMonth` as month 1, one on every multiple of
// `period`, plus maturity — which always settles whatever principal is left.
export function principalPaymentMonths(startMonth, tenorMonths, period) {
  const out = [];
  for (let m = startMonth; m <= tenorMonths; m++) {
    if ((m - startMonth + 1) % period === 0) out.push(m);
  }
  if (out[out.length - 1] !== tenorMonths && tenorMonths >= startMonth) out.push(tenorMonths);
  return out;
}

// Default interest treatment for a month, before any per-month override from the grid.
// Principal months are forced Paid; otherwise the frequency decides, counting from month 1.
function defaultIntFlag(m, interestPeriod, principalSet, tenorMonths) {
  if (principalSet.has(m) || m === tenorMonths) return 'paid';
  return (m % interestPeriod === 0) ? 'paid' : 'accrued';
}

export function buildSplitSchedule(p) {
  const {
    loanAmount, ratePerYear, tenorMonths,
    interestPeriod = 1,        // months per interest period: 1 / 3 / 6 / 12
    principalPeriod = 1,       // months per principal period: 1 / 3 / 6 / 12
    principalStartMonth = 1,   // where the principal frequency starts counting
    principalBasis = 'fixed',  // 'fixed' (equal) | 'custom' (one amount per date)
    customPrincipals = [],     // used when principalBasis === 'custom'
    intFlags = [],             // per-month override: 'paid' | 'accrued' | 'capitalized'
    cofRate = 0,
  } = p;

  const monthlyRate = ratePerYear / 12;
  const monthlyCof = cofRate / 12;
  const pMonths = principalPaymentMonths(principalStartMonth, tenorMonths, principalPeriod);
  const pSet = new Set(pMonths);

  const rows = [{ sl: 0, installment: 0, interest: 0, principal: 0, urpa: loanAmount, interestExpense: 0, idpReceivable: 0 }];
  let urpa = loanAmount;
  let accruedReceivable = 0;
  let paidIdx = 0;

  for (let m = 1; m <= tenorMonths; m++) {
    // Interest first, then principal — so the month's interest is always on the opening balance.
    const interest = urpa * monthlyRate;
    accruedReceivable += interest;
    const flag = pSet.has(m) ? 'paid'
      : (intFlags[m - 1] || defaultIntFlag(m, interestPeriod, pSet, tenorMonths));

    let installment = 0, principal = 0;
    if (flag === 'capitalized') {
      urpa += accruedReceivable;          // compounds from here (convention 5)
      accruedReceivable = 0;
    } else if (flag === 'paid') {
      installment = accruedReceivable;
      accruedReceivable = 0;
    }

    if (pSet.has(m)) {
      const remaining = pMonths.length - paidIdx;
      if (m === tenorMonths) {
        principal = urpa;                 // maturity settles the balance, whatever it is
      } else if (principalBasis === 'custom') {
        principal = Math.min(customPrincipals[paidIdx] || 0, urpa);
      } else {
        // Fixed: re-divide the live balance across the dates still to come. With nothing
        // perturbing it this yields identical payments; after a capitalized month it
        // re-sizes the remainder automatically, which is the behaviour the user specified.
        principal = urpa / remaining;
      }
      installment += principal;
      urpa = Math.max(0, urpa - principal);
      paidIdx++;
    }

    rows.push({
      sl: m, installment, interest, principal, urpa,
      interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
      isPaymentMonth: installment > 0, intFlag: flag, isPrincipalMonth: pSet.has(m),
    });
  }

  applyIntExpenseAccrual(rows, monthlyCof);
  return { rows, accruedReceivable, capitalizedPrincipal: loanAmount, principalMonths: pMonths };
}

// ============================================================
// Metrics — keep only what's needed; surface ERR primarily, plus NIM% and NII$
// ============================================================
export function computeMetrics(schedule, params) {
  const { loanAmount, ratePerYear, cofRate = 0, paymentMode, securityKind, numInst = 0,
          tenorMonths = 0, moratoriumMonths = 0 } = params;
  let securityAmount = params.securityAmount || 0;
  let securityRate = params.securityRate || 0;
  const ppy = periodsPerYear(paymentMode || 'EMI');
  const rows = schedule.rows;

  // Installment-equivalent funded security.
  //   "EMI after Moratorium" = PMT(rate/12, regularMonths, -loan)   (B14 in the sample)
  //   "EQI after Moratorium" = PMT(rate/4,  regularQuarters, -loan)
  //   "Installment"          = the post-moratorium Equal-Principal installment (baseInstallment)
  // These are the hypothetical EMI/EQI size after moratorium, independent of the actual layers.
  const kind = String(securityKind || '');
  const regularMonths = Math.max(0, tenorMonths - moratoriumMonths);
  // When moratorium interest is capitalized the post-moratorium principal is larger,
  // so the funded-security installment is sized on that (grown) principal.
  const basePrincipal = schedule.capitalizedPrincipal || loanAmount;
  if (kind.startsWith('EMI') || kind.startsWith('EQI') || kind === 'Installment') {
    let unitInstallment = 0;
    if (kind.startsWith('EMI')) {
      unitInstallment = regularMonths > 0 ? PMT(ratePerYear / 12, regularMonths, -basePrincipal) : 0;
    } else if (kind.startsWith('EQI')) {
      const q = Math.max(1, Math.round(regularMonths / 3));
      unitInstallment = PMT(ratePerYear / 4, q, -basePrincipal);
    } else if (schedule.baseInstallment) {
      // Structured Equal-Principal: first installment size
      unitInstallment = schedule.baseInstallment;
    }
    securityAmount = unitInstallment * (numInst || 1);
    // securityRate keeps the passed value — the "Funded Security Rate" for EMI/EQI after
    // Moratorium (0 for plain "Installment", which has no rate field).
  }

  // Avg portfolio = avg of URPA at start of each month (rows 0..N-1)
  const urpaSeries = rows.slice(0, -1).map(r => r.urpa);
  const avgPortfolio = urpaSeries.length ? urpaSeries.reduce((s, v) => s + v, 0) / urpaSeries.length : 0;

  const totalInterest = rows.reduce((s, r) => s + r.interest, 0);
  // Int expense accrues at end of every month from sl 1 through sl N (last month).
  const totalInterestExpense = rows.slice(1).reduce((s, r) => s + (r.interestExpense || 0), 0);
  const totalMonths = rows.length - 1;
  const tenorYears = totalMonths / 12;

  const csBenefit = (cofRate - securityRate) * (securityAmount || 0) * tenorYears;
  const netInterestExpense = totalInterestExpense - csBenefit;

  // NII ($) per user: Interest received + CS savings - Interest paid
  const nii = totalInterest + csBenefit - totalInterestExpense;
  const nim = avgPortfolio > 0 && tenorYears > 0 ? (nii / avgPortfolio) / tenorYears : 0;

  // ERR — match Excel I4 = E6 + I3 = COF + NIM. Internally computed.
  const effectiveRate = cofRate + nim;

  return {
    effectiveRate,
    nim,
    nii,
    // Auxiliary (not displayed, used for context)
    avgPortfolio,
    totalInterest,
    totalInterestExpense,
    csBenefit,
    netInterestExpense,
    tenorYears,
    derivedSecurityAmount: securityAmount,
    derivedSecurityRate: securityRate,
  };
}

// ============================================================
// Rate Revision — Structured
// ============================================================

// Add months to a date keeping the loan's due day-of-month, per the banking convention:
// a due day of 28/29/30/31 falls on February's LAST day (28 non-leap, 29 leap) whenever
// February occurs in the schedule, and a 31st due day falls on the 30th in 30-day months.
// Always computed from the original anchor date, so the due day never drifts (after a
// 28-Feb due the schedule returns to the 29th/30th/31st in March).
export function addMonthsDue(d, m) {
  const x = new Date(d);
  const day = x.getDate();
  x.setDate(1);
  x.setMonth(x.getMonth() + m);
  const lastDay = new Date(x.getFullYear(), x.getMonth() + 1, 0).getDate();
  const isFeb = x.getMonth() === 1;
  x.setDate(isFeb && day >= 28 ? lastDay : Math.min(day, lastDay));
  return x;
}

// Excel DAYS360 (US/NASD method): months are 30 days, years 360. Used for the day-count split
// of the Loan Security Benefit when a security amount is taken on a non-installment date.
export function days360(startISO, endISO) {
  const [sy, sm, sdRaw] = String(startISO).split('-').map(Number);
  const [ey, em, edRaw] = String(endISO).split('-').map(Number);
  let sd = sdRaw, ed = edRaw;
  if (sd === 31) sd = 30;
  if (ed === 31 && sd === 30) ed = 30;
  return (ey - sy) * 360 + (em - sm) * 30 + (ed - sd);
}

export function buildRateRevisionStructured(p) {
  const {
    initialLoanAmount,
    disbursementDate,
    moratoriumMonths = 0,
    idpFlags = [],
    capFlags = [],
    paymentModality,
    tenorMonths,
    rateLayers,
    securityLayers,
    cofData = null,     // sorted array of { date: 'YYYY-MM-DD', rate } from the uploaded COF file
  } = p;

  const ppy = periodsPerYear(paymentModality);
  const start = new Date(disbursementDate);

  const addMonths = addMonthsDue;
  function fmt(d) { return d.toISOString().slice(0, 10); }

  // Lending rate effective on a date = the LAST layer whose fromDate <= date (extends the last
  // rate forward past the final layer's To, so the maturity-month payment keeps the last rate).
  function getRateOn(dateStr) {
    if (!rateLayers || !rateLayers.length) return 0;
    let result = rateLayers[0].activeRate || 0;
    let bestFrom = null;
    for (const r of rateLayers) {
      if (r.fromDate && r.fromDate <= dateStr && (bestFrom === null || r.fromDate >= bestFrom)) {
        bestFrom = r.fromDate; result = r.activeRate || 0;
      }
    }
    return result;
  }
  // COF effective on a date = the LAST uploaded COF entry whose date <= the given date.
  // 0 if the date precedes the first entry (uncovered period). cofData is pre-filtered to
  // entries effective on/before maturity, so the last entry naturally extends to maturity.
  function getCofOn(dateStr) {
    if (!cofData || !cofData.length) return 0;
    let result = 0;
    for (const e of cofData) {
      if (e.date <= dateStr) result = e.rate; else break;
    }
    return result;
  }
  // Interest earned over an accrual period (startISO, endISO] on a constant balance.
  // If no rate layer takes effect strictly inside the period, the nominal convention
  // applies: balance * rate-at-period-start / nominalDivisor (12 monthly, 4 quarterly).
  // If a revision's From Date falls inside the period (a non-due date), the period is
  // split into day-count segments: balance * rate_i * DAYS360(segStart,segEnd) / 360 per
  // segment — the fractional-interest treatment from the rectified file (Excel DAYS360 basis).
  function periodInterest(balance, startISO, endISO, nominalDivisor) {
    const cuts = [...new Set((rateLayers || [])
      .map(l => l.fromDate)
      .filter(fd => fd && fd > startISO && fd < endISO))].sort();
    if (!cuts.length) {
      return { interest: balance * getRateOn(startISO) / nominalDivisor, segments: null };
    }
    const bounds = [startISO, ...cuts, endISO];
    const segments = [];
    let interest = 0;
    for (let i = 0; i < bounds.length - 1; i++) {
      const segStart = bounds[i], segEnd = bounds[i + 1];
      const days = days360(segStart, segEnd);
      const rate = getRateOn(segStart);
      interest += balance * rate / 360 * days;
      segments.push({ rate, from: segStart, to: segEnd, days });
    }
    return { interest, segments };
  }

  // Loan security effective on a date: the most recent layer whose From <= date wins — its
  // Amount is the security balance in effect (a total, not an increment) and its rate is the
  // rate in effect, until the next layer's From. No cumulative sum, no weighted average.
  function getSecurityOn(dateStr) {
    if (!securityLayers || !securityLayers.length) return { amount: 0, rate: 0 };
    let result = { amount: 0, rate: 0 }, bestFrom = null;
    for (const r of securityLayers) {
      if (r.fromDate && r.fromDate <= dateStr && (bestFrom === null || r.fromDate >= bestFrom)) {
        bestFrom = r.fromDate; result = { amount: r.amount || 0, rate: r.activeRate || 0 };
      }
    }
    return result;
  }

  const rows = [];
  let urpa = initialLoanAmount;
  const date0 = fmt(start);
  const sec0 = getSecurityOn(date0);
  rows.push({
    sl: 0, date: date0, installment: 0, interest: 0, principal: 0, urpa,
    interestExpense: 0,
    rate: getRateOn(date0), cof: getCofOn(date0),
    securityAmount: sec0.amount, securityRate: sec0.rate,
  });

  let accruedReceivable = 0;

  for (let m = 1; m <= tenorMonths; m++) {
    const d = fmt(addMonths(start, m));
    // COF is keyed to the row's OWN month (this row's date); interest expense uses the
    // PREVIOUS row's URPA and COF via applyIntExpenseAccrual — matching the rectified files.
    const accrualStart = fmt(addMonths(start, m - 1));
    const cofPm = getCofOn(d);
    const sec = getSecurityOn(d); // security uses the current month's date

    let installment = 0, interest = 0, principal = 0;
    // rowRate: rate shown/used for this row's accrual (period-start rate; after a
    // mid-period revision the row carries the NEW rate so block logic continues from it).
    let rowRate = getRateOn(accrualStart), splitSegments = null, splitFromSl, stubMonthsOut;
    if (m <= moratoriumMonths) {
      // Moratorium month: interest accrues monthly; a revision mid-month splits by days/360.
      const pi = periodInterest(urpa, accrualStart, d, 12);
      interest = pi.interest;
      if (pi.segments) { splitSegments = pi.segments; splitFromSl = m - 1; rowRate = getRateOn(d); }
      accruedReceivable += interest;
      if (capFlags[m - 1]) {
        // Capitalized: fold accrued-but-unpaid interest into principal at this month-end.
        urpa += accruedReceivable;
        accruedReceivable = 0;
      } else if (idpFlags[m - 1]) {
        // Paid: settle accrued-but-unpaid interest in cash this month.
        installment = accruedReceivable;
        accruedReceivable = 0;
      }
    } else {
      const monthsSinceMora = m - moratoriumMonths;
      const isPmtMonth = (ppy === 12) || (monthsSinceMora % 3 === 0);
      // Quarterly grid that doesn't divide the remaining tenor evenly leaves a 1-2 month
      // stub; the leftover principal + its accrued interest is settled at maturity.
      const stubAtMaturity = !isPmtMonth && m === tenorMonths;
      if (isPmtMonth) {
        // Payments required to COVER the period from this due date to maturity — a partial
        // final quarter still needs one payment, so quarterly uses ceil((months from the
        // period start)/3). Equals the plain grid count when the tenor divides evenly;
        // otherwise the installment is sized over the phantom final quarter and the
        // residual principal lands on the maturity stub row.
        const remainingPayments = (ppy === 12)
          ? (tenorMonths - m + 1)
          : Math.ceil((tenorMonths - m + 3) / 3);
        const step = (ppy === 12) ? 1 : 3;
        // The payment period runs from the previous due date (or moratorium end) to this
        // due date; a rate revision inside it splits the interest by days/360.
        const periodStartISO = fmt(addMonths(start, m - step));
        const pi = periodInterest(urpa, periodStartISO, d, ppy);
        interest = pi.interest;
        const rateAtStart = getRateOn(periodStartISO);
        rowRate = pi.segments ? getRateOn(d) : rateAtStart;
        if (pi.segments) { splitSegments = pi.segments; splitFromSl = m - step; }
        if (paymentModality === 'EMI' || paymentModality === 'EQI') {
          if (pi.segments) {
            // Mid-period revision: derive the constant installment X that amortises the
            // balance to exactly 0 at maturity given this period's split interest and the
            // new rate onward — closed form of the rectified file's Goal Seek:
            //   X = (balance + split interest) / (1 + annuity(rNew, remainingPayments - 1))
            const rNew = getRateOn(d) / ppy;
            const k = remainingPayments - 1;
            const annuity = k <= 0 ? 0 : (rNew === 0 ? k : (1 - Math.pow(1 + rNew, -k)) / rNew);
            installment = (urpa + interest) / (1 + annuity);
          } else {
            const periodRate = rateAtStart / ppy;
            installment = remainingPayments > 0 ? PMT(periodRate, remainingPayments, -urpa) : urpa * (1 + periodRate);
          }
          principal = installment - interest;
        } else {
          principal = urpa / remainingPayments;
          installment = principal + interest;
        }
        if (m === tenorMonths) {
          principal = urpa;
          installment = principal + interest;
        }
        // first regular installment carries accrued
        if (monthsSinceMora === 1 && accruedReceivable > 0) {
          installment += accruedReceivable;
          accruedReceivable = 0;
        } else if (ppy === 4 && monthsSinceMora === 3 && accruedReceivable > 0) {
          installment += accruedReceivable;
          accruedReceivable = 0;
        }
        urpa = Math.max(0, urpa - principal);
      } else if (stubAtMaturity) {
        // Maturity falls 1-2 months after the last quarterly due date: pay the outstanding
        // principal plus the interest accrued on it since that due date (nominal monthly
        // twelfths; a rate revision inside the stub splits by days/360).
        const stubMonths = monthsSinceMora % 3;
        const periodStartISO = fmt(addMonths(start, m - stubMonths));
        const pi = periodInterest(urpa, periodStartISO, d, 12 / stubMonths);
        interest = pi.interest;
        rowRate = pi.segments ? getRateOn(d) : getRateOn(periodStartISO);
        if (pi.segments) { splitSegments = pi.segments; splitFromSl = m - stubMonths; }
        stubMonthsOut = stubMonths;
        principal = urpa;
        installment = principal + interest;
        if (accruedReceivable > 0) { // moratorium accrued never collected (no grid payment fit)
          installment += accruedReceivable;
          accruedReceivable = 0;
        }
        urpa = 0;
      }
    }

    rows.push({
      sl: m, date: d, installment, interest, principal, urpa,
      interestExpense: 0, // populated below per applyIntExpenseAccrual
      rate: rowRate, cof: cofPm,
      securityAmount: sec.amount, securityRate: sec.rate,
      idpReceivable: accruedReceivable,
      splitSegments: splitSegments || undefined, splitFromSl,
      stubMonths: stubMonthsOut,
    });
  }

  // Loan Security Balance is released at maturity: show 0 on the final payment row.
  if (rows.length > 1) rows[rows.length - 1].securityAmount = 0;

  // Loan Security Benefit per row (one-month lag): row N's benefit uses row N-1's COF and the
  // security state across [date(N-1), date(N)]. A security amount TAKEN mid-period (its fromDate
  // strictly inside that window — i.e. on a non-installment day) splits the month by day-count
  // (DAYS360): each segment accrues on its own cumulative balance and weighted-average rate,
  // both against the period-start COF. Clean (on-installment) dates keep the flat /12 accrual.
  if (rows.length) rows[0].securityBenefit = 0;
  for (let i = 1; i < rows.length; i++) {
    const d0 = rows[i - 1].date, d1 = rows[i].date, cof = rows[i - 1].cof || 0;
    const cuts = securityLayers
      .filter(l => l.fromDate && l.fromDate > d0 && l.fromDate < d1)
      .map(l => l.fromDate)
      .sort();
    if (!cuts.length) {
      const s = getSecurityOn(d0);
      rows[i].securityBenefit = s.amount * (cof - s.rate) / 12;
    } else {
      const bounds = [d0, ...cuts, d1];
      let ben = 0;
      for (let k = 0; k < bounds.length - 1; k++) {
        const s = getSecurityOn(bounds[k]);
        ben += s.amount * (cof - s.rate) * days360(bounds[k], bounds[k + 1]) / 360;
      }
      rows[i].securityBenefit = ben;
    }
  }

  // Int. Expense at month N = URPA(N-1) * COF(N-1) / 12 — the PREVIOUS month's URPA and
  // COF (the COF prevailing at the start of the accrual month), per the rectified file.
  applyIntExpenseAccrual(rows, 0, (i) => (rows[i - 1].cof || 0) / 12);
  return { rows };
}

// Convert uploaded COF rows -> sorted, maturity-filtered effective-date list.
// Returns { cofData, warning }. cofData = [{date, rate}] ascending; warning set if the first
// COF effective date is after disbursement (uncovered head period uses 0%).
export function buildCofData(uploadedCof, disbursementISO, maturityISO) {
  if (!uploadedCof || !uploadedCof.length) return { cofData: [], warning: null };
  const sorted = uploadedCof
    .filter(e => e.date && (e.rate !== null && e.rate !== undefined))
    .map(e => ({ date: e.date, rate: Number(e.rate), cof: e.cof, isc: e.isc }))
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  // Only consider records effective on/before maturity.
  const filtered = maturityISO ? sorted.filter(e => e.date <= maturityISO) : sorted;
  let warning = null;
  if (filtered.length && disbursementISO && filtered[0].date > disbursementISO) {
    warning = `The uploaded COF data does not cover the full loan timeline — no COF before ${filtered[0].date}. ` +
              `0.00% COF (no interest expense) is applied from disbursement (${disbursementISO}) until then. ` +
              `Add an earlier COF record to cover the whole period.`;
  }
  return { cofData: filtered, warning };
}

export function computeRevisionMetrics(schedule) {
  const rows = schedule.rows;
  const urpaSeries = rows.slice(0, -1).map(r => r.urpa);
  const avgPortfolio = urpaSeries.length ? urpaSeries.reduce((s, v) => s + v, 0) / urpaSeries.length : 0;
  const totalInterest = rows.reduce((s, r) => s + r.interest, 0);
  const totalInterestExpense = rows.slice(1).reduce((s, r) => s + (r.interestExpense || 0), 0);
  const totalMonths = rows.length - 1;
  const tenorYears = totalMonths / 12;

  // Loan Security Benefit is precomputed per row by buildRateRevisionStructured (one-month lag,
  // with a day-count split for mid-period security additions). Sum it; fall back to the flat
  // prior-month accrual for any row lacking the precomputed value.
  let csBenefit = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    csBenefit += r.securityBenefit != null
      ? r.securityBenefit
      : ((rows[i - 1].cof || 0) - (rows[i - 1].securityRate || 0)) * (rows[i - 1].securityAmount || 0) / 12;
  }
  const nii = totalInterest + csBenefit - totalInterestExpense;
  const nim = avgPortfolio > 0 && tenorYears > 0 ? (nii / avgPortfolio) / tenorYears : 0;
  // Effective COF = total interest expense / avg portfolio / tenor years; ERR = NIM + effective COF.
  const effectiveCof = avgPortfolio > 0 && tenorYears > 0 ? (totalInterestExpense / avgPortfolio) / tenorYears : 0;
  const effectiveRate = nim + effectiveCof;

  return {
    effectiveRate, nim, nii,
    avgPortfolio, totalInterest, totalInterestExpense, csBenefit, effectiveCof,
    tenorYears,
  };
}

export function computeRevisionCustomizedMetrics(uploadedRows, { securityLayers = [], cofData = null } = {}) {
  const rows = uploadedRows;
  if (!rows || rows.length < 2) return null;

  const dates = rows.map(r => new Date(r.date));
  const periodicRates = [];
  const exposure = [];
  for (let i = 1; i < rows.length; i++) {
    const prev = rows[i - 1];
    const cur = rows[i];
    const days = Math.max(1, (dates[i] - dates[i - 1]) / 86400000);
    const r = prev.urpa > 0 ? (cur.interestAmount / prev.urpa) : 0;
    const annualR = r * (365 / days);
    periodicRates.push(annualR);
    exposure.push(prev.urpa * days);
  }
  const totalExposure = exposure.reduce((s, v) => s + v, 0);
  const yolWeighted = totalExposure > 0
    ? periodicRates.reduce((s, r, i) => s + r * exposure[i], 0) / totalExposure
    : 0;

  const totalInterest = rows.reduce((s, r) => s + (r.interestAmount || 0), 0);
  const totalDays = (dates[dates.length - 1] - dates[0]) / 86400000;
  const avgPortfolio = totalDays > 0 ? totalExposure / totalDays : 0;
  const tenorYears = totalDays / 365;

  let csBenefit = 0, totalInterestExpense = 0, effectiveRate = yolWeighted;
  if (cofData && cofData.length) {
    // COF effective on a date = the LAST uploaded entry whose date <= the given date
    // (0 before the first entry / uncovered period). Mirrors Rate Revision — Structured.
    const getCofOn = (dateStr) => {
      let result = 0;
      for (const e of cofData) { if (e.date <= dateStr) result = e.rate; else break; }
      return result;
    };
    const getSecOn = (dateStr) => {
      // Most recent security layer whose From <= date wins — its Amount (the balance in effect)
      // and its rate, until the next layer's From. No cumulative sum, no weighted average.
      let result = { amount: 0, rate: 0 }, bestFrom = null;
      for (const r of (securityLayers || []))
        if (r.fromDate && r.fromDate <= dateStr && (bestFrom === null || r.fromDate >= bestFrom)) {
          bestFrom = r.fromDate; result = { amount: r.amount || 0, rate: r.activeRate || 0 };
        }
      return result;
    };
    for (let i = 1; i < rows.length; i++) {
      const prev = rows[i - 1];
      const days = Math.max(1, (dates[i] - dates[i - 1]) / 86400000);
      const dateStr = rows[i - 1].date;
      const cof = getCofOn(dateStr);
      const sec = getSecOn(dateStr);
      totalInterestExpense += prev.urpa * cof * (days / 365);
      csBenefit += (cof - sec.rate) * sec.amount * (days / 365);
    }
    const nii = totalInterest + csBenefit - totalInterestExpense;
    const nim = avgPortfolio > 0 && tenorYears > 0 ? (nii / avgPortfolio) / tenorYears : 0;
    const effectiveCof = avgPortfolio > 0 && tenorYears > 0 ? (totalInterestExpense / avgPortfolio) / tenorYears : 0;
    effectiveRate = nim + effectiveCof;
    return { effectiveRate, nim, nii, avgPortfolio, totalInterest, totalInterestExpense, csBenefit, tenorYears };
  }
  const nii = totalInterest - totalInterestExpense + csBenefit;
  const nim = avgPortfolio > 0 && tenorYears > 0 ? (nii / avgPortfolio) / tenorYears : 0;
  return { effectiveRate, nim, nii, avgPortfolio, totalInterest, totalInterestExpense, csBenefit, tenorYears };
}
