// Excel / Word / PDF I/O via CDN libs
import { formatMoney as fmtM, formatPercent as fmtP } from './formatting.js?v=20260908d';

// Round a cell value to 2 decimals (numeric — kept distinct from fmtM which returns a string).
function num(v) {
  if (v === null || v === undefined || isNaN(v)) return 0;
  return Math.round(Number(v) * 100) / 100;
}

// Full-precision numeric (no 2dp rounding) — used for COF rates so 11.01% / 10.99% survive
// instead of all collapsing to 11.00%.
function numHi(v) {
  if (v === null || v === undefined || isNaN(v)) return 0;
  return Math.round(Number(v) * 1e8) / 1e8;
}

// Format an ISO date (YYYY-MM-DD) as DD-Mmm-YYYY without timezone drift.
function fmtDateDMY(iso) {
  const parts = String(iso || '').split('-');
  if (parts.length < 3) return String(iso || '');
  const [y, m, d] = parts.map(Number);
  const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][m - 1] || '';
  return `${String(d).padStart(2, '0')}-${mon}-${y}`;
}

export function downloadScheduleAsExcel(filename, schedule, meta = {}) {
  const aoa = buildScheduleAoA(schedule, meta);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // Center every cell horizontally + vertically.
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell) cell.s = { ...(cell.s || {}), alignment: { horizontal: 'center', vertical: 'center' } };
    }
  }
  ws['!cols'] = (aoa[aoa.length - 1] || []).map(() => ({ wch: 16 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Amortization');
  XLSX.writeFile(wb, filename, { cellStyles: true });
}

// AoA: a small header (meta.header key/value pairs) then the schedule + a Total row.
function buildScheduleAoA(schedule, meta) {
  const aoa = [];
  if (meta.header) {
    Object.entries(meta.header).forEach(([k, v]) => aoa.push([k, v]));
    aoa.push([]);
  }
  const rows = schedule.rows || [];
  const hasDate = rows.length && rows[0].date !== undefined;
  const hasIDP = rows.some(r => (r.idpReceivable || 0) > 0);
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  if (hasIDP) headers.push('Accrued Interest');
  aoa.push(headers);

  let tInst = 0, tInt = 0, tPrin = 0, tAcc = 0;
  for (const r of rows) {
    const row = [r.sl];
    if (hasDate) row.push(r.date || '');
    row.push(num(r.installment), num(r.interest), num(r.principal), num(r.urpa));
    if (hasIDP) row.push(num(r.idpReceivable || 0));
    aoa.push(row);
    tInst += Number(r.installment) || 0; tInt += Number(r.interest) || 0;
    tPrin += Number(r.principal) || 0; tAcc += Number(r.idpReceivable) || 0;
  }
  const total = ['Total', ...(hasDate ? [''] : []), num(tInst), num(tInt), num(tPrin), ''];
  if (hasIDP) total.push(num(tAcc));
  aoa.push(total);
  return aoa;
}

// Word
export async function downloadScheduleAsWord(filename, schedule, meta = {}) {
  if (typeof docx === 'undefined') { alert('Word export library not loaded.'); return; }
  const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, TextRun, AlignmentType, VerticalAlign } = docx;
  const rows = schedule.rows || [];
  const hasDate = rows[0]?.date !== undefined;
  const hasIDP = rows.some(r => (r.idpReceivable || 0) > 0);
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  if (hasIDP) headers.push('Accrued Interest');

  // Cell centered both horizontally and vertically.
  const cell = (text, bold = false) => new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(text), bold })] })],
  });
  const headerRow = new TableRow({ tableHeader: true, children: headers.map(h => cell(h, true)) });
  let tInst = 0, tInt = 0, tPrin = 0, tAcc = 0;
  const dataRows = rows.map(r => {
    const cells = [r.sl];
    if (hasDate) cells.push(r.date || '');
    cells.push(fmtM(r.installment), fmtM(r.interest), fmtM(r.principal), fmtM(r.urpa));
    if (hasIDP) cells.push(fmtM(r.idpReceivable || 0));
    tInst += Number(r.installment) || 0; tInt += Number(r.interest) || 0;
    tPrin += Number(r.principal) || 0; tAcc += Number(r.idpReceivable) || 0;
    return new TableRow({ children: cells.map(c => cell(c)) });
  });
  const totalCells = ['Total', ...(hasDate ? [''] : []), fmtM(tInst), fmtM(tInt), fmtM(tPrin), ''];
  if (hasIDP) totalCells.push(fmtM(tAcc));
  const totalRow = new TableRow({ children: totalCells.map(c => cell(c, true)) });

  const children = [];
  if (meta.header) {
    Object.entries(meta.header).forEach(([k, v]) =>
      children.push(new Paragraph({ children: [new TextRun({ text: `${k}: `, bold: true }), new TextRun(String(v))] })));
    children.push(new Paragraph(''));
  }
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows, totalRow] }));

  const doc = new Document({ sections: [{ properties: {}, children }] });
  const blob = await Packer.toBlob(doc);
  saveBlob(blob, filename);
}

// PDF (schedule)
export function downloadScheduleAsPDF(filename, schedule, meta = {}) {
  const { jsPDF } = window.jspdf || {};
  if (!jsPDF) { alert('PDF export library not loaded.'); return; }
  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
  let y = 40;
  if (meta.header) {
    doc.setFontSize(10);
    Object.entries(meta.header).forEach(([k, v]) => { doc.text(`${k}: ${v}`, 40, y); y += 14; });
    y += 6;
  }
  const rows = schedule.rows || [];
  const hasDate = rows[0]?.date !== undefined;
  const hasIDP = rows.some(r => (r.idpReceivable || 0) > 0);
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  if (hasIDP) headers.push('Accrued Interest');
  let tInst = 0, tInt = 0, tPrin = 0, tAcc = 0;
  const body = rows.map(r => {
    const row = [String(r.sl)];
    if (hasDate) row.push(r.date || '');
    row.push(fmtM(r.installment), fmtM(r.interest), fmtM(r.principal), fmtM(r.urpa));
    if (hasIDP) row.push(fmtM(r.idpReceivable || 0));
    tInst += Number(r.installment) || 0; tInt += Number(r.interest) || 0;
    tPrin += Number(r.principal) || 0; tAcc += Number(r.idpReceivable) || 0;
    return row;
  });
  const foot = ['Total', ...(hasDate ? [''] : []), fmtM(tInst), fmtM(tInt), fmtM(tPrin), ''];
  if (hasIDP) foot.push(fmtM(tAcc));
  doc.autoTable({
    head: [headers], body, foot: [foot], startY: y,
    showFoot: 'lastPage',
    styles: { fontSize: 8, cellPadding: 3, halign: 'center', valign: 'middle' },
    headStyles: { fillColor: [37, 70, 224], textColor: 255, halign: 'center', valign: 'middle' },
    footStyles: { fillColor: [230, 235, 255], textColor: 20, fontStyle: 'bold', halign: 'center' },
    margin: { left: 40, right: 40 },
  });
  doc.save(filename);
}

// =====================================================================
// Verify Calculation Excel — exactly mirrors "Sample format for verification.xlsx"
// Styling: navy title, green section headers w/ white bold text, accounting/percent
// number formats matching sample, formulas linking Results -> Inputs + Schedule.
// =====================================================================
const FMT = {
  ACCOUNTING: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  ACCOUNTING_INT: '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
  PCT2: '0.00%',
  PCT4: '0.0000%',
};
const WHITE_BOLD = { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 };
const STYLE = {
  title: {
    font: { ...WHITE_BOLD, sz: 12 },
    fill: { fgColor: { rgb: '002060' } },
    alignment: { horizontal: 'center', vertical: 'center' },
  },
  greenHeader: {
    font: WHITE_BOLD,
    fill: { fgColor: { rgb: '00B050' } },
    alignment: { horizontal: 'center', vertical: 'center' },
  },
  cellCenter: { alignment: { horizontal: 'center' } },
};

// Rate Revision verification look (matches the rectified file): bigger navy title, green
// banners/headers with DARK text. Number formats are the plain red-negative variants used on
// the Inputs/Results/Layers sheets (the Schedule keeps the accounting format above).
const RR_FMT = {
  NUM2: '#,##0.00_);[Red](#,##0.00)',
  NUM0: '#,##0_);[Red](#,##0)',
  INT0: '0_);[Red](0)',
  PCT2: '0.00%',
  PCT4: '0.0000%',
  DATE_LONG: 'dd-mmm-yyyy',
  DATE_SHORT: '[$-409]d-mmm-yy;@',
};
const RR_STYLE = {
  title: { font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 16 }, fill: { fgColor: { rgb: '002060' } }, alignment: { horizontal: 'center', vertical: 'center' } },
  subtitle: { font: { bold: true, color: { rgb: '000000' }, sz: 11 }, fill: { fgColor: { rgb: '00B050' } }, alignment: { horizontal: 'center', vertical: 'center' } },
  banner: { font: { bold: true, color: { rgb: '000000' }, sz: 14 }, fill: { fgColor: { rgb: '00B050' } }, alignment: { horizontal: 'center', vertical: 'center' } },
  colHead: { font: { bold: true, color: { rgb: '000000' }, sz: 11 }, fill: { fgColor: { rgb: '00B050' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } },
  cell: { alignment: { horizontal: 'center', vertical: 'center' } },
};

// Build a SheetJS cell object. Pass exactly one of `value` (literal) or `f` (formula).
function setCell(ws, addr, value, opts = {}) {
  const cell = opts.f
    ? { t: 'n', f: opts.f }
    : opts.text
      ? { t: 's', v: value }
      : { t: 'n', v: value };
  if (opts.z) cell.z = opts.z;
  if (opts.s) cell.s = opts.s;
  ws[addr] = cell;
}

export function downloadVerificationExcel(filename, ctx) {
  // Rate Revision — Structured has its own column layout (Date, Total COF, CS Balance, CS Benefit).
  if (ctx.pageType === 'revisionStructured') return downloadRevisionStructuredVerify(filename, ctx);

  const wb = XLSX.utils.book_new();
  const { schedule, inputs, metrics, pageTitle, pageType, params = {} } = ctx;
  const cof = params.cofRate || 0;
  const mora = params.moratoriumMonths || 0;
  // When moratorium interest is capitalized, the outstanding grows during the moratorium.
  // For those rows (and the EPI-principal / funded-security cells that depend on the grown
  // principal) the verify sheet uses the exact engine value so it matches the app.
  const capitalize = (params.capFlags || []).some(Boolean);

  // ----- Schedule sheet (built first so we know row counts)
  // Each row uses cell-referenced formulas where possible so users can trace every calculation.
  // Column layout: A=Sl, B=Installment, C=Interest, D=Principal, E=URPA, F=Int.Expense, G=Accrued.
  // Excel row r corresponds to schedule index i = r - 2 (header at row 1, sl 0 at row 2).
  //
  // Formulas (sheet "Inputs & Results" prefix omitted in code; resolved via inputsSheetRef):
  //   Row 2 (sl 0): E2 = Inputs!B<loan>; everything else 0
  //   Row r (>=3):  C = E_prev * Rate / 12   (or /4 for EQI/Q-EPI payment months only)
  //                 D = B - C                (or = E_prev on the final row to settle remaining)
  //                 E = E_prev - D
  //                 F = E_prev * COF / 12    (always non-zero from sl=1 through final)
  //   Where the page has no single Rate/COF input (Rate Revision Customized), the formula
  //   falls back to a hardcoded value so the column still totals correctly.
  const schedHeaders = ['Sl.', 'Installment', 'Interest', 'Principal', 'URPA', 'Int. Expense (URPA*COF/12)', 'Accrued Interest',
                        'NIM (Yield to Maturity)', 'ERR (Yield to Maturity)'];
  const wsSched = {};
  schedHeaders.forEach((h, c) => setCell(wsSched, XLSX.utils.encode_cell({ r: 0, c }), h, { text: true, s: STYLE.greenHeader }));

  // Pre-compute input-cell references on the Inputs & Results sheet so the Schedule can quote them
  const inputsForRefs = { ...inputs, _derived: {
    derivedSecurityAmount: metrics.derivedSecurityAmount,
    derivedSecurityRate: metrics.derivedSecurityRate,
  }};
  const _inputLines = collectInputLinesFor(pageType, inputsForRefs);
  const _idx = inputIndex(_inputLines, 5); // input rows start at Excel row 5
  const inputsSheet = `'Inputs & Results'!`;
  const RATE_REF = _idx.offeredRate ? `${inputsSheet}B${_idx.offeredRate}` : null;
  const LOAN_REF = _idx.loanAmount ? `${inputsSheet}B${_idx.loanAmount}` : null;
  const COF_REF = _idx.totalCof ? `${inputsSheet}B${_idx.totalCof}` : null;
  const TENOR_REF = _idx.loanTenor ? `${inputsSheet}B${_idx.loanTenor}` : null;
  const MORA_REF = _idx.moratoriumPeriod ? `${inputsSheet}B${_idx.moratoriumPeriod}` : null;
  // Equal-Principal pages can express principal as a constant formula (Loan / number-of-periods).
  // Only valid on the Structured page where the whole regular tenor is one EPI stream.
  const epiCapable = pageType === 'regular' && LOAN_REF && TENOR_REF && MORA_REF;
  const regularMonthsExpr = TENOR_REF && MORA_REF ? `(${TENOR_REF}-${MORA_REF})` : null;

  // Customized Equal-Principal layers: principal is constant within a layer =
  // (layer-start balance) / (periods remaining to MATURITY). The layer-start balance and
  // its Sl are real Schedule cells, so the formula is fully traceable:
  //   Monthly:   D = E{startRow} / (Tenor - A{startRow})
  //   Quarterly: D = E{startRow} / ((Tenor - A{startRow}) / 3)
  // where startRow = Excel row of sl=(from-1) = from+1, A{startRow}=that Sl, E{startRow}=its URPA.
  const epiLayerBySl = {};
  if (pageType === 'customized' && Array.isArray(inputs.paymentLayers) && TENOR_REF) {
    inputs.paymentLayers.forEach((L) => {
      const isQ = L.paymentType === 'Equal Principal + Interest (Quarterly)';
      const isM = L.paymentType === 'Equal Principal + Interest (Monthly)';
      if (!isQ && !isM) return;
      const from = Number(L.fromInstallment), to = Number(L.toInstallment);
      if (!from || !to) return;
      const startRow = from + 1; // Excel row of sl=(from-1)
      const startCell = `E${startRow}`;
      const slCell = `A${startRow}`;
      const denom = isQ ? `((${TENOR_REF}-${slCell})/3)` : `(${TENOR_REF}-${slCell})`;
      const formula = `${startCell}/${denom}`;
      for (let m = from; m <= to; m++) epiLayerBySl[m] = { formula };
    });
  }

  schedule.rows.forEach((r, i) => {
    const xr = i + 2; // 1-indexed Excel row
    const pr = xr - 1; // previous-row Excel index
    setCell(wsSched, `A${xr}`, r.sl, { s: STYLE.cellCenter });

    if (i === 0) {
      // Disbursement row — URPA is the loan amount (formula if available)
      setCell(wsSched, `B${xr}`, 0, { z: FMT.ACCOUNTING });
      setCell(wsSched, `C${xr}`, 0, { z: FMT.ACCOUNTING });
      setCell(wsSched, `D${xr}`, 0, { z: FMT.ACCOUNTING });
      if (LOAN_REF) setCell(wsSched, `E${xr}`, 0, { f: LOAN_REF, z: FMT.ACCOUNTING });
      else setCell(wsSched, `E${xr}`, num(r.urpa), { z: FMT.ACCOUNTING });
      setCell(wsSched, `F${xr}`, 0, { z: FMT.ACCOUNTING });
      setCell(wsSched, `G${xr}`, num(r.idpReceivable || 0), { z: FMT.ACCOUNTING });
      return;
    }

    const isPaymentRow = (r.installment || 0) > 0;
    const isQuarterly = r.paymentType === 'EQI' || r.paymentType === 'Equal Principal + Interest (Quarterly)' || r.paymentType === 'Customized Principal (Quarterly)';
    const isEPI = r.paymentType === 'Equal Principal + Interest (Monthly)' || r.paymentType === 'Equal Principal + Interest (Quarterly)';
    const divisor = (isQuarterly && r.interest > 0) ? 4 : 12;

    // Interest (C) = URPA_prev * Rate / (12 or 4). Always a formula when a rate input exists.
    // Maturity stub (quarterly grid short of maturity) accrues nominal months: Rate*(1 or 2)/12.
    if (RATE_REF && r.interest > 0) {
      const intFormula = r.stubMonths
        ? `E${pr}*${RATE_REF}${r.stubMonths === 2 ? '*2' : ''}/12`
        : `E${pr}*${RATE_REF}/${divisor}`;
      setCell(wsSched, `C${xr}`, 0, { f: intFormula, z: FMT.ACCOUNTING });
    } else {
      setCell(wsSched, `C${xr}`, num(r.interest), { z: FMT.ACCOUNTING });
    }

    // Principal (D) — independent of accrued interest (the previous B-C formula leaked accrued
    // into principal). Equal-Principal: constant = balance / number-of-payment-periods.
    const epiLayer = epiLayerBySl[r.sl];
    if (!isPaymentRow || !(r.principal > 0)) {
      setCell(wsSched, `D${xr}`, num(r.principal || 0), { z: FMT.ACCOUNTING });
    } else if (isEPI && epiCapable && !capitalize) {
      // Structured: one EPI stream over the whole regular tenor.
      // Monthly: Loan / regularMonths ; Quarterly: Loan / ROUNDUP(regularMonths/3) — the
      // cover-based payment count (a partial final quarter still needs a payment).
      const denom = isQuarterly ? `ROUNDUP(${regularMonthsExpr}/3,0)` : regularMonthsExpr;
      setCell(wsSched, `D${xr}`, 0, { f: `${LOAN_REF}/${denom}`, z: FMT.ACCOUNTING });
    } else if (isEPI && epiLayer && !capitalize) {
      // Customized EPI layer: constant = (layer-start URPA cell) / (periods remaining to maturity)
      setCell(wsSched, `D${xr}`, 0, { f: epiLayer.formula, z: FMT.ACCOUNTING });
    } else {
      // EMI / EQI / Custom Principal — use the engine's (already accrued-free) base principal.
      setCell(wsSched, `D${xr}`, num(r.principal), { z: FMT.ACCOUNTING });
    }

    // Installment (B) = Interest + Principal + Accrued-carried-from-previous-row.
    // This is the relationship that actually holds and avoids any accrued leak.
    if (isPaymentRow) {
      setCell(wsSched, `B${xr}`, 0, { f: `C${xr}+D${xr}+G${pr}`, z: FMT.ACCOUNTING });
    } else {
      setCell(wsSched, `B${xr}`, num(r.installment || 0), { z: FMT.ACCOUNTING });
    }

    // URPA (E) = URPA_prev - Principal. Under capitalization the outstanding grows during
    // the moratorium, so use the engine value on those rows (guaranteed to match the app).
    if (capitalize && r.sl >= 1 && r.sl <= mora) {
      setCell(wsSched, `E${xr}`, num(r.urpa), { z: FMT.ACCOUNTING });
    } else {
      setCell(wsSched, `E${xr}`, 0, { f: `E${pr}-D${xr}`, z: FMT.ACCOUNTING });
    }

    // Int Expense (F) = URPA_prev * COF / 12 — accrues every month sl=1..N
    if (COF_REF) {
      setCell(wsSched, `F${xr}`, 0, { f: `E${pr}*${COF_REF}/12`, z: FMT.ACCOUNTING });
    } else {
      setCell(wsSched, `F${xr}`, num(r.interestExpense || 0), { z: FMT.ACCOUNTING });
    }

    // Accrued Interest (G) — the running unpaid-interest balance from the engine
    setCell(wsSched, `G${xr}`, num(r.idpReceivable || 0), { z: FMT.ACCOUNTING });
  });
  const lastDataRow = schedule.rows.length + 1; // 1-indexed last data row
  const totalRowIdx = schedule.rows.length + 1; // 0-indexed for setCell
  // TOTAL row (skip col E — URPA isn't summed; G — accrued shown as-is)
  setCell(wsSched, XLSX.utils.encode_cell({ r: totalRowIdx, c: 0 }), 'TOTAL', { text: true, s: { font: { bold: true }, alignment: { horizontal: 'center' } } });
  ['B', 'C', 'D', null, 'F'].forEach((col, k) => {
    if (!col) return;
    setCell(wsSched, XLSX.utils.encode_cell({ r: totalRowIdx, c: k + 1 }), 0,
      { f: `SUM(${col}2:${col}${lastDataRow})`, z: FMT.ACCOUNTING, s: { font: { bold: true } } });
  });
  // NIM / ERR (Yield to Maturity): standing right AFTER each row's payment, the annualized
  // NIM and ERR over the months remaining to maturity. Sl 0 reproduces the headline figures;
  // the final row is left blank (nothing remains).
  if (pageType === 'revisionCustomized') {
    // Uploaded schedules run on arbitrary dates — compute day-count values (same method as
    // the app's metrics): per period, expense = URPA_prev * COF * days/365, benefit =
    // (COF - secRate) * secAmount * days/365, avg portfolio = exposure-weighted by days.
    const rws = schedule.rows;
    const cofData = params.cofData || [];
    const secLayers = params.securityLayers || [];
    const getCofOn = (ds) => { let r0 = 0; for (const e of cofData) { if (e.date <= ds) r0 = e.rate; else break; } return r0; };
    const getSecOn = (ds) => {
      // Incremental layers: cumulative amount active by the date; amount-weighted average rate.
      let cum = 0, weighted = 0;
      for (const l of secLayers) if (l.fromDate && l.fromDate <= ds) { const a = l.amount || 0; cum += a; weighted += a * (l.activeRate || 0); }
      return { amount: cum, rate: cum > 0 ? weighted / cum : 0 };
    };
    for (let i = 0; i < rws.length - 1; i++) {
      let inc = 0, exp = 0, ben = 0, exposure = 0;
      for (let j = i + 1; j < rws.length; j++) {
        const days = Math.max(1, (new Date(rws[j].date) - new Date(rws[j - 1].date)) / 86400000);
        const cof = getCofOn(rws[j - 1].date);
        const sec = getSecOn(rws[j - 1].date);
        inc += rws[j].interest || 0;
        exp += (rws[j - 1].urpa || 0) * cof * (days / 365);
        ben += (cof - sec.rate) * sec.amount * (days / 365);
        exposure += (rws[j - 1].urpa || 0) * days;
      }
      const totalDays = (new Date(rws[rws.length - 1].date) - new Date(rws[i].date)) / 86400000;
      const avgPort = totalDays > 0 ? exposure / totalDays : 0;
      const years = totalDays / 365;
      const effCof = avgPort > 0 && years > 0 ? exp / avgPort / years : 0;
      const nim = avgPort > 0 && years > 0 ? (inc + ben - exp) / avgPort / years : 0;
      setCell(wsSched, `H${i + 2}`, nim, { z: FMT.PCT4 });
      setCell(wsSched, `I${i + 2}`, nim + effCof, { z: FMT.PCT4 });
    }
  } else {
    // Formula columns. Window rows (xr+1..last): income C, expense F; Loan Security Benefit
    // = monthly benefit x remaining months (constant-rate modules); avg portfolio =
    // AVERAGE(E xr..last-1); years = remainingMonths/12.
    const benefitMonthly = (_idx.csAmount && _idx.totalCof && _idx.csRate)
      ? `(${inputsSheet}B${_idx.csAmount}*(${inputsSheet}B${_idx.totalCof}-${inputsSheet}B${_idx.csRate})/12)` : null;
    for (let xr = 2; xr < lastDataRow; xr++) {
      const n = lastDataRow - xr; // remaining months
      const benTerm = benefitMonthly ? `+${benefitMonthly}*${n}` : '';
      const base = `AVERAGE(E${xr}:E${lastDataRow - 1})/(${n}/12)`;
      setCell(wsSched, `H${xr}`, 0, { f: `(SUM(C${xr + 1}:C${lastDataRow})${benTerm}-SUM(F${xr + 1}:F${lastDataRow}))/${base}`, z: FMT.PCT4 });
      setCell(wsSched, `I${xr}`, 0, { f: `H${xr}+SUM(F${xr + 1}:F${lastDataRow})/${base}`, z: FMT.PCT4 });
    }
  }

  wsSched['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: totalRowIdx, c: 8 } });
  wsSched['!cols'] = schedHeaders.map((h, c) => ({ wch: c >= 7 ? 20 : 18 }));

  // ----- Inputs & Results sheet
  const wsInputs = {};
  const inputsForExport = { ...inputs, _derived: {
    derivedSecurityAmount: metrics.derivedSecurityAmount,
    derivedSecurityRate: metrics.derivedSecurityRate,
  }};
  const inputLines = collectInputLinesFor(pageType, inputsForExport);

  // Title rows (A1: navy, A2: green)
  setCell(wsInputs, 'A1', 'ERR Calculator – Verification', { text: true, s: STYLE.title });
  setCell(wsInputs, 'A2', pageTitle, { text: true, s: STYLE.greenHeader });
  // Section header row 4
  setCell(wsInputs, 'A4', 'Inputs', { text: true, s: STYLE.greenHeader });
  setCell(wsInputs, 'B4', 'Values', { text: true, s: STYLE.greenHeader });

  // Input rows starting at row 5
  const inputStartRow = 5;
  const idx = inputIndex(inputLines, inputStartRow);
  inputLines.forEach(([label, value], i) => {
    const r = inputStartRow + i;
    setCell(wsInputs, `A${r}`, label, { text: true });
    const labelLower = label.toLowerCase();
    const isPct = ['offered rate', 'total cof', 'cash security / fdr rate'].some(k => labelLower.includes(k));
    const isMoney = ['loan amount', 'cash security / fdr amount', 'initial loan amount'].some(k => labelLower.includes(k));
    const isInt = ['moratorium period', 'loan tenor', 'number of installments'].some(k => labelLower.includes(k));
    if (typeof value === 'number') {
      if (isPct) setCell(wsInputs, `B${r}`, value, { z: FMT.PCT2 });
      else if (isMoney) setCell(wsInputs, `B${r}`, value, { z: FMT.ACCOUNTING });
      else if (isInt) setCell(wsInputs, `B${r}`, value, { z: FMT.ACCOUNTING_INT });
      else setCell(wsInputs, `B${r}`, value);
    } else {
      setCell(wsInputs, `B${r}`, String(value || ''), { text: true });
    }
  });

  // Funded Security "EMI/EQI after Moratorium": write the security amount as a traced PMT
  // formula (matches sample B14 = PMT($B$5/12, B9-B8, -B6,,0)).
  const secKind = String(inputs.fundedSecurityType || '');
  if (idx.csAmount && (secKind.startsWith('EMI') || secKind.startsWith('EQI'))) {
    if (capitalize) {
      // Sized on the capitalized (grown) principal — write the derived amount directly.
      setCell(wsInputs, `B${idx.csAmount}`, num(metrics.derivedSecurityAmount || 0), { z: FMT.ACCOUNTING });
    } else if (RATE_REF && LOAN_REF && TENOR_REF && MORA_REF) {
      const nMul = idx.numInst ? `*B${idx.numInst}` : '';
      const pmt = secKind.startsWith('EQI')
        ? `PMT(${RATE_REF}/4,(${TENOR_REF}-${MORA_REF})/3,-${LOAN_REF},,0)`
        : `PMT(${RATE_REF}/12,${TENOR_REF}-${MORA_REF},-${LOAN_REF},,0)`;
      setCell(wsInputs, `B${idx.csAmount}`, 0, { f: `${pmt}${nMul}`, z: FMT.ACCOUNTING });
    }
  }

  // Results section
  const resHeaderRow = inputStartRow + inputLines.length + 1;
  setCell(wsInputs, `A${resHeaderRow}`, 'Results', { text: true, s: STYLE.greenHeader });

  const totalRowExcel = totalRowIdx + 1; // 1-indexed
  const tenor = idx.tenorMonths || idx.loanTenor;
  const cofRow = idx.totalCof;
  const csAmt = idx.csAmount;
  const csRate = idx.csRate;

  // Declarative result rows. Each row's formula can reference earlier rows via the `ref` lookup.
  // null formula -> show 0 with no formula (used when prerequisite inputs are missing).
  const csBenefitF = (csAmt && cofRow && tenor && csRate)
    ? `B${csAmt}*(B${cofRow}-B${csRate})*B${tenor}/12` : null;
  const nimF = tenor ? (ref) =>
    `IF(B${ref.avgPortfolio}*(B${tenor}/12)=0,0,B${ref.netII}/B${ref.avgPortfolio}/(B${tenor}/12))` : null;
  const errF = cofRow ? (ref) => `B${cofRow}+B${ref.nim}` : null;

  const resultRows = [
    { key: 'totalIntReceived', label: 'Total Interest Received', f: `Schedule!C${totalRowExcel}`,         z: FMT.ACCOUNTING },
    { key: 'totalIntExpense',  label: 'Total Interest Expense',  f: `Schedule!F${totalRowExcel}`,         z: FMT.ACCOUNTING },
    { key: 'csBenefit',        label: 'Loan Security Benefit',   f: csBenefitF,                            z: FMT.ACCOUNTING },
    { key: 'netII',            label: 'Net Interest Income',     f: (ref) => `B${ref.totalIntReceived}+B${ref.csBenefit}-B${ref.totalIntExpense}`, z: FMT.ACCOUNTING },
    // Avg Portfolio averages the START-of-month balances (sl 0 .. N-1), EXCLUDING the final
    // row whose URPA is 0 after the last payment. Matches computeMetrics (rows.slice(0,-1))
    // and the original "Sample format for verification" (E2:E61 for a 60-mo loan, not E62).
    { key: 'avgPortfolio',     label: 'Avg Portfolio',           f: `AVERAGE(Schedule!E2:E${lastDataRow - 1})`, z: FMT.ACCOUNTING },
    { key: 'nim',              label: 'NIM',                     f: nimF,                                  z: FMT.PCT4 },
    { key: 'err',              label: 'Effective Rate (ERR)',    f: errF,                                  z: FMT.PCT4 },
  ];
  const ref = {};
  resultRows.forEach((row, i) => {
    const r = resHeaderRow + 1 + i;
    ref[row.key] = r;
    setCell(wsInputs, `A${r}`, row.label, { text: true });
    const f = typeof row.f === 'function' ? row.f(ref) : row.f;
    setCell(wsInputs, `B${r}`, 0, f ? { f, z: row.z } : { z: row.z });
  });
  let lastRow = resHeaderRow + resultRows.length;

  // ===========================================================
  // Yearly Summary table (mirrors "from Fin" L8..L15 methodology)
  // Year | Income | Expense | LS Benefit | Net Int Income | Avg Portfolio | NIM
  // ===========================================================
  const totalMonths = schedule.rows.length - 1;
  const years = Math.ceil(totalMonths / 12);
  const ySummaryHeaderRow = lastRow + 2;
  setCell(wsInputs, `A${ySummaryHeaderRow}`, 'Yearly Summary', { text: true, s: STYLE.greenHeader });
  const yColHeaderRow = ySummaryHeaderRow + 1;
  const ySumHeaders = ['Year', 'Interest Income', 'Interest Expense', 'Loan Security Benefit',
                       'Net Interest Income', 'Average Portfolio', 'NIM', 'YoY ERR'];
  ySumHeaders.forEach((h, c) => setCell(wsInputs, XLSX.utils.encode_cell({ r: yColHeaderRow - 1, c }),
    h, { text: true, s: STYLE.greenHeader }));

  // For each year y (1..years), compute month range and write formulas
  for (let y = 1; y <= years; y++) {
    const yRow = yColHeaderRow + y; // 1-indexed Excel row
    const firstSl = (y - 1) * 12 + 1;
    const lastSl = Math.min(y * 12, totalMonths);
    // Schedule rows: sl m is at Excel row m+2
    const schedFromExcel = firstSl + 2;
    const schedToExcel = lastSl + 2;
    // For Avg Portfolio: average URPA at start of each month in year y
    // URPA at start of month m is rows[m-1].urpa → Excel row m+1
    const avgPortFromExcel = firstSl + 1; // sl (firstSl - 1) row = Excel row firstSl+1
    const avgPortToExcel = lastSl + 1;

    setCell(wsInputs, `A${yRow}`, `Year ${String(y).padStart(2, '0')}`, { text: true, s: STYLE.cellCenter });
    setCell(wsInputs, `B${yRow}`, 0, { f: `SUM(Schedule!C${schedFromExcel}:C${schedToExcel})`, z: FMT.ACCOUNTING });
    setCell(wsInputs, `C${yRow}`, 0, { f: `SUM(Schedule!F${schedFromExcel}:F${schedToExcel})`, z: FMT.ACCOUNTING });
    // Loan Security Benefit per year: CS_Amount * (COF - CS_Rate). Constant per year.
    if (csAmt && cofRow && csRate) {
      setCell(wsInputs, `D${yRow}`, 0, { f: `B${csAmt}*(B${cofRow}-B${csRate})`, z: FMT.ACCOUNTING });
    } else {
      setCell(wsInputs, `D${yRow}`, 0, { z: FMT.ACCOUNTING });
    }
    // Net Interest Income = Income + LS Benefit - Expense
    setCell(wsInputs, `E${yRow}`, 0, { f: `B${yRow}+D${yRow}-C${yRow}`, z: FMT.ACCOUNTING });
    // Avg Portfolio = AVERAGE of URPA across months in the year
    setCell(wsInputs, `F${yRow}`, 0, { f: `AVERAGE(Schedule!E${avgPortFromExcel}:E${avgPortToExcel})`, z: FMT.ACCOUNTING });
    // A final stub year covers fewer than 12 months, so both legs are annualised by
    // 12/months — otherwise the last year reads a raw period figure (a 2-month year
    // showing 2.5% instead of ~15%). Full years scale by 1 and are unchanged.
    const yMonths = lastSl - firstSl + 1;
    // NIM = NetII / AvgPort, annualised
    setCell(wsInputs, `G${yRow}`, 0, { f: `IF(F${yRow}=0,0,E${yRow}/F${yRow}*12/${yMonths})`, z: FMT.PCT4 });
    // YoY ERR = (yearly Interest Expense / yearly Avg Portfolio, annualised) + yearly NIM
    setCell(wsInputs, `H${yRow}`, 0, { f: `(C${yRow}/F${yRow}*12/${yMonths})+G${yRow}`, z: FMT.PCT4 });
  }
  lastRow = yColHeaderRow + years;

  wsInputs['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow - 1, c: 7 } });
  wsInputs['!cols'] = [{ wch: 28 }, { wch: 18 }, { wch: 18 }, { wch: 20 }, { wch: 20 }, { wch: 18 }, { wch: 12 }, { wch: 12 }];
  // Merges: title + page header + Results header + Yearly Summary header (across A:H)
  wsInputs['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } },
    { s: { r: resHeaderRow - 1, c: 0 }, e: { r: resHeaderRow - 1, c: 1 } },
    { s: { r: ySummaryHeaderRow - 1, c: 0 }, e: { r: ySummaryHeaderRow - 1, c: 7 } },
  ];

  XLSX.utils.book_append_sheet(wb, wsInputs, 'Inputs & Results');
  // Rate Revision — Customized: add the Loan Security + COF + ISC layer tables on their own sheet
  // (no Lending Rate table — the rate schedule is already baked into the uploaded amortization).
  if (pageType === 'revisionCustomized') {
    const wsL = buildLayersSheet([], params.securityLayers || [], params.cofData || [], false);
    XLSX.utils.book_append_sheet(wb, wsL, 'Rate_Security_COF_Layers');
  }
  XLSX.utils.book_append_sheet(wb, wsSched, 'Schedule');

  XLSX.writeFile(wb, filename, { cellStyles: true });
}

// =====================================================================
// Verify Calculation Excel — Rate Revision (Structured). Columns:
//   A Sl | B Date | C Installment | D Interest | E Principal | F URPA
//   G Total COF | H Int.Expense | I Accrued | J CS Balance | K CS Benefit
// Mirrors the rectified "Rate_Revision_Structured_*_Rectified Calculation" files.
// =====================================================================
function downloadRevisionStructuredVerify(filename, ctx) {
  const wb = XLSX.utils.book_new();
  const { schedule, inputs, metrics, pageTitle, params = {} } = ctx;
  const rows = schedule.rows;
  const tenor = params.tenorMonths || (rows.length - 1);
  const mora = params.moratoriumMonths || 0;
  const capitalize = (params.capFlags || []).some(Boolean);
  const pct = (v) => `${+(Number(v) * 100).toFixed(6)}%`; // inline percent literal, e.g. 0.14 -> "14%"

  // Layer tables live on the Rate_Security_COF_Layers sheet (rows 3+). The Schedule cell-references
  // them by date so edits to a rate / amount / COF value flow straight into the schedule (point 3).
  const LAYER = 'Rate_Security_COF_Layers';
  const rateLayers = (params.rateLayers || []).filter(l => l.fromDate);
  const securityLayers = (params.securityLayers || []).filter(l => l.fromDate);
  const cofRecs = params.cofData || [];
  const lastIdxLE = (arr, key, d) => { let k = -1; for (let j = 0; j < arr.length; j++) if (arr[j][key] && arr[j][key] <= d) k = j; return k; };
  const lendCell = (d) => `${LAYER}!$E$${3 + Math.max(0, lastIdxLE(rateLayers, 'fromDate', d))}`;     // active lending rate on date d
  const cofCell = (d) => `${LAYER}!$U$${3 + Math.max(0, lastIdxLE(cofRecs, 'date', d))}`;               // Eligible COF on date d
  const secActive = (d) => securityLayers.filter(l => l.fromDate <= d).length;                          // # security layers active by d
  // Most recent active layer wins: its own Amount ($K) is the balance, its Active Rate ($L) the rate.
  const secAmtCell = (d) => { const n = secActive(d); return n ? `${LAYER}!$K$${3 + n - 1}` : null; };
  const secRateCell = (d) => { const n = secActive(d); return n ? `${LAYER}!$L$${3 + n - 1}` : null; };

  // ----- Schedule sheet -----
  const heads = ['Sl.', 'Date', 'Installment', 'Interest', 'Principal', 'URPA',
                 'Total COF', 'Int. Expense (URPA*COF/12)', 'Accrued Interest', 'Loan Security Balance', 'Loan Security Benefit',
                 'NIM (Yield to Maturity)', 'ERR (Yield to Maturity)'];
  const ws = {};
  heads.forEach((h, c) => setCell(ws, XLSX.utils.encode_cell({ r: 0, c }), h, { text: true, s: RR_STYLE.colHead }));

  const RATE = (modality) => (modality === 'EQI' || modality === 'Equal Principal + Interest (Quarterly)') ? 4 : 12;
  const ppyDiv = RATE(params.paymentModality);
  // Annuity (EMI/EQI) => installment via PMT; otherwise Equal Principal => constant principal.
  const isAnnuity = params.paymentModality === 'EMI' || params.paymentModality === 'EQI';
  const moraRow = mora + 2;       // Excel row holding sl = mora (balance at end of moratorium)
  const afterMoraRow = mora + 3;  // Excel row holding sl = mora+1 (first post-moratorium month)

  // Precompute rate-block starts for regular (post-moratorium) payment rows so EMI/EQI can recompute
  // the installment via PMT(blockRate, remainingPeriods, -balanceAtBlockStart) when the rate changes.
  // blockStartSl[sl] = the Sl that began the current rate block.
  const blockStartBySl = {};
  let curRate = null, curBlockStart = null;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.sl <= mora) continue;
    if (r.installment > 0 && ppyDiv === 12) {
      if (r.rate !== curRate) { curRate = r.rate; curBlockStart = r.sl; }
      blockStartBySl[r.sl] = curBlockStart;
    } else if (r.installment > 0) {
      // quarterly — block on rate change too
      if (r.rate !== curRate) { curRate = r.rate; curBlockStart = r.sl; }
      blockStartBySl[r.sl] = curBlockStart;
    }
  }

  // Day-count split interest for a row whose accrual period contains a rate revision on a
  // non-due date: each segment = F(periodStart) * rate/360 * days. Period bounds use cell
  // refs; the revision date(s) are DATE() literals. Mirrors the rectified file's D formula.
  const dateLit = (iso) => { const [y, mo, da] = String(iso).split('-').map(Number); return `DATE(${y},${mo},${da})`; };
  const splitDFormula = (r, xr) => {
    const psXr = (r.splitFromSl || 0) + 2; // Excel row holding sl = splitFromSl (period start)
    const segs = r.splitSegments;
    return segs.map((s, i) => {
      const startRef = i === 0 ? `B${psXr}` : dateLit(s.from);
      const endRef = i === segs.length - 1 ? `B${xr}` : dateLit(s.to);
      return `(F${psXr}*${lendCell(s.from)}/360*DAYS360(${startRef},${endRef}))`;
    }).join('+');
  };

  let eqpPrevPayRow = null; // Excel row of the previous Equal-Principal payment (quarterly chains =E(prev))
  let chainPayXr = null, chainRate = null; // annuity chaining after a goal-sought split payment
  rows.forEach((r, i) => {
    const xr = i + 2, pr = xr - 1;
    const d = new Date(r.date);
    setCell(ws, `A${xr}`, r.sl, { z: FMT.ACCOUNTING_INT, s: RR_STYLE.cell });
    setCell(ws, `B${xr}`, 0, { f: `DATE(${d.getFullYear()},${d.getMonth() + 1},${d.getDate()})`, z: 'dd-mmm-yyyy', s: RR_STYLE.cell });

    if (i === 0) {
      ['C', 'D', 'E'].forEach(c => setCell(ws, `${c}${xr}`, 0, { z: FMT.ACCOUNTING }));
      setCell(ws, `F${xr}`, 0, { f: `'Inputs & Results'!D5`, z: FMT.ACCOUNTING });
      setCell(ws, `G${xr}`, 0, { f: cofCell(r.date), z: FMT.PCT2 });
      setCell(ws, `H${xr}`, 0, { z: FMT.ACCOUNTING });
      setCell(ws, `I${xr}`, num(r.idpReceivable || 0), { z: FMT.ACCOUNTING });
      const sAmt0 = secAmtCell(r.date);
      if (sAmt0) setCell(ws, `J${xr}`, 0, { f: sAmt0, z: FMT.ACCOUNTING });
      else setCell(ws, `J${xr}`, 0, { z: FMT.ACCOUNTING });
      // Loan Security Benefit starts one row later (one-month lag); disbursement row carries none.
      setCell(ws, `K${xr}`, 0, { z: FMT.ACCOUNTING });
      return;
    }

    const isMora = r.sl <= mora;
    const isPaymentRow = (r.installment || 0) > 0;

    // D Interest = URPA_prev * lending rate / (12 or 4); day-count split when a rate
    // revision falls inside the accrual period; maturity stub accrues nominal months.
    const accrualStart = rows[i - 1].date; // accrual period starts on the previous row's date
    if (r.splitSegments) {
      setCell(ws, `D${xr}`, 0, { f: splitDFormula(r, xr), z: FMT.ACCOUNTING });
    } else if (isMora) {
      setCell(ws, `D${xr}`, 0, { f: `F${pr}*${lendCell(accrualStart)}/12`, z: FMT.ACCOUNTING });
    } else if (isPaymentRow && r.stubMonths) {
      setCell(ws, `D${xr}`, 0, { f: `F${pr}*${lendCell(accrualStart)}${r.stubMonths === 2 ? '*2' : ''}/12`, z: FMT.ACCOUNTING });
    } else if (isPaymentRow) {
      setCell(ws, `D${xr}`, 0, { f: `F${pr}*${lendCell(accrualStart)}/${ppyDiv}`, z: FMT.ACCOUNTING });
    } else {
      setCell(ws, `D${xr}`, 0, { z: FMT.ACCOUNTING });
    }

    // C Installment & E Principal
    if (isMora) {
      // paid month: C = D + E + accrued_prev (E=0). unpaid: C=0.
      if (isPaymentRow) setCell(ws, `C${xr}`, 0, { f: `D${xr}+E${xr}+I${pr}`, z: FMT.ACCOUNTING });
      else setCell(ws, `C${xr}`, 0, { z: FMT.ACCOUNTING });
      setCell(ws, `E${xr}`, 0, { z: FMT.ACCOUNTING });
    } else if (isPaymentRow && r.stubMonths) {
      // Maturity stub (quarterly grid short of maturity): pay off the full outstanding
      // balance plus the stub interest (+ any uncollected moratorium accrued).
      setCell(ws, `E${xr}`, 0, { f: `F${pr}`, z: FMT.ACCOUNTING });
      setCell(ws, `C${xr}`, 0, { f: `D${xr}+E${xr}+I${pr}`, z: FMT.ACCOUNTING });
    } else if (isPaymentRow && isAnnuity && r.splitSegments) {
      // Mid-period revision: the installment is the goal-sought constant, written as a VALUE
      // (per the rectified file); unpaid moratorium accrued (I prev) rides on top of it, so
      // principal excludes it: E = C - D - I(prev). Full precision — the whole chain of
      // subsequent =C(prev) installments references this cell, so 2dp would leave a residual.
      setCell(ws, `C${xr}`, numHi(r.installment), { z: FMT.ACCOUNTING });
      setCell(ws, `E${xr}`, 0, { f: `C${xr}-D${xr}-I${pr}`, z: FMT.ACCOUNTING });
      chainPayXr = xr; chainRate = r.rate;
    } else if (isPaymentRow && isAnnuity && chainPayXr !== null && r.rate === chainRate) {
      // Same rate block as the goal-sought payment: installments stay constant, =C(prev payment).
      setCell(ws, `C${xr}`, 0, { f: `C${chainPayXr}`, z: FMT.ACCOUNTING });
      setCell(ws, `E${xr}`, 0, { f: `C${xr}-D${xr}`, z: FMT.ACCOUNTING });
      chainPayXr = xr;
    } else if (isPaymentRow && isAnnuity) {
      // EMI/EQI annuity: installment via PMT (reset per rate block), principal = installment − interest.
      chainPayXr = null; chainRate = null;
      const bStart = blockStartBySl[r.sl];
      const bRow = bStart + 1; // Excel row of sl=(bStart-1) holding balance going into the block
      // Quarterly: payments required to COVER the remaining period (phantom final quarter
      // when the tenor doesn't divide evenly) = ROUNDUP((tenor - blockStartSl + 1 + 2)/3).
      const periodsExpr = ppyDiv === 12
        ? `'Inputs & Results'!$D$10-Schedule!$A$${bRow}`
        : `ROUNDUP(('Inputs & Results'!$D$10-Schedule!$A$${bRow}+2)/3,0)`;
      const pmt = `PMT(${lendCell(accrualStart)}/${ppyDiv},${periodsExpr},-Schedule!$F$${bRow},,0)`;
      setCell(ws, `C${xr}`, 0, { f: `${pmt}+I${pr}`, z: FMT.ACCOUNTING });
      setCell(ws, `E${xr}`, 0, { f: `C${xr}-D${xr}`, z: FMT.ACCOUNTING });
    } else if (isPaymentRow) {
      // Equal Principal: CONSTANT principal sized over the post-moratorium payment count (tenor − mora).
      // Monthly repeats the absolute formula every row; quarterly seeds the first payment then chains
      // =E(prevQuarter). Installment = Interest + Principal + Accrued_prev. Matches rectified EMPP/EQPP.
      if (ppyDiv === 12) {
        setCell(ws, `E${xr}`, 0, { f: `$F$${moraRow}/('Inputs & Results'!$D$10-Schedule!$A$${moraRow})`, z: FMT.ACCOUNTING });
      } else if (eqpPrevPayRow === null) {
        setCell(ws, `E${xr}`, 0, { f: `$F$${afterMoraRow}/('Inputs & Results'!$D$10-Schedule!$A$${moraRow})*3`, z: FMT.ACCOUNTING });
      } else {
        setCell(ws, `E${xr}`, 0, { f: `E${eqpPrevPayRow}`, z: FMT.ACCOUNTING });
      }
      eqpPrevPayRow = xr;
      setCell(ws, `C${xr}`, 0, { f: `D${xr}+E${xr}+I${pr}`, z: FMT.ACCOUNTING });
    } else {
      setCell(ws, `C${xr}`, 0, { z: FMT.ACCOUNTING });
      setCell(ws, `E${xr}`, 0, { z: FMT.ACCOUNTING });
    }

    // F URPA = URPA_prev - Principal. Under capitalization the outstanding grows during
    // the moratorium, so use the engine value on those rows.
    if (capitalize && r.sl >= 1 && r.sl <= mora) {
      setCell(ws, `F${xr}`, num(r.urpa), { z: FMT.ACCOUNTING });
    } else {
      setCell(ws, `F${xr}`, 0, { f: `F${pr}-E${xr}`, z: FMT.ACCOUNTING });
    }
    // G Total COF (Eligible COF referenced from the COF table by date); H Int Expense = prev URPA * prev COF / 12
    setCell(ws, `G${xr}`, 0, { f: cofCell(r.date), z: FMT.PCT2 });
    setCell(ws, `H${xr}`, 0, { f: `F${pr}*G${pr}/12`, z: FMT.ACCOUNTING });
    // I Accrued
    setCell(ws, `I${xr}`, num(r.idpReceivable || 0), { z: FMT.ACCOUNTING });
    // J Loan Security Balance (cumulative amount active by the date; released to 0 at maturity);
    // K Loan Security Benefit = PRIOR row balance*(COF − weighted-avg security rate)/12 (one-month lag)
    const isLast = i === rows.length - 1;
    const sAmt = secAmtCell(r.date);
    if (isLast || !sAmt) setCell(ws, `J${xr}`, 0, { z: FMT.ACCOUNTING });
    else setCell(ws, `J${xr}`, 0, { f: sAmt, z: FMT.ACCOUNTING });
    // K Loan Security Benefit. A security amount taken mid-period (fromDate strictly inside
    // [prev date, this date]) splits the month by DAYS360 — each segment on its own cumulative
    // balance ($M$) and weighted rate ($N$), both vs the period-start COF. Else flat /12.
    const d0s = rows[i - 1].date, d1s = r.date;
    const benCuts = securityLayers.filter(l => l.fromDate > d0s && l.fromDate < d1s).map(l => l.fromDate).sort();
    if (!benCuts.length) {
      const sRatePrev = secRateCell(d0s);
      setCell(ws, `K${xr}`, 0, { f: `J${pr}*(G${pr}-${sRatePrev || '0'})/12`, z: FMT.ACCOUNTING });
    } else {
      const layerJCell = (fromDate) => `${LAYER}!$J$${3 + securityLayers.findIndex(l => l.fromDate === fromDate)}`;
      const bounds = [d0s, ...benCuts, d1s];
      const terms = [];
      for (let k = 0; k < bounds.length - 1; k++) {
        const segStart = bounds[k];
        const bal = k === 0 ? `J${pr}` : (secAmtCell(segStart) || '0');
        const rate = secRateCell(segStart) || '0';
        const startRef = k === 0 ? `B${pr}` : layerJCell(segStart);
        const endRef = k === bounds.length - 2 ? `B${xr}` : layerJCell(bounds[k + 1]);
        terms.push(`${bal}*(G${pr}-${rate})*DAYS360(${startRef},${endRef})/360`);
      }
      setCell(ws, `K${xr}`, 0, { f: terms.join('+'), z: FMT.ACCOUNTING });
    }
  });

  const lastDataRow = rows.length + 1;        // 1-indexed
  const totalRow = rows.length + 2;           // TOTAL row (1-indexed)
  const tIdx = rows.length + 1;               // 0-indexed for encode_cell
  setCell(ws, `A${totalRow}`, 'TOTAL', { text: true, s: { font: { bold: true }, alignment: { horizontal: 'center' } } });
  ['C', 'D', 'E', 'H', 'K'].forEach((col) => {
    setCell(ws, `${col}${totalRow}`, 0, { f: `SUM(${col}2:${col}${lastDataRow})`, z: FMT.ACCOUNTING, s: { font: { bold: true } } });
  });
  // NIM / ERR (Yield to Maturity): standing right AFTER each row's payment, the annualized
  // NIM and ERR over the months remaining to maturity. Window rows (xr+1..last): income D,
  // Loan Security Benefit K, expense H; avg portfolio = start-of-month balances F(xr..last-1);
  // years = remainingMonths/12. Sl 0 therefore reproduces the headline NIM/ERR; the final
  // row is left blank (nothing remains).
  for (let xr = 2; xr < lastDataRow; xr++) {
    const n = lastDataRow - xr; // remaining months
    const base = `AVERAGE(F${xr}:F${lastDataRow - 1})/(${n}/12)`;
    setCell(ws, `L${xr}`, 0, { f: `(SUM(D${xr + 1}:D${lastDataRow})+SUM(K${xr + 1}:K${lastDataRow})-SUM(H${xr + 1}:H${lastDataRow}))/${base}`, z: FMT.PCT4 });
    setCell(ws, `M${xr}`, 0, { f: `L${xr}+SUM(H${xr + 1}:H${lastDataRow})/${base}`, z: FMT.PCT4 });
  }

  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: tIdx, c: 12 } });
  ws['!cols'] = heads.map((h, c) => ({ wch: c === 0 ? 6 : c === 1 ? 12 : c >= 11 ? 20 : 17 }));

  // ----- Inputs & Results sheet (Inputs in A:D, Results in F:H, side by side) -----
  const wsI = {};
  setCell(wsI, 'A1', 'ERR Calculator — Verification', { text: true, s: RR_STYLE.title });
  setCell(wsI, 'A2', pageTitle, { text: true, s: RR_STYLE.subtitle });
  setCell(wsI, 'A4', 'Inputs', { text: true, s: RR_STYLE.banner });
  setCell(wsI, 'D4', 'Values', { text: true, s: RR_STYLE.banner });
  setCell(wsI, 'F4', 'Results', { text: true, s: RR_STYLE.banner });

  const moraYes = (inputs.moratoriumAvail === 'Yes' || inputs.moratoriumAvail === true);
  // 6 inputs (the Lending Rate / Loan Security / COF layer summaries now live on their own sheet)
  const inputRows = [
    ['Initial Loan Amount', inputs.initialAmount ?? 0, 'num'],
    ['Disbursement Date', inputs.disbursementDate ?? '', 'date'],
    ['Moratorium Given at Disbursement?', moraYes ? 'Yes' : 'No', 'text'],
    ['Moratorium Period (Months)', inputs.moratoriumPeriod ?? 0, 'int'],
    [moraYes ? 'Payment Modality after Moratorium Period' : 'Payment Modality', inputs.paymentModality ?? '', 'text'],
    ['Loan Tenor including Moratorium at Disbursement (Months)', inputs.tenorMonths ?? 0, 'int'],
  ];
  inputRows.forEach(([label, value, kind], i) => {
    const r = 5 + i;
    setCell(wsI, `A${r}`, label, { text: true, s: RR_STYLE.cell });
    if (kind === 'date' && value) {
      const dd = new Date(value);
      setCell(wsI, `D${r}`, 0, { f: `DATE(${dd.getFullYear()},${dd.getMonth() + 1},${dd.getDate()})`, z: RR_FMT.DATE_LONG, s: RR_STYLE.cell });
    } else if (kind === 'num') {
      setCell(wsI, `D${r}`, value, { z: RR_FMT.NUM2, s: RR_STYLE.cell });
    } else if (kind === 'int') {
      setCell(wsI, `D${r}`, value, { z: RR_FMT.NUM0, s: RR_STYLE.cell });
    } else {
      setCell(wsI, `D${r}`, String(value || ''), { text: true, s: RR_STYLE.cell });
    }
  });

  // Results occupy F5:H11 — fixed cells, so the formulas can reference them (and D10 tenor) directly.
  const RES = [
    ['Total Interest Received', `Schedule!D${totalRow}`, RR_FMT.NUM2],
    ['Total Interest Expense', `Schedule!H${totalRow}`, RR_FMT.NUM2],
    ['Loan Security Benefit', `Schedule!K${totalRow}`, RR_FMT.NUM2],
    ['Net Interest Income', `H5+H7-H6`, RR_FMT.NUM2],
    ['Avg Portfolio', `AVERAGE(Schedule!F2:F${lastDataRow - 1})`, RR_FMT.NUM2],
    ['NIM', `IF(H9*(D10/12)=0,0,H8/H9/(D10/12))`, RR_FMT.PCT4],
    ['Effective Rate (ERR)', `H10+(H6/H9/D10*12)`, RR_FMT.PCT4],
  ];
  RES.forEach(([label, f, z], i) => {
    const r = 5 + i;
    setCell(wsI, `F${r}`, label, { text: true, s: RR_STYLE.cell });
    setCell(wsI, `H${r}`, 0, { f, z, s: RR_STYLE.cell });
  });

  // ----- Yearly Summary (fixed: banner row 14, headers row 15, data from row 16) -----
  const years = Math.ceil((rows.length - 1) / 12);
  setCell(wsI, 'A14', 'Yearly Summary', { text: true, s: RR_STYLE.banner });
  ['Year', 'Interest Income', 'Interest Expense', 'Loan Security Benefit', 'Net Interest Income', 'Average Portfolio', 'NIM', 'YoY ERR']
    .forEach((h, c) => setCell(wsI, XLSX.utils.encode_cell({ r: 14, c }), h, { text: true, s: RR_STYLE.colHead }));
  for (let y = 1; y <= years; y++) {
    const yr = 15 + y;
    const firstSl = (y - 1) * 12 + 1;
    const lastSl = Math.min(y * 12, rows.length - 1);
    const dFrom = firstSl + 2, dTo = lastSl + 2;             // interest/expense rows (sl)
    const kFrom = firstSl + 2, kTo = lastSl + 2;             // Loan Security Benefit (lagged: Sl s -> row K(s+2))
    const fFrom = firstSl + 1, fTo = lastSl + 1;             // URPA start-of-month
    setCell(wsI, `A${yr}`, `Year ${String(y).padStart(2, '0')}`, { text: true, s: RR_STYLE.cell });
    setCell(wsI, `B${yr}`, 0, { f: `SUM(Schedule!D${dFrom}:D${dTo})`, z: RR_FMT.NUM2, s: RR_STYLE.cell });
    setCell(wsI, `C${yr}`, 0, { f: `SUM(Schedule!H${dFrom}:H${dTo})`, z: RR_FMT.NUM2, s: RR_STYLE.cell });
    setCell(wsI, `D${yr}`, 0, { f: `SUM(Schedule!K${kFrom}:K${kTo})`, z: RR_FMT.NUM2, s: RR_STYLE.cell });
    setCell(wsI, `E${yr}`, 0, { f: `B${yr}+D${yr}-C${yr}`, z: RR_FMT.NUM2, s: RR_STYLE.cell });
    setCell(wsI, `F${yr}`, 0, { f: `AVERAGE(Schedule!F${fFrom}:F${fTo})`, z: RR_FMT.NUM2, s: RR_STYLE.cell });
    // A final stub year covers fewer than 12 months — annualise both legs by 12/months
    // so it is comparable with the full years above it (full years scale by 1).
    const yMonths = lastSl - firstSl + 1;
    setCell(wsI, `G${yr}`, 0, { f: `IF(F${yr}=0,0,E${yr}/F${yr}*12/${yMonths})`, z: RR_FMT.PCT4, s: RR_STYLE.cell });
    // YoY ERR = (yearly Interest Expense / yearly Avg Portfolio, annualised) + yearly NIM
    setCell(wsI, `H${yr}`, 0, { f: `(C${yr}/F${yr}*12/${yMonths})+G${yr}`, z: RR_FMT.PCT4, s: RR_STYLE.cell });
  }
  const lastRow = 15 + years;

  wsI['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow - 1, c: 7 } });
  wsI['!cols'] = [{ wch: 24 }, { wch: 13 }, { wch: 12 }, { wch: 15 }, { wch: 14 }, { wch: 18 }, { wch: 12 }, { wch: 14 }];
  wsI['!rows'] = []; wsI['!rows'][0] = { hpt: 21 }; wsI['!rows'][3] = { hpt: 19.5 }; wsI['!rows'][13] = { hpt: 19.5 }; wsI['!rows'][14] = { hpt: 30.75 };
  const mergeRC = (r1, c1, r2, c2) => ({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
  wsI['!merges'] = [
    mergeRC(0, 0, 0, 7), mergeRC(1, 0, 1, 7),        // title + subtitle
    mergeRC(3, 0, 3, 2), mergeRC(3, 5, 3, 7),        // Inputs banner (A:C), Results banner (F:H)
    mergeRC(13, 0, 13, 7),                            // Yearly Summary banner
  ];
  for (let i = 0; i < inputRows.length; i++) wsI['!merges'].push(mergeRC(4 + i, 0, 4 + i, 2)); // input labels A:C
  for (let i = 0; i < RES.length; i++) wsI['!merges'].push(mergeRC(4 + i, 5, 4 + i, 6));        // result labels F:G

  // ----- Rate_Security_COF_Layers sheet (3 tables: Lending Rate A:E, Loan Security G:O, COF+ISC Q:W) -----
  const wsL = buildLayersSheet(rateLayers, securityLayers, cofRecs);

  XLSX.utils.book_append_sheet(wb, wsI, 'Inputs & Results');
  XLSX.utils.book_append_sheet(wb, wsL, 'Rate_Security_COF_Layers');
  XLSX.utils.book_append_sheet(wb, ws, 'Schedule');
  XLSX.writeFile(wb, filename, { cellStyles: true });
}

// Build the Rate_Security_COF_Layers worksheet. Three side-by-side tables; the Schedule
// cell-references their values, so the row count of each table follows its input (point 5) and
// editing a value flows into the schedule (points 3 & 7). Returns a SheetJS worksheet object.
function buildLayersSheet(rateLayers, securityLayers, cofRecs, includeLending = true) {
  const wsL = {};
  const ymd = (iso) => { const [y, m, d] = String(iso || '').split('-').map(Number); return [y || 0, m || 0, d || 0]; };
  const title = (addr, text) => setCell(wsL, addr, text, { text: true, s: RR_STYLE.title });
  const head = (addr, text) => setCell(wsL, addr, text, { text: true, s: RR_STYLE.colHead });
  const intCell = (addr, v) => setCell(wsL, addr, v, { z: RR_FMT.INT0, s: RR_STYLE.cell });
  const pctCell = (addr, v) => setCell(wsL, addr, v, { z: RR_FMT.PCT2, s: RR_STYLE.cell });
  const numCell = (addr, v) => setCell(wsL, addr, v, { z: RR_FMT.NUM2, s: RR_STYLE.cell });
  const dateCell = (addr, f) => setCell(wsL, addr, 0, { f, z: RR_FMT.DATE_SHORT, s: RR_STYLE.cell });
  const fNum = (addr, f, z) => setCell(wsL, addr, 0, { f, z, s: RR_STYLE.cell });

  // Table 1 — Lending Rate Layers (A:E). Omitted for Rate Revision — Customized, where the rate
  // schedule is already baked into the uploaded amortization (no separate lending-rate layers).
  if (includeLending) {
    title('A1', 'Lending Rate Layers');
    ['Year', 'Month', 'Day', 'From Date', 'Lending Rate'].forEach((h, c) => head(XLSX.utils.encode_cell({ r: 1, c }), h));
    rateLayers.forEach((l, i) => {
      const r = 3 + i; const [y, m, d] = ymd(l.fromDate);
      intCell(`A${r}`, y); intCell(`B${r}`, m); intCell(`C${r}`, d);
      dateCell(`D${r}`, `DATE(A${r},B${r},C${r})`);
      pctCell(`E${r}`, l.activeRate || 0);
    });
  }

  // Table 2 — Loan Security Layers (G:M). The most recent active layer wins: its Amount is the
  // balance in effect (a total, not an increment) and its Active Rate is the rate in effect.
  title('G1', 'Loan Security Layers');
  ['Year', 'Month', 'Day', 'From Date', 'Amount', 'Active Rate', 'Security Identifier']
    .forEach((h, c) => head(XLSX.utils.encode_cell({ r: 1, c: 6 + c }), h));
  securityLayers.forEach((l, i) => {
    const r = 3 + i; const [y, m, d] = ymd(l.fromDate);
    intCell(`G${r}`, y); intCell(`H${r}`, m); intCell(`I${r}`, d);
    dateCell(`J${r}`, `DATE(G${r},H${r},I${r})`);
    numCell(`K${r}`, l.amount || 0);
    pctCell(`L${r}`, l.activeRate || 0);
    setCell(wsL, `M${r}`, `Loan Security ${i + 1}`, { text: true, s: RR_STYLE.cell });
  });

  // Table 3 — COF + ISC Layers (O:U). Year/Month/Day are integers; Eligible COF = MAX(COF,ISC)+0.3%.
  title('O1', 'COF + ISC Layers');
  ['Year', 'Month', 'Day', 'Date', 'COF\n(M-o-M)', 'ISC', 'Eligible COF']
    .forEach((h, c) => head(XLSX.utils.encode_cell({ r: 1, c: 14 + c }), h));
  cofRecs.forEach((rec, i) => {
    const r = 3 + i; const [y, m, d] = ymd(rec.date);
    intCell(`O${r}`, y); intCell(`P${r}`, m); intCell(`Q${r}`, d);
    dateCell(`R${r}`, `DATE(O${r},P${r},Q${r})`);
    const hasRaw = (rec.cof != null && rec.isc != null);
    pctCell(`S${r}`, rec.cof != null ? rec.cof : rec.rate);
    if (rec.isc != null) pctCell(`T${r}`, rec.isc);
    if (hasRaw) fNum(`U${r}`, `MAX(S${r}:T${r})+0.3%`, RR_FMT.PCT2);
    else pctCell(`U${r}`, rec.rate);
  });

  // Table 4 — Loan Security breakdown (W:Y). Each individual security behind the combined rows in
  // G:M (where K = Σ these amounts and L = the amount-weighted-average of these rates). Purely
  // informational — the Schedule never references this table, so the combined table stays intact.
  title('W1', 'Loan Security — Breakdown');
  ['From Date', 'Outstanding Amount', 'Interest Rate'].forEach((h, c) => head(XLSX.utils.encode_cell({ r: 1, c: 22 + c }), h));
  let br = 3;
  securityLayers.forEach((l) => {
    const secs = (l.securities && l.securities.length) ? l.securities : [{ amount: l.amount, rate: l.activeRate }];
    secs.forEach((s) => {
      setCell(wsL, `W${br}`, fmtDateDMY(l.fromDate), { text: true, s: RR_STYLE.cell });
      numCell(`X${br}`, s.amount || 0);
      pctCell(`Y${br}`, s.rate || 0);
      br++;
    });
  });
  const breakdownRows = br - 3;

  const maxRows = Math.max(rateLayers.length, securityLayers.length, cofRecs.length, breakdownRows, 1);
  wsL['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRows + 1, c: 24 } });
  const W = (n) => ({ wch: n });
  wsL['!cols'] = [W(7), W(6), W(6), W(9), W(9), W(3), W(7), W(6), W(6), W(9), W(13), W(8), W(15), W(3), W(7), W(6), W(6), W(10), W(11), W(8), W(10), W(3), W(11), W(15), W(11)];
  wsL['!rows'] = []; wsL['!rows'][0] = { hpt: 21.75 }; wsL['!rows'][1] = { hpt: 28 };
  wsL['!merges'] = [
    { s: { r: 0, c: 6 }, e: { r: 0, c: 12 } },   // G1:M1 (Loan Security)
    { s: { r: 0, c: 14 }, e: { r: 0, c: 20 } },  // O1:U1 (COF + ISC)
    { s: { r: 0, c: 22 }, e: { r: 0, c: 24 } },  // W1:Y1 (Loan Security — Breakdown)
  ];
  if (includeLending) wsL['!merges'].unshift({ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }); // A1:E1 (Lending Rate)
  return wsL;
}

// If security type is installment-based, the input csAmount is empty —
// substitute the model's derived amount stored on metrics (passed as `inp._derived`).
function securityAmtFor(inp, key) {
  return inp._derived && inp._derived[key] !== undefined ? inp._derived[key] : null;
}

// Build list of [label, value] for the Inputs section based on page type
function collectInputLinesFor(pageType, inp) {
  const yesNo = (v) => (v === 'Yes' || v === true ? 'Yes' : 'No');
  if (pageType === 'regular') {
    return [
      ['Offered Rate', inp.offeredRate ?? 0],
      ['Loan Amount', inp.loanAmount ?? 0],
      ['Moratorium Available?', yesNo(inp.moratoriumAvail)],
      ['Moratorium Period (Months)', inp.moratoriumPeriod ?? 0],
      ['Loan Tenor including Moratorium (Months)', inp.loanTenor ?? 0],
      ['Payment Mode', inp.paymentMode ?? ''],
      ['Total COF (COF/ISC + OPEX)', inp.totalCof ?? 0],
      ['Funded Security Type', inp.fundedSecurityType ?? ''],
      ['Number of Installments (security)', inp.numInst ?? 0],
      ['Cash Security / FDR Amount', securityAmtFor(inp, 'derivedSecurityAmount') ?? (inp.csAmount ?? 0)],
      ['Cash Security / FDR Rate', inp.csRate ?? 0],
    ];
  }
  if (pageType === 'customized') {
    return [
      ['Offered Rate', inp.offeredRate ?? 0],
      ['Loan Amount', inp.loanAmount ?? 0],
      ['Moratorium Available?', yesNo(inp.moratoriumAvail)],
      ['Moratorium Period (Months)', inp.moratoriumPeriod ?? 0],
      ['Loan Tenor including Moratorium (Months)', inp.loanTenor ?? 0],
      ['Payment Layers', (inp.paymentLayers || []).length + ' layer(s)'],
      ['Total COF (COF/ISC + OPEX)', inp.totalCof ?? 0],
      ['Funded Security Type', inp.fundedSecurityType ?? ''],
      ['Number of Installments (security)', inp.numInst ?? 0],
      ['Cash Security / FDR Amount', securityAmtFor(inp, 'derivedSecurityAmount') ?? (inp.csAmount ?? 0)],
      ['Cash Security / FDR Rate', inp.csRate ?? 0],
    ];
  }
  if (pageType === 'revisionStructured') {
    const moraYes = inp.moratoriumAvail === 'Yes' || inp.moratoriumAvail === true;
    const rateLayersStr = (inp.rateLayers || []).length
      ? inp.rateLayers.map(l => `${fmtP(l.activeRate || 0)} from ${fmtDateDMY(l.fromDate)}`).join(', ')
      : '0 layer(s)';
    const secLayersStr = (inp.securityLayers || []).length
      ? inp.securityLayers.map(l => `+${fmtM(l.amount || 0)} from ${fmtDateDMY(l.fromDate)} at ${fmtP(l.activeRate || 0)}`).join(', ')
      : '0 layer(s)';
    return [
      ['Initial Loan Amount', inp.initialAmount ?? 0],
      ['Disbursement Date', inp.disbursementDate ?? ''],
      ['Moratorium Given at Disbursement?', yesNo(inp.moratoriumAvail)],
      ['Moratorium Period (Months)', inp.moratoriumPeriod ?? 0],
      [moraYes ? 'Payment Modality after Moratorium Period' : 'Payment Modality', inp.paymentModality ?? ''],
      ['Loan Tenor including Moratorium at Disbursement (Months)', inp.tenorMonths ?? 0],
      ['Lending Rate Layers', rateLayersStr],
      ['Loan Security Layers', secLayersStr],
      ['Cost of Fund Layers', (inp.cofRecordCount || 0) + ' record(s)'],
    ];
  }
  // revisionCustomized
  return [
    ['Uploaded Schedule', (inp.uploadedRowsCount || 0) + ' rows'],
    ['Loan Security Layers', (inp.securityLayers || []).length + ' layer(s)'],
    ['Cost of Fund Records', (inp.cofRecords || 0) + ' record(s)'],
  ];
}

function inputIndex(lines, start) {
  const idx = {};
  lines.forEach(([label], i) => {
    const row = start + i;
    const key = label.toLowerCase();
    if (key === 'offered rate') idx.offeredRate = row;
    else if (key === 'loan amount' || key === 'initial loan amount') idx.loanAmount = row;
    else if (key.startsWith('loan tenor including moratorium at disbursement')) idx.tenorMonths = row;
    else if (key.startsWith('loan tenor')) idx.loanTenor = row;
    else if (key.startsWith('moratorium period')) idx.moratoriumPeriod = row;
    else if (key.startsWith('total cof')) idx.totalCof = row;
    else if (key.startsWith('number of installments')) idx.numInst = row;
    else if (key === 'cash security / fdr amount') idx.csAmount = row;
    else if (key === 'cash security / fdr rate') idx.csRate = row;
  });
  return idx;
}

// Yearly Summary rows from a schedule — mirrors the Calculation Verification Excel:
// Income/Expense over months sl=firstSl..lastSl; Avg Portfolio + Loan Security Benefit over the
// start-of-month window sl=firstSl-1..lastSl-1; NIM = NII/AvgPort; YoY ERR = (Expense/AvgPort)+NIM.
function computeYearlySummary(ctx) {
  const rows = (ctx.schedule && ctx.schedule.rows) || [];
  if (rows.length < 2) return [];
  const m = ctx.metrics || {}, inp = ctx.inputs || {}, params = ctx.params || {};
  const isRevision = ctx.pageType === 'revisionStructured' || ctx.pageType === 'revisionCustomized';
  const cofRate = params.cofRate != null ? params.cofRate : (inp.totalCof || 0);
  const secAmt = m.derivedSecurityAmount != null ? m.derivedSecurityAmount : (inp.csAmount || 0);
  const secRate = m.derivedSecurityRate != null ? m.derivedSecurityRate : (inp.csRate || 0);

  // Per-row interest expense + CS benefit (units = per month).
  const per = rows.map((r, i) => {
    const prevUrpa = i > 0 ? (rows[i - 1].urpa || 0) : (r.urpa || 0);
    const intExp = isRevision ? (r.interestExpense || 0) : (i > 0 ? prevUrpa * cofRate / 12 : 0);
    const csBen = isRevision
      ? ((r.cof || 0) - (r.securityRate || 0)) * (r.securityAmount || 0) / 12
      : (cofRate - secRate) * secAmt / 12;
    return { interest: r.interest || 0, intExp, csBen, urpa: r.urpa || 0 };
  });

  const totalMonths = rows.length - 1;
  const years = Math.ceil(totalMonths / 12);
  const out = [];
  for (let y = 0; y < years; y++) {
    const firstSl = y * 12 + 1, lastSl = Math.min((y + 1) * 12, totalMonths);
    let income = 0, expense = 0, lsBen = 0, urpaSum = 0, urpaN = 0;
    for (let sl = firstSl; sl <= lastSl; sl++) {
      if (per[sl]) { income += per[sl].interest; expense += per[sl].intExp; }
      if (per[sl - 1]) { lsBen += per[sl - 1].csBen; urpaSum += per[sl - 1].urpa; urpaN++; }
    }
    const avgPort = urpaN ? urpaSum / urpaN : 0;
    const nii = income + lsBen - expense;
    // Annualise both legs by 12/months so a final stub year is comparable with the full
    // years above it; full years scale by 1 and are unchanged.
    const ann = 12 / (lastSl - firstSl + 1);
    const nim = avgPort ? (nii / avgPort) * ann : 0;
    const yoyErr = avgPort ? (expense / avgPort) * ann + nim : 0;
    out.push({ year: y + 1, income, expense, lsBen, nii, avgPort, nim, yoyErr });
  }
  return out;
}

// =====================================================================
// Download Report — PDF with inputs + results + schedule (replication-ready)
// =====================================================================
export function downloadReportPDF(filename, ctx) {
  const { jsPDF } = window.jspdf || {};
  if (!jsPDF) { alert('PDF library not loaded.'); return; }
  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
  let y = 36;
  doc.setFontSize(16); doc.setFont(undefined, 'bold');
  doc.text(ctx.pageTitle, 40, y); y += 22;
  doc.setFontSize(10); doc.setFont(undefined, 'normal');
  doc.text('Generated: ' + new Date().toLocaleString(), 40, y); y += 16;

  // Inputs
  doc.setFontSize(12); doc.setFont(undefined, 'bold');
  doc.text('Inputs', 40, y); y += 10;
  const inputLines = collectInputLinesFor(ctx.pageType, ctx.inputs);
  // Rename "Payment Mode" / "Payment Modality" → "Payment After Moratorium" when moratorium exists.
  const reportMora = ctx.inputs.moratoriumAvail === 'Yes' || (Number(ctx.inputs.moratoriumPeriod) || 0) > 0;
  if (reportMora) inputLines.forEach((line) => {
    if (line[0] === 'Payment Mode' || line[0] === 'Payment Modality') line[0] = 'Payment After Moratorium';
  });
  // Add layer detail
  const extras = [];
  if (Array.isArray(ctx.inputs.paymentLayers)) {
    ctx.inputs.paymentLayers.forEach((L, i) => {
      extras.push([`Payment Layer ${i + 1}`, `${L.paymentType}: Month ${L.fromInstallment}–${L.toInstallment}` + (L.customPrincipal ? ` @ ${fmtM(L.customPrincipal)}/inst` : '')]);
    });
  }
  if (Array.isArray(ctx.inputs.rateLayers)) {
    ctx.inputs.rateLayers.forEach((L, i) => {
      extras.push([`Rate Layer ${i + 1}`, `${L.fromDate} – ${L.toDate}: ${fmtP(L.activeRate)}`]);
    });
  }
  if (Array.isArray(ctx.inputs.securityLayers) && ctx.inputs.securityLayers.length) {
    ctx.inputs.securityLayers.forEach((L, i) => {
      extras.push([`Security Layer ${i + 1}`, `${L.fromDate} – ${L.toDate}: amt ${fmtM(L.amount)} @ ${fmtP(L.activeRate)}`]);
    });
  }
  if (Array.isArray(ctx.inputs.cofLayers) && ctx.inputs.cofLayers.length) {
    ctx.inputs.cofLayers.forEach((L, i) => {
      extras.push([`COF Layer ${i + 1}`, `${L.fromDate} – ${L.toDate}: ${fmtP(L.cofRate)}`]);
    });
  }
  if (Array.isArray(ctx.inputs.idpFlags) && ctx.inputs.idpFlags.some(Boolean)) {
    const months = ctx.inputs.idpFlags.map((v, i) => v ? i + 1 : null).filter(Boolean);
    extras.push(['Interest-paid Moratorium Months', months.map(m => `M${String(m).padStart(2, '0')}`).join(', ')]);
  }

  // Render input rows
  const formatLine = ([k, v]) => {
    let val = v;
    if (typeof v === 'number') {
      if (k.toLowerCase().includes('rate') || k.toLowerCase().startsWith('total cof')) val = fmtP(v);
      else val = fmtM(v);
    }
    return [k, String(val)];
  };
  doc.autoTable({
    head: [['Field', 'Value']],
    body: [...inputLines.map(formatLine), ...extras],
    startY: y,
    styles: { fontSize: 9, cellPadding: 4, halign: 'center' },
    headStyles: { fillColor: [37, 70, 224], textColor: 255, halign: 'center' },
    margin: { left: 40, right: 40 },
    columnStyles: { 0: { cellWidth: 240, halign: 'center' }, 1: { halign: 'center' } },
  });
  y = doc.lastAutoTable.finalY + 18;

  // Results
  doc.setFont(undefined, 'bold'); doc.setFontSize(12);
  doc.text('Results', 40, y); y += 10;
  doc.autoTable({
    head: [['Metric', 'Value']],
    body: [
      ['Effective Rate (ERR)', fmtP(ctx.metrics.effectiveRate)],
      ['NIM', fmtP(ctx.metrics.nim)],
      ['Net Interest Income', fmtM(ctx.metrics.nii)],
      ['Avg Portfolio', fmtM(ctx.metrics.avgPortfolio)],
      ['Total Interest', fmtM(ctx.metrics.totalInterest)],
      ['Tenor (years)', String(Number(ctx.metrics.tenorYears || 0).toFixed(2))],
    ],
    startY: y,
    styles: { fontSize: 9, cellPadding: 4, halign: 'center' },
    headStyles: { fillColor: [37, 70, 224], textColor: 255, halign: 'center' },
    margin: { left: 40, right: 40 },
    columnStyles: { 0: { cellWidth: 240, halign: 'center' }, 1: { halign: 'center' } },
  });
  y = doc.lastAutoTable.finalY + 18;

  // Yearly Summary — same table as the Calculation Verification Excel.
  const ySummary = computeYearlySummary(ctx);
  if (ySummary.length) {
    doc.setFont(undefined, 'bold'); doc.setFontSize(12);
    doc.text('Yearly Summary', 40, y); y += 10;
    doc.autoTable({
      head: [['Year', 'Interest Income', 'Interest Expense', 'Loan Security Benefit', 'Net Interest Income', 'Average Portfolio', 'NIM', 'YoY ERR']],
      body: ySummary.map(r => [
        `Year ${String(r.year).padStart(2, '0')}`,
        fmtM(r.income), fmtM(r.expense), fmtM(r.lsBen), fmtM(r.nii), fmtM(r.avgPort), fmtP(r.nim), fmtP(r.yoyErr),
      ]),
      startY: y,
      styles: { fontSize: 7, cellPadding: 3, halign: 'center', valign: 'middle' },
      headStyles: { fillColor: [0, 176, 80], textColor: 255, halign: 'center' },
      margin: { left: 40, right: 40 },
    });
    y = doc.lastAutoTable.finalY + 18;
  }

  // Schedule
  doc.setFont(undefined, 'bold'); doc.setFontSize(12);
  doc.text('Amortization Schedule', 40, y); y += 10;
  const hasDate = ctx.schedule.rows[0]?.date !== undefined;
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  let tInst = 0, tInt = 0, tPrin = 0;
  const body = ctx.schedule.rows.map(r => {
    const row = [String(r.sl)];
    if (hasDate) row.push(r.date || '');
    row.push(fmtM(r.installment), fmtM(r.interest), fmtM(r.principal), fmtM(r.urpa));
    tInst += Number(r.installment) || 0; tInt += Number(r.interest) || 0; tPrin += Number(r.principal) || 0;
    return row;
  });
  const schedFoot = ['Total', ...(hasDate ? [''] : []), fmtM(tInst), fmtM(tInt), fmtM(tPrin), ''];
  doc.autoTable({
    head: [headers], body, foot: [schedFoot],
    startY: y,
    showFoot: 'lastPage',
    styles: { fontSize: 8, cellPadding: 3, halign: 'center', valign: 'middle' },
    headStyles: { fillColor: [37, 70, 224], textColor: 255, halign: 'center' },
    footStyles: { fillColor: [230, 235, 255], textColor: 20, fontStyle: 'bold', halign: 'center' },
    margin: { left: 40, right: 40 },
  });

  doc.save(filename);
}

export function downloadSampleAmortization() {
  const aoa = [
    ['Date', 'Installment Amount', 'Interest Amount', 'Principal Amount', 'URPA'],
    ['NOTE: First data row = DISBURSEMENT. Enter disbursement date and disbursement amount in URPA. Leave Installment/Interest/Principal blank or 0.', '', '', '', ''],
    ['2024-01-15', 0, 0, 0, 50000000],
    ['2024-02-15', 1112222, 500000, 612222, 49387778],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 18 }, { wch: 18 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Amortization');
  XLSX.writeFile(wb, 'Sample_Amortization_Schedule.xlsx');
}

export function readUploadedSchedule(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
        const out = [];
        for (const r of rows) {
          const date = r['Date'] ?? r['date'];
          if (!date) continue;
          let dateStr;
          if (date instanceof Date) dateStr = date.toISOString().slice(0, 10);
          else if (typeof date === 'string') {
            const d = new Date(date);
            if (isNaN(d)) continue;
            dateStr = d.toISOString().slice(0, 10);
          } else continue;
          out.push({
            date: dateStr,
            installmentAmount: Number(r['Installment Amount'] ?? r['installment'] ?? 0) || 0,
            interestAmount: Number(r['Interest Amount'] ?? r['interest'] ?? 0) || 0,
            principalAmount: Number(r['Principal Amount'] ?? r['principal'] ?? 0) || 0,
            urpa: Number(r['URPA'] ?? r['urpa'] ?? 0) || 0,
          });
        }
        if (out.length < 2) return reject(new Error('Uploaded file must contain disbursement row + at least one installment row.'));
        resolve(out);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Failed to read file.'));
    reader.readAsArrayBuffer(file);
  });
}

// COF Data sample — serve the exact bundled template (preserves Table, Date formula, styling).
export async function downloadCofSample() {
  try {
    const res = await fetch('assets/COF_Data_Sample.xlsx', { cache: 'no-store' });
    if (!res.ok) throw new Error('sample not found');
    const blob = await res.blob();
    saveBlob(blob, 'COF_Data_Sample.xlsx');
  } catch (err) {
    alert('Could not load the COF sample file. Make sure assets/COF_Data_Sample.xlsx is deployed.');
  }
}

// Parse a COF worksheet grid → [{ date:'YYYY-MM-DD', rate }] from a Year/Month/Day/Date/COF
// table (legend rows above the header are skipped). Returns [] if no COF column is found.
// Rate column preference: exact "COF" (legacy template) → "Eligible COF" (current template,
// the computed all-in rate the schedule uses) → any other header containing "COF".
function parseCofGrid(ws) {
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  let headerRow = -1;
  for (let i = 0; i < grid.length; i++) {
    const cells = (grid[i] || []).map(c => String(c ?? '').trim().toLowerCase());
    if (cells.some(c => c.includes('cof'))) { headerRow = i; break; }
  }
  if (headerRow < 0) return [];
  const headers = grid[headerRow].map(c => String(c ?? '').trim().toLowerCase());
  const col = (name) => headers.indexOf(name);
  const iYear = col('year'), iMonth = col('month'), iDay = col('day'), iDate = col('date');
  const iEligible = col('eligible cof');
  const iIsc = col('isc');
  // Raw monthly COF column: an exact 'cof', else a cof-containing header that isn't 'eligible cof'.
  let iRawCof = col('cof');
  if (iRawCof < 0) iRawCof = headers.findIndex(h => h.includes('cof') && !h.includes('eligible'));
  // Rate the engine uses = Eligible COF when present (the all-in rate), else the raw COF column.
  let iRate = iEligible >= 0 ? iEligible : iRawCof;
  if (iRate < 0) iRate = headers.findIndex(h => h.includes('cof'));
  const numOrU = (v) => (v === null || v === undefined || v === '' ? undefined : Number(v));
  const out = [];
  for (let i = headerRow + 1; i < grid.length; i++) {
    const row = grid[i] || [];
    const rateRaw = iRate >= 0 ? row[iRate] : null;
    if (rateRaw === null || rateRaw === undefined || rateRaw === '') continue;
    let iso = null;
    if (iYear >= 0 && iMonth >= 0 && iDay >= 0 && row[iYear] && row[iMonth] && row[iDay]) {
      const y = Number(row[iYear]), mo = Number(row[iMonth]), da = Number(row[iDay]);
      if (y && mo && da) iso = `${y}-${String(mo).padStart(2, '0')}-${String(da).padStart(2, '0')}`;
    }
    if (!iso && iDate >= 0 && row[iDate]) {
      const dv = row[iDate];
      const d = dv instanceof Date ? dv : new Date(dv);
      if (!isNaN(d)) iso = d.toISOString().slice(0, 10);
    }
    if (!iso) continue;
    // rate = Eligible COF (engine); cof/isc = raw inputs (verification sheet rebuilds Eligible = MAX(cof,isc)+0.3%).
    out.push({ date: iso, rate: Number(rateRaw), cof: numOrU(iRawCof >= 0 ? row[iRawCof] : null), isc: numOrU(iIsc >= 0 ? row[iIsc] : null) });
  }
  return out;
}

// Read an uploaded COF file → [{ date:'YYYY-MM-DD', rate }] (first sheet).
export function readUploadedCof(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        const out = parseCofGrid(wb.Sheets[wb.SheetNames[0]]);
        if (!out.length) return reject(new Error('No COF records found. Fill the Year/Month/Day and COF columns.'));
        resolve(out);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Failed to read file.'));
    reader.readAsArrayBuffer(file);
  });
}

// Download the combined Rate Revision — Customized template (Schedule + COF Layers sheets).
export async function downloadCustomizedRevisionSample() {
  try {
    const res = await fetch('assets/Sample for customized rate revision.xlsx', { cache: 'no-store' });
    if (!res.ok) throw new Error('sample not found');
    saveBlob(await res.blob(), 'Sample for customized rate revision.xlsx');
  } catch (err) {
    alert('Could not load the sample file. Make sure "assets/Sample for customized rate revision.xlsx" is deployed.');
  }
}

// Read the combined Rate Revision — Customized file → { scheduleRows, cofRows }.
// Amortization comes from the 'Schedule' sheet; COF from the 'COF Layers' sheet.
export function readCustomizedRevisionFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        const byName = (re) => { const n = wb.SheetNames.find(nm => re.test(nm)); return n ? wb.Sheets[n] : null; };
        const hasHeader = (ws, h) => ((XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })[0] || [])
          .map(c => String(c ?? '').trim().toLowerCase()).includes(h));
        let schedWs = byName(/schedule|amort/i)
          || wb.SheetNames.map(n => wb.Sheets[n]).find(ws => hasHeader(ws, 'urpa'));
        if (!schedWs) return reject(new Error('Could not find the amortization Schedule sheet (Date / Installment / Interest / Principal / URPA).'));
        const numv = v => Number(String(v ?? '').replace(/,/g, '')) || 0;
        const json = XLSX.utils.sheet_to_json(schedWs, { defval: null });
        const scheduleRows = [];
        for (const r of json) {
          const date = r['Date'] ?? r['date'];
          if (!date) continue;
          let dateStr;
          if (date instanceof Date) dateStr = date.toISOString().slice(0, 10);
          else { const d = new Date(date); if (isNaN(d)) continue; dateStr = d.toISOString().slice(0, 10); }
          scheduleRows.push({
            date: dateStr,
            installmentAmount: numv(r['Installment Amount'] ?? r['Installment'] ?? r['installment']),
            interestAmount: numv(r['Interest Amount'] ?? r['Interest'] ?? r['interest']),
            principalAmount: numv(r['Principal Amount'] ?? r['Principal'] ?? r['principal']),
            urpa: numv(r['URPA'] ?? r['urpa']),
          });
        }
        if (scheduleRows.length < 2) return reject(new Error('The Schedule sheet needs a disbursement row + at least one installment row.'));
        const cofWs = byName(/cof/i);
        const cofRows = cofWs ? parseCofGrid(cofWs) : [];
        resolve({ scheduleRows, cofRows });
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Failed to read file.'));
    reader.readAsArrayBuffer(file);
  });
}

function saveBlob(blob, filename) {
  if (window.saveAs) return window.saveAs(blob, filename);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 0);
}
