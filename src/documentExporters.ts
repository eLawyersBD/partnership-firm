import * as XLSX from 'xlsx';
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  PageBreak,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';

export interface ExportReportPayload {
  firmName: string;
  reportDate: string;
  prevDate: string;
  regNo: string;
  estDate: string;
  address: string;
  businessNature: string;
  approvalDate: string;
  includeAudit: boolean;
  auditorFirm: string;
  auditorPartner: string;
  txtOps: string;
  txtRisks: string;
  txtFuture: string;
  txtResp: string;
  partners: Array<{
    id: number;
    name: string;
    role: string;
    ratio: number;
    open: number;
    add: number;
    draw: number;
    prof: number;
    close: number;
  }>;
  totals: {
    ratioSum: number;
    totOpen: number;
    totAdd: number;
    totProf: number;
    totDraw: number;
    totClose: number;
  };
  pnl: {
    revLabel: string;
    revNote: string;
    revCurr: number;
    revPrev: number;
    cosLabel: string;
    cosNote: string;
    cosCurr: number;
    cosPrev: number;
    gpCurr: number;
    gpPrev: number;
    othIncLabel: string;
    othIncNote: string;
    othIncCurr: number;
    othIncPrev: number;
    expenses: Array<{ label: string; note: string; curr: number; prev: number }>;
    opCurr: number;
    opPrev: number;
    finLabel: string;
    finNote: string;
    finCurr: number;
    finPrev: number;
    pbtCurr: number;
    pbtPrev: number;
    taxLabel: string;
    taxNote: string;
    taxCurr: number;
    taxPrev: number;
    npCurr: number;
    npPrev: number;
    ociLabel: string;
    ociCurr: number;
    ociPrev: number;
    tciCurr: number;
    tciPrev: number;
  };
  bs: {
    nca: Array<{ label: string; note: string; curr: number; prev: number }>;
    ca: Array<{ label: string; note: string; curr: number; prev: number }>;
    ncl: Array<{ label: string; note: string; curr: number; prev: number }>;
    cl: Array<{ label: string; note: string; curr: number; prev: number }>;
    ncaCurr: number;
    ncaPrev: number;
    caCurr: number;
    caPrev: number;
    totAssetsCurr: number;
    totAssetsPrev: number;
    eqCurr: number;
    eqPrev: number;
    nclCurr: number;
    nclPrev: number;
    clCurr: number;
    clPrev: number;
    totLiabCurr: number;
    totLiabPrev: number;
    totEqLiabCurr: number;
    totEqLiabPrev: number;
  };
  cf: {
    cfOpsGen: number;
    netOps: number;
    investing: Array<{ label: string; amount: number }>;
    netInv: number;
    financing: Array<{ label: string; amount: number }>;
    netFin: number;
    netCashChange: number;
    cfOpenCash: number;
    cfCloseCash: number;
  };
  ratios: {
    gpm: string;
    npm: string;
    currRatio: string;
    roe: string;
  };
  notes8to14: any;
  notes15to25: any;
  disclosureNotes: Array<{ title: string; text: string }>;
  tax: {
    includeTax: boolean;
    taxPeriodTitle: string;
    slabRatePct: number;
    grossReceipts: number;
    taxInvoiceRef: string;
    taxableProfit: number;
    regTaxRatePct: number;
    taxPbtNote: string;
    aitPaid: number;
    taxAitRef: string;
    deferredTaxAmt: number;
    taxClosingLabel: string;
    taxStatutoryNote: string;
    minTaxA: number;
    regTaxB: number;
    applicableTax: number;
    netCurrentTaxPayable: number;
  };
  it116: {
    includeIt116: boolean;
    tin: string;
    bin: string;
    circle: string;
    zone: string;
    taxYear: string;
    category: string;
    resident: string;
    regDate: string;
    phone: string;
    email: string;
    rentInc: number;
    agriInc: number;
    bizInc: number;
    capGain: number;
    finInc: number;
    othSrcInc: number;
    foreignInc: number;
    totalIncome: number;
    regularTaxOnTotalInc: number;
    minTaxOnTds: number;
    turnoverTaxCalc: number;
    taxPayableHigher: number;
    envSurcharge: number;
    penaltyAmt: number;
    totalTaxPayable: number;
    tdsPaid: number;
    aitPaidReturn: number;
    balanceTaxWithReturn: number;
    totalTaxPaid: number;
    excessPayment: number;
    exemptInc: number;
    docsList: string;
  };
}

function fmtNum(n: number, parens = false): string {
  if (isNaN(n) || n === null) return '0';
  const r = Math.round(n);
  const abs = Math.abs(r).toLocaleString('en-US');
  if (r < 0 || parens) {
    return r === 0 ? '0' : `(${abs})`;
  }
  return abs;
}

function fmtDec2(n: number, parens = false): string {
  if (isNaN(n) || n === null) return '0.00';
  const abs = Math.abs(n).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  if (n < 0 || parens) {
    return Math.abs(n) < 0.005 ? '0.00' : `(${abs})`;
  }
  return abs;
}

function autoFitSheetCols(aoa: any[][]): XLSX.ColInfo[] {
  const maxCols = aoa.reduce((m, r) => Math.max(m, r ? r.length : 0), 0);
  const cols: XLSX.ColInfo[] = [];
  for (let c = 0; c < maxCols; c++) {
    let maxLen = 12;
    for (let r = 0; r < aoa.length; r++) {
      const val = aoa[r]?.[c];
      if (val !== undefined && val !== null) {
        const len = String(val).length;
        if (len > maxLen && len < 65) maxLen = len;
      }
    }
    cols.push({ wch: maxLen + 3 });
  }
  return cols;
}

// ==================== 1. EXCEL (.XLSX) GENERATOR USING `xlsx` ====================
export function downloadReportAsXlsx(data: ExportReportPayload): void {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Overview, Partners & Ratios
  const s1: any[][] = [
    [data.firmName.toUpperCase()],
    [`ANNUAL FINANCIAL REPORT — FOR THE FINANCIAL YEAR ENDED ${data.reportDate.toUpperCase()}`],
    ['Prepared in accordance with IFRS, IAS & Bangladesh FRC Reporting Standards'],
    [],
    ['ENTITY PROFILE', 'DETAILS'],
    ['Partnership Firm Name', data.firmName],
    ['Financial Year Ended', data.reportDate],
    ['Prior Year Ended', data.prevDate],
    ['Registration No. (RJSC)', data.regNo],
    ['Date of Establishment', data.estDate],
    ['Registered Office Address', data.address],
    ['Nature of Business', data.businessNature],
    ['Approval / Audit Date', data.approvalDate],
    [],
    ['PARTNERSHIP STRUCTURE & CAPITAL SUMMARY'],
    ['Partner Name', 'Role / Designation', 'Profit Share (%)', 'Opening Capital (BDT)', 'Additional Capital (BDT)', 'Profit Allocated (BDT)', 'Drawings (BDT)', 'Closing Capital (BDT)'],
    ...data.partners.map((p) => [
      p.name,
      p.role,
      p.ratio,
      p.open,
      p.add,
      p.prof,
      -p.draw,
      p.close,
    ]),
    [
      'Total',
      '',
      data.totals.ratioSum,
      data.totals.totOpen,
      data.totals.totAdd,
      data.totals.totProf,
      -data.totals.totDraw,
      data.totals.totClose,
    ],
    [],
    ['EXECUTIVE SUMMARY — KEY FINANCIAL HIGHLIGHTS & RATIOS'],
    ['Financial Indicator', 'Current Year (BDT)', 'Prior Year (BDT)'],
    ['Total Revenue', data.pnl.revCurr, data.pnl.revPrev],
    ['Gross Profit', data.pnl.gpCurr, data.pnl.gpPrev],
    ['Net Profit for the Year', data.pnl.npCurr, data.pnl.npPrev],
    ['Total Assets', data.bs.totAssetsCurr, data.bs.totAssetsPrev],
    ['Total Liabilities', data.bs.totLiabCurr, data.bs.totLiabPrev],
    ["Total Partners' Equity", data.bs.eqCurr, data.bs.eqPrev],
    ['Net Operating Cash Flow', data.cf.netOps, ''],
    ['Closing Cash & Cash Equivalents', data.cf.cfCloseCash, ''],
    [],
    ['KEY FINANCIAL RATIOS', 'VALUE'],
    ['Gross Profit Margin', `${data.ratios.gpm}%`],
    ['Net Profit Margin', `${data.ratios.npm}%`],
    ['Current Ratio', `${data.ratios.currRatio}x`],
    ['Return on Equity (ROE)', `${data.ratios.roe}%`],
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(s1);
  ws1['!cols'] = autoFitSheetCols(s1);
  XLSX.utils.book_append_sheet(wb, ws1, '1_Overview_Ratios');

  // Sheet 2: Statement of Financial Position (Balance Sheet)
  const s2: any[][] = [
    [data.firmName.toUpperCase()],
    ['STATEMENT OF FINANCIAL POSITION (IAS 1)'],
    [`As at ${data.reportDate}`],
    [],
    ['Particulars', 'Notes', `${data.reportDate} (BDT)`, `${data.prevDate} (BDT)`],
    ['ASSETS', '', '', ''],
    ['Non-current assets', '', '', ''],
    ...data.bs.nca.map((i) => [i.label, i.note, i.curr, i.prev]),
    ['Total non-current assets', '', data.bs.ncaCurr, data.bs.ncaPrev],
    [],
    ['Current assets', '', '', ''],
    ...data.bs.ca.map((i) => [i.label, i.note, i.curr, i.prev]),
    ['Total current assets', '', data.bs.caCurr, data.bs.caPrev],
    ['TOTAL ASSETS', '', data.bs.totAssetsCurr, data.bs.totAssetsPrev],
    [],
    ['EQUITY AND LIABILITIES', '', '', ''],
    ["Partners' Equity (IAS 32)", '', '', ''],
    ["Partners' capital and current accounts", 'Note 12', data.bs.eqCurr, data.bs.eqPrev],
    [],
    ['Non-current liabilities', '', '', ''],
    ...data.bs.ncl.map((i) => [i.label, i.note, i.curr, i.prev]),
    ['Total non-current liabilities', '', data.bs.nclCurr, data.bs.nclPrev],
    [],
    ['Current liabilities', '', '', ''],
    ...data.bs.cl.map((i) => [i.label, i.note, i.curr, i.prev]),
    ['Total current liabilities', '', data.bs.clCurr, data.bs.clPrev],
    ['TOTAL EQUITY AND LIABILITIES', '', data.bs.totEqLiabCurr, data.bs.totEqLiabPrev],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(s2);
  ws2['!cols'] = autoFitSheetCols(s2);
  XLSX.utils.book_append_sheet(wb, ws2, '2_Balance_Sheet');

  // Sheet 3: Profit or Loss & Changes in Partners' Equity
  const s3: any[][] = [
    [data.firmName.toUpperCase()],
    ['STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME (IAS 1)'],
    [`For the financial year ended ${data.reportDate}`],
    [],
    ['Particulars', 'Notes', 'Current Year (BDT)', 'Previous Year (BDT)'],
    [data.pnl.revLabel, data.pnl.revNote, data.pnl.revCurr, data.pnl.revPrev],
    [`Less: ${data.pnl.cosLabel}`, data.pnl.cosNote, -data.pnl.cosCurr, -data.pnl.cosPrev],
    ['Gross Profit', '', data.pnl.gpCurr, data.pnl.gpPrev],
    [data.pnl.othIncLabel, data.pnl.othIncNote, data.pnl.othIncCurr, data.pnl.othIncPrev],
    ...data.pnl.expenses.map((e) => [e.label, e.note, -e.curr, -e.prev]),
    ['Operating Profit', '', data.pnl.opCurr, data.pnl.opPrev],
    [data.pnl.finLabel, data.pnl.finNote, -data.pnl.finCurr, -data.pnl.finPrev],
    ['Profit Before Tax', '', data.pnl.pbtCurr, data.pnl.pbtPrev],
    [data.pnl.taxLabel, data.pnl.taxNote, -data.pnl.taxCurr, -data.pnl.taxPrev],
    ['Profit for the Year', '', data.pnl.npCurr, data.pnl.npPrev],
    [data.pnl.ociLabel, '', data.pnl.ociCurr, data.pnl.ociPrev],
    ['TOTAL COMPREHENSIVE INCOME FOR THE YEAR', '', data.pnl.tciCurr, data.pnl.tciPrev],
    [],
    ['ALLOCATION OF PROFIT AMONG PARTNERS (NOTE 14)'],
    ['Partner', 'Profit Share Ratio', 'Allocated Profit (BDT)'],
    ...data.partners.map((p) => [p.name, `${p.ratio}%`, p.prof]),
    ["Total Profit Allocated to Partners' Equity", `${data.totals.ratioSum}%`, data.totals.totProf],
    [],
    ["STATEMENT OF CHANGES IN PARTNERS' EQUITY (IAS 1 & IAS 32)"],
    ['Particulars', ...data.partners.map((p) => `${p.name} (BDT)`), 'Total Equity (BDT)'],
    ['Opening capital balance', ...data.partners.map((p) => p.open), data.totals.totOpen],
    ['Additional capital introduced', ...data.partners.map((p) => p.add), data.totals.totAdd],
    ['Profit allocated for the year', ...data.partners.map((p) => p.prof), data.totals.totProf],
    ["Less: Partners' drawings", ...data.partners.map((p) => -p.draw), -data.totals.totDraw],
    ['CLOSING CAPITAL BALANCE', ...data.partners.map((p) => p.close), data.totals.totClose],
  ];
  const ws3 = XLSX.utils.aoa_to_sheet(s3);
  ws3['!cols'] = autoFitSheetCols(s3);
  XLSX.utils.book_append_sheet(wb, ws3, '3_Profit_Loss_Equity');

  // Sheet 4: Statement of Cash Flows (IAS 7)
  const s4: any[][] = [
    [data.firmName.toUpperCase()],
    ['STATEMENT OF CASH FLOWS (IAS 7)'],
    [`For the financial year ended ${data.reportDate}`],
    [],
    ['Particulars', 'Current Year (BDT)'],
    ['A. Cash Flows from Operating Activities', ''],
    ['Cash generated from operations', data.cf.cfOpsGen],
    ['Interest paid', -data.pnl.finCurr],
    ['Income taxes paid', -data.pnl.taxCurr],
    ['Net Cash from Operating Activities', data.cf.netOps],
    [],
    ['B. Cash Flows from Investing Activities', ''],
    ...data.cf.investing.map((i) => [i.label, i.amount]),
    ['Net Cash from / (Used in) Investing Activities', data.cf.netInv],
    [],
    ['C. Cash Flows from Financing Activities', ''],
    ['Capital introduced by partners', data.totals.totAdd],
    ["Partners' withdrawals (drawings)", -data.totals.totDraw],
    ...data.cf.financing.map((i) => [i.label, i.amount]),
    ['Net Cash from / (Used in) Financing Activities', data.cf.netFin],
    [],
    ['Net Increase / (Decrease) in Cash and Cash Equivalents (A + B + C)', data.cf.netCashChange],
    ['Opening cash and cash equivalents', data.cf.cfOpenCash],
    ['CLOSING CASH AND CASH EQUIVALENTS', data.cf.cfCloseCash],
  ];
  const ws4 = XLSX.utils.aoa_to_sheet(s4);
  ws4['!cols'] = autoFitSheetCols(s4);
  XLSX.utils.book_append_sheet(wb, ws4, '4_Cash_Flows');

  // Sheet 5: Notes 8 to 25 Schedules
  const s5: any[][] = [
    [data.firmName.toUpperCase()],
    [`NOTES TO THE FINANCIAL STATEMENTS (SCHEDULES 8 TO 25) — FY ENDED ${data.reportDate}`],
    [],
    [data.notes8to14.n8.title],
    ['Category', 'Cost (BDT)', 'Accumulated Depreciation (BDT)', 'Carrying Amount (BDT)'],
    ...data.notes8to14.n8.rows.map((r: any) => [r.cat, r.cost, r.dep, r.cost + r.dep]),
    [data.notes8to14.n8.footer],
    [],
    [data.notes8to14.n9.title],
    ['Particulars', 'Current Year (BDT)', 'Prior Year (BDT)'],
    ...data.notes8to14.n9.rows.map((r: any) => [r.label, r.curr, r.prev]),
    [],
    [data.notes8to14.n10.title],
    ['Particulars', 'Current Year (BDT)', 'Prior Year (BDT)'],
    ...data.notes8to14.n10.rows.map((r: any) => [r.label, r.curr, r.prev]),
    [],
    [data.notes8to14.n11.title],
    ['Particulars', 'Current Year (BDT)', 'Prior Year (BDT)'],
    ...data.notes8to14.n11.rows.map((r: any) => [r.label, r.curr, r.prev]),
    [],
    [data.notes15to25.n15.title],
    ['Particulars', 'Current Year (BDT)', 'Prior Year (BDT)'],
    ...data.notes15to25.n15.rows.map((r: any) => [r.label, r.curr ?? '', r.prev ?? '']),
    [],
    [data.notes15to25.n16.title],
    ['Particulars', 'Current Year (BDT)', 'Prior Year (BDT)'],
    ...data.notes15to25.n16.rows.map((r: any) => [r.label, r.curr, r.prev]),
    [],
    [data.notes15to25.n18.title],
    ['Particulars', 'Current Year (BDT)', 'Prior Year (BDT)'],
    ...data.notes15to25.n18.rows.map((r: any) => [r.label, r.curr, r.prev]),
    [],
    ...data.disclosureNotes.map((n) => [n.title, n.text]),
  ];
  const ws5 = XLSX.utils.aoa_to_sheet(s5);
  ws5['!cols'] = autoFitSheetCols(s5);
  XLSX.utils.book_append_sheet(wb, ws5, '5_Notes_Schedules');

  // Sheet 6: IAS 12 & S.R.O. 324/2026 Taxation Schedule
  const s6: any[][] = [
    [data.firmName.toUpperCase()],
    [`NOTE 19: CURRENT AND DEFERRED TAXATION (${data.tax.taxPeriodTitle})`],
    ['Compliant with IAS 12 & S.R.O. No. 324-Ain/Aykor-05/2026 (Income Tax Act 2023)'],
    [],
    ['SL', 'Corporate Tax Assessment Item', 'Statutory Reference / Rate', 'Calculation Base', 'Amount (BDT)'],
    ['1', 'Gross Receipts / Turnover from Contracts with Customers', 'S.R.O. 324/2026', data.tax.taxInvoiceRef, data.tax.grossReceipts],
    ['2', `Minimum Income Tax @ ${data.tax.slabRatePct.toFixed(2)}% of Gross Receipts (A)`, 'Section 163(6) / S.R.O. 324', `${fmtDec2(data.tax.grossReceipts)} @ ${data.tax.slabRatePct.toFixed(2)}%`, data.tax.minTaxA],
    ['3', 'Taxable Business Profit for the Period', 'Income Statement', data.tax.taxPbtNote, data.tax.taxableProfit],
    ['4', `Tax Applied on Taxable Profit @ ${data.tax.regTaxRatePct.toFixed(1)}% (B)`, `Regular Tax Rate @ ${data.tax.regTaxRatePct.toFixed(2)}%`, `${fmtDec2(data.tax.taxableProfit)} @ ${data.tax.regTaxRatePct.toFixed(2)}%`, data.tax.regTaxB],
    ['—', 'APPLICABLE CORPORATE INCOME TAX (HIGHER OF A AND B)', 'Higher of (A) & (B)', data.tax.regTaxB >= data.tax.minTaxA ? 'Regular Tax (B) Applies' : 'Minimum Tax (A) Applies', data.tax.applicableTax],
    ['5', 'Less: Advance Tax Paid (AIT) / TDS Deducted at Source', 'Section 163 / AIT Challan', data.tax.taxAitRef, -data.tax.aitPaid],
    ['6', 'Deferred Tax Expense / (Income) Movement (IAS 12)', 'IAS 12 Temporary Diff.', 'Deferred Tax Recon.', data.tax.deferredTaxAmt],
    ['—', data.tax.taxClosingLabel, 'Net Payable to NBR', data.tax.netCurrentTaxPayable === 0 ? 'Balance Due (TDS covers liability)' : 'Net Tax Payable', data.tax.netCurrentTaxPayable],
    [],
    ['Note on Statutory Assessment:', data.tax.taxStatutoryNote],
  ];
  const ws6 = XLSX.utils.aoa_to_sheet(s6);
  ws6['!cols'] = autoFitSheetCols(s6);
  XLSX.utils.book_append_sheet(wb, ws6, '6_Tax_IAS12_SRO324');

  // Sheet 7: NBR Form IT-116 (2023) & IT-11 Cola
  if (data.it116.includeIt116) {
    const sellingExp = data.pnl.expenses[0]?.curr || 0;
    const adminAndOtherExp =
      data.pnl.expenses.reduce((s, e) => s + e.curr, 0) - sellingExp;
    const netOpProfitBeforeOther =
      data.pnl.gpCurr - adminAndOtherExp - sellingExp - data.pnl.finCurr;

    const s7: any[][] = [
      ['NATIONAL BOARD OF REVENUE — FORM IT-116 (2023) & IT-11 COLA (2023)'],
      [`Income Tax Return for Taxpayers Other Than Company — Income Year Ended ${data.reportDate}`],
      [],
      ['PART I: TAXPAYER PROFILE', 'VALUE'],
      ['1. Name of Taxpayer', data.firmName],
      ['2. TIN (12-Digit)', data.it116.tin],
      ['3. Business Identification Number (BIN)', data.it116.bin],
      ['4(a). Circle', data.it116.circle],
      ['4(b). Tax Zone', data.it116.zone],
      ['5. Tax / Assessment Year', data.it116.taxYear],
      ['6. Category of Taxpayer', data.it116.category.toUpperCase()],
      ['7. Residential Status', data.it116.resident.toUpperCase()],
      ['8. Contact Address', data.address],
      ['Telephone & E-mail', `${data.it116.phone} | ${data.it116.email}`],
      ['9. Registration No. & 10. Date', `${data.regNo} (${data.it116.regDate})`],
      [],
      ['PART II: STATEMENT OF TOTAL INCOME', 'AMOUNT IN TAKA (BDT)'],
      ['1. Income from Rent (Schedule 1)', data.it116.rentInc],
      ['2. Income from Agriculture (Schedule 2)', data.it116.agriInc],
      ['3. Income from Business (Schedule 3)', data.it116.bizInc],
      ['4. Capital Gain', data.it116.capGain],
      ['5. Income from Financial Assets', data.it116.finInc],
      ['6. Income from Other Sources', data.it116.othSrcInc],
      ['7. Foreign Sourced Taxable Income', data.it116.foreignInc],
      ['8. TOTAL INCOME (Sum of Serial 1 to 7)', data.it116.totalIncome],
      [],
      ['COMPUTATION OF TAX & STATEMENT OF TAX PAYMENT', 'AMOUNT IN TAKA (BDT)'],
      ['9. Income tax imposable on total taxable income', data.it116.regularTaxOnTotalInc],
      ['10. Minimum Tax', data.it116.minTaxOnTds],
      ['11. Turnover Tax', data.it116.turnoverTaxCalc],
      ['12. Tax Payable (Higher of Serial 9, 10, 11)', data.it116.taxPayableHigher],
      ['13. Environment Surcharge', data.it116.envSurcharge],
      ['14. Delay interest, penalty or other amount payable', data.it116.penaltyAmt],
      ['15. TOTAL TAX PAYABLE (12+13+14)', data.it116.totalTaxPayable],
      ['16. Tax deducted / collected at source (TDS)', data.it116.tdsPaid],
      ['17. Advance tax paid (AIT)', data.it116.aitPaidReturn],
      ['19. Balance tax paid along with this return', data.it116.balanceTaxWithReturn],
      ['20. TOTAL TAX PAID (Sum of 16 to 19)', data.it116.totalTaxPaid],
      ['21. Excess payment (if any)', data.it116.excessPayment],
      ['22. Tax exempted / tax-free income', data.it116.exemptInc],
      [],
      ['SCHEDULE 3: SUMMARY OF BUSINESS INCOME & BALANCE SHEET', 'AMOUNT IN TAKA (BDT)'],
      ['1. Sales / Turnover / Receipts', data.pnl.revCurr],
      ['3. Net Sales / Turnover / Receipts', data.pnl.revCurr],
      ['4. Cost of Goods Sold', data.pnl.cosCurr],
      ['5. Gross Profit (3-4)', data.pnl.gpCurr],
      ['6. General administrative and other expenses', adminAndOtherExp],
      ['7. Selling and marketing expenses', sellingExp],
      ['8. Interest and other expenses', data.pnl.finCurr],
      ['9. Net Profit (5-6-7-8)', netOpProfitBeforeOther],
      ['10. Other income', data.pnl.othIncCurr],
      ['11. Pre-tax income / profit', data.pnl.pbtCurr],
      ['Total Assets (Schedule 3 Serial 8)', data.bs.totAssetsCurr],
      ['Closing Partners Capital (Schedule 3 Serial 12)', data.totals.totClose],
      ['Total Capital and Liabilities (Schedule 3 Serial 15)', data.bs.totEqLiabCurr],
    ];
    const ws7 = XLSX.utils.aoa_to_sheet(s7);
    ws7['!cols'] = autoFitSheetCols(s7);
    XLSX.utils.book_append_sheet(wb, ws7, '7_NBR_IT116_Return');
  }

  const filename = `${data.firmName} - IFRS Financial Report & IT-116 (${data.reportDate}).xlsx`;
  XLSX.writeFile(wb, filename);
}

// ==================== 2. WORD (.DOCX) GENERATOR USING `docx` ====================
const CELL_BORDER = {
  top: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
  left: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
  right: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
};

function makeDocxTable(
  headers: string[],
  rows: Array<{ cells: string[]; bold?: boolean; shaded?: boolean }>,
  rightAlignFromCol = 1
): Table {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map(
      (h, idx) =>
        new TableCell({
          borders: CELL_BORDER,
          shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 70, bottom: 70, left: 100, right: 100 },
          children: [
            new Paragraph({
              alignment:
                idx >= rightAlignFromCol
                  ? AlignmentType.RIGHT
                  : AlignmentType.LEFT,
              children: [
                new TextRun({
                  text: h,
                  bold: true,
                  color: 'FFFFFF',
                  size: 19,
                  font: 'Calibri',
                }),
              ],
            }),
          ],
        })
    ),
  });

  const dataRows = rows.map(
    (r) =>
      new TableRow({
        children: r.cells.map(
          (c, idx) =>
            new TableCell({
              borders: CELL_BORDER,
              shading: r.shaded
                ? { fill: 'F1F5F9', type: ShadingType.CLEAR, color: 'auto' }
                : undefined,
              margins: { top: 55, bottom: 55, left: 100, right: 100 },
              children: [
                new Paragraph({
                  alignment:
                    idx >= rightAlignFromCol
                      ? AlignmentType.RIGHT
                      : AlignmentType.LEFT,
                  children: [
                    new TextRun({
                      text: String(c ?? ''),
                      bold: Boolean(r.bold),
                      size: 19,
                      font: 'Calibri',
                    }),
                  ],
                }),
              ],
            })
        ),
      })
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows],
  });
}

function sectionTitle(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 26,
        color: '065F46',
        font: 'Calibri',
      }),
    ],
  });
}

function subHeading(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 80 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 22,
        color: '0F172A',
        font: 'Calibri',
      }),
    ],
  });
}

function bodyPara(text: string, boldPrefix?: string): Paragraph {
  const children: TextRun[] = [];
  if (boldPrefix) {
    children.push(
      new TextRun({
        text: boldPrefix + ' ',
        bold: true,
        size: 20,
        font: 'Calibri',
      })
    );
  }
  children.push(
    new TextRun({
      text,
      size: 20,
      font: 'Calibri',
    })
  );
  return new Paragraph({
    spacing: { after: 110 },
    children,
  });
}

export async function downloadReportAsDocx(
  data: ExportReportPayload
): Promise<void> {
  const children: Array<Paragraph | Table> = [];

  // PAGE 1: COVER & ENTITY DETAILS
  children.push(
    new Paragraph({
      spacing: { before: 400, after: 120 },
      children: [
        new TextRun({
          text: 'ANNUAL FINANCIAL REPORT',
          bold: true,
          size: 24,
          color: '065F46',
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({
          text: data.firmName,
          bold: true,
          size: 44,
          color: '0F172A',
          font: 'Calibri',
        }),
      ],
    }),
    bodyPara(`For the Financial Year Ended ${data.reportDate}`),
    bodyPara(
      'Prepared in accordance with International Financial Reporting Standards (IFRS) and International Accounting Standards (IAS) as issued by the IASB and adopted in Bangladesh by the Financial Reporting Council (FRC).'
    ),
    makeDocxTable(
      ['Entity Attribute', 'Registered Particulars'],
      [
        { cells: ['Partnership Firm Name', data.firmName], bold: true },
        { cells: ['Financial Year Ended', data.reportDate] },
        { cells: ['Comparative Prior Year Ended', data.prevDate] },
        { cells: ['Registration Number (RJSC)', data.regNo] },
        { cells: ['Date of Establishment', data.estDate] },
        { cells: ['Registered Office Address', data.address] },
        { cells: ['Nature of Business & Activities', data.businessNature] },
        { cells: ['Reporting Currency', 'Bangladeshi Taka (BDT)'] },
      ],
      99
    ),
    new Paragraph({ children: [new PageBreak()] })
  );

  // PAGE 2: SECTION A — MANAGEMENT REPORT & EXECUTIVE SUMMARY
  children.push(
    sectionTitle('SECTION A — MANAGEMENT REPORT & EXECUTIVE SUMMARY'),
    subHeading('1.1 Introduction & Business Operations'),
    bodyPara(data.txtOps),
    subHeading('1.2 Partnership Structure & Capital Summary'),
    makeDocxTable(
      [
        'Partner Name',
        'Designation / Role',
        'Profit Ratio',
        'Opening Cap (BDT)',
        'Added Cap (BDT)',
        'Profit Share (BDT)',
        'Drawings (BDT)',
        'Closing Cap (BDT)',
      ],
      [
        ...data.partners.map((p) => ({
          cells: [
            p.name,
            p.role,
            `${p.ratio}%`,
            fmtNum(p.open),
            fmtNum(p.add),
            fmtNum(p.prof),
            fmtNum(p.draw, true),
            fmtNum(p.close),
          ],
        })),
        {
          cells: [
            'Total Partners Equity',
            '—',
            `${data.totals.ratioSum}%`,
            fmtNum(data.totals.totOpen),
            fmtNum(data.totals.totAdd),
            fmtNum(data.totals.totProf),
            fmtNum(data.totals.totDraw, true),
            fmtNum(data.totals.totClose),
          ],
          bold: true,
          shaded: true,
        },
      ],
      2
    ),
    subHeading('1.3 Risk Management & Future Outlook'),
    bodyPara(data.txtRisks, 'Risk Management:'),
    bodyPara(data.txtFuture, 'Future Outlook:'),
    bodyPara(data.txtResp, 'Management Responsibility:'),
    subHeading('2. Executive Summary — Key Financial Highlights & Ratios'),
    makeDocxTable(
      ['Financial Indicator / Ratio', 'Current Period', 'Comparative / Formula'],
      [
        {
          cells: [
            'Total Revenue',
            `BDT ${fmtNum(data.pnl.revCurr)}`,
            `Prior: BDT ${fmtNum(data.pnl.revPrev)}`,
          ],
          bold: true,
        },
        {
          cells: [
            'Gross Profit & Margin',
            `BDT ${fmtNum(data.pnl.gpCurr)}`,
            `Gross Margin: ${data.ratios.gpm}%`,
          ],
        },
        {
          cells: [
            'Net Profit & Margin',
            `BDT ${fmtNum(data.pnl.npCurr)}`,
            `Net Margin: ${data.ratios.npm}%`,
          ],
          bold: true,
          shaded: true,
        },
        {
          cells: [
            'Total Assets',
            `BDT ${fmtNum(data.bs.totAssetsCurr)}`,
            `Prior: BDT ${fmtNum(data.bs.totAssetsPrev)}`,
          ],
        },
        {
          cells: [
            "Total Partners' Equity & ROE",
            `BDT ${fmtNum(data.bs.eqCurr)}`,
            `Return on Equity (ROE): ${data.ratios.roe}%`,
          ],
          bold: true,
        },
        {
          cells: [
            'Current Ratio (Liquidity)',
            `${data.ratios.currRatio}x`,
            `CA: BDT ${fmtNum(data.bs.caCurr)} ÷ CL: BDT ${fmtNum(data.bs.clCurr)}`,
          ],
        },
      ],
      1
    ),
    new Paragraph({ children: [new PageBreak()] })
  );

  // SECTION B: INDEPENDENT AUDITOR'S REPORT (IF ENABLED)
  if (data.includeAudit) {
    children.push(
      sectionTitle("SECTION B — INDEPENDENT AUDITOR'S REPORT (ISA 700)"),
      bodyPara(`To the Partners of ${data.firmName}`, 'Addressee:'),
      bodyPara(
        `We have audited the financial statements of ${data.firmName}, which comprise the Statement of Financial Position as at ${data.reportDate}, and the Statement of Profit or Loss and Other Comprehensive Income, Statement of Changes in Partners' Equity, and Statement of Cash Flows for the year then ended, and notes to the financial statements, including material accounting policy information.`,
        'Opinion:'
      ),
      bodyPara(
        `In our opinion, the accompanying financial statements present fairly, in all material respects, the financial position of the Firm as at ${data.reportDate}, and its financial performance and its cash flows for the year then ended in accordance with International Financial Reporting Standards (IFRS).`
      ),
      bodyPara(
        'We conducted our audit in accordance with International Standards on Auditing (ISAs). We are independent of the Firm in accordance with the ethical requirements relevant to our audit of the financial statements in Bangladesh.',
        'Basis for Opinion:'
      ),
      bodyPara(
        `${data.auditorFirm} | Engagement Partner: ${data.auditorPartner} | Date: ${data.approvalDate} | Dhaka, Bangladesh`,
        'Signed:'
      ),
      new Paragraph({ children: [new PageBreak()] })
    );
  }

  // SECTION C.1: STATEMENT OF FINANCIAL POSITION (BALANCE SHEET)
  children.push(
    sectionTitle('SECTION C.1 — STATEMENT OF FINANCIAL POSITION (IAS 1)'),
    bodyPara(`As at ${data.reportDate} (Amounts in Bangladeshi Taka)`),
    makeDocxTable(
      [
        'Particulars',
        'Notes',
        `${data.reportDate} (BDT)`,
        `${data.prevDate} (BDT)`,
      ],
      [
        { cells: ['ASSETS — Non-Current Assets', '', '', ''], bold: true, shaded: true },
        ...data.bs.nca.map((i) => ({
          cells: [i.label, i.note, fmtNum(i.curr), fmtNum(i.prev)],
        })),
        {
          cells: [
            'Total Non-Current Assets',
            '',
            fmtNum(data.bs.ncaCurr),
            fmtNum(data.bs.ncaPrev),
          ],
          bold: true,
        },
        { cells: ['ASSETS — Current Assets', '', '', ''], bold: true, shaded: true },
        ...data.bs.ca.map((i) => ({
          cells: [i.label, i.note, fmtNum(i.curr), fmtNum(i.prev)],
        })),
        {
          cells: [
            'Total Current Assets',
            '',
            fmtNum(data.bs.caCurr),
            fmtNum(data.bs.caPrev),
          ],
          bold: true,
        },
        {
          cells: [
            'TOTAL ASSETS',
            '',
            fmtNum(data.bs.totAssetsCurr),
            fmtNum(data.bs.totAssetsPrev),
          ],
          bold: true,
          shaded: true,
        },
        {
          cells: ["EQUITY & LIABILITIES — Partners' Equity (IAS 32)", '', '', ''],
          bold: true,
          shaded: true,
        },
        {
          cells: [
            "Partners' capital and current accounts",
            'Note 12',
            fmtNum(data.bs.eqCurr),
            fmtNum(data.bs.eqPrev),
          ],
          bold: true,
        },
        { cells: ['Non-Current Liabilities', '', '', ''], bold: true, shaded: true },
        ...data.bs.ncl.map((i) => ({
          cells: [i.label, i.note, fmtNum(i.curr), fmtNum(i.prev)],
        })),
        {
          cells: [
            'Total Non-Current Liabilities',
            '',
            fmtNum(data.bs.nclCurr),
            fmtNum(data.bs.nclPrev),
          ],
          bold: true,
        },
        { cells: ['Current Liabilities', '', '', ''], bold: true, shaded: true },
        ...data.bs.cl.map((i) => ({
          cells: [i.label, i.note, fmtNum(i.curr), fmtNum(i.prev)],
        })),
        {
          cells: [
            'Total Current Liabilities',
            '',
            fmtNum(data.bs.clCurr),
            fmtNum(data.bs.clPrev),
          ],
          bold: true,
        },
        {
          cells: [
            'TOTAL EQUITY AND LIABILITIES',
            '',
            fmtNum(data.bs.totEqLiabCurr),
            fmtNum(data.bs.totEqLiabPrev),
          ],
          bold: true,
          shaded: true,
        },
      ],
      2
    ),
    new Paragraph({ children: [new PageBreak()] })
  );

  // SECTION C.2 & C.3: PROFIT OR LOSS & CHANGES IN PARTNERS' EQUITY
  children.push(
    sectionTitle(
      'SECTION C.2 — STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME'
    ),
    bodyPara(`For the Financial Year Ended ${data.reportDate}`),
    makeDocxTable(
      ['Particulars', 'Notes', 'Current Year (BDT)', 'Previous Year (BDT)'],
      [
        {
          cells: [
            data.pnl.revLabel,
            data.pnl.revNote,
            fmtNum(data.pnl.revCurr),
            fmtNum(data.pnl.revPrev),
          ],
          bold: true,
        },
        {
          cells: [
            `Less: ${data.pnl.cosLabel}`,
            data.pnl.cosNote,
            fmtNum(data.pnl.cosCurr, true),
            fmtNum(data.pnl.cosPrev, true),
          ],
        },
        {
          cells: [
            'Gross Profit',
            '',
            fmtNum(data.pnl.gpCurr),
            fmtNum(data.pnl.gpPrev),
          ],
          bold: true,
          shaded: true,
        },
        {
          cells: [
            data.pnl.othIncLabel,
            data.pnl.othIncNote,
            fmtNum(data.pnl.othIncCurr),
            fmtNum(data.pnl.othIncPrev),
          ],
        },
        ...data.pnl.expenses.map((e) => ({
          cells: [
            e.label,
            e.note,
            fmtNum(e.curr, true),
            fmtNum(e.prev, true),
          ],
        })),
        {
          cells: [
            'Operating Profit',
            '',
            fmtNum(data.pnl.opCurr),
            fmtNum(data.pnl.opPrev),
          ],
          bold: true,
          shaded: true,
        },
        {
          cells: [
            data.pnl.finLabel,
            data.pnl.finNote,
            fmtNum(data.pnl.finCurr, true),
            fmtNum(data.pnl.finPrev, true),
          ],
        },
        {
          cells: [
            'Profit Before Tax',
            '',
            fmtNum(data.pnl.pbtCurr),
            fmtNum(data.pnl.pbtPrev),
          ],
          bold: true,
        },
        {
          cells: [
            data.pnl.taxLabel,
            data.pnl.taxNote,
            fmtNum(data.pnl.taxCurr, true),
            fmtNum(data.pnl.taxPrev, true),
          ],
        },
        {
          cells: [
            'Profit for the Year',
            '',
            fmtNum(data.pnl.npCurr),
            fmtNum(data.pnl.npPrev),
          ],
          bold: true,
          shaded: true,
        },
        {
          cells: [
            data.pnl.ociLabel,
            '',
            fmtNum(data.pnl.ociCurr),
            fmtNum(data.pnl.ociPrev),
          ],
        },
        {
          cells: [
            'TOTAL COMPREHENSIVE INCOME FOR THE YEAR',
            '',
            fmtNum(data.pnl.tciCurr),
            fmtNum(data.pnl.tciPrev),
          ],
          bold: true,
          shaded: true,
        },
      ],
      2
    ),
    subHeading("SECTION C.3 — Statement of Changes in Partners' Equity (IAS 1 & IAS 32)"),
    makeDocxTable(
      [
        'Particulars',
        ...data.partners.map((p) => `${p.name} (${p.ratio}%)`),
        'Total Equity (BDT)',
      ],
      [
        {
          cells: [
            'Opening capital balance',
            ...data.partners.map((p) => fmtNum(p.open)),
            fmtNum(data.totals.totOpen),
          ],
        },
        {
          cells: [
            'Additional capital introduced',
            ...data.partners.map((p) => fmtNum(p.add)),
            fmtNum(data.totals.totAdd),
          ],
        },
        {
          cells: [
            'Profit allocated for the year',
            ...data.partners.map((p) => fmtNum(p.prof)),
            fmtNum(data.totals.totProf),
          ],
        },
        {
          cells: [
            "Less: Partners' drawings",
            ...data.partners.map((p) => fmtNum(p.draw, true)),
            fmtNum(data.totals.totDraw, true),
          ],
        },
        {
          cells: [
            'CLOSING CAPITAL BALANCE',
            ...data.partners.map((p) => fmtNum(p.close)),
            fmtNum(data.totals.totClose),
          ],
          bold: true,
          shaded: true,
        },
      ],
      1
    ),
    new Paragraph({ children: [new PageBreak()] })
  );

  // SECTION C.4: STATEMENT OF CASH FLOWS (IAS 7) & NOTES SUMMARY
  children.push(
    sectionTitle('SECTION C.4 — STATEMENT OF CASH FLOWS (IAS 7)'),
    bodyPara(`For the Financial Year Ended ${data.reportDate}`),
    makeDocxTable(
      ['Cash Flow Activity Particulars', 'Current Year (BDT)'],
      [
        {
          cells: ['A. Cash Flows from Operating Activities', ''],
          bold: true,
          shaded: true,
        },
        { cells: ['Cash generated from operations', fmtNum(data.cf.cfOpsGen)] },
        { cells: ['Interest paid', fmtNum(data.pnl.finCurr, true)] },
        { cells: ['Income taxes paid', fmtNum(data.pnl.taxCurr, true)] },
        {
          cells: ['Net Cash from Operating Activities', fmtNum(data.cf.netOps)],
          bold: true,
        },
        {
          cells: ['B. Cash Flows from Investing Activities', ''],
          bold: true,
          shaded: true,
        },
        ...data.cf.investing.map((i) => ({
          cells: [i.label, fmtNum(i.amount)],
        })),
        {
          cells: [
            'Net Cash from / (Used in) Investing Activities',
            fmtNum(data.cf.netInv),
          ],
          bold: true,
        },
        {
          cells: ['C. Cash Flows from Financing Activities', ''],
          bold: true,
          shaded: true,
        },
        {
          cells: ['Capital introduced by partners', fmtNum(data.totals.totAdd)],
        },
        {
          cells: [
            "Partners' withdrawals (drawings)",
            fmtNum(data.totals.totDraw, true),
          ],
        },
        ...data.cf.financing.map((i) => ({
          cells: [i.label, fmtNum(i.amount)],
        })),
        {
          cells: [
            'Net Cash from / (Used in) Financing Activities',
            fmtNum(data.cf.netFin),
          ],
          bold: true,
        },
        {
          cells: [
            'Net Increase / (Decrease) in Cash and Cash Equivalents (A + B + C)',
            fmtNum(data.cf.netCashChange),
          ],
          bold: true,
        },
        {
          cells: [
            'Opening cash and cash equivalents',
            fmtNum(data.cf.cfOpenCash),
          ],
        },
        {
          cells: [
            'CLOSING CASH AND CASH EQUIVALENTS',
            fmtNum(data.cf.cfCloseCash),
          ],
          bold: true,
          shaded: true,
        },
      ],
      1
    ),
    subHeading('SECTION C.5 — Key Schedules to the Financial Statements (Notes 8 to 25)'),
    bodyPara(data.notes8to14.n8.title),
    makeDocxTable(
      ['Asset Category', 'Cost (BDT)', 'Accum. Depreciation (BDT)', 'Carrying Amount (BDT)'],
      data.notes8to14.n8.rows.map((r: any) => ({
        cells: [
          r.cat,
          fmtNum(r.cost),
          fmtNum(r.dep),
          fmtNum(r.cost + r.dep),
        ],
        bold: Boolean(r.isTotal),
        shaded: Boolean(r.isTotal),
      })),
      1
    ),
    ...data.disclosureNotes.map((n) => bodyPara(n.text, `${n.title}:`))
  );

  // SECTION C.6: IAS 12 & S.R.O. 324/2026 TAXATION SCHEDULE
  if (data.tax.includeTax) {
    children.push(
      new Paragraph({ children: [new PageBreak()] }),
      sectionTitle(
        `SECTION C.6 — NOTE 19: CURRENT & DEFERRED TAXATION (${data.tax.taxPeriodTitle})`
      ),
      bodyPara(
        'Statutory Assessment under IAS 12 and S.R.O. No. 324-Ain/Aykor-05/2026 (Income Tax Act 2023):'
      ),
      makeDocxTable(
        [
          'SL',
          'Corporate Tax Assessment Item',
          'Statutory Reference / Rate',
          'Calculation Base',
          'Amount (BDT)',
        ],
        [
          {
            cells: [
              '1',
              'Gross Receipts / Turnover from Contracts',
              'S.R.O. 324/2026',
              data.tax.taxInvoiceRef,
              fmtDec2(data.tax.grossReceipts),
            ],
          },
          {
            cells: [
              '2',
              `Minimum Income Tax @ ${data.tax.slabRatePct.toFixed(2)}% of Gross Receipts (A)`,
              'Section 163(6) / S.R.O. 324',
              `${fmtDec2(data.tax.grossReceipts)} @ ${data.tax.slabRatePct.toFixed(2)}%`,
              fmtDec2(data.tax.minTaxA),
            ],
            bold: true,
          },
          {
            cells: [
              '3',
              'Taxable Business Profit for the Period',
              'Income Statement',
              data.tax.taxPbtNote,
              fmtDec2(data.tax.taxableProfit),
            ],
          },
          {
            cells: [
              '4',
              `Tax Applied on Taxable Profit @ ${data.tax.regTaxRatePct.toFixed(1)}% (B)`,
              `Regular Rate @ ${data.tax.regTaxRatePct.toFixed(2)}%`,
              `${fmtDec2(data.tax.taxableProfit)} @ ${data.tax.regTaxRatePct.toFixed(2)}%`,
              fmtDec2(data.tax.regTaxB),
            ],
            bold: true,
          },
          {
            cells: [
              '—',
              'APPLICABLE CORPORATE INCOME TAX (HIGHER OF A AND B)',
              'Higher of (A) & (B)',
              data.tax.regTaxB >= data.tax.minTaxA
                ? 'Regular Tax (B) Applies'
                : 'Minimum Tax (A) Applies',
              fmtDec2(data.tax.applicableTax),
            ],
            bold: true,
            shaded: true,
          },
          {
            cells: [
              '5',
              'Less: Advance Tax Paid (AIT) / TDS Deducted at Source',
              'Section 163 / AIT Challan',
              data.tax.taxAitRef,
              fmtDec2(data.tax.aitPaid, true),
            ],
          },
          {
            cells: [
              '—',
              data.tax.taxClosingLabel,
              'Net Payable to NBR',
              data.tax.netCurrentTaxPayable === 0
                ? 'Balance Due (TDS covers liability)'
                : 'Net Tax Payable',
              fmtDec2(data.tax.netCurrentTaxPayable),
            ],
            bold: true,
            shaded: true,
          },
        ],
        4
      ),
      bodyPara(data.tax.taxStatutoryNote, 'Note on Statutory Assessment:')
    );
  }

  // SECTION D: PARTNERS' AUTHORIZATION & SIGNATURES
  children.push(
    subHeading("SECTION D — Approval and Authorization of Financial Statements"),
    bodyPara(
      `The partners of ${data.firmName} hereby approve and authorize the issue of the financial statements for the financial year ended ${data.reportDate}, together with all accompanying notes and disclosures, on ${data.approvalDate} in Dhaka, Bangladesh.`
    ),
    makeDocxTable(
      ['Partner Name', 'Designation / Role', 'Date & Place of Authorization'],
      data.partners.map((p) => ({
        cells: [p.name, p.role, `${data.approvalDate} — Dhaka, Bangladesh`],
        bold: true,
      })),
      3
    )
  );

  // SECTION E: NBR FORM IT-116 (2023) & IT-11 COLA
  if (data.it116.includeIt116) {
    children.push(
      new Paragraph({ children: [new PageBreak()] }),
      sectionTitle(
        'SECTION E — STATUTORY INCOME TAX RETURN: NBR FORM IT-116 (2023) & IT-11 COLA'
      ),
      bodyPara(
        `National Board of Revenue (www.nbr.gov.bd) — Return of Income for Taxpayers Other Than Company u/s 2(31) of the Income Tax Act, 2023`
      ),
      makeDocxTable(
        ['Part I — Taxpayer Particular', 'Registered Return Information'],
        [
          { cells: ['1. Name of Taxpayer', data.firmName], bold: true },
          { cells: ['2. 12-Digit Taxpayer TIN', data.it116.tin], bold: true },
          { cells: ['3. Business Identification Number (BIN)', data.it116.bin] },
          {
            cells: [
              '4. Tax Circle & Tax Zone',
              `${data.it116.circle} | ${data.it116.zone}`,
            ],
          },
          { cells: ['5. Assessment / Tax Year', data.it116.taxYear], bold: true },
          {
            cells: [
              '6. Category & 7. Residential Status',
              `${data.it116.category.toUpperCase()} (${data.it116.resident.toUpperCase()})`,
            ],
          },
          {
            cells: [
              '8. Contact Address, Phone & Email',
              `${data.address} | ${data.it116.phone} | ${data.it116.email}`,
            ],
          },
          {
            cells: [
              '9. RJSC Registration No. & 10. Date',
              `${data.regNo} (${data.it116.regDate})`,
            ],
          },
        ],
        99
      ),
      subHeading('Statement of Total Income, Computation of Tax & Tax Payment (IT-116 Page 2)'),
      makeDocxTable(
        ['Serial & Head of Income / Tax Computation', 'Amount in Taka (BDT)'],
        [
          {
            cells: [
              '1. Income from Rent (Schedule 1)',
              fmtNum(data.it116.rentInc),
            ],
          },
          {
            cells: [
              '2. Income from Agriculture (Schedule 2)',
              fmtNum(data.it116.agriInc),
            ],
          },
          {
            cells: [
              '3. Income from Business (Schedule 3 Pre-tax Profit)',
              fmtNum(data.it116.bizInc),
            ],
            bold: true,
          },
          { cells: ['4. Capital Gain', fmtNum(data.it116.capGain)] },
          {
            cells: [
              '5. Income from Financial Assets',
              fmtNum(data.it116.finInc),
            ],
          },
          {
            cells: [
              '7. Foreign Sourced Taxable Income',
              fmtNum(data.it116.foreignInc),
            ],
          },
          {
            cells: [
              '8. TOTAL INCOME (Sum of Serial 1 to 7)',
              fmtNum(data.it116.totalIncome),
            ],
            bold: true,
            shaded: true,
          },
          {
            cells: [
              '9. Income tax imposable on total taxable income',
              fmtNum(data.it116.regularTaxOnTotalInc),
            ],
          },
          {
            cells: ['10. Minimum Tax', fmtNum(data.it116.minTaxOnTds)],
          },
          {
            cells: ['11. Turnover Tax', fmtNum(data.it116.turnoverTaxCalc)],
          },
          {
            cells: [
              '12. Tax Payable (Higher among Serial 9, 10 and 11)',
              fmtNum(data.it116.taxPayableHigher),
            ],
            bold: true,
          },
          {
            cells: [
              '13. Environment Surcharge & 14. Delay Interest/Penalty',
              fmtNum(data.it116.envSurcharge + data.it116.penaltyAmt),
            ],
          },
          {
            cells: [
              '15. TOTAL TAX PAYABLE (12+13+14)',
              fmtNum(data.it116.totalTaxPayable),
            ],
            bold: true,
            shaded: true,
          },
          {
            cells: [
              '16. Tax deducted / collected at source (TDS)',
              fmtNum(data.it116.tdsPaid),
            ],
          },
          {
            cells: [
              '17. Advance tax paid (AIT)',
              fmtNum(data.it116.aitPaidReturn),
            ],
          },
          {
            cells: [
              '19. Balance tax paid along with this return',
              fmtNum(data.it116.balanceTaxWithReturn),
            ],
          },
          {
            cells: [
              '20. TOTAL TAX PAID (Sum of Serial 16 to 19) [IT-11 Cola Certified]',
              fmtNum(data.it116.totalTaxPaid),
            ],
            bold: true,
            shaded: true,
          },
        ],
        1
      ),
      bodyPara(data.it116.docsList, 'Attached Documents (IT-116 Page 6):')
    );
  }

  const doc = new Document({
    creator: data.firmName,
    title: `${data.firmName} - Annual Financial Report & IT-116 (${data.reportDate})`,
    description:
      'IFRS / IAS Compliant Partnership Annual Financial Report & NBR Form IT-116 Return',
    sections: [
      {
        properties: {
          page: {
            size: {
              width: 11906, // A4 width in twips (210mm)
              height: 16838, // A4 height in twips (297mm)
            },
            margin: {
              top: 720,
              right: 850,
              bottom: 720,
              left: 850,
            },
          },
        },
        children,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${data.firmName} - IFRS Financial Report & IT-116 (${data.reportDate}).docx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
