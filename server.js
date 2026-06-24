const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// --- Shared password protection ---
// Set APP_PASSWORD as an environment variable on your host (e.g. Railway).
// If it's not set, the app falls back to a default for local testing only.
const APP_PASSWORD = process.env.APP_PASSWORD || 'changeme';

app.use((req, res, next) => {
  const auth = req.headers.authorization;
  if (auth && auth.startsWith('Basic ')) {
    const decoded = Buffer.from(auth.slice(6), 'base64').toString('utf8');
    const sepIdx = decoded.indexOf(':');
    const pass = sepIdx === -1 ? decoded : decoded.slice(sepIdx + 1);
    if (pass === APP_PASSWORD) return next();
  }
  res.set('WWW-Authenticate', 'Basic realm="ATT Processor"');
  return res.status(401).send('Authentication required.');
});

app.use(express.static(path.join(__dirname, 'public')));

function parseDate(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    return new Date(d.y, d.m - 1, d.d);
  }
  if (val instanceof Date) return val;
  // Handle MM/DD/YYYY explicitly (output of formatDate)
  const m = String(val).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[1]) - 1, parseInt(m[2]));
  return new Date(val);
}

function formatDate(val) {
  const d = parseDate(val);
  if (!d || isNaN(d)) return val || '';
  return `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}/${d.getFullYear()}`;
}

// Format date as MMDDYY for period display
function fmtPeriod(val) {
  const d = parseDate(val);
  if (!d || isNaN(d)) return '';
  return String(d.getMonth()+1).padStart(2,'0') +
         String(d.getDate()).padStart(2,'0') +
         String(d.getFullYear()).slice(-2);
}

function mbToGB(description) {
  const match = String(description).match(/Billable Data Usage\s+([\d.,]+)\s*MB/i);
  if (!match) return '';
  const mb = parseFloat(match[1].replace(',', ''));
  return isNaN(mb) ? '' : (mb / 1024).toFixed(4) + ' GB';
}

function planFromSOC(soc) {
  const s = String(soc || '').toLowerCase();
  if (s.includes('10gb')) return '10GB Voice and Data';
  if (s.includes('5gb'))  return '5GB Voice and Data';
  if (s.includes('1gb'))  return '1GB SmartPhone';
  return '';
}

// Flexible column reader — trims key and compares lowercase
function getCol(row, ...candidates) {
  for (const key of Object.keys(row)) {
    if (candidates.some(c => key.trim().toLowerCase() === c.toLowerCase())) {
      return String(row[key] || '').trim();
    }
  }
  return '';
}

// Determine billing period remark based on where the period falls
// Returns: 'prior' | 'current' | 'next'
function classifyPeriod(dateStart, dateEnd, sortedStandardPeriods) {
  const startKey = fmtPeriod(dateStart);
  const endKey   = fmtPeriod(dateEnd);
  const periodKey = startKey + '-' + endKey;
  const idx = sortedStandardPeriods.findIndex(p => p.key === periodKey);
  if (idx === -1) {
    // Partial/non-standard period — classify by end date matching a standard period
    const endMatch = sortedStandardPeriods.findIndex(p => p.endKey === endKey);
    if (endMatch === -1) return 'prior';
    if (endMatch === sortedStandardPeriods.length - 1) return 'next';
    return 'current';
  }
  if (idx === sortedStandardPeriods.length - 1) return 'next';
  if (idx === sortedStandardPeriods.length - 2) return 'current';
  return 'prior';
}

function processFiles(rawFilePath, prevWorkingPath, nalMasterlistPath) {
  const rawWb = XLSX.readFile(rawFilePath);
  const rawData = XLSX.utils.sheet_to_json(rawWb.Sheets[rawWb.SheetNames[0]], { defval: '' });

  // --- Previous Working Copy lookup ---
  const prevLookup = new Map();
  let prevWorkingWarning = null;
  if (prevWorkingPath) {
    const prevWb = XLSX.readFile(prevWorkingPath);
    const prevData = XLSX.utils.sheet_to_json(prevWb.Sheets[prevWb.SheetNames[0]], { defval: '' });
    const trimmedCols = Object.keys(prevData[0] || {}).map(c => c.trim().toLowerCase());
    const hasCLEC = ['clec name','clec'].some(c => trimmedCols.includes(c));
    const hasSOC  = ['soc-code nrt plan if na consult support','soc code','soc'].some(c => trimmedCols.includes(c));
    const hasPlan = ['plan rate customer rate','plan rate'].some(c => trimmedCols.includes(c));
    if (!hasCLEC || !hasSOC || !hasPlan) {
      prevWorkingWarning = `Wrong file uploaded as "Previous Working Copy". Could not find columns: ${[!hasCLEC && 'CLEC', !hasSOC && 'SOC Code', !hasPlan && 'Plan Rate'].filter(Boolean).join(', ')}. Make sure you are uploading last month's completed Working Copy (not the raw file).`;
    }
    for (const row of prevData) {
      const mdn = String(row['Number'] || '').trim();
      if (!mdn || prevLookup.has(mdn)) continue;
      prevLookup.set(mdn, {
        clec:     getCol(row, 'clec name', 'clec'),
        soc:      getCol(row, 'soc-code nrt plan if na consult support', 'soc code', 'soc'),
        planRate: getCol(row, 'plan rate customer rate', 'plan rate'),
      });
    }
    console.log(`prevLookup: ${prevLookup.size} unique MDNs`);
  }

  // --- NAL Masterlist lookup (Combine- 2026 tab) ---
  const nalLookup = new Map();
  if (nalMasterlistPath) {
    const nalWb = XLSX.readFile(nalMasterlistPath);
    const combineSheet = nalWb.Sheets['Combine- 2026'];
    if (combineSheet) {
      const nalData = XLSX.utils.sheet_to_json(combineSheet, { defval: '' });
      nalData.sort((a, b) => {
        const da = parseDate(a['Request Date']);
        const db = parseDate(b['Request Date']);
        return (!da || !db) ? 0 : da - db;
      });
      for (const row of nalData) {
        const mdn    = String(row['MDN '] || row['MDN'] || '').trim();
        const toPlan = String(row['To ']  || row['To']  || '').trim();
        if (mdn && toPlan) nalLookup.set(mdn, toPlan);
      }
      console.log(`nalLookup: ${nalLookup.size} unique MDNs from Combine- 2026`);
    }
  }

  // --- Identify standard billing periods (start day=16, end day=15) ---
  const standardPeriods = new Map();
  for (const row of rawData) {
    const ds = parseDate(row['Date Start']);
    const de = parseDate(row['Date End']);
    if (!ds || !de) continue;
    if (ds.getDate() === 16 && de.getDate() === 15) {
      const key = fmtPeriod(row['Date Start']) + '-' + fmtPeriod(row['Date End']);
      if (!standardPeriods.has(key)) {
        standardPeriods.set(key, { key, startKey: fmtPeriod(row['Date Start']), endKey: fmtPeriod(row['Date End']), startDate: ds });
      }
    }
  }
  const sortedStandardPeriods = [...standardPeriods.values()].sort((a, b) => a.startDate - b.startDate);

  // --- Process rows ---
  const output = [];
  const unidentifiedMap = new Map();
  const mdnsWithUsage = new Set();
  const billingSummaryMap = new Map(); // key → { description, period, mdns: Set, amount }

  const summary = {
    total: rawData.length,
    identified: 0,
    unidentifiedCLEC: 0,
    unidentifiedSOC: 0,
    unidentifiedPlanRate: 0,
    nalOverrides: 0,
    byClec: {},
    byPlan: {},
    byCategory: { NRT: 0, COG: 0, Other: 0 },
  };

  for (const row of rawData) {
    const mdn            = String(row['Number'] || '').trim();
    const description    = String(row['Description'] || '').trim();
    const chargeCategory = String(row['Charge Category'] || '').trim();
    const amount         = parseFloat(row['Amount']) || 0;

    // VLOOKUP from previous working copy
    const prev = prevLookup.get(mdn) || { clec: '', soc: '', planRate: '' };
    let clec     = prev.clec;
    let soc      = prev.soc;
    let planRate = prev.planRate;

    // NAL Masterlist: populate for ALL MDNs in Combine tab
    let nalMasterlistVal = nalLookup.get(mdn) || '';
    const clecLower = clec.toLowerCase();
    if (clecLower.includes('north american') && nalMasterlistVal) {
      planRate = nalMasterlistVal;
      summary.nalOverrides++;
    }

    // Telzeq fallback from SOC
    if (clecLower.includes('telzeq') && !planRate && soc) {
      planRate = planFromSOC(soc);
    }

    // Billing category & usage conversion
    let billingCategory = 'Other';
    let usageGB = '';
    if (/1 GB (Smartphone|Data Only) Plan - NRT/i.test(description) || /^Credit - 1 GB/i.test(description)) {
      billingCategory = 'NRT';
      summary.byCategory.NRT++;
    } else if (/Billable Data Usage/i.test(description)) {
      billingCategory = 'COG';
      usageGB = mbToGB(description);
      mdnsWithUsage.add(mdn);
      summary.byCategory.COG++;
    } else {
      summary.byCategory.Other++;
    }

    // Missing field tracking — count each type separately
    const missingFields = [];
    if (!clec)     { missingFields.push('CLEC');      summary.unidentifiedCLEC++; }
    if (!soc)      { missingFields.push('SOC Code');  summary.unidentifiedSOC++; }
    if (!planRate && chargeCategory === 'MRC') {
      missingFields.push('Plan Rate');
      summary.unidentifiedPlanRate++;
    }

    if (missingFields.length > 0) {
      if (!unidentifiedMap.has(mdn)) {
        unidentifiedMap.set(mdn, { mdn, name: row['Name'], missing: [] });
      }
      const existing = unidentifiedMap.get(mdn);
      for (const f of missingFields) {
        if (!existing.missing.includes(f)) existing.missing.push(f);
      }
    } else {
      summary.identified++;
    }

    if (clec)     summary.byClec[clec]    = (summary.byClec[clec]    || 0) + 1;
    if (planRate) summary.byPlan[planRate] = (summary.byPlan[planRate] || 0) + 1;

    // --- Billing summary grouping ---
    const normalizedDesc = /Billable Data Usage/i.test(description) ? 'Billable Usage' : description;
    const periodKey = fmtPeriod(row['Date Start']) + '-' + fmtPeriod(row['Date End']);
    const bKey = normalizedDesc + '||' + periodKey;
    if (!billingSummaryMap.has(bKey)) {
      const periodClass = classifyPeriod(row['Date Start'], row['Date End'], sortedStandardPeriods);
      const isCredit = /^Credit/i.test(normalizedDesc);
      let remarks = '';
      if (/Billable Usage/i.test(normalizedDesc)) {
        remarks = 'Optimized';
      } else if (periodClass === 'next') {
        remarks = isCredit ? 'Month In Advance - Credit' : 'Month In Advance';
      } else if (periodClass === 'prior') {
        remarks = 'Credit';
      } else {
        remarks = isCredit ? 'Credit' : 'MRC';
      }
      billingSummaryMap.set(bKey, { description: normalizedDesc, period: periodKey, mdns: new Set(), amount: 0, remarks });
    }
    const bEntry = billingSummaryMap.get(bKey);
    bEntry.mdns.add(mdn);
    bEntry.amount += amount;

    output.push({
      'Statement ID':                             row['Statement ID'],
      'AcctNum':                                  row['AcctNum'],
      'Name':                                     row['Name'],
      'Number':                                   row['Number'],
      'CLEC name':                                clec,
      'SOC-Code NRT Plan if NA consult support':  soc,
      'Plan Rate Customer rate':                  planRate,
      'NAL Masterlist':                           nalMasterlistVal,
      'Billing Category':                         billingCategory,
      'Usage GB':                                 usageGB,
      'Transaction Date':                         formatDate(row['Transaction Date']),
      'Date Start':                               formatDate(row['Date Start']),
      'Date End':                                 formatDate(row['Date End']),
      'Transaction ID':                           row['Transaction ID'],
      'Package ID':                               row['Package ID'],
      'Package Description':                      row['Package Description'],
      'Description':                              description,
      'Amount':                                   row['Amount'],
      'Charge Category':                          chargeCategory,
      'Statement Date':                           formatDate(row['Statement Date']),
      'Due Date':                                 formatDate(row['Due Date']),
      'Time Zone':                                row['Time Zone'],
      'Missing Fields':                           missingFields.join(', '),
    });
  }

  summary.mdnsWithUsage = mdnsWithUsage.size;

  // Convert billing summary map to sorted array
  const periodOrder = { prior: 0, current: 1, next: 2 };
  const billingSummary = [...billingSummaryMap.values()]
    .map(e => ({
      description: e.description,
      period:      e.period,
      mdnCount:    e.mdns.size,
      amount:      parseFloat(e.amount.toFixed(2)),
      remarks:     e.remarks,
    }))
    .sort((a, b) => {
      if (a.period !== b.period) return a.period.localeCompare(b.period);
      return a.description.localeCompare(b.description);
    });

  return { output, summary, unidentifiedList: [...unidentifiedMap.values()], billingSummary, prevWorkingWarning };
}

function buildWorkbook(output, summary, unidentifiedList, billingSummary) {
  const wb = XLSX.utils.book_new();

  // --- Working Copy sheet ---
  const ws = XLSX.utils.json_to_sheet(output);
  const headerStyle = {
    font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Arial', sz: 10 },
    fill: { fgColor: { rgb: '1F4E79' } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  };
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let C = range.s.c; C <= range.e.c; C++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: C });
    if (ws[addr]) ws[addr].s = headerStyle;
  }
  for (let R = 1; R <= range.e.r; R++) {
    const cell = ws[XLSX.utils.encode_cell({ r: R, c: range.e.c })];
    if (cell && cell.v) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[addr]) ws[addr] = { v: '', t: 's' };
        ws[addr].s = { fill: { fgColor: { rgb: 'FFFF00' } }, font: { name: 'Arial', sz: 10 } };
      }
    }
  }
  ws['!cols'] = [
    {wch:12},{wch:8},{wch:30},{wch:14},{wch:22},{wch:38},{wch:22},{wch:22},
    {wch:14},{wch:12},{wch:14},{wch:12},{wch:12},{wch:14},{wch:12},{wch:25},{wch:38},
    {wch:10},{wch:14},{wch:14},{wch:12},{wch:10},{wch:20}
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Working Copy');

  // --- Billing Summary sheet ---
  const totalAmount = parseFloat(billingSummary.reduce((s, e) => s + e.amount, 0).toFixed(2));
  const bRows = [
    ...billingSummary.map(e => ({
      'Description':    e.description,
      'Period Covered': e.period,
      'No. of MDN':     e.mdnCount || '',
      'Amount':         parseFloat(e.amount.toFixed(2)),
      'Remarks':        e.remarks,
    })),
    { 'Description': 'TOTAL', 'Period Covered': '', 'No. of MDN': '', 'Amount': totalAmount, 'Remarks': '' },
  ];
  const wsB = XLSX.utils.json_to_sheet(bRows);
  const bHeaderStyle = { font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: '1F4E79' } } };
  const bRange = XLSX.utils.decode_range(wsB['!ref'] || 'A1:E1');
  for (let C = bRange.s.c; C <= bRange.e.c; C++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: C });
    if (wsB[addr]) wsB[addr].s = bHeaderStyle;
  }
  // Style the total row — bold with top border
  const totalRow = bRows.length; // 1-indexed in Excel = data rows + 1 header
  for (let C = 0; C <= 4; C++) {
    const addr = XLSX.utils.encode_cell({ r: totalRow, c: C });
    if (!wsB[addr]) wsB[addr] = { v: '', t: 's' };
    wsB[addr].s = { font: { bold: true, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: 'D9E1F2' } } };
  }
  wsB['!cols'] = [{wch:40},{wch:16},{wch:12},{wch:14},{wch:26}];
  XLSX.utils.book_append_sheet(wb, wsB, 'Billing Summary');

  // --- Unidentified MDNs sheet ---
  if (unidentifiedList.length > 0) {
    const rows = unidentifiedList.map(u => ({
      MDN: u.mdn, Name: u.name,
      'Missing Fields': u.missing.join(', '),
      'Action Required': 'Send to Gopal/Camilo for identification',
    }));
    const wsU = XLSX.utils.json_to_sheet(rows);
    wsU['!cols'] = [{wch:14},{wch:30},{wch:25},{wch:42}];
    XLSX.utils.book_append_sheet(wb, wsU, 'Unidentified MDNs');
  }

  // --- Summary sheet ---
  const totalUnid = new Set([
    ...Array(summary.unidentifiedCLEC).fill(0).map((_,i)=>'c'+i),
  ]).size; // just use the counts directly
  const summaryRows = [
    { Metric: 'Total Rows Processed',           Value: summary.total },
    { Metric: 'Fully Identified Rows',           Value: summary.identified },
    { Metric: '',                                Value: '' },
    { Metric: '--- Unidentified Breakdown ---',  Value: '' },
    { Metric: 'Unidentified CLEC (rows)',        Value: summary.unidentifiedCLEC },
    { Metric: 'Unidentified SOC Code (rows)',    Value: summary.unidentifiedSOC },
    { Metric: 'Unidentified Plan Rate (MRC rows)', Value: summary.unidentifiedPlanRate },
    { Metric: '',                                Value: '' },
    { Metric: '--- Plan Rate & Usage ---',       Value: '' },
    { Metric: 'MDNs with Usage',                 Value: summary.mdnsWithUsage },
    { Metric: 'NAL Plan Overrides Applied',      Value: summary.nalOverrides },
    { Metric: '',                                Value: '' },
    { Metric: '--- By Billing Category ---',     Value: '' },
    { Metric: 'NRT',                             Value: summary.byCategory.NRT },
    { Metric: 'COG (Billable Usage)',            Value: summary.byCategory.COG },
    { Metric: 'Other',                           Value: summary.byCategory.Other },
    { Metric: '',                                Value: '' },
    { Metric: '--- By CLEC ---',                Value: '' },
    ...Object.entries(summary.byClec).sort((a,b)=>b[1]-a[1]).map(([k,v])=>({ Metric: k, Value: v })),
    { Metric: '',                                Value: '' },
    { Metric: '--- By Plan Rate ---',            Value: '' },
    ...Object.entries(summary.byPlan).sort((a,b)=>b[1]-a[1]).map(([k,v])=>({ Metric: k, Value: v })),
  ];
  const wsS = XLSX.utils.json_to_sheet(summaryRows);
  wsS['!cols'] = [{wch:38},{wch:12}];
  XLSX.utils.book_append_sheet(wb, wsS, 'Summary');

  return wb;
}

function cleanupFiles(...files) {
  for (const f of files) {
    if (f && fs.existsSync(f.path)) fs.unlinkSync(f.path);
  }
}

const uploadFields = upload.fields([
  { name: 'rawFile',       maxCount: 1 },
  { name: 'prevWorking',   maxCount: 1 },
  { name: 'nalMasterlist', maxCount: 1 },
]);

app.post('/api/summary', uploadFields, (req, res) => {
  const rawFile     = req.files['rawFile']?.[0];
  const prevWorking = req.files['prevWorking']?.[0];
  const nalFile     = req.files['nalMasterlist']?.[0];
  if (!rawFile) return res.status(400).json({ error: 'Raw file is required.' });
  try {
    const { summary, unidentifiedList, billingSummary, prevWorkingWarning } =
      processFiles(rawFile.path, prevWorking?.path, nalFile?.path);
    cleanupFiles(rawFile, prevWorking, nalFile);
    res.json({ summary, unidentifiedList, billingSummary, prevWorkingWarning });
  } catch (err) {
    console.error(err);
    cleanupFiles(rawFile, prevWorking, nalFile);
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/process', uploadFields, (req, res) => {
  const rawFile     = req.files['rawFile']?.[0];
  const prevWorking = req.files['prevWorking']?.[0];
  const nalFile     = req.files['nalMasterlist']?.[0];
  if (!rawFile) return res.status(400).json({ error: 'Raw file is required.' });
  try {
    const { output, summary, unidentifiedList, billingSummary } =
      processFiles(rawFile.path, prevWorking?.path, nalFile?.path);
    const wbOut = buildWorkbook(output, summary, unidentifiedList, billingSummary);
    const outPath = path.join('uploads', `working_copy_${Date.now()}.xlsx`);
    XLSX.writeFile(wbOut, outPath, { bookType: 'xlsx', type: 'binary', cellStyles: true });
    cleanupFiles(rawFile, prevWorking, nalFile);
    res.download(outPath, 'Working_Copy_Processed.xlsx', () => {
      if (fs.existsSync(outPath)) fs.unlinkSync(outPath);
    });
  } catch (err) {
    console.error(err);
    cleanupFiles(rawFile, prevWorking, nalFile);
    res.status(500).json({ error: err.message });
  }
});

// ============================================================
// STEP 2 — FINALIZE + CUSTOMER FILE GENERATION
// ============================================================

const TELZEQ_RATES = {
  '1GB SmartPhone':        7.00,
  '1GB Voice and Data':    7.00,
  '1GB-HCD':              15.00,
  '10GB Voice and Data':  25.50,
  'Small Data Plan 50MB':  5.00,
};

const NAL_RATES = {
  '1GB Voice and Data': { mia: 8.00,  passthrough: 1.51 },
  '5GB Voice and Data': { mia: 14.00, passthrough: 1.51 },
};

const NAL_OVERAGE_PER_GB = 11;
const NAL_OVERAGE_CAP_GB = 3;

function parseGBUsage(val) {
  const m = String(val || '').match(/([\d.]+)\s*GB/i);
  return m ? parseFloat(m[1]) : 0;
}

// Number of calendar days in the month containing `date`
function daysInMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
}

// Short date string MMDDYY for summary period labels
function shortDate(d) {
  return String(d.getMonth()+1).padStart(2,'0') + String(d.getDate()).padStart(2,'0') + String(d.getFullYear()).slice(-2);
}
function periodShort(pStart, pEnd) {
  return `${shortDate(pStart)}-${shortDate(pEnd)}`;
}

// Apply $#,##0.00 number format to cells in specified columns (matched by header name)
function applyDollarFormat(ws, dollarColNames) {
  if (!ws['!ref']) return;
  const range = XLSX.utils.decode_range(ws['!ref']);
  const dollarCols = new Set();
  for (let C = range.s.c; C <= range.e.c; C++) {
    const h = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
    if (h && dollarColNames.has(String(h.v || ''))) dollarCols.add(C);
  }
  for (let R = 1; R <= range.e.r; R++) {
    for (const C of dollarCols) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell && typeof cell.v === 'number') cell.z = '$#,##0.00';
    }
  }
}

// Append a TOTAL row summing all numeric columns in `rows`, label first col as 'TOTAL'
function withTotalsRow(rows) {
  if (!rows.length) return rows;
  const keys = Object.keys(rows[0]);
  const totals = {};
  for (const k of keys) {
    const s = rows.reduce((acc, r) => acc + (typeof r[k] === 'number' ? r[k] : 0), 0);
    totals[k] = s !== 0 ? parseFloat(s.toFixed(2)) : (k === keys[0] ? 'TOTAL' : '');
  }
  totals[keys[0]] = 'TOTAL';
  return [...rows, totals];
}

// Prorated amount: rate / daysInEndMonth × billableDays (End - Start + 1)
function prorateAmount(rate, startDate, endDate) {
  const billDays = Math.round((endDate - startDate) / 86400000) + 1;
  const dims     = daysInMonth(endDate);
  return { amount: parseFloat((rate / dims * billDays).toFixed(2)), days: billDays };
}

// Maps any date to the 16th–15th billing period that contains it.
// Day ≥ 16 → period starts on the 16th of that month, ends on 15th of next month.
// Day ≤ 15 → period starts on the 16th of prior month, ends on 15th of this month.
function periodForDate(date) {
  if (!date) return null;
  const day = date.getDate();
  let pStart, pEnd;
  if (day >= 16) {
    pStart = new Date(date.getFullYear(), date.getMonth(), 16);
    pEnd   = new Date(date.getFullYear(), date.getMonth() + 1, 15);
  } else {
    pStart = new Date(date.getFullYear(), date.getMonth() - 1, 16);
    pEnd   = new Date(date.getFullYear(), date.getMonth(), 15);
  }
  return { pStart, pEnd };
}

function tabName(pStart, pEnd) {
  return `${formatDate(pStart).replace(/\//g,'.')} - ${formatDate(pEnd).replace(/\//g,'.')}`;
}

// From processed WC rows (dates already MM/DD/YYYY), find MIA and current billing periods
function identifyBillingPeriods(rows) {
  const periodMap = new Map();
  for (const row of rows) {
    const dsStr = String(row['Date Start'] || '').trim();
    const deStr = String(row['Date End']   || '').trim();
    const ds = parseDate(dsStr);
    if (!ds || ds.getDate() !== 16) continue;
    if (!periodMap.has(dsStr)) {
      periodMap.set(dsStr, { start: ds, end: parseDate(deStr), startStr: dsStr, endStr: deStr });
    }
  }
  const sorted = [...periodMap.values()].sort((a, b) => a.start - b.start);
  return {
    miaPeriod:     sorted.length >= 1 ? sorted[sorted.length - 1] : null,
    currentPeriod: sorted.length >= 2 ? sorted[sorted.length - 2] : null,
  };
}

function mergeWithGopal(processedWCPath, gopalFilePath) {
  const wb     = XLSX.readFile(processedWCPath);
  const wsName = wb.SheetNames.find(s => /working.?copy/i.test(s)) || wb.SheetNames[0];
  const rows   = XLSX.utils.sheet_to_json(wb.Sheets[wsName], { defval: '' });

  const gopalLookup = new Map();
  if (gopalFilePath) {
    const gWb   = XLSX.readFile(gopalFilePath);
    const gData = XLSX.utils.sheet_to_json(gWb.Sheets[gWb.SheetNames[0]], { defval: '' });
    for (const row of gData) {
      const mdn = String(row['Number'] || row['MDN'] || '').trim();
      if (!mdn) continue;
      gopalLookup.set(mdn, {
        clec:     getCol(row, 'clec name', 'clec'),
        soc:      getCol(row, 'soc-code nrt plan if na consult support', 'soc code', 'soc'),
        planRate: getCol(row, 'plan rate customer rate', 'plan rate', 'plan'),
      });
    }
    console.log(`Gopal lookup: ${gopalLookup.size} MDNs`);
  }

  const merged = rows.map(row => {
    const mdn   = String(row['Number'] || '').trim();
    const gopal = gopalLookup.get(mdn) || {};
    return {
      ...row,
      'CLEC name':                               row['CLEC name']                               || gopal.clec     || '',
      'SOC-Code NRT Plan if NA consult support': row['SOC-Code NRT Plan if NA consult support'] || gopal.soc      || '',
      'Plan Rate Customer rate':                 row['Plan Rate Customer rate']                  || gopal.planRate || '',
      'Missing Fields': '',
    };
  });

  // For NAL/EBBP Badger MIA rows only: override Plan Rate with NAL Masterlist if it has a real value
  const { miaPeriod } = identifyBillingPeriods(merged);
  if (miaPeriod) {
    for (const row of merged) {
      const clecL = String(row['CLEC name'] || '').toLowerCase();
      if (!clecL.includes('north american') && !clecL.includes('ebbp') && !clecL.includes('badger')) continue;

      const ds = parseDate(row['Date Start']);
      if (!ds || ds.getTime() !== miaPeriod.start.getTime()) continue;

      const nalPlan = String(row['NAL Masterlist'] || '').trim();
      if (!nalPlan || nalPlan.toLowerCase() === 'n/a') continue;

      row['Plan Rate Customer rate'] = nalPlan;
    }
  }

  return merged;
}

function buildFinalizedWCWorkbook(mergedRows, originalWCPath) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(mergedRows);
  const headerStyle = {
    font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Arial', sz: 10 },
    fill: { fgColor: { rgb: '1F4E79' } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  };
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  for (let C = range.s.c; C <= range.e.c; C++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: C });
    if (ws[addr]) ws[addr].s = headerStyle;
  }
  ws['!cols'] = [
    {wch:12},{wch:8},{wch:30},{wch:14},{wch:22},{wch:38},{wch:22},{wch:22},
    {wch:14},{wch:12},{wch:14},{wch:12},{wch:12},{wch:14},{wch:12},{wch:25},{wch:38},
    {wch:10},{wch:14},{wch:14},{wch:12},{wch:10},{wch:20},
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Working Copy');
  // Copy Billing Summary sheet from the original processed Working Copy (Step 1 output)
  if (originalWCPath) {
    try {
      const origWb   = XLSX.readFile(originalWCPath);
      const bsSheet  = origWb.Sheets['Billing Summary'];
      if (bsSheet) XLSX.utils.book_append_sheet(wb, bsSheet, 'Billing Summary');
    } catch (e) { /* missing sheet is non-fatal */ }
  }
  return wb;
}

function buildTelzeqWorkbook(mergedRows, hcdMDNs = new Set()) {
  const { miaPeriod } = identifyBillingPeriods(mergedRows);
  if (!miaPeriod) throw new Error('Cannot identify MIA billing period from the Working Copy.');

  const cpStart = new Date(miaPeriod.start.getFullYear(), miaPeriod.start.getMonth() - 1, 16);
  const cpEnd   = new Date(miaPeriod.start.getFullYear(), miaPeriod.start.getMonth(),     15);
  const cpTab   = tabName(cpStart, cpEnd);
  const miaTab  = tabName(miaPeriod.start, miaPeriod.end);
  const cpShort  = periodShort(cpStart, cpEnd);
  const miaShort = periodShort(miaPeriod.start, miaPeriod.end);

  const cpRows = [], miaRows = [];
  const cpPlanAgg  = new Map(); // plan → { mdns: Set, total }
  const miaPlanAgg = new Map();

  for (const row of mergedRows) {
    const clec = String(row['CLEC name'] || '').trim();
    if (!clec.toLowerCase().includes('telzeq')) continue;

    const mdn  = String(row['Number'] || '').trim();
    const desc = String(row['Description'] || '').trim();
    const amt  = parseFloat(row['Amount']) || 0;
    const cat  = String(row['Charge Category'] || '').trim();
    const plan = hcdMDNs.has(mdn) ? '1GB-HCD' : String(row['Plan Rate Customer rate'] || '').trim();
    if (!mdn || cat === 'Usage') continue;

    const ds = parseDate(row['Date Start']);
    const de = parseDate(row['Date End']);
    if (!ds || !de) continue;

    const period = periodForDate(ds);
    const isMIA  = period.pStart.getTime() === miaPeriod.start.getTime();
    const isCP   = period.pStart.getTime() === cpStart.getTime();
    if (!isMIA && !isCP) continue;

    const isCredit     = /^credit/i.test(desc) || amt < 0;
    const isFullPeriod = ds.getTime() === period.pStart.getTime();
    const rate         = TELZEQ_RATES[plan] || 0;

    let amount;
    if (isFullPeriod) {
      amount = isCredit ? -rate : rate;
    } else {
      const pr = prorateAmount(rate, ds, de);
      amount = isCredit ? -pr.amount : pr.amount;
    }
    const roundedAmt = parseFloat(amount.toFixed(2));

    const outRow = {
      MDN: mdn, Name: String(row['Name'] || ''), 'CLEC Name': clec, 'Rate Plan': plan,
      'Date Start': formatDate(row['Date Start']), 'Date End': formatDate(row['Date End']),
    };
    if (!isFullPeriod) outRow['Days'] = Math.round((de - ds) / 86400000) + 1;

    if (isMIA) {
      outRow['MIA']   = roundedAmt;
      outRow['Total'] = roundedAmt;
      miaRows.push(outRow);
      const a = miaPlanAgg.get(plan) || { mdns: new Set(), total: 0 };
      a.mdns.add(mdn); a.total += roundedAmt;
      miaPlanAgg.set(plan, a);
    } else {
      outRow['MRC']   = roundedAmt;
      outRow['Total'] = roundedAmt;
      cpRows.push(outRow);
      const a = cpPlanAgg.get(plan) || { mdns: new Set(), total: 0 };
      a.mdns.add(mdn); a.total += roundedAmt;
      cpPlanAgg.set(plan, a);
    }
  }

  const wb = XLSX.utils.book_new();
  let grandTotal = 0;

  if (cpRows.length > 0) {
    const ws = XLSX.utils.json_to_sheet(withTotalsRow(cpRows));
    ws['!cols'] = [{wch:14},{wch:30},{wch:20},{wch:24},{wch:12},{wch:12},{wch:6},{wch:12},{wch:12}];
    applyDollarFormat(ws, new Set(['MRC', 'Total']));
    XLSX.utils.book_append_sheet(wb, ws, cpTab.slice(0, 31));
    grandTotal += [...cpPlanAgg.values()].reduce((s, a) => s + a.total, 0);
  }

  if (miaRows.length > 0) {
    const ws = XLSX.utils.json_to_sheet(withTotalsRow(miaRows));
    ws['!cols'] = [{wch:14},{wch:30},{wch:20},{wch:24},{wch:12},{wch:12},{wch:12},{wch:12}];
    applyDollarFormat(ws, new Set(['MIA', 'Total']));
    XLSX.utils.book_append_sheet(wb, ws, miaTab.slice(0, 31));
    grandTotal += [...miaPlanAgg.values()].reduce((s, a) => s + a.total, 0);
  }

  // Build new-format Summary
  const sumRows = [{ Description: 'Telzeq', Period: '', 'No. of MDN': '', Amount: '', Type: '' }];
  for (const [plan, agg] of cpPlanAgg) {
    sumRows.push({ Description: plan, Period: cpShort, 'No. of MDN': agg.mdns.size, Amount: parseFloat(agg.total.toFixed(2)), Type: 'MRC' });
  }
  for (const [plan, agg] of miaPlanAgg) {
    sumRows.push({ Description: plan, Period: miaShort, 'No. of MDN': agg.mdns.size, Amount: parseFloat(agg.total.toFixed(2)), Type: 'Month In Advance' });
  }
  sumRows.push({ Description: '', Period: '', 'No. of MDN': '', Amount: parseFloat(grandTotal.toFixed(2)), Type: '' });

  const wsSummary = XLSX.utils.json_to_sheet(sumRows);
  wsSummary['!cols'] = [{ wch: 26 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 20 }];
  applyDollarFormat(wsSummary, new Set(['Amount']));
  wb.SheetNames.unshift('Summary');
  wb.Sheets['Summary'] = wsSummary;

  return wb;
}

function buildNALWorkbook(mergedRows) {
  const { miaPeriod } = identifyBillingPeriods(mergedRows);
  if (!miaPeriod) throw new Error('Cannot identify MIA billing period from the Working Copy.');

  const cpStart  = new Date(miaPeriod.start.getFullYear(), miaPeriod.start.getMonth() - 1, 16);
  const cpEnd    = new Date(miaPeriod.start.getFullYear(), miaPeriod.start.getMonth(),     15);
  const cpTab    = tabName(cpStart, cpEnd);
  const miaTab   = tabName(miaPeriod.start, miaPeriod.end);
  const cpShort  = periodShort(cpStart, cpEnd);
  const miaShort = periodShort(miaPeriod.start, miaPeriod.end);

  // Pass 1: collect usage GB and identify MIA credit MDNs
  const usageByMDN     = new Map();
  const usagePlanByMDN = new Map();
  const usageInfoByMDN = new Map();
  const miaMDNInfo     = new Map();
  const miaCreditMDNs  = new Set();

  for (const row of mergedRows) {
    const clecL = String(row['CLEC name'] || '').toLowerCase();
    if (!clecL.includes('north american') && !clecL.includes('ebbp') && !clecL.includes('badger')) continue;
    const mdn = String(row['Number'] || '').trim();
    if (!mdn) continue;

    const cat  = String(row['Charge Category'] || '').trim();
    const desc = String(row['Description'] || '').trim();
    const plan = String(row['Plan Rate Customer rate'] || '').trim();

    const gb = parseGBUsage(row['Usage GB']);
    if (gb > 0) {
      usageByMDN.set(mdn, Math.max(usageByMDN.get(mdn) || 0, gb));
      if (!usagePlanByMDN.has(mdn) && plan) usagePlanByMDN.set(mdn, plan);
      if (!usageInfoByMDN.has(mdn)) {
        usageInfoByMDN.set(mdn, { name: String(row['Name']||''), clec: String(row['CLEC name']||'').trim(), plan });
      }
    }

    const ds = parseDate(row['Date Start']);
    if (!ds) continue;
    const period = periodForDate(ds);
    if (period.pStart.getTime() !== miaPeriod.start.getTime()) continue;
    if (cat === 'Usage') continue;

    if (/^credit/i.test(desc)) {
      miaCreditMDNs.add(mdn);
    } else {
      if (!miaMDNInfo.has(mdn)) {
        miaMDNInfo.set(mdn, { name: String(row['Name']||''), clec: String(row['CLEC name']||'').trim(), plan });
      }
    }
  }

  const cpSheetRows = [], miaSheetRows = [];
  const cpPartialPlanAgg = new Map(); // plan → { mdns: Set, total } for partial MRC rows (MRC+PT combined)
  const ovgAgg           = { mdns: new Set(), total: 0 };
  const miaPlanAgg       = new Map(); // plan → { mdns: Set, total } for MIA rows (MIA+PT combined)

  // Pass 2: build billing rows
  for (const row of mergedRows) {
    const clec  = String(row['CLEC name'] || '').trim();
    const clecL = clec.toLowerCase();
    if (!clecL.includes('north american') && !clecL.includes('ebbp') && !clecL.includes('badger')) continue;

    const mdn  = String(row['Number'] || '').trim();
    const desc = String(row['Description'] || '').trim();
    const cat  = String(row['Charge Category'] || '').trim();
    const plan = String(row['Plan Rate Customer rate'] || '').trim();
    if (!mdn || cat === 'Usage') continue;

    const ds = parseDate(row['Date Start']);
    const de = parseDate(row['Date End']);
    if (!ds || !de) continue;

    const period       = periodForDate(ds);
    const isMIA        = period.pStart.getTime() === miaPeriod.start.getTime();
    const isCP         = period.pStart.getTime() === cpStart.getTime();
    if (!isMIA && !isCP) continue;

    const isCredit     = /^credit/i.test(desc);
    const isFullPeriod = ds.getTime() === period.pStart.getTime();
    const rateInfo     = NAL_RATES[plan] || { mia: 0, passthrough: 1.51 };

    if (isMIA) {
      if (miaCreditMDNs.has(mdn)) continue;
      if (isCredit) continue;

      let miaAmt = isFullPeriod ? rateInfo.mia : prorateAmount(rateInfo.mia, ds, de).amount;
      const ptAmt = rateInfo.passthrough;
      const total = parseFloat((miaAmt + ptAmt).toFixed(2));

      const a = miaPlanAgg.get(plan) || { mdns: new Set(), total: 0 };
      a.mdns.add(mdn); a.total += total;
      miaPlanAgg.set(plan, a);

      const outRow = {
        MDN: mdn, Name: String(row['Name']||''), 'CLEC Name': clec, 'Rate Plan': plan,
        'Date Start': formatDate(row['Date Start']), 'Date End': formatDate(row['Date End']),
      };
      if (!isFullPeriod) outRow['Days'] = Math.round((de - ds) / 86400000) + 1;
      outRow['MIA']         = parseFloat(miaAmt.toFixed(2));
      outRow['Passthrough'] = parseFloat(ptAmt.toFixed(2));
      outRow['Total']       = total;
      miaSheetRows.push(outRow);

    } else {
      // CP tab: partial rows only → prorated MRC + $1.51 PT
      if (!isFullPeriod) {
        const pr     = prorateAmount(rateInfo.mia, ds, de);
        const sign   = isCredit ? -1 : 1;
        const mrcAmt = parseFloat((sign * pr.amount).toFixed(2));
        const ptAmt  = parseFloat((sign * rateInfo.passthrough).toFixed(2));
        const total  = parseFloat((mrcAmt + ptAmt).toFixed(2));

        const a = cpPartialPlanAgg.get(plan) || { mdns: new Set(), total: 0 };
        a.mdns.add(mdn); a.total += total;
        cpPartialPlanAgg.set(plan, a);

        cpSheetRows.push({
          MDN: mdn, Name: String(row['Name']||''), 'CLEC Name': clec, 'Rate Plan': plan,
          'Date Start': formatDate(row['Date Start']), 'Date End': formatDate(row['Date End']),
          Days: pr.days, MRC: mrcAmt, Passthrough: ptAmt, Total: total,
        });
      }
    }
  }

  // Pass 3: CP tab — one full-period row per MDN with usage (overage)
  for (const mdn of usageByMDN.keys()) {
    if (miaCreditMDNs.has(mdn)) continue;
    const info       = miaMDNInfo.get(mdn) || usageInfoByMDN.get(mdn);
    if (!info) continue;
    const usageGB    = usageByMDN.get(mdn) || 0;
    const cpPlan     = usagePlanByMDN.get(mdn) || info.plan;
    const is1GB      = /1\s*gb/i.test(cpPlan);
    const overageGB  = is1GB && usageGB > 1 ? Math.min(usageGB - 1, NAL_OVERAGE_CAP_GB) : 0;
    const overageAmt = parseFloat((overageGB * NAL_OVERAGE_PER_GB).toFixed(2));

    if (overageAmt > 0) {
      ovgAgg.mdns.add(mdn);
      ovgAgg.total += overageAmt;
    }

    cpSheetRows.push({
      MDN: mdn, Name: info.name, 'CLEC Name': info.clec, 'Rate Plan': cpPlan,
      'Date Start': formatDate(cpStart), 'Date End': formatDate(cpEnd),
      'GB Usages': usageGB > 0 ? parseFloat(usageGB.toFixed(4)) : '',
      Overage: overageAmt || '', Total: overageAmt || '',
    });
  }

  const wb = XLSX.utils.book_new();

  {
    const ws = XLSX.utils.json_to_sheet(cpSheetRows.length ? withTotalsRow(cpSheetRows) : [{}]);
    ws['!cols'] = [{wch:14},{wch:30},{wch:24},{wch:24},{wch:12},{wch:12},{wch:6},{wch:12},{wch:12},{wch:10},{wch:12}];
    applyDollarFormat(ws, new Set(['MRC', 'Passthrough', 'Overage', 'Total']));
    XLSX.utils.book_append_sheet(wb, ws, cpTab.slice(0, 31));
  }

  if (miaSheetRows.length > 0) {
    const ws = XLSX.utils.json_to_sheet(withTotalsRow(miaSheetRows));
    ws['!cols'] = [{wch:14},{wch:30},{wch:24},{wch:24},{wch:12},{wch:12},{wch:6},{wch:12},{wch:12},{wch:12}];
    applyDollarFormat(ws, new Set(['MIA', 'Passthrough', 'Total']));
    XLSX.utils.book_append_sheet(wb, ws, miaTab.slice(0, 31));
  }

  // Build new-format Summary
  const cpPartialTotal  = [...cpPartialPlanAgg.values()].reduce((s, a) => s + a.total, 0);
  const miaTotal        = [...miaPlanAgg.values()].reduce((s, a) => s + a.total, 0);
  const grandTotal      = parseFloat((ovgAgg.total + cpPartialTotal + miaTotal).toFixed(2));

  const sumRows = [{ Description: 'NAL', Period: '', 'No. of MDN': 'No. of MDN', Amount: 'Amount', Type: '' }];

  if (ovgAgg.total > 0) {
    sumRows.push({
      Description: `${cpShort} Overage`, Period: cpShort,
      'No. of MDN': ovgAgg.mdns.size, Amount: parseFloat(ovgAgg.total.toFixed(2)), Type: 'Overage',
    });
  }
  for (const [plan, agg] of cpPartialPlanAgg) {
    sumRows.push({
      Description: plan, Period: cpShort,
      'No. of MDN': agg.mdns.size, Amount: parseFloat(agg.total.toFixed(2)), Type: 'MRC',
    });
  }
  for (const [plan, agg] of miaPlanAgg) {
    sumRows.push({
      Description: plan, Period: miaShort,
      'No. of MDN': agg.mdns.size, Amount: parseFloat(agg.total.toFixed(2)), Type: 'Month In Advance',
    });
  }
  sumRows.push({ Description: '', Period: '', 'No. of MDN': '', Amount: grandTotal, Type: '' });

  const wsSummary = XLSX.utils.json_to_sheet(sumRows);
  wsSummary['!cols'] = [{ wch: 26 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 20 }];
  applyDollarFormat(wsSummary, new Set(['Amount']));
  wb.SheetNames.unshift('Summary');
  wb.Sheets['Summary'] = wsSummary;

  return wb;
}

const step2Upload = upload.fields([
  { name: 'processedWC',   maxCount: 1 },
  { name: 'gopalFile',     maxCount: 1 },
  { name: 'telzeqPrevFile', maxCount: 1 },
]);

app.post('/api/step2/finalized', step2Upload, (req, res) => {
  const processedWC = req.files?.['processedWC']?.[0];
  const gopalFile   = req.files?.['gopalFile']?.[0];
  if (!processedWC) return res.status(400).json({ error: 'Processed Working Copy is required.' });
  try {
    const merged  = mergeWithGopal(processedWC.path, gopalFile?.path);
    const wb      = buildFinalizedWCWorkbook(merged, processedWC.path);
    const outPath = path.join('uploads', `finalized_wc_${Date.now()}.xlsx`);
    XLSX.writeFile(wb, outPath, { bookType: 'xlsx', type: 'binary', cellStyles: true });
    cleanupFiles(processedWC, gopalFile);
    res.download(outPath, 'Finalized_Working_Copy.xlsx', () => { if (fs.existsSync(outPath)) fs.unlinkSync(outPath); });
  } catch (err) {
    console.error(err);
    cleanupFiles(processedWC, gopalFile);
    res.status(500).json({ error: err.message });
  }
});

function extractHCDMDNs(filePath) {
  const hcdMDNs = new Set();
  if (!filePath) return hcdMDNs;
  const wb = XLSX.readFile(filePath);
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
    for (const row of rows) {
      const mdn  = String(row['MDN'] || '').trim();
      const plan = String(row['RATE PLAN'] || row['Rate Plan'] || row['Plan'] || '').trim();
      if (mdn && /hcd/i.test(plan)) hcdMDNs.add(mdn);
    }
  }
  return hcdMDNs;
}

app.post('/api/step2/telzeq', step2Upload, (req, res) => {
  const processedWC    = req.files?.['processedWC']?.[0];
  const gopalFile      = req.files?.['gopalFile']?.[0];
  const telzeqPrevFile = req.files?.['telzeqPrevFile']?.[0];
  if (!processedWC) return res.status(400).json({ error: 'Processed Working Copy is required.' });
  try {
    const hcdMDNs = extractHCDMDNs(telzeqPrevFile?.path);
    const merged  = mergeWithGopal(processedWC.path, gopalFile?.path);
    const wb      = buildTelzeqWorkbook(merged, hcdMDNs);
    const outPath = path.join('uploads', `telzeq_${Date.now()}.xlsx`);
    XLSX.writeFile(wb, outPath, { bookType: 'xlsx', type: 'binary', cellStyles: true });
    cleanupFiles(processedWC, gopalFile, telzeqPrevFile);
    res.download(outPath, 'Telzeq_Customer_File.xlsx', () => { if (fs.existsSync(outPath)) fs.unlinkSync(outPath); });
  } catch (err) {
    console.error(err);
    cleanupFiles(processedWC, gopalFile, telzeqPrevFile);
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/step2/nal', step2Upload, (req, res) => {
  const processedWC = req.files?.['processedWC']?.[0];
  const gopalFile   = req.files?.['gopalFile']?.[0];
  if (!processedWC) return res.status(400).json({ error: 'Processed Working Copy is required.' });
  try {
    const merged  = mergeWithGopal(processedWC.path, gopalFile?.path);
    const wb      = buildNALWorkbook(merged);
    const outPath = path.join('uploads', `nal_${Date.now()}.xlsx`);
    XLSX.writeFile(wb, outPath, { bookType: 'xlsx', type: 'binary', cellStyles: true });
    cleanupFiles(processedWC, gopalFile);
    res.download(outPath, 'NAL_Customer_File.xlsx', () => { if (fs.existsSync(outPath)) fs.unlinkSync(outPath); });
  } catch (err) {
    console.error(err);
    cleanupFiles(processedWC, gopalFile);
    res.status(500).json({ error: err.message });
  }
});

if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`AT&T Data Processor running at http://localhost:${PORT}`));
