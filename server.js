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

// ── Flag helpers ──────────────────────────────────────────────────────────────

function isUsageRow(row) {
  return String(row['Billing Category'] || '').trim() === 'COG';
}

function isCreditRow(row) {
  return parseFloat(row['Amount']) < 0 || /^Credit/i.test(String(row['Description'] || ''));
}

function periodsOverlap(startA, endA, startB, endB) {
  if (!startA || !endA || !startB || !endB) return false;
  const sA = parseDate(startA), eA = parseDate(endA);
  const sB = parseDate(startB), eB = parseDate(endB);
  if (!sA || !eA || !sB || !eB || isNaN(sA) || isNaN(eA) || isNaN(sB) || isNaN(eB)) return false;
  return sA <= eB && sB <= eA;
}

// Union-Find used by Flag 1 to identify connected overlapping-period groups
function makeUF(n) {
  const p = Array.from({ length: n }, (_, i) => i);
  function find(i) { return p[i] === i ? i : (p[i] = find(p[i])); }
  function union(i, j) { p[find(i)] = find(j); }
  return { find, union };
}

function computeFlaggedItems(output, prevBilledRows) {
  const flagged = [];

  // ── Flag 1: Duplicate MRC Charge, Same/Overlapping Period ─────────────────
  const byMDN = new Map();
  for (let i = 0; i < output.length; i++) {
    const mdn = String(output[i]['Number'] || '').trim();
    if (!mdn) continue;
    if (!byMDN.has(mdn)) byMDN.set(mdn, []);
    byMDN.get(mdn).push(i);
  }

  for (const [mdn, indices] of byMDN) {
    if (indices.length < 2) continue;
    const n = indices.length;
    const { find, union } = makeUF(n);

    for (let i = 0; i < n; i++) {
      for (let j = i + 1; j < n; j++) {
        const a = output[indices[i]], b = output[indices[j]];
        if (periodsOverlap(a['Date Start'], a['Date End'], b['Date Start'], b['Date End'])) {
          union(i, j);
        }
      }
    }

    // Collect components with ≥ 2 members
    const components = new Map();
    for (let i = 0; i < n; i++) {
      const root = find(i);
      if (!components.has(root)) components.set(root, []);
      components.get(root).push(indices[i]);
    }

    for (const [, group] of components) {
      if (group.length < 2) continue;
      const rows = group.map(idx => output[idx]);
      // Suppress if any row in the group is a Credit or Usage
      if (rows.some(r => isUsageRow(r) || isCreditRow(r))) continue;

      for (const r of rows) {
        flagged.push({
          'MDN':              mdn,
          'Flag Type':        'Duplicate-SamePeriod',
          'Period Start':     r['Date Start'],
          'Period End':       r['Date End'],
          'Prior Period Start': '',
          'Prior Period End':   '',
          'Description':      r['Description'],
          'Amount':           parseFloat(r['Amount']) || 0,
          'Prior Amount':     '',
          'Row Count in Group': group.length,
          'Flag Reason':      'Duplicate MRC charge — same/overlapping period, no credit/usage offset',
          'Status':           '',
        });
      }
    }
  }

  // ── Flag 2: High Billable Usage (> 10 GB) ─────────────────────────────────
  for (const row of output) {
    if (!isUsageRow(row)) continue;
    const gb = parseGBUsage(row['Usage GB']);
    if (gb <= 10) continue;
    flagged.push({
      'MDN':              String(row['Number'] || '').trim(),
      'Flag Type':        'HighUsage',
      'Period Start':     row['Date Start'],
      'Period End':       row['Date End'],
      'Prior Period Start': '',
      'Prior Period End':   '',
      'Description':      row['Description'],
      'Amount':           parseFloat(row['Amount']) || 0,
      'Prior Amount':     '',
      'Row Count in Group': '',
      'Flag Reason':      `High billable usage — ${gb.toFixed(4)} GB exceeds 10 GB threshold`,
      'Status':           '',
    });
  }

  // ── Flag 3: Already Billed in Prior Month ─────────────────────────────────
  if (prevBilledRows && prevBilledRows.length > 0) {
    const priorByMDN = new Map();
    for (const pr of prevBilledRows) {
      if (!pr.mdn) continue;
      if (!priorByMDN.has(pr.mdn)) priorByMDN.set(pr.mdn, []);
      priorByMDN.get(pr.mdn).push(pr);
    }

    for (const row of output) {
      const mdn = String(row['Number'] || '').trim();
      if (!mdn) continue;
      if (isUsageRow(row) || isCreditRow(row)) continue; // current row suppresses itself

      const priorRows = priorByMDN.get(mdn);
      if (!priorRows) continue;

      for (const pr of priorRows) {
        if (pr.isCredit || pr.isUsage) continue; // prior row suppresses itself
        if (!periodsOverlap(row['Date Start'], row['Date End'], pr.dateStart, pr.dateEnd)) continue;

        flagged.push({
          'MDN':              mdn,
          'Flag Type':        'Duplicate-PriorMonth',
          'Period Start':     row['Date Start'],
          'Period End':       row['Date End'],
          'Prior Period Start': pr.dateStart,
          'Prior Period End':   pr.dateEnd,
          'Description':      row['Description'],
          'Amount':           parseFloat(row['Amount']) || 0,
          'Prior Amount':     pr.amount,
          'Row Count in Group': '',
          'Flag Reason':      'MDN already billed in overlapping period in prior month\'s working copy',
          'Status':           '',
        });
        break; // one flag per current row is sufficient
      }
    }
  }

  // Sort by MDN, then Flag Type
  flagged.sort((a, b) =>
    a['MDN'] !== b['MDN']
      ? String(a['MDN']).localeCompare(String(b['MDN']))
      : a['Flag Type'].localeCompare(b['Flag Type'])
  );

  return flagged;
}

// ── End flag helpers ──────────────────────────────────────────────────────────

function processFiles(rawFilePath, prevWorkingPath, nalMasterlistPath) {
  const rawWb = XLSX.readFile(rawFilePath);
  const rawData = XLSX.utils.sheet_to_json(rawWb.Sheets[rawWb.SheetNames[0]], { defval: '' });

  // --- Previous Working Copy lookup ---
  const prevLookup = new Map();
  const prevBilledRows = []; // for Flag 3
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
      if (!mdn) continue;
      if (!prevLookup.has(mdn)) {
        prevLookup.set(mdn, {
          clec:     getCol(row, 'clec name', 'clec'),
          soc:      getCol(row, 'soc-code nrt plan if na consult support', 'soc code', 'soc'),
          planRate: getCol(row, 'plan rate customer rate', 'plan rate'),
        });
      }
      // Collect all rows (including duplicates) for Flag 3 overlap check
      const amt = parseFloat(row['Amount']) || 0;
      const desc = String(row['Description'] || '').trim();
      const billingCat = String(row['Billing Category'] || '').trim();
      prevBilledRows.push({
        mdn,
        dateStart: String(row['Date Start'] || '').trim(),
        dateEnd:   String(row['Date End']   || '').trim(),
        amount:    amt,
        isCredit:  amt < 0 || /^Credit/i.test(desc),
        isUsage:   billingCat === 'COG',
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

  return { output, summary, unidentifiedList: [...unidentifiedMap.values()], billingSummary, prevWorkingWarning, prevBilledRows };
}

function buildWorkbook(output, summary, unidentifiedList, billingSummary, prevBilledRows) {
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

  // --- Flagged Items sheet ---
  {
    const flaggedItems = computeFlaggedItems(output, prevBilledRows || []);

    const FLAG_COLS = [
      'MDN', 'Flag Type', 'Period Start', 'Period End',
      'Prior Period Start', 'Prior Period End',
      'Description', 'Amount', 'Prior Amount', 'Row Count in Group',
      'Flag Reason', 'Status',
    ];

    // aoa_to_sheet guarantees every cell exists — no missing-cell style-drop issues
    const aoa = [FLAG_COLS];
    if (flaggedItems.length > 0) {
      for (const item of flaggedItems) {
        aoa.push(FLAG_COLS.map(k => item[k] === undefined || item[k] === '' ? '' : item[k]));
      }
    } else {
      aoa.push(['', '', '', '', '', '', '', '', '', '', 'No flagged items found', '']);
    }

    const wsF = XLSX.utils.aoa_to_sheet(aoa);

    const fHeaderStyle = {
      font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Arial', sz: 10 },
      fill: { fgColor: { rgb: '7B2C2C' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    };
    // Row-background colours keyed by Flag Type value (column index 1)
    const FLAG_COLOURS = {
      'Duplicate-SamePeriod': 'FFF2CC', // yellow
      'HighUsage':            'FCE4D6', // orange
      'Duplicate-PriorMonth': 'DDEBF7', // blue
    };
    const numCols = FLAG_COLS.length;
    const numRows = aoa.length;

    for (let C = 0; C < numCols; C++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c: C });
      if (!wsF[addr]) wsF[addr] = { v: FLAG_COLS[C], t: 's' };
      wsF[addr].s = fHeaderStyle;
    }

    for (let R = 1; R < numRows; R++) {
      const flagTypeCell = wsF[XLSX.utils.encode_cell({ r: R, c: 1 })];
      const flagType = flagTypeCell ? String(flagTypeCell.v || '') : '';
      const bgRgb = FLAG_COLOURS[flagType] || null;
      const rowStyle = bgRgb
        ? { fill: { fgColor: { rgb: bgRgb } }, font: { name: 'Arial', sz: 10 } }
        : { font: { name: 'Arial', sz: 10 } };

      for (let C = 0; C < numCols; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!wsF[addr]) wsF[addr] = { v: '', t: 's' };
        wsF[addr].s = rowStyle;
      }
    }

    wsF['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: numRows - 1, c: numCols - 1 } });
    wsF['!cols'] = [
      { wch: 14 }, // MDN
      { wch: 22 }, // Flag Type
      { wch: 13 }, // Period Start
      { wch: 13 }, // Period End
      { wch: 18 }, // Prior Period Start
      { wch: 18 }, // Prior Period End
      { wch: 40 }, // Description
      { wch: 12 }, // Amount
      { wch: 12 }, // Prior Amount
      { wch: 8  }, // Row Count in Group
      { wch: 62 }, // Flag Reason
      { wch: 18 }, // Status
    ];
    XLSX.utils.book_append_sheet(wb, wsF, 'Flagged Items');
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
    const { output, summary, unidentifiedList, billingSummary, prevBilledRows } =
      processFiles(rawFile.path, prevWorking?.path, nalFile?.path);
    const wbOut = buildWorkbook(output, summary, unidentifiedList, billingSummary, prevBilledRows);
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

// ============================================================
// TAB 2 — NRT 499Q/A GL REVENUE REPORT PROCESSOR
// ============================================================

function normalizeGLCustomer(name) {
  const n = String(name || '').toLowerCase();
  if (n.includes('north american local')) return 'NAL';
  if (n.includes('telzeq')) return 'Telzeq';
  if (n.includes('metro communications')) return 'Metro';
  return String(name || '').trim();
}

function getOveragePlan(desc) {
  const d = String(desc || '').toLowerCase();
  if (d.includes('hcd') || d.includes('home connect')) return 'Home Connect Device Plan';
  if (/10\s*gb/.test(d)) return '10GB Voice and Data';
  if (/5\s*gb/.test(d)) return '5GB Voice and Data';
  return '1GB Voice and Data';
}

function parseGLFile(glFilePath) {
  const wb = XLSX.readFile(glFilePath);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });

  const period = String(rows[2]?.[0] || '').trim();

  // att[planKey][customer] = amount  (overage already distributed into plan)
  const att = {};
  // nonAtt[categoryLabel] = { [customer]: amount }
  const nonAtt = {};

  const MAIN_WITH_SUBS = new Set(['AT&T', 'Billable Contracts Hour', 'T -MOBILE']);
  const SKIP_PREFIXES = ['National Relief', 'Transaction Report', 'Cash Basis', 'TOTAL', 'Total for'];
  const DATE_RE = /^\d{2}\/\d{2}\/\d{4}$/;

  let mainCat = null;
  let subCat = null;
  let inSub = false;

  for (const row of rows) {
    const c0 = String(row[0] || '').trim();
    const c1 = String(row[1] || '').trim();
    const amt = typeof row[8] === 'number' ? row[8] : 0;
    const name = String(row[4] || '').trim();
    const desc = String(row[5] || '').trim();

    if (!c0 && !c1) continue;

    if (c0.startsWith('Total for') && c0.includes('with sub-accounts')) {
      mainCat = null; subCat = null; inSub = false;
      continue;
    }
    if (c0.startsWith('Total for') || c0 === 'TOTAL') { subCat = null; continue; }
    if (SKIP_PREFIXES.some(p => c0.startsWith(p))) continue;
    if (c0 === '') {
      // Possible transaction row
      if (DATE_RE.test(c1) && amt !== 0) {
        // Fall back to Num field (col 3) when Name (col 4) is empty — handles bad-debt Journal Entries
        const nameOrNum = name || String(row[3] || '').trim();
        const cust = normalizeGLCustomer(nameOrNum);
        if (mainCat === 'AT&T' && subCat) {
          let planKey = subCat === 'AT&T Overage' ? getOveragePlan(desc) : subCat;
          if (!att[planKey]) att[planKey] = {};
          att[planKey][cust] = (att[planKey][cust] || 0) + amt;
        } else if (mainCat && mainCat !== 'AT&T') {
          const key = subCat ? `${mainCat} — ${subCat}` : mainCat;
          if (!nonAtt[key]) nonAtt[key] = {};
          nonAtt[key][cust] = (nonAtt[key][cust] || 0) + amt;
        }
      }
      continue;
    }

    // Category header (c0 non-empty, not total/skip)
    if (MAIN_WITH_SUBS.has(c0)) {
      mainCat = c0; subCat = null; inSub = true;
    } else if (inSub) {
      subCat = c0;
    } else {
      mainCat = c0; subCat = null; inSub = false;
    }
  }

  return { period, att, nonAtt };
}

function sumCustomers(obj) {
  return Object.values(obj || {}).reduce((s, v) => s + v, 0);
}
function round2(n) { return Math.round((n || 0) * 100) / 100; }

function buildGLWorkbook(period, att, nonAtt) {
  const wb = XLSX.utils.book_new();

  // ── Dynamically discover customers from GL data ──
  const PREFERRED_CUSTOMERS = ['NAL', 'Telzeq', 'Metro'];
  const STATE_MAP = { NAL: 'Florida', Telzeq: 'New York', Metro: 'Georgia' };
  const CUSTOMER_DISPLAY = { NAL: 'North American Local' };
  const stateLabel = (c) => STATE_MAP[c] || c;
  const custDisplay = (c) => CUSTOMER_DISPLAY[c] || c;

  const customerSet = new Set();
  for (const planData of Object.values(att)) {
    for (const c of Object.keys(planData)) customerSet.add(c);
  }
  const customers = [
    ...PREFERRED_CUSTOMERS.filter(c => customerSet.has(c)),
    ...[...customerSet].filter(c => !PREFERRED_CUSTOMERS.includes(c)).sort(),
  ];

  // ── Dynamically discover AT&T plans from GL data ──
  const PREFERRED_PLANS = [
    '10GB Voice and Data', '1GB Voice and Data', '5GB Voice and Data',
    'Home Connect Device Plan', '50MB Small Data Plan', 'Feature Phone',
    'AT&T MRC Credit Sale', 'AT&T Passthrough Fees',
  ];
  const PLAN_LABEL = { 'AT&T Passthrough Fees': 'AT&T Passthrough' };
  const leftPlans = [
    ...PREFERRED_PLANS.filter(p => att[p]),
    ...Object.keys(att).filter(p => !PREFERRED_PLANS.includes(p)).sort(),
  ];

  // Plan type classification — right-side table columns
  // Any plan not in this map is treated as Voice and Data automatically
  const PLAN_TYPE_MAP = {
    'Home Connect Device Plan': 'hcd',
    'Feature Phone': 'fp',
    'AT&T Passthrough Fees': 'pt',
    'AT&T MRC Credit Sale': 'mrc',
  };
  const planType = (p) => PLAN_TYPE_MAP[p] || 'vd';

  const planAmt = (plan, cust) => round2((att[plan] || {})[cust] || 0);

  // ── Per-customer aggregates ──
  const stateAgg = {};
  for (const c of customers) {
    const agg = { vd: 0, fp: 0, hcd: 0, pt: 0, mrc: 0 };
    for (const plan of leftPlans) {
      const t = planType(plan);
      agg[t] = round2(agg[t] + planAmt(plan, c));
    }
    agg.total = round2(agg.vd + agg.fp + agg.hcd + agg.pt + agg.mrc);
    stateAgg[c] = agg;
  }

  // Zero-out customers whose net total rounds to 0
  const zeroedCustomers = new Set(customers.filter(c => Math.abs(stateAgg[c].total) < 0.01));
  for (const c of zeroedCustomers) {
    stateAgg[c] = { vd: 0, fp: 0, hcd: 0, pt: 0, mrc: 0, total: 0 };
  }

  const attGrandTotal = round2(customers.reduce((s, c) => s + stateAgg[c].total, 0));

  // ── Column layout (dynamic based on number of customers) ──
  const N = customers.length;
  const colTotal  = N + 1;
  const colState  = N + 3;
  const colVD     = N + 4;
  const colFP     = N + 5;
  const colHCD    = N + 6;
  const colPT     = N + 7;
  const colMRC    = N + 8;
  const colRTotal = N + 9;

  // ── Cell styles ──
  const blueHeader = { font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: '1F4E79' } }, alignment: { horizontal: 'center' } };
  const boldBlue   = { font: { bold: true, color: { rgb: '1F4E79' }, name: 'Arial', sz: 10 } };
  const boldBlack  = { font: { bold: true, name: 'Arial', sz: 10 } };
  const normal     = { font: { name: 'Arial', sz: 10 } };
  const numFmt     = '$#,##0.00;($#,##0.00);"-"';

  function setCell(ws, r, c, v, s, z) {
    const addr = XLSX.utils.encode_cell({ r, c });
    ws[addr] = { v, t: typeof v === 'number' ? 'n' : 's' };
    if (s) ws[addr].s = s;
    if (z) ws[addr].z = z;
  }

  // ── AT&T Sheet ──
  const attWS = XLSX.utils.aoa_to_sheet([]);

  setCell(attWS, 0, 0, 'National Relief Telecom', boldBlue);
  setCell(attWS, 1, 0, period, normal);

  // Row 3: headers
  setCell(attWS, 3, 0, '', blueHeader);
  customers.forEach((c, i) => setCell(attWS, 3, 1 + i, custDisplay(c), blueHeader));
  setCell(attWS, 3, colTotal,  'Total',               blueHeader);
  setCell(attWS, 3, colState,  '',                    blueHeader);
  setCell(attWS, 3, colVD,     'Voice and Data',      blueHeader);
  setCell(attWS, 3, colFP,     'Feature Phone',       blueHeader);
  setCell(attWS, 3, colHCD,    'HCD Plan',            blueHeader);
  setCell(attWS, 3, colPT,     'AT&T Passthrough',    blueHeader);
  setCell(attWS, 3, colMRC,    'AT&T MRC Credit Sale',blueHeader);
  setCell(attWS, 3, colRTotal, 'Total',               blueHeader);

  // Row 4: state sub-labels for left table columns + first customer's right-table row
  const subLabelStyle = { font: { bold: true, name: 'Arial', sz: 10 }, alignment: { horizontal: 'center' } };
  customers.forEach((c, i) => setCell(attWS, 4, 1 + i, stateLabel(c), subLabelStyle));

  function writeRightRow(ws, r, c) {
    const agg = stateAgg[c];
    setCell(ws, r, colState, stateLabel(c), boldBlack);
    if (agg.vd)  setCell(ws, r, colVD,    agg.vd,    normal, numFmt);
    if (agg.fp)  setCell(ws, r, colFP,    agg.fp,    normal, numFmt);
    if (agg.hcd) setCell(ws, r, colHCD,   agg.hcd,   normal, numFmt);
    if (agg.pt)  setCell(ws, r, colPT,    agg.pt,    normal, numFmt);
    if (agg.mrc) setCell(ws, r, colMRC,   agg.mrc,   normal, numFmt);
    setCell(ws, r, colRTotal, agg.total, boldBlack, numFmt);
  }

  if (customers[0]) writeRightRow(attWS, 4, customers[0]);

  // Left table plan rows + right-side state rows
  leftPlans.forEach((plan, i) => {
    const r = 5 + i;
    setCell(attWS, r, 0, PLAN_LABEL[plan] || plan, normal);

    let rowTotal = 0;
    customers.forEach((c, ci) => {
      const zeroed = zeroedCustomers.has(c);
      const amt = zeroed ? 0 : planAmt(plan, c);
      // Show 0 explicitly for zeroed customers; skip cell for non-zeroed zeros
      if (zeroed || amt) setCell(attWS, r, 1 + ci, amt, normal, numFmt);
      if (!zeroed) rowTotal = round2(rowTotal + planAmt(plan, c));
    });
    if (rowTotal) setCell(attWS, r, colTotal, rowTotal, boldBlack, numFmt);

    // Right-side: customers[1] goes on row 5 (i=0), customers[2] on row 6 (i=1), etc.
    const stateIdx = i + 1;
    if (stateIdx < customers.length) {
      writeRightRow(attWS, r, customers[stateIdx]);
    } else if (stateIdx === customers.length) {
      // Place right-table grand total on the row right after the last state
      setCell(attWS, r, colMRC,    'Total',        boldBlack);
      setCell(attWS, r, colRTotal, attGrandTotal,  boldBlack, numFmt);
    }
  });

  // Left table grand total row
  const totalRow = 5 + leftPlans.length;
  customers.forEach((c, ci) => {
    const zeroed = zeroedCustomers.has(c);
    const custTotal = zeroed ? 0 : round2(leftPlans.reduce((s, p) => s + planAmt(p, c), 0));
    setCell(attWS, totalRow, 1 + ci, custTotal, boldBlack, numFmt);
  });
  setCell(attWS, totalRow, colTotal, attGrandTotal, boldBlack, numFmt);

  const attCols = [{ wch: 28 }];
  for (let i = 0; i < N; i++) attCols.push({ wch: i === 0 ? 16 : 10 });
  attCols.push({ wch: 14 }, { wch: 3 }, { wch: 12 }, { wch: 16 }, { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 18 }, { wch: 14 });
  attWS['!cols'] = attCols;
  attWS['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: totalRow, c: colRTotal } });
  XLSX.utils.book_append_sheet(wb, attWS, 'AT&T');

  // ── Summary Sheet ──
  const sumWS = XLSX.utils.aoa_to_sheet([]);
  const sumHeaders = ['State','Voice and Data','Feature Phone','HCD Plan','Pass-Through','Telecom','Non-Telecom','Total'];
  sumHeaders.forEach((h, c) => setCell(sumWS, 0, c, h, blueHeader));

  let sumRow = 1;
  let telecomTotal = 0;

  for (const c of customers) {
    if (!Object.values(att).some(v => c in v)) continue;
    const agg = stateAgg[c];
    // V&D in Summary = VD + MRC credit (netted in), per existing business logic
    const vdNet = round2(agg.vd + agg.mrc);
    const telecom = round2(vdNet + agg.fp + agg.hcd + agg.pt);
    setCell(sumWS, sumRow, 0, stateLabel(c), boldBlack);
    if (vdNet)   setCell(sumWS, sumRow, 1, vdNet,   normal, numFmt);
    if (agg.fp)  setCell(sumWS, sumRow, 2, agg.fp,  normal, numFmt);
    if (agg.hcd) setCell(sumWS, sumRow, 3, agg.hcd, normal, numFmt);
    if (agg.pt)  setCell(sumWS, sumRow, 4, agg.pt,  normal, numFmt);
    setCell(sumWS, sumRow, 5, telecom, normal, numFmt);
    telecomTotal = round2(telecomTotal + telecom);
    sumRow++;
  }

  // ── Dynamic non-telecom items ──
  // Group sub-categories (e.g. "Billable Contracts Hour — Admin") by parent category
  const DISPLAY_LABEL_MAP = { 'Billable Contracts Hour': 'Billable Contract Hours' };
  const nonAttGrouped = {};
  for (const [key, custData] of Object.entries(nonAtt)) {
    const parent = key.split(' — ')[0].trim();
    const label = DISPLAY_LABEL_MAP[parent] || parent;
    nonAttGrouped[label] = round2((nonAttGrouped[label] || 0) + sumCustomers(custData));
  }

  // Preferred order for known items; unknown new items appended alphabetically
  const PREFERRED_NON_TELECOM = [
    'AT&T Sim Cards', 'Billable Contract Hours', 'Bookkeeping Revenue',
    'Exclusive Brand – One-Time Development Fee', 'Lifeline Revenue',
    'Marketing Revenue', 'Monthly Network Storage Fee',
  ];
  const nonTelecomKeys = [
    ...PREFERRED_NON_TELECOM.filter(k => nonAttGrouped[k] && Math.abs(nonAttGrouped[k]) > 0.005),
    ...Object.keys(nonAttGrouped).filter(k => !PREFERRED_NON_TELECOM.includes(k) && Math.abs(nonAttGrouped[k]) > 0.005).sort(),
  ];

  let nonTelecomTotal = 0;
  for (const label of nonTelecomKeys) {
    const amt = round2(nonAttGrouped[label]);
    setCell(sumWS, sumRow, 0, label,  boldBlack);
    setCell(sumWS, sumRow, 6, amt,    normal, numFmt);
    nonTelecomTotal = round2(nonTelecomTotal + amt);
    sumRow++;
  }

  const grandTotal = round2(telecomTotal + nonTelecomTotal);
  setCell(sumWS, sumRow, 0, 'Total',          boldBlack);
  setCell(sumWS, sumRow, 5, telecomTotal,     boldBlack, numFmt);
  setCell(sumWS, sumRow, 6, nonTelecomTotal,  boldBlack, numFmt);
  setCell(sumWS, sumRow, 7, grandTotal,       boldBlack, numFmt);

  sumWS['!cols'] = [{ wch: 44 }, { wch: 16 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
  sumWS['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: sumRow, c: 7 } });
  XLSX.utils.book_append_sheet(wb, sumWS, 'Summary');

  return wb;
}

const glUpload = upload.fields([{ name: 'glFile', maxCount: 1 }]);

app.post('/api/499qa/process', glUpload, (req, res) => {
  const glFile = req.files?.['glFile']?.[0];
  if (!glFile) return res.status(400).json({ error: 'GL file is required.' });
  try {
    const { period, att, nonAtt } = parseGLFile(glFile.path);
    const wb = buildGLWorkbook(period, att, nonAtt);
    const outPath = path.join('uploads', `nrt_499qa_${Date.now()}.xlsx`);
    XLSX.writeFile(wb, outPath, { bookType: 'xlsx', type: 'binary', cellStyles: true });
    cleanupFiles(glFile);
    const label = `NRT_Revenue_Report_${period.replace(/[^a-zA-Z0-9]/g,'_')}.xlsx`;
    res.download(outPath, label, () => { if (fs.existsSync(outPath)) fs.unlinkSync(outPath); });
  } catch (err) {
    console.error(err);
    cleanupFiles(glFile);
    res.status(500).json({ error: err.message });
  }
});

if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

// ============================================================
// TAB 3 — 321 Comm TFN Data Processing
// ============================================================

let pdfParse = null;
try { pdfParse = require('pdf-parse'); } catch(e) { /* graceful fallback */ }

const archiver = require('archiver');

const TFN_FEES = {
  tfnNumberFee: 1.75,  // per TFN number/month
  didNumberFee: 2.00,  // per DID number/month
};

const MARKUP = 1.15;

// r4 for per-minute rates (4 decimal places for precision)
function r4(n) { return Math.round((n || 0) * 10000) / 10000; }

// Compute dynamic per-minute rates from vendor invoices + total minutes
// invoices: { amt382, amtMCI, amtLumenInter, amtLumenIntra, amtIPC }
// totMins:  { tfnIn, didIn, ldInter, ldIntra, ipcOut }
function computeRates(invoices, totMins) {
  const safe = (amt, mins) => (amt && mins) ? r4((amt / mins) * MARKUP) : 0;
  return {
    inbound:      safe(invoices.amt382,        totMins.tfnIn),
    didTraffic:   safe(invoices.amtMCI,        totMins.didIn),
    ldInterstate: safe(invoices.amtLumenInter, totMins.ldInter),
    ldIntrastate: safe(invoices.amtLumenIntra, totMins.ldIntra),
    ipcOutbound:  safe(invoices.amtIPC,        totMins.ipcOut),
  };
}

// Extract total minutes across all customers from parsed Gopal summary + IPC file
function computeTotalMins(gopalSummary, ipcOutbound) {
  const { tfnInbound, didInbound, lumenLD } = gopalSummary;
  const sum = obj => Object.values(obj).reduce((s, v) => s + (v || 0), 0);
  const ldInter = Object.values(lumenLD).reduce((s, v) => s + (v['Long Distance Interstate'] || 0), 0);
  const ldIntra = Object.values(lumenLD).reduce((s, v) => s + (v['Long Distance Intrastate'] || 0), 0);
  return {
    tfnIn:  r2(sum(tfnInbound)),
    didIn:  r2(sum(didInbound)),
    ldInter: r2(ldInter),
    ldIntra: r2(ldIntra),
    ipcOut: r2(sum(ipcOutbound)),
  };
}

function r2(n) { return Math.round((n || 0) * 100) / 100; }

// ── Parse TFN Inventory for number counts per customer ────────
function parseTFNInventory(filePath) {
  const wb    = XLSX.readFile(filePath);
  const rows  = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
  const counts = {};
  for (const row of rows) {
    const cust = String(row['Customer Name'] || '').trim();
    if (!cust) continue;
    counts[cust] = (counts[cust] || 0) + 1;
  }
  return counts; // { 'Paricus': 25, 'Centercom': 81, 'NAL': 7, ... }
}

// ── Parse DID Inventory for DID number counts per customer ────
function parseDIDInventory(filePath) {
  const wb   = XLSX.readFile(filePath);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
  const headers = rows[0];
  const custIdx = headers.indexOf('Customers assigned');
  if (custIdx < 0) return {};
  const counts = {};
  for (const row of rows.slice(1)) {
    const cust = String(row[custIdx] || '').trim();
    if (!cust) continue;
    counts[cust] = (counts[cust] || 0) + 1;
  }
  return counts; // { 'Paricus': 50, 'Centercom': 87, ... }
}

// ── Parse IPC Outbound file for outbound minutes per customer ─
function parseIPCOutboundFile(filePath) {
  const wb      = XLSX.readFile(filePath);
  const summWS  = wb.Sheets['Summary'];
  const minutes = {}; // customerName → total outbound minutes
  if (!summWS) return minutes;
  const rows = XLSX.utils.sheet_to_json(summWS, { header: 1, defval: '' });
  let inSection = false;
  for (const row of rows) {
    const c0 = String(row[0] || '').trim();
    if (c0.includes('IPC') || c0.includes('Termination')) { inSection = true; continue; }
    if (!inSection) continue;
    if (c0 === 'OriginationcarrierName' || !c0) continue;
    if (typeof row[2] === 'number') minutes[c0] = (minutes[c0] || 0) + row[2];
  }
  return minutes;
}

// ── Parse Gopal new-version CDR file ─────────────────────────
function parseGopalFile(filePath) {
  const wb   = XLSX.readFile(filePath);
  const data = { summary: { tfnInbound: {}, didInbound: {}, lumenLD: {} }, cdrs: {} };

  const summaryWS = wb.Sheets['Summary'];
  if (summaryWS) {
    const rows    = XLSX.utils.sheet_to_json(summaryWS, { header: 1, defval: '' });
    let section   = null;
    for (const row of rows) {
      const c0 = String(row[0] || '').trim();
      const c1 = String(row[1] || '').trim();
      if (c0.includes('382'))                { section = 'tfn'; continue; }
      if (c0.includes('DID--'))              { section = 'did'; continue; }
      if (c0.includes('Lumen') || c0.includes('Termination Calls')) { section = 'lumen'; continue; }
      if (!c0 || c0 === 'CustomerName' || c0 === 'OriginationcarrierName') continue;
      if (section === 'tfn'   && typeof row[1] === 'number') data.summary.tfnInbound[c0] = row[1];
      if (section === 'did'   && typeof row[1] === 'number') data.summary.didInbound[c0] = row[1];
      if (section === 'lumen' && typeof row[2] === 'number') {
        if (!data.summary.lumenLD[c0]) data.summary.lumenLD[c0] = {};
        data.summary.lumenLD[c0][c1] = (data.summary.lumenLD[c0][c1] || 0) + row[2];
      }
    }
  }

  for (const sn of wb.SheetNames) {
    if (sn === 'Summary') continue;
    data.cdrs[sn] = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '' });
  }
  return data;
}

// ── Slice column group from raw rows (array-of-arrays) ───────
function sliceCols(rows, colMap) {
  // colMap: { fieldName: colIndex, ... }
  const firstIdx = Object.values(colMap)[0];
  return rows.slice(1)
    .filter(r => r[firstIdx] !== '' && r[firstIdx] !== null && r[firstIdx] !== undefined)
    .map(r => {
      const obj = {};
      for (const [field, idx] of Object.entries(colMap)) obj[field] = r[idx];
      return obj;
    });
}

// ── Compute per-customer data from parsed Gopal ───────────────
function computeTFNCustomers(gopal, inventoryCounts, ipcOutbound, didInventoryCounts, rates) {
  const { summary, cdrs } = gopal;
  inventoryCounts    = inventoryCounts    || {};
  ipcOutbound        = ipcOutbound        || {};
  didInventoryCounts = didInventoryCounts || {};
  rates              = rates              || {};

  const numCount = (invKey, cdrNums) => inventoryCounts[invKey]    || cdrNums.length;
  const didCount = (invKey, cdrNums) => didInventoryCounts[invKey] || cdrNums.length;

  const outMins = (ipcKey, cdrRows) =>
    ipcOutbound[ipcKey] != null
      ? r2(ipcOutbound[ipcKey])
      : r2(cdrRows.reduce((s, r) => s + (parseFloat(r.Duration_Minutes) || 0), 0));

  // ── NAL ──
  const nalSheet   = cdrs['NAL CDR-In and Out'] || [];
  const nalIn      = sliceCols(nalSheet, { SourceNumber: 0, TerminationNumber: 1, calldatetime: 2, Callduration_Minutes: 3 });
  const nalOutCDR  = sliceCols(nalSheet, { OriginationcarrierName: 11, sourcenumber: 12, terminationnumber: 13, starttime: 14, Duration_Minutes: 15 });
  const nalTFNNums = [...new Set(nalIn.map(r => r.TerminationNumber).filter(Boolean))];

  const nalTFNMin   = r2(summary.tfnInbound['NorthAmericanLocal'] || 0);
  const nalLDMins   = summary.lumenLD['NorthAmericanLocal'] || {};
  const nalLDInt    = r2(nalLDMins['Long Distance Interstate'] || 0);
  const nalOutMin   = outMins('NorthAmericanLocal', nalOutCDR);
  const nalTFNCount = numCount('NorthAmericanLocal', nalTFNNums);

  const nalAmtInbound  = r2(nalTFNMin   * (rates.inbound      || 0));
  const nalAmtTFNFee   = r2(nalTFNCount * TFN_FEES.tfnNumberFee);
  const nalAmtLDInt    = r2(nalLDInt    * (rates.ldInterstate || 0));
  const nalAmtOutbound = r2(nalOutMin   * (rates.ipcOutbound  || 0));
  const nalTotal       = r2(nalAmtInbound + nalAmtTFNFee + nalAmtLDInt + nalAmtOutbound);

  // ── Paricus ──
  const parInSheet  = cdrs['Paricus - Inbound']  || [];
  const parOutSheet = cdrs['Paricus - Dial Out']  || [];
  const parIn      = sliceCols(parInSheet,  { SourceNumber: 0, TerminationNumber: 1, calldatetime: 2, Callduration_Minutes: 3 });
  const parDID     = sliceCols(parInSheet,  { CustomerName: 10, DID: 11, SourceNumber: 12, CallConnectTime: 13, Duration_Minutes: 14 });
  const parOutCDR  = sliceCols(parOutSheet, { OriginationcarrierName: 0, sourcenumber: 1, calledno: 2, starttime: 3, Duration_Minutes: 4 });
  const parTFNNums = [...new Set(parIn.map(r => r.TerminationNumber).filter(Boolean))];
  const parDIDNums = [...new Set(parDID.map(r => r.DID).filter(Boolean))];

  const parTFNMin   = r2(summary.tfnInbound['Paricus'] || 0);
  const parDIDMin   = r2(summary.didInbound['Paricus']  || 0);
  const parLDMins   = summary.lumenLD['Paricus'] || {};
  const parLDInt    = r2(parLDMins['Long Distance Interstate'] || 0);
  const parOutMin   = outMins('Paricus', parOutCDR);
  const parTFNCount = numCount('Paricus', parTFNNums);
  const parDIDCount = didCount('Paricus', parDIDNums);

  const parAmtInbound  = r2(parTFNMin   * (rates.inbound      || 0));
  const parAmtTFNFee   = r2(parTFNCount * TFN_FEES.tfnNumberFee);
  const parAmtLDInt    = r2(parLDInt    * (rates.ldInterstate || 0));
  const parAmtOutbound = r2(parOutMin   * (rates.ipcOutbound  || 0));
  const parAmtDIDMin   = r2(parDIDMin   * (rates.didTraffic   || 0));
  const parAmtDIDFee   = r2(parDIDCount * TFN_FEES.didNumberFee);
  const parTotal       = r2(parAmtInbound + parAmtTFNFee + parAmtLDInt + parAmtOutbound + parAmtDIDMin + parAmtDIDFee);

  // ── Torch (Centercom) ──
  const torchSheet    = cdrs['Centercom CDR--In and OUt'] || [];
  const torchIn       = sliceCols(torchSheet, { SourceNumber: 0, TerminationNumber: 1, calldatetime: 2, Callduration_Minutes: 3 });
  const torchOutCDR   = sliceCols(torchSheet, { OriginationcarrierName: 8, sourcenumber: 9, terminationnumber: 10, starttime: 11, Duration_Minutes: 12 });
  const torchDID      = sliceCols(torchSheet, { CustomerName: 18, DID: 19, SourceNumber: 20, CallConnectTime: 21, Duration_Minutes: 22 });
  const torchTFNNums  = [...new Set(torchIn.map(r => r.TerminationNumber).filter(Boolean))];
  const torchDIDNums  = [...new Set(torchDID.map(r => r.DID).filter(Boolean))];

  const torchTFNMin   = r2(summary.tfnInbound['Centercom'] || 0);
  const torchDIDMin   = r2(summary.didInbound['Centercom']  || 0);
  const torchLDMins   = summary.lumenLD['Centercom'] || {};
  const torchLDInt    = r2(torchLDMins['Long Distance Interstate']  || 0);
  const torchLDIntra  = r2(torchLDMins['Long Distance Intrastate']  || 0);
  const torchOutMin   = outMins('Centercom', torchOutCDR);
  const torchTFNCount = numCount('Centercom', torchTFNNums);
  const torchDIDCount = didCount('Centercom', torchDIDNums);

  const torchAmtInbound  = r2(torchTFNMin   * (rates.inbound      || 0));
  const torchAmtTFNFee   = r2(torchTFNCount * TFN_FEES.tfnNumberFee);
  const torchAmtLDInt    = r2(torchLDInt    * (rates.ldInterstate || 0));
  const torchAmtLDIntra  = r2(torchLDIntra  * (rates.ldIntrastate || 0));
  const torchAmtOutbound = r2(torchOutMin   * (rates.ipcOutbound  || 0));
  const torchAmtDIDMin   = r2(torchDIDMin   * (rates.didTraffic   || 0));
  const torchAmtDIDFee   = r2(torchDIDCount * TFN_FEES.didNumberFee);
  const torchTotal       = r2(torchAmtInbound + torchAmtTFNFee + torchAmtLDInt + torchAmtLDIntra + torchAmtOutbound + torchAmtDIDMin + torchAmtDIDFee);

  return {
    nal: {
      inRows: nalIn, outRows: nalOutCDR, tfnNumbers: nalTFNNums,
      tfnCount: nalTFNCount, tfnMin: nalTFNMin, ldInt: nalLDInt, outMin: nalOutMin,
      amtInbound: nalAmtInbound, amtTFNFee: nalAmtTFNFee,
      amtLDInt: nalAmtLDInt, amtOutbound: nalAmtOutbound, total: nalTotal,
    },
    paricus: {
      inRows: parIn, outRows: parOutCDR, didRows: parDID,
      tfnNumbers: parTFNNums, didNumbers: parDIDNums,
      tfnCount: parTFNCount, didCount: parDIDCount,
      tfnMin: parTFNMin, didMin: parDIDMin, ldInt: parLDInt, outMin: parOutMin,
      amtInbound: parAmtInbound, amtTFNFee: parAmtTFNFee,
      amtLDInt: parAmtLDInt, amtOutbound: parAmtOutbound,
      amtDIDMin: parAmtDIDMin, amtDIDFee: parAmtDIDFee, total: parTotal,
    },
    torch: {
      inRows: torchIn, outRows: torchOutCDR, didRows: torchDID,
      tfnNumbers: torchTFNNums, didNumbers: torchDIDNums,
      tfnCount: torchTFNCount, didCount: torchDIDCount,
      tfnMin: torchTFNMin, didMin: torchDIDMin,
      ldInt: torchLDInt, ldIntra: torchLDIntra, outMin: torchOutMin,
      amtInbound: torchAmtInbound, amtTFNFee: torchAmtTFNFee,
      amtLDInt: torchAmtLDInt, amtLDIntra: torchAmtLDIntra,
      amtOutbound: torchAmtOutbound,
      amtDIDMin: torchAmtDIDMin, amtDIDFee: torchAmtDIDFee, total: torchTotal,
    },
    rates,
  };
}

// ── Cell style helpers ────────────────────────────────────────
const TFN_STYLE = {
  hdr:   { font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: '1F4E79' } }, alignment: { horizontal: 'center' } },
  bold:  { font: { bold: true, name: 'Arial', sz: 10 } },
  blue:  { font: { bold: true, color: { rgb: '2e75b6' }, name: 'Arial', sz: 10 } },
  norm:  { font: { name: 'Arial', sz: 10 } },
  total: { font: { bold: true, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: 'D9E1F2' } } },
};

function setC(ws, r, c, v, s, z) {
  const addr = XLSX.utils.encode_cell({ r, c });
  ws[addr] = { v, t: typeof v === 'number' ? 'n' : 's' };
  if (s) ws[addr].s = s;
  if (z) ws[addr].z = z;
  return ws[addr];
}

const DOLLAR_FMT = '$#,##0.00;($#,##0.00);"-"';

function setRef(ws, maxR, maxC) {
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
}

// ── Build per-customer workbooks ──────────────────────────────

// Helper: apply header style to row 0 of a sheet
function applyHdrRow(ws) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  for (let C = range.s.c; C <= range.e.c; C++) {
    const a = XLSX.utils.encode_cell({ r: 0, c: C });
    if (ws[a]) ws[a].s = TFN_STYLE.hdr;
  }
}

// Helper: build a CDR detail sheet (incalls / out calls / DID) with Amount ($) column
function buildDetailSheet(headers, dataRows, amtColIdx, ratePerMin, totalLabel) {
  const totalRow = new Array(headers.length).fill('');
  totalRow[0] = totalLabel || 'TOTAL';
  let sumAmt = 0;
  const bodyRows = dataRows.map(vals => {
    const row = [...vals];
    const mins = parseFloat(row[amtColIdx]) || 0;
    const amt  = r4(mins * (ratePerMin || 0));
    sumAmt += amt;
    row.push(amt);
    return row;
  });
  sumAmt = r2(sumAmt);
  const sumRow = new Array(headers.length + 1).fill('');
  sumRow[0] = 'TOTAL';
  sumRow[headers.length] = sumAmt;

  const aoa = [[...headers, 'Amount ($)'], ...bodyRows, sumRow];
  const ws  = XLSX.utils.aoa_to_sheet(aoa);
  applyHdrRow(ws);

  // Style body amount col and total row
  const lastRow = aoa.length - 1;
  const lastCol = headers.length;
  for (let r = 1; r <= lastRow; r++) {
    const a = XLSX.utils.encode_cell({ r, c: lastCol });
    if (ws[a] && typeof ws[a].v === 'number') {
      ws[a].s = r === lastRow ? TFN_STYLE.total : TFN_STYLE.norm;
      ws[a].z = DOLLAR_FMT;
    }
  }
  // Style total label
  const tLabel = XLSX.utils.encode_cell({ r: lastRow, c: 0 });
  if (ws[tLabel]) ws[tLabel].s = TFN_STYLE.total;

  return ws;
}

function buildNALWorkbookTFN(c, rates) {
  rates = rates || {};
  const wb = XLSX.utils.book_new();

  // Summary sheet
  {
    const ws = XLSX.utils.aoa_to_sheet([]);
    const HDR = ['Item', 'Service', 'Duration (min)', 'Amount ($)'];
    const inSub  = r2(c.amtInbound + c.amtTFNFee);
    const outSub = r2(c.amtLDInt + c.amtOutbound);
    const rows = [
      HDR,
      ['NAL TFN', 'NorthAmericanLocal- Inbound',  c.tfnMin,  c.amtInbound],
      ['',        'TFN Numbers',                   c.tfnCount, c.amtTFNFee],
      ['',        'Inbound Subtotal',               '',         inSub],
      [],
      HDR,
      ['NAL TFN', 'Long Distance Interstate',       c.ldInt,   c.amtLDInt],
      ['',        'NorthAmericanLocal- Outbound',   c.outMin,  c.amtOutbound],
      ['',        'Outbound Subtotal',              '',         outSub],
      [],
      ['', '', 'TOTAL', c.total],
    ];
    const hdrRows   = new Set([0, 5]);
    const subRows   = new Set([3, 8]);
    const totalRows = new Set([rows.length - 1]);
    rows.forEach((row, r) => row.forEach((v, col) => {
      if (v === '' || v == null) return;
      let s = TFN_STYLE.norm;
      if (hdrRows.has(r))             s = TFN_STYLE.hdr;
      else if (subRows.has(r))        s = TFN_STYLE.total;
      else if (totalRows.has(r) && col === 2) s = TFN_STYLE.bold;
      const z = (col === 3 && typeof v === 'number') ? DOLLAR_FMT : undefined;
      setC(ws, r, col, v, s, z);
    }));
    setRef(ws, rows.length - 1, 3);
    ws['!cols'] = [{wch:14},{wch:34},{wch:18},{wch:14}];
    XLSX.utils.book_append_sheet(wb, ws, 'Summary');
  }

  // Incalls sheet
  {
    const hdrs = ['SourceNumber','TerminationNumber','calldatetime','Callduration_Minutes','','','Number'];
    const data  = c.inRows.map((row, i) => [
      row.SourceNumber, row.TerminationNumber, row.calldatetime, row.Callduration_Minutes,
      '', '', i < c.tfnNumbers.length ? c.tfnNumbers[i] : '',
    ]);
    const ws = buildDetailSheet(hdrs, data, 3, rates.inbound);
    ws['!cols'] = [{wch:14},{wch:14},{wch:22},{wch:20},{wch:4},{wch:4},{wch:14},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, ' Incalls');
  }

  // Out Calls sheet
  {
    const OCOLS = ['OriginationcarrierName','sourcenumber','terminationnumber','starttime','Duration_Minutes'];
    const data  = c.outRows.map(row => OCOLS.map(k => row[k]));
    const ws = buildDetailSheet(OCOLS, data, 4, rates.ipcOutbound);
    ws['!cols'] = [{wch:24},{wch:14},{wch:14},{wch:22},{wch:18},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, 'Out Calls');
  }

  return wb;
}

function buildParicusWorkbookTFN(c, rates) {
  rates = rates || {};
  const wb = XLSX.utils.book_new();

  // Summary
  {
    const ws  = XLSX.utils.aoa_to_sheet([]);
    const HDR = ['Item', 'Service', 'Duration (min)', 'Amount ($)'];
    const inSub  = r2(c.amtInbound + c.amtTFNFee);
    const outSub = r2(c.amtLDInt + c.amtOutbound);
    const didSub = r2(c.amtDIDMin + c.amtDIDFee);
    const rows = [
      HDR,
      ['Paricus TFN', 'Paricus-Inbound',                  c.tfnMin,   c.amtInbound],
      ['',            'TFN Numbers',                       c.tfnCount, c.amtTFNFee],
      ['',            'Inbound Subtotal',                  '',          inSub],
      [],
      HDR,
      ['',            'Paricus Long Distance Interstate',  c.ldInt,    c.amtLDInt],
      ['',            'Paricus-Outbound',                  c.outMin,   c.amtOutbound],
      ['',            'Outbound Subtotal',                 '',          outSub],
      [],
      HDR,
      ['',            'DID Traffic',                       c.didMin,   c.amtDIDMin],
      ['',            'DID Numbers',                       c.didCount, c.amtDIDFee],
      ['',            'DID Subtotal',                      '',          didSub],
      [],
      ['', '', 'TOTAL', c.total],
    ];
    const hdrRows = new Set([0, 5, 10]);
    const subRows = new Set([3, 8, 13]);
    const totalRow = rows.length - 1;
    rows.forEach((row, r) => row.forEach((v, col) => {
      if (v === '' || v == null) return;
      let s = TFN_STYLE.norm;
      if (hdrRows.has(r))                         s = TFN_STYLE.hdr;
      else if (subRows.has(r))                    s = TFN_STYLE.total;
      else if (r === totalRow && col === 2)       s = TFN_STYLE.bold;
      const z = (col === 3 && typeof v === 'number') ? DOLLAR_FMT : undefined;
      setC(ws, r, col, v, s, z);
    }));
    setRef(ws, rows.length - 1, 3);
    ws['!cols'] = [{wch:14},{wch:36},{wch:18},{wch:14}];
    XLSX.utils.book_append_sheet(wb, ws, 'Summary');
  }

  // Paricus - Inbound Toll Free 382
  {
    const hdrs = ['SourceNumber','TerminationNumber','calldatetime','Callduration_Minutes','','','TFN Numbers'];
    const data  = c.inRows.map((row, i) => [
      row.SourceNumber, row.TerminationNumber, row.calldatetime, row.Callduration_Minutes,
      '', '', i < c.tfnNumbers.length ? c.tfnNumbers[i] : '',
    ]);
    const ws = buildDetailSheet(hdrs, data, 3, rates.inbound);
    ws['!cols'] = [{wch:14},{wch:14},{wch:22},{wch:20},{wch:4},{wch:4},{wch:14},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, 'Paricus - Inbound Toll Free 382');
  }

  // Paricus - Dial Out
  {
    const OCOLS = ['OriginationcarrierName','sourcenumber','calledno','starttime','Duration_Minutes'];
    const data  = c.outRows.map(row => OCOLS.map(k => row[k]));
    const ws = buildDetailSheet(OCOLS, data, 4, rates.ipcOutbound);
    ws['!cols'] = [{wch:24},{wch:14},{wch:14},{wch:22},{wch:18},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, 'Paricus - Dial Out ');
  }

  // DID - Verizon MCI Inbound
  {
    const hdrs = ['CustomerName','DID','SourceNumber','CallConnectTime','Duration_Minutes','','','','','','','DID Numbers'];
    const data  = c.didRows.map((row, i) => [
      row.CustomerName, row.DID, row.SourceNumber, row.CallConnectTime, row.Duration_Minutes,
      '','','','','','', i < c.didNumbers.length ? c.didNumbers[i] : '',
    ]);
    const ws = buildDetailSheet(hdrs, data, 4, rates.didTraffic);
    ws['!cols'] = [{wch:14},{wch:14},{wch:14},{wch:22},{wch:18},{wch:4},{wch:4},{wch:4},{wch:4},{wch:4},{wch:4},{wch:14},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, 'DID  - Verizon MCI Inbound');
  }

  return wb;
}

function buildTorchWorkbookTFN(c, rates) {
  rates = rates || {};
  const wb = XLSX.utils.book_new();

  // Summary
  {
    const ws  = XLSX.utils.aoa_to_sheet([]);
    const HDR = ['Item', 'Service', 'Duration (min)', 'Amount ($)'];
    const inSub  = r2(c.amtInbound + c.amtTFNFee);
    const outSub = r2(c.amtLDInt + c.amtLDIntra + c.amtOutbound);
    const didSub = r2(c.amtDIDMin + c.amtDIDFee);
    const rows = [
      HDR,
      ['Torch Wireless TFN', 'Centercom- Inbound',                 c.tfnMin,   c.amtInbound],
      ['',                   'TFN Numbers',                         c.tfnCount, c.amtTFNFee],
      ['',                   'Inbound Subtotal',                    '',          inSub],
      [],
      HDR,
      ['',                   'Centercom Long Distance Interstate',  c.ldInt,    c.amtLDInt],
      ['',                   'Centercom Long Distance Intrastate',  c.ldIntra,  c.amtLDIntra],
      ['',                   'Centercom- Outbound',                 c.outMin,   c.amtOutbound],
      ['',                   'Outbound Subtotal',                   '',          outSub],
      [],
      HDR,
      ['',                   'DID Duration',                        c.didMin,   c.amtDIDMin],
      ['',                   'DID Numbers',                         c.didCount, c.amtDIDFee],
      ['',                   'DID Subtotal',                        '',          didSub],
      [],
      ['', '', 'TOTAL', c.total],
    ];
    const hdrRows = new Set([0, 5, 11]);
    const subRows = new Set([3, 9, 14]);
    const totalRow = rows.length - 1;
    rows.forEach((row, r) => row.forEach((v, col) => {
      if (v === '' || v == null) return;
      let s = TFN_STYLE.norm;
      if (hdrRows.has(r))                         s = TFN_STYLE.hdr;
      else if (subRows.has(r))                    s = TFN_STYLE.total;
      else if (r === totalRow && col === 2)       s = TFN_STYLE.bold;
      const z = (col === 3 && typeof v === 'number') ? DOLLAR_FMT : undefined;
      setC(ws, r, col, v, s, z);
    }));
    setRef(ws, rows.length - 1, 3);
    ws['!cols'] = [{wch:18},{wch:36},{wch:18},{wch:14}];
    XLSX.utils.book_append_sheet(wb, ws, 'Summary');
  }

  // Incalls
  {
    const hdrs = ['SourceNumber','TerminationNumber','calldatetime','Callduration_Minutes','','','Number'];
    const data  = c.inRows.map((row, i) => [
      row.SourceNumber, row.TerminationNumber, row.calldatetime, row.Callduration_Minutes,
      '', '', i < c.tfnNumbers.length ? c.tfnNumbers[i] : '',
    ]);
    const ws = buildDetailSheet(hdrs, data, 3, rates.inbound);
    ws['!cols'] = [{wch:14},{wch:14},{wch:22},{wch:20},{wch:4},{wch:4},{wch:14},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, ' Incalls');
  }

  // Out Calls
  {
    const OCOLS = ['OriginationcarrierName','sourcenumber','terminationnumber','starttime','Duration_Minutes'];
    const data  = c.outRows.map(row => OCOLS.map(k => row[k]));
    const ws = buildDetailSheet(OCOLS, data, 4, rates.ipcOutbound);
    ws['!cols'] = [{wch:24},{wch:14},{wch:14},{wch:22},{wch:18},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, 'Out Calls');
  }

  // DID sheet
  {
    const hdrs = ['CustomerName','DID','SourceNumber','CallConnectTime','Duration_Minutes','','','','DID Numbers'];
    const data  = c.didRows.map((row, i) => [
      row.CustomerName, row.DID, row.SourceNumber, row.CallConnectTime, row.Duration_Minutes,
      '','','', i < c.didNumbers.length ? c.didNumbers[i] : '',
    ]);
    const ws = buildDetailSheet(hdrs, data, 4, rates.didTraffic);
    ws['!cols'] = [{wch:14},{wch:14},{wch:14},{wch:22},{wch:18},{wch:4},{wch:4},{wch:4},{wch:14},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, 'DID ');
  }

  return wb;
}

// ── Build Audit sheet ─────────────────────────────────────────
function buildAuditSheet(wb, customers, vendorBills, rates) {
  const { nal, paricus, torch } = customers;
  rates = rates || {};

  const b382  = vendorBills.find(b => b.label === '382 Communications') || {};
  const bMCI  = vendorBills.find(b => b.label === 'MCI/Verizon')        || {};
  const bLI   = vendorBills.find(b => b.label === 'Lumen Interstate')   || {};
  const bLa   = vendorBills.find(b => b.label === 'Lumen Intrastate')   || {};
  const bIPC  = vendorBills.find(b => b.label === 'IPC')                || {};

  // Profit% formula matching reference: P&L / (vendor + invoice)
  const pct = (pl, vendor, invoice) =>
    (vendor + invoice) > 0 ? r2((pl / (vendor + invoice)) * 100) : 0;

  const ws  = XLSX.utils.aoa_to_sheet([]);
  const W   = DOLLAR_FMT;
  const PCT = '0.00"%"';

  // Style aliases
  const S = {
    title:   { font: { bold: true, name: 'Arial', sz: 11 } },
    hdr:     TFN_STYLE.hdr,
    subhdr:  { font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: '2E75B6' } } },
    secLeft: { font: { bold: true, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: 'D6E4F0' } } },
    norm:    TFN_STYLE.norm,
    bold:    TFN_STYLE.bold,
    total:   { font: { bold: true, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: 'BDD7EE' } } },
    grand:   { font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Arial', sz: 11 }, fill: { fgColor: { rgb: '1F4E79' } } },
    plPos:   { font: { bold: true, color: { rgb: '276749' }, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: 'BDD7EE' } } },
    plNeg:   { font: { bold: true, color: { rgb: 'C53030' }, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: 'BDD7EE' } } },
    numFee:  { font: { italic: true, name: 'Arial', sz: 10 }, fill: { fgColor: { rgb: 'EBF5FB' } } },
  };

  // Column layout:
  // 0: Label        1: Bill Amt   2: Duration   3: Calls/Units  4: Vendor Rate
  // 5: (sep)
  // 6: Customer     7: Duration   8: Invoice    9: Rate w/ Markup
  // 10: (sep)   11: P&L   12: Profit%
  const NCOLS = 13;

  let row = 0;

  const sc = (r, c, v, s, z) => { if (v !== '' && v != null) setC(ws, r, c, v, s, z); };

  // ── Title row ──────────────────────────────────────────────
  sc(row, 0, 'TFN Billing Audit', S.title);
  sc(row, 11, 'Profit / Loss', S.bold);
  sc(row, 12, 'Profit %', S.bold);
  row++;
  sc(row, 1, 'Vendors', S.bold);
  sc(row, 6, 'Invoice to Client', S.bold);
  sc(row, 11, '(note: P&L / (cost + revenue))', { font: { italic: true, name: 'Arial', sz: 9, color: { rgb: '718096' } } });
  row++;
  row++; // blank

  // ── Section builder ─────────────────────────────────────────
  // leftRows: [[label, billAmt, duration, calls, vendorRate], ...]
  // rightRows: [[customer, duration, invoiceAmt, rate], ...]  or null for blank line
  // numFeeRows: [[label, count, amount, unitRate], ...] — styled differently
  function writeSection(sectionLabel, leftRows, rightRows, numFeeRows, vendorTotal, invoiceTotal) {
    // Section header
    for (let c = 0; c < NCOLS; c++) sc(row, c, c === 0 ? sectionLabel : '', S.secLeft);
    row++;

    // Column headers
    const hdrL = ['Date / Item', 'Bill Amount', 'Duration (min)', 'Calls', 'Vendor Rate'];
    const hdrR = ['Customer', 'Duration (min)', 'Invoice Amount', 'Rate w/ Markup'];
    hdrL.forEach((v, c) => sc(row, c, v, S.hdr));
    hdrR.forEach((v, c) => sc(row, 6 + c, v, S.hdr));
    row++;

    // Merge left + right traffic rows
    const maxTraffic = Math.max(leftRows.length, rightRows.length);
    for (let i = 0; i < maxTraffic; i++) {
      const L = leftRows[i]  || [null, null, null, null, null];
      const R = rightRows[i] || [null, null, null, null];
      if (R[0] === null) { row++; continue; } // blank row
      sc(row, 0, L[0], S.norm);
      if (L[1] != null) sc(row, 1, L[1], S.norm, W);
      if (L[2] != null) sc(row, 2, L[2], S.norm);
      if (L[3] != null) sc(row, 3, L[3], S.norm);
      if (L[4] != null) sc(row, 4, L[4], S.norm, '$0.0000');
      sc(row, 6, R[0], S.norm);
      if (R[1] != null) sc(row, 7, R[1], S.norm);
      if (R[2] != null) sc(row, 8, R[2], S.norm, W);
      if (R[3] != null) sc(row, 9, R[3], S.norm, '$0.0000');
      row++;
    }

    // Number fee rows (italic, shaded differently)
    if (numFeeRows && numFeeRows.length) {
      for (const nf of numFeeRows) {
        sc(row, 6, nf[0], S.numFee);
        sc(row, 7, nf[1], S.numFee);
        sc(row, 8, nf[2], S.numFee, W);
        sc(row, 9, nf[3], S.numFee, '$0.0000');
        row++;
      }
    }

    // Total row
    const pl  = r2(invoiceTotal - vendorTotal);
    const pp  = pct(pl, vendorTotal, invoiceTotal);
    const plS = pl >= 0 ? S.plPos : S.plNeg;
    sc(row, 0, 'Total:', S.total);
    sc(row, 1, vendorTotal,   S.total, W);
    sc(row, 6, 'Invoice Total', S.total);
    sc(row, 8, invoiceTotal,  S.total, W);
    sc(row, 11, pl,  plS, W);
    sc(row, 12, pp,  plS, PCT);
    row++;
    row++; // blank between sections

    return { vendor: vendorTotal, invoice: invoiceTotal, pl, pp };
  }

  // CDR minute totals (derived from customer data for vendor-side comparison)
  const cdrTFNMin  = r2(nal.tfnMin  + paricus.tfnMin  + torch.tfnMin);
  const cdrLDInter = r2(nal.ldInt   + paricus.ldInt   + torch.ldInt);
  const cdrLDIntra = r2(torch.ldIntra || 0);
  const cdrDIDMin  = r2((paricus.didMin || 0) + (torch.didMin || 0));
  const cdrOutMin  = r2(nal.outMin  + paricus.outMin  + torch.outMin);

  // ── 382 Communications ───────────────────────────────────────
  const weeks = b382.weeks || [];
  const wkLabels = ['Week 1','Week 2','Week 3','Week 4'];
  const left382 = wkLabels.map((lbl, i) => {
    const w = weeks[i] || {};
    return w.amount != null ? [lbl, w.amount, null, null, null] : null;
  }).filter(Boolean);
  // Append CDR minute total row on vendor side for minute reconciliation
  left382.push(['CDR Total Minutes (billed)', null, cdrTFNMin, null, null]);

  const right382 = [
    ['Centercom (Torch)', torch.tfnMin, torch.amtInbound, rates.inbound],
    ['NorthAmericanLocal (NAL)', nal.tfnMin, nal.amtInbound, rates.inbound],
    ['Paricus', paricus.tfnMin, paricus.amtInbound, rates.inbound],
  ];

  const numFees382 = [
    ['NAL — TFN Numbers',    nal.tfnCount,    nal.amtTFNFee,    TFN_FEES.tfnNumberFee],
    ['Torch — TFN Numbers',  torch.tfnCount,  torch.amtTFNFee,  TFN_FEES.tfnNumberFee],
    ['Paricus — TFN Numbers',paricus.tfnCount,paricus.amtTFNFee,TFN_FEES.tfnNumberFee],
  ];

  const vendor382  = b382.amount || 0;
  const invoice382 = r2(nal.amtInbound + paricus.amtInbound + torch.amtInbound +
                        nal.amtTFNFee  + paricus.amtTFNFee  + torch.amtTFNFee);
  const s382 = writeSection('382 Communications', left382, right382, numFees382, vendor382, invoice382);

  // ── Lumen ────────────────────────────────────────────────────
  const leftLumen = [
    bLI.amount != null ? ['Long Distance Interstate', bLI.amount, cdrLDInter, null, null] : ['Long Distance Interstate', null, cdrLDInter, null, null],
    bLa.amount != null ? ['Long Distance Intrastate', bLa.amount, cdrLDIntra, null, null] : ['Long Distance Intrastate', null, cdrLDIntra, null, null],
  ];

  const rightLumen = [
    ['Centercom — LD Interstate', torch.ldInt,   torch.amtLDInt,   rates.ldInterstate],
    ['Centercom — LD Intrastate', torch.ldIntra, torch.amtLDIntra, rates.ldIntrastate],
    ['NAL — LD Interstate',       nal.ldInt,     nal.amtLDInt,     rates.ldInterstate],
    ['Paricus — LD Interstate',   paricus.ldInt, paricus.amtLDInt, rates.ldInterstate],
  ];

  const vendorLumen  = r2((bLI.amount || 0) + (bLa.amount || 0));
  const invoiceLumen = r2(torch.amtLDInt + torch.amtLDIntra + nal.amtLDInt + paricus.amtLDInt);
  const sLumen = writeSection('Lumen — Long Distance', leftLumen, rightLumen, null, vendorLumen, invoiceLumen);

  // ── MCI / Verizon ────────────────────────────────────────────
  const leftMCI = [
    ['DID Inbound / Outbound', bMCI.amount != null ? bMCI.amount : null, cdrDIDMin, null, null],
  ];

  const rightMCI = [
    ['Centercom — DID Traffic', torch.didMin,   torch.amtDIDMin,   rates.didTraffic],
    ['Paricus — DID Traffic',   paricus.didMin, paricus.amtDIDMin, rates.didTraffic],
  ];

  const numFeesMCI = [
    ['Torch — DID Numbers',  torch.didCount,   torch.amtDIDFee,   TFN_FEES.didNumberFee],
    ['Paricus — DID Numbers',paricus.didCount, paricus.amtDIDFee, TFN_FEES.didNumberFee],
  ];

  const vendorMCI  = bMCI.amount || 0;
  const invoiceMCI = r2(torch.amtDIDMin + paricus.amtDIDMin + torch.amtDIDFee + paricus.amtDIDFee);
  const sMCI = writeSection('MCI / Verizon — DID', leftMCI, rightMCI, numFeesMCI, vendorMCI, invoiceMCI);

  // ── IPC ──────────────────────────────────────────────────────
  const leftIPC = [
    ['IPC Outbound', bIPC.amount != null ? bIPC.amount : null, cdrOutMin, null, null],
  ];

  const rightIPC = [
    ['NAL — Outbound',    nal.outMin,    nal.amtOutbound,    rates.ipcOutbound],
    ['Torch — Outbound',  torch.outMin,  torch.amtOutbound,  rates.ipcOutbound],
    ['Paricus — Outbound',paricus.outMin,paricus.amtOutbound,rates.ipcOutbound],
  ];

  const vendorIPC  = bIPC.amount || 0;
  const invoiceIPC = r2(nal.amtOutbound + torch.amtOutbound + paricus.amtOutbound);
  const sIPC = writeSection('IPC — Outbound', leftIPC, rightIPC, null, vendorIPC, invoiceIPC);

  // ── Grand Total ──────────────────────────────────────────────
  const gVendor  = r2(s382.vendor  + sLumen.vendor  + sMCI.vendor  + sIPC.vendor);
  const gInvoice = r2(s382.invoice + sLumen.invoice + sMCI.invoice + sIPC.invoice);
  const gPL      = r2(gInvoice - gVendor);
  const gPct     = pct(gPL, gVendor, gInvoice);
  const gS       = gPL >= 0 ? S.plPos : S.plNeg;

  for (let c = 0; c < NCOLS; c++) sc(row, c, '', S.grand);
  sc(row, 0, 'GRAND TOTAL',  S.grand);
  sc(row, 1, gVendor,        { ...S.grand }, W);
  sc(row, 6, 'Total Billed', { ...S.grand });
  sc(row, 8, gInvoice,       { ...S.grand }, W);
  sc(row, 11, gPL,           { ...gS, fill: S.grand.fill }, W);
  sc(row, 12, gPct,          { ...gS, fill: S.grand.fill }, PCT);

  setRef(ws, row, NCOLS - 1);
  ws['!cols'] = [
    {wch:32},{wch:14},{wch:14},{wch:10},{wch:13},{wch:2},
    {wch:30},{wch:14},{wch:14},{wch:14},{wch:2},{wch:14},{wch:12},
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Audit');
}

// ── Build workpaper comparison ────────────────────────────────
function buildWorkpaperTFN(customers, vendorBills, period, rates) {
  rates = rates || {};
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([]);

  const { nal, paricus, torch } = customers;
  const totalCustomer = r2(nal.total + paricus.total + torch.total);
  const totalVendor   = r2(vendorBills.reduce((s, b) => s + (b.amount || 0), 0));
  const netProfit     = r2(totalCustomer - totalVendor);
  const margin        = totalCustomer > 0 ? r2((netProfit / totalCustomer) * 100) : 0;

  let row = 0;
  const W = '$#,##0.00;($#,##0.00);"-"';

  // Title
  setC(ws, row,   0, '321 Communications — TFN Billing Workpaper', { font: { bold: true, name: 'Arial', sz: 13 } });
  setC(ws, row+1, 0, period || '', TFN_STYLE.norm);
  row += 3;

  // ── VENDOR COSTS ──
  setC(ws, row, 0, 'VENDOR COSTS',       TFN_STYLE.hdr);
  setC(ws, row, 1, 'Invoice / File',     TFN_STYLE.hdr);
  setC(ws, row, 2, 'Amount',             TFN_STYLE.hdr);
  setC(ws, row, 3, 'Parse Status',       TFN_STYLE.hdr);
  row++;

  for (const bill of vendorBills) {
    setC(ws, row, 1, bill.filename, TFN_STYLE.norm);
    if (bill.total !== null) {
      setC(ws, row, 2, bill.total, TFN_STYLE.norm, W);
      setC(ws, row, 3, 'Auto-parsed', { font: { color: { rgb: '276749' }, name: 'Arial', sz: 10 } });
    } else {
      setC(ws, row, 2, bill.manualTotal || 0, TFN_STYLE.norm, W);
      setC(ws, row, 3, 'Manual entry', { font: { color: { rgb: 'DD6B20' }, name: 'Arial', sz: 10 } });
    }
    row++;
  }

  setC(ws, row, 0, 'Total Vendor Cost',  TFN_STYLE.total);
  setC(ws, row, 2, totalVendor,          TFN_STYLE.total, W);
  row += 2;

  // ── CUSTOMER BILLINGS ──
  setC(ws, row, 0, 'CUSTOMER BILLINGS', TFN_STYLE.hdr);
  setC(ws, row, 1, 'Customer',          TFN_STYLE.hdr);
  setC(ws, row, 2, 'Amount Billed',     TFN_STYLE.hdr);
  setC(ws, row, 3, 'Detail',            TFN_STYLE.hdr);
  row++;

  const custRows = [
    { name: 'North American Local (NAL)', total: nal.total,
      detail: `TFN In ${nal.tfnMin} min | LD Int ${nal.ldInt} min | Out ${nal.outMin} min | ${nal.tfnCount} TFN#` },
    { name: 'Paricus', total: paricus.total,
      detail: `TFN In ${paricus.tfnMin} min | LD Int ${paricus.ldInt} min | Out ${paricus.outMin} min | DID ${paricus.didMin} min | ${paricus.tfnCount} TFN# | ${paricus.didCount} DID#` },
    { name: 'Torch Wireless (Centercom)', total: torch.total,
      detail: `TFN In ${torch.tfnMin} min | LD Int ${torch.ldInt} min | Intra ${torch.ldIntra} min | Out ${torch.outMin} min | DID ${torch.didMin} min | ${torch.tfnCount} TFN# | ${torch.didCount} DID#` },
  ];

  for (const c of custRows) {
    setC(ws, row, 1, c.name,   TFN_STYLE.norm);
    setC(ws, row, 2, c.total,  TFN_STYLE.norm, W);
    setC(ws, row, 3, c.detail, TFN_STYLE.norm);
    row++;
  }

  setC(ws, row, 0, 'Total Customer Billing', TFN_STYLE.total);
  setC(ws, row, 2, totalCustomer,             TFN_STYLE.total, W);
  row += 2;

  // ── PROFIT SUMMARY ──
  setC(ws, row, 0, 'PROFIT SUMMARY', TFN_STYLE.hdr);
  row++;
  setC(ws, row, 0, 'Total Vendor Cost',     TFN_STYLE.norm);
  setC(ws, row, 2, totalVendor,              TFN_STYLE.norm, W);
  row++;
  setC(ws, row, 0, 'Total Customer Billing', TFN_STYLE.norm);
  setC(ws, row, 2, totalCustomer,             TFN_STYLE.norm, W);
  row++;
  setC(ws, row, 0, 'Net Profit / (Loss)', TFN_STYLE.total);
  setC(ws, row, 2, netProfit,              { font: { bold: true, name: 'Arial', sz: 10, color: { rgb: netProfit >= 0 ? '276749' : 'C53030' } } }, W);
  row++;
  setC(ws, row, 0, 'Margin %', TFN_STYLE.norm);
  setC(ws, row, 2, margin / 100, TFN_STYLE.norm, '0.0%;(0.0%);"-"');

  setRef(ws, row, 3);
  ws['!cols'] = [{wch:30},{wch:34},{wch:18},{wch:70}];
  XLSX.utils.book_append_sheet(wb, ws, 'TFN Workpaper');

  // Rate Reference sheet
  const refWS = XLSX.utils.aoa_to_sheet([]);
  const refRows = [
    ['Rate Reference — 321 Comm TFN', '', ''],
    ['', '', ''],
    ['Service', 'Unit', 'Rate'],
    ['TFN Inbound (382)', 'per minute', rates.inbound      || 0],
    ['TFN Number Fee', 'per number/month', TFN_FEES.tfnNumberFee],
    ['DID Traffic (MCI/Verizon)', 'per minute', rates.didTraffic   || 0],
    ['DID Number Fee', 'per number/month', TFN_FEES.didNumberFee],
    ['Long Distance Interstate (Lumen)', 'per minute', rates.ldInterstate || 0],
    ['Long Distance Intrastate (Lumen)', 'per minute', rates.ldIntrastate || 0],
    ['IPC Outbound', 'per minute', rates.ipcOutbound  || 0],
  ];
  refRows.forEach((r, ri) => r.forEach((v, ci) => {
    if (v === '') return;
    const s = ri === 0 ? { font: { bold: true, name: 'Arial', sz: 12 } } :
              ri === 2 ? TFN_STYLE.hdr :
              TFN_STYLE.norm;
    const z = (ci === 2 && typeof v === 'number') ? '$0.0000' : undefined;
    setC(refWS, ri, ci, v, s, z);
  }));
  setRef(refWS, refRows.length - 1, 2);
  refWS['!cols'] = [{wch:36},{wch:22},{wch:14}];
  XLSX.utils.book_append_sheet(wb, refWS, 'Rate Reference');

  buildAuditSheet(wb, customers, vendorBills, rates);

  return wb;
}

// ── Try to extract total from PDF text ───────────────────────
async function extractPDFTotal(filePath) {
  if (!pdfParse) return null;
  try {
    const buf  = fs.readFileSync(filePath);
    const data = await pdfParse(buf);
    const text = data.text || '';

    // Match patterns like "Total Due $1,234.56" or "Amount Due: 1,234.56"
    const patterns = [
      /(?:total\s+(?:amount\s+)?due|amount\s+due|invoice\s+total|total\s+charges?|balance\s+due)[:\s$]*(\d[\d,]*\.\d{2})/i,
      /(?:total)[:\s$]+(\d[\d,]*\.\d{2})\s*$/im,
    ];
    for (const re of patterns) {
      const m = text.match(re);
      if (m) return parseFloat(m[1].replace(/,/g, ''));
    }

    // Last-resort: find all dollar amounts in last 500 chars and take the largest
    const tail    = text.slice(-500);
    const amounts = [...tail.matchAll(/\$?\s*(\d[\d,]*\.\d{2})/g)]
      .map(m => parseFloat(m[1].replace(/,/g, '')))
      .filter(n => n > 0);
    if (amounts.length > 0) return Math.max(...amounts);
  } catch(e) { /* ignore */ }
  return null;
}

// ── In-memory temp file store ────────────────────────────────
const tfnTempStore = new Map(); // token → { files: {name→path}, ts }

function tfnStorePaths(paths) {
  const token = `tfn_${Date.now()}_${Math.random().toString(36).slice(2,7)}`;
  tfnTempStore.set(token, { files: paths, ts: Date.now() });
  // Expire after 30 minutes
  setTimeout(() => {
    const entry = tfnTempStore.get(token);
    if (entry) {
      for (const p of Object.values(entry.files)) { try { if (fs.existsSync(p)) fs.unlinkSync(p); } catch(e){} }
      tfnTempStore.delete(token);
    }
  }, 30 * 60 * 1000);
  return token;
}

// ── Upload config ────────────────────────────────────────────
const tfnUpload = upload.fields([
  { name: 'gopalFile',        maxCount: 1 },
  { name: 'ipcFile',          maxCount: 1 },
  { name: 'inventoryFile',    maxCount: 1 },
  { name: 'didInventoryFile', maxCount: 1 },
  { name: 'bill382',          maxCount: 4 },
  { name: 'billMCI',          maxCount: 1 },
  { name: 'billLumen',        maxCount: 1 },
  { name: 'billIPC',          maxCount: 1 },
]);

// ── Try to extract a named amount from PDF text ──────────────
// Looks for keyword near a dollar amount in the last ~1500 chars
async function extractPDFAmount(filePath, keywords) {
  if (!pdfParse || !filePath) return null;
  try {
    const data = await pdfParse(fs.readFileSync(filePath));
    const text = data.text || '';
    const tail = text.slice(-1500);
    for (const kw of keywords) {
      const re = new RegExp(kw + '[^\\d$]*\\$?([\\d,]+\\.\\d{2})', 'i');
      const m  = tail.match(re) || text.match(re);
      if (m) return parseFloat(m[1].replace(/,/g, ''));
    }
    // fallback: largest dollar amount in tail
    const nums = [...tail.matchAll(/\$?([\d,]+\.\d{2})/g)]
      .map(m => parseFloat(m[1].replace(/,/g, ''))).filter(n => n > 0);
    return nums.length ? Math.max(...nums) : null;
  } catch { return null; }
}

// ── POST /api/tfn/process ────────────────────────────────────
app.post('/api/tfn/process', tfnUpload, async (req, res) => {
  const gopalFile        = req.files?.['gopalFile']?.[0];
  const ipcFile          = req.files?.['ipcFile']?.[0];
  const inventoryFile    = req.files?.['inventoryFile']?.[0];
  const didInventoryFile = req.files?.['didInventoryFile']?.[0];
  const bill382Files     = req.files?.['bill382'] || [];
  const billMCIFile      = req.files?.['billMCI']?.[0];
  const billLumenFile    = req.files?.['billLumen']?.[0];
  const billIPCFile      = req.files?.['billIPC']?.[0];
  const manualBody       = req.body || {};

  if (!gopalFile) return res.status(400).json({ error: 'Gopal CDR file is required.' });

  const cleanupFiles = (...files) => files.forEach(f => { try { if (f?.path && fs.existsSync(f.path)) fs.unlinkSync(f.path); } catch(e){} });

  try {
    const gopal              = parseGopalFile(gopalFile.path);
    const inventoryCounts    = inventoryFile    ? parseTFNInventory(inventoryFile.path)    : {};
    const ipcOutbound        = ipcFile          ? parseIPCOutboundFile(ipcFile.path)       : {};
    const didInventoryCounts = didInventoryFile ? parseDIDInventory(didInventoryFile.path) : {};

    // ── Parse vendor bill amounts ─────────────────────────────
    const parseAmt = async (file, keywords, manualKey) => {
      const manual = parseFloat(manualBody[manualKey] || '') || null;
      if (manual) return { amount: manual, parsedOK: false, manualUsed: true };
      const parsed = await extractPDFAmount(file?.path, keywords);
      return { amount: parsed, parsedOK: parsed !== null, manualUsed: false };
    };

    // Parse 4 weekly 382 invoices
    const weeks382 = await Promise.all(
      [1,2,3,4].map(w => parseAmt(bill382Files[w-1], ['Total Due','Amount Due','Invoice Total'], `manual382_w${w}`))
    );
    // Total 382 = sum of all weeks that have an amount
    const total382Amount = weeks382.reduce((s, w) => s + (w.amount || 0), 0);
    const any382Parsed   = weeks382.some(w => w.parsedOK);
    const any382Manual   = weeks382.some(w => w.manualUsed);
    const b382 = {
      amount:     total382Amount || null,
      parsedOK:   any382Parsed,
      manualUsed: any382Manual,
      weeks:      weeks382,
    };

    const [bMCI, bLumenInter, bLumenIntra, bIPC] = await Promise.all([
      parseAmt(billMCIFile,   ['Total Due','Amount Due','Invoice Total'], 'manualMCI'),
      parseAmt(billLumenFile, ['Interstate'],                             'manualLumenInter'),
      parseAmt(billLumenFile, ['Intrastate'],                            'manualLumenIntra'),
      parseAmt(billIPCFile,   ['Total Due','Amount Due','Invoice Total'], 'manualIPC'),
    ]);

    const vendorBills = [
      { label: '382 Communications', filenames: bill382Files.map(f=>f.originalname), ...b382 },
      { label: 'MCI/Verizon',        filename: billMCIFile?.originalname  || null, ...bMCI },
      { label: 'Lumen Interstate',   filename: billLumenFile?.originalname|| null, ...bLumenInter },
      { label: 'Lumen Intrastate',   filename: billLumenFile?.originalname|| null, ...bLumenIntra },
      { label: 'IPC',                filename: billIPCFile?.originalname  || null, ...bIPC },
    ];

    const needsManualEntry = vendorBills.some(b => b.amount == null);

    if (needsManualEntry && !Object.keys(manualBody).length) {
      cleanupFiles(gopalFile, ipcFile, inventoryFile, didInventoryFile, ...bill382Files, billMCIFile, billLumenFile, billIPCFile);
      return res.json({ needsManualEntry: true, vendorBills });
    }

    // ── Compute dynamic rates ─────────────────────────────────
    const totMins = computeTotalMins(gopal.summary, ipcOutbound);
    const invoices = {
      amt382:        b382.amount        || 0,
      amtMCI:        bMCI.amount        || 0,
      amtLumenInter: bLumenInter.amount || 0,
      amtLumenIntra: bLumenIntra.amount || 0,
      amtIPC:        bIPC.amount        || 0,
    };
    const rates = computeRates(invoices, totMins);

    const customers = computeTFNCustomers(gopal, inventoryCounts, ipcOutbound, didInventoryCounts, rates);

    // Determine period label
    const summaryWS = XLSX.readFile(gopalFile.path).Sheets['Summary'];
    const firstCell = summaryWS ? summaryWS['A1'] : null;
    let periodLabel = '';
    if (firstCell && typeof firstCell.v === 'number') {
      const d = XLSX.SSF.parse_date_code(firstCell.v);
      periodLabel = `${d.y}`;
    }

    // Build Excel files
    const nalWb     = buildNALWorkbookTFN(customers.nal,     rates);
    const paricusWb = buildParicusWorkbookTFN(customers.paricus, rates);
    const torchWb   = buildTorchWorkbookTFN(customers.torch,   rates);
    const wpWb      = buildWorkpaperTFN(customers, vendorBills, `321 Comm TFN — ${periodLabel}`, rates);

    const ts = Date.now();
    const outPaths = {
      'NAL':       path.join('uploads', `tfn_nal_${ts}.xlsx`),
      'Paricus':   path.join('uploads', `tfn_paricus_${ts}.xlsx`),
      'Torch':     path.join('uploads', `tfn_torch_${ts}.xlsx`),
      'Workpaper': path.join('uploads', `tfn_workpaper_${ts}.xlsx`),
    };

    const xlsxOpts = { bookType: 'xlsx', type: 'binary', cellStyles: true };
    XLSX.writeFile(nalWb,     outPaths['NAL'],       xlsxOpts);
    XLSX.writeFile(paricusWb, outPaths['Paricus'],   xlsxOpts);
    XLSX.writeFile(torchWb,   outPaths['Torch'],     xlsxOpts);
    XLSX.writeFile(wpWb,      outPaths['Workpaper'], xlsxOpts);

    cleanupFiles(gopalFile, ipcFile, inventoryFile, didInventoryFile, ...bill382Files, billMCIFile, billLumenFile, billIPCFile);

    const token = tfnStorePaths(outPaths);

    res.json({
      token,
      vendorBills,
      rates,
      totMins,
      summary: {
        nal:    { total: customers.nal.total,    tfnNums: customers.nal.tfnCount,    tfnMin: customers.nal.tfnMin },
        paricus:{ total: customers.paricus.total, tfnNums: customers.paricus.tfnCount, didNums: customers.paricus.didCount, tfnMin: customers.paricus.tfnMin, didMin: customers.paricus.didMin },
        torch:  { total: customers.torch.total,  tfnNums: customers.torch.tfnCount,  didNums: customers.torch.didCount, tfnMin: customers.torch.tfnMin, didMin: customers.torch.didMin },
        totalCustomer: r2(customers.nal.total + customers.paricus.total + customers.torch.total),
        totalVendor:   r2(vendorBills.reduce((s, b) => s + (b.amount || 0), 0)),
      },
      needsManualEntry,
    });
  } catch (err) {
    console.error(err);
    cleanupFiles(gopalFile, ipcFile, inventoryFile, didInventoryFile, ...bill382Files, billMCIFile, billLumenFile, billIPCFile);
    res.status(500).json({ error: err.message });
  }
});

// ── GET /api/tfn/download/:token/:type ──────────────────────
app.get('/api/tfn/download/:token/:type', (req, res) => {
  const entry = tfnTempStore.get(req.params.token);
  if (!entry) return res.status(404).json({ error: 'File not found or expired. Please re-process.' });
  const filePath = entry.files[req.params.type];
  if (!filePath || !fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found.' });

  const fileNames = {
    'NAL':      'NAL_TFN_Customer_File.xlsx',
    'Paricus':  'Paricus_TFN_Customer_File.xlsx',
    'Torch':    'Torch_TFN_Customer_File.xlsx',
    'Workpaper':'TFN_Workpaper_Summary.xlsx',
  };
  res.download(filePath, fileNames[req.params.type] || 'output.xlsx');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`AT&T Data Processor running at http://localhost:${PORT}`));
