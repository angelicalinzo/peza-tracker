const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const excelPath = path.join('C:', 'Users', 'C5269573', 'OneDrive - SAP SE', '@SAP', 'AI Learning', 'Concur PEZA Registered Devices', 'PEZA', 'PEZA Registered Devices Main Tracker.xlsx');

const wb   = XLSX.readFile(excelPath);
const ws   = wb.Sheets['Main'];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });

// Helper: format Excel serial date → "DD-Mon-YY"
function fmtDate(v) {
  if (!v) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return String(v);
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return String(d.d).padStart(2,'0') + '-' + months[d.m - 1] + '-' + String(d.y).slice(-2);
  }
  return String(v).trim();
}

function clean(v) {
  if (v === null || v === undefined) return '';
  const s = String(v).trim();
  return s === 'N/A' ? 'N/A' : s;
}

function fmt(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' && v > 1000) {
    return 'PHP ' + v.toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return clean(v);
}

const dataRows = rows.slice(1).filter(r => r && r.some(c => c !== null && c !== undefined && c !== ''));

// Excel column indices (0-based):
// [0]  Booklet No.
// [1]  Device Category
// [2]  TUSA/ConcurOnly/SAP-I4P-ISP
// [3]  Description
// [4]  Qty
// [5]  Type of Permit
// [6]  PEZA Permit Date
// [7]  PEZA Permit No.
// [8]  Name of Supplier
// [9]  Address of Supplier
// [10] Invoice Date
// [11] Invoice Number
// [12] DR Number
// [13] PO Number
// [14] Serial Number
// [15] Equipment Number
// [16] Asset Number
// [17] Invoice Copy? (Y/N)
// [18] DR Copy? (Y/N)
// [19] Acquisition Cost/Value* (USD/PhP)
// [20] Acquisition Cost/Value* (PhP)
// [21] Netbook Value* (USD/PhP)
// [22] Acquisition Date
// [23] Age of Equipment
// [24] Import/Local (long text)
// [25] Condition
// [26] Floor Location
// [27] Status
// [28] ABC Indicator
// [29] Remarks
// [30] SIMS PAC Order

const records = dataRows.map((r, i) => {
  const g = idx => (r[idx] !== null && r[idx] !== undefined) ? r[idx] : '';
  return [
    i + 1,            // [0]  rowNum
    clean(g(27)),     // [1]  status
    clean(g(14)),     // [2]  serialNumber
    clean(g(15)),     // [3]  equipmentNumber
    clean(g(1)),      // [4]  deviceCategory
    clean(g(3)),      // [5]  description
    clean(g(16)),     // [6]  assetNumber
    clean(g(2)),      // [7]  tusaType
    clean(g(0)),      // [8]  bookletNo
    clean(g(7)),      // [9]  pezaPermitNo
    fmtDate(g(6)),    // [10] pezaPermitDate
    clean(g(8)),      // [11] supplierName
    clean(g(9)),      // [12] supplierAddress
    fmtDate(g(10)),   // [13] invoiceDate
    clean(g(11)),     // [14] invoiceNumber
    clean(g(13)),     // [15] poNumber
    fmt(g(19)),       // [16] acqCostUsd
    fmt(g(20)),       // [17] acqCostPhp
    fmt(g(21)),       // [18] netbookValue
    fmtDate(g(22)),   // [19] acquisitionDate
    clean(g(23)),     // [20] ageOfEquip
    clean(g(24)),     // [21] importLocal
    clean(g(25)),     // [22] condition
    clean(g(28)),     // [23] abcIndicator
    clean(g(29)),     // [24] remarks
    clean(g(30)),     // [25] simsPacOrder
    // Extra fields for detail modal (not shown as table columns)
    clean(g(5)),      // [26] permitType
    clean(g(26)),     // [27] floorLocation
    clean(g(12)),     // [28] drNumber
    clean(g(17)),     // [29] invoiceCopy
    clean(g(18)),     // [30] drCopy
    clean(g(4)),      // [31] qty
  ];
});

const js = 'const PEZA_DATA=' + JSON.stringify(records) + ';';
fs.writeFileSync(path.join(__dirname, 'data.js'), js, 'utf8');
console.log('data.js written —', records.length, 'records');
