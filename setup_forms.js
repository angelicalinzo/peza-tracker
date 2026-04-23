const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const excelPath = path.join('C:', 'Users', 'C5269573', 'OneDrive - SAP SE', '@SAP', 'AI Learning', 'Concur PEZA Registered Devices', 'PEZA', 'PEZA Registered Devices Main Tracker.xlsx');
const formsBase = path.join('C:', 'Users', 'C5269573', 'OneDrive - SAP SE', '@SAP', 'AI Learning', 'Concur PEZA Registered Devices', 'PEZA', 'PEZA 8105 Forms');
const tusaSrc   = path.join('C:', 'Users', 'C5269573', 'OneDrive - SAP SE', '@SAP', 'AI Learning', 'Concur PEZA Registered Devices', 'PEZA', 'TUSA', '8112_v4 signed.pdf');
const outForms  = path.join(__dirname, 'forms');

// 1. Create forms/ dir
if (!fs.existsSync(outForms)) fs.mkdirSync(outForms);

// 2. Copy all PDFs and build map: permitNo → relative path
const pdfMap = {};
const booklets = fs.readdirSync(formsBase);
booklets.forEach(booklet => {
  const bDir = path.join(formsBase, booklet);
  if (!fs.statSync(bDir).isDirectory()) return;
  const outBooklet = path.join(outForms, booklet);
  if (!fs.existsSync(outBooklet)) fs.mkdirSync(outBooklet);
  fs.readdirSync(bDir).forEach(f => {
    if (!f.toLowerCase().endsWith('.pdf')) return;
    fs.copyFileSync(path.join(bDir, f), path.join(outBooklet, f));
    const key = f.replace(/\.pdf$/i,'').replace(/\s*\(.*\)$/,'').trim();
    const rel = encodeURIComponent(booklet) + '/' + encodeURIComponent(f);
    pdfMap[key] = rel;
  });
});
console.log('PDFs copied:', Object.keys(pdfMap).length);

// 3. Copy TUSA 8112 PDF
if (fs.existsSync(tusaSrc)) {
  const tusaDir = path.join(outForms, 'TUSA');
  if (!fs.existsSync(tusaDir)) fs.mkdirSync(tusaDir);
  fs.copyFileSync(tusaSrc, path.join(tusaDir, '8112_v4 signed.pdf'));
  console.log('TUSA 8112 PDF copied');
}

// 4. Write forms_map.js (consumed by index.html via <script src>)
const mapEntries = Object.entries(pdfMap)
  .map(([k, v]) => `  '${k}':'${v}'`)
  .join(',\n');

const formsMapJs =
  `const PEZA_FORMS_BASE = 'forms/';\n` +
  `const TUSA_8112_PATH = 'forms/TUSA/8112_v4%20signed.pdf';\n` +
  `const PEZA_FORMS_MAP = {\n${mapEntries}\n};\n`;

fs.writeFileSync(path.join(__dirname, 'forms_map.js'), formsMapJs, 'utf8');
console.log('forms_map.js written with', Object.keys(pdfMap).length, 'entries');
