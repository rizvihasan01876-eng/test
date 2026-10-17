/**
 * BUZZBLU HUB — STOCK MANAGEMENT WEB APP
 * Designed to be bound to a NEW Google Spreadsheet converted from the supplied XLSX.
 * Setup is non-destructive: existing sheets and values are not cleared.
 */
const SM = {
  sheets: {
    summary: 'Stock Summery',
    stockIn: 'Stock In',
    outbound: 'Outbond',
    returns: 'Return',
    damage: 'Demage',
    manual: 'Manual Stock Update',
    monthly: 'Month of Octobre Stock Summery Report'
  },
  products: [
    'Baggy pant- Beige- M size',
    'Baggy pant- Beige- L size',
    'Baggy pant- Beige- xl size',
    'Baggy pant- Beige- xxl size',
    'Baggy pant- Nevy Blue- M size',
    'Baggy pant- Nevy Blue- L size',
    'Baggy pant- Nevy Blue- Xl size',
    'Baggy pant- Nevy Blue- xxl size'
  ],
  types: {
    'Stock In': 'Stock In',
    'Outbound': 'Outbond',
    'Return': 'Return',
    'Damage': 'Demage',
    'Manual Stock Update': 'Manual Stock Update'
  }
};

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('BUZZBLU STOCK')
      .addItem('Setup / Repair system', 'setupStockManagement')
      .addItem('Open dashboard', 'showDashboard')
      .addToUi();
  } catch (err) {
    console.log('onOpen skipped: ' + err.message);
  }
}

function setupStockManagement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the target Google Spreadsheet.');
  PropertiesService.getScriptProperties().setProperty('SM_SPREADSHEET_ID', ss.getId());

  // Create missing tabs only; never clear an existing sheet.
  Object.values(SM.sheets).forEach(name => {
    if (!getSheetByTrimmedName_(ss, name)) ss.insertSheet(name);
  });

  // Preserve existing titles and data. Add missing date/product headers where safe.
  [SM.sheets.stockIn, SM.sheets.outbound, SM.sheets.returns, SM.sheets.damage, SM.sheets.manual].forEach(name => {
    const sh = ss.getSheetByName(name);
    const headerRow = findOrPrepareHeaderRow_(sh);
    if (!String(sh.getRange(headerRow, 1).getDisplayValue()).trim()) {
      sh.getRange(headerRow, 1).setValue('Date');
    }
    for (let i = 0; i < SM.products.length; i++) {
      const cell = sh.getRange(headerRow, i + 2);
      if (!String(cell.getDisplayValue()).trim()) cell.setValue(SM.products[i]);
    }
    const noteCol = 10;
    if (!String(sh.getRange(headerRow, noteCol).getDisplayValue()).trim()) sh.getRange(headerRow, noteCol).setValue('NOTE');
    sh.setFrozenRows(headerRow);
    sh.getRange(headerRow, 1, 1, 10).setFontWeight('bold').setWrap(true);
    sh.setColumnWidth(1, 115);
    sh.setColumnWidths(2, 8, 175);
    sh.setColumnWidth(10, 220);
    const dateRange = sh.getRange(headerRow + 1, 1, Math.max(1, sh.getMaxRows() - headerRow), 1);
    dateRange.setNumberFormat('dd-mmm-yyyy');
  });

  setupSummarySheet_(ss.getSheetByName(SM.sheets.summary));
  setupMonthlySheetIfBlank_(getSheetByTrimmedName_(ss, SM.sheets.monthly));
  SpreadsheetApp.flush();
  return { ok: true, message: 'Setup complete. Existing data was preserved.', spreadsheetUrl: ss.getUrl() };
}

function showDashboard() {
  const html = HtmlService.createHtmlOutput(buildDashboardHtml_()).setWidth(1180).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'BUZZBLU Stock Dashboard');
}

function doGet() {
  return HtmlService.createHtmlOutput(buildDashboardHtml_())
    .setTitle('BUZZBLU HUB | Stock Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppData() {
  const ss = getSpreadsheet_();
  const today = new Date();
  return {
    ok: true,
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    products: SM.products,
    types: Object.keys(SM.types),
    today: Utilities.formatDate(today, Session.getScriptTimeZone() || 'Asia/Dhaka', 'yyyy-MM-dd'),
    summary: getSummaryData(),
    recent: getRecentEntries(12)
  };
}

function submitStockEntry(payload) {
  validatePayload_(payload);
  const ss = getSpreadsheet_();
  const sheetName = SM.types[payload.type];
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Required sheet not found: ' + sheetName + '. Run setupStockManagement first.');
  const headerRow = findOrPrepareHeaderRow_(sh);
  const date = parseDate_(payload.date);
  const quantities = SM.products.map((_, i) => {
    const n = Number((payload.quantities || [])[i] || 0);
    if (!Number.isFinite(n) || n < 0) throw new Error('Quantity must be zero or a positive number.');
    return n;
  });
  if (!quantities.some(n => n !== 0) && !String(payload.note || '').trim()) {
    throw new Error('At least one quantity or a note is required.');
  }

  const existingRow = payload.row ? Number(payload.row) : findEntryRow_(sh, headerRow, date);
  const row = [date].concat(quantities, [String(payload.note || '').trim()]);
  if (existingRow && existingRow > headerRow && existingRow <= sh.getMaxRows()) {
    // Update matching date row if row explicitly selected or same date exists.
    sh.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    const appendAt = Math.max(sh.getLastRow() + 1, headerRow + 1);
    sh.getRange(appendAt, 1, 1, row.length).setValues([row]);
  }
  sh.getRange(existingRow || sh.getLastRow(), 1).setNumberFormat('dd-mmm-yyyy');
  SpreadsheetApp.flush();
  return { ok: true, message: existingRow ? 'Entry updated successfully.' : 'Entry saved successfully.', summary: getSummaryData(), recent: getRecentEntries(12) };
}

function getSummaryData(dateString) {
  const ss = getSpreadsheet_();
  const date = dateString ? parseDate_(dateString) : new Date();
  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const end = new Date(start.getTime() + 24 * 60 * 60 * 1000);
  const types = ['Stock In', 'Outbound', 'Return', 'Damage', 'Manual Stock Update'];
  const totals = {};
  types.forEach(type => totals[type] = Array(8).fill(0));
  const allTotals = {};
  types.forEach(type => allTotals[type] = Array(8).fill(0));
  const counts = {};

  types.forEach(type => {
    const sh = ss.getSheetByName(SM.types[type]);
    if (!sh) return;
    const hr = findOrPrepareHeaderRow_(sh);
    const last = sh.getLastRow();
    counts[type] = 0;
    if (last <= hr) return;
    const values = sh.getRange(hr + 1, 1, last - hr, Math.min(10, sh.getMaxColumns())).getValues();
    values.forEach(r => {
      const d = normalizeDate_(r[0]);
      if (!d) return;
      for (let i = 0; i < 8; i++) {
        const v = Number(r[i + 1]) || 0;
        allTotals[type][i] += v;
        if (d >= start && d < end) totals[type][i] += v;
      }
      if (d >= start && d < end) counts[type]++;
    });
  });

  // Base stock is read from the Stock Summery sheet's Stock column where available.
  const summarySh = ss.getSheetByName(SM.sheets.summary);
  const base = Array(8).fill(0);
  if (summarySh) {
    for (let i = 0; i < 8; i++) {
      const label = summarySh.getRange(i + 3, 1).getDisplayValue();
      if (label) {
        const val = summarySh.getRange(i + 3, 2).getValue();
        base[i] = Number(val) || 0;
      }
    }
  }
  const current = SM.products.map((_, i) => Math.max(0,
    base[i] + allTotals['Stock In'][i] + allTotals['Return'][i] + allTotals['Manual Stock Update'][i]
    - allTotals['Outbound'][i] - allTotals['Damage'][i]
  ));
  return {
    date: Utilities.formatDate(start, Session.getScriptTimeZone() || 'Asia/Dhaka', 'yyyy-MM-dd'),
    products: SM.products.map((name, i) => ({
      name, base: base[i], current: current[i],
      stockIn: totals['Stock In'][i], outbound: totals['Outbound'][i],
      returns: totals['Return'][i], damage: totals['Damage'][i], manual: totals['Manual Stock Update'][i]
    })),
    todayTotals: types.reduce((acc, type) => acc + totals[type].reduce((a, b) => a + b, 0), 0),
    currentTotal: current.reduce((a, b) => a + b, 0),
    counts
  };
}

function getRecentEntries(limit) {
  const ss = getSpreadsheet_();
  const out = [];
  Object.keys(SM.types).forEach(type => {
    const sh = ss.getSheetByName(SM.types[type]);
    if (!sh) return;
    const hr = findOrPrepareHeaderRow_(sh);
    const last = sh.getLastRow();
    if (last <= hr) return;
    const start = hr + 1;
    const vals = sh.getRange(start, 1, last - hr, 10).getValues();
    vals.forEach((r, idx) => {
      const d = normalizeDate_(r[0]);
      if (!d) return;
      out.push({ type, sheet: SM.types[type], row: start + idx, date: Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Dhaka', 'yyyy-MM-dd'), total: r.slice(1, 9).reduce((a, v) => a + (Number(v) || 0), 0), note: r[9] || '' });
    });
  });
  out.sort((a, b) => b.date.localeCompare(a.date));
  return out.slice(0, Math.max(Number(limit) || 12, 1));
}

function getEntryForEdit(sheetName, rowNumber) {
  if (!Object.values(SM.types).includes(sheetName)) throw new Error('Invalid sheet.');
  const sh = getSpreadsheet_().getSheetByName(sheetName);
  const hr = findOrPrepareHeaderRow_(sh);
  const row = Number(rowNumber);
  if (!row || row <= hr || row > sh.getLastRow()) throw new Error('Entry row not found.');
  const values = sh.getRange(row, 1, 1, 10).getValues()[0];
  const date = normalizeDate_(values[0]);
  return { sheet: sheetName, type: Object.keys(SM.types).find(k => SM.types[k] === sheetName), row, date: date ? Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Dhaka', 'yyyy-MM-dd') : '', quantities: SM.products.map((_, i) => Number(values[i + 1]) || 0), note: values[9] || '' };
}

function setupSummarySheet_(sh) {
  // Only create a basic summary if the sheet is empty; don't overwrite an existing dashboard.
  if (sh.getLastRow() > 0) return;
  sh.getRange('A1').setValue('Stock Management Dashboard').setFontSize(16).setFontWeight('bold');
  sh.getRange('D1').setValue(new Date()).setNumberFormat('dd mmmm yyyy');
  sh.getRange(2, 1, 1, 9).setValues([['Product', 'Stock', 'Outbound', 'Return', 'Others Out', 'Damage', 'Stock In', 'Transfer To Divine Cloth', 'Auto stock']]);
  sh.getRange(3, 1, 8, 1).setValues(SM.products.map(p => [p]));
  sh.getRange(2, 1, 1, 9).setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(2);
  sh.setColumnWidths(1, 9, 155);
  sh.setColumnWidth(1, 250);
}

function setupMonthlySheetIfBlank_(sh) {
  if (sh.getLastRow() > 0) return;
  sh.getRange(1, 1, 1, 10).setValues([['Date'].concat(SM.products, ['NOTE'])]);
  sh.getRange(1, 1, 1, 10).setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(1);
}

function findOrPrepareHeaderRow_(sh) {
  const maxRows = Math.min(Math.max(sh.getLastRow(), 3), 5);
  const maxCols = Math.min(Math.max(sh.getLastColumn(), 10), 15);
  const rows = sh.getRange(1, 1, maxRows, maxCols).getDisplayValues();
  // Prefer row containing product labels; workbook usually has product headers on row 2.
  for (let r = 0; r < rows.length; r++) {
    const joined = rows[r].join(' ').toLowerCase();
    if (joined.includes('baggy') || joined.includes('colour m') || joined.includes('color m')) return r + 1;
  }
  for (let r = 0; r < rows.length; r++) if (String(rows[r][0]).trim().toLowerCase() === 'date') return r + 1;
  return sh.getLastRow() === 0 ? 1 : Math.min(2, sh.getLastRow() + 1);
}

function findEntryRow_(sh, headerRow, date) {
  const last = sh.getLastRow();
  if (last <= headerRow) return null;
  const target = dateKey_(date);
  const values = sh.getRange(headerRow + 1, 1, last - headerRow, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const d = normalizeDate_(values[i][0]);
    if (d && dateKey_(d) === target) return headerRow + 1 + i;
  }
  return null;
}

function validatePayload_(p) {
  if (!p || !SM.types[p.type]) throw new Error('Please choose a valid entry type.');
  if (!p.date) throw new Error('Please choose a date.');
  if (!Array.isArray(p.quantities) || p.quantities.length < 8) throw new Error('Product quantities are missing.');
}
function parseDate_(s) {
  if (s instanceof Date && !isNaN(s)) return new Date(s.getFullYear(), s.getMonth(), s.getDate());
  const str = String(s || '').trim();
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  const d = new Date(str);
  if (isNaN(d)) throw new Error('Invalid date.');
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function normalizeDate_(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (v === '' || v == null) return null;
  try { return parseDate_(v); } catch (e) { return null; }
}
function dateKey_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Dhaka', 'yyyy-MM-dd'); }
function getSheetByTrimmedName_(ss, wantedName) {
  const wanted = String(wantedName).trim().toLowerCase();
  return ss.getSheets().find(sh => sh.getName().trim().toLowerCase() === wanted) || null;
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SM_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet not connected. Run setupStockManagement from the target spreadsheet first.');
  return ss;
}

function buildDashboardHtml_() {
  return `<!DOCTYPE html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1"><style>
  :root{--bg:#f4f6fa;--ink:#172033;--muted:#697386;--gold:#d6b56d;--dark:#101827;--line:#e5e9f0;--green:#16845b;--red:#c34848}*{box-sizing:border-box}body{margin:0;background:var(--bg);font:14px Arial,sans-serif;color:var(--ink)}.shell{max-width:1250px;margin:auto;padding:22px}.top{background:linear-gradient(120deg,#101827,#273448);color:#fff;border-radius:18px;padding:24px;display:flex;justify-content:space-between;gap:16px;align-items:center}.eyebrow{color:var(--gold);letter-spacing:2px;font-size:11px;font-weight:bold}.top h1{margin:8px 0;font-size:25px}.top p{margin:0;color:#c7cfda}.top a{color:#111827;background:var(--gold);padding:11px 15px;border-radius:9px;text-decoration:none;font-weight:bold}.cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:16px 0}.card,.panel{background:#fff;border:1px solid var(--line);border-radius:14px;padding:17px}.metric{font-size:25px;font-weight:bold;margin-top:9px}.muted{color:var(--muted);font-size:12px}.layout{display:grid;grid-template-columns:1.1fr .9fr;gap:16px}.panel h2{font-size:16px;margin:0 0 15px}.formgrid{display:grid;grid-template-columns:1fr 1fr;gap:12px}.field{display:flex;flex-direction:column;gap:6px}.field label{font-size:12px;color:var(--muted);font-weight:bold}.field input,.field select{border:1px solid #d9dfe8;border-radius:8px;padding:10px;background:#fff;font:inherit;min-width:0}.products{display:grid;grid-template-columns:1fr 100px;gap:7px;margin-top:13px}.prod{padding:9px 0;border-bottom:1px solid #f0f2f5;font-size:12px}.products input{width:100%;padding:8px;border:1px solid #d9dfe8;border-radius:7px}.actions{display:flex;gap:9px;margin-top:16px;flex-wrap:wrap}button{border:0;border-radius:8px;padding:11px 15px;cursor:pointer;font-weight:bold;background:var(--dark);color:white}button.alt{background:#edf0f5;color:#273448}button.gold{background:var(--gold);color:#1a2330}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;font-size:12px}th,td{text-align:left;padding:10px;border-bottom:1px solid var(--line);white-space:nowrap}th{color:var(--muted);font-weight:bold;background:#fafbfc}.pill{padding:5px 8px;background:#f2f4f8;border-radius:6px}.status{margin-top:12px;color:var(--green);min-height:18px}.status.err{color:var(--red)}.full{margin-top:16px}.product-table td:first-child{white-space:normal;min-width:210px}.loading{opacity:.6;pointer-events:none}.smallbtn{padding:6px 9px;font-size:11px;background:#edf0f5;color:#172033}@media(max-width:850px){.layout{grid-template-columns:1fr}.cards{grid-template-columns:repeat(2,1fr)}.shell{padding:12px}.top{align-items:flex-start;flex-direction:column}.formgrid{grid-template-columns:1fr}}
  </style></head><body><div class="shell" id="app"><header class="top"><div><div class="eyebrow">BUZZBLU HUB · INVENTORY CONTROL</div><h1>Stock Management</h1><p>Daily stock entries, live summary and entry management</p></div><a id="sheetLink" href="#" target="_blank">Open Google Sheet ↗</a></header>
  <section class="cards"><div class="card"><div class="muted">CURRENT STOCK</div><div class="metric" id="currentTotal">—</div><div class="muted">All 8 variants</div></div><div class="card"><div class="muted">TODAY'S MOVEMENT</div><div class="metric" id="todayTotal">—</div><div class="muted">Units recorded today</div></div><div class="card"><div class="muted">TODAY'S ENTRIES</div><div class="metric" id="entryCount">—</div><div class="muted">Across entry types</div></div><div class="card"><div class="muted">PRODUCT VARIANTS</div><div class="metric">8</div><div class="muted">Beige & Navy Blue sizes</div></div></section>
  <div class="layout"><section class="panel"><h2 id="formTitle">Add stock entry</h2><div class="formgrid"><div class="field"><label>ENTRY TYPE</label><select id="type"><option>Stock In</option><option>Outbound</option><option>Return</option><option>Damage</option><option>Manual Stock Update</option></select></div><div class="field"><label>DATE</label><input type="date" id="date"></div></div><div class="products" id="products"></div><div class="field" style="margin-top:12px"><label>NOTE (OPTIONAL)</label><input id="note" placeholder="Write a short note"></div><div class="actions"><button class="gold" onclick="saveEntry()" id="saveBtn">Save entry</button><button class="alt" onclick="resetForm()">Clear</button><button class="alt" onclick="refreshAll()">↻ Refresh</button></div><div class="status" id="status"></div></section>
  <section class="panel"><h2>Current stock by variant</h2><div class="tablewrap"><table class="product-table"><thead><tr><th>Product variant</th><th>Base</th><th>Current</th></tr></thead><tbody id="summaryRows"><tr><td colspan="3">Loading…</td></tr></tbody></table></div><p class="muted" style="margin-bottom:0">Current stock = base stock + stock in + return + manual update − outbound − damage. Please verify that the base stock values in Stock Summery are correct.</p></section></div>
  <section class="panel full"><h2>Recent entries</h2><div class="tablewrap"><table><thead><tr><th>Date</th><th>Type</th><th>Total units</th><th>Note</th><th>Action</th></tr></thead><tbody id="recentRows"><tr><td colspan="5">Loading…</td></tr></tbody></table></div></section></div>
  <script>
  let products=[];let editTarget=null;function el(id){return document.getElementById(id)}
  function init(data){products=data.products||[];el('sheetLink').href=data.spreadsheetUrl;el('date').value=data.today;el('products').innerHTML=products.map((p,i)=>'<div class="prod">'+esc(p)+'</div><input type="number" min="0" step="1" id="q'+i+'" value="0" aria-label="'+esc(p)+' quantity">').join('');refreshAll()}
  function refreshAll(){google.script.run.withSuccessHandler(renderData).withFailureHandler(showError).getAppData()}
  function renderData(d){if(!d.ok)return;el('sheetLink').href=d.spreadsheetUrl;const s=d.summary;el('currentTotal').textContent=fmt(s.currentTotal);el('todayTotal').textContent=fmt(s.todayTotals);el('entryCount').textContent=Object.values(s.counts||{}).reduce((a,b)=>a+b,0);el('summaryRows').innerHTML=s.products.map(p=>'<tr><td>'+esc(p.name)+'</td><td>'+fmt(p.base)+'</td><td><b>'+fmt(p.current)+'</b></td></tr>').join('');el('recentRows').innerHTML=(d.recent||[]).map(r=>'<tr><td>'+esc(r.date)+'</td><td><span class="pill">'+esc(r.type)+'</span></td><td>'+fmt(r.total)+'</td><td>'+esc(r.note||'—')+'</td><td><button class="smallbtn" onclick="editEntry('+r.row+',\''+escJs(r.sheet)+'\')">Edit</button></td></tr>').join('')||'<tr><td colspan="5">No entries found.</td></tr>'}
  function saveEntry(){const p={type:el('type').value,date:el('date').value,note:el('note').value,quantities:products.map((_,i)=>Number(el('q'+i).value||0)),row:editTarget?editTarget.row:null};el('saveBtn').disabled=true;el('status').className='status';el('status').textContent='Saving…';google.script.run.withSuccessHandler(r=>{el('saveBtn').disabled=false;el('status').textContent=r.message;editTarget=null;el('formTitle').textContent='Add stock entry';el('saveBtn').textContent='Save entry';refreshAll()}).withFailureHandler(e=>{el('saveBtn').disabled=false;showError(e)}).submitStockEntry(p)}
  function editEntry(row,sheet){el('status').textContent='Loading entry…';google.script.run.withSuccessHandler(d=>{editTarget={row:d.row,sheet:d.sheet};el('type').value=d.type;el('date').value=d.date;el('note').value=d.note||'';d.quantities.forEach((v,i)=>el('q'+i).value=v);el('formTitle').textContent='Edit existing entry · row '+d.row;el('type').disabled=true;el('saveBtn').textContent='Update entry';el('status').textContent='Entry loaded for editing.';window.scrollTo({top:0,behavior:'smooth'})}).withFailureHandler(showError).getEntryForEdit(sheet,row)}
  function resetForm(){editTarget=null;el('type').disabled=false;el('type').value='Stock In';el('date').value=new Date().toISOString().slice(0,10);el('note').value='';products.forEach((_,i)=>el('q'+i).value=0);el('formTitle').textContent='Add stock entry';el('saveBtn').textContent='Save entry';el('status').textContent=''}
  function showError(e){el('status').className='status err';el('status').textContent=(e&&e.message)||String(e)}function fmt(n){return Number(n||0).toLocaleString()}function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}function escJs(s){return String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'")}
  google.script.run.withSuccessHandler(init).withFailureHandler(showError).getAppData();
  </script></body></html>`;
}
