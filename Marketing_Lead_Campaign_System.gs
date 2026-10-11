/**
 * MARKETING CONTENT PERFORMANCE SYSTEM — LEAD + CAMPAIGN
 * Google Sheets + Apps Script
 *
 * Funnel:
 * Impression -> 3s View/Hook -> Click -> Landing Page View -> Lead
 * Final conversion = Lead
 * Sales / Purchase / Revenue are NOT tracked here.
 *
 * Required sheets:
 * Dashboard
 * Project Summary
 * Campaign Summary
 * Content Summary
 * Daily Performance
 * Content Master
 * Raw Import
 */

const CONFIG = {
  SHEETS: {
    DASHBOARD: 'Dashboard',
    PROJECT_SUMMARY: 'Project Summary',
    CAMPAIGN_SUMMARY: 'Campaign Summary',
    CONTENT_SUMMARY: 'Content Summary',
    DAILY: 'Daily Performance',
    MASTER: 'Content Master',
    RAW: 'Raw Import'
  },

  COLORS: {
    DARK: '#0F172A',
    BLUE: '#2563EB',
    LIGHT_BLUE: '#DBEAFE',
    GREEN: '#16A34A',
    LIGHT_GREEN: '#DCFCE7',
    YELLOW: '#D97706',
    LIGHT_YELLOW: '#FEF3C7',
    RED: '#DC2626',
    LIGHT_RED: '#FEE2E2',
    WHITE: '#FFFFFF',
    GRAY: '#64748B',
    LIGHT_GRAY: '#F8FAFC'
  },

  // IMPORTANT: CTA/Clicks must be ACTUAL CLICK COUNT, not CTR percentage.
  RAW_HEADERS: [
    'Date',                 // A
    'Project',              // B
    'Campaign ID',          // C
    'Campaign Name',        // D
    'Content ID',           // E
    'Spend',                // F
    'Impressions',          // G
    'Reach',                // H
    '3s Views',             // I
    '25% Views',            // J
    '50% Views',            // K
    '75% Views',            // L
    '95% Views',            // M
    'CTA/Clicks',           // N
    'Landing Page Views',   // O
    'Leads'                 // P
  ],

  // A:P = imported fields, Q:AA = calculated metrics.
  DAILY_HEADERS: [
    'Date',                 // A
    'Project',              // B
    'Campaign ID',          // C
    'Campaign Name',        // D
    'Content ID',           // E
    'Spend',                // F
    'Impressions',          // G
    'Reach',                // H
    '3s Views',             // I
    '25% Views',            // J
    '50% Views',            // K
    '75% Views',            // L
    '95% Views',            // M
    'CTA/Clicks',           // N
    'Landing Page Views',   // O
    'Leads',                // P
    'Hook Rate',            // Q
    'CTR',                  // R
    'CPC',                  // S
    'CPM',                  // T
    'Cost/CTA',             // U
    'LPV Rate',             // V
    'Cost/LPV',             // W
    'Lead Rate',            // X
    'CPL',                  // Y
    'Frequency',            // Z
    'Performance Score'     // AA
  ],

  BENCHMARKS: {
    HOOK_RATE: 0.40,  // 40%
    CTR: 0.03,        // 3%
    LPV_RATE: 0.80,   // 80%
    LEAD_RATE: 0.12,  // 12%
    TARGET_CPL: 100   // ৳100; lower is better
  }
};


/***************************************************************
 * MENU
 ***************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Marketing Report')
    .addItem('1. Setup / Migrate System', 'setupMarketingSystem')
    .addSeparator()
    .addItem('2. Update Raw Data', 'updateReport')
    .addItem('3. Refresh All Reports', 'refreshAllReports')
    .addSeparator()
    .addItem('4. Refresh Content Summary', 'generateContentSummary')
    .addItem('5. Refresh Campaign Summary', 'generateCampaignSummary')
    .addItem('6. Refresh Project Summary', 'generateProjectSummary')
    .addItem('7. Refresh Dashboard', 'generateDashboard')
    .addSeparator()
    .addItem('8. Clear Raw Import Data', 'clearRawImport')
    .addSeparator()
    .addItem('9. Create Daily Automation', 'createDailyTrigger')
    .addItem('10. Remove Automation', 'removeAllTriggers')
    .addSeparator()
    .addItem('11. Setup Dashboard Controls', 'setupDashboardControls')
    .addToUi();
}


/***************************************************************
 * SETUP / MIGRATION
 * Existing old lead-only Raw/Daily data is remapped by header.
 * Old rows without Campaign ID remain preserved with blank campaign.
 ***************************************************************/
function setupMarketingSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.values(CONFIG.SHEETS).forEach(createSheetIfMissing_);

  migrateTableByHeaders_(CONFIG.SHEETS.RAW, CONFIG.RAW_HEADERS, true);
  migrateTableByHeaders_(CONFIG.SHEETS.DAILY, CONFIG.DAILY_HEADERS.slice(0, 16), true);

  setupRawImportSheet_();
  setupDailyPerformanceSheet_();
  setupContentMasterSheet_();
  setupContentSummarySheet_();
  setupCampaignSummarySheet_();
  setupProjectSummarySheet_();
  setupDashboardSheet_();

  rebuildAllDailyMetrics_();
  generateContentSummary();
  generateCampaignSummary();
  generateProjectSummary();
  generateDashboard();

  SpreadsheetApp.flush();
  ss.toast('Campaign + lead marketing system is ready.', 'Setup Complete', 5);
}

function createSheetIfMissing_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(sheetName)) ss.insertSheet(sheetName);
}

function migrateTableByHeaders_(sheetName, targetHeaders, preserveRows) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  let oldRows = [];
  if (preserveRows && lastRow >= 2 && lastCol >= 1) {
    const oldHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
    const oldData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const index = {};
    oldHeaders.forEach((h, i) => { if (h) index[h.toLowerCase()] = i; });

    oldRows = oldData
      .filter(r => r.some(v => v !== '' && v !== null))
      .map(r => targetHeaders.map(h => {
        const key = h.toLowerCase();
        if (index[key] !== undefined) return r[index[key]];
        // Backward compatibility aliases.
        if (h === 'CTA/Clicks' && index['clicks'] !== undefined) return r[index['clicks']];
        if (h === 'Landing Page Views' && index['lpv'] !== undefined) return r[index['lpv']];
        return '';
      }));
  }

  ensureMinimumColumns_(sheet, targetHeaders.length);
  sheet.clear();
  sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
  if (oldRows.length) sheet.getRange(2, 1, oldRows.length, targetHeaders.length).setValues(oldRows);
}


/***************************************************************
 * SHEET SETUP
 ***************************************************************/
function setupRawImportSheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.RAW);
  ensureMinimumColumns_(sheet, CONFIG.RAW_HEADERS.length);

  sheet.getRange(1, 1, 1, CONFIG.RAW_HEADERS.length).setValues([CONFIG.RAW_HEADERS]);
  formatHeader_(sheet.getRange(1, 1, 1, CONFIG.RAW_HEADERS.length));
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, CONFIG.RAW_HEADERS.length, 125);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 220);
  sheet.getRange('A:A').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('A2').setNote(
    'Required key: Date + Project + Campaign ID + Content ID. CTA/Clicks must be actual click count, NOT CTR %.'
  );
  sheet.getRange('N1').setNote('IMPORTANT: Put actual click count here. Example: 76, not CTR 2.01%.');
}

function setupDailyPerformanceSheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.DAILY);
  ensureMinimumColumns_(sheet, CONFIG.DAILY_HEADERS.length);

  sheet.getRange(1, 1, 1, CONFIG.DAILY_HEADERS.length).setValues([CONFIG.DAILY_HEADERS]);
  formatHeader_(sheet.getRange(1, 1, 1, CONFIG.DAILY_HEADERS.length));
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(5);

  const widths = [110,150,150,220,220];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  for (let c = 6; c <= CONFIG.DAILY_HEADERS.length; c++) sheet.setColumnWidth(c, 120);

  sheet.getRange('A:A').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('F:P').setNumberFormat('#,##0.00');
  sheet.getRange('Q:R').setNumberFormat('0.00%');
  sheet.getRange('S:U').setNumberFormat('৳#,##0.00');
  sheet.getRange('V:V').setNumberFormat('0.00%');
  sheet.getRange('W:W').setNumberFormat('৳#,##0.00');
  sheet.getRange('X:X').setNumberFormat('0.00%');
  sheet.getRange('Y:Y').setNumberFormat('৳#,##0.00');
  sheet.getRange('Z:Z').setNumberFormat('0.00');
  sheet.getRange('AA:AA').setNumberFormat('0.0');
}

function setupContentMasterSheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.MASTER);
  const headers = [
    'Content ID','Project','Content Name','Content Type','Hook Type','Angle/Topic',
    'CTA','Creator/Presenter','Duration (sec)','Publish Date','Status','Notes'
  ];
  ensureMinimumColumns_(sheet, headers.length);
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  formatHeader_(sheet.getRange(1,1,1,headers.length));
  sheet.setFrozenRows(1);
  const widths = [160,150,260,130,150,180,140,170,110,120,120,250];
  widths.forEach((w,i) => sheet.setColumnWidth(i+1,w));
  sheet.getRange('J:J').setNumberFormat('dd-mmm-yyyy');
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active','Testing','Winner','Paused','Stopped'], true)
    .setAllowInvalid(true).build();
  sheet.getRange('K2:K1000').setDataValidation(statusRule);
}

function setupContentSummarySheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.CONTENT_SUMMARY);
  sheet.clear();
  const headers = [
    'Content ID','Project','Content Name','Total Spend','Impressions','3s Views','Clicks','LPV','Leads',
    'Hook Rate','CTR','LPV Rate','Lead Rate','CPC','Cost/LPV','CPL','Frequency','Score','Status'
  ];
  ensureMinimumColumns_(sheet, headers.length);
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  formatHeader_(sheet.getRange(1,1,1,headers.length));
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,220); sheet.setColumnWidth(2,150); sheet.setColumnWidth(3,260);
  for (let c=4;c<=headers.length;c++) sheet.setColumnWidth(c,120);
}

function setupCampaignSummarySheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.CAMPAIGN_SUMMARY);
  sheet.clear();
  const headers = [
    'Campaign ID','Campaign Name','Project','Total Spend','Impressions','Reach','3s Views','Clicks','LPV','Leads',
    'Hook Rate','CTR','LPV Rate','Lead Rate','CPC','Cost/LPV','CPL','Frequency','Score','Status'
  ];
  ensureMinimumColumns_(sheet, headers.length);
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  formatHeader_(sheet.getRange(1,1,1,headers.length));
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,160); sheet.setColumnWidth(2,260); sheet.setColumnWidth(3,150);
  for (let c=4;c<=headers.length;c++) sheet.setColumnWidth(c,120);
}

function setupProjectSummarySheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.PROJECT_SUMMARY);
  sheet.clear();
  const headers = [
    'Project','Spend','Impressions','Reach','3s Views','Clicks','LPV','Leads',
    'Hook Rate','CTR','LPV Rate','Lead Rate','CPC','Cost/LPV','CPL','Frequency','Score','Status'
  ];
  ensureMinimumColumns_(sheet, headers.length);
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  formatHeader_(sheet.getRange(1,1,1,headers.length));
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,180);
  for (let c=2;c<=headers.length;c++) sheet.setColumnWidth(c,120);
}

function setupDashboardSheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.DASHBOARD);
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.getRange('A1:H2').merge();
  sheet.getRange('A1')
    .setValue('MARKETING CONTENT PERFORMANCE DASHBOARD — LEAD + CAMPAIGN')
    .setBackground(CONFIG.COLORS.DARK).setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold').setFontSize(18)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  [170,220,30,130,130,130,130,130].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  createDashboardControlPanel();
}


/***************************************************************
 * UPDATE RAW -> DAILY
 * Raw rows are first GROUPED by Date+Project+Campaign ID+Content ID.
 * This prevents same-key breakdown rows from overwriting each other.
 * Re-importing the same grouped key updates Daily instead of duplicating.
 ***************************************************************/
function updateReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Processing Raw Import data...', 'Marketing Report', 3);
  try {
    const warnings = processRawImport_();
    generateContentSummary();
    generateCampaignSummary();
    generateProjectSummary();
    generateDashboard();
    SpreadsheetApp.flush();
    if (warnings.length) {
      ss.toast('Updated with ' + warnings.length + ' data warning(s). Check Raw Import notes / click values.', 'Complete with warning', 8);
    } else {
      ss.toast('Lead + campaign report updated successfully.', 'Complete', 5);
    }
  } catch (error) {
    console.error(error);
    ss.toast('ERROR: ' + error.message, 'Marketing Report', 8);
  }
}

function processRawImport_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW);
  const dailySheet = getSheet_(CONFIG.SHEETS.DAILY);
  const lastRow = rawSheet.getLastRow();
  if (lastRow < 2) throw new Error('Raw Import sheet has no data.');

  const rawData = rawSheet.getRange(2,1,lastRow-1,CONFIG.RAW_HEADERS.length).getValues();
  const cleanRows = rawData.filter(row => row[0] && row[1] && row[2] && row[4]);
  if (!cleanRows.length) {
    throw new Error('No valid rows. Date, Project, Campaign ID and Content ID are required.');
  }

  const grouped = {};
  const warnings = [];

  cleanRows.forEach((row, index) => {
    const n = normalizeRawRow_(row);
    const key = createUniqueKey_(n[0], n[1], n[2], n[4]);

    if (!grouped[key]) {
      grouped[key] = n.slice();
    } else {
      // Keep Date, Project, Campaign ID, Campaign Name, Content ID.
      if (!grouped[key][3] && n[3]) grouped[key][3] = n[3];
      for (let i = 5; i <= 15; i++) grouped[key][i] += n[i];
    }

    const clicks = n[13];
    const lpv = n[14];
    if (clicks > 0 && !Number.isInteger(clicks)) {
      warnings.push('Row ' + (index + 2) + ': CTA/Clicks is fractional (' + clicks + '). It may be CTR %, not click count.');
    }
    if (clicks > 0 && lpv > clicks * 1.2) {
      warnings.push('Row ' + (index + 2) + ': LPV (' + lpv + ') is much higher than Clicks (' + clicks + '). Check CTA/Clicks field.');
    }
  });

  const existingMap = getExistingDailyMap_();
  Object.keys(grouped).forEach(key => {
    const row = grouped[key];
    if (existingMap[key]) {
      const rowNumber = existingMap[key];
      dailySheet.getRange(rowNumber,1,1,16).setValues([row]);
      writeCalculatedMetrics_(dailySheet,rowNumber);
    } else {
      const newRow = dailySheet.getLastRow()+1;
      dailySheet.getRange(newRow,1,1,16).setValues([row]);
      writeCalculatedMetrics_(dailySheet,newRow);
      existingMap[key] = newRow;
    }
  });

  sortDailySheet_();
  if (warnings.length) console.warn(warnings.join('\n'));
  return warnings;
}

function normalizeRawRow_(row) {
  return [
    normalizeDate_(row[0]),
    String(row[1]).trim(),
    String(row[2]).trim(),
    String(row[3] || '').trim(),
    String(row[4]).trim(),
    number_(row[5]), number_(row[6]), number_(row[7]), number_(row[8]),
    number_(row[9]), number_(row[10]), number_(row[11]), number_(row[12]),
    number_(row[13]), number_(row[14]), number_(row[15])
  ];
}

function getExistingDailyMap_() {
  const sheet = getSheet_(CONFIG.SHEETS.DAILY);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const values = sheet.getRange(2,1,lastRow-1,5).getValues();
  values.forEach((row,index)=>{
    if (!row[0] || !row[1] || !row[2] || !row[4]) return;
    map[createUniqueKey_(row[0],row[1],row[2],row[4])] = index+2;
  });
  return map;
}

function createUniqueKey_(date, project, campaignId, contentId) {
  const formattedDate = Utilities.formatDate(normalizeDate_(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return [formattedDate, project, campaignId, contentId]
    .map(v => String(v).trim().toLowerCase()).join('||');
}


/***************************************************************
 * DAILY CALCULATED METRICS
 ***************************************************************/
function writeCalculatedMetrics_(sheet,row) {
  const formulas = [[
    `=IFERROR(I${row}/G${row},0)`,                 // Q Hook Rate
    `=IFERROR(N${row}/G${row},0)`,                 // R CTR
    `=IFERROR(F${row}/N${row},0)`,                 // S CPC
    `=IFERROR(F${row}/G${row}*1000,0)`,            // T CPM
    `=IFERROR(F${row}/N${row},0)`,                 // U Cost/CTA
    `=IFERROR(O${row}/N${row},0)`,                 // V LPV Rate
    `=IFERROR(F${row}/O${row},0)`,                 // W Cost/LPV
    `=IFERROR(P${row}/O${row},0)`,                 // X Lead Rate
    `=IFERROR(F${row}/P${row},0)`,                 // Y CPL
    `=IFERROR(G${row}/H${row},0)`,                 // Z Frequency
    createPerformanceScoreFormula_(row)            // AA Score
  ]];
  sheet.getRange(row,17,1,11).setFormulas(formulas);
  formatDailyMetricsRow_(sheet,row);
}

function rebuildAllDailyMetrics_() {
  const sheet = getSheet_(CONFIG.SHEETS.DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  for (let r=2;r<=lastRow;r++) {
    if (sheet.getRange(r,1).getValue() && sheet.getRange(r,5).getValue()) writeCalculatedMetrics_(sheet,r);
  }
}

function createPerformanceScoreFormula_(row) {
  const b = CONFIG.BENCHMARKS;
  // Existing weighting preserved: 25 Hook + 25 CTR + 20 LPV + 20 Lead Rate + 10 CPL = 100.
  return `=MIN(100,MIN(Q${row}/${b.HOOK_RATE},1)*25+MIN(R${row}/${b.CTR},1)*25+MIN(V${row}/${b.LPV_RATE},1)*20+MIN(X${row}/${b.LEAD_RATE},1)*20+IF(Y${row}=0,0,MIN(${b.TARGET_CPL}/Y${row},1)*10))`;
}

function formatDailyMetricsRow_(sheet,row) {
  sheet.getRange(row,17,1,2).setNumberFormat('0.00%');   // Q:R
  sheet.getRange(row,19,1,3).setNumberFormat('৳#,##0.00'); // S:U
  sheet.getRange(row,22).setNumberFormat('0.00%');      // V
  sheet.getRange(row,23).setNumberFormat('৳#,##0.00');  // W
  sheet.getRange(row,24).setNumberFormat('0.00%');      // X
  sheet.getRange(row,25).setNumberFormat('৳#,##0.00');  // Y
  sheet.getRange(row,26).setNumberFormat('0.00');       // Z
  sheet.getRange(row,27).setNumberFormat('0.0');        // AA
}

function sortDailySheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  sheet.getRange(2,1,lastRow-1,CONFIG.DAILY_HEADERS.length).sort([
    {column:1,ascending:true}, {column:2,ascending:true},
    {column:3,ascending:true}, {column:5,ascending:true}
  ]);
}


/***************************************************************
 * CONTENT SUMMARY — combines a content across all campaigns
 * within the same project.
 ***************************************************************/
function generateContentSummary() {
  const daily = getSheet_(CONFIG.SHEETS.DAILY);
  const master = getSheet_(CONFIG.SHEETS.MASTER);
  const out = getSheet_(CONFIG.SHEETS.CONTENT_SUMMARY);
  setupContentSummarySheet_();

  const lastRow = daily.getLastRow();
  if (lastRow < 2) return;
  const data = daily.getRange(2,1,lastRow-1,27).getValues();
  const masterMap = getContentMasterMap_(master);
  const summary = {};

  data.forEach(row => {
    const project = String(row[1]).trim();
    const contentId = String(row[4]).trim();
    if (!project || !contentId) return;
    const key = project.toLowerCase() + '||' + contentId.toLowerCase();
    if (!summary[key]) summary[key] = baseMetricObject_({project,contentId});
    addMetrics_(summary[key],row);
  });

  const rows = Object.values(summary).map(item => {
    const m = deriveMetrics_(item);
    const masterItem = masterMap[(item.project+'||'+item.contentId).toLowerCase()] || masterMap[item.contentId.toLowerCase()] || {};
    const score = calculateScore_(m.hookRate,m.ctr,m.lpvRate,m.leadRate,m.cpl);
    return [
      item.contentId,item.project,masterItem.contentName||'',item.spend,item.impressions,item.views3s,
      item.clicks,item.lpv,item.leads,m.hookRate,m.ctr,m.lpvRate,m.leadRate,m.cpc,m.costLpv,m.cpl,m.frequency,
      score,getPerformanceStatus_(score,item.leads)
    ];
  });

  rows.sort((a,b)=>b[17]-a[17]);
  if (rows.length) {
    out.getRange(2,1,rows.length,19).setValues(rows);
    formatLeadSummary_(out,rows.length,10,14,17,18,19);
  }
}

function getContentMasterMap_(sheet) {
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  sheet.getRange(2,1,lastRow-1,12).getValues().forEach(row=>{
    const id=String(row[0]).trim(); const project=String(row[1]).trim();
    if(!id) return;
    const obj={project,contentName:row[2],contentType:row[3],hookType:row[4],angle:row[5],cta:row[6],creator:row[7],duration:row[8],publishDate:row[9],status:row[10],notes:row[11]};
    map[id.toLowerCase()]=obj;
    if(project) map[(project+'||'+id).toLowerCase()]=obj;
  });
  return map;
}


/***************************************************************
 * CAMPAIGN SUMMARY
 ***************************************************************/
function generateCampaignSummary() {
  const daily = getSheet_(CONFIG.SHEETS.DAILY);
  const out = getSheet_(CONFIG.SHEETS.CAMPAIGN_SUMMARY);
  setupCampaignSummarySheet_();

  const lastRow = daily.getLastRow();
  if (lastRow < 2) return;
  const data = daily.getRange(2,1,lastRow-1,27).getValues();
  const summary = {};

  data.forEach(row=>{
    const project=String(row[1]).trim();
    const campaignId=String(row[2]).trim();
    const campaignName=String(row[3]).trim();
    if(!project || !campaignId) return;
    const key=(project+'||'+campaignId).toLowerCase();
    if(!summary[key]) summary[key]=baseMetricObject_({project,campaignId,campaignName});
    if(!summary[key].campaignName && campaignName) summary[key].campaignName=campaignName;
    addMetrics_(summary[key],row);
  });

  const rows=Object.values(summary).map(item=>{
    const m=deriveMetrics_(item);
    const score=calculateScore_(m.hookRate,m.ctr,m.lpvRate,m.leadRate,m.cpl);
    return [
      item.campaignId,item.campaignName||'',item.project,item.spend,item.impressions,item.reach,item.views3s,
      item.clicks,item.lpv,item.leads,m.hookRate,m.ctr,m.lpvRate,m.leadRate,m.cpc,m.costLpv,m.cpl,m.frequency,
      score,getPerformanceStatus_(score,item.leads)
    ];
  });

  rows.sort((a,b)=>b[18]-a[18]);
  if(rows.length){
    out.getRange(2,1,rows.length,20).setValues(rows);
    out.getRange(2,4,rows.length,7).setNumberFormat('#,##0.00');
    out.getRange(2,11,rows.length,4).setNumberFormat('0.00%');
    out.getRange(2,15,rows.length,3).setNumberFormat('৳#,##0.00');
    out.getRange(2,18,rows.length,1).setNumberFormat('0.00');
    out.getRange(2,19,rows.length,1).setNumberFormat('0.0');
    applyStatusColors_(out,rows.length,20);
  }
}


/***************************************************************
 * PROJECT SUMMARY
 ***************************************************************/
function generateProjectSummary() {
  const daily=getSheet_(CONFIG.SHEETS.DAILY);
  const out=getSheet_(CONFIG.SHEETS.PROJECT_SUMMARY);
  setupProjectSummarySheet_();
  const lastRow=daily.getLastRow();
  if(lastRow<2) return;
  const data=daily.getRange(2,1,lastRow-1,27).getValues();
  const summary={};

  data.forEach(row=>{
    const project=String(row[1]).trim();
    if(!project) return;
    const key=project.toLowerCase();
    if(!summary[key]) summary[key]=baseMetricObject_({project});
    addMetrics_(summary[key],row);
  });

  const rows=Object.values(summary).map(item=>{
    const m=deriveMetrics_(item);
    const score=calculateScore_(m.hookRate,m.ctr,m.lpvRate,m.leadRate,m.cpl);
    return [
      item.project,item.spend,item.impressions,item.reach,item.views3s,item.clicks,item.lpv,item.leads,
      m.hookRate,m.ctr,m.lpvRate,m.leadRate,m.cpc,m.costLpv,m.cpl,m.frequency,score,getPerformanceStatus_(score,item.leads)
    ];
  });

  rows.sort((a,b)=>b[16]-a[16]);
  if(rows.length){
    out.getRange(2,1,rows.length,18).setValues(rows);
    out.getRange(2,2,rows.length,7).setNumberFormat('#,##0.00');
    out.getRange(2,9,rows.length,4).setNumberFormat('0.00%');
    out.getRange(2,13,rows.length,3).setNumberFormat('৳#,##0.00');
    out.getRange(2,16,rows.length,1).setNumberFormat('0.00');
    out.getRange(2,17,rows.length,1).setNumberFormat('0.0');
    applyStatusColors_(out,rows.length,18);
  }
}

function baseMetricObject_(extra) {
  return Object.assign({spend:0,impressions:0,reach:0,views3s:0,clicks:0,lpv:0,leads:0},extra||{});
}

function addMetrics_(item,row) {
  item.spend += number_(row[5]);
  item.impressions += number_(row[6]);
  item.reach += number_(row[7]);
  item.views3s += number_(row[8]);
  item.clicks += number_(row[13]);
  item.lpv += number_(row[14]);
  item.leads += number_(row[15]);
}

function deriveMetrics_(item) {
  return {
    hookRate:safeDivide_(item.views3s,item.impressions),
    ctr:safeDivide_(item.clicks,item.impressions),
    lpvRate:safeDivide_(item.lpv,item.clicks),
    leadRate:safeDivide_(item.leads,item.lpv),
    cpc:safeDivide_(item.spend,item.clicks),
    costLpv:safeDivide_(item.spend,item.lpv),
    cpl:safeDivide_(item.spend,item.leads),
    frequency:safeDivide_(item.impressions,item.reach)
  };
}

function formatLeadSummary_(sheet,rowCount,percentStart,costStart,frequencyCol,scoreCol,statusCol) {
  sheet.getRange(2,4,rowCount,6).setNumberFormat('#,##0.00');
  sheet.getRange(2,percentStart,rowCount,4).setNumberFormat('0.00%');
  sheet.getRange(2,costStart,rowCount,3).setNumberFormat('৳#,##0.00');
  sheet.getRange(2,frequencyCol,rowCount,1).setNumberFormat('0.00');
  sheet.getRange(2,scoreCol,rowCount,1).setNumberFormat('0.0');
  applyStatusColors_(sheet,rowCount,statusCol);
}


/***************************************************************
 * DASHBOARD
 * Project filter = B4
 * Campaign filter = B5
 ***************************************************************/
function generateDashboard() {
  const dashboard=getSheet_(CONFIG.SHEETS.DASHBOARD);
  const oldProject=dashboard.getRange('B4').getValue() || 'All Projects';
  const oldCampaign=dashboard.getRange('B5').getValue() || 'All Campaigns';

  setupDashboardSheet_();
  createProjectDropdown_();
  dashboard.getRange('B4').setValue(oldProject);
  createCampaignDropdown_(oldProject);

  const campaignOptions=getValidationValues_(dashboard.getRange('B5'));
  dashboard.getRange('B5').setValue(campaignOptions.includes(oldCampaign)?oldCampaign:'All Campaigns');
  const selectedProject=dashboard.getRange('B4').getValue() || 'All Projects';
  const selectedCampaign=dashboard.getRange('B5').getValue() || 'All Campaigns';

  const data=getDashboardData_(selectedProject,selectedCampaign);
  writeDashboardKPIs_(dashboard,data);
  writeDashboardFunnel_(dashboard,data);
  writeDashboardContentTable_(dashboard,selectedProject,selectedCampaign);
  createDashboardCharts_(dashboard);
}

function createProjectDropdown_() {
  const dashboard=getSheet_(CONFIG.SHEETS.DASHBOARD);
  const daily=getSheet_(CONFIG.SHEETS.DAILY);
  let projects=['All Projects'];
  const lastRow=daily.getLastRow();
  if(lastRow>=2){
    const vals=daily.getRange(2,2,lastRow-1,1).getValues().flat().map(v=>String(v).trim()).filter(Boolean);
    projects=projects.concat([...new Set(vals)].sort());
  }
  dashboard.getRange('B4').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(projects,true).setAllowInvalid(false).build()
  );
}

function createCampaignDropdown_(selectedProject) {
  const dashboard=getSheet_(CONFIG.SHEETS.DASHBOARD);
  const daily=getSheet_(CONFIG.SHEETS.DAILY);
  let options=['All Campaigns'];
  const lastRow=daily.getLastRow();
  if(lastRow>=2){
    const vals=daily.getRange(2,2,lastRow-1,3).getValues(); // Project, Campaign ID, Name
    const labels=[];
    vals.forEach(r=>{
      const project=String(r[0]).trim(); const id=String(r[1]).trim(); const name=String(r[2]).trim();
      if(!id) return;
      if(selectedProject!=='All Projects' && project!==selectedProject) return;
      labels.push(campaignLabel_(id,name));
    });
    options=options.concat([...new Set(labels)].sort());
  }
  dashboard.getRange('B5').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(options,true).setAllowInvalid(false).build()
  );
  if(!options.includes(String(dashboard.getRange('B5').getValue()))) dashboard.getRange('B5').setValue('All Campaigns');
}

function campaignLabel_(id,name){ return name ? id+' | '+name : id; }
function campaignIdFromLabel_(label){ return String(label).split(' | ')[0].trim(); }

function getDashboardData_(selectedProject,selectedCampaign) {
  const daily=getSheet_(CONFIG.SHEETS.DAILY);
  const result=baseMetricObject_({});
  const lastRow=daily.getLastRow();
  if(lastRow<2) return result;
  const campaignId=selectedCampaign==='All Campaigns'?'':campaignIdFromLabel_(selectedCampaign);
  daily.getRange(2,1,lastRow-1,16).getValues().forEach(row=>{
    const project=String(row[1]).trim(); const cid=String(row[2]).trim();
    if(selectedProject!=='All Projects' && project!==selectedProject) return;
    if(campaignId && cid!==campaignId) return;
    addMetrics_(result,row);
  });
  return result;
}

function writeDashboardKPIs_(sheet,data) {
  const m=deriveMetrics_(data);
  const cards=[
    ['A8','Total Spend',data.spend,'৳#,##0.00'], ['C8','Impressions',data.impressions,'#,##0'],
    ['E8','Clicks',data.clicks,'#,##0'], ['G8','CTR',m.ctr,'0.00%'],
    ['A11','Leads',data.leads,'#,##0'], ['C11','CPL',m.cpl,'৳#,##0.00'],
    ['E11','Hook Rate',m.hookRate,'0.00%'], ['G11','LPV Rate',m.lpvRate,'0.00%'],
    ['A14','Lead Rate',m.leadRate,'0.00%'], ['C14','LPV',data.lpv,'#,##0'],
    ['E14','3s Views',data.views3s,'#,##0'], ['G14','Frequency',m.frequency,'0.00']
  ];
  cards.forEach(card=>{
    const label=sheet.getRange(card[0]); const row=label.getRow(), col=label.getColumn();
    label.setValue(card[1]).setBackground(CONFIG.COLORS.LIGHT_BLUE).setFontColor(CONFIG.COLORS.DARK)
      .setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(row+1,col).setValue(card[2]).setNumberFormat(card[3]).setFontWeight('bold')
      .setFontSize(15).setHorizontalAlignment('center').setBackground(CONFIG.COLORS.WHITE);
  });
}

function writeDashboardFunnel_(sheet,data) {
  const m=deriveMetrics_(data);
  sheet.getRange('A17:F17').merge().setValue('MARKETING FUNNEL — FINAL RESULT = LEAD')
    .setBackground(CONFIG.COLORS.BLUE).setFontColor(CONFIG.COLORS.WHITE).setFontWeight('bold');
  const headers=['Stage','Count','Conversion Rate','Cost / Stage','Diagnosis','Suggested Benchmark'];
  sheet.getRange('A18:F18').setValues([headers]); formatHeader_(sheet.getRange('A18:F18'));
  const rows=[
    ['Impressions',data.impressions,'',safeDivide_(data.spend*1000,data.impressions),'Ad delivery / audience cost','Track CPM'],
    ['3s Views',data.views3s,m.hookRate,safeDivide_(data.spend,data.views3s),'Opening hook quality','Hook Rate 30-40%+'],
    ['Clicks',data.clicks,m.ctr,safeDivide_(data.spend,data.clicks),'Creative + CTA strength','CTR 2%+'],
    ['Landing Page Views',data.lpv,m.lpvRate,safeDivide_(data.spend,data.lpv),'Click quality + page load','LPV Rate 70-80%+'],
    ['Leads',data.leads,m.leadRate,m.cpl,'Final marketing conversion','Track CPL + Lead Rate']
  ];
  sheet.getRange(19,1,rows.length,6).setValues(rows);
  sheet.getRange(19,3,rows.length,1).setNumberFormat('0.00%');
  sheet.getRange(19,4,rows.length,1).setNumberFormat('৳#,##0.00');
}

function writeDashboardContentTable_(dashboard,selectedProject,selectedCampaign) {
  const daily=getSheet_(CONFIG.SHEETS.DAILY);
  dashboard.getRange('A27:I27').merge().setValue('TOP CONTENT PERFORMANCE — CURRENT FILTER')
    .setBackground(CONFIG.COLORS.BLUE).setFontColor(CONFIG.COLORS.WHITE).setFontWeight('bold');
  const headers=['Content ID','Content Name','Spend','Hook Rate','CTR','LPV Rate','Leads','CPL','Score'];
  dashboard.getRange('A28:I28').setValues([headers]); formatHeader_(dashboard.getRange('A28:I28'));

  const lastRow=daily.getLastRow(); if(lastRow<2) return;
  const masterMap=getContentMasterMap_(getSheet_(CONFIG.SHEETS.MASTER));
  const campaignId=selectedCampaign==='All Campaigns'?'':campaignIdFromLabel_(selectedCampaign);
  const sum={};
  daily.getRange(2,1,lastRow-1,27).getValues().forEach(row=>{
    const project=String(row[1]).trim(), cid=String(row[2]).trim(), contentId=String(row[4]).trim();
    if(selectedProject!=='All Projects' && project!==selectedProject) return;
    if(campaignId && cid!==campaignId) return;
    if(!contentId) return;
    const key=(project+'||'+contentId).toLowerCase();
    if(!sum[key]) sum[key]=baseMetricObject_({project,contentId});
    addMetrics_(sum[key],row);
  });

  const rows=Object.values(sum).map(item=>{
    const m=deriveMetrics_(item); const score=calculateScore_(m.hookRate,m.ctr,m.lpvRate,m.leadRate,m.cpl);
    const master=masterMap[(item.project+'||'+item.contentId).toLowerCase()]||masterMap[item.contentId.toLowerCase()]||{};
    return [item.contentId,master.contentName||'',item.spend,m.hookRate,m.ctr,m.lpvRate,item.leads,m.cpl,score];
  }).sort((a,b)=>b[8]-a[8]).slice(0,10);

  if(rows.length){
    dashboard.getRange(29,1,rows.length,9).setValues(rows);
    dashboard.getRange(29,3,rows.length,1).setNumberFormat('৳#,##0.00');
    dashboard.getRange(29,4,rows.length,3).setNumberFormat('0.00%');
    dashboard.getRange(29,8,rows.length,1).setNumberFormat('৳#,##0.00');
    dashboard.getRange(29,9,rows.length,1).setNumberFormat('0.0');
  }
}

function createDashboardCharts_(dashboard) {
  dashboard.getCharts().forEach(c=>dashboard.removeChart(c));
  const funnel=dashboard.newChart().setChartType(Charts.ChartType.BAR)
    .addRange(dashboard.getRange('A18:B23')).setPosition(17,11,0,0)
    .setOption('title','Lead Funnel').setOption('legend',{position:'none'}).build();
  dashboard.insertChart(funnel);
  const content=dashboard.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(dashboard.getRange('A28:A38')).addRange(dashboard.getRange('G28:G38'))
    .setPosition(29,11,0,0).setOption('title','Top Content by Leads')
    .setOption('legend',{position:'none'}).build();
  dashboard.insertChart(content);
}


/***************************************************************
 * REFRESH / CLEAR / AUTOMATION
 ***************************************************************/
function refreshAllReports() {
  generateContentSummary(); generateCampaignSummary(); generateProjectSummary(); generateDashboard();
  SpreadsheetApp.getActiveSpreadsheet().toast('All reports refreshed.','Marketing Report',4);
}

function clearRawImport() {
  const ui=SpreadsheetApp.getUi();
  const result=ui.alert('Clear Raw Import?','This deletes Raw Import rows only. Daily Performance history stays safe.',ui.ButtonSet.YES_NO);
  if(result!==ui.Button.YES) return;
  clearRawImportFromControl_();
}

function createDailyTrigger() {
  removeSpecificTrigger_('refreshAllReports');
  ScriptApp.newTrigger('refreshAllReports').timeBased().everyDays(1).atHour(23).create();
  SpreadsheetApp.getUi().alert('Daily report refresh automation created for around 11 PM.');
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('All Apps Script triggers have been removed.');
}

function removeSpecificTrigger_(functionName) {
  ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()===functionName) ScriptApp.deleteTrigger(t); });
}


/***************************************************************
 * PERFORMANCE SCORE
 ***************************************************************/
function calculateScore_(hookRate,ctr,lpvRate,leadRate,cpl) {
  const b=CONFIG.BENCHMARKS;
  const hookScore=Math.min(hookRate/b.HOOK_RATE,1)*25;
  const ctrScore=Math.min(ctr/b.CTR,1)*25;
  const lpvScore=Math.min(lpvRate/b.LPV_RATE,1)*20;
  const leadScore=Math.min(leadRate/b.LEAD_RATE,1)*20;
  const cplScore=cpl>0?Math.min(b.TARGET_CPL/cpl,1)*10:0;
  return Number((hookScore+ctrScore+lpvScore+leadScore+cplScore).toFixed(2));
}

function getPerformanceStatus_(score,leads) {
  if(leads<=0) return '🔴 No Lead';
  if(score>=80) return '🔥 Winner';
  if(score>=65) return '🟢 Good';
  if(score>=45) return '🟡 Average';
  return '🔴 Poor';
}

function applyStatusColors_(sheet,rowCount,statusCol) {
  if(rowCount<=0) return;
  const range=sheet.getRange(2,statusCol,rowCount,1);
  const rules=[
    ['Winner',CONFIG.COLORS.LIGHT_GREEN,'#166534'],
    ['Good',CONFIG.COLORS.LIGHT_BLUE,'#1E40AF'],
    ['Average',CONFIG.COLORS.LIGHT_YELLOW,'#92400E'],
    ['Poor',CONFIG.COLORS.LIGHT_RED,'#991B1B'],
    ['No Lead',CONFIG.COLORS.LIGHT_RED,'#991B1B']
  ].map(x=>SpreadsheetApp.newConditionalFormatRule().whenTextContains(x[0]).setBackground(x[1]).setFontColor(x[2]).setRanges([range]).build());
  sheet.setConditionalFormatRules(rules);
}


/***************************************************************
 * DASHBOARD CONTROL PANEL
 ***************************************************************/
const CONTROL_PANEL={
  SHEET:'Dashboard',
  BUTTONS:{
    D5:{name:'UPDATE REPORT',action:'updateReport'},
    E5:{name:'REFRESH ALL',action:'refreshAllReports'},
    F5:{name:'DASHBOARD',action:'generateDashboard'},
    G5:{name:'CLEAR RAW',action:'clearRawImportFromControl_'},
    H5:{name:'AUTOMATION',action:'createDailyTriggerFromControl_'}
  }
};

function createDashboardControlPanel() {
  const sheet=getSheet_(CONTROL_PANEL.SHEET);
  sheet.getRange('A4').setValue('SELECT PROJECT').setBackground(CONFIG.COLORS.BLUE).setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B4').setValue(sheet.getRange('B4').getValue()||'All Projects').setBackground('#EFF6FF').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('A5').setValue('SELECT CAMPAIGN').setBackground(CONFIG.COLORS.BLUE).setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B5').setValue(sheet.getRange('B5').getValue()||'All Campaigns').setBackground('#EFF6FF').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('A6').setValue('LAST UPDATED').setBackground('#F1F5F9').setFontColor(CONFIG.COLORS.GRAY).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B6').setBackground(CONFIG.COLORS.WHITE).setHorizontalAlignment('center').setNumberFormat('dd-mmm-yyyy hh:mm:ss AM/PM');

  sheet.getRange('D4:H4').setValues([['UPDATE REPORT','REFRESH ALL','DASHBOARD','CLEAR RAW','AUTOMATION']])
    .setBackground(CONFIG.COLORS.DARK).setFontColor(CONFIG.COLORS.WHITE).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.getRange('D5:H5').insertCheckboxes().uncheck().setHorizontalAlignment('center');
  sheet.getRange('D5').setBackground(CONFIG.COLORS.LIGHT_GREEN);
  sheet.getRange('E5').setBackground(CONFIG.COLORS.LIGHT_BLUE);
  sheet.getRange('F5').setBackground('#E0F2FE');
  sheet.getRange('G5').setBackground(CONFIG.COLORS.LIGHT_RED);
  sheet.getRange('H5').setBackground(CONFIG.COLORS.LIGHT_YELLOW);

  [4,5,6].forEach(r=>sheet.setRowHeight(r,35));
  sheet.getRange('D5').setNote('Tick to import Raw Import and refresh all reports.');
  sheet.getRange('E5').setNote('Tick to refresh summaries and dashboard from Daily Performance.');
  sheet.getRange('F5').setNote('Tick to refresh Dashboard only.');
  sheet.getRange('G5').setNote('Tick to clear Raw Import only. Daily history stays safe.');
  sheet.getRange('H5').setNote('Tick to create daily refresh automation around 11 PM.');
}

function setupDashboardControls() {
  createDashboardControlPanel(); createProjectDropdown_();
  createCampaignDropdown_(getSheet_(CONFIG.SHEETS.DASHBOARD).getRange('B4').getValue()||'All Projects');
  installControlPanelTrigger(); updateLastUpdated_(); SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard controls activated.','Setup Complete',5);
}

function installControlPanelTrigger() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t=>{if(t.getHandlerFunction()==='controlPanelOnEdit') ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('controlPanelOnEdit').forSpreadsheet(ss).onEdit().create();
  ss.toast('Control Panel trigger installed.','Setup Complete',5);
}

function controlPanelOnEdit(e) {
  if(!e||!e.range) return;
  const range=e.range, sheet=range.getSheet();
  if(sheet.getName()!==CONTROL_PANEL.SHEET) return;
  const a1=range.getA1Notation();

  if(a1==='B4') { handleProjectChange_(); return; }
  if(a1==='B5') { handleCampaignChange_(); return; }

  const btn=CONTROL_PANEL.BUTTONS[a1];
  if(!btn||e.value!=='TRUE') return;
  const lock=LockService.getDocumentLock();
  try{
    lock.waitLock(30000); setControlStatus_('RUNNING: '+btn.name); runControlAction_(btn.action);
    updateLastUpdated_(); setControlStatus_('DONE: '+btn.name);
  }catch(error){ console.error(error); setControlStatus_('ERROR: '+error.message); }
  finally{ range.uncheck(); try{lock.releaseLock();}catch(err){console.error(err);} }
}

function runControlAction_(actionName) {
  switch(actionName){
    case 'updateReport': updateReport(); break;
    case 'refreshAllReports': refreshAllReports(); break;
    case 'generateDashboard': generateDashboard(); break;
    case 'clearRawImportFromControl_': clearRawImportFromControl_(); break;
    case 'createDailyTriggerFromControl_': createDailyTriggerFromControl_(); break;
    default: throw new Error('Unknown Control Panel action: '+actionName);
  }
}

function clearRawImportFromControl_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=getSheet_(CONFIG.SHEETS.RAW), lastRow=sheet.getLastRow();
  if(lastRow<=1){ss.toast('Raw Import is already empty.','Marketing Report',3);return;}
  sheet.getRange(2,1,lastRow-1,CONFIG.RAW_HEADERS.length).clearContent();
  ss.toast('Raw Import cleared. Daily history is safe.','Marketing Report',4);
}

function createDailyTriggerFromControl_() {
  removeSpecificTrigger_('refreshAllReports');
  ScriptApp.newTrigger('refreshAllReports').timeBased().everyDays(1).atHour(23).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Daily refresh automation enabled.','Automation',4);
}

function handleProjectChange_() {
  const sheet=getSheet_(CONFIG.SHEETS.DASHBOARD);
  const project=sheet.getRange('B4').getValue()||'All Projects';
  createCampaignDropdown_(project); sheet.getRange('B5').setValue('All Campaigns');
  setControlStatus_('Updating project dashboard...'); generateDashboard(); updateLastUpdated_(); setControlStatus_('Project dashboard updated');
}

function handleCampaignChange_() {
  setControlStatus_('Updating campaign dashboard...'); generateDashboard(); updateLastUpdated_(); setControlStatus_('Campaign dashboard updated');
}

function updateLastUpdated_() {
  getSheet_(CONTROL_PANEL.SHEET).getRange('B6').setValue(new Date()).setNumberFormat('dd-mmm-yyyy hh:mm:ss AM/PM');
}

function setControlStatus_(message) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName(CONTROL_PANEL.SHEET);
  if(sheet) sheet.getRange('A1').setNote(message);
  ss.toast(message,'Marketing Control Panel',3);
}


/***************************************************************
 * UTILITIES
 ***************************************************************/
function formatHeader_(range) {
  range.setBackground(CONFIG.COLORS.DARK).setFontColor(CONFIG.COLORS.WHITE).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

function getSheet_(sheetName) {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sheet) throw new Error(`Sheet "${sheetName}" not found. Run setupMarketingSystem first.`);
  return sheet;
}

function ensureMinimumColumns_(sheet,neededColumns) {
  const current=sheet.getMaxColumns();
  if(current<neededColumns) sheet.insertColumnsAfter(current,neededColumns-current);
}

function number_(value) {
  if(value===''||value===null||value===undefined) return 0;
  if(typeof value==='number') return value;
  const text=String(value).replace(/,/g,'').replace(/৳/g,'').replace(/\$/g,'').replace(/%/g,'').trim();
  const num=parseFloat(text); return isNaN(num)?0:num;
}

function safeDivide_(numerator,denominator) {
  numerator=number_(numerator); denominator=number_(denominator);
  return denominator===0?0:numerator/denominator;
}

function normalizeDate_(value) {
  if(value instanceof Date && !isNaN(value)) return value;
  if(typeof value==='number') return new Date(Math.round((value-25569)*86400*1000));
  if(typeof value==='string'){
    const cleaned=value.trim();
    const parsedDirect=new Date(cleaned);
    if(!isNaN(parsedDirect)) return parsedDirect;
    const parts=cleaned.split(/[-/.]/);
    if(parts.length===3){
      let day=parseInt(parts[0],10), month=parseInt(parts[1],10), year=parseInt(parts[2],10);
      if(parts[0].length===4){year=parseInt(parts[0],10);month=parseInt(parts[1],10);day=parseInt(parts[2],10);}
      if(year<100) year+=2000;
      const d=new Date(year,month-1,day); if(!isNaN(d)) return d;
    }
  }
  throw new Error('Invalid date detected: '+value);
}

function getValidationValues_(range) {
  const rule=range.getDataValidation();
  if(!rule) return [];
  const criteria=rule.getCriteriaValues();
  return Array.isArray(criteria[0])?criteria[0]:[];
}
