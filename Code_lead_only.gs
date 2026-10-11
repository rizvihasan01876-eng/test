/**
 * MARKETING CONTENT PERFORMANCE SYSTEM — LEAD ONLY
 * Google Sheets + Apps Script
 *
 * Funnel:
 * Impression -> 3s View/Hook -> Click -> Landing Page View -> Lead
 *
 * Final conversion = Lead
 * Sales / Purchase / Revenue are intentionally NOT tracked here.
 *
 * Required sheets:
 * Dashboard
 * Project Summary
 * Content Summary
 * Daily Performance
 * Content Master
 * Raw Import
 */

const CONFIG = {
  SHEETS: {
    DASHBOARD: 'Dashboard',
    PROJECT_SUMMARY: 'Project Summary',
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

  // User-entered/imported fields only.
  RAW_HEADERS: [
    'Date',
    'Project',
    'Content ID',
    'Spend',
    'Impressions',
    'Reach',
    '3s Views',
    '25% Views',
    '50% Views',
    '75% Views',
    '95% Views',
    'CTA/Clicks',
    'Landing Page Views',
    'Leads'
  ],

  // A:N = input data, O:X = calculated metrics.
  DAILY_HEADERS: [
    'Date',                 // A
    'Project',              // B
    'Content ID',           // C
    'Spend',                // D
    'Impressions',          // E
    'Reach',                // F
    '3s Views',             // G
    '25% Views',            // H
    '50% Views',            // I
    '75% Views',            // J
    '95% Views',            // K
    'CTA/Clicks',           // L
    'Landing Page Views',   // M
    'Leads',                // N
    'Hook Rate',            // O
    'CTR',                  // P
    'CPC',                  // Q
    'CPM',                  // R
    'Cost/CTA',             // S
    'LPV Rate',             // T
    'Cost/LPV',             // U
    'Lead Rate',            // V
    'CPL',                  // W
    'Frequency',            // X
    'Performance Score'     // Y
  ],

  // Scoring benchmark. Change these later if your real business data suggests better thresholds.
  BENCHMARKS: {
    HOOK_RATE: 0.40,        // 40%
    CTR: 0.03,              // 3%
    LPV_RATE: 0.80,         // 80%
    LEAD_RATE: 0.12,        // 12%
    TARGET_CPL: 100         // ৳100; lower is better
  }
};


/***************************************************************
 * MENU
 ***************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Marketing Report')
    .addItem('1. Setup / Migrate Lead-Only System', 'setupMarketingSystem')
    .addSeparator()
    .addItem('2. Update Raw Data', 'updateReport')
    .addItem('3. Refresh All Reports', 'refreshAllReports')
    .addSeparator()
    .addItem('4. Refresh Content Summary', 'generateContentSummary')
    .addItem('5. Refresh Project Summary', 'generateProjectSummary')
    .addItem('6. Refresh Dashboard', 'generateDashboard')
    .addSeparator()
    .addItem('7. Clear Raw Import Data', 'clearRawImport')
    .addSeparator()
    .addItem('8. Create Daily Automation', 'createDailyTrigger')
    .addItem('9. Remove Automation', 'removeAllTriggers')
    .addSeparator()
    .addItem('10. Setup Dashboard Controls', 'setupDashboardControls')
    .addToUi();
}


/***************************************************************
 * SAFE SETUP / MIGRATION
 *
 * Existing A:N Daily data is preserved.
 * Old Purchase/Revenue/CPA/ROAS metric area is cleared.
 ***************************************************************/
function setupMarketingSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheetIfMissing_(CONFIG.SHEETS.RAW);
  createSheetIfMissing_(CONFIG.SHEETS.DAILY);
  createSheetIfMissing_(CONFIG.SHEETS.MASTER);
  createSheetIfMissing_(CONFIG.SHEETS.CONTENT_SUMMARY);
  createSheetIfMissing_(CONFIG.SHEETS.PROJECT_SUMMARY);
  createSheetIfMissing_(CONFIG.SHEETS.DASHBOARD);

  setupRawImportSheet_();
  setupDailyPerformanceSheet_();
  setupContentMasterSheet_();
  setupContentSummarySheet_();
  setupProjectSummarySheet_();
  setupDashboardSheet_();

  // Rebuild formulas for existing Daily Performance rows.
  rebuildAllDailyMetrics_();

  generateContentSummary();
  generateProjectSummary();
  generateDashboard();

  SpreadsheetApp.flush();
  ss.toast('Lead-only marketing system is ready.', 'Setup Complete', 5);
}

function createSheetIfMissing_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}


/***************************************************************
 * RAW IMPORT SETUP
 ***************************************************************/
function setupRawImportSheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.RAW);

  ensureMinimumColumns_(sheet, CONFIG.RAW_HEADERS.length);

  // Preserve A:N values. Remove old Purchase/Revenue columns O:P and beyond.
  if (sheet.getMaxColumns() > CONFIG.RAW_HEADERS.length) {
    const extraCols = sheet.getMaxColumns() - CONFIG.RAW_HEADERS.length;
    sheet.getRange(1, CONFIG.RAW_HEADERS.length + 1, sheet.getMaxRows(), extraCols).clearContent().clearFormat();
  }

  sheet.getRange(1, 1, 1, CONFIG.RAW_HEADERS.length).setValues([CONFIG.RAW_HEADERS]);
  formatHeader_(sheet.getRange(1, 1, 1, CONFIG.RAW_HEADERS.length));

  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, CONFIG.RAW_HEADERS.length, 130);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 140);
  sheet.getRange('A:A').setNumberFormat('dd-mmm-yyyy');

  sheet.getRange('A2').setNote(
    'Paste daily data from row 2. Required key: Date + Project + Content ID. Final conversion is Leads.'
  );
}


/***************************************************************
 * DAILY PERFORMANCE SETUP
 ***************************************************************/
function setupDailyPerformanceSheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.DAILY);

  ensureMinimumColumns_(sheet, CONFIG.DAILY_HEADERS.length);

  // Preserve A:N raw history. Clear all old calculated / purchase / revenue columns from O onward.
  const maxCols = sheet.getMaxColumns();
  if (maxCols >= 15) {
    sheet.getRange(1, 15, sheet.getMaxRows(), maxCols - 14).clearContent().clearFormat();
  }

  sheet.getRange(1, 1, 1, CONFIG.DAILY_HEADERS.length).setValues([CONFIG.DAILY_HEADERS]);
  formatHeader_(sheet.getRange(1, 1, 1, CONFIG.DAILY_HEADERS.length));

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 140);
  for (let col = 4; col <= CONFIG.DAILY_HEADERS.length; col++) {
    sheet.setColumnWidth(col, 120);
  }

  sheet.getRange('A:A').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('D:N').setNumberFormat('#,##0.00');
  sheet.getRange('O:P').setNumberFormat('0.00%');
  sheet.getRange('Q:S').setNumberFormat('৳#,##0.00');
  sheet.getRange('T:T').setNumberFormat('0.00%');
  sheet.getRange('U:U').setNumberFormat('৳#,##0.00');
  sheet.getRange('V:V').setNumberFormat('0.00%');
  sheet.getRange('W:W').setNumberFormat('৳#,##0.00');
  sheet.getRange('X:X').setNumberFormat('0.00');
  sheet.getRange('Y:Y').setNumberFormat('0.0');
}


/***************************************************************
 * CONTENT MASTER SETUP
 ***************************************************************/
function setupContentMasterSheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.MASTER);

  const headers = [
    'Content ID',
    'Project',
    'Content Name',
    'Content Type',
    'Hook Type',
    'Angle/Topic',
    'CTA',
    'Creator/Presenter',
    'Duration (sec)',
    'Publish Date',
    'Status',
    'Notes'
  ];

  ensureMinimumColumns_(sheet, headers.length);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader_(sheet.getRange(1, 1, 1, headers.length));
  sheet.setFrozenRows(1);

  const widths = [140, 160, 260, 130, 150, 180, 140, 170, 110, 120, 120, 250];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange('J:J').setNumberFormat('dd-mmm-yyyy');

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Testing', 'Winner', 'Paused', 'Stopped'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange('K2:K1000').setDataValidation(statusRule);
}


/***************************************************************
 * CONTENT SUMMARY SETUP
 ***************************************************************/
function setupContentSummarySheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.CONTENT_SUMMARY);
  sheet.clear();

  const headers = [
    'Content ID',
    'Project',
    'Content Name',
    'Total Spend',
    'Impressions',
    '3s Views',
    'Clicks',
    'LPV',
    'Leads',
    'Hook Rate',
    'CTR',
    'LPV Rate',
    'Lead Rate',
    'CPC',
    'Cost/LPV',
    'CPL',
    'Frequency',
    'Score',
    'Status'
  ];

  ensureMinimumColumns_(sheet, headers.length);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader_(sheet.getRange(1, 1, 1, headers.length));
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 260);
  for (let col = 4; col <= headers.length; col++) sheet.setColumnWidth(col, 120);
}


/***************************************************************
 * PROJECT SUMMARY SETUP
 ***************************************************************/
function setupProjectSummarySheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.PROJECT_SUMMARY);
  sheet.clear();

  const headers = [
    'Project',
    'Spend',
    'Impressions',
    'Reach',
    '3s Views',
    'Clicks',
    'LPV',
    'Leads',
    'Hook Rate',
    'CTR',
    'LPV Rate',
    'Lead Rate',
    'CPC',
    'Cost/LPV',
    'CPL',
    'Frequency'
  ];

  ensureMinimumColumns_(sheet, headers.length);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader_(sheet.getRange(1, 1, 1, headers.length));
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 180);
  for (let col = 2; col <= headers.length; col++) sheet.setColumnWidth(col, 120);
}


/***************************************************************
 * DASHBOARD SETUP
 ***************************************************************/
function setupDashboardSheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.DASHBOARD);
  sheet.clear();
  sheet.clearConditionalFormatRules();

  sheet.getRange('A1:H2').merge();
  sheet.getRange('A1')
    .setValue('MARKETING CONTENT PERFORMANCE DASHBOARD — LEAD ONLY')
    .setBackground(CONFIG.COLORS.DARK)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 190);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 150);

  createDashboardControlPanel();
}


/***************************************************************
 * MAIN UPDATE
 ***************************************************************/
function updateReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Processing Raw Import data...', 'Marketing Report', 3);

  try {
    processRawImport_();
    generateContentSummary();
    generateProjectSummary();
    generateDashboard();
    SpreadsheetApp.flush();
    ss.toast('Lead report updated successfully.', 'Complete', 5);
  } catch (error) {
    console.error(error);
    ss.toast('ERROR: ' + error.message, 'Marketing Report', 8);
  }
}


/***************************************************************
 * RAW IMPORT -> DAILY PERFORMANCE
 ***************************************************************/
function processRawImport_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW);
  const dailySheet = getSheet_(CONFIG.SHEETS.DAILY);
  const lastRow = rawSheet.getLastRow();

  if (lastRow < 2) throw new Error('Raw Import sheet has no data.');

  const rawData = rawSheet
    .getRange(2, 1, lastRow - 1, CONFIG.RAW_HEADERS.length)
    .getValues();

  const cleanRows = rawData.filter(row => row[0] && row[1] && row[2]);
  if (cleanRows.length === 0) {
    throw new Error('No valid rows found. Date, Project and Content ID are required.');
  }

  const existingMap = getExistingDailyMap_();

  cleanRows.forEach(row => {
    const normalized = normalizeRawRow_(row);
    const key = createUniqueKey_(normalized[0], normalized[1], normalized[2]);

    if (existingMap[key]) {
      const rowNumber = existingMap[key];
      dailySheet.getRange(rowNumber, 1, 1, 14).setValues([normalized]);
      writeCalculatedMetrics_(dailySheet, rowNumber);
    } else {
      const newRow = dailySheet.getLastRow() + 1;
      dailySheet.getRange(newRow, 1, 1, 14).setValues([normalized]);
      writeCalculatedMetrics_(dailySheet, newRow);
      existingMap[key] = newRow;
    }
  });

  sortDailySheet_();
}

function normalizeRawRow_(row) {
  return [
    normalizeDate_(row[0]),
    String(row[1]).trim(),
    String(row[2]).trim(),
    number_(row[3]),
    number_(row[4]),
    number_(row[5]),
    number_(row[6]),
    number_(row[7]),
    number_(row[8]),
    number_(row[9]),
    number_(row[10]),
    number_(row[11]),
    number_(row[12]),
    number_(row[13])
  ];
}

function getExistingDailyMap_() {
  const sheet = getSheet_(CONFIG.SHEETS.DAILY);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  values.forEach((row, index) => {
    if (!row[0] || !row[1] || !row[2]) return;
    const key = createUniqueKey_(row[0], row[1], row[2]);
    map[key] = index + 2;
  });
  return map;
}

function createUniqueKey_(date, project, contentId) {
  const formattedDate = Utilities.formatDate(
    normalizeDate_(date),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
  return formattedDate + '||' + String(project).trim().toLowerCase() + '||' + String(contentId).trim().toLowerCase();
}


/***************************************************************
 * DAILY CALCULATED METRICS
 * O:Y = 11 calculated fields
 ***************************************************************/
function writeCalculatedMetrics_(sheet, row) {
  const formulas = [[
    `=IFERROR(G${row}/E${row},0)`,                         // O Hook Rate
    `=IFERROR(L${row}/E${row},0)`,                         // P CTR
    `=IFERROR(D${row}/L${row},0)`,                         // Q CPC
    `=IFERROR(D${row}/E${row}*1000,0)`,                    // R CPM
    `=IFERROR(D${row}/L${row},0)`,                         // S Cost/CTA
    `=IFERROR(M${row}/L${row},0)`,                         // T LPV Rate
    `=IFERROR(D${row}/M${row},0)`,                         // U Cost/LPV
    `=IFERROR(N${row}/M${row},0)`,                         // V Lead Rate
    `=IFERROR(D${row}/N${row},0)`,                         // W CPL
    `=IFERROR(E${row}/F${row},0)`,                         // X Frequency
    createPerformanceScoreFormula_(row)                    // Y Score
  ]];

  sheet.getRange(row, 15, 1, 11).setFormulas(formulas);
  formatDailyMetricsRow_(sheet, row);
}

function rebuildAllDailyMetrics_() {
  const sheet = getSheet_(CONFIG.SHEETS.DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  for (let row = 2; row <= lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() && sheet.getRange(row, 3).getValue()) {
      writeCalculatedMetrics_(sheet, row);
    }
  }
}

function createPerformanceScoreFormula_(row) {
  const b = CONFIG.BENCHMARKS;

  // 25 Hook + 25 CTR + 20 LPV + 20 Lead Rate + 10 CPL efficiency = 100
  // For CPL: target CPL or lower gets full 10 points.
  return (
    `=MIN(100,` +
    `MIN(O${row}/${b.HOOK_RATE},1)*25+` +
    `MIN(P${row}/${b.CTR},1)*25+` +
    `MIN(T${row}/${b.LPV_RATE},1)*20+` +
    `MIN(V${row}/${b.LEAD_RATE},1)*20+` +
    `IF(W${row}=0,0,MIN(${b.TARGET_CPL}/W${row},1)*10)` +
    `)`
  );
}

function formatDailyMetricsRow_(sheet, row) {
  sheet.getRange(row, 15, 1, 2).setNumberFormat('0.00%'); // O:P
  sheet.getRange(row, 17, 1, 3).setNumberFormat('৳#,##0.00'); // Q:S
  sheet.getRange(row, 20).setNumberFormat('0.00%'); // T
  sheet.getRange(row, 21).setNumberFormat('৳#,##0.00'); // U
  sheet.getRange(row, 22).setNumberFormat('0.00%'); // V
  sheet.getRange(row, 23).setNumberFormat('৳#,##0.00'); // W
  sheet.getRange(row, 24).setNumberFormat('0.00'); // X
  sheet.getRange(row, 25).setNumberFormat('0.0'); // Y
}

function sortDailySheet_() {
  const sheet = getSheet_(CONFIG.SHEETS.DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  sheet.getRange(2, 1, lastRow - 1, CONFIG.DAILY_HEADERS.length).sort([
    { column: 1, ascending: true },
    { column: 2, ascending: true },
    { column: 3, ascending: true }
  ]);
}


/***************************************************************
 * CONTENT SUMMARY
 ***************************************************************/
function generateContentSummary() {
  const dailySheet = getSheet_(CONFIG.SHEETS.DAILY);
  const masterSheet = getSheet_(CONFIG.SHEETS.MASTER);
  const outputSheet = getSheet_(CONFIG.SHEETS.CONTENT_SUMMARY);

  setupContentSummarySheet_();

  const dailyLastRow = dailySheet.getLastRow();
  if (dailyLastRow < 2) return;

  const dailyData = dailySheet.getRange(2, 1, dailyLastRow - 1, 25).getValues();
  const masterMap = getContentMasterMap_(masterSheet);
  const summary = {};

  dailyData.forEach(row => {
    const project = row[1];
    const contentId = String(row[2]).trim();
    if (!contentId) return;

    if (!summary[contentId]) {
      summary[contentId] = {
        contentId,
        project,
        spend: 0,
        impressions: 0,
        reach: 0,
        views3s: 0,
        clicks: 0,
        lpv: 0,
        leads: 0
      };
    }

    const item = summary[contentId];
    item.spend += number_(row[3]);
    item.impressions += number_(row[4]);
    item.reach += number_(row[5]);
    item.views3s += number_(row[6]);
    item.clicks += number_(row[11]);
    item.lpv += number_(row[12]);
    item.leads += number_(row[13]);
  });

  const rows = [];

  Object.keys(summary).forEach(contentId => {
    const item = summary[contentId];
    const master = masterMap[contentId] || {};

    const hookRate = safeDivide_(item.views3s, item.impressions);
    const ctr = safeDivide_(item.clicks, item.impressions);
    const lpvRate = safeDivide_(item.lpv, item.clicks);
    const leadRate = safeDivide_(item.leads, item.lpv);
    const cpc = safeDivide_(item.spend, item.clicks);
    const costLpv = safeDivide_(item.spend, item.lpv);
    const cpl = safeDivide_(item.spend, item.leads);
    const frequency = safeDivide_(item.impressions, item.reach);

    const score = calculateScore_(hookRate, ctr, lpvRate, leadRate, cpl);
    const status = getPerformanceStatus_(score, item.leads);

    rows.push([
      item.contentId,
      item.project,
      master.contentName || '',
      item.spend,
      item.impressions,
      item.views3s,
      item.clicks,
      item.lpv,
      item.leads,
      hookRate,
      ctr,
      lpvRate,
      leadRate,
      cpc,
      costLpv,
      cpl,
      frequency,
      score,
      status
    ]);
  });

  rows.sort((a, b) => b[17] - a[17]);

  if (rows.length > 0) {
    outputSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    outputSheet.getRange(2, 4, rows.length, 6).setNumberFormat('#,##0.00');
    outputSheet.getRange(2, 10, rows.length, 4).setNumberFormat('0.00%');
    outputSheet.getRange(2, 14, rows.length, 3).setNumberFormat('৳#,##0.00');
    outputSheet.getRange(2, 17, rows.length, 1).setNumberFormat('0.00');
    outputSheet.getRange(2, 18, rows.length, 1).setNumberFormat('0.0');
    applySummaryStatusColors_(outputSheet, rows.length);
  }
}

function getContentMasterMap_(sheet) {
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  values.forEach(row => {
    const contentId = String(row[0]).trim();
    if (!contentId) return;
    map[contentId] = {
      project: row[1],
      contentName: row[2],
      contentType: row[3],
      hookType: row[4],
      angle: row[5],
      cta: row[6],
      creator: row[7],
      duration: row[8],
      publishDate: row[9],
      status: row[10],
      notes: row[11]
    };
  });
  return map;
}


/***************************************************************
 * PROJECT SUMMARY
 ***************************************************************/
function generateProjectSummary() {
  const dailySheet = getSheet_(CONFIG.SHEETS.DAILY);
  const outputSheet = getSheet_(CONFIG.SHEETS.PROJECT_SUMMARY);

  setupProjectSummarySheet_();

  const lastRow = dailySheet.getLastRow();
  if (lastRow < 2) return;

  const values = dailySheet.getRange(2, 1, lastRow - 1, 25).getValues();
  const summary = {};

  values.forEach(row => {
    const project = String(row[1]).trim();
    if (!project) return;

    if (!summary[project]) {
      summary[project] = {
        spend: 0,
        impressions: 0,
        reach: 0,
        views3s: 0,
        clicks: 0,
        lpv: 0,
        leads: 0
      };
    }

    const item = summary[project];
    item.spend += number_(row[3]);
    item.impressions += number_(row[4]);
    item.reach += number_(row[5]);
    item.views3s += number_(row[6]);
    item.clicks += number_(row[11]);
    item.lpv += number_(row[12]);
    item.leads += number_(row[13]);
  });

  const rows = [];

  Object.keys(summary).forEach(project => {
    const item = summary[project];
    const hookRate = safeDivide_(item.views3s, item.impressions);
    const ctr = safeDivide_(item.clicks, item.impressions);
    const lpvRate = safeDivide_(item.lpv, item.clicks);
    const leadRate = safeDivide_(item.leads, item.lpv);
    const cpc = safeDivide_(item.spend, item.clicks);
    const costLpv = safeDivide_(item.spend, item.lpv);
    const cpl = safeDivide_(item.spend, item.leads);
    const frequency = safeDivide_(item.impressions, item.reach);

    rows.push([
      project,
      item.spend,
      item.impressions,
      item.reach,
      item.views3s,
      item.clicks,
      item.lpv,
      item.leads,
      hookRate,
      ctr,
      lpvRate,
      leadRate,
      cpc,
      costLpv,
      cpl,
      frequency
    ]);
  });

  // Best (lowest) CPL first, but zero-lead projects go to bottom.
  rows.sort((a, b) => {
    const aCpl = a[14] || Number.MAX_SAFE_INTEGER;
    const bCpl = b[14] || Number.MAX_SAFE_INTEGER;
    return aCpl - bCpl;
  });

  if (rows.length > 0) {
    outputSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    outputSheet.getRange(2, 2, rows.length, 7).setNumberFormat('#,##0.00');
    outputSheet.getRange(2, 9, rows.length, 4).setNumberFormat('0.00%');
    outputSheet.getRange(2, 13, rows.length, 3).setNumberFormat('৳#,##0.00');
    outputSheet.getRange(2, 16, rows.length, 1).setNumberFormat('0.00');
  }
}


/***************************************************************
 * DASHBOARD
 ***************************************************************/
function generateDashboard() {
  const dashboard = getSheet_(CONFIG.SHEETS.DASHBOARD);
  const selectedProject = dashboard.getRange('B4').getValue() || 'All Projects';

  setupDashboardSheet_();
  createProjectDropdown_();
  dashboard.getRange('B4').setValue(selectedProject);

  const projectData = getProjectDashboardData_(selectedProject);
  writeDashboardKPIs_(dashboard, projectData);
  writeDashboardFunnel_(dashboard, projectData);
  writeDashboardContentTable_(dashboard, selectedProject);
  createDashboardCharts_(dashboard);
}

function createProjectDropdown_() {
  const dashboard = getSheet_(CONFIG.SHEETS.DASHBOARD);
  const projectSheet = getSheet_(CONFIG.SHEETS.PROJECT_SUMMARY);
  const lastRow = projectSheet.getLastRow();

  let projects = ['All Projects'];
  if (lastRow >= 2) {
    const values = projectSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Boolean);
    projects = projects.concat(values);
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(projects, true)
    .setAllowInvalid(false)
    .build();
  dashboard.getRange('B4').setDataValidation(rule);
}

function getProjectDashboardData_(selectedProject) {
  const dailySheet = getSheet_(CONFIG.SHEETS.DAILY);
  const lastRow = dailySheet.getLastRow();

  const result = {
    spend: 0,
    impressions: 0,
    reach: 0,
    views3s: 0,
    clicks: 0,
    lpv: 0,
    leads: 0
  };

  if (lastRow < 2) return result;

  const values = dailySheet.getRange(2, 1, lastRow - 1, 14).getValues();
  values.forEach(row => {
    const project = String(row[1]).trim();
    if (selectedProject !== 'All Projects' && project !== selectedProject) return;

    result.spend += number_(row[3]);
    result.impressions += number_(row[4]);
    result.reach += number_(row[5]);
    result.views3s += number_(row[6]);
    result.clicks += number_(row[11]);
    result.lpv += number_(row[12]);
    result.leads += number_(row[13]);
  });

  return result;
}

function writeDashboardKPIs_(sheet, data) {
  const ctr = safeDivide_(data.clicks, data.impressions);
  const hookRate = safeDivide_(data.views3s, data.impressions);
  const lpvRate = safeDivide_(data.lpv, data.clicks);
  const leadRate = safeDivide_(data.leads, data.lpv);
  const cpl = safeDivide_(data.spend, data.leads);
  const frequency = safeDivide_(data.impressions, data.reach);

  const cards = [
    ['A7', 'Total Spend', data.spend, '৳#,##0.00'],
    ['C7', 'Impressions', data.impressions, '#,##0'],
    ['E7', 'Clicks', data.clicks, '#,##0'],
    ['G7', 'CTR', ctr, '0.00%'],
    ['A10', 'Leads', data.leads, '#,##0'],
    ['C10', 'CPL', cpl, '৳#,##0.00'],
    ['E10', 'Hook Rate', hookRate, '0.00%'],
    ['G10', 'LPV Rate', lpvRate, '0.00%'],
    ['A13', 'Lead Rate', leadRate, '0.00%'],
    ['C13', 'LPV', data.lpv, '#,##0'],
    ['E13', '3s Views', data.views3s, '#,##0'],
    ['G13', 'Frequency', frequency, '0.00']
  ];

  cards.forEach(card => {
    const labelCell = sheet.getRange(card[0]);
    const col = labelCell.getColumn();
    const row = labelCell.getRow();

    labelCell
      .setValue(card[1])
      .setBackground(CONFIG.COLORS.LIGHT_BLUE)
      .setFontColor(CONFIG.COLORS.DARK)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    sheet.getRange(row + 1, col)
      .setValue(card[2])
      .setNumberFormat(card[3])
      .setFontWeight('bold')
      .setFontSize(15)
      .setHorizontalAlignment('center')
      .setBackground(CONFIG.COLORS.WHITE);
  });
}

function writeDashboardFunnel_(sheet, data) {
  sheet.getRange('A16:F16').merge()
    .setValue('MARKETING FUNNEL — FINAL RESULT = LEAD')
    .setBackground(CONFIG.COLORS.BLUE)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold');

  const headers = ['Stage', 'Count', 'Conversion Rate', 'Cost / Stage', 'Diagnosis', 'Suggested Benchmark'];
  sheet.getRange('A17:F17').setValues([headers]);
  formatHeader_(sheet.getRange('A17:F17'));

  const rows = [
    [
      'Impressions',
      data.impressions,
      '',
      safeDivide_(data.spend * 1000, data.impressions),
      'Ad delivery / audience cost',
      'Track CPM'
    ],
    [
      '3s Views',
      data.views3s,
      safeDivide_(data.views3s, data.impressions),
      safeDivide_(data.spend, data.views3s),
      'Opening hook quality',
      'Hook Rate 30-40%+'
    ],
    [
      'Clicks',
      data.clicks,
      safeDivide_(data.clicks, data.impressions),
      safeDivide_(data.spend, data.clicks),
      'Creative + CTA strength',
      'CTR 2%+'
    ],
    [
      'Landing Page Views',
      data.lpv,
      safeDivide_(data.lpv, data.clicks),
      safeDivide_(data.spend, data.lpv),
      'Click quality + page load',
      'LPV Rate 70-80%+'
    ],
    [
      'Leads',
      data.leads,
      safeDivide_(data.leads, data.lpv),
      safeDivide_(data.spend, data.leads),
      'Final marketing conversion',
      'Track CPL + Lead Rate'
    ]
  ];

  sheet.getRange(18, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(18, 3, rows.length, 1).setNumberFormat('0.00%');
  sheet.getRange(18, 4, rows.length, 1).setNumberFormat('৳#,##0.00');
}

function writeDashboardContentTable_(dashboard, selectedProject) {
  const contentSheet = getSheet_(CONFIG.SHEETS.CONTENT_SUMMARY);
  const lastRow = contentSheet.getLastRow();

  dashboard.getRange('A26:I26').merge()
    .setValue('TOP CONTENT PERFORMANCE — LEAD GENERATION')
    .setBackground(CONFIG.COLORS.BLUE)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold');

  const headers = ['Content ID', 'Content Name', 'Spend', 'Hook Rate', 'CTR', 'LPV Rate', 'Leads', 'CPL', 'Score'];
  dashboard.getRange('A27:I27').setValues([headers]);
  formatHeader_(dashboard.getRange('A27:I27'));

  if (lastRow < 2) return;

  let values = contentSheet.getRange(2, 1, lastRow - 1, 19).getValues();
  if (selectedProject !== 'All Projects') values = values.filter(row => row[1] === selectedProject);

  values.sort((a, b) => b[17] - a[17]);
  const top10 = values.slice(0, 10);

  const output = top10.map(row => [
    row[0], // Content ID
    row[2], // Name
    row[3], // Spend
    row[9], // Hook Rate
    row[10], // CTR
    row[11], // LPV Rate
    row[8], // Leads
    row[15], // CPL
    row[17] // Score
  ]);

  if (output.length > 0) {
    dashboard.getRange(28, 1, output.length, 9).setValues(output);
    dashboard.getRange(28, 3, output.length, 1).setNumberFormat('৳#,##0.00');
    dashboard.getRange(28, 4, output.length, 3).setNumberFormat('0.00%');
    dashboard.getRange(28, 8, output.length, 1).setNumberFormat('৳#,##0.00');
    dashboard.getRange(28, 9, output.length, 1).setNumberFormat('0.0');
  }
}

function createDashboardCharts_(dashboard) {
  dashboard.getCharts().forEach(chart => dashboard.removeChart(chart));

  const funnelChart = dashboard.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dashboard.getRange('A17:B22'))
    .setPosition(16, 11, 0, 0)
    .setOption('title', 'Lead Funnel')
    .setOption('legend', { position: 'none' })
    .build();
  dashboard.insertChart(funnelChart);

  const contentChart = dashboard.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dashboard.getRange('A27:A37'))
    .addRange(dashboard.getRange('G27:G37'))
    .setPosition(28, 11, 0, 0)
    .setOption('title', 'Top Content by Leads')
    .setOption('legend', { position: 'none' })
    .build();
  dashboard.insertChart(contentChart);
}


/***************************************************************
 * REFRESH / CLEAR / AUTOMATION
 ***************************************************************/
function refreshAllReports() {
  generateContentSummary();
  generateProjectSummary();
  generateDashboard();
  SpreadsheetApp.getActiveSpreadsheet().toast('All lead reports refreshed.', 'Marketing Report', 4);
}

function clearRawImport() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear Raw Import?',
    'This deletes Raw Import rows only. Daily Performance history stays safe.',
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;
  clearRawImportFromControl_();
}

function createDailyTrigger() {
  removeSpecificTrigger_('refreshAllReports');
  ScriptApp.newTrigger('refreshAllReports').timeBased().everyDays(1).atHour(23).create();
  SpreadsheetApp.getUi().alert('Daily report refresh automation created for around 11 PM.');
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  SpreadsheetApp.getUi().alert('All Apps Script triggers have been removed.');
}

function removeSpecificTrigger_(functionName) {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) ScriptApp.deleteTrigger(trigger);
  });
}


/***************************************************************
 * PERFORMANCE SCORE
 ***************************************************************/
function calculateScore_(hookRate, ctr, lpvRate, leadRate, cpl) {
  const b = CONFIG.BENCHMARKS;

  const hookScore = Math.min(hookRate / b.HOOK_RATE, 1) * 25;
  const ctrScore = Math.min(ctr / b.CTR, 1) * 25;
  const lpvScore = Math.min(lpvRate / b.LPV_RATE, 1) * 20;
  const leadScore = Math.min(leadRate / b.LEAD_RATE, 1) * 20;
  const cplScore = cpl > 0 ? Math.min(b.TARGET_CPL / cpl, 1) * 10 : 0;

  return Number((hookScore + ctrScore + lpvScore + leadScore + cplScore).toFixed(2));
}

function getPerformanceStatus_(score, leads) {
  if (leads <= 0) return '🔴 No Lead';
  if (score >= 80) return '🔥 Winner';
  if (score >= 65) return '🟢 Good';
  if (score >= 45) return '🟡 Average';
  return '🔴 Poor';
}

function applySummaryStatusColors_(sheet, rowCount) {
  if (rowCount <= 0) return;

  const range = sheet.getRange(2, 19, rowCount, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Winner')
      .setBackground(CONFIG.COLORS.LIGHT_GREEN)
      .setFontColor('#166534')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Good')
      .setBackground(CONFIG.COLORS.LIGHT_BLUE)
      .setFontColor('#1E40AF')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Average')
      .setBackground(CONFIG.COLORS.LIGHT_YELLOW)
      .setFontColor('#92400E')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Poor')
      .setBackground(CONFIG.COLORS.LIGHT_RED)
      .setFontColor('#991B1B')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('No Lead')
      .setBackground(CONFIG.COLORS.LIGHT_RED)
      .setFontColor('#991B1B')
      .setRanges([range])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
}


/***************************************************************
 * DASHBOARD CONTROL PANEL
 ***************************************************************/
const CONTROL_PANEL = {
  SHEET: 'Dashboard',
  BUTTONS: {
    D5: { name: 'UPDATE REPORT', action: 'updateReport' },
    E5: { name: 'REFRESH ALL', action: 'refreshAllReports' },
    F5: { name: 'DASHBOARD', action: 'generateDashboard' },
    G5: { name: 'CLEAR RAW', action: 'clearRawImportFromControl_' },
    H5: { name: 'AUTOMATION', action: 'createDailyTriggerFromControl_' }
  }
};

function createDashboardControlPanel() {
  const sheet = getSheet_(CONTROL_PANEL.SHEET);

  sheet.getRange('A4')
    .setValue('SELECT PROJECT')
    .setBackground(CONFIG.COLORS.BLUE)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('B4')
    .setValue(sheet.getRange('B4').getValue() || 'All Projects')
    .setBackground('#EFF6FF')
    .setFontColor(CONFIG.COLORS.DARK)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('D4:H4').setValues([[
    'UPDATE REPORT', 'REFRESH ALL', 'DASHBOARD', 'CLEAR RAW', 'AUTOMATION'
  ]]);

  sheet.getRange('D4:H4')
    .setBackground(CONFIG.COLORS.DARK)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('D5:H5').insertCheckboxes().uncheck().setHorizontalAlignment('center');
  sheet.getRange('D5').setBackground(CONFIG.COLORS.LIGHT_GREEN);
  sheet.getRange('E5').setBackground(CONFIG.COLORS.LIGHT_BLUE);
  sheet.getRange('F5').setBackground('#E0F2FE');
  sheet.getRange('G5').setBackground(CONFIG.COLORS.LIGHT_RED);
  sheet.getRange('H5').setBackground(CONFIG.COLORS.LIGHT_YELLOW);

  sheet.getRange('A5')
    .setValue('LAST UPDATED')
    .setBackground('#F1F5F9')
    .setFontColor(CONFIG.COLORS.GRAY)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('B5')
    .setBackground(CONFIG.COLORS.WHITE)
    .setFontColor(CONFIG.COLORS.DARK)
    .setHorizontalAlignment('center')
    .setNumberFormat('dd-mmm-yyyy hh:mm:ss AM/PM');

  sheet.setRowHeight(4, 35);
  sheet.setRowHeight(5, 35);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 130);

  sheet.getRange('D5').setNote('Tick to import Raw Import data and refresh lead reports.');
  sheet.getRange('E5').setNote('Tick to refresh summaries and dashboard using existing Daily Performance data.');
  sheet.getRange('F5').setNote('Tick to refresh Dashboard only.');
  sheet.getRange('G5').setNote('Tick to clear Raw Import only. Daily history is preserved.');
  sheet.getRange('H5').setNote('Tick to create daily refresh automation around 11 PM.');
}

function setupDashboardControls() {
  createDashboardControlPanel();
  createProjectDropdown_();
  installControlPanelTrigger();
  updateLastUpdated_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard controls activated.', 'Setup Complete', 5);
}

function installControlPanelTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'controlPanelOnEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('controlPanelOnEdit').forSpreadsheet(ss).onEdit().create();
  ss.toast('Control Panel trigger installed.', 'Setup Complete', 5);
}

function controlPanelOnEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== CONTROL_PANEL.SHEET) return;

  const a1 = range.getA1Notation();

  if (a1 === 'B4') {
    handleProjectChange_();
    return;
  }

  const buttonConfig = CONTROL_PANEL.BUTTONS[a1];
  if (!buttonConfig || e.value !== 'TRUE') return;

  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(30000);
    setControlStatus_('RUNNING: ' + buttonConfig.name);
    runControlAction_(buttonConfig.action);
    updateLastUpdated_();
    setControlStatus_('DONE: ' + buttonConfig.name);
  } catch (error) {
    console.error(error);
    setControlStatus_('ERROR: ' + error.message);
  } finally {
    range.uncheck();
    try { lock.releaseLock(); } catch (err) { console.error(err); }
  }
}

function runControlAction_(actionName) {
  switch (actionName) {
    case 'updateReport':
      updateReport();
      break;
    case 'refreshAllReports':
      refreshAllReports();
      break;
    case 'generateDashboard':
      generateDashboard();
      break;
    case 'clearRawImportFromControl_':
      clearRawImportFromControl_();
      break;
    case 'createDailyTriggerFromControl_':
      createDailyTriggerFromControl_();
      break;
    default:
      throw new Error('Unknown Control Panel action: ' + actionName);
  }
}

function clearRawImportFromControl_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(CONFIG.SHEETS.RAW);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    ss.toast('Raw Import is already empty.', 'Marketing Report', 3);
    return;
  }

  sheet.getRange(2, 1, lastRow - 1, CONFIG.RAW_HEADERS.length).clearContent();
  ss.toast('Raw Import cleared. Daily history is safe.', 'Marketing Report', 4);
}

function createDailyTriggerFromControl_() {
  removeSpecificTrigger_('refreshAllReports');
  ScriptApp.newTrigger('refreshAllReports').timeBased().everyDays(1).atHour(23).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Daily refresh automation enabled.', 'Automation', 4);
}

function handleProjectChange_() {
  setControlStatus_('Updating project dashboard...');
  generateDashboard();
  updateLastUpdated_();
  setControlStatus_('Project dashboard updated');
}

function updateLastUpdated_() {
  const sheet = getSheet_(CONTROL_PANEL.SHEET);
  sheet.getRange('B5').setValue(new Date()).setNumberFormat('dd-mmm-yyyy hh:mm:ss AM/PM');
}

function setControlStatus_(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONTROL_PANEL.SHEET);
  if (sheet) sheet.getRange('A1').setNote(message);
  ss.toast(message, 'Marketing Control Panel', 3);
}


/***************************************************************
 * UTILITIES
 ***************************************************************/
function formatHeader_(range) {
  range
    .setBackground(CONFIG.COLORS.DARK)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}

function getSheet_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found. Run setupMarketingSystem first.`);
  return sheet;
}

function ensureMinimumColumns_(sheet, neededColumns) {
  const current = sheet.getMaxColumns();
  if (current < neededColumns) sheet.insertColumnsAfter(current, neededColumns - current);
}

function number_(value) {
  if (value === '' || value === null || value === undefined) return 0;
  if (typeof value === 'number') return value;

  const text = String(value)
    .replace(/,/g, '')
    .replace(/৳/g, '')
    .replace(/\$/g, '')
    .replace(/%/g, '')
    .trim();

  const num = parseFloat(text);
  return isNaN(num) ? 0 : num;
}

function safeDivide_(numerator, denominator) {
  numerator = number_(numerator);
  denominator = number_(denominator);
  return denominator === 0 ? 0 : numerator / denominator;
}

function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value)) return value;

  if (typeof value === 'number') {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }

  if (typeof value === 'string') {
    const cleaned = value.trim();
    const parts = cleaned.split(/[-/.]/);

    if (parts.length === 3) {
      let day = parseInt(parts[0], 10);
      let month = parseInt(parts[1], 10);
      let year = parseInt(parts[2], 10);

      if (parts[0].length === 4) {
        year = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10);
        day = parseInt(parts[2], 10);
      }

      if (year < 100) year += 2000;
      const constructed = new Date(year, month - 1, day);
      if (!isNaN(constructed)) return constructed;
    }

    const parsed = new Date(cleaned);
    if (!isNaN(parsed)) return parsed;
  }

  throw new Error('Invalid date detected: ' + value);
}
