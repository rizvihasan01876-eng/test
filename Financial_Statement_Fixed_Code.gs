/**
 * PROFESSIONAL FINANCIAL STATEMENT
 * Source sheets are never edited or deleted.
 * Creates separate dashboard sheets with linked data.
 */

const FS_CONFIG = {
  dashboardPrefix: "FS - ",
  navy: "#123B5D",
  blue: "#1F4E78",
  lightBlue: "#DCE6F1",
  paleBlue: "#F3F7FB",
  white: "#FFFFFF",
  gray: "#64748B",
  red: "#FCE8E6",
  green: "#E2F0D9",
  border: "#D5DDE5"
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("FINANCIAL TOOLS")
    .addItem("Create / Refresh Dashboard", "buildFinancialDashboards")
    .addItem("Export Active Dashboard PDF", "exportFinancialDashboardPDF")
    .addToUi();
}

function buildFinancialDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sources = ss.getSheets().filter(s =>
    !s.getName().startsWith(FS_CONFIG.dashboardPrefix)
  );

  if (!sources.length) {
    SpreadsheetApp.getUi().alert("No source sheets found.");
    return;
  }

  sources.forEach(source => {
    const data = source.getDataRange().getDisplayValues();
    if (!data.length || !data.some(r => r.some(v => v !== ""))) return;

    const name = (FS_CONFIG.dashboardPrefix + source.getName())
      .substring(0, 99);

    let dash = ss.getSheetByName(name);
    if (!dash) {
      dash = ss.insertSheet(name);
    }

    renderFinancialDashboard_(ss, source, dash, data);
  });

  SpreadsheetApp.getUi().alert(
    "Dashboard created/refreshed!\nOriginal sheets were not edited."
  );
}

function renderFinancialDashboard_(ss, source, dash, data) {
  const C = FS_CONFIG;

  // Only clear the generated dashboard sheet.
  // Source sheets remain untouched.
  dash.getRange(
    1, 1, dash.getMaxRows(), dash.getMaxColumns()
  ).breakApart();

  dash.clear();
  dash.setHiddenGridlines(true);
  dash.setTabColor(C.navy);
  dash.setFrozenRows(3);

  const maxSourceCols = Math.max(
    ...data.map(row => row.length), 2
  );

  const labelCol = 1;
  const amountCol = Math.min(2, maxSourceCols);

  // Read source data and classify rows by their labels.
  const assets = [];
  const liabilities = [];
  const equity = [];
  const other = [];

  let section = "";
  let totalAssets = 0;
  let totalLiabilities = 0;
  let totalEquity = 0;

  data.forEach((row, index) => {
    const label = String(row[0] || "").trim();
    const joined = row.join(" ").toLowerCase();
    const lower = label.toLowerCase();
    const amount = parseMoney_(row[amountCol - 1]);

    if (!label && row.every(v => !v)) return;

    if (/\basset(s)?\b/.test(lower)) {
      section = "assets";
      return;
    }

    if (/liabilit|payable|due\b/.test(lower)) {
      section = "liabilities";
      if (/owner.?s equity|equity|capital/.test(lower)) {
        section = "equity";
      }
      return;
    }

    if (/owner.?s equity|shareholder.?s equity|capital/.test(lower)) {
      section = "equity";
      return;
    }

    if (/financial statement/.test(lower)) return;

    // Preserve each non-empty source row as a linked detail row.
    const item = {
      sourceRow: index + 1,
      label: label || "(Details)",
      amount: amount,
      row: row
    };

    if (section === "assets") {
      assets.push(item);
      if (!/total/i.test(label)) totalAssets += amount;
    } else if (section === "liabilities") {
      liabilities.push(item);
      if (!/total/i.test(label)) totalLiabilities += amount;
    } else if (section === "equity") {
      equity.push(item);
      if (!/total/i.test(label)) totalEquity += amount;
    } else {
      other.push(item);
    }
  });

  // Layout: A:H
  dash.setColumnWidth(1, 28);
  dash.setColumnWidth(2, 205);
  dash.setColumnWidth(3, 125);
  dash.setColumnWidth(4, 24);
  dash.setColumnWidth(5, 205);
  dash.setColumnWidth(6, 125);
  dash.setColumnWidth(7, 24);
  dash.setColumnWidth(8, 110);

  dash.getRange("A1:H2").merge()
    .setValue("FINANCIAL STATEMENT")
    .setBackground(C.navy)
    .setFontColor(C.white)
    .setFontFamily("Arial")
    .setFontSize(20)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  dash.setRowHeight(1, 30);
  dash.setRowHeight(2, 24);

  dash.getRange("A3:H3").merge()
    .setValue(
      source.getName() + "  |  " +
      Utilities.formatDate(
        new Date(), ss.getSpreadsheetTimeZone(), "dd MMM yyyy"
      )
    )
    .setBackground(C.lightBlue)
    .setFontColor(C.navy)
    .setFontSize(10)
    .setHorizontalAlignment("center");

  // Summary cards
  const cardRows = [
    ["TOTAL ASSETS", totalAssets],
    ["TOTAL LIABILITIES", totalLiabilities],
    ["TOTAL EQUITY", totalEquity]
  ];

  const cardStarts = [1, 3, 5];

  cardRows.forEach((card, i) => {
    const col = cardStarts[i];
    dash.getRange(5, col, 1, 2).merge()
      .setValue(card[0])
      .setBackground(C.blue)
      .setFontColor(C.white)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    dash.getRange(6, col, 2, 2).merge()
      .setValue(card[1])
      .setNumberFormat('"৳" #,##0.00;[Red]-"৳" #,##0.00')
      .setBackground(C.paleBlue)
      .setFontColor(C.navy)
      .setFontSize(15)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  });

  // Side-by-side tables
  writeTable_(dash, 9, 1, "ASSETS", assets, source, 2);
  writeTable_(dash, 9, 5, "LIABILITIES & EQUITY",
    liabilities.concat(equity), source, 2);

  const leftEnd = 10 + assets.length;
  const rightEnd = 10 + liabilities.length + equity.length;
  const balanceRow = Math.max(leftEnd, rightEnd) + 2;

  dash.getRange(balanceRow, 1, 1, 6).merge()
    .setValue("FINANCIAL BALANCE SUMMARY")
    .setBackground(C.navy)
    .setFontColor(C.white)
    .setFontWeight("bold")
    .setFontSize(12)
    .setHorizontalAlignment("center");

  dash.getRange(balanceRow + 1, 1, 1, 3).merge()
    .setValue("Assets − (Liabilities + Equity)")
    .setBackground(C.lightBlue)
    .setFontColor(C.navy)
    .setFontWeight("bold");

  const difference = totalAssets - totalLiabilities - totalEquity;
  dash.getRange(balanceRow + 1, 4, 1, 3).merge()
    .setValue(difference)
    .setNumberFormat('"৳" #,##0.00;[Red]-"৳" #,##0.00')
    .setFontWeight("bold")
    .setFontSize(13)
    .setHorizontalAlignment("center")
    .setBackground(Math.abs(difference) < 0.005 ? C.green : C.red);

  dash.getRange(balanceRow + 2, 1, 1, 6).merge()
    .setValue(
      Math.abs(difference) < 0.005
        ? "BALANCED"
        : "NOT BALANCED — Please verify source data and classification."
    )
    .setFontWeight("bold")
    .setFontColor(
      Math.abs(difference) < 0.005 ? "#38761D" : "#990000"
    )
    .setHorizontalAlignment("center");

  // Additional source rows that couldn't be classified.
  if (other.length) {
    const otherStart = balanceRow + 5;
    dash.getRange(otherStart, 1, 1, 6).merge()
      .setValue("OTHER SOURCE DATA — REVIEW")
      .setBackground(C.blue)
      .setFontColor(C.white)
      .setFontWeight("bold");

    other.forEach((item, i) => {
      const r = otherStart + 1 + i;
      dash.getRange(r, 1, 1, 6).merge()
        .setFormula(
          "='" + source.getName().replace(/'/g, "''") +
          "'!A" + item.sourceRow
        );
      dash.getRange(r, 1, 1, 6)
        .setBackground(i % 2 ? C.paleBlue : C.white);
    });
  }

  const usedRows = dash.getLastRow();
  const usedCols = 6;

  dash.getRange(1, 1, usedRows, usedCols)
    .setFontFamily("Arial")
    .setVerticalAlignment("middle");

  dash.getRange(1, 1, usedRows, usedCols)
    .setBorder(
      true, true, true, true, true, true,
      C.border, SpreadsheetApp.BorderStyle.SOLID
    );

  dash.setRowHeights(9, Math.max(1, usedRows - 8), 25);
}

function writeTable_(dash, startRow, startCol, title, items, source, amountIndex) {
  const C = FS_CONFIG;

  dash.getRange(startRow, startCol, 1, 2).merge()
    .setValue(title)
    .setBackground(C.navy)
    .setFontColor(C.white)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  dash.getRange(startRow + 1, startCol, 1, 2)
    .setValues([["Description", "Amount (BDT)"]])
    .setBackground(C.lightBlue)
    .setFontColor(C.navy)
    .setFontWeight("bold");

  if (!items.length) {
    dash.getRange(startRow + 2, startCol, 1, 2)
      .setValues([["No classified rows", 0]]);
    return;
  }

  items.forEach((item, i) => {
    const r = startRow + 2 + i;
    const sourceName = source.getName().replace(/'/g, "''");

    // Live link to source description and amount.
    dash.getRange(r, startCol)
      .setFormula("='" + sourceName + "'!A" + item.sourceRow);

    const amountColumnLetter = columnLetter_(amountIndex);
    dash.getRange(r, startCol + 1)
      .setFormula(
        "='" + sourceName + "'!" +
        amountColumnLetter + item.sourceRow
      );

    dash.getRange(r, startCol, 1, 2)
      .setBackground(i % 2 ? C.paleBlue : C.white);
  });

  const totalRow = startRow + 2 + items.length;
  dash.getRange(totalRow, startCol)
    .setValue("TOTAL")
    .setFontWeight("bold")
    .setBackground(C.lightBlue);

  dash.getRange(totalRow, startCol + 1)
    .setFormula(
      "=SUM(" + columnLetter_(startCol + 1) +
      (startRow + 2) + ":" +
      columnLetter_(startCol + 1) + (totalRow - 1) + ")"
    )
    .setFontWeight("bold")
    .setBackground(C.lightBlue)
    .setNumberFormat('"৳" #,##0.00;[Red]-"৳" #,##0.00');

  dash.getRange(
    startRow + 2, startCol + 1, items.length, 1
  ).setNumberFormat('"৳" #,##0.00;[Red]-"৳" #,##0.00');
}

function parseMoney_(value) {
  if (typeof value === "number") return value;
  if (value === null || value === "") return 0;

  const cleaned = String(value)
    .replace(/[৳,$\s]/g, "")
    .replace(/,/g, "");

  const result = Number(cleaned);
  return Number.isFinite(result) ? result : 0;
}

function columnLetter_(column) {
  let result = "";
  while (column > 0) {
    const rem = (column - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    column = Math.floor((column - 1) / 26);
  }
  return result;
}

function exportFinancialDashboardPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet.getName().startsWith(FS_CONFIG.dashboardPrefix)) {
    SpreadsheetApp.getUi().alert(
      "First open a generated FS - dashboard sheet."
    );
    return;
  }

  const url = "https://docs.google.com/spreadsheets/d/" +
    ss.getId() + "/export?format=pdf" +
    "&gid=" + sheet.getSheetId() +
    "&portrait=false&fitw=true&sheetnames=false" +
    "&printtitle=false&pagenumbers=false&gridlines=false" +
    "&fzr=false";

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    }
  });

  const file = DriveApp.createFile(
    response.getBlob().setName(sheet.getName() + ".pdf")
  );

  SpreadsheetApp.getUi().alert(
    "PDF created in Google Drive:\n" + file.getUrl()
  );
}
/************************************************************
 * FINANCIAL STATEMENT PRO UPGRADE
 * Add this code at the VERY BOTTOM of your existing Code.gs
 * Existing functions are not replaced.
 ************************************************************/

const FS_PRO = {
  PREFIX: 'FS PRO - ',
  HEADER_COLOR: '#142D4E',
  ACCENT_COLOR: '#2878D0',
  LIGHT_COLOR: '#EEF3F8',
  TEXT_COLOR: '#243247',
  SUCCESS_COLOR: '#DDF5E5',
  WARNING_COLOR: '#FFF0D5'
};

/**
 * MAIN FUNCTION
 * Run this function manually from Apps Script.
 */
function runFinancialStatementUpgrade() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Google Spreadsheet open kore tarpor run korun.');
  }

  const sourceSheets = ss.getSheets().filter(function(sheet) {
    const name = sheet.getName();
    return !name.startsWith(FS_PRO.PREFIX) &&
           !name.startsWith(FS_CONFIG.dashboardPrefix);
  });

  if (!sourceSheets.length) {
    throw new Error('Source data sheet paoa jayni.');
  }

  const report = getOrCreateFSProSheet_(ss, 'Dashboard');
  const audit = getOrCreateFSProSheet_(ss, 'Data Audit');
  const data = getOrCreateFSProSheet_(ss, 'Data Snapshot');

  buildFSProDashboard_(report, sourceSheets);
  buildFSProAudit_(audit, sourceSheets);
  buildFSProSnapshot_(data, sourceSheets);

  SpreadsheetApp.flush();

  ss.toast(
    'Upgrade complete! Dashboard, Data Audit and Data Snapshot check korun.',
    'Financial Statement Pro',
    8
  );
}

/**
 * Creates a new report sheet without modifying source sheets.
 */
function getOrCreateFSProSheet_(ss, suffix) {
  const name = FS_PRO.PREFIX + suffix;
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  sheet.clear();
  sheet.setHiddenGridlines(true);
  return sheet;
}

/**
 * DASHBOARD
 */
function buildFSProDashboard_(sheet, sourceSheets) {
  sheet.setTabColor(FS_PRO.HEADER_COLOR);
  sheet.setColumnWidths(1, 8, 135);

  sheet.getRange('A1:H2')
    .merge()
    .setValue('FINANCIAL STATEMENT | PRO DASHBOARD')
    .setBackground(FS_PRO.HEADER_COLOR)
    .setFontColor('#FFFFFF')
    .setFontSize(17)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('A3:H3')
    .merge()
    .setValue('Generated: ' + new Date().toLocaleString())
    .setFontColor('#64748B')
    .setHorizontalAlignment('center');

  const labels = [
    ['A5:B5', 'SOURCE SHEETS'],
    ['C5:D5', 'NON-EMPTY CELLS'],
    ['E5:F5', 'FORMULA CELLS'],
    ['G5:H5', 'DATA ROWS']
  ];

  labels.forEach(function(item) {
    sheet.getRange(item[0])
      .merge()
      .setValue(item[1])
      .setBackground(FS_PRO.LIGHT_COLOR)
      .setFontColor(FS_PRO.HEADER_COLOR)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  let nonEmpty = 0;
  let formulas = 0;
  let dataRows = 0;

  sourceSheets.forEach(function(source) {
    const range = source.getDataRange();
    const values = range.getDisplayValues();
    const formulaValues = range.getFormulas();

    values.forEach(function(row, r) {
      let rowHasData = false;

      row.forEach(function(value, c) {
        if (String(value).trim() !== '') {
          nonEmpty++;
          rowHasData = true;
        }

        if (formulaValues[r] && formulaValues[r][c]) {
          formulas++;
        }
      });

      if (rowHasData) dataRows++;
    });
  });

  const metrics = [
    sourceSheets.length,
    nonEmpty,
    formulas,
    dataRows
  ];

  const metricRanges = ['A6:B7', 'C6:D7', 'E6:F7', 'G6:H7'];

  metricRanges.forEach(function(range, i) {
    sheet.getRange(range)
      .merge()
      .setValue(metrics[i])
      .setBackground('#FFFFFF')
      .setFontColor(FS_PRO.ACCENT_COLOR)
      .setFontSize(19)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, '#D5DEE9',
        SpreadsheetApp.BorderStyle.SOLID);
  });

  sheet.getRange('A9:H9')
    .merge()
    .setValue('SOURCE SHEET REGISTER')
    .setBackground(FS_PRO.HEADER_COLOR)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  sheet.getRange('A10:D10').setValues([[
    'Sheet Name', 'Rows', 'Columns', 'Last Row'
  ]]);

  sheet.getRange('A10:D10')
    .setBackground(FS_PRO.LIGHT_COLOR)
    .setFontWeight('bold');

  const rows = sourceSheets.map(function(source) {
    return [
      source.getName(),
      source.getDataRange().getNumRows(),
      source.getDataRange().getNumColumns(),
      source.getLastRow()
    ];
  });

  if (rows.length) {
    sheet.getRange(11, 1, rows.length, 4).setValues(rows);
  }

  sheet.getRange('F10:H10').merge().setValue('QUICK ACTIONS');
  sheet.getRange('F10:H10')
    .setBackground(FS_PRO.LIGHT_COLOR)
    .setFontWeight('bold');

  sheet.getRange('F11:H11').merge().setValue('1. Check Data Audit');
  sheet.getRange('F12:H12').merge().setValue('2. Review Data Snapshot');
  sheet.getRange('F13:H13').merge().setValue('3. Export Dashboard PDF');

  sheet.getRange('F15:H15')
    .merge()
    .setValue('Source sheets are read-only in this upgrade.');

  sheet.getRange('F15:H15')
    .setWrap(true)
    .setFontColor('#64748B');

  sheet.setFrozenRows(3);
  sheet.getDataRange().setVerticalAlignment('middle');
  sheet.getDataRange().setWrap(true);
}

/**
 * DATA AUDIT
 * Finds where data and formulas actually exist.
 */
function buildFSProAudit_(sheet, sourceSheets) {
  sheet.setTabColor('#D97706');
  sheet.setColumnWidths(1, 7, 145);

  sheet.getRange('A1:G2')
    .merge()
    .setValue('SOURCE DATA AUDIT')
    .setBackground(FS_PRO.HEADER_COLOR)
    .setFontColor('#FFFFFF')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const headers = [[
    'Source Sheet',
    'Cell / Range',
    'Displayed Value',
    'Formula',
    'Row',
    'Column',
    'Type'
  ]];

  sheet.getRange(4, 1, 1, 7)
    .setValues(headers)
    .setBackground(FS_PRO.LIGHT_COLOR)
    .setFontWeight('bold');

  const output = [];

  sourceSheets.forEach(function(source) {
    const range = source.getDataRange();
    const values = range.getDisplayValues();
    const formulas = range.getFormulas();

    values.forEach(function(row, r) {
      row.forEach(function(value, c) {
        const formula = formulas[r][c];

        if (String(value).trim() !== '' || formula) {
          const cell = source.getRange(r + 1, c + 1);

          let type = 'TEXT';

          if (formula) {
            type = 'FORMULA';
          } else if (
            typeof cell.getValue() === 'number' &&
            String(value).trim() !== ''
          ) {
            type = 'NUMBER';
          } else if (String(value).trim() === '') {
            type = 'EMPTY DISPLAY';
          }

          output.push([
            source.getName(),
            cell.getA1Notation(),
            value,
            formula,
            r + 1,
            c + 1,
            type
          ]);
        }
      });
    });
  });

  if (output.length) {
    sheet.getRange(5, 1, output.length, 7).setValues(output);
  } else {
    sheet.getRange('A5').setValue('No data found in source sheets.');
  }

  sheet.setFrozenRows(4);
  sheet.getDataRange().setWrap(true);
  sheet.getDataRange().setVerticalAlignment('middle');
  sheet.getRange('C:C').setNumberFormat('@');
}

/**
 * SNAPSHOT
 * Copies displayed source values into a separate review sheet.
 * Original sheets are not changed.
 */
function buildFSProSnapshot_(sheet, sourceSheets) {
  sheet.setTabColor('#2878D0');
  sheet.setColumnWidths(1, 6, 150);

  sheet.getRange('A1:F2')
    .merge()
    .setValue('SOURCE DATA SNAPSHOT')
    .setBackground(FS_PRO.HEADER_COLOR)
    .setFontColor('#FFFFFF')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  let nextRow = 4;

  sourceSheets.forEach(function(source) {
    const values = source.getDataRange().getDisplayValues();

    sheet.getRange(nextRow, 1, 1, 3)
      .merge()
      .setValue('SOURCE: ' + source.getName())
      .setBackground(FS_PRO.LIGHT_COLOR)
      .setFontColor(FS_PRO.HEADER_COLOR)
      .setFontWeight('bold');

    nextRow++;

    if (values.length && values[0].length) {
      const width = Math.min(
        values.reduce(function(max, row) {
          return Math.max(max, row.length);
        }, 1),
        26
      );

      const normalized = values.map(function(row) {
        return row.slice(0, width);
      });

      sheet.getRange(nextRow, 1, normalized.length, width)
        .setNumberFormat('@')
        .setValues(normalized);

      nextRow += normalized.length + 2;
    }
  });

  sheet.setFrozenRows(2);
  sheet.getDataRange().setWrap(true);
  sheet.getDataRange().setVerticalAlignment('middle');
}

/**
 * EXPORT THE PRO DASHBOARD AS PDF
 */
function exportFSProDashboardPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FS_PRO.PREFIX + 'Dashboard');

  if (!sheet) {
    throw new Error(
      'First run runFinancialStatementUpgrade function.'
    );
  }

  const url =
    'https://docs.google.com/spreadsheets/d/' +
    ss.getId() +
    '/export?format=pdf' +
    '&gid=' + sheet.getSheetId() +
    '&size=A4' +
    '&portrait=false' +
    '&fitw=true' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=true' +
    '&gridlines=false' +
    '&fzr=false';

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'PDF export failed. Permission check korun. HTTP ' +
      response.getResponseCode()
    );
  }

  const file = DriveApp.createFile(
    response.getBlob().setName(
      'Financial_Statement_Dashboard_' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd_HH-mm'
      ) +
      '.pdf'
    )
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'PDF created: ' + file.getName(),
    'Financial Statement Pro',
    8
  );

  Logger.log('PDF URL: ' + file.getUrl());
  return file.getUrl();
}

/**
 * Adds a custom menu when the spreadsheet opens.
 * If your project already has onOpen(), this function intentionally
 * uses a different name. Run installFSProMenu once to install trigger.
 */
function installFSProMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onOpenFSPro') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onOpenFSPro')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  onOpenFSPro();

  ss.toast('Financial Statement Pro menu installed.', 'Success', 5);
}

function onOpenFSPro() {
  SpreadsheetApp.getUi()
    .createMenu('FINANCIAL STATEMENT PRO')
    .addItem('Build / Refresh Dashboard', 'runFinancialStatementUpgrade')
    .addItem('Export Dashboard PDF', 'exportFSProDashboardPDF')
    .addSeparator()
    .addItem('Reinstall Menu', 'installFSProMenu')
    .addToUi();
}
