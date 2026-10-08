/**
 * Rizvi Task Hub
 * Personal productivity + money management system for Google Sheets.
 *
 * Setup:
 * 1) Create a blank Google Sheet.
 * 2) Extensions > Apps Script.
 * 3) Replace Code.gs with this file.
 * 4) Add a new HTML file named Sidebar and paste Sidebar.html.
 * 5) Run setupRizviTaskHub() once and authorize.
 * 6) Use Rizvi Hub > Install/Refresh Automations.
 */

const APP = {
  NAME: 'Rizvi Task Hub',
  VERSION: '1.0.0',
  TZ: 'Asia/Dhaka',
  CURRENCY: '৳',
  MONTHLY_BUDGET: 3000,
  BIG_TASK_LIMIT: 3,
  SMALL_TASK_RECOMMENDATION: 10,
  WORK_START: '09:00',
  WORK_END: '19:00',
  MORNING_PLANNING: '09:00',
  DAILY_REVIEW: '19:05',
  WEEKLY_REVIEW_DAY: 'FRIDAY',
  SHEETS: {
    DASHBOARD: 'Dashboard',
    TASKS: 'Tasks',
    EXPENSES: 'Expenses',
    INCOME: 'Income',
    GOALS: 'Goals',
    MONTHLY_GOALS: 'Monthly Goals',
    DAILY_REVIEW: 'Daily Review',
    WEEKLY_REVIEW: 'Weekly Review',
    PROJECTS: 'Projects',
    SETTINGS: 'Settings',
    LOGS: 'Logs',
    ANALYTICS: 'Analytics'
  },
  PROJECTS: ['Personal', 'DiaFit', 'Tenova', 'Divine Cloth', 'Quicon'],
  EXPENSE_CATEGORIES: ['Food', 'Transport', 'Bills'],
  PRIORITIES: ['Critical', 'High', 'Medium', 'Low'],
  TASK_TYPES: ['BIG', 'SMALL'],
  TASK_STATUSES: ['Backlog', 'Planned', 'In Progress', 'Blocked', 'Done', 'Cancelled'],
  NEED_WANT: ['Need', 'Want'],
  PAYMENT_METHODS: ['Cash', 'bKash', 'Nagad', 'Card', 'Bank', 'Other']
};

const COLORS = {
  navy: '#172554',
  blue: '#2563EB',
  sky: '#E0F2FE',
  green: '#16A34A',
  greenLight: '#DCFCE7',
  yellow: '#CA8A04',
  yellowLight: '#FEF9C3',
  red: '#DC2626',
  redLight: '#FEE2E2',
  purple: '#7C3AED',
  purpleLight: '#EDE9FE',
  gray: '#64748B',
  light: '#F8FAFC',
  white: '#FFFFFF',
  border: '#E2E8F0',
  black: '#0F172A'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Rizvi Hub')
    .addItem('🏠 Open Dashboard', 'showDashboard')
    .addItem('➕ New Task', 'showSidebar')
    .addItem('🔥 Plan Today\'s Big 3', 'planTodaysBig3')
    .addItem('💰 Add Expense', 'showExpenseSidebar')
    .addItem('💵 Add Income', 'showIncomeSidebar')
    .addItem('⏱ Focus Timer', 'showFocusTimer')
    .addSeparator()
    .addItem('📅 Sync Tasks to Calendar', 'syncAllCalendarTasks')
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addItem('🌙 Daily Review', 'showDailyReview')
    .addItem('📊 Weekly Review', 'generateWeeklyReview')
    .addSeparator()
    .addItem('⚙️ Install/Refresh Automations', 'installAutomations')
    .addItem('🧰 Repair Formatting & Validations', 'repairHub')
    .addToUi();
}

function setupRizviTaskHub() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(APP.TZ);
  ensureSheets_();
  setupSettings_();
  setupProjects_();
  setupTasks_();
  setupExpenses_();
  setupIncome_();
  setupGoals_();
  setupMonthlyGoals_();
  setupDailyReview_();
  setupWeeklyReview_();
  setupLogs_();
  setupAnalytics_();
  setupDashboard_();
  applyValidations_();
  applyFormatting_();
  createNamedRanges_();
  installAutomations();
  refreshDashboard();
  ss.setActiveSheet(ss.getSheetByName(APP.SHEETS.DASHBOARD));
  SpreadsheetApp.getUi().alert('Rizvi Task Hub is ready. Open the Dashboard to start.');
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActive();
  Object.values(APP.SHEETS).forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
}

function setupSettings_() {
  const sh = getSheet_('SETTINGS');
  sh.clear();
  sh.getRange('A1:D1').merge().setValue(APP.NAME + ' — Settings');
  sh.getRange('A3:B16').setValues([
    ['Setting', 'Value'],
    ['Hub Name', APP.NAME],
    ['Currency', 'BDT'],
    ['Monthly Budget', APP.MONTHLY_BUDGET],
    ['Big Task Daily Limit', APP.BIG_TASK_LIMIT],
    ['Recommended Small Task Limit', APP.SMALL_TASK_RECOMMENDATION],
    ['Work Start', APP.WORK_START],
    ['Work End', APP.WORK_END],
    ['Morning Planning', APP.MORNING_PLANNING],
    ['Daily Review', APP.DAILY_REVIEW],
    ['Weekly Review Day', APP.WEEKLY_REVIEW_DAY],
    ['Reminder Email', Session.getEffectiveUser().getEmail() || ''],
    ['Time Zone', APP.TZ],
    ['Calendar Name', 'Rizvi Task Hub']
  ]);
  sh.getRange('D3:D12').setValues([
    ['Expense Categories'],
    ...APP.EXPENSE_CATEGORIES.map(x => [x]),
    ['Projects'],
    ...APP.PROJECTS.map(x => [x])
  ]);
}

function setupProjects_() {
  const sh = getSheet_('PROJECTS');
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,6).setValues([['Project ID','Project','Color','Status','Notes','Created At']]);
    sh.getRange(2,1,APP.PROJECTS.length,6).setValues(APP.PROJECTS.map((p,i)=>[
      'PRJ-' + String(i+1).padStart(3,'0'), p, '', 'Active', '', new Date()
    ]));
  }
}

function setupTasks_() {
  const sh = getSheet_('TASKS');
  const headers = ['Task ID','Task','Project','Type','Priority','Status','Due Date','Start Time','End Time','Est. Minutes','Actual Minutes','Notes','Calendar Event ID','Created At','Completed At','Score'];
  ensureHeader_(sh, headers);
  if (sh.getLastRow() === 1) {
    sh.getRange(2,1,1,headers.length).clearContent();
  }
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('G:G').setNumberFormat('dd-mmm-yyyy');
  sh.getRange('H:I').setNumberFormat('hh:mm AM/PM');
  sh.getRange('J:K').setNumberFormat('0');
  sh.getRange('N:O').setNumberFormat('dd-mmm-yyyy hh:mm');
}

function setupExpenses_() {
  const sh = getSheet_('EXPENSES');
  ensureHeader_(sh, ['Expense ID','Date','Category','Description','Amount','Payment Method','Need/Want','Project','Month','Notes','Created At']);
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.getRange('E:E').setNumberFormat('৳#,##0.00');
  sh.getRange('I:I').setNumberFormat('@');
  if (sh.getLastRow() > 1) sh.getRange(2,9,sh.getLastRow()-1,1).setFormulaR1C1('=TEXT(RC[-7],"yyyy-mm")');
}

function setupIncome_() {
  const sh = getSheet_('INCOME');
  ensureHeader_(sh, ['Income ID','Date','Source','Description','Amount','Month','Notes','Created At']);
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.getRange('E:E').setNumberFormat('৳#,##0.00');
  if (sh.getLastRow() > 1) sh.getRange(2,6,sh.getLastRow()-1,1).setFormulaR1C1('=TEXT(RC[-4],"yyyy-mm")');
}

function setupGoals_() {
  const sh = getSheet_('GOALS');
  ensureHeader_(sh, ['Goal ID','Goal','Type','Target','Current','Unit','Deadline','Status','Progress %','Notes','Created At']);
  sh.setFrozenRows(1);
}

function setupMonthlyGoals_() {
  const sh = getSheet_('MONTHLY_GOALS');
  ensureHeader_(sh, ['Goal ID','Month','Goal','Target','Current','Unit','Status','Progress %','Notes','Created At']);
  sh.setFrozenRows(1);
}

function setupDailyReview_() {
  const sh = getSheet_('DAILY_REVIEW');
  ensureHeader_(sh, ['Date','Big Wins','Big Completed','Small Completed','Focus Minutes','Productivity %','What Went Well','What Did Not Get Done','Tomorrow #1','Notes']);
  sh.setFrozenRows(1);
}

function setupWeeklyReview_() {
  const sh = getSheet_('WEEKLY_REVIEW');
  ensureHeader_(sh, ['Week Ending','Tasks Completed','Big Tasks Completed','Total Tasks','Focus Minutes','Productivity %','Total Income','Total Expenses','Savings','Best Project','Notes']);
  sh.setFrozenRows(1);
}

function setupLogs_() {
  const sh = getSheet_('LOGS');
  ensureHeader_(sh, ['Timestamp','Action','Entity','ID','Details']);
  sh.setFrozenRows(1);
}

function setupAnalytics_() {
  const sh = getSheet_('ANALYTICS');
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange('A1:H1').merge().setValue('📊 RIZVI HUB — ANALYTICS');
  sh.getRange('A3:D3').setValues([['Expense Category','Amount','Share %','Month']]);
  sh.getRange('F3:H3').setValues([['Week','Productivity %','Tasks Completed']]);
  sh.getRange('A1:H1').setBackground(COLORS.navy).setFontColor(COLORS.white).setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A3:D3').setBackground(COLORS.blue).setFontColor(COLORS.white).setFontWeight('bold');
  sh.getRange('F3:H3').setBackground(COLORS.purple).setFontColor(COLORS.white).setFontWeight('bold');
}

function refreshAnalytics_() {
  const sh = getSheet_('ANALYTICS');
  sh.getRange('A4:D20').clearContent();
  sh.getRange('F4:H20').clearContent();
  const month = monthKey_(new Date());
  const exp = getSheet_('EXPENSES');
  const totals = {};
  if (exp.getLastRow() >= 2) exp.getRange(2,1,exp.getLastRow()-1,11).getValues().forEach(r=>{ if(String(r[8])===month) totals[r[2]]=(totals[r[2]]||0)+Number(r[4]||0); });
  const total = Object.values(totals).reduce((a,b)=>a+b,0);
  APP.EXPENSE_CATEGORIES.forEach((c,i)=>sh.getRange(4+i,1,1,4).setValues([[c,totals[c]||0,total?((totals[c]||0)/total):0,month]]));
  sh.getRange(4,2,APP.EXPENSE_CATEGORIES.length,1).setNumberFormat('৳#,##0.00');
  sh.getRange(4,3,APP.EXPENSE_CATEGORIES.length,1).setNumberFormat('0%');
  const wr=getSheet_('WEEKLY_REVIEW');
  if(wr.getLastRow()>=2){
    const rows=wr.getRange(2,1,Math.min(8,wr.getLastRow()-1),11).getValues();
    rows.forEach((r,i)=>sh.getRange(4+i,6,1,3).setValues([[dateOnly_(new Date(r[0])),Number(r[5]||0),Number(r[1]||0)]]));
    sh.getRange(4,7,Math.max(1,rows.length),1).setNumberFormat('0%');
  }
  sh.getCharts().forEach(c=>sh.removeChart(c));
  const catChart=sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(sh.getRange('A3:B6')).setPosition(10,1,0,0).setOption('title','Current Month Expenses by Category').setOption('legend',{position:'none'}).build();
  sh.insertChart(catChart);
  const prodChart=sh.newChart().setChartType(Charts.ChartType.LINE).addRange(sh.getRange('F3:G12')).setPosition(10,6,0,0).setOption('title','Weekly Productivity').setOption('vAxis',{format:'percent',minValue:0,maxValue:1}).build();
  sh.insertChart(prodChart);
}

function setupDashboard_() {
  const sh = getSheet_('DASHBOARD');
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange('A1:L1').merge().setValue('🚀 RIZVI TASK HUB');
  sh.getRange('A2:L2').merge().setValue('Personal Productivity + Money Control Center');
  sh.getRange('A4:L4').merge().setValue('TODAY');
  sh.getRange('A5:D5').merge().setValue('🔥 BIG 3');
  sh.getRange('F5:H5').merge().setValue('⚡ QUICK TASKS');
  sh.getRange('J5:L5').merge().setValue('💰 MONEY');
  sh.getRange('A10:D10').merge().setValue('📊 PRODUCTIVITY');
  sh.getRange('F10:H10').merge().setValue('🎯 GOALS');
  sh.getRange('J10:L10').merge().setValue('📅 UPCOMING');
  sh.getRange('A18:L18').merge().setValue('Use Rizvi Hub menu or the sidebar to add and manage items.');
  for (let c=1;c<=12;c++) sh.setColumnWidth(c, 120);
  sh.setColumnWidth(2, 240);
  sh.setColumnWidth(7, 240);
  sh.setColumnWidth(11, 180);
  sh.setRowHeight(1, 42);
  sh.setRowHeight(2, 28);
  sh.getRange('A1:L1').setBackground(COLORS.navy).setFontColor(COLORS.white).setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A2:L2').setBackground(COLORS.sky).setFontColor(COLORS.navy).setFontSize(11).setHorizontalAlignment('center');
}

function ensureHeader_(sh, headers) {
  if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.getRange(1,1,1,headers.length).setBackground(COLORS.navy).setFontColor(COLORS.white).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
}

function applyValidations_() {
  const ss = SpreadsheetApp.getActive();
  const task = ss.getSheetByName(APP.SHEETS.TASKS);
  setValidation_(task.getRange('C2:C'), APP.PROJECTS);
  setValidation_(task.getRange('D2:D'), APP.TASK_TYPES);
  setValidation_(task.getRange('E2:E'), APP.PRIORITIES);
  setValidation_(task.getRange('F2:F'), APP.TASK_STATUSES);
  const exp = ss.getSheetByName(APP.SHEETS.EXPENSES);
  setValidation_(exp.getRange('C2:C'), APP.EXPENSE_CATEGORIES);
  setValidation_(exp.getRange('F2:F'), APP.PAYMENT_METHODS);
  setValidation_(exp.getRange('G2:G'), APP.NEED_WANT);
  setValidation_(exp.getRange('H2:H'), APP.PROJECTS);
  const goals = ss.getSheetByName(APP.SHEETS.GOALS);
  setValidation_(goals.getRange('H2:H'), ['Active','Completed','Paused','Cancelled']);
  const mg = ss.getSheetByName(APP.SHEETS.MONTHLY_GOALS);
  setValidation_(mg.getRange('G2:G'), ['Active','Completed','Paused','Cancelled']);
}

function setValidation_(range, values) {
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build());
}

function applyFormatting_() {
  const ss = SpreadsheetApp.getActive();
  Object.values(APP.SHEETS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    if (name !== APP.SHEETS.DASHBOARD) {
      sh.getDataRange().setFontFamily('Arial').setVerticalAlignment('middle');
      sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).setBackground(COLORS.navy).setFontColor(COLORS.white).setFontWeight('bold');
      sh.autoResizeColumns(1, Math.min(sh.getLastColumn(), 12));
    }
  });
  applyConditionalFormatting_();
}

function applyConditionalFormatting_() {
  const sh = getSheet_('TASKS');
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Done').setBackground(COLORS.greenLight).setFontColor(COLORS.green).setRanges([sh.getRange('F2:F')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('In Progress').setBackground(COLORS.sky).setRanges([sh.getRange('F2:F')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Blocked').setBackground(COLORS.redLight).setFontColor(COLORS.red).setRanges([sh.getRange('F2:F')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Critical').setBackground(COLORS.redLight).setFontColor(COLORS.red).setRanges([sh.getRange('E2:E')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('High').setBackground('#FFEDD5').setRanges([sh.getRange('E2:E')]).build());
  sh.setConditionalFormatRules(rules);
}

function createNamedRanges_() {
  const ss = SpreadsheetApp.getActive();
  ss.getNamedRanges().filter(n => n.getName().startsWith('RIZVI_')).forEach(n => n.remove());
  ss.setNamedRange('RIZVI_MONTHLY_BUDGET', getSheet_('SETTINGS').getRange('B6'));
}

function getSheet_(key) {
  const name = APP.SHEETS[key] || key;
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

function getSettings_() {
  const sh = getSheet_('SETTINGS');
  const values = sh.getRange('A4:B16').getValues();
  const out = {};
  values.forEach(r => { if (r[0]) out[String(r[0])] = r[1]; });
  return out;
}

function getEmail_() {
  return getSettings_()['Reminder Email'] || Session.getEffectiveUser().getEmail() || '';
}

function todayKey_() { return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd'); }
function monthKey_(d) { return Utilities.formatDate(d || new Date(), APP.TZ, 'yyyy-MM'); }
function dateOnly_(d) { return Utilities.formatDate(d || new Date(), APP.TZ, 'dd-MMM-yyyy'); }
function parseTime_(s) {
  const parts = String(s).split(':');
  const d = new Date();
  d.setHours(Number(parts[0] || 0), Number(parts[1] || 0), 0, 0);
  return d;
}

function nextId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random()*900+100);
}

function addTask(data) {
  data = data || {};
  const task = String(data.task || '').trim();
  if (!task) throw new Error('Task name is required.');
  const type = data.type === 'BIG' ? 'BIG' : 'SMALL';
  const due = data.dueDate ? new Date(data.dueDate) : new Date();
  const start = data.startTime || '';
  const end = data.endTime || '';
  const priority = data.priority || 'Medium';
  const project = data.project || 'Personal';
  const status = data.status || (type === 'BIG' ? 'Planned' : 'Backlog');
  const score = calculateScore_(priority, type, data.impact, data.urgency, data.revenue);
  if (type === 'BIG' && countBigTasksForDate_(due) >= APP.BIG_TASK_LIMIT) {
    throw new Error('Today already has 3 Big Tasks. Move this task to another date or make it SMALL.');
  }
  const sh = getSheet_('TASKS');
  const id = nextId_('TSK');
  sh.appendRow([id, task, project, type, priority, status, due, start, end, Number(data.estMinutes||0), 0, data.notes||'', '', new Date(), '', score]);
  formatTaskRow_(sh.getLastRow());
  log_('CREATE','TASK',id,task);
  refreshDashboard();
  if (data.calendar && start && end) syncTaskToCalendar(id);
  return id;
}

function calculateScore_(priority, type, impact, urgency, revenue) {
  const p = {Critical:5,High:4,Medium:3,Low:1}[priority] || 3;
  const t = type === 'BIG' ? 2 : 0;
  return p + t + Number(impact||0) + Number(urgency||0) + Number(revenue||0);
}

function countBigTasksForDate_(date) {
  const sh = getSheet_('TASKS');
  if (sh.getLastRow() < 2) return 0;
  const vals = sh.getRange(2,1,sh.getLastRow()-1,16).getValues();
  const key = Utilities.formatDate(new Date(date), APP.TZ, 'yyyy-MM-dd');
  return vals.filter(r => r[3] === 'BIG' && r[5] !== 'Cancelled' && r[6] && Utilities.formatDate(new Date(r[6]), APP.TZ, 'yyyy-MM-dd') === key).length;
}

function formatTaskRow_(row) {
  const sh = getSheet_('TASKS');
  sh.getRange(row,7).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(row,8,1,2).setNumberFormat('hh:mm AM/PM');
  sh.getRange(row,10,1,2).setNumberFormat('0');
  sh.getRange(row,14,1,2).setNumberFormat('dd-mmm-yyyy hh:mm');
}

function updateTaskStatus(taskId, status) {
  const sh = getSheet_('TASKS');
  const row = findRowById_(sh, taskId, 1);
  if (!row) throw new Error('Task not found: ' + taskId);
  sh.getRange(row,6).setValue(status);
  if (status === 'Done') sh.getRange(row,15).setValue(new Date());
  else sh.getRange(row,15).clearContent();
  log_('STATUS','TASK',taskId,status);
  refreshDashboard();
}

function planTodaysBig3() {
  const sh = getSheet_('TASKS');
  const today = new Date();
  const key = Utilities.formatDate(today, APP.TZ, 'yyyy-MM-dd');
  const data = sh.getLastRow() < 2 ? [] : sh.getRange(2,1,sh.getLastRow()-1,16).getValues();
  const candidates = data.map((r,i)=>({row:i+2,score:Number(r[15]||0),type:r[3],status:r[5],due:r[6],task:r[1]}))
    .filter(x => x.type === 'BIG' && x.status !== 'Done' && x.status !== 'Cancelled' && x.due && Utilities.formatDate(new Date(x.due),APP.TZ,'yyyy-MM-dd') === key)
    .sort((a,b)=>b.score-a.score);
  if (candidates.length === 0) {
    SpreadsheetApp.getUi().alert('No Big Tasks are scheduled for today. Add up to 3 Big Tasks first.');
    return;
  }
  candidates.slice(0,3).forEach((x,i)=>sh.getRange(x.row,6).setValue(i===0?'In Progress':'Planned'));
  refreshDashboard();
  SpreadsheetApp.getUi().alert('Today\'s Big 3 are prioritized by task score.');
}

function addExpense(data) {
  data = data || {};
  const amount = Number(data.amount);
  if (!(amount > 0)) throw new Error('Expense amount must be greater than 0.');
  const date = data.date ? new Date(data.date) : new Date();
  const month = monthKey_(date);
  const budget = getMonthlyBudget_(month);
  const spent = getMonthlyExpenses_(month);
  const remaining = budget - spent;
  const over = amount > remaining;
  if (over && !data.override) {
    return {ok:false, warning:true, message:'This expense will exceed the monthly budget by ' + money_(amount-remaining) + '.', remaining:remaining, budget:budget, spent:spent};
  }
  const sh = getSheet_('EXPENSES');
  const id = nextId_('EXP');
  sh.appendRow([id,date,data.category||'Food',data.description||'',amount,data.paymentMethod||'Cash',data.needWant||'Need',data.project||'Personal',month,data.notes||'',new Date()]);
  sh.getRange(sh.getLastRow(),2).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(sh.getLastRow(),5).setNumberFormat('৳#,##0.00');
  log_('CREATE','EXPENSE',id,money_(amount));
  refreshDashboard();
  return {ok:true,id:id,warning:over,message:over?'Budget exceeded, but override was accepted.':'Expense saved.'};
}

function addIncome(data) {
  data = data || {};
  const amount = Number(data.amount);
  if (!(amount > 0)) throw new Error('Income amount must be greater than 0.');
  const date = data.date ? new Date(data.date) : new Date();
  const sh = getSheet_('INCOME');
  const id = nextId_('INC');
  sh.appendRow([id,date,data.source||'Other',data.description||'',amount,monthKey_(date),data.notes||'',new Date()]);
  sh.getRange(sh.getLastRow(),2).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(sh.getLastRow(),5).setNumberFormat('৳#,##0.00');
  log_('CREATE','INCOME',id,money_(amount));
  refreshDashboard();
  return id;
}

function getMonthlyBudget_(month) {
  const sh = getSheet_('SETTINGS');
  const v = Number(sh.getRange('B6').getValue());
  return v || APP.MONTHLY_BUDGET;
}

function getMonthlyExpenses_(month) {
  const sh = getSheet_('EXPENSES');
  if (sh.getLastRow() < 2) return 0;
  return sh.getRange(2,5,sh.getLastRow()-1,5).getValues().reduce((sum,r)=>sum + (String(r[4])===month ? Number(r[0]) : 0),0);
}

function getMonthlyIncome_(month) {
  const sh = getSheet_('INCOME');
  if (sh.getLastRow() < 2) return 0;
  return sh.getRange(2,5,sh.getLastRow()-1,2).getValues().reduce((sum,r)=>sum + (String(r[1])===month ? Number(r[0]) : 0),0);
}

function money_(n) { return APP.CURRENCY + Number(n||0).toLocaleString('en-BD',{minimumFractionDigits:2,maximumFractionDigits:2}); }

function getDashboardData() {
  const now = new Date();
  const month = monthKey_(now);
  const budget = getMonthlyBudget_(month);
  const spent = getMonthlyExpenses_(month);
  const income = getMonthlyIncome_(month);
  const daysInMonth = Number(Utilities.formatDate(new Date(now.getFullYear(), now.getMonth()+1,0),APP.TZ,'dd'));
  const todayDay = Number(Utilities.formatDate(now,APP.TZ,'dd'));
  const remainingDays = Math.max(1,daysInMonth-todayDay+1);
  const remaining = budget-spent;
  const safeDaily = Math.max(0,remaining/remainingDays);
  const big3 = getTasksForToday_('BIG').slice(0,3);
  const small = getTasksForToday_('SMALL').filter(x=>x.status!=='Done'&&x.status!=='Cancelled').slice(0,10);
  const allToday = getTasksForToday_();
  const done = allToday.filter(x=>x.status==='Done').length;
  const productivity = allToday.length ? Math.round(done/allToday.length*100) : 0;
  const goals = getActiveGoals_().slice(0,3);
  const upcoming = getUpcomingTasks_().slice(0,5);
  return {date:dateOnly_(now),month,budget,spent,income,remaining,safeDaily,big3,small,productivity,done,totalToday:allToday.length,goals,upcoming,budgetHealth:getBudgetHealth_(budget,spent,remaining,daysInMonth,todayDay)};
}

function getTasksForToday_(type) {
  const sh = getSheet_('TASKS');
  if (sh.getLastRow() < 2) return [];
  const key = todayKey_();
  return sh.getRange(2,1,sh.getLastRow()-1,16).getValues().map((r,i)=>({
    row:i+2,id:r[0],task:r[1],project:r[2],type:r[3],priority:r[4],status:r[5],due:r[6],start:r[7],end:r[8],est:r[9],score:r[15]
  })).filter(x=>x.due && Utilities.formatDate(new Date(x.due),APP.TZ,'yyyy-MM-dd')===key && (!type || x.type===type)).sort((a,b)=>Number(b.score||0)-Number(a.score||0));
}

function getUpcomingTasks_() {
  const sh = getSheet_('TASKS');
  if (sh.getLastRow() < 2) return [];
  const today = new Date(); today.setHours(0,0,0,0);
  return sh.getRange(2,1,sh.getLastRow()-1,16).getValues().map(r=>({id:r[0],task:r[1],project:r[2],type:r[3],status:r[5],due:r[6]}))
    .filter(x=>x.due && x.status!=='Done' && x.status!=='Cancelled' && new Date(x.due)>=today)
    .sort((a,b)=>new Date(a.due)-new Date(b.due)).slice(0,5);
}

function getBudgetHealth_(budget,spent,remaining,daysInMonth,todayDay) {
  if (remaining < 0) return {label:'DANGER',color:COLORS.red,detail:'Monthly budget exceeded by '+money_(-remaining)};
  const elapsed = Math.max(1,todayDay);
  const projected = spent / elapsed * daysInMonth;
  if (projected > budget) return {label:'WARNING',color:COLORS.yellow,detail:'Current spending pace may exceed the monthly budget.'};
  return {label:'SAFE',color:COLORS.green,detail:'Spending pace is within the monthly budget.'};
}

function refreshDashboard() {
  const sh = getSheet_('DASHBOARD');
  const d = getDashboardData();
  sh.getRange('A4:L4').setValue('TODAY — ' + d.date).setBackground(COLORS.blue).setFontColor(COLORS.white).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A5:D5').setBackground(COLORS.redLight).setFontColor(COLORS.red).setFontWeight('bold');
  sh.getRange('F5:H5').setBackground(COLORS.sky).setFontColor(COLORS.blue).setFontWeight('bold');
  sh.getRange('J5:L5').setBackground(COLORS.greenLight).setFontColor(COLORS.green).setFontWeight('bold');
  sh.getRange('A6:D8').clearContent();
  d.big3.forEach((x,i)=>sh.getRange(6+i,1,1,4).setValues([[i+1+'️⃣ '+x.task,x.project,x.status,x.priority]]));
  sh.getRange('F6:H15').clearContent();
  d.small.forEach((x,i)=>sh.getRange(6+i,6,1,3).setValues([[x.task,x.project,x.status]]));
  sh.getRange('J6:L9').clearContent();
  sh.getRange('J6:L9').setValues([
    ['Budget',money_(d.budget),''],
    ['Spent',money_(d.spent),''],
    ['Remaining',money_(d.remaining),''],
    ['Safe Daily',money_(d.safeDaily),'']
  ]);
  sh.getRange('L9').setValue(d.budgetHealth.label).setBackground(d.budgetHealth.color).setFontColor(COLORS.white).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A10:D10').setBackground(COLORS.purpleLight).setFontColor(COLORS.purple).setFontWeight('bold');
  sh.getRange('F10:H10').setBackground(COLORS.yellowLight).setFontColor(COLORS.yellow).setFontWeight('bold');
  sh.getRange('J10:L10').setBackground(COLORS.sky).setFontColor(COLORS.blue).setFontWeight('bold');
  sh.getRange('A11:D14').clearContent();
  sh.getRange('A11:D14').setValues([
    ['Productivity',d.productivity+'%','',''],
    ['Today Done',d.done+' / '+d.totalToday,'',''],
    ['Big 3 Done',d.big3.filter(x=>x.status==='Done').length+' / '+d.big3.length,'',''],
    ['Income',money_(d.income),'','']
  ]);
  sh.getRange('F11:H14').clearContent();
  d.goals.forEach((g,i)=>sh.getRange(11+i,6,1,3).setValues([[g.goal,g.progress+'%',g.status]]));
  sh.getRange('J11:L15').clearContent();
  d.upcoming.forEach((x,i)=>sh.getRange(11+i,10,1,3).setValues([[dateOnly_(new Date(x.due)),x.task,x.project]]));
  sh.getRange('A18:L18').setBackground(COLORS.light).setFontColor(COLORS.gray).setHorizontalAlignment('center');
  refreshAnalytics_();
  SpreadsheetApp.flush();
}

function getActiveGoals_() {
  const sh = getSheet_('GOALS');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,11).getValues().map(r=>({goal:r[1],status:r[7],progress:Math.round(Number(r[3]||0)?Number(r[4]||0)/Number(r[3])*100:0)})).filter(x=>x.status==='Active').sort((a,b)=>b.progress-a.progress);
}

function addGoal(data) {
  const sh = getSheet_('GOALS');
  const id = nextId_('GOL');
  sh.appendRow([id,data.goal||'',data.type||'General',Number(data.target||0),Number(data.current||0),data.unit||'',data.deadline?new Date(data.deadline):'',data.status||'Active','',data.notes||'',new Date()]);
  const row=sh.getLastRow();
  sh.getRange(row,9).setFormula('=IFERROR(E'+row+'/D'+row+',0)').setNumberFormat('0%');
  refreshDashboard();
  return id;
}

function addMonthlyGoal(data) {
  const sh = getSheet_('MONTHLY_GOALS');
  const month = data.month || monthKey_(new Date());
  const id = nextId_('MGO');
  sh.appendRow([id,month,data.goal||'',Number(data.target||0),Number(data.current||0),data.unit||'',data.status||'Active','',data.notes||'',new Date()]);
  const row=sh.getLastRow();
  sh.getRange(row,8).setFormula('=IFERROR(E'+row+'/D'+row+',0)').setNumberFormat('0%');
  return id;
}

function addSavingsGoal(data) {
  data = data || {};
  return addGoal({goal:data.goal||'Savings Goal',type:'Savings',target:data.target,current:data.current||0,unit:data.unit||'BDT',deadline:data.deadline,status:'Active',notes:data.notes});
}

function showDashboard() { SpreadsheetApp.getActive().setActiveSheet(getSheet_('DASHBOARD')); refreshDashboard(); }

function showSidebar() { showHtmlSidebar_('Task'); }
function showExpenseSidebar() { showHtmlSidebar_('Expense'); }
function showIncomeSidebar() { showHtmlSidebar_('Income'); }
function showFocusTimer() { showHtmlSidebar_('Timer'); }
function showDailyReview() { showHtmlSidebar_('Review'); }
function showHtmlSidebar_(mode) {
  const t = HtmlService.createTemplateFromFile('Sidebar');
  t.mode = mode;
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('Rizvi Task Hub'));
}

function installAutomations() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('installedEditHandler').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('morningAutomation').timeBased().everyDays(1).atHour(9).nearMinute(0).create();
  ScriptApp.newTrigger('eveningAutomation').timeBased().everyDays(1).atHour(19).nearMinute(5).create();
  ScriptApp.newTrigger('weeklyAutomation').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(18).nearMinute(0).create();
  ScriptApp.newTrigger('refreshDashboard').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Automations installed/refreshed.');
}

function installedEditHandler(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() === APP.SHEETS.TASKS && e.range.getRow() > 1) {
    const row = e.range.getRow();
    const status = sh.getRange(row,6).getValue();
    const type = sh.getRange(row,4).getValue();
    if (status === 'Done' && !sh.getRange(row,15).getValue()) sh.getRange(row,15).setValue(new Date());
    if (type === 'BIG' && status !== 'Cancelled') {
      const due = sh.getRange(row,7).getValue();
      if (due && countBigTasksForDate_(due) > APP.BIG_TASK_LIMIT) {
        sh.getRange(row,6).setValue('Backlog');
        SpreadsheetApp.getActive().toast('Daily Big Task limit is 3. This task was moved back to Backlog.','Rizvi Hub',5);
      }
    }
    if ([7,8,9].some(c=>e.range.getColumn()===c)) {
      const id=sh.getRange(row,1).getValue();
      if (id && sh.getRange(row,8).getValue() && sh.getRange(row,9).getValue()) syncTaskToCalendar(id);
    }
    refreshDashboard();
  }
  if (sh.getName() === APP.SHEETS.EXPENSES && e.range.getRow()>1) refreshDashboard();
  if (sh.getName() === APP.SHEETS.INCOME && e.range.getRow()>1) refreshDashboard();
}

function morningAutomation() {
  refreshDashboard();
  sendEmail_('🌅 Rizvi Task Hub — Morning Plan', buildMorningEmail_());
}

function eveningAutomation() {
  refreshDashboard();
  sendEmail_('🌙 Rizvi Task Hub — Daily Review', buildEveningEmail_());
}

function weeklyAutomation() {
  generateWeeklyReview();
  sendEmail_('📊 Rizvi Task Hub — Weekly Review', buildWeeklyEmail_());
}

function sendEmail_(subject, body) {
  const email = getEmail_();
  if (!email) return;
  MailApp.sendEmail({to:email,subject:subject,body:body});
}

function buildMorningEmail_() {
  const d=getDashboardData();
  const deadlines=getTasksForToday_().filter(x=>x.status!=='Done'&&x.status!=='Cancelled');
  return 'Good morning!\n\nToday: '+d.date+'\n\nBIG 3:\n'+d.big3.map((x,i)=>(i+1)+'. '+x.task+' — '+x.project).join('\n')+'\n\nDEADLINES TODAY:\n'+(deadlines.length?deadlines.map(x=>'- '+x.task+' ['+x.priority+']').join('\n'):'None')+'\n\nBudget remaining: '+money_(d.remaining)+'\nSafe daily spending: '+money_(d.safeDaily)+'\n\nOpen your Rizvi Task Hub Dashboard to plan the day.';
}
function buildEveningEmail_() {
  const d=getDashboardData();
  return 'Daily Review — '+d.date+'\n\nProductivity: '+d.productivity+'%\nBig 3: '+d.big3.filter(x=>x.status==='Done').length+'/'+d.big3.length+' done\nTasks: '+d.done+'/'+d.totalToday+' done\n\nMoney remaining: '+money_(d.remaining)+'\nBudget health: '+d.budgetHealth.label+'\n\nOpen the sheet for your Daily Review.';
}
function buildWeeklyEmail_() {
  const sh=getSheet_('WEEKLY_REVIEW');
  if (sh.getLastRow()<2) return 'No weekly review data yet.';
  const r=sh.getRange(2,1,1,11).getValues()[0];
  return 'Weekly Review\n\nTasks completed: '+r[1]+'\nBig tasks completed: '+r[2]+'\nFocus minutes: '+r[4]+'\nProductivity: '+Math.round(Number(r[5]||0)*100)+'%\nIncome: '+money_(r[6])+'\nExpenses: '+money_(r[7])+'\nSavings: '+money_(r[8]);
}

function syncTaskToCalendar(taskId) {
  const sh=getSheet_('TASKS');
  const row=findRowById_(sh,taskId,1);
  if(!row) throw new Error('Task not found.');
  const r=sh.getRange(row,1,1,16).getValues()[0];
  if(!r[6] || !r[7] || !r[8]) return false;
  const cal=getOrCreateCalendar_();
  const start=combineDateTime_(new Date(r[6]),r[7]);
  const end=combineDateTime_(new Date(r[6]),r[8]);
  let event;
  const oldId=r[12];
  if(oldId){ try { event=cal.getEventById(oldId); } catch(err) {} }
  if(event){ event.setTitle((r[3]==='BIG'?'🔥 ':'')+r[1]+' — '+r[2]); event.setTime(start,end); event.setDescription('Rizvi Task Hub\nTask ID: '+r[0]+'\nPriority: '+r[4]+'\nStatus: '+r[5]); }
  else { event=cal.createEvent((r[3]==='BIG'?'🔥 ':'')+r[1]+' — '+r[2],start,end,{description:'Rizvi Task Hub\nTask ID: '+r[0]+'\nPriority: '+r[4]+'\nStatus: '+r[5]}); event.addPopupReminder(15); }
  sh.getRange(row,13).setValue(event.getId());
  log_('CALENDAR_SYNC','TASK',taskId,event.getId());
  return true;
}

function syncAllCalendarTasks() {
  const sh=getSheet_('TASKS');
  if(sh.getLastRow()<2) return;
  let count=0;
  sh.getRange(2,1,sh.getLastRow()-1,16).getValues().forEach(r=>{if(r[0]&&r[6]&&r[7]&&r[8]){try{if(syncTaskToCalendar(r[0])) count++;}catch(e){log_('ERROR','TASK',r[0],String(e));}}});
  SpreadsheetApp.getUi().alert(count+' task(s) synced to Google Calendar.');
}

function getOrCreateCalendar_() {
  const settings=getSettings_();
  const name=settings['Calendar Name']||APP.NAME;
  const list=CalendarApp.getCalendarsByName(name);
  if(list.length) return list[0];
  return CalendarApp.createCalendar(name,{timeZone:APP.TZ,description:'Rizvi Task Hub task calendar'});
}

function createRecurringCalendarRoutines_() {
  const cal=getOrCreateCalendar_();
  const title1='🌅 Rizvi Hub — Morning Planning';
  const title2='🌙 Rizvi Hub — Daily Review';
  cal.getEventsForDay(new Date()).filter(e=>[title1,title2].includes(e.getTitle())).forEach(e=>e.deleteEvent());
  const d=new Date();
  const s1=combineDateTime_(d,APP.MORNING_PLANNING); const e1=new Date(s1.getTime()+15*60000);
  const s2=combineDateTime_(d,APP.DAILY_REVIEW); const e2=new Date(s2.getTime()+15*60000);
  cal.createEventSeries(title1,s1,e1,CalendarApp.newRecurrence().addDailyRule().until(new Date(d.getFullYear()+1,d.getMonth(),d.getDate())),{description:'Plan today\'s Big 3 and review your dashboard.'});
  cal.createEventSeries(title2,s2,e2,CalendarApp.newRecurrence().addDailyRule().until(new Date(d.getFullYear()+1,d.getMonth(),d.getDate())),{description:'Complete your Rizvi Hub Daily Review.'});
}

function combineDateTime_(date,time) {
  const d=new Date(date); const parts=String(time).split(':');
  d.setHours(Number(parts[0]||0),Number(parts[1]||0),0,0); return d;
}

function generateWeeklyReview() {
  const now=new Date();
  const end=new Date(now); end.setHours(23,59,59,999);
  const start=new Date(end); start.setDate(start.getDate()-6); start.setHours(0,0,0,0);
  const sh=getSheet_('TASKS');
  const tasks=sh.getLastRow()<2?[]:sh.getRange(2,1,sh.getLastRow()-1,16).getValues().filter(r=>r[5]==='Done'&&r[14]&&new Date(r[14])>=start&&new Date(r[14])<=end);
  const total=sh.getLastRow()<2?0:sh.getRange(2,1,sh.getLastRow()-1,16).getValues().filter(r=>r[6]&&new Date(r[6])>=start&&new Date(r[6])<=end).length;
  const big=tasks.filter(r=>r[3]==='BIG').length;
  const focus=tasks.reduce((s,r)=>s+Number(r[10]||r[9]||0),0);
  const exp=getExpensesBetween_(start,end); const inc=getIncomeBetween_(start,end);
  const productivity=total?tasks.length/total:0;
  const projectCounts={}; tasks.forEach(r=>projectCounts[r[2]]=(projectCounts[r[2]]||0)+1);
  const best=Object.keys(projectCounts).sort((a,b)=>projectCounts[b]-projectCounts[a])[0]||'';
  const ws=getSheet_('WEEKLY_REVIEW');
  ws.insertRowBefore(2);
  ws.getRange(2,1,1,11).setValues([[end,tasks.length,big,total,focus,productivity,inc,exp,inc-exp,best,'']]);
  ws.getRange(2,1).setNumberFormat('dd-mmm-yyyy'); ws.getRange(2,6).setNumberFormat('0%'); ws.getRange(2,7,1,3).setNumberFormat('৳#,##0.00');
  refreshDashboard();
}

function getExpensesBetween_(start,end){
  const sh=getSheet_('EXPENSES'); if(sh.getLastRow()<2)return 0;
  return sh.getRange(2,1,sh.getLastRow()-1,11).getValues().reduce((s,r)=>{const d=new Date(r[1]);return d>=start&&d<=end?s+Number(r[4]||0):s;},0);
}
function getIncomeBetween_(start,end){
  const sh=getSheet_('INCOME'); if(sh.getLastRow()<2)return 0;
  return sh.getRange(2,1,sh.getLastRow()-1,8).getValues().reduce((s,r)=>{const d=new Date(r[1]);return d>=start&&d<=end?s+Number(r[4]||0):s;},0);
}

function addFocusMinutes(taskId, minutes) {
  minutes = Number(minutes || 0);
  if (!(minutes > 0)) throw new Error('Focus minutes must be greater than 0.');
  if (!taskId) return {ok:true,minutes:minutes};
  const sh=getSheet_('TASKS');
  const row=findRowById_(sh,taskId,1);
  if(!row) throw new Error('Task ID not found: '+taskId);
  const cell=sh.getRange(row,11);
  cell.setValue(Number(cell.getValue()||0)+minutes);
  log_('FOCUS','TASK',taskId,minutes+' minutes');
  refreshDashboard();
  return {ok:true,minutes:minutes};
}

function submitDailyReview(data) {
  const sh=getSheet_('DAILY_REVIEW');
  const key=todayKey_();
  const values=sh.getLastRow()<2?[]:sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
  const rowIndex=values.findIndex(r=>r[0]&&Utilities.formatDate(new Date(r[0]),APP.TZ,'yyyy-MM-dd')===key);
  const d=getDashboardData();
  const row=[new Date(),d.big3.filter(x=>x.status==='Done').map(x=>x.task).join(' | '),d.big3.filter(x=>x.status==='Done').length,d.done-d.big3.filter(x=>x.status==='Done').length,Number(data.focusMinutes||0),d.productivity/100,data.whatWentWell||'',data.notDone||'',data.tomorrow||'',data.notes||''];
  if(rowIndex>=0) sh.getRange(rowIndex+2,1,1,10).setValues([row]); else sh.appendRow(row);
  sh.getRange(rowIndex>=0?rowIndex+2:sh.getLastRow(),6).setNumberFormat('0%');
  return true;
}

function findRowById_(sh,id,col){
  if(sh.getLastRow()<2)return 0;
  const vals=sh.getRange(2,col,sh.getLastRow()-1,1).getValues();
  const idx=vals.findIndex(r=>String(r[0])===String(id)); return idx<0?0:idx+2;
}
function log_(action,entity,id,details){getSheet_('LOGS').appendRow([new Date(),action,entity,id,details]);}
function repairHub(){setupProjects_();setupTasks_();setupExpenses_();setupIncome_();setupGoals_();setupMonthlyGoals_();setupDailyReview_();setupWeeklyReview_();setupLogs_();setupAnalytics_();setupDashboard_();applyValidations_();applyFormatting_();createNamedRanges_();refreshDashboard();SpreadsheetApp.getUi().alert('Rizvi Task Hub formatting and validations repaired. Your settings and data were preserved.');}

