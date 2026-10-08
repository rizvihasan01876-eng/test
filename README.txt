RIZVI TASK HUB
==============

Files:
- Code.gs       -> Google Apps Script backend
- Sidebar.html  -> Google Apps Script sidebar UI

INSTALLATION
------------
1. Create a new blank Google Sheet.
2. Extensions > Apps Script.
3. In the Apps Script editor, replace Code.gs with the supplied Code.gs.
4. Add a new HTML file named exactly: Sidebar
5. Paste the supplied Sidebar.html into that file.
6. Save the project.
7. Run setupRizviTaskHub() manually once from the Apps Script editor.
8. Accept the Google authorization prompts.
9. Return to the Sheet and reload it.
10. Open the "🚀 Rizvi Hub" menu.
11. If Calendar/Email automation needs permission, run "Install/Refresh Automations" once.

IMPORTANT
---------
- The first setup creates the full sheet system and keeps your requested defaults:
  Monthly budget: 3000 BDT
  Big tasks per day: 3
  Projects: Personal, DiaFit, Tenova, Divine Cloth, Quicon
  Expense categories: Food, Transport, Bills
  Work hours: 09:00-19:00
  Morning planning: 09:00
  Daily review: 19:05
  Weekly review: Friday
- Change Monthly Budget in Settings!B6.
- Set Reminder Email in Settings!B14 if the detected email is blank or not the desired address.
- Google Calendar tasks sync when a task has Due Date + Start Time + End Time.
- The script creates a separate calendar named "Rizvi Task Hub" if one does not already exist.
- Email reminders are sent only if Reminder Email is populated.
- "Repair Formatting & Validations" preserves existing settings and data.

SHEETS CREATED
--------------
Dashboard, Tasks, Expenses, Income, Goals, Monthly Goals, Daily Review,
Weekly Review, Projects, Settings, Logs, Analytics

MAIN FEATURES
-------------
- Daily Big 3 with a hard maximum of 3 per date
- Unlimited small tasks, with 10 shown as the dashboard quick-task recommendation
- Critical/High/Medium/Low priority
- Project tracking for Personal, DiaFit, Tenova, Divine Cloth, Quicon
- Task scoring
- Google Calendar task sync
- Morning, daily review and weekly email automations
- Monthly expense budget protection with smart override
- Safe daily spending calculation
- Income tracking
- Savings/goal tracking
- Focus timer and focus-minute logging
- Daily and weekly review
- Analytics charts
