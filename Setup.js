/**
 * @fileoverview A helper script for one-time setup of the Habits_Main sheet header.
 * This function should be run manually once to set up the sheet.
 */

/**
 * Creates the header row for the Habits_Main sheet.
 * This should be run once after the sheet is created.
 */
function createHabitsMainHeader() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.HABITS);
    if (!sheet) {
      log('ERROR', `Sheet not found: ${SHEET_NAMES.HABITS}`);
      return;
    }

    const headers = [
      'HabitID', 'HabitName', 'StartDate', 'EndDate', 'Frequency',
      'DesiredTimes', 'Status', 'CreatedDate', 'Notes'
    ];

    // Set the header row
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Set formatting for the header row
    sheet.setFrozenRows(1); 
    sheet.getRange('A1:I1').setFontWeight('bold');
    
    // Adjust column widths for better readability
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 80);
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 200);

    log('INFO', 'Habits_Main header row created successfully.');
  } catch (error) {
    log('ERROR', 'Failed to create Habits_Main header:', error.message);
  }
}
