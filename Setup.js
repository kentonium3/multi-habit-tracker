/**
 * @fileoverview Updated Setup.js with frequency-based tracking migration
 * This function should be run manually once to set up the sheet.
 */

/**
 * Creates the header row for the Habits_Main sheet with frequency-based fields.
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
      'FrequencyPerPeriod', 'Status', 'CreatedDate', 'Notes'
    ];

    // Set the header row
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Set formatting for the header row
    sheet.setFrozenRows(1); 
    sheet.getRange('A1:I1').setFontWeight('bold');
    
    // Adjust column widths for better readability
    sheet.setColumnWidth(1, 80);   // HabitID
    sheet.setColumnWidth(2, 150);  // HabitName
    sheet.setColumnWidth(3, 100);  // StartDate
    sheet.setColumnWidth(4, 100);  // EndDate
    sheet.setColumnWidth(5, 100);  // Frequency
    sheet.setColumnWidth(6, 120);  // FrequencyPerPeriod
    sheet.setColumnWidth(7, 80);   // Status
    sheet.setColumnWidth(8, 120);  // CreatedDate
    sheet.setColumnWidth(9, 200);  // Notes

    log('INFO', 'Habits_Main header row created successfully with frequency tracking.');
  } catch (error) {
    log('ERROR', 'Failed to create Habits_Main header:', error.message);
  }
}

/**
 * Creates the header row for the Daily_Tracking sheet with frequency-based fields.
 */
function createTrackingSheetHeader() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.TRACKING);
    if (!sheet) {
      log('ERROR', `Sheet not found: ${SHEET_NAMES.TRACKING}`);
      return;
    }

    const headers = [
      'Timestamp', 'HabitID', 'HabitName', 'Frequency', 'TargetFrequencyPerPeriod', 'ActualCompletions',
      'Success', 'Comments', 'CreatedDate'
    ];

    // Set the header row
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Set formatting for the header row
    sheet.setFrozenRows(1); 
    sheet.getRange('A1:H1').setFontWeight('bold');
    
    // Adjust column widths for better readability
    sheet.setColumnWidth(1, 120);  // Timestamp
    sheet.setColumnWidth(2, 80);   // HabitID
    sheet.setColumnWidth(3, 150);  // HabitName
    sheet.setColumnWidth(4, 100);  // Frequency
    sheet.setColumnWidth(5, 150);  // TargetFrequencyPerPeriod
    sheet.setColumnWidth(6, 130);  // ActualCompletions
    sheet.setColumnWidth(7, 80);   // Success
    sheet.setColumnWidth(8, 200);  // Comments
    sheet.setColumnWidth(9, 120);  // CreatedDate

    log('INFO', 'Daily_Tracking header row created successfully with frequency tracking.');
  } catch (error) {
    log('ERROR', 'Failed to create Daily_Tracking header:', error.message);
  }
}

/**
 * ONE-TIME MIGRATION FUNCTION: Migrates existing sheets to frequency-based tracking
 * This function analyzes the current state and makes the necessary changes
 * 
 * **IMPORTANT: This should only be run ONCE to migrate from time-based to frequency-based tracking**
 */
function migrateToFrequencyTracking() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Migrate to Frequency Tracking',
      'This will convert your habit tracker from time-based to frequency-based tracking.\n\n' +
      'Changes:\n' +
      '• Replace "DesiredTimes" with "FrequencyPerDay" (1x, 2x, 3x, 4x)\n' +
      '• Update tracking data structure\n' +
      '• Modify form questions\n' +
      '• Update dashboard analytics\n\n' +
      'This operation cannot be undone. Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      log('INFO', 'Migration cancelled by user');
      return;
    }

    log('INFO', 'Starting migration to frequency-based tracking...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Migrate Habits_Main sheet
    migrateHabitsMainSheet(ss);
    
    // 2. Migrate Daily_Tracking sheet
    migrateTrackingSheet(ss);
    
    // 3. Update Dashboard sheet headers
    migrateDashboardSheet(ss);
    
    // 4. Update form for frequency tracking
    updateFormForFrequencyTracking();
    
    // 5. Update data validation rules
    updateDataValidationForFrequency(ss);
    
    log('INFO', 'Migration to frequency-based tracking completed successfully!');
    
    ui.alert(
      'Migration Complete!',
      'Your habit tracker has been successfully converted to frequency-based tracking.\n\n' +
      'Next steps:\n' +
      '1. Review your habits and set appropriate FrequencyPerDay values (1-4)\n' +
      '2. Test the updated form\n' +
      '3. Check the dashboard for new frequency-based analytics',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    log('ERROR', 'Migration failed:', error.message, error.stack);
    SpreadsheetApp.getUi().alert('Migration Failed', `Error: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Migrates the Habits_Main sheet structure
 * @param {SpreadsheetApp.Spreadsheet} ss The spreadsheet object
 */
function migrateHabitsMainSheet(ss) {
  const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
  if (!habitsSheet) {
    log('ERROR', 'Habits_Main sheet not found during migration');
    return;
  }
  
  log('INFO', 'Migrating Habits_Main sheet structure...');
  
  // Get current data
  const lastRow = habitsSheet.getLastRow();
  const lastCol = habitsSheet.getLastColumn();
  
  if (lastRow < 1) {
    // Empty sheet, just create new headers
    createHabitsMainHeader();
    return;
  }
  
  const allData = habitsSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = allData[0];
  
  // Check current structure and migrate
  const desiredTimesIndex = headers.indexOf('DesiredTimes');
  const frequencyIndex = headers.indexOf('Frequency');
  const frequencyPerDayIndex = headers.indexOf('FrequencyPerDay');
  
  if (desiredTimesIndex !== -1) {
    // Replace DesiredTimes column with FrequencyPerPeriod
    habitsSheet.getRange(1, desiredTimesIndex + 1).setValue('FrequencyPerPeriod');
    
    // Convert existing data to frequency format
    for (let row = 2; row <= lastRow; row++) {
      const cell = habitsSheet.getRange(row, desiredTimesIndex + 1);
      const currentValue = cell.getValue();
      
      // Convert time-based values to frequency (default to 1 if unclear)
      let frequency = 1;
      if (typeof currentValue === 'string') {
        // Try to extract number from time strings
        const matches = currentValue.match(/\d+/);
        if (matches) {
          frequency = Math.min(4, Math.max(1, parseInt(matches[0])));
        }
      } else if (typeof currentValue === 'number') {
        frequency = Math.min(4, Math.max(1, Math.round(currentValue)));
      }
      
      cell.setValue(frequency);
    }
    
    log('INFO', 'Converted DesiredTimes column to FrequencyPerPeriod');
  } else if (frequencyPerDayIndex !== -1) {
    // Update FrequencyPerDay column to FrequencyPerPeriod
    habitsSheet.getRange(1, frequencyPerDayIndex + 1).setValue('FrequencyPerPeriod');
    log('INFO', 'Renamed FrequencyPerDay column to FrequencyPerPeriod');
  } else if (frequencyIndex !== -1) {
    // Update existing frequency column logic...
    habitsSheet.getRange(1, frequencyIndex + 1).setValue('FrequencyPerDay');
    
    for (let row = 2; row <= lastRow; row++) {
      const cell = habitsSheet.getRange(row, frequencyIndex + 1);
      const currentValue = cell.getValue();
      
      // Convert text frequency to numbers
      let frequency = 1;
      if (typeof currentValue === 'string') {
        const lower = currentValue.toLowerCase();
        if (lower.includes('daily') || lower.includes('1')) frequency = 1;
        else if (lower.includes('2') || lower.includes('twice')) frequency = 2;
        else if (lower.includes('3') || lower.includes('thrice')) frequency = 3;
        else if (lower.includes('4') || lower.includes('four')) frequency = 4;
      } else if (typeof currentValue === 'number') {
        frequency = Math.min(4, Math.max(1, Math.round(currentValue)));
      }
      
      cell.setValue(frequency);
    }
    
    log('INFO', 'Updated Frequency column to FrequencyPerDay with numeric values');
  } else {
    // No frequency column exists, need to add one
    const newHeaders = ['HabitID', 'HabitName', 'StartDate', 'EndDate', 'Frequency', 'FrequencyPerPeriod', 'Status', 'CreatedDate', 'Notes'];
    
    // Insert new column if needed
    if (lastCol < 9) {
      habitsSheet.insertColumnsAfter(lastCol, 9 - lastCol);
    }
    
    // Set new headers
    habitsSheet.getRange(1, 1, 1, 9).setValues([newHeaders]);
    
    // Set default frequency for existing habits
    if (lastRow > 1) {
      const frequencyRange = habitsSheet.getRange(2, 5, lastRow - 1, 1);
      const defaultFrequencyValues = Array(lastRow - 1).fill(['Daily']);
      frequencyRange.setValues(defaultFrequencyValues);
      
      const frequencyPerPeriodRange = habitsSheet.getRange(2, 6, lastRow - 1, 1);
      const defaultPerPeriodValues = Array(lastRow - 1).fill([1]);
      frequencyPerPeriodRange.setValues(defaultPerPeriodValues);
    }
    
    log('INFO', 'Added Frequency and FrequencyPerPeriod columns with default values');
  }
}

/**
 * Migrates the Daily_Tracking sheet structure
 * @param {SpreadsheetApp.Spreadsheet} ss The spreadsheet object
 */
function migrateTrackingSheet(ss) {
  const trackingSheet = ss.getSheetByName(SHEET_NAMES.TRACKING);
  if (!trackingSheet) {
    log('INFO', 'Daily_Tracking sheet not found, will be created on first form submission');
    return;
  }
  
  log('INFO', 'Migrating Daily_Tracking sheet structure...');
  
  const lastRow = trackingSheet.getLastRow();
  
  if (lastRow < 1) {
    // Empty sheet, create new headers
    createTrackingSheetHeader();
    return;
  }
  
  // For existing data, we'll preserve what we can and add new columns
  const newHeaders = ['Timestamp', 'HabitID', 'HabitName', 'TargetFrequency', 'ActualCompletions', 'Success', 'Comments', 'CreatedDate'];
  
  // Backup existing data by creating a backup sheet
  const backupSheet = ss.insertSheet('Tracking_Backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd'));
  const existingData = trackingSheet.getDataRange().getValues();
  if (existingData.length > 0) {
    backupSheet.getRange(1, 1, existingData.length, existingData[0].length).setValues(existingData);
    log('INFO', 'Created backup of existing tracking data');
  }
  
  // Clear and recreate tracking sheet with new structure
  trackingSheet.clear();
  trackingSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  
  // Set formatting
  trackingSheet.setFrozenRows(1);
  trackingSheet.getRange('A1:H1').setFontWeight('bold');
  
  log('INFO', 'Updated Daily_Tracking sheet structure (existing data backed up)');
}

/**
 * Migrates the Dashboard sheet
 * @param {SpreadsheetApp.Spreadsheet} ss The spreadsheet object
 */
function migrateDashboardSheet(ss) {
  const dashboardSheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD);
  if (!dashboardSheet) {
    log('INFO', 'Dashboard sheet not found, will be created when dashboard is updated');
    return;
  }
  
  log('INFO', 'Clearing dashboard for frequency-based analytics...');
  
  // Clear dashboard - it will be repopulated with new structure on next update
  dashboardSheet.clear();
  
  // The dashboard will be rebuilt automatically when updateDashboard() is called
}

/**
 * Updates data validation rules for frequency-based tracking
 * @param {SpreadsheetApp.Spreadsheet} ss The spreadsheet object
 */
function updateDataValidationForFrequency(ss) {
  const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
  if (!habitsSheet) return;
  
  log('INFO', 'Updating data validation for frequency tracking...');
  
  // Clear existing data validation
  habitsSheet.getRange('A:Z').clearDataValidations();
  
  // Set correct data validation rules based on actual column structure
  // Actual Headers: HabitID, HabitName, StartDate, EndDate, Frequency, FrequencyPerDay, Status, CreatedDate, Notes
  
  // Column C: StartDate - Date validation
  const startDateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  habitsSheet.getRange('C2:C').setDataValidation(startDateRule);
  
  // Column D: EndDate - Date validation  
  const endDateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  habitsSheet.getRange('D2:D').setDataValidation(endDateRule);

  // Column E: Frequency - Daily/Weekly validation
  const frequencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Daily', 'Weekly'])
    .setAllowInvalid(false)
    .build();
  habitsSheet.getRange('E2:E').setDataValidation(frequencyRule);

  // Column F: FrequencyPerPeriod - Times per period validation (1, 2, 3, 4)
  const timesPerPeriodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4'])
    .setAllowInvalid(false)
    .build();
  habitsSheet.getRange('F2:F').setDataValidation(timesPerPeriodRule);

  // Column G: Status - Status validation (Active, Paused, Completed)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Paused', 'Completed'])
    .setAllowInvalid(false)
    .build();
  habitsSheet.getRange('G2:G').setDataValidation(statusRule);

  // Column H: CreatedDate - No validation needed (auto-generated)
  // Column I: Notes - No validation needed (free text)
  
  log('INFO', 'Updated data validation rules with correct column assignments');
}

/**
 * One-time fix function to correct data validation on existing sheet
 * Run this if data validation was applied to wrong columns
 */
function fixDataValidationColumns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
    
    if (!habitsSheet) {
      log('ERROR', 'Habits_Main sheet not found');
      return;
    }
    
    log('INFO', 'Fixing data validation column assignments...');
    
    // Get current headers to verify structure
    const headers = habitsSheet.getRange(1, 1, 1, 10).getValues()[0];
    log('INFO', 'Current headers:', headers.join(', '));
    
    // Clear all existing data validation
    habitsSheet.getRange('A:Z').clearDataValidations();
    
    // Apply correct data validation
    updateDataValidationForFrequency(ss);
    
    log('INFO', 'Data validation columns fixed successfully');
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Data Validation Fixed', 
      'Column data validation has been corrected:\n\n' +
      '• Column E (Frequency): Daily, Weekly\n' +
      '• Column F (FrequencyPerPeriod): 1, 2, 3, 4\n' +
      '• Column G (Status): Active, Paused, Completed\n' +
      '• Columns C & D: Date validation\n\n' +
      'Your habit data is preserved.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    log('ERROR', 'Failed to fix data validation:', error.message);
    SpreadsheetApp.getUi().alert('Fix Failed', `Error: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}