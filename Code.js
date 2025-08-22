/**
 * @fileoverview Updated Code.js for frequency-based habit tracking
 * Updated data validation and menu items
 */

/**
 * --- GLOBAL CONFIGURATION ---
 * Modify these constants to match your sheet names.
 * Do not change them once the system is live.
 */
const SHEET_NAMES = {
  HABITS: 'Habits_Main',
  TRACKING: 'Daily_Tracking',
  DASHBOARD: 'Dashboard',
  CONFIG: 'Config'
};

/**
 * Trigger function that fires on a new Google Form submission.
 * It's the primary entry point for processing new habit entries.
 * @param {Object} e The event object containing form submission data.
 */
function onFormSubmit(e) {
  // Use a try-catch block for robust error handling
  try {
    const config = getConfig();
    if (config.debugMode) {
      log('INFO', 'onFormSubmit triggered.');
      log('DEBUG', 'Event object received:', JSON.stringify(e));
    }

    // Process the new form entry
    const entryData = e.namedValues;
    processNewEntry(entryData);

    // Update the dashboard with new data
    updateDashboard();

    // Check if it's time to send an email
    const lastSentTimestamp = new Date(config.lastEmailSent);
    if (shouldSendEmail(lastSentTimestamp, config.emailFrequency)) {
      sendAccountabilityEmail();
    }

    log('INFO', 'onFormSubmit completed successfully.');

  } catch (error) {
    log('ERROR', 'An error occurred during form submission processing:', error.message, error.stack);
    // You could also send an alert email to the admin here.
  }
}

/**
 * Trigger function that fires on a spreadsheet edit.
 * This is a simple trigger and does not require manual setup in the Apps Script Triggers menu.
 * Uses the robust HabitID system that won't break when rows are deleted.
 * @param {Object} e The event object containing information about the edit.
 */
function onEdit(e) {
  try {
    const config = getConfig();
    if (config.debugMode) {
      log('DEBUG', 'onEdit triggered.');
    }
    
    const range = e.range;
    const sheet = range.getSheet();
    
    // Check if the edit is in the Habits_Main sheet, in the second column (HabitName),
    // and if a new habit is being entered.
    if (sheet.getName() === SHEET_NAMES.HABITS && range.getColumn() === 2) {
      const habitIdCell = sheet.getRange(range.getRow(), 1);
      const habitName = range.getValue();
      
      // Check if the HabitID cell is empty and a HabitName has been entered
      if (habitIdCell.getValue() === '' && habitName !== '') {
        // Use the robust ID generation system from IDGenerator.js
        const newId = getNextHabitID();
        
        if (newId !== 'ERROR') {
          habitIdCell.setValue(newId);
          log('INFO', `Auto-populated HabitID ${newId} for new habit: ${habitName}`);
          
          // Update the form update button to signal a required update
          setFormUpdateButtonStatus('pending');
        } else {
          log('ERROR', 'Failed to generate HabitID for new habit');
        }
      }
    }
  } catch (error) {
    log('ERROR', 'An error occurred during onEdit processing:', error.message, error.stack);
  }
}

/**
 * Trigger function that runs when the spreadsheet is opened.
 * It creates a custom menu to make core functions easily accessible.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⚙️ My Habit Tracker')
      .addItem('Setup Initial Configuration', 'setupInitialConfig')
      .addSeparator()
      .addItem('Update Form Dropdown', 'updateFormDropdownAndStatus')
      .addItem('Repair Missing HabitIDs', 'repairMissingHabitIDs')
      .addItem('Test Email Functionality', 'testEmailFunctionality')
      .addSeparator()
      .addItem('🔄 Migrate to Frequency Tracking', 'migrateToFrequencyTracking')
      .addItem('🔧 Fix Data Validation Columns', 'fixDataValidationColumns')
      .addToUi();
}

/**
 * A wrapper function to update the form dropdown and reset the button status.
 * This function is attached to the custom menu item.
 */
function updateFormDropdownAndStatus() {
  try {
    updateHabitFormDropdown();
    setFormUpdateButtonStatus('success');
  } catch (error) {
    log('ERROR', 'Failed to update form dropdown and status:', error.message);
  }
}

/**
 * Sets the status of the form update button.
 * @param {string} status The status to set ('pending' or 'success').
 */
function setFormUpdateButtonStatus(status) {
  const habitsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.HABITS);
  if (!habitsSheet) {
    log('ERROR', 'Habits_Main sheet not found. Cannot set button status.');
    return;
  }
  
  const buttonCell = habitsSheet.getRange('J1');
  
  if (status === 'pending') {
    buttonCell.setValue('⚠️ Form Update Needed');
    buttonCell.setBackground('#f4e868'); // Yellow color
  } else if (status === 'success') {
    buttonCell.setValue('✅ Form Updated');
    buttonCell.setBackground('#a8e4a0'); // Green color
  }
}

/**
 * Initializes the entire system with frequency-based tracking.
 * This function should be run once manually after setup.
 * It creates the required sheets, sets initial config, and sets up triggers.
 */
function setupInitialConfig() {
  try {
    log('INFO', 'Starting initial configuration setup...');
    
    // Create and initialize the Config sheet
    createConfigSheet();

    // Initialize the HabitID counter system (uses functions from IDGenerator.js)
    initializeHabitIDCounter();

    // Set up the form and its triggers
    const formUrl = setupFormTrigger();
    log('INFO', 'Form is configured and trigger is set. Form URL:', formUrl);
    
    // Update the form dropdown for the first time
    updateHabitFormDropdown();

    // Set up data validation for the Habits_Main sheet with frequency-based fields
    const habitsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.HABITS);
    if (habitsSheet) {
      // Data validation for Frequency column (Daily, Weekly)
      const frequencyRule = SpreadsheetApp.newDataValidation().requireValueInList(['Daily', 'Weekly']).setAllowInvalid(false).build();
      habitsSheet.getRange('E2:E').setDataValidation(frequencyRule);

      // Data validation for FrequencyPerPeriod column (1, 2, 3, 4)
      const timesPerPeriodRule = SpreadsheetApp.newDataValidation().requireValueInList(['1', '2', '3', '4']).setAllowInvalid(false).build();
      habitsSheet.getRange('F2:F').setDataValidation(timesPerPeriodRule);

      // Data validation for Status column
      const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Paused', 'Completed']).setAllowInvalid(false).build();
      habitsSheet.getRange('G2:G').setDataValidation(statusRule);

      // Data validation for StartDate and EndDate columns (must be a valid date)
      const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
      habitsSheet.getRange('C2:C').setDataValidation(dateRule);
      habitsSheet.getRange('D2:D').setDataValidation(dateRule);

      log('INFO', 'Data validation rules added to Habits_Main sheet for frequency tracking.');
    }

    // All done, set the last email sent timestamp to prevent immediate sending
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
    const lastSentCell = configSheet.getRange('B3');
    lastSentCell.setValue(new Date());

    log('INFO', 'Initial configuration completed successfully. You can now use the form!');
    
  } catch (error) {
    log('ERROR', 'Initial configuration failed:', error.message, error.stack);
  }
}