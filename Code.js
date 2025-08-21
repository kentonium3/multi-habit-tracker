/**
 * @fileoverview Main script file containing core functions and triggers.
 * Handles form submissions, initializes the system, and orchestrates other modules.
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
        // Use the robust ID generation system
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
 * Initializes the entire system.
 * This function should be run once manually after setup.
 * It creates the required sheets, sets initial config, and sets up triggers.
 */
function setupInitialConfig() {
  try {
    log('INFO', 'Starting initial configuration setup...');
    
    // Create and initialize the Config sheet
    createConfigSheet();

    // Initialize the HabitID counter system (NEW - ROBUST ID SYSTEM)
    initializeHabitIDCounter();

    // Set up the form and its triggers
    const formUrl = setupFormTrigger();
    log('INFO', 'Form is configured and trigger is set. Form URL:', formUrl);
    
    // Update the form dropdown for the first time
    updateHabitFormDropdown();

    // Set up data validation for the Habits_Main sheet
    const habitsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.HABITS);
    if (habitsSheet) {
      // Data validation for Status column
      const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Paused', 'Completed']).setAllowInvalid(false).build();
      habitsSheet.getRange('G2:G').setDataValidation(statusRule);

      // Data validation for Frequency column
      const frequencyRule = SpreadsheetApp.newDataValidation().requireValueInList(['Daily', 'Weekly', 'Monthly']).setAllowInvalid(false).build();
      habitsSheet.getRange('E2:E').setDataValidation(frequencyRule);

      // Data validation for StartDate and EndDate columns (must be a valid date)
      const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
      habitsSheet.getRange('C2:C').setDataValidation(dateRule);
      habitsSheet.getRange('D2:D').setDataValidation(dateRule);

      // Data validation for DesiredTimes (must be a time value or 'Any')
      const timeOrAnyRule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied('=OR(F2="Any", ISNUMBER(F2))').setAllowInvalid(false).build();
      habitsSheet.getRange('F2:F').setDataValidation(timeOrAnyRule);

      log('INFO', 'Data validation rules added to Habits_Main sheet.');
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

/**
 * ============================================================================
 * ROBUST HABITID GENERATION SYSTEM
 * ============================================================================
 */

/**
 * Gets the next available HabitID by checking the Config sheet counter
 * This approach is deletion-proof and doesn't rely on scanning existing rows
 * @return {string} A unique habit ID (e.g., "H001", "H002")
 */
function getNextHabitID() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    
    if (!configSheet) {
      log('ERROR', 'Config sheet not found. Creating it now...');
      createConfigSheet();
      return getNextHabitID(); // Retry after creating the sheet
    }
    
    // Get current highest ID from Config sheet
    let highestID = getStoredHighestHabitID(configSheet);
    
    // For extra safety, verify against actual data (in case of corruption)
    const actualHighestID = findHighestIDInHabitsSheet();
    
    // Use whichever is higher (handles edge cases)
    highestID = Math.max(highestID, actualHighestID);
    
    // Increment for next ID
    const nextID = highestID + 1;
    
    // Store the new highest ID in Config sheet
    storeHighestHabitID(configSheet, nextID);
    
    // Format as H001, H002, etc.
    const formattedID = 'H' + nextID.toString().padStart(3, '0');
    
    log('INFO', `Generated new HabitID: ${formattedID}`);
    return formattedID;
    
  } catch (error) {
    log('ERROR', 'Failed to generate HabitID:', error.message, error.stack);
    return 'ERROR';
  }
}

/**
 * Gets the stored highest HabitID number from the Config sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet The Config sheet
 * @return {number} The highest HabitID number used so far
 */
function getStoredHighestHabitID(configSheet) {
  try {
    const data = configSheet.getDataRange().getValues();
    
    // Look for the 'Highest_Habit_ID' setting
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'Highest_Habit_ID') {
        const value = data[i][1];
        return (typeof value === 'number') ? value : 0;
      }
    }
    
    // If not found, initialize it
    log('INFO', 'Highest_Habit_ID not found in Config. Initializing...');
    configSheet.appendRow(['Highest_Habit_ID', 0]);
    return 0;
    
  } catch (error) {
    log('ERROR', 'Error reading highest HabitID from Config:', error.message);
    return 0;
  }
}

/**
 * Stores the new highest HabitID number in the Config sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet The Config sheet
 * @param {number} newHighestID The new highest ID number to store
 */
function storeHighestHabitID(configSheet, newHighestID) {
  try {
    const data = configSheet.getDataRange().getValues();
    
    // Find and update the existing row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'Highest_Habit_ID') {
        configSheet.getRange(i + 1, 2).setValue(newHighestID);
        log('DEBUG', `Updated Highest_Habit_ID to: ${newHighestID}`);
        return;
      }
    }
    
    // If not found, add it
    configSheet.appendRow(['Highest_Habit_ID', newHighestID]);
    log('DEBUG', `Added Highest_Habit_ID: ${newHighestID}`);
    
  } catch (error) {
    log('ERROR', 'Error storing highest HabitID in Config:', error.message);
  }
}

/**
 * Scans the actual Habits_Main sheet to find the highest ID number
 * This is used as a safety check to prevent ID conflicts
 * @return {number} The highest ID number found in the sheet
 */
function findHighestIDInHabitsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
    
    if (!habitsSheet) {
      log('WARNING', 'Habits_Main sheet not found');
      return 0;
    }
    
    const lastRow = habitsSheet.getLastRow();
    if (lastRow <= 1) {
      return 0; // No data rows
    }
    
    // Get all HabitID values (column A, starting from row 2)
    const idRange = habitsSheet.getRange(2, 1, lastRow - 1, 1);
    const ids = idRange.getValues().flat();
    
    let maxID = 0;
    
    for (const id of ids) {
      if (typeof id === 'string' && id.startsWith('H')) {
        const numericPart = parseInt(id.substring(1));
        if (!isNaN(numericPart) && numericPart > maxID) {
          maxID = numericPart;
        }
      }
    }
    
    log('DEBUG', `Highest ID found in sheet: ${maxID}`);
    return maxID;
    
  } catch (error) {
    log('ERROR', 'Error scanning Habits sheet for highest ID:', error.message);
    return 0;
  }
}

/**
 * Ensures the Config sheet has the Highest_Habit_ID entry
 * Call this during setup to initialize the system
 */
function initializeHabitIDCounter() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    
    if (!configSheet) {
      log('ERROR', 'Config sheet not found. Run setupInitialConfig first.');
      return;
    }
    
    // Get the actual highest ID from the sheet
    const actualHighest = findHighestIDInHabitsSheet();
    
    // Store it in the Config sheet
    storeHighestHabitID(configSheet, actualHighest);
    
    log('INFO', `Initialized HabitID counter to: ${actualHighest}`);
    
  } catch (error) {
    log('ERROR', 'Failed to initialize HabitID counter:', error.message);
  }
}

/**
 * Manual function to repair any missing HabitIDs in existing rows
 * Run this if you have rows without IDs that need to be assigned
 * Available through the custom menu: ⚙️ My Habit Tracker > Repair Missing HabitIDs
 */
function repairMissingHabitIDs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
    
    if (!habitsSheet) {
      log('ERROR', 'Habits_Main sheet not found');
      return;
    }
    
    const lastRow = habitsSheet.getLastRow();
    if (lastRow <= 1) {
      log('INFO', 'No data rows found');
      return;
    }
    
    let repairedCount = 0;
    
    // Check each row for missing HabitID
    for (let row = 2; row <= lastRow; row++) {
      const habitIDCell = habitsSheet.getRange(row, 1);
      const habitNameCell = habitsSheet.getRange(row, 2);
      
      // If there's a habit name but no ID, assign one
      if (habitIDCell.getValue() === '' && habitNameCell.getValue() !== '') {
        const newID = getNextHabitID();
        habitIDCell.setValue(newID);
        repairedCount++;
        log('INFO', `Assigned ${newID} to row ${row} (${habitNameCell.getValue()})`);
      }
    }
    
    log('INFO', `Repair complete. Assigned ${repairedCount} missing HabitIDs.`);
    
    // Show result to user
    const ui = SpreadsheetApp.getUi();
    ui.alert('Repair Complete', `Assigned ${repairedCount} missing HabitIDs.`, ui.ButtonSet.OK);
    
  } catch (error) {
    log('ERROR', 'Failed to repair missing HabitIDs:', error.message);
  }
}