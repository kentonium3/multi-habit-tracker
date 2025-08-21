/**
 * @fileoverview Robust HabitID generation system that persists through row deletions
 * Stores the highest ID counter in the Config sheet to avoid permission issues
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
    
  } catch (error) {
    log('ERROR', 'Failed to repair missing HabitIDs:', error.message);
  }
}