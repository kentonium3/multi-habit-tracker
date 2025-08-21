/**
 * @fileoverview A collection of helper and utility functions.
 */

/**
 * A custom logging function that can be disabled with a config toggle.
 * @param {string} level The log level (e.g., 'INFO', 'WARN', 'ERROR').
 * @param {...*} message The messages to log.
 */
function log(level, ...message) {
  try {
    const config = getConfig();
    if (level === 'DEBUG' && (!config || !config.debugMode)) {
      return;
    }
  } catch (e) {
    // If getConfig() fails (e.g., during initial setup), log regardless of level
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
    console.log(`[${timestamp}] [${level}]`, ...message);
    return;
  }
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  console.log(`[${timestamp}] [${level}]`, ...message);
}

/**
 * Processes a new form entry and appends it to the tracking sheet.
 * @param {Object} entryData The named values from the form submission.
 */
function processNewEntry(entryData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackingSheet = ss.getSheetByName(SHEET_NAMES.TRACKING);
  const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
  
  const habitName = entryData['Habit Selection'][0];
  const statusText = entryData['Success/Miss'][0]; // Changed from 'Success/Failure'
  const comments = entryData['Comments'][0];
  const actualTime = entryData['Time of completion'][0];
  const entryTimestamp = new Date();
  
  // Find the HabitID for the submitted habit
  const habitData = habitsSheet.getRange(2, 1, habitsSheet.getLastRow() - 1, 2).getValues();
  const habitRow = habitData.find(row => row[1] === habitName);
  const habitId = habitRow ? habitRow[0] : null;

  if (!habitId) {
    log('ERROR', `Habit ID not found for habit name: ${habitName}`);
    return;
  }

  // Determine success based on the new status text
  const success = statusText === 'Success';
  
  // Append new row to the tracking sheet
  const newRow = [
    entryTimestamp,
    habitId,
    habitName,
    'Not Applicable', // This would require more complex logic for target times
    actualTime,
    success,
    comments,
    entryTimestamp
  ];
  
  trackingSheet.appendRow(newRow);
  log('INFO', `New entry appended to ${SHEET_NAMES.TRACKING}:`, JSON.stringify(newRow));
}

/**
 * Calculates the number of days between two dates.
 * @param {Date} date1 The start date.
 * @param {Date} date2 The end date.
 * @return {number} The difference in days.
 */
function getDaysBetweenDates(date1, date2) {
  const oneDay = 1000 * 60 * 60 * 24;
  const diffTime = Math.abs(date2.getTime() - date1.getTime());
  return Math.ceil(diffTime / oneDay);
}

/**
 * Checks if two Date objects represent the same calendar day.
 * @param {Date} d1 The first date.
 * @param {Date} d2 The second date.
 * @return {boolean} True if they are the same day.
 */
function isSameDay(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}