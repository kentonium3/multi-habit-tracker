/**
 * @fileoverview Updated utility functions for frequency-based habit tracking
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
 * Processes a new form entry for frequency-based habit tracking
 * @param {Object} entryData The named values from the form submission.
 */
function processNewEntry(entryData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackingSheet = ss.getSheetByName(SHEET_NAMES.TRACKING);
  const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
  
  const habitName = entryData['Habit Selection'][0];
  const completionStatus = entryData['Completion Status'][0] || entryData['Success/Miss'][0]; // Handle both old and new forms
  const completionCount = entryData['How many times completed today?'][0] || '1'; // Default to 1 for backward compatibility
  const comments = entryData['Comments'][0] || '';
  const entryTimestamp = new Date();
  
  // Find the HabitID and target frequency for the submitted habit
  const habitData = habitsSheet.getRange(2, 1, habitsSheet.getLastRow() - 1, 7).getValues(); // Get through column G (Status)
  const habitRow = habitData.find(row => row[1] === habitName);
  
  if (!habitRow) {
    log('ERROR', `Habit not found for habit name: ${habitName}`);
    return;
  }
  
  const habitId = habitRow[0];
  const frequency = habitRow[4] || 'Daily'; // Column E is Frequency
  const targetFrequencyPerPeriod = habitRow[5] || 1; // Column F is FrequencyPerPeriod
  
  // Parse completion count
  let actualCount = 1;
  if (completionCount === '5+') {
    actualCount = 5;
  } else {
    actualCount = parseInt(completionCount) || 1;
  }
  
  // Determine success based on completion status and count vs target
  let success = false;
  if (completionStatus === 'Completed' || completionStatus === 'Success') {
    success = actualCount >= targetFrequencyPerPeriod;
  }
  
  // Append new row to the tracking sheet with frequency-based data
  const newRow = [
    entryTimestamp,              // A: Timestamp
    habitId,                     // B: HabitID  
    habitName,                   // C: HabitName
    frequency,                   // D: Frequency (Daily/Weekly)
    targetFrequencyPerPeriod,    // E: TargetFrequencyPerPeriod
    actualCount,                 // F: ActualCompletions
    success,                     // G: Success (true if actualCount >= targetFrequencyPerPeriod)
    comments,                    // H: Comments
    entryTimestamp               // I: CreatedDate
  ];
  
  trackingSheet.appendRow(newRow);
  log('INFO', `New frequency-based entry appended to ${SHEET_NAMES.TRACKING}:`, JSON.stringify(newRow));
  log('INFO', `Habit: ${habitName}, Frequency: ${frequency}, Target: ${targetFrequencyPerPeriod}x, Actual: ${actualCount}x, Success: ${success}`);
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

/**
 * Gets the total completions for a habit on a specific date
 * @param {Array<Array>} trackingData All tracking data
 * @param {string} habitId The habit ID to check
 * @param {Date} date The date to check
 * @return {number} Total completions for that habit on that date
 */
function getHabitCompletionsForDate(trackingData, habitId, date) {
  const dateKey = date.toISOString().split('T')[0];
  
  return trackingData
    .filter(entry => {
      const entryDate = new Date(entry[0]);
      const entryDateKey = entryDate.toISOString().split('T')[0];
      return entry[1] === habitId && entryDateKey === dateKey;
    })
    .reduce((total, entry) => total + (entry[5] || 1), 0); // Sum ActualCompletions (column F)
}

/**
 * Checks if a habit was successful on a specific date based on frequency
 * @param {Array<Array>} trackingData All tracking data
 * @param {string} habitId The habit ID to check
 * @param {Date} date The date to check
 * @param {number} targetFrequency Required completions per day
 * @return {boolean} True if habit met its frequency target
 */
function wasHabitSuccessfulOnDate(trackingData, habitId, date, targetFrequency) {
  const completions = getHabitCompletionsForDate(trackingData, habitId, date);
  return completions >= targetFrequency;
}