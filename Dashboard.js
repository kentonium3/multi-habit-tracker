/**
 * @fileoverview Updated Dashboard.js for frequency-based habit tracking
 * Handles all complex calculations like streaks and success rates based on daily frequency targets
 */

/**
 * Updates the Dashboard sheet with frequency-based analytics
 * Clears old data and populates a new rolling 30-day view
 */
function updateDashboard() {
  log('INFO', 'Updating dashboard with frequency-based analytics...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
  const trackingSheet = ss.getSheetByName(SHEET_NAMES.TRACKING);
  const dashboardSheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD);

  // Clear existing dashboard data (except headers)
  if (dashboardSheet.getLastRow() > 1) {
    dashboardSheet.getRange(2, 1, dashboardSheet.getLastRow() - 1, dashboardSheet.getLastColumn()).clearContent();
  }

  const habitsData = habitsSheet.getRange(2, 1, habitsSheet.getLastRow() - 1, habitsSheet.getLastColumn()).getValues();
  const trackingData = trackingSheet.getRange(2, 1, trackingSheet.getLastRow() - 1, trackingSheet.getLastColumn()).getValues();
  
  if (!habitsData || habitsData.length === 0 || habitsData[0].length === 0) {
    log('WARN', 'No habits found in Habits_Main sheet.');
    return;
  }
  
  const activeHabits = habitsData.filter(row => row[6] === 'Active'); // Status is now column G (index 6)
  if (activeHabits.length === 0) {
    log('WARN', 'No active habits to display on the dashboard.');
    return;
  }
  
  const dashboardData = [];
  const today = new Date();
  
  // Updated headers for frequency-based dashboard
  const headers = [
    'Habit Name', 
    'Frequency',
    'Target/Period', 
    'Days Since Start', 
    'Days Remaining', 
    'Current Streak', 
    'Success Rate', 
    'Avg Completions/Period',
    ...getRollingDates()
  ];
  dashboardData.push(headers);

  // Process each active habit with frequency-based metrics
  activeHabits.forEach(habit => {
    const habitId = habit[0];
    const habitName = habit[1];
    const startDate = new Date(habit[2]);
    const endDate = new Date(habit[3]);
    const frequency = habit[4] || 'Daily'; // Frequency column (Daily/Weekly)
    const targetFrequencyPerPeriod = habit[5] || 1; // FrequencyPerPeriod column
    
    const relevantTrackingData = trackingData.filter(entry => entry[1] === habitId);
    
    // Calculate frequency-based dashboard metrics
    const daysSinceStart = getDaysBetweenDates(startDate, today);
    const daysRemaining = getDaysBetweenDates(today, endDate);
    const streak = calculateFrequencyBasedStreak(relevantTrackingData, habitId, targetFrequencyPerPeriod);
    const successRate = calculateFrequencyBasedSuccessRate(relevantTrackingData, habitId, targetFrequencyPerPeriod, startDate, today);
    const avgCompletions = calculateAverageCompletionsPerPeriod(relevantTrackingData, frequency, startDate, today);
    
    // Prepare the rolling 30-day view based on frequency targets
    const thirtyDayView = getFrequencyBasedThirtyDayView(relevantTrackingData, habitId, targetFrequencyPerPeriod);
    
    const row = [
      habitName, 
      frequency,
      `${targetFrequencyPerPeriod}x`,
      daysSinceStart, 
      daysRemaining, 
      streak, 
      successRate.toFixed(1) + '%',
      avgCompletions.toFixed(1),
      ...thirtyDayView
    ];
    dashboardData.push(row);
  });

  // Write the data to the dashboard sheet
  dashboardSheet.getRange(1, 1, dashboardData.length, dashboardData[0].length).setValues(dashboardData);
  dashboardSheet.setFrozenRows(1);
  dashboardSheet.setColumnWidth(1, 150);
  dashboardSheet.setColumnWidth(2, 80);
  
  // Apply conditional formatting for frequency-based success/failure
  applyFrequencyBasedConditionalFormatting(dashboardSheet, dashboardData.length, dashboardData[0].length);
  
  log('INFO', 'Frequency-based dashboard updated successfully.');
}

/**
 * Calculates current streak based on meeting daily frequency targets
 * @param {Array<Array>} trackingData The tracking data for a single habit
 * @param {string} habitId The habit ID
 * @param {number} targetFrequency Required completions per day
 * @return {number} The current streak length in days
 */
function calculateFrequencyBasedStreak(trackingData, habitId, targetFrequency) {
  if (!trackingData || trackingData.length === 0) return 0;
  
  const today = new Date();
  let streak = 0;
  
  // Go backwards from today, checking each day
  for (let i = 0; i < 365; i++) { // Check up to a year back
    const checkDate = new Date();
    checkDate.setDate(today.getDate() - i);
    
    const wasSuccessful = wasHabitSuccessfulOnDate(trackingData, habitId, checkDate, targetFrequency);
    
    if (wasSuccessful) {
      streak++;
    } else {
      // Check if there's any data for this date - if no data, don't break streak for recent dates
      const hasDataForDate = trackingData.some(entry => {
        const entryDate = new Date(entry[0]);
        return isSameDay(entryDate, checkDate);
      });
      
      // If it's within the last 2 days and no data, don't break streak yet
      if (i <= 1 && !hasDataForDate) {
        continue;
      } else {
        break; // Streak is broken
      }
    }
  }
  
  return streak;
}

/**
 * Calculates success rate based on meeting daily frequency targets
 * @param {Array<Array>} trackingData The tracking data for a single habit
 * @param {string} habitId The habit ID
 * @param {number} targetFrequency Required completions per day
 * @param {Date} startDate The start of the tracking window
 * @param {Date} endDate The end of the tracking window
 * @return {number} The success rate percentage
 */
function calculateFrequencyBasedSuccessRate(trackingData, habitId, targetFrequency, startDate, endDate) {
  const daysBetween = getDaysBetweenDates(startDate, endDate);
  let successfulDays = 0;
  let totalDaysWithData = 0;
  
  for (let i = 0; i < daysBetween; i++) {
    const checkDate = new Date();
    checkDate.setTime(startDate.getTime() + (i * 24 * 60 * 60 * 1000));
    
    const hasDataForDate = trackingData.some(entry => {
      const entryDate = new Date(entry[0]);
      return entry[1] === habitId && isSameDay(entryDate, checkDate);
    });
    
    if (hasDataForDate) {
      totalDaysWithData++;
      if (wasHabitSuccessfulOnDate(trackingData, habitId, checkDate, targetFrequency)) {
        successfulDays++;
      }
    }
  }
  
  return totalDaysWithData > 0 ? (successfulDays / totalDaysWithData) * 100 : 0;
}

/**
 * Calculates average completions per period (day or week)
 * @param {Array<Array>} trackingData The tracking data for a single habit
 * @param {string} frequency The frequency type (Daily or Weekly)
 * @param {Date} startDate The start of the tracking window
 * @param {Date} endDate The end of the tracking window
 * @return {number} Average completions per period
 */
function calculateAverageCompletionsPerPeriod(trackingData, frequency, startDate, endDate) {
  if (!trackingData || trackingData.length === 0) return 0;
  
  const totalCompletions = trackingData
    .filter(entry => {
      const entryDate = new Date(entry[0]);
      return entryDate >= startDate && entryDate <= endDate;
    })
    .reduce((sum, entry) => sum + (entry[5] || 1), 0); // Sum ActualCompletions (column F)
  
  const daysBetween = Math.max(1, getDaysBetweenDates(startDate, endDate));
  
  if (frequency === 'Weekly') {
    const weeksBetween = Math.max(1, daysBetween / 7);
    return totalCompletions / weeksBetween;
  } else {
    // Daily
    return totalCompletions / daysBetween;
  }
}

/**
 * Gets a frequency-based rolling 30-day view
 * @param {Array<Array>} trackingData The tracking data for a single habit
 * @param {string} habitId The habit ID
 * @param {number} targetFrequency Required completions per day
 * @return {Array<string>} Array of indicators for the last 30 days
 */
function getFrequencyBasedThirtyDayView(trackingData, habitId, targetFrequency) {
  const view = [];
  const today = new Date();
  
  for (let i = 29; i >= 0; i--) {
    const currentDate = new Date();
    currentDate.setDate(today.getDate() - i);
    
    const completions = getHabitCompletionsForDate(trackingData, habitId, currentDate);
    const wasSuccessful = completions >= targetFrequency;
    
    if (completions === 0) {
      view.push('-');
    } else if (wasSuccessful) {
      view.push(completions > targetFrequency ? `${completions}✓` : '✓');
    } else {
      view.push(`${completions}/${targetFrequency}`);
    }
  }
  
  return view;
}

/**
 * Gets the dates for the last 30 days for the dashboard headers.
 * @return {Array<string>} An array of date strings in 'MM/DD' format.
 */
function getRollingDates() {
  const dates = [];
  const today = new Date();
  for (let i = 29; i >= 0; i--) {
    const d = new Date();
    d.setDate(today.getDate() - i);
    dates.push(Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd'));
  }
  return dates;
}

/**
 * Applies conditional formatting for frequency-based indicators
 * @param {Sheet} sheet The dashboard sheet
 * @param {number} numRows Number of rows to format
 * @param {number} numCols Number of columns to format
 */
function applyFrequencyBasedConditionalFormatting(sheet, numRows, numCols) {
  const range = sheet.getRange(2, 8, numRows - 1, numCols - 7); // Start from column H (after Avg Completions)
  
  const ruleSuccess = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('✓')
    .setBackground('#b6d7a8') // Light green
    .setBold(true)
    .setFontColor('#073763')
    .setRanges([range])
    .build();
    
  const rulePartial = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('/')
    .setBackground('#ffd966') // Light yellow for partial completion
    .setBold(true)
    .setFontColor('#cc7a00')
    .setRanges([range])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(ruleSuccess, rulePartial);
  sheet.setConditionalFormatRules(rules);
}