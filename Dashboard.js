/**
 * @fileoverview Functions for populating and updating the dashboard sheet.
 * Handles all complex calculations like streaks and percentages.
 */

/**
 * Updates the Dashboard sheet with the latest data from the tracking sheet.
 * Clears old data and populates a new rolling 30-day view.
 */
function updateDashboard() {
  log('INFO', 'Updating dashboard...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
  const trackingSheet = ss.getSheetByName(SHEET_NAMES.TRACKING);
  const dashboardSheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD);

  // Clear existing dashboard data (except headers)
  dashboardSheet.getRange(2, 1, dashboardSheet.getLastRow(), dashboardSheet.getLastColumn()).clearContent();

  const habitsData = habitsSheet.getRange(2, 1, habitsSheet.getLastRow() - 1, habitsSheet.getLastColumn()).getValues();
  const trackingData = trackingSheet.getRange(2, 1, trackingSheet.getLastRow() - 1, trackingSheet.getLastColumn()).getValues();
  
  if (!habitsData || habitsData.length === 0 || habitsData[0].length === 0) {
    log('WARN', 'No habits found in Habits_Main sheet.');
    return;
  }
  
  const activeHabits = habitsData.filter(row => row[5] === 'Active');
  if (activeHabits.length === 0) {
    log('WARN', 'No active habits to display on the dashboard.');
    return;
  }
  
  const dashboardData = [];
  const today = new Date();
  
  // Headers for the dashboard
  const headers = ['Habit Name', 'Days Since Start', 'Days Remaining', 'Current Streak', 'Success Percentage', ...getRollingDates()];
  dashboardData.push(headers);

  // Process each active habit
  activeHabits.forEach(habit => {
    const habitId = habit[0];
    const habitName = habit[1];
    const startDate = new Date(habit[2]);
    const endDate = new Date(habit[3]);
    
    const relevantTrackingData = trackingData.filter(entry => entry[1] === habitId);
    
    // Calculate dashboard metrics
    const daysSinceStart = getDaysBetweenDates(startDate, today);
    const daysRemaining = getDaysBetweenDates(today, endDate);
    const streak = calculateStreak(relevantTrackingData);
    const successRate = calculateSuccessRate(relevantTrackingData, startDate, today);
    
    // Prepare the rolling 30-day view
    const thirtyDayView = getThirtyDayView(relevantTrackingData);
    
    const row = [
      habitName, 
      daysSinceStart, 
      daysRemaining, 
      streak, 
      successRate.toFixed(2) + '%', 
      ...thirtyDayView
    ];
    dashboardData.push(row);
  });

  // Write the data to the dashboard sheet
  dashboardSheet.getRange(1, 1, dashboardData.length, dashboardData[0].length).setValues(dashboardData);
  dashboardSheet.setFrozenRows(1);
  dashboardSheet.setColumnWidth(1, 150);
  
  // Optional: Apply conditional formatting for success/failure
  applyConditionalFormatting(dashboardSheet, dashboardData.length, dashboardData[0].length);
  
  log('INFO', 'Dashboard updated successfully.');
}

/**
 * Calculates the current streak for a given habit.
 * @param {Array<Array>} data The tracking data for a single habit.
 * @return {number} The current streak length.
 */
function calculateStreak(data) {
  if (!data || data.length === 0) return 0;
  
  const sortedData = data.sort((a, b) => new Date(b[0]) - new Date(a[0]));
  let streak = 0;
  let lastDate = new Date();
  
  for (let i = 0; i < sortedData.length; i++) {
    const entryDate = new Date(sortedData[i][0]);
    const success = sortedData[i][5];
    
    if (isSameDay(entryDate, lastDate)) {
      if (success) {
        // If there are multiple entries on the same day, we only count one success for the streak
        if (streak > 0) continue; // Already counted a success for this day
      } else {
        // A single failure on a day breaks the streak
        return 0;
      }
    } else if (getDaysBetweenDates(entryDate, lastDate) === 1) {
      if (success) {
        streak++;
      } else {
        return 0;
      }
    } else {
      // Gap in dates, streak is over
      return streak;
    }
    lastDate = entryDate;
  }
  return streak;
}

/**
 * Calculates the success rate within a specific window.
 * @param {Array<Array>} data The tracking data for a single habit.
 * @param {Date} startDate The start of the tracking window.
 * @param {Date} endDate The end of the tracking window.
 * @return {number} The success rate percentage.
 */
function calculateSuccessRate(data, startDate, endDate) {
  const relevantData = data.filter(entry => {
    const entryDate = new Date(entry[0]);
    return entryDate >= startDate && entryDate <= endDate;
  });
  
  if (relevantData.length === 0) return 0;
  
  let successCount = 0;
  let totalCount = 0;
  
  // Use a map to handle multiple entries per day correctly for simple percentage
  const uniqueDates = new Map();
  
  relevantData.forEach(entry => {
    const dateKey = entry[0].toISOString().split('T')[0];
    const success = entry[5];
    if (!uniqueDates.has(dateKey)) {
      uniqueDates.set(dateKey, success);
    } else if (uniqueDates.get(dateKey) === false) {
      // If a previous entry for the day was a failure, a success overrides it for percentage calculation
      uniqueDates.set(dateKey, success);
    }
  });
  
  uniqueDates.forEach(isSuccess => {
    totalCount++;
    if (isSuccess) successCount++;
  });
  
  return (successCount / totalCount) * 100;
}

/**
 * Gets a rolling 30-day view of success/failure.
 * @param {Array<Array>} data The tracking data for a single habit.
 * @return {Array<string>} An array of success indicators (✓/✗) for the last 30 days.
 */
function getThirtyDayView(data) {
  const view = [];
  const today = new Date();
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(today.getDate() - 29);
  
  const datesMap = new Map();
  data.forEach(entry => {
    const entryDate = new Date(entry[0]);
    const dateKey = entryDate.toISOString().split('T')[0];
    const success = entry[5];
    // A failure on a day with a success will overwrite it for the view
    datesMap.set(dateKey, success);
  });
  
  for (let i = 29; i >= 0; i--) {
    const currentDate = new Date();
    currentDate.setDate(today.getDate() - i);
    const dateKey = currentDate.toISOString().split('T')[0];
    
    if (datesMap.has(dateKey)) {
      view.push(datesMap.get(dateKey) ? '✓' : '✗');
    } else {
      view.push('-');
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
 * Applies conditional formatting to the dashboard for visual indicators.
 * @param {Sheet} sheet The dashboard sheet.
 * @param {number} numRows Number of rows to format.
 * @param {number} numCols Number of columns to format.
 */
function applyConditionalFormatting(sheet, numRows, numCols) {
  const range = sheet.getRange(2, 6, numRows - 1, numCols - 5);
  
  const ruleSuccess = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('✓')
    .setBackground('#b6d7a8') // Light green
    .setBold(true)
    .setFontColor('#073763')
    .setRanges([range])
    .build();
    
  const ruleFailure = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('✗')
    .setBackground('#ea9999') // Light red
    .setBold(true)
    .setFontColor('#cc0000')
    .setRanges([range])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(ruleSuccess, ruleFailure);
  sheet.setConditionalFormatRules(rules);
}