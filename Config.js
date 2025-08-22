/**
 * @fileoverview Manages system configuration and project properties.
 */

/**
 * Reads the configuration settings from the 'Config' sheet.
 * @return {Object} An object containing all configuration settings.
 */
function getConfig() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
  if (!configSheet) {
    throw new Error(`Config sheet not found. Please ensure a sheet named '${SHEET_NAMES.CONFIG}' exists.`);
  }

  const values = configSheet.getDataRange().getValues();
  const config = {};
  for (let i = 0; i < values.length; i++) {
    const key = values[i][0];
    const value = values[i][1];
    switch (key) {
      case 'AccountabilityEmails':
        config.accountabilityEmails = value.split(',').map(email => email.trim()).filter(email => email);
        break;
      case 'EmailFrequency':
        config.emailFrequency = value;
        break;
      case 'LastEmailSent':
        config.lastEmailSent = value;
        break;
      case 'DebugMode':
        // Handle both string and boolean values
        if (typeof value === 'boolean') {
          config.debugMode = value;
        } else if (typeof value === 'string') {
          config.debugMode = value.toLowerCase() === 'true';
        } else {
          config.debugMode = false; // Default fallback
        }
        break;
      default:
        break;
    }
  }
  return config;
}

/**
 * Creates and populates the Config sheet with default values.
 * This is run only once during the initial setup.
 */
function createConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_NAMES.CONFIG);
    log('INFO', `Created new sheet: ${SHEET_NAMES.CONFIG}`);
  }

  const headers = ['Setting', 'Value'];
  const data = [
    ['AccountabilityEmails', ''],
    ['EmailFrequency', 'Daily'],
    ['LastEmailSent', ''],
    ['DebugMode', 'FALSE']
  ];

  configSheet.getRange(1, 1, data.length, 2).setValues(data);
  configSheet.getRange('B2').setNote('Daily, Weekly, or Bi-weekly');
  configSheet.setColumnWidth(1, 200);
}