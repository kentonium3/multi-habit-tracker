/**
 * @fileoverview Functions for generating and sending accountability emails.
 * It's responsible for the content and delivery of email reports.
 */

/**
 * Checks if an email should be sent based on the last sent timestamp and frequency.
 * @param {Date} lastSentTimestamp The timestamp of the last email sent.
 * @param {string} frequency The configured email frequency (e.g., Daily).
 * @return {boolean} True if an email should be sent.
 */
function shouldSendEmail(lastSentTimestamp, frequency) {
  const now = new Date();
  const timeDiff = now.getTime() - lastSentTimestamp.getTime();
  const dayInMillis = 1000 * 60 * 60 * 24;

  switch (frequency) {
    case 'Daily':
      return timeDiff >= dayInMillis;
    case 'Weekly':
      return timeDiff >= dayInMillis * 7;
    case 'Bi-weekly':
      return timeDiff >= dayInMillis * 14;
    default:
      return false;
  }
}

/**
 * Sends a clean, easy-to-read HTML email report to accountability friends.
 */
function sendAccountabilityEmail() {
  const config = getConfig();
  if (config.accountabilityEmails.length === 0) {
    log('WARN', 'No accountability emails configured. Skipping email send.');
    return;
  }
  
  if (config.debugMode) {
    log('INFO', 'Debug mode is active. Email will not be sent.');
    return;
  }

  log('INFO', 'Generating and sending accountability email...');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD);
  
  // Get dashboard data to build the HTML table
  const dashboardData = dashboardSheet.getDataRange().getValues();
  if (dashboardData.length <= 1) {
    log('WARN', 'Dashboard is empty. No email will be sent.');
    return;
  }

  // Get the latest comment to include in the email
  const trackingSheet = ss.getSheetByName(SHEET_NAMES.TRACKING);
  const comments = getRecentComments(trackingSheet);

  const subject = 'Your Daily Habit Tracker Update';
  
  const htmlBody = generateHtmlEmail(dashboardData, comments);

  try {
    MailApp.sendEmail({
      to: config.accountabilityEmails.join(','),
      subject: subject,
      htmlBody: htmlBody
    });

    // Update the last sent timestamp in the config sheet
    const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    configSheet.getRange('B3').setValue(new Date());

    log('INFO', 'Accountability email sent successfully.');
  } catch (error) {
    log('ERROR', 'Failed to send accountability email:', error.message);
  }
}

/**
 * Generates the HTML body for the email report.
 * @param {Array<Array>} dashboardData The data from the dashboard sheet.
 * @param {string} comments The latest user comments.
 * @return {string} The complete HTML string for the email body.
 */
function generateHtmlEmail(dashboardData, comments) {
  let html = `
    <html>
      <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #4CAF50;">Daily Habit Progress Report</h2>
        <p>Hello,</p>
        <p>Here is the latest update on your habit tracking. Keep up the great work!</p>
  `;

  if (comments) {
    html += `
      <div style="margin-top: 20px; padding: 10px; border: 1px solid #ccc; background-color: #f9f9f9;">
        <p><strong>Latest Comment:</strong></p>
        <p>${comments}</p>
      </div>
    `;
  }

  html += `
    <h3 style="margin-top: 20px;">Habit Dashboard</h3>
    <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
      <thead>
        <tr style="background-color: #f2f2f2;">
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Habit Name</th>
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Days Since Start</th>
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Days Remaining</th>
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Current Streak</th>
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Success %</th>
          <th colspan="7" style="padding: 8px; border: 1px solid #ddd; text-align: center;">Last 7 Days</th>
        </tr>
        <tr>
          <td colspan="5"></td>
          ${dashboardData[0].slice(-7).map(date => `<th style="padding: 8px; border: 1px solid #ddd; text-align: center;">${date}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
  `;

  // Filter out the header row
  const bodyData = dashboardData.slice(1);

  // Loop through rows to build the table body
  bodyData.forEach(row => {
    html += `<tr>`;
    // Habit details (first 5 columns)
    html += row.slice(0, 5).map(cell => `<td style="padding: 8px; border: 1px solid #ddd;">${cell}</td>`).join('');
    // Last 7 days view
    html += row.slice(-7).map(cell => {
      let cellColor = '';
      if (cell === '✓') cellColor = '#b6d7a8';
      else if (cell === '✗') cellColor = '#ea9999';
      return `<td style="padding: 8px; border: 1px solid #ddd; background-color: ${cellColor}; text-align: center;">${cell}</td>`;
    }).join('');
    html += `</tr>`;
  });

  html += `
      </tbody>
    </table>
    <p style="margin-top: 20px; font-size: 12px; color: #888;">
      This email was generated automatically by your Habit Tracker system.
    </p>
  </body>
  </html>
  `;
  
  return html;
}

/**
 * Retrieves the most recent comment from the tracking sheet.
 * @param {Sheet} trackingSheet The Daily_Tracking sheet object.
 * @return {string} The latest comment, or an empty string if none is found.
 */
function getRecentComments(trackingSheet) {
  const lastRow = trackingSheet.getLastRow();
  if (lastRow < 2) return '';
  
  const lastEntry = trackingSheet.getRange(lastRow, 1, 1, trackingSheet.getLastColumn()).getValues()[0];
  const comment = lastEntry[6]; // Comments are in the 7th column (index 6)
  return comment || '';
}