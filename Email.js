/**
 * @fileoverview Enhanced email functions with robust thread management and error handling.
 * Incorporates proven patterns from gsheet-update-triggered-email.js
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
 * Sends a robust, thread-managed HTML email report to accountability friends.
 * Uses proven thread management and error handling patterns.
 */
function sendAccountabilityEmail() {
  try {
    const config = getConfig();
    
    // Validate email configuration
    if (!config.accountabilityEmails || config.accountabilityEmails.length === 0) {
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
    
    // Validate dashboard data
    if (!dashboardSheet) {
      log('ERROR', 'Dashboard sheet not found. Cannot send email.');
      return;
    }
    
    const dashboardData = dashboardSheet.getDataRange().getDisplayValues();
    if (dashboardData.length <= 1) {
      log('WARN', 'Dashboard is empty. No email will be sent.');
      return;
    }

    // Get the latest comment to include in the email
    const trackingSheet = ss.getSheetByName(SHEET_NAMES.TRACKING);
    const comments = getRecentComments(trackingSheet);

    const subject = 'Your Daily Habit Tracker Update';
    const recipientEmail = config.accountabilityEmails.join(',');
    
    // Generate enhanced HTML email
    const htmlBody = generateEnhancedHtmlEmail(dashboardData, comments);

    // Send email with robust thread management
    sendEmailWithThreadManagement(recipientEmail, subject, htmlBody);

    // Update the last sent timestamp in the config sheet
    updateLastEmailSentTimestamp();

    log('INFO', 'Accountability email sent successfully with thread management.');
    
  } catch (error) {
    log('ERROR', 'Failed to send accountability email:', error.message, error.stack);
  }
}

/**
 * Sends email with robust thread management and error handling.
 * Maintains conversation threads and handles thread recovery.
 * @param {string} recipientEmail The recipient email address(es)
 * @param {string} subject The email subject
 * @param {string} htmlBody The HTML email body
 */
function sendEmailWithThreadManagement(recipientEmail, subject, htmlBody) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const currentThreadId = properties.getProperty('habitTrackerThreadId');
    let thread = null;
    
    // Check if a thread ID exists and try to find the thread using robust search
    if (currentThreadId) {
      log('DEBUG', `Searching for existing thread ID: ${currentThreadId}`);
      
      // Use the search function to find the thread by ID, regardless of its labels
      // This is more reliable than getThreadById()
      const threads = GmailApp.search("thread:" + currentThreadId);
      if (threads.length > 0) {
        thread = threads[0];
        log('DEBUG', `Found existing thread with ${thread.getMessageCount()} messages`);
      } else {
        log('INFO', `Thread ${currentThreadId} not found. Will create new thread.`);
      }
    }

    if (thread) {
      // A thread ID exists and is valid, so reply to the existing thread
      log('INFO', `Replying to existing thread ID: ${currentThreadId}`);
      
      const messages = thread.getMessages();
      const lastMessageId = messages[messages.length - 1].getId();
      const messageIds = messages.map(msg => msg.getId());

      GmailApp.sendEmail(recipientEmail, subject, "", {
        htmlBody: htmlBody,
        inReplyTo: lastMessageId,
        references: messageIds.join(" ")
      });
      
      log('INFO', 'Successfully replied to existing email thread');
      
    } else {
      // The thread was not found, so create a new one
      if (currentThreadId) {
        // Store the previous thread ID for debugging purposes
        properties.setProperty('lastKnownHabitTrackerThreadId', currentThreadId);
        log('INFO', `Previous thread ID (${currentThreadId}) stored for debugging. Starting new thread.`);
      } else {
        log('INFO', 'No previous thread ID found. Starting new thread.');
      }
      
      // Use createDraft and send() to reliably get the thread ID
      const draft = GmailApp.createDraft(recipientEmail, subject, "", {htmlBody: htmlBody});
      const newThread = draft.send().getThread();
      const newThreadId = newThread.getId();
      
      properties.setProperty('habitTrackerThreadId', newThreadId);
      log('INFO', `Created new thread with ID: ${newThreadId}`);
    }
    
  } catch (error) {
    log('ERROR', `Error in thread management: ${error.toString()}`, error.stack);
    
    // Fallback to simple email if thread management fails
    try {
      log('INFO', 'Attempting fallback email without thread management...');
      GmailApp.sendEmail(recipientEmail, subject, "", {htmlBody: htmlBody});
      log('INFO', 'Fallback email sent successfully');
    } catch (fallbackError) {
      log('ERROR', `Fallback email also failed: ${fallbackError.toString()}`);
      throw fallbackError;
    }
  }
}

/**
 * Generates enhanced HTML email with improved styling and formatting.
 * @param {Array<Array>} dashboardData The data from the dashboard sheet
 * @param {string} comments The latest user comments
 * @return {string} The complete HTML string for the email body
 */
function generateEnhancedHtmlEmail(dashboardData, comments) {
  let html = `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th { background-color: #4CAF50; color: white; padding: 8px; border: 1px solid #ddd; text-align: left; }
          .even-row { background-color: #f0f0f0; }
          .odd-row { background-color: #ffffff; }
          td { padding: 8px; border: 1px solid #cccccc; }
          .comment-section { margin-top: 20px; padding: 10px; border: 1px solid #ccc; background-color: #f9f9f9; }
          .success { background-color: #b6d7a8; font-weight: bold; text-align: center; }
          .failure { background-color: #ea9999; font-weight: bold; text-align: center; }
          .neutral { text-align: center; }
        </style>
      </head>
      <body>
        <h2 style="color: #4CAF50;">Daily Habit Progress Report</h2>
        <p>Hello,</p>
        <p>Here is your latest habit tracking update. Keep up the great work!</p>
  `;

  if (comments) {
    html += `
      <div class="comment-section">
        <p><strong>Latest Comment:</strong></p>
        <p>${comments}</p>
      </div>
    `;
  }

  html += `
    <h3 style="margin-top: 20px;">Habit Dashboard</h3>
    <table>
      <thead>
        <tr>
          <th>Habit Name</th>
          <th>Days Since Start</th>
          <th>Days Remaining</th>
          <th>Current Streak</th>
          <th>Success %</th>
          <th colspan="7">Last 7 Days</th>
        </tr>
        <tr>
          <td colspan="5"></td>
          ${dashboardData[0].slice(-7).map(date => `<th>${date}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
  `;

  // Filter out the header row and create table body with alternating colors
  const bodyData = dashboardData.slice(1);
  
  bodyData.forEach((row, rowIndex) => {
    const rowClass = rowIndex % 2 === 0 ? 'even-row' : 'odd-row';
    html += `<tr class="${rowClass}">`;
    
    // Habit details (first 5 columns)
    html += row.slice(0, 5).map(cell => `<td>${cell}</td>`).join('');
    
    // Last 7 days view with enhanced styling
    html += row.slice(-7).map(cell => {
      let cellClass = 'neutral';
      if (cell === '✓') cellClass = 'success';
      else if (cell === '✗') cellClass = 'failure';
      return `<td class="${cellClass}">${cell}</td>`;
    }).join('');
    
    html += `</tr>`;
  });

  html += `
      </tbody>
    </table>
    <p style="margin-top: 20px; font-size: 12px; color: #888;">
      This email was generated automatically by your Habit Tracker system at ${new Date().toLocaleString()}.
    </p>
  </body>
  </html>
  `;
  
  return html;
}

/**
 * Updates the last email sent timestamp in the Config sheet.
 */
function updateLastEmailSentTimestamp() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    
    if (!configSheet) {
      log('ERROR', 'Config sheet not found. Cannot update last email timestamp.');
      return;
    }
    
    // Find the LastEmailSent row and update it
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'LastEmailSent') {
        configSheet.getRange(i + 1, 2).setValue(new Date());
        log('DEBUG', 'Updated LastEmailSent timestamp');
        return;
      }
    }
    
    // If not found, add it
    configSheet.appendRow(['LastEmailSent', new Date()]);
    log('DEBUG', 'Added LastEmailSent timestamp');
    
  } catch (error) {
    log('ERROR', 'Failed to update last email timestamp:', error.message);
  }
}

/**
 * Retrieves the most recent comment from the tracking sheet.
 * @param {Sheet} trackingSheet The Daily_Tracking sheet object.
 * @return {string} The latest comment, or an empty string if none is found.
 */
function getRecentComments(trackingSheet) {
  try {
    if (!trackingSheet) {
      log('WARN', 'Tracking sheet not found for comments');
      return '';
    }
    
    const lastRow = trackingSheet.getLastRow();
    if (lastRow < 2) return '';
    
    const lastEntry = trackingSheet.getRange(lastRow, 1, 1, trackingSheet.getLastColumn()).getValues()[0];
    const comment = lastEntry[6]; // Comments are in the 7th column (index 6)
    return comment || '';
    
  } catch (error) {
    log('ERROR', 'Error retrieving recent comments:', error.message);
    return '';
  }
}

/**
 * Manual function to test email functionality
 * Available through custom menu for debugging
 */
function testEmailFunctionality() {
  try {
    log('INFO', 'Testing email functionality...');
    sendAccountabilityEmail();
    log('INFO', 'Email test completed. Check execution logs for results.');
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Email Test Complete', 'Check the execution logs to see the results of the email test.', ui.ButtonSet.OK);
    
  } catch (error) {
    log('ERROR', 'Email test failed:', error.message);
    const ui = SpreadsheetApp.getUi();
    ui.alert('Email Test Failed', `Error: ${error.message}`, ui.ButtonSet.OK);
  }
}