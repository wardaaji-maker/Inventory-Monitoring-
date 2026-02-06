// Code.gs - Modified for your structure

// Configuration
const CONFIG = {
  SPREADSHEET_ID: '1r8MbuBCBx2dTkfk7BVgNhxaZrPWPY6EvsKQBmE176VM',
  SHEET_NAME: 'Leads', // Your sheet name
  MAX_FOLLOW_UPS: 10, // Maximum follow-up date columns
  FOLLOW_UP_START_COL: 10 // Column K (0-indexed 10), since Content was index 10 and is now removed from main cols
};

// Column indices based on your structure
// New Structure: ... Status (9), FollowUp1_Date (10), FollowUp1_Feedback (11), FollowUp1_Content (12) ...
const COLUMNS = {
  NO: 0,           // A
  CLIENT_ID: 1,    // B
  CLIENT_NAME: 2,  // C
  STORE_CODE: 3,   // D
  PHONE: 4,        // E
  EMAIL: 5,        // F
  ADDRESS: 6,      // G
  CREATED_AT: 7,   // H
  PIC: 8,          // I
  STATUS: 9,       // J
  // Content column removed from main area
  FOLLOW_UP_START: 10 // K
};

// Main web app
function doGet() {
  // Inspect and repair headers on load
  checkAndRepairHeaders();

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Client Follow-up Monitor')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Automatically inspect header and auto repair if wrong
function checkAndRepairHeaders() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    }

    // Check if sheet has data
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

      // Heuristic: If "Progress" is at Column K (Index 10) OR "Follow Up date 1" is at Column L (Index 11)
      // This indicates the old structure (packed dates starting at L, Progress at K)
      const isOldStructure = (currentHeaders.length > 10 && (currentHeaders[10] === 'Progress' || currentHeaders[10] === 'J')) ||
                             (currentHeaders.length > 11 && (currentHeaders[11] === 'Follow Up date 1' || currentHeaders[11] === 'Follow Up Date 1'));

      if (isOldStructure) {
        Logger.log('Detected Old Structure. Migrating Data...');
        migrateData(sheet);
        return; // Migration handles header setting
      }
    }

    // Standard Repair Logic (just overwrites headers if needed)
    const headers = generateHeaders();
    const requiredCols = headers.length;

    if (sheet.getMaxColumns() < requiredCols) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
    }

    const range = sheet.getRange(1, 1, 1, requiredCols);
    const currentHeaders = range.getValues()[0];

    let needsRepair = false;
    for (let i = 0; i < headers.length; i++) {
      if (currentHeaders[i] !== headers[i]) {
        needsRepair = true;
        break;
      }
    }

    if (needsRepair) {
      range.setValues([headers]);
      Logger.log('Headers repaired');
    }

    // Check for Settings sheet
    let settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('Settings');
      settingsSheet.getRange('A1').setValue('Promotion Types');
      settingsSheet.getRange('A2').setValue('Monthly Sale');
      settingsSheet.getRange('A3').setValue('New Arrival');
      settingsSheet.getRange('A4').setValue('Special Offer');
    }

  } catch (e) {
    Logger.log('Error checking headers: ' + e.toString());
  }
}

function generateHeaders() {
  const headers = [
    'No', 'Client ID', 'Client Name', 'Store Code', 'Phone Number',
    'Email', 'Address', 'Created At', 'PIC', 'Status'
  ];

  // Add follow-up headers (Date + Feedback + Content triplets)
  for (let i = 1; i <= CONFIG.MAX_FOLLOW_UPS; i++) {
    headers.push(`Follow Up Date ${i}`);
    headers.push(`Feedback ${i}`);
    headers.push(`Content ${i}`);
  }
  return headers;
}

function migrateData(sheet) {
   const lastRow = sheet.getLastRow();
   if (lastRow <= 1) { // Only headers
      const newHeaders = generateHeaders();
      const requiredCols = newHeaders.length;
      if (sheet.getMaxColumns() < requiredCols) {
         sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
      }
      sheet.getRange(1, 1, 1, requiredCols).setValues([newHeaders]);
      return;
   }

   const lastCol = sheet.getLastColumn();
   // Read all data (including headers to be safe with indices, but we skip row 1)
   const oldData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

   // Transform Data
   const newData = oldData.map(row => {
      // Static cols 0-9 (A-J). Note: Old K(10) was Progress, discard it.
      const newRow = [];
      for(let i=0; i<=9; i++) {
          newRow[i] = (row[i] !== undefined) ? row[i] : '';
      }

      // Old Dates started at L(11) (Index 11)
      // New Dates start at K(10) with stride 3
      let oldDateIdx = 11;
      let newTripletIdx = 10;

      // Process packed dates
      while(oldDateIdx < row.length) {
         const dateVal = row[oldDateIdx];
         if (dateVal) {
            newRow[newTripletIdx] = dateVal; // Date
            newRow[newTripletIdx + 1] = '';  // Feedback (Empty)
            newRow[newTripletIdx + 2] = '';  // Content (Empty)
         }
         oldDateIdx++; // Move to next packed date
         newTripletIdx += 3; // Move to next triplet
      }
      return newRow;
   });

   // Clear old data
   sheet.clearContents();

   // Set New Headers
   const newHeaders = generateHeaders();
   const requiredCols = newHeaders.length;

   if (sheet.getMaxColumns() < requiredCols) {
     sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
   }

   sheet.getRange(1, 1, 1, requiredCols).setValues([newHeaders]);

   // Write Data
   if (newData.length > 0) {
     const maxLen = Math.max(...newData.map(r => r.length));
     const normalizedData = newData.map(r => {
        while(r.length < maxLen) r.push('');
        return r;
     });

     sheet.getRange(2, 1, normalizedData.length, maxLen).setValues(normalizedData);
   }

   SpreadsheetApp.flush();
   Logger.log('Migration Completed.');
}

// Get Promotion Types from Settings
function getPromotionTypes() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) return [];

  const lastRow = settingsSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = settingsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map(r => r[0]).filter(String);
}

// Get all clients with enhanced data
function getAllClients() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const totalCols = 10 + (CONFIG.MAX_FOLLOW_UPS * 3);
  // Ensure we don't read beyond existing columns if sheet is small (though checkHeaders should fix it)
  const actualCols = Math.min(totalCols, sheet.getLastColumn());

  const data = sheet.getRange(2, 1, lastRow-1, actualCols).getValues();

  const clients = data.map((row, index) => {
    // Collect all follow-up dates, feedback, and content
    const followUps = [];
    let latestFollowUpDate = null;
    let latestFeedback = '';
    let latestContent = '';

    for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
      const dateColIdx = COLUMNS.FOLLOW_UP_START + (i * 3);
      const feedbackColIdx = dateColIdx + 1;
      const contentColIdx = dateColIdx + 2;

      // Safety check for indices
      if (dateColIdx >= row.length) break;

      const dateVal = row[dateColIdx];
      const feedbackVal = (feedbackColIdx < row.length) ? row[feedbackColIdx] : '';
      const contentVal = (contentColIdx < row.length) ? row[contentColIdx] : '';

      if (dateVal && dateVal instanceof Date) {
        followUps.push({
          date: dateVal,
          feedback: feedbackVal || '',
          content: contentVal || ''
        });

        if (!latestFollowUpDate || dateVal > latestFollowUpDate) {
          latestFollowUpDate = dateVal;
          latestFeedback = feedbackVal || '';
          latestContent = contentVal || '';
        }
      }
    }

    // Find next follow-up date logic
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let nextFollowUp = null;
    let overdueFollowUps = 0;

    if (latestContent === 'TERHUBUNG') {
        nextFollowUp = null;
    } else {
        if (latestFollowUpDate) {
             nextFollowUp = new Date(latestFollowUpDate);
             nextFollowUp.setDate(nextFollowUp.getDate() + 4);
             nextFollowUp.setHours(0,0,0,0);
        } else if (row[COLUMNS.CREATED_AT] instanceof Date) {
             nextFollowUp = new Date(row[COLUMNS.CREATED_AT]);
             nextFollowUp.setDate(nextFollowUp.getDate() + 4);
             nextFollowUp.setHours(0,0,0,0);
        }
    }

    if (nextFollowUp && nextFollowUp < today) {
         overdueFollowUps = 1;
    }

    return {
      row: index + 2,
      no: row[COLUMNS.NO],
      clientId: row[COLUMNS.CLIENT_ID],
      clientName: row[COLUMNS.CLIENT_NAME],
      storeCode: row[COLUMNS.STORE_CODE],
      phone: row[COLUMNS.PHONE],
      email: row[COLUMNS.EMAIL],
      address: row[COLUMNS.ADDRESS],
      createdAt: row[COLUMNS.CREATED_AT],
      pic: row[COLUMNS.PIC],
      status: row[COLUMNS.STATUS],
      content: latestContent, // Derived from latest follow-up
      followUps: followUps,
      latestFeedback: latestFeedback,
      nextFollowUp: nextFollowUp,
      overdueFollowUps: overdueFollowUps,
      totalFollowUps: followUps.length
    };
  }).filter(client => client.clientId);

  return clients;
}

// Get clients needing follow-up today
function getClientsNeedingFollowup() {
  const clients = getAllClients();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const dueClients = clients.filter(client => {
    if (!client.nextFollowUp) return false;
    const followUpDate = new Date(client.nextFollowUp);
    followUpDate.setHours(0, 0, 0, 0);
    return followUpDate <= today;
  });

  return dueClients;
}

// Helper function to save a single follow-up
function saveFollowupInternal(sheet, clientRow, followupData) {
  // Find the next empty follow-up slot
  let dateColumn = null;
  let feedbackColumn = null;
  let contentColumn = null;

  for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
    const colIdx = COLUMNS.FOLLOW_UP_START + (i * 3);
    const cellValue = sheet.getRange(clientRow, colIdx + 1).getValue();
    if (!cellValue) {
      dateColumn = colIdx + 1;
      feedbackColumn = colIdx + 2;
      contentColumn = colIdx + 3;
      break;
    }
  }

  if (!dateColumn) {
    const lastIdx = COLUMNS.FOLLOW_UP_START + ((CONFIG.MAX_FOLLOW_UPS - 1) * 3);
    dateColumn = lastIdx + 1;
    feedbackColumn = lastIdx + 2;
    contentColumn = lastIdx + 3;
  }

  // Record Date, Feedback, Content
  const followupDate = new Date(followupData.date);
  sheet.getRange(clientRow, dateColumn).setValue(followupDate);
  sheet.getRange(clientRow, feedbackColumn).setValue(followupData.notes);
  sheet.getRange(clientRow, contentColumn).setValue(followupData.content);

  // Update status
  if (followupData.status) {
    sheet.getRange(clientRow, COLUMNS.STATUS + 1).setValue(followupData.status);
  }

  // Log the follow-up activity
  logFollowupActivity({
    clientId: sheet.getRange(clientRow, COLUMNS.CLIENT_ID + 1).getValue(),
    clientName: sheet.getRange(clientRow, COLUMNS.CLIENT_NAME + 1).getValue(),
    date: followupDate,
    type: followupData.type,
    notes: followupData.notes,
    pic: Session.getActiveUser().getEmail(),
    status: followupData.status,
    content: followupData.content
  });

  // Send email notification if requested
  if (followupData.sendEmail) {
    sendFollowupNotification(sheet, clientRow, followupData);
  }
}

// Record a new follow-up (Single)
function recordFollowup(clientRow, followupData) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  saveFollowupInternal(sheet, clientRow, followupData);

  SpreadsheetApp.flush();

  return { success: true };
}

// Record bulk follow-up
function recordBulkFollowup(clientRows, followupData) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  for (let i = 0; i < clientRows.length; i++) {
    saveFollowupInternal(sheet, clientRows[i], followupData);
  }

  SpreadsheetApp.flush();
  return { success: true, count: clientRows.length };
}

// Log follow-up activity in a separate sheet
function logFollowupActivity(activity) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let logSheet = ss.getSheetByName('Follow-up Log');

  if (!logSheet) {
    logSheet = ss.insertSheet('Follow-up Log');
    logSheet.getRange(1, 1, 1, 9).setValues([[
      'Timestamp', 'Client ID', 'Client Name', 'Follow-up Date',
      'Type', 'Notes', 'PIC', 'Status', 'Content'
    ]]);
  }

  const timestamp = new Date();
  logSheet.appendRow([
    timestamp,
    activity.clientId,
    activity.clientName,
    activity.date,
    activity.type,
    activity.notes,
    activity.pic,
    activity.status || '',
    activity.content || ''
  ]);
}

// Send follow-up notification email
function sendFollowupNotification(sheet, clientRow, followupData) {
  const clientData = {
    name: sheet.getRange(clientRow, COLUMNS.CLIENT_NAME + 1).getValue(),
    id: sheet.getRange(clientRow, COLUMNS.CLIENT_ID + 1).getValue(),
    storeCode: sheet.getRange(clientRow, COLUMNS.STORE_CODE + 1).getValue(),
    phone: sheet.getRange(clientRow, COLUMNS.PHONE + 1).getValue(),
    email: sheet.getRange(clientRow, COLUMNS.EMAIL + 1).getValue(),
    pic: sheet.getRange(clientRow, COLUMNS.PIC + 1).getValue()
  };

  const subject = `Follow-up Recorded: ${clientData.name}`;
  const htmlBody = `
    <h2>Follow-up Recorded</h2>
    <p><strong>Client:</strong> ${clientData.name}</p>
    <p><strong>Follow-up Date:</strong> ${formatDate(followupData.date)}</p>
    <p><strong>Notes:</strong> ${followupData.notes}</p>
    <p><strong>Content:</strong> ${followupData.content || 'No change'}</p>
  `;

  const recipients = [];
  if (clientData.pic) recipients.push(clientData.pic);
  recipients.push(Session.getActiveUser().getEmail());

  GmailApp.sendEmail(recipients.join(','), subject, '', {
    htmlBody: htmlBody,
    name: 'Client Follow-up System'
  });
}

// Get dashboard statistics
function getDashboardStats() {
  const clients = getAllClients();
  const dueClients = getClientsNeedingFollowup();
  const overdueClients = clients.filter(client => client.overdueFollowUps > 0);

  // Helper for counting
  const countBy = (items, keyFn) => {
    const counts = {};
    items.forEach(item => {
      const key = keyFn(item) || 'Unknown';
      counts[key] = (counts[key] || 0) + 1;
    });
    return counts;
  };

  // Helper for aggregation
  const aggregate = (keyFn) => {
    const groups = {};
    clients.forEach(client => {
      const key = keyFn(client) || 'Unassigned';
      if (!groups[key]) {
        groups[key] = {
           total: 0,
           followedUp: 0,
           overdue: 0,
           content: {},
           feedback: {}
        };
      }

      const g = groups[key];
      g.total++;
      if (client.totalFollowUps > 0) g.followedUp++;
      if (client.overdueFollowUps > 0) g.overdue++;

      const c = client.content || 'No Content';
      g.content[c] = (g.content[c] || 0) + 1;

      const f = client.latestFeedback || 'No Feedback';
      g.feedback[f] = (g.feedback[f] || 0) + 1;
    });

    // Add Contribution %
    Object.values(groups).forEach(g => {
       g.contribution = g.total > 0 ? ((g.followedUp / g.total) * 100).toFixed(1) : 0;
    });

    return groups;
  };

  const stats = {
    totalClients: clients.length,
    dueFollowups: dueClients.length,
    overdueClients: overdueClients.length,
    statusCounts: countBy(clients, c => c.status),
    contentCounts: countBy(clients, c => c.content),
    feedbackCounts: countBy(clients, c => c.latestFeedback),
    storeStats: aggregate(c => c.storeCode),
    picStats: aggregate(c => c.pic),
    avgFollowUps: clients.length > 0 ?
      (clients.reduce((sum, client) => sum + client.totalFollowUps, 0) / clients.length).toFixed(1) : 0
  };

  return stats;
}

// Get follow-up history for a client
function getClientFollowupHistory(clientRow) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  const followUps = [];
  for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
    const colIdx = COLUMNS.FOLLOW_UP_START + (i * 3);
    const date = sheet.getRange(clientRow, colIdx + 1).getValue();
    const feedback = sheet.getRange(clientRow, colIdx + 2).getValue();
    const content = sheet.getRange(clientRow, colIdx + 3).getValue();

    if (date && date instanceof Date) {
      followUps.push({
        date: date.toISOString(),
        feedback: feedback || '',
        content: content || '',
        sequence: i + 1
      });
    }
  }

  followUps.sort((a, b) => new Date(b.date) - new Date(a.date));
  return followUps;
}

// Helper function to format dates
function formatDate(date) {
  if (!date) return 'Not scheduled';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MMM dd, yyyy');
}

// Export data
function exportClientData() {
  const clients = getAllClients();
  const headers = [
    'Client ID', 'Client Name', 'Store Code', 'Phone', 'Email',
    'Status', 'Content', 'PIC', 'Next Follow-up', 'Latest Feedback'
  ];

  const csvData = clients.map(client => [
    client.clientId,
    `"${client.clientName}"`,
    client.storeCode,
    client.phone,
    client.email,
    client.status,
    client.content,
    client.pic,
    formatDate(client.nextFollowUp),
    `"${client.latestFeedback}"`
  ]);

  const csvContent = [headers, ...csvData].map(row => row.join(',')).join('\n');
  return {
    content: csvContent,
    filename: `clients_export_${new Date().toISOString().slice(0,10)}.csv`,
    mimeType: 'text/csv'
  };
}

// Schedule daily reminders
function scheduleDailyReminders() {
  const clients = getClientsNeedingFollowup();

  if (clients.length === 0) return { sent: 0 };

  // Group by PIC
  const clientsByPIC = {};
  clients.forEach(client => {
    const pic = client.pic || Session.getActiveUser().getEmail();
    if (!clientsByPIC[pic]) clientsByPIC[pic] = [];
    clientsByPIC[pic].push(client);
  });

  // Send email to each PIC
  Object.keys(clientsByPIC).forEach(pic => {
    const picClients = clientsByPIC[pic];
    const subject = `Follow-up Reminder: ${picClients.length} clients need attention`;

    let htmlBody = `
      <h2>Follow-up Reminder</h2>
      <p>Dear ${pic.split('@')[0]},</p>
      <p>You have ${picClients.length} clients requiring follow-up:</p>
      <table border="1" cellpadding="5" style="border-collapse: collapse;">
        <tr style="background-color: #f2f2f2;">
          <th>Client ID</th>
          <th>Client Name</th>
          <th>Store Code</th>
          <th>Status</th>
          <th>Next Follow-up</th>
          <th>Content</th>
        </tr>
    `;

    picClients.forEach(client => {
      htmlBody += `
        <tr>
          <td>${client.clientId}</td>
          <td><strong>${client.clientName}</strong></td>
          <td>${client.storeCode}</td>
          <td>${client.status || 'No Status'}</td>
          <td style="color: #e74c3c;"><strong>${formatDate(client.nextFollowUp)}</strong></td>
          <td>${client.content || 'N/A'}</td>
        </tr>
      `;
    });

    htmlBody += `
      </table>
      <p style="margin-top: 20px;">
        <a href="${ScriptApp.getService().getUrl()}" style="
          background-color: #3498db;
          color: white;
          padding: 10px 20px;
          text-decoration: none;
          border-radius: 5px;
          display: inline-block;
        ">Open Follow-up System</a>
      </p>
    `;

    GmailApp.sendEmail(pic, subject, '', {
      htmlBody: htmlBody,
      name: 'Client Follow-up System'
    });
  });

  return { sent: Object.keys(clientsByPIC).length };
}

// Create daily reminder trigger
function createDailyReminderTrigger() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'scheduleDailyReminders') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger for 8 AM daily
  ScriptApp.newTrigger('scheduleDailyReminders')
    .timeBased()
    .atHour(8)
    .nearMinute(30)
    .everyDays(1)
    .create();
}
