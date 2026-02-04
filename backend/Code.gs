// Code.gs - Modified for your structure

// Configuration
const CONFIG = {
  SPREADSHEET_ID: '1r8MbuBCBx2dTkfk7BVgNhxaZrPWPY6EvsKQBmE176VM',
  SHEET_NAME: 'Leads', // Your sheet name
  MAX_FOLLOW_UPS: 10, // Maximum follow-up date columns
  FOLLOW_UP_START_COL: 11 // Column L (0-indexed 11)
};

// Column indices based on your structure
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
  CONTENT: 10,     // K (Renamed from PROGRESS)
  FOLLOW_UP_START: 11 // L
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

    // Define expected headers
    const headers = [
      'No', 'Client ID', 'Client Name', 'Store Code', 'Phone Number',
      'Email', 'Address', 'Created At', 'PIC', 'Status', 'Content'
    ];

    // Add follow-up headers (Date + Feedback pairs)
    for (let i = 1; i <= CONFIG.MAX_FOLLOW_UPS; i++) {
      headers.push(`Follow Up Date ${i}`);
      headers.push(`Feedback ${i}`);
    }

    // Get current headers
    const lastCol = headers.length;

    // Ensure sufficient columns
    if (sheet.getMaxColumns() < lastCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), lastCol - sheet.getMaxColumns());
    }

    const range = sheet.getRange(1, 1, 1, lastCol);
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

  const totalCols = 11 + (CONFIG.MAX_FOLLOW_UPS * 2);
  const data = sheet.getRange(2, 1, lastRow-1, totalCols).getValues();

  const clients = data.map((row, index) => {
    // Collect all follow-up dates and feedback
    const followUps = [];
    let latestFollowUpDate = null;
    let latestFeedback = '';

    for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
      const dateColIdx = COLUMNS.FOLLOW_UP_START + (i * 2);
      const feedbackColIdx = dateColIdx + 1;

      const dateVal = row[dateColIdx];
      const feedbackVal = row[feedbackColIdx];

      if (dateVal && dateVal instanceof Date) {
        followUps.push({
          date: dateVal,
          feedback: feedbackVal || ''
        });

        if (!latestFollowUpDate || dateVal > latestFollowUpDate) {
          latestFollowUpDate = dateVal;
          latestFeedback = feedbackVal || '';
        }
      }
    }

    // Find next follow-up date logic
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let nextFollowUp = null;
    let overdueFollowUps = 0;

    if (latestFollowUpDate) {
         nextFollowUp = new Date(latestFollowUpDate);
         nextFollowUp.setDate(nextFollowUp.getDate() + 4); // Example logic: +4 days
         nextFollowUp.setHours(0,0,0,0);
    } else if (row[COLUMNS.CREATED_AT] instanceof Date) {
         nextFollowUp = new Date(row[COLUMNS.CREATED_AT]);
         nextFollowUp.setDate(nextFollowUp.getDate() + 4);
         nextFollowUp.setHours(0,0,0,0);
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
      content: row[COLUMNS.CONTENT],
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

// Record a new follow-up
function recordFollowup(clientRow, followupData) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // Find the next empty follow-up slot
  let dateColumn = null;
  let feedbackColumn = null;

  for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
    const colIdx = COLUMNS.FOLLOW_UP_START + (i * 2);
    const cellValue = sheet.getRange(clientRow, colIdx + 1).getValue();
    if (!cellValue) {
      dateColumn = colIdx + 1;
      feedbackColumn = colIdx + 2;
      break;
    }
  }

  if (!dateColumn) {
    const lastIdx = COLUMNS.FOLLOW_UP_START + ((CONFIG.MAX_FOLLOW_UPS - 1) * 2);
    dateColumn = lastIdx + 1;
    feedbackColumn = lastIdx + 2;
  }

  // Record Date and Feedback
  const followupDate = new Date(followupData.date);
  sheet.getRange(clientRow, dateColumn).setValue(followupDate);
  sheet.getRange(clientRow, feedbackColumn).setValue(followupData.notes);

  // Update status and content
  if (followupData.status) {
    sheet.getRange(clientRow, COLUMNS.STATUS + 1).setValue(followupData.status);
  }

  if (followupData.content) {
    sheet.getRange(clientRow, COLUMNS.CONTENT + 1).setValue(followupData.content);
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
    sendFollowupNotification(clientRow, followupData);
  }

  SpreadsheetApp.flush();

  return { success: true };
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
function sendFollowupNotification(clientRow, followupData) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

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
  // Since we have data in the row, we could parse it from getAllClients logic,
  // but to be safe and simple, we read again or just implement based on structure
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  const followUps = [];
  for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
    const colIdx = COLUMNS.FOLLOW_UP_START + (i * 2);
    const date = sheet.getRange(clientRow, colIdx + 1).getValue();
    const feedback = sheet.getRange(clientRow, colIdx + 2).getValue();

    if (date && date instanceof Date) {
      followUps.push({
        date: date.toISOString(),
        feedback: feedback || '',
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
  // ... (keep logic simple or same)
  return { sent: 0 };
}
