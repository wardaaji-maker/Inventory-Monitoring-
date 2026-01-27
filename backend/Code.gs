// Code.gs - Modified for your structure

// Configuration
const CONFIG = {
  SPREADSHEET_ID: '1r8MbuBCBx2dTkfk7BVgNhxaZrPWPY6EvsKQBmE176VM',
  SHEET_NAME: 'Leads', // Your sheet name
  MAX_FOLLOW_UPS: 10, // Maximum follow-up date columns
  FOLLOW_UP_DATE_COLUMNS: ['L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U'] // Columns for follow-up dates
};

// Expected Headers
const HEADERS = [
  'No',
  'Client ID',
  'Client Name',
  'Store Code',
  'Phone Number',
  'Email',
  'Address',
  'Created At',
  'PIC',
  'Status',
  'Progress'
];
// Add Follow Up headers
for (let i = 1; i <= CONFIG.MAX_FOLLOW_UPS; i++) {
  HEADERS.push(`Follow Up date ${i}`);
}

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
  PROGRESS: 10,    // K
  FOLLOW_UP_1: 11, // L
  // Continue for more follow-up columns...
};

function checkAndRepairHeaders() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  const currentHeadersRange = sheet.getRange(1, 1, 1, HEADERS.length);
  const currentHeaders = currentHeadersRange.getValues()[0];

  let needsRepair = false;
  if (currentHeaders.length !== HEADERS.length) {
    needsRepair = true;
  } else {
    for (let i = 0; i < HEADERS.length; i++) {
      if (currentHeaders[i] !== HEADERS[i]) {
        needsRepair = true;
        break;
      }
    }
  }

  if (needsRepair) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    console.log('Headers repaired');
  }
}

// Main web app
function doGet() {
  checkAndRepairHeaders();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Client Follow-up Monitor')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Get paginated clients with filtering
function getClientsPaginated(page = 1, pageSize = 50, search = '', statusFilter = '') {
  const allClients = getAllClients();
  let filteredClients = allClients;

  // Apply search
  if (search) {
    const searchLower = search.toLowerCase();
    filteredClients = filteredClients.filter(client =>
      (client.clientId && client.clientId.toString().toLowerCase().includes(searchLower)) ||
      (client.clientName && client.clientName.toString().toLowerCase().includes(searchLower)) ||
      (client.storeCode && client.storeCode.toString().toLowerCase().includes(searchLower)) ||
      (client.phone && client.phone.toString().includes(searchLower)) ||
      (client.email && client.email.toString().toLowerCase().includes(searchLower))
    );
  }

  // Apply status filter
  if (statusFilter) {
    filteredClients = filteredClients.filter(client => client.status === statusFilter);
  }

  // Pagination
  const total = filteredClients.length;
  const totalPages = Math.ceil(total / pageSize);
  const startIndex = (page - 1) * pageSize;
  const paginatedClients = filteredClients.slice(startIndex, startIndex + pageSize);

  return {
    clients: paginatedClients,
    total: total,
    page: page,
    totalPages: totalPages,
    pageSize: pageSize
  };
}

// Get all clients with enhanced data
function getAllClients() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow-1, 11 + CONFIG.MAX_FOLLOW_UPS).getValues();

  const clients = data.map((row, index) => {
    // Collect all follow-up dates
    const followUpDates = [];
    for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
      if (row[COLUMNS.FOLLOW_UP_1 + i]) {
        followUpDates.push(row[COLUMNS.FOLLOW_UP_1 + i]);
      }
    }

    // Find next follow-up date (most recent future date)
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let nextFollowUp = null;
    let overdueFollowUps = 0;

    followUpDates.forEach(date => {
      if (date instanceof Date) {
        const followUpDate = new Date(date);
        followUpDate.setHours(0, 0, 0, 0);

        if (followUpDate < today) {
          overdueFollowUps++;
        } else if (!nextFollowUp || followUpDate < nextFollowUp) {
          nextFollowUp = followUpDate;
        }
      }
    });

    // If no future dates, find the most recent past date
    if (!nextFollowUp && followUpDates.length > 0) {
      const pastDates = followUpDates
        .filter(date => date instanceof Date)
        .sort((a, b) => b - a);
      if (pastDates.length > 0) {
        nextFollowUp = pastDates[0];
      }
    }

    return {
      row: index + 2, // +2 because we start from row 2 and arrays are 0-indexed
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
      progress: row[COLUMNS.PROGRESS],
      followUpDates: followUpDates,
      nextFollowUp: nextFollowUp,
      overdueFollowUps: overdueFollowUps,
      totalFollowUps: followUpDates.filter(d => d).length
    };
  }).filter(client => client.clientId); // Filter out empty rows

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

    // Check if follow-up is today or overdue
    return followUpDate <= today;
  });

  return dueClients;
}

// Get overdue clients
function getOverdueClients() {
  const clients = getAllClients();
  return clients.filter(client => client.overdueFollowUps > 0);
}

// Get single client details
function getClientDetails(row) {
  const clients = getAllClients(); // Inefficient but safe. Could optimize to read single row.
  return clients.find(c => c.row === row);
}

// Record a new follow-up
function recordFollowup(clientRow, followupData) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // OPTIMIZATION: Read all follow-up columns in one go
  const range = sheet.getRange(clientRow, COLUMNS.FOLLOW_UP_1 + 1, 1, CONFIG.MAX_FOLLOW_UPS);
  const followUpValues = range.getValues()[0];

  // Find the next empty follow-up column
  let followUpColumnIndex = -1; // 0-based index relative to the range
  for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
    if (!followUpValues[i]) {
      followUpColumnIndex = i;
      break;
    }
  }

  let targetColumn;
  if (followUpColumnIndex !== -1) {
      // Found an empty slot
      // Calculate actual sheet column index (1-based)
      targetColumn = COLUMNS.FOLLOW_UP_1 + 1 + followUpColumnIndex;
  } else {
    // All follow-up columns are filled, find the earliest one to overwrite
    let earliestDate = new Date(9999, 11, 31);
    let earliestIndex = 0;

    for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
      const date = followUpValues[i];
      if (date && new Date(date) < earliestDate) {
        earliestDate = new Date(date);
        earliestIndex = i;
      }
    }
    targetColumn = COLUMNS.FOLLOW_UP_1 + 1 + earliestIndex;
  }

  // Record the follow-up date
  const followupDate = new Date(followupData.date);
  sheet.getRange(clientRow, targetColumn).setValue(followupDate);

  // Update status and progress if provided
  if (followupData.status) {
    sheet.getRange(clientRow, COLUMNS.STATUS + 1).setValue(followupData.status);
  }

  if (followupData.progress) {
    sheet.getRange(clientRow, COLUMNS.PROGRESS + 1).setValue(followupData.progress);
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
    progress: followupData.progress
  });

  // Send email notification if requested
  if (followupData.sendEmail) {
    sendFollowupNotification(clientRow, followupData);
  }

  return {
    success: true,
    followupDate: followupDate,
    column: targetColumn
  };
}

// Log follow-up activity in a separate sheet
function logFollowupActivity(activity) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let logSheet = ss.getSheetByName('Follow-up Log');

  if (!logSheet) {
    logSheet = ss.insertSheet('Follow-up Log');
    logSheet.getRange(1, 1, 1, 8).setValues([[
      'Timestamp', 'Client ID', 'Client Name', 'Follow-up Date',
      'Type', 'Notes', 'PIC', 'Status', 'Progress'
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
    activity.progress || ''
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
    <p><strong>Client ID:</strong> ${clientData.id}</p>
    <p><strong>Store Code:</strong> ${clientData.storeCode}</p>
    <p><strong>Follow-up Date:</strong> ${formatDate(followupData.date)}</p>
    <p><strong>Type:</strong> ${followupData.type}</p>
    <p><strong>Notes:</strong> ${followupData.notes}</p>
    <p><strong>Status Updated:</strong> ${followupData.status || 'No change'}</p>
    <p><strong>Progress Updated:</strong> ${followupData.progress || 'No change'}</p>
    <p><strong>Recorded by:</strong> ${Session.getActiveUser().getEmail()}</p>
  `;

  // Send to assigned PIC and the person who recorded it
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
  const today = new Date();
  const dueClients = getClientsNeedingFollowup();

  // Categorize by status
  const statusCounts = {};
  clients.forEach(client => {
    const status = client.status || 'No Status';
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  });

  // Count overdue follow-ups
  const overdueClients = clients.filter(client => client.overdueFollowUps > 0);

  const stats = {
    totalClients: clients.length,
    dueFollowups: dueClients.length,
    overdueClients: overdueClients.length,
    statusCounts: statusCounts,
    avgFollowUps: clients.length > 0 ?
      (clients.reduce((sum, client) => sum + client.totalFollowUps, 0) / clients.length).toFixed(1) : 0
  };

  return stats;
}

// Update client status
function updateClientStatus(clientRow, status, progress) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (status) {
    sheet.getRange(clientRow, COLUMNS.STATUS + 1).setValue(status);
  }

  if (progress) {
    sheet.getRange(clientRow, COLUMNS.PROGRESS + 1).setValue(progress);
  }

  return { success: true };
}

// Get follow-up history for a client
function getClientFollowupHistory(clientRow) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  const followUpDates = [];
  for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
    const date = sheet.getRange(clientRow, COLUMNS.FOLLOW_UP_1 + i + 1).getValue();
    if (date) {
      followUpDates.push({
        date: date,
        sequence: i + 1
      });
    }
  }

  // Sort by date, most recent first
  followUpDates.sort((a, b) => new Date(b.date) - new Date(a.date));

  return followUpDates;
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
          <th>Phone</th>
        </tr>
    `;

    picClients.forEach(client => {
      const phoneLink = client.phone ?
        `<a href="tel:${client.phone}">${client.phone}</a>` : 'N/A';

      htmlBody += `
        <tr>
          <td>${client.clientId}</td>
          <td><strong>${client.clientName}</strong></td>
          <td>${client.storeCode}</td>
          <td>${client.status || 'No Status'}</td>
          <td style="color: #e74c3c;"><strong>${formatDate(client.nextFollowUp)}</strong></td>
          <td>${phoneLink}</td>
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

// Helper function to format dates
function formatDate(date) {
  if (!date) return 'Not scheduled';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MMM dd, yyyy');
}

// Export data for reporting
function exportClientData(format = 'csv') {
  const clients = getAllClients();

  if (format === 'csv') {
    const headers = [
      'Client ID', 'Client Name', 'Store Code', 'Phone', 'Email',
      'Status', 'Progress', 'PIC', 'Next Follow-up', 'Total Follow-ups',
      'Overdue Follow-ups', 'Created At'
    ];

    const csvData = clients.map(client => [
      client.clientId,
      `"${client.clientName}"`,
      client.storeCode,
      client.phone,
      client.email,
      client.status,
      client.progress,
      client.pic,
      formatDate(client.nextFollowUp),
      client.totalFollowUps,
      client.overdueFollowUps,
      formatDate(client.createdAt)
    ]);

    const csvContent = [headers, ...csvData]
      .map(row => row.join(','))
      .join('\n');

    return {
      content: csvContent,
      filename: `clients_export_${new Date().toISOString().slice(0,10)}.csv`,
      mimeType: 'text/csv'
    };
  }

  return { error: 'Unsupported format' };
}
