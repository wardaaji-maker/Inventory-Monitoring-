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

// --- CORE DATA FUNCTIONS ---

// Fetch raw data with row indices, filtering empty rows
function getRawClientData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  // Get all values at once
  const data = sheet.getRange(2, 1, lastRow-1, 11 + CONFIG.MAX_FOLLOW_UPS).getValues();

  // Map to wrapper object to preserve row index, filtering out empty client IDs
  return data
    .map((row, index) => ({
      data: row,
      rowIndex: index + 2
    }))
    .filter(item => item.data[COLUMNS.CLIENT_ID]);
}

// Process a single raw row into a rich client object
function processClientRow(rowData, rowIndex) {
  const row = rowData;

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
    row: rowIndex,
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
}

// Optimized Pagination
function getClientsPaginated(page = 1, pageSize = 50, search = '', statusFilter = '') {
  // 1. Get raw data (fast)
  let rawItems = getRawClientData();

  // 2. Filter raw data (fast, string comparisons)
  if (search) {
    const searchLower = search.toLowerCase();
    rawItems = rawItems.filter(item => {
      const row = item.data;
      return (
        (row[COLUMNS.CLIENT_ID] && row[COLUMNS.CLIENT_ID].toString().toLowerCase().includes(searchLower)) ||
        (row[COLUMNS.CLIENT_NAME] && row[COLUMNS.CLIENT_NAME].toString().toLowerCase().includes(searchLower)) ||
        (row[COLUMNS.STORE_CODE] && row[COLUMNS.STORE_CODE].toString().toLowerCase().includes(searchLower)) ||
        (row[COLUMNS.PHONE] && row[COLUMNS.PHONE].toString().includes(searchLower)) ||
        (row[COLUMNS.EMAIL] && row[COLUMNS.EMAIL].toString().toLowerCase().includes(searchLower))
      );
    });
  }

  if (statusFilter) {
    rawItems = rawItems.filter(item => item.data[COLUMNS.STATUS] === statusFilter);
  }

  // 3. Pagination calculation
  const total = rawItems.length;
  const totalPages = Math.ceil(total / pageSize);

  // Adjust page if out of bounds
  if (page < 1) page = 1;
  if (page > totalPages && totalPages > 0) page = totalPages;

  const startIndex = (page - 1) * pageSize;
  const slicedItems = rawItems.slice(startIndex, startIndex + pageSize);

  // 4. Process only the sliced rows (heavy lifting)
  const paginatedClients = slicedItems.map(item => processClientRow(item.data, item.rowIndex));

  return {
    clients: paginatedClients,
    total: total,
    page: page,
    totalPages: totalPages,
    pageSize: pageSize
  };
}

// Optimized Dashboard Stats
function getDashboardStats() {
  const rawItems = getRawClientData();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let dueCount = 0;
  let overdueCount = 0;
  let totalFollowUps = 0;
  const statusCounts = {};

  // Iterate raw data once
  for (const item of rawItems) {
    const row = item.data;

    // Status count
    const status = row[COLUMNS.STATUS] || 'No Status';
    statusCounts[status] = (statusCounts[status] || 0) + 1;

    // Date Logic - Optimization: avoid creating full object
    // Just find nextFollowUp and overdue status for this row
    let rowHasOverdue = false;
    let rowNextFollowUp = null;
    let rowFollowUpsCount = 0;

    // Scan follow-up columns
    for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
      const val = row[COLUMNS.FOLLOW_UP_1 + i];
      if (val && val instanceof Date) {
        rowFollowUpsCount++;
        const date = new Date(val);
        date.setHours(0, 0, 0, 0);

        if (date < today) {
          rowHasOverdue = true;
        } else if (!rowNextFollowUp || date < rowNextFollowUp) {
          rowNextFollowUp = date;
        }
      }
    }

    // Check overdue
    if (rowHasOverdue) overdueCount++;

    // Check due today
    if (rowNextFollowUp) {
       // If next follow up is today or earlier (but if earlier it's also overdue,
       // but typically due means 'needs action now'. If it's overdue, it needs action now too.
       // The original logic for 'Due' was: nextFollowUp <= today.
       if (rowNextFollowUp <= today) {
         dueCount++;
       }
    } else {
        // Fallback: if no future date, check if there was a past date (which would be the nextFollowUp logic fallback)
        // In the full logic: "If no future dates, find the most recent past date" -> set that as nextFollowUp.
        // If that date <= today (which it is), then it counts as due.
        // So effectively, if there are ANY dates, and no future dates, it is due/overdue.
        if (rowFollowUpsCount > 0 && !rowNextFollowUp) {
             // Logic match: nextFollowUp becomes the most recent past date.
             // Since it is past, it is < today. So it is due.
             dueCount++;
        }
    }

    totalFollowUps += rowFollowUpsCount;
  }

  return {
    totalClients: rawItems.length,
    dueFollowups: dueCount,
    overdueClients: overdueCount,
    statusCounts: statusCounts,
    avgFollowUps: rawItems.length > 0 ? (totalFollowUps / rawItems.length).toFixed(1) : 0
  };
}

// Revert getAllClients to use the optimized flow if needed,
// or keep it for legacy/export but use the new processing function
function getAllClients() {
  const rawItems = getRawClientData();
  return rawItems.map(item => processClientRow(item.data, item.rowIndex));
}

// Get clients needing follow-up today (Optimized)
function getClientsNeedingFollowup() {
  // This one still needs to return rich objects, but we can filter first?
  // Actually, deciding if it needs follow-up requires the date logic.
  // So we must process rows. But maybe we can do it lazily?
  // For now, let's use the full process but strictly for this view.
  // If the user has 5000 due clients, this will still be slow.
  // But typically due clients are a subset.

  // Better approach: Calculate criteria on raw data first
  const rawItems = getRawClientData();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const dueClients = [];

  for (const item of rawItems) {
    const row = item.data;
    let isDue = false;

    // Inline date logic check
    let nextFollowUp = null;
    let hasDates = false;

    for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
      const val = row[COLUMNS.FOLLOW_UP_1 + i];
      if (val && val instanceof Date) {
        hasDates = true;
        const date = new Date(val);
        date.setHours(0, 0, 0, 0);
        if (date >= today) {
           if (!nextFollowUp || date < nextFollowUp) {
             nextFollowUp = date;
           }
        }
      }
    }

    if (nextFollowUp && nextFollowUp.getTime() === today.getTime()) {
      isDue = true;
    } else if (!nextFollowUp && hasDates) {
      // If all dates are past, it is effectively due/overdue
      isDue = true;
    } else if (nextFollowUp && nextFollowUp < today) {
       // Should be covered by !nextFollowUp check if we only looked for future dates,
       // but let's be safe.
       isDue = true;
    }

    if (isDue) {
      dueClients.push(processClientRow(item.data, item.rowIndex));
    }
  }

  return dueClients;
}

// Get overdue clients (Optimized)
function getOverdueClients() {
  const rawItems = getRawClientData();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const overdueClients = [];

  for (const item of rawItems) {
    const row = item.data;
    let hasOverdue = false;

    for (let i = 0; i < CONFIG.MAX_FOLLOW_UPS; i++) {
       const val = row[COLUMNS.FOLLOW_UP_1 + i];
       if (val && val instanceof Date) {
         const date = new Date(val);
         date.setHours(0, 0, 0, 0);
         if (date < today) {
           hasOverdue = true;
           break;
         }
       }
    }

    if (hasOverdue) {
      overdueClients.push(processClientRow(item.data, item.rowIndex));
    }
  }
  return overdueClients;
}

// Get single client details
function getClientDetails(row) {
  // Fetch just that row if possible, but getRawClientData reads all.
  // Optimization: Read specific row directly.
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // Row index is 1-based.
  if (row < 2 || row > sheet.getLastRow()) return null;

  const values = sheet.getRange(row, 1, 1, 11 + CONFIG.MAX_FOLLOW_UPS).getValues()[0];

  if (!values[COLUMNS.CLIENT_ID]) return null;

  return processClientRow(values, row);
}

// Record a new follow-up
function recordFollowup(clientRow, followupData) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // Read all follow-up columns in one go
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
