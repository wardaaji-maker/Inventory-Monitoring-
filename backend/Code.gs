const CONFIG = {
  sheetId: '1RG1rfZ-eKAkUOCmvOrlUoq45tfXGGZFNm2VvDgJAad4', // Updated spreadsheet ID
  sheets: {
    leads: 'Leads',
    followups: 'FollowUpHistory',
    transactions: 'Transaction',
    settings: 'Settings',
    templates: 'MessageTemplates',
    monthlyLeadSummary: 'MonthlyLeadSummary',
    dailyDashboardStats: 'DailyDashboardStats',
    teamPerformanceDashboard: 'TeamPerformanceDashboard'
  }
};

function getTeamMembers() {
  return ['HERU', 'FAHMY', 'HILLARY', 'DWI PUJI', 'RIZKY M.', 'MELLINDA', 'NANA', 'SYAFEI', 'RIZKI PIK', 'DESI', 'DIMAS', 'DANU', 'CISCO', 'DESTIAN'];
}

// New function to get status options
function getStatusOptions() {
  return ['New', 'In Progress', 'Prospecting', 'Convincing', 'Negotiation', 'Won', 'Lost', 'Drop'];
}

// New function to get source options
function getSourceOptions() {
  return ['Website', 'Social Media', 'Referral', 'Advertisement', 'Walk In', 'Other'];
}

// New function to get showroom options
function getShowroomOptions() {
  return ['South78', 'PIM', 'SCT', 'PIK 2'];
}

// Enhanced getSheet function that works with your spreadsheet
function getSheet(name) {
  try {
    // Validate input parameter
    if (!name || typeof name !== 'string') {
      throw new Error(`Invalid sheet name: ${name}`);
    }

    // Get the specific spreadsheet by ID
    const ss = SpreadsheetApp.openById(CONFIG.sheetId);

    // Get the sheet name from configuration
    const sheetName = CONFIG.sheets[name];
    if (!sheetName) {
      throw new Error(`Sheet configuration for '${name}' not found. Available: ${Object.keys(CONFIG.sheets).join(', ')}`);
    }

    // Try to get the sheet
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`Sheet '${sheetName}' not found. Creating it now...`);
      sheet = ss.insertSheet(sheetName);

      // Set up headers for the new sheet
      setupSheetHeaders(sheetName, sheet);
      Logger.log(`Created and initialized sheet: ${sheetName}`);
    }

    return sheet;

  } catch (e) {
    Logger.log(`Error in getSheet('${name}'): ${e.toString()}`);
    throw e;
  }
}

// Initialize the application for your spreadsheet
function initApp() {
  try {
    Logger.log('Initializing application for your spreadsheet...');

    const ss = SpreadsheetApp.openById(CONFIG.sheetId);
    Logger.log(`Connected to spreadsheet: ${ss.getName()}`);

    setupSheets(); // Ensures sheets and headers exist
    setupDefaultData(); // Adds default settings and templates

    // Add sample leads if the Leads sheet is empty
    addSampleData(); // Assuming you have this function from previous script

    Logger.log('Application initialized successfully!');
    return {
      success: true,
      message: 'App initialized successfully for your spreadsheet',
      spreadsheet: ss.getName()
    };

  } catch (e) {
    Logger.log('Error in initApp: ' + e.toString());
    return {
      success: false,
      error: e.toString(),
      suggestion: 'Make sure the spreadsheet URL is correct and you have edit permissions'
    };
  }
}

// You need to include the addSampleData function in this Code.gs file if it's not present.
function addSampleData() {
  try {
    console.log('Adding sample data...');
    const sampleLeads = [
      ['L001', 'John Doe', '+1234567890', 'john@test.com', 'Website', 'South78', 'Product A', new Date(), 'New', new Date(Date.now() + 3 * 24 * 60 * 60 * 1000), 0, 'Interested in pricing', 'Team', new Date()],
      ['L002', 'Jane Smith', '+1234567891', 'jane@test.com', 'Social Media', 'PIM', 'Product B', new Date(), 'In Progress', new Date(Date.now() + 1 * 24 * 60 * 60 * 1000), 1, 'Follow up tomorrow', 'Team', new Date()],
      ['L003', 'Mike Johnson', '+1234567892', 'mike@test.com', 'Referral', 'SCT', 'Product C', new Date(), 'New', new Date(Date.now() + 5 * 24 * 60 * 60 * 1000), 0, 'Hot lead', 'Team', new Date()]
    ];

    const sheet = getSheet('leads');
    if (sheet.getLastRow() <= 1) { // Check if only headers or empty
      sheet.getRange(2, 1, sampleLeads.length, sampleLeads[0].length).setValues(sampleLeads);
      console.log('Added sample leads successfully');
      return { success: true, message: 'Added 3 sample leads' };
    } else {
      console.log('Leads sheet already has data');
      return { success: false, message: 'Leads sheet already contains data' };
    }
  } catch (error) {
    console.error('Error adding sample data:', error);
    return { success: false, error: error.toString() };
  }
}

// Set up all required sheets
function setupSheets() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.sheetId);

    Object.keys(CONFIG.sheets).forEach(sheetKey => {
      const sheetName = CONFIG.sheets[sheetKey];
      let sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        Logger.log(`Created sheet: ${sheetName}`);
      }

      // Set up headers
      setupSheetHeaders(sheetName, sheet);
    });

  } catch (e) {
    Logger.log('Error in setupSheets: ' + e.toString());
    throw e;
  }
}

// Set up headers for each sheet type
function setupSheetHeaders(sheetName, sheet) {
  try {
    // Clear existing data if needed (only if empty)
    if (sheet.getLastRow() === 0) {
      let headers = [];

      switch(sheetName) {
        case 'Leads':
          headers = ['LeadID', 'Name', 'Phone', 'Email', 'Source', 'Showroom', 'Product', 'CreatedAt', 'Status', 'NextFollowUp', 'FollowUpCount', 'Notes', 'AssignedTo', 'LastContact', 'DealValue', 'ReceiptNumber'];
          break;
        case 'FollowUpHistory':
          headers = ['LeadID', 'Date', 'Type', 'Status', 'Notes', 'User'];
          break;
        case 'Transaction':
          headers = ['LeadID', 'TransactionDate', 'Product', 'Amount', 'PaymentStatus', 'DeliveryStatus', 'Notes'];
          break;
        case 'Settings':
          headers = ['Key', 'Value', 'Description', 'Example'];
          break;
        case 'MessageTemplates':
          headers = ['TemplateName', 'Message', 'Type', 'Step'];
          break;
        case 'MonthlyLeadSummary':
          headers = ['Date', 'LeadID', 'AssignedTo', 'Name', 'Source', 'Phone', 'Note', 'Status', 'DealValue', 'LastContact'];
          break;
        case 'DailyDashboardStats':
            headers = [
              'Date',
              'LeadsCreated_Count', 'LeadsCreated_Value',
              'BecameInProgress_Count', 'BecameInProgress_Value',
              'BecameProspecting_Count', 'BecameProspecting_Value',
              'BecameConvincing_Count', 'BecameConvincing_Value',
              'BecameNegotiation_Count', 'BecameNegotiation_Value',
              'BecameWon_Count', 'BecameWon_Value',
              'BecameLost_Count', 'BecameLost_Value',
              'BecameDrop_Count', 'BecameDrop_Value',
              'LeadsContacted_Count'
            ];
            break;
        case 'TeamPerformanceDashboard':
            headers = [
              'Month', 'TeamMember',
              'NewLeads_Count', 'NewLeads_Value',
              'InProgress_Count', 'InProgress_Value',
              'Prospecting_Count', 'Prospecting_Value',
              'Convincing_Count', 'Convincing_Value',
              'Negotiation_Count', 'Negotiation_Value',
              'Won_Count', 'Won_Value',
              'Lost_Count', 'Lost_Value',
              'Drop_Count', 'Drop_Value'
            ];
            break;
        default:
          return; // No headers for unknown sheets
      }

      if (headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        Logger.log(`Headers set up for ${sheetName}`);
      }
    }
  } catch (e) {
    Logger.log(`Error setting up headers for ${sheetName}: ${e.toString()}`);
  }
}
// Set up default data
function setupDefaultData() {
  try {
    // Setup default settings
    setupDefaultSettings();

    // Setup default templates
    setupDefaultTemplates();

  } catch (e) {
    Logger.log('Error in setupDefaultData: ' + e.toString());
  }
}

function setupDefaultSettings() {
  try {
    const sheet = getSheet('settings');
    if (sheet.getLastRow() <= 1) { // Only headers or empty
      const defaultSettings = [
        ['FOLLOWUP_SCHEDULE', '3,5,7,14', 'Follow-up days sequence', '3,5,7,14'],
        ['AUTO_NOTIFY', 'TRUE', 'Auto notifications', 'TRUE'],
        ['BUSINESS_NAME', 'Your Business Name', 'Business name for messages', 'ABC Company'],
        ['DEFAULT_ASSIGNEE', 'Team', 'Default lead assignee', 'Sales Team'],
        ['NOTIFICATION_EMAIL', '', 'Team notification email', 'team@company.com']
      ];

      sheet.getRange(2, 1, defaultSettings.length, 4).setValues(defaultSettings);
      Logger.log('Default settings added');
    }
  } catch (e) {
    Logger.log('Error in setupDefaultSettings: ' + e.toString());
  }
}

function setupDefaultTemplates() {
  try {
    const sheet = getSheet('templates');
    if (sheet.getLastRow() <= 1) { // Only headers or empty
      const defaultTemplates = [
        ['First Follow-up', 'Hello {name}, this is {business} following up on your inquiry about {product}. Are you still interested?', 'WhatsApp', '1'],
        ['Second Follow-up', 'Hi {name}, just checking in about your interest in {product}. Let me know if you have any questions!', 'WhatsApp', '2'],
        ['Third Follow-up', 'Hello {name}, wanted to follow up once more about {product}. We have some great options available!', 'WhatsApp', '3'],
        ['Closing Follow-up', 'Hi {name}, this will be my last follow-up about {product}. Feel free to reach out if you change your mind!', 'WhatsApp', '4'],
        ['Welcome Message', 'Welcome {name}! Thank you for your interest in {product}. How can I help you?', 'WhatsApp', '0']
      ];

      sheet.getRange(2, 1, defaultTemplates.length, 4).setValues(defaultTemplates);
      Logger.log('Default templates added');
    }
  } catch (e) {
    Logger.log('Error in setupDefaultTemplates: ' + e.toString());
  }
}

// FIXED: getLeads function for your spreadsheet
// ENHANCED: getLeads function with better error handling
// Alias for frontend compatibility if needed
function getActiveLeads() {
  const allLeads = getLeads();
  const today = new Date();
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();

  return allLeads.filter(lead => {
      // Filter Logic:
      // 1. If lead was contacted/created this month (MTD), always show it.
      // 2. If lead is NOT MTD:
      //    - If Status is 'Won', 'Lost', or 'Drop', HIDE it (Historical closed leads).
      //    - If Status is Active (New, In Progress, etc.), SHOW it (Active pipeline/Overdue).

      let activityDate;
      if (lead.LastContact) {
        activityDate = new Date(lead.LastContact);
      } else if (lead.CreatedAt) {
        activityDate = new Date(lead.CreatedAt);
      } else {
        activityDate = new Date(0);
      }

      const isMTD = activityDate.getMonth() === currentMonth && activityDate.getFullYear() === currentYear;
      const isClosed = ['Won', 'Lost', 'Drop'].includes(lead.Status);

      if (isMTD) return true;
      if (!isClosed) return true; // Keep active pipeline regardless of date

      return false; // Hide old closed leads
  });
}

function getLeads() {
  try {
    const sheet = getSheet('leads');
    const lastRow = sheet.getLastRow();

    // If no data (only headers or empty)
    if (!lastRow || lastRow < 2) {
      Logger.log('No leads data found - returning empty array');
      return [];
    }

    const dataRange = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
    const allData = dataRange.getValues();

    const headers = allData[0];
    const data = allData.slice(1); // Remove header row

    if (!data || data.length === 0) {
      return [];
    }

    const leads = data.map(row => {
      const lead = {};
      headers.forEach((header, index) => {
        // Handle date objects properly
        let value = row[index];
        if (value instanceof Date) {
          value = value.toISOString();
        }
        lead[header] = value || '';
      });
      return lead;
    });

    Logger.log(`Loaded ${leads.length} leads successfully`);
    return leads;

  } catch (e) {
    Logger.log('Error in getLeads: ' + e.toString());
    // Return sample data for testing if real data fails
    return getSampleLeadsForTesting();
  }
}

// Fallback function if main data loading fails
function getSampleLeadsForTesting() {
  return [
    {
      LeadID: 'L001',
      Name: 'John Doe',
      Phone: '+1234567890',
      Email: 'john@test.com',
      Source: 'Website',
      Showroom: 'South78',
      Product: 'Product A',
      CreatedAt: new Date().toISOString(),
      Status: 'New',
      NextFollowUp: new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString(),
      FollowUpCount: 0,
      Notes: 'Interested in pricing',
      AssignedTo: 'Team',
      LastContact: new Date().toISOString()
    }
  ];
}

function getDashboardStats() {
  try {
    const leads = getLeads();
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();

    const stats = {
      total: 0,
      new: 0,
      inProgress: 0,
      prospecting: 0,
      convincing: 0,
      negotiation: 0,
      won: 0,
      lost: 0,
      drop: 0,
      dueToday: 0,
      overdue: 0,
      monthlyStats: {
        prospectingValue: 0,
        convincingValue: 0,
        negotiationValue: 0,
        wonValue: 0,
        lostValue: 0,
        dropValue: 0,
        totalValue: 0
      }
    };

    today.setHours(0, 0, 0, 0);

    leads.forEach(lead => {
      const status = lead.Status || 'New';
      const dealValue = parseFloat(lead.DealValue) || 0;

      // Determine the reference date for activity: LastContact > CreatedAt > Epoch
      let activityDate;
      if (lead.LastContact) {
        activityDate = new Date(lead.LastContact);
      } else if (lead.CreatedAt) {
        activityDate = new Date(lead.CreatedAt);
      } else {
        activityDate = new Date(0); // 1970 (inactive)
      }

      const creationDate = lead.CreatedAt ? new Date(lead.CreatedAt) : new Date(0);

      // 1. Total Leads: New Inflow (Created This Month)
      if (creationDate.getMonth() === currentMonth && creationDate.getFullYear() === currentYear) {
          stats.total++;
      }

      const isMTDActivity = activityDate.getMonth() === currentMonth && activityDate.getFullYear() === currentYear;

      // 2. Active Pipeline (Snapshot): Count All Active Leads regardless of date
      if (['New', 'In Progress', 'Prospecting', 'Convincing', 'Negotiation'].includes(status)) {
          if (status === 'New') stats.new++;
          else if (status === 'In Progress') stats.inProgress++;
          else if (status === 'Prospecting') stats.prospecting++;
          else if (status === 'Convincing') stats.convincing++;
          else if (status === 'Negotiation') stats.negotiation++;

          // Accumulate Value for Active Pipeline (Snapshot)
          if (status === 'Prospecting') {
            stats.monthlyStats.prospectingValue += dealValue;
          } else if (status === 'Convincing') {
            stats.monthlyStats.convincingValue += dealValue;
          } else if (status === 'Negotiation') {
            stats.monthlyStats.negotiationValue += dealValue;
          }

          // Total Active Pipeline Value
          if (['Prospecting', 'Convincing', 'Negotiation'].includes(status)) {
            stats.monthlyStats.totalValue += dealValue;
          }
      }

      // 3. Closed Leads (Performance): Count Only MTD Activity
      if (isMTDActivity) {
          if (status === 'Won') {
              stats.won++;
              stats.monthlyStats.wonValue += dealValue;
          } else if (status === 'Lost') {
              stats.lost++;
              stats.monthlyStats.lostValue += dealValue;
          } else if (status === 'Drop') {
              stats.drop++;
              stats.monthlyStats.dropValue += dealValue;
          }
      }

      // Check follow-up dates (Global check - unrelated to monthly stats)
      if (status !== 'Won' && status !== 'Lost' && status !== 'Drop' && lead.NextFollowUp) {
        try {
          const followUpDate = new Date(lead.NextFollowUp);
          followUpDate.setHours(0, 0, 0, 0);

          const daysDiff = Math.floor((today - followUpDate) / (1000 * 60 * 60 * 24));

          if (daysDiff === 0) stats.dueToday++;
          else if (daysDiff > 0) stats.overdue++;
        } catch (e) {
          // Ignore date parsing errors
        }
      }
    });

    // Debug logging to help identify issues
    Logger.log('Dashboard Stats Calculation:');
    Logger.log(`Total leads: ${stats.total}`);
    Logger.log(`Prospecting: ${stats.prospecting} with value: ${stats.monthlyStats.prospectingValue}`);
    Logger.log(`Convincing: ${stats.convincing} with value: ${stats.monthlyStats.convincingValue}`);
    Logger.log(`Negotiation: ${stats.negotiation} with value: ${stats.monthlyStats.negotiationValue}`);
    Logger.log(`Won: ${stats.won} with value: ${stats.monthlyStats.wonValue}`);
    Logger.log(`Total Monthly Value: ${stats.monthlyStats.totalValue}`);

    return stats;

  } catch (e) {
    Logger.log('Error in getDashboardStats: ' + e.toString());
    return {
      total: 0, new: 0, inProgress: 0, prospecting: 0, convincing: 0,
      negotiation: 0, won: 0, lost: 0, drop: 0, dueToday: 0, overdue: 0,
      monthlyStats: {
        prospectingValue: 0,
        convincingValue: 0,
        negotiationValue: 0,
        wonValue: 0,
        lostValue: 0,
        dropValue: 0,
        totalValue: 0
      }
    };
  }
}
// FIXED: addLead function
function addLead(leadData) {
  try {
    if (!leadData || !leadData.name || !leadData.phone) {
      return { success: false, error: 'Name and Phone are required' };
    }

    const sheet = getSheet('leads');
    const leadId = 'L' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const now = new Date();

    const newLead = [
      leadId,
      leadData.name,
      leadData.phone,
      leadData.email || '',
      leadData.source || 'Website',
      leadData.showroom || '', // Added Showroom
      leadData.product || '',
      now,
      'New',
      calculateNextFollowUp(now, 0),
      0,
      leadData.notes || '',
      leadData.assignedTo || 'Team',
      now
    ];

    sheet.appendRow(newLead);

    // Add to history
    addFollowUpHistory(leadId, now, 'Initial Contact', 'Completed', 'Lead created', 'System');

    return { success: true, leadId: leadId };

  } catch (e) {
    Logger.log('Error in addLead: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// --- NEW HELPER FUNCTION: updateLeadNote ---
function updateLeadNote(leadId, note) {
  try {
    if (!leadId || !note) {
      Logger.log(`Skipping note update for LeadID '${leadId}' due to missing data.`);
      return; // Don't throw an error, just log and exit.
    }

    const sheet = getSheet('leads');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const leadIDColIndex = headers.indexOf('LeadID');
    const notesColIndex = headers.indexOf('Notes');

    if (leadIDColIndex === -1 || notesColIndex === -1) {
      throw new Error('Could not find LeadID or Notes column in the Leads sheet.');
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][leadIDColIndex] === leadId) {
        sheet.getRange(i + 1, notesColIndex + 1).setValue(note);
        Logger.log(`Successfully updated note for LeadID '${leadId}'.`);
        return; // Exit after finding and updating the lead.
      }
    }

    Logger.log(`Could not find LeadID '${leadId}' to update note.`);

  } catch (e) {
    Logger.log(`Error in updateLeadNote for LeadID '${leadId}': ${e.toString()}`);
    // We don't re-throw the error as this is a non-critical update.
  }
}

// FIXED: getMessageTemplates
function getMessageTemplates() {
  try {
    const sheet = getSheet('templates');
    const lastRow = sheet.getLastRow();

    if (!lastRow || lastRow < 2) {
      return []; // Return empty array if no templates
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    return data.map(row => {
      const template = {};
      headers.forEach((header, index) => {
        template[header] = row[index] || '';
      });
      return template;
    });

  } catch (e) {
    Logger.log('Error in getMessageTemplates: ' + e.toString());
    return [];
  }
}

// Helper functions
function calculateNextFollowUp(baseDate, followUpCount) {
  const schedule = [3, 5, 7, 14]; // Default schedule
  const daysToAdd = followUpCount < schedule.length ? schedule[followUpCount] : 14;

  const nextDate = new Date(baseDate);
  nextDate.setDate(nextDate.getDate() + daysToAdd);
  return nextDate;
}

function addFollowUpHistory(leadId, date, type, status, notes, user) {
  try {
    const sheet = getSheet('followups');
    sheet.appendRow([leadId, date, type, status, notes, user]);
  } catch (e) {
    Logger.log('Error adding follow-up history: ' + e.toString());
  }
}

// Web app functions
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getCurrentUser() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    return 'Team Member';
  }
}

// Test function for your spreadsheet
function testConnection() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.sheetId);
    const sheets = ss.getSheets();

    Logger.log(`Connected to: ${ss.getName()}`);
    Logger.log(`Sheets found: ${sheets.map(s => s.getName()).join(', ')}`);

    // Test getSheet function
    ['leads', 'settings', 'templates'].forEach(sheet => {
      try {
        const s = getSheet(sheet);
        Logger.log(`✓ ${s.getName()} sheet accessible`);
      } catch (e) {
        Logger.log(`✗ Error accessing ${sheet}: ${e}`);
      }
    });

    return {
      success: true,
      spreadsheet: ss.getName(),
      sheets: sheets.map(s => s.getName())
    };

  } catch (e) {
    Logger.log('Connection test failed: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// FIXED: updateLeadStatus function (frontend calls this as updateStatus)
function updateStatus(leadId, newStatus, dealValue = '', receiptNumber = '', lostReason = '') {
  try {
    if (!leadId || !newStatus) {
      return { success: false, error: 'Lead ID and new status are required.' };
    }

    if (newStatus === 'Lost' && !lostReason) {
      return { success: false, error: 'A reason is required when marking a lead as Lost.' };
    }

    const sheet = getSheet('leads');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Find column indices
    const leadIDColIndex = headers.indexOf('LeadID');
    const statusColIndex = headers.indexOf('Status');
    const lastContactColIndex = headers.indexOf('LastContact');
    const followUpCountColIndex = headers.indexOf('FollowUpCount');
    const nextFollowUpColIndex = headers.indexOf('NextFollowUp');
    const dealValueColIndex = headers.indexOf('DealValue');
    const receiptNumberColIndex = headers.indexOf('ReceiptNumber');

    if (leadIDColIndex === -1 || statusColIndex === -1) {
      throw new Error('Required headers not found in Leads sheet.');
    }

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    let leadFound = false;

    for (let i = 1; i < data.length; i++) {
      // Use loose comparison or string conversion for LeadID to be safe
      if (String(data[i][leadIDColIndex]) === String(leadId)) {
        // Update the row
        sheet.getRange(i + 1, statusColIndex + 1).setValue(newStatus);
        sheet.getRange(i + 1, lastContactColIndex + 1).setValue(new Date());

        // Update NextFollowUp for active statuses to clear "Overdue"
        // Also clear if status is Won, Lost or Drop
        if (newStatus !== 'Won' && newStatus !== 'Lost' && newStatus !== 'Drop' && nextFollowUpColIndex !== -1) {
          const currentFollowUpCount = parseInt(data[i][followUpCountColIndex]) || 0;
          const nextDate = calculateNextFollowUp(new Date(), currentFollowUpCount);
          sheet.getRange(i + 1, nextFollowUpColIndex + 1).setValue(nextDate);
        }

        // Update deal value if provided
        if (dealValue && dealValueColIndex !== -1) {
          sheet.getRange(i + 1, dealValueColIndex + 1).setValue(dealValue);
        }

        // Update receipt number if provided and status is Won
        if (receiptNumber && receiptNumberColIndex !== -1 && newStatus === 'Won') {
          sheet.getRange(i + 1, receiptNumberColIndex + 1).setValue(receiptNumber);
        }

        // If Won, Lost or Drop, clear next follow-up (check if column exists)
        if ((newStatus === 'Won' || newStatus === 'Lost' || newStatus === 'Drop') && nextFollowUpColIndex !== -1) {
          sheet.getRange(i + 1, nextFollowUpColIndex + 1).setValue('');
        }

        leadFound = true;

        // Add history entry
        let historyNotes = `Lead status changed to '${newStatus}'`;
        if (dealValue) historyNotes += ` with deal value: ${dealValue}`;
        if (receiptNumber && newStatus === 'Won') historyNotes += ` | Receipt: ${receiptNumber}`;
        if (lostReason && newStatus === 'Lost') historyNotes += ` | Reason: ${lostReason}`;

        addFollowUpHistory(leadId, new Date(), 'Status Change', newStatus, historyNotes, getCurrentUser());

        Logger.log(`Lead ${leadId} status updated to ${newStatus}`);
        return { success: true, message: `Lead ${leadId} status updated to ${newStatus}` };
      }
    }

    if (!leadFound) {
      return { success: false, error: `Lead with ID '${leadId}' not found.` };
    }

    } catch (e) {
    Logger.log(`Error in updateStatus for LeadID '${leadId}': ${e.toString()}`);
    return { success: false, error: e.toString() };
  }
}
// Add this function to handle the frontend's updateStatus call
function updateLeadDetails(payload) {
    const { leadId, newStatus, dealValue, receiptNumber, reasonForLoss } = payload;
    return updateStatus(leadId, newStatus, dealValue, receiptNumber, reasonForLoss);
}

function updateDealValue(leadId, dealValue) {
  try {
    const sheet = getSheet('leads');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const leadIDColIndex = headers.indexOf('LeadID');
    const dealValueColIndex = headers.indexOf('DealValue');

    if (dealValueColIndex === -1) {
      // Add DealValue column if it doesn't exist
      sheet.getRange(1, headers.length + 1).setValue('DealValue');
      return updateDealValue(leadId, dealValue); // Retry with new column
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][leadIDColIndex] === leadId) {
        sheet.getRange(i + 1, dealValueColIndex + 1).setValue(dealValue);

        // Add to history
        addFollowUpHistory(leadId, new Date(), 'Deal Value Update', 'Completed',
                          `Deal value set to: ${dealValue}`, getCurrentUser());

        return { success: true, message: `Deal value updated for ${leadId}` };
      }
    }

    return { success: false, error: 'Lead not found' };

  } catch (e) {
    Logger.log(`Error updating deal value: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }
}
// --- NEW FUNCTION: getTemplateVariables ---
function getTemplateVariables(leadId) {
  try {
    let variables = {
      name: 'Customer',
      product: 'our product',
      phone: '',
      email: '',
      business: 'Your Business Name' // Default value
    };

    // Get lead data
    if (leadId) {
      const leads = getLeads(); // Re-use existing getLeads
      const lead = leads.find(l => l.LeadID === leadId);
      if (lead) {
        variables.name = lead.Name || variables.name;
        variables.product = lead.Product || variables.product;
        variables.phone = lead.Phone || variables.phone;
        variables.email = lead.Email || variables.email;
      }
    }

    // Get business name from settings
    try {
      const settingsSheet = getSheet('settings');
      const settingsData = settingsSheet.getDataRange().getValues();
      const headers = settingsData[0];
      const keyColIndex = headers.indexOf('Key');
      const valueColIndex = headers.indexOf('Value');

      if (keyColIndex !== -1 && valueColIndex !== -1) {
        for (let i = 1; i < settingsData.length; i++) {
          if (settingsData[i][keyColIndex] === 'BUSINESS_NAME') {
            variables.business = settingsData[i][valueColIndex] || variables.business;
            break;
          }
        }
      }
    } catch (settingError) {
      Logger.log('Could not load BUSINESS_NAME from settings: ' + settingError.toString());
      // Continue with default business name
    }

    Logger.log(`Template variables for ${leadId || 'default'}:`, variables);
    return variables;

  } catch (e) {
    Logger.log(`Error in getTemplateVariables for LeadID '${leadId}': ${e.toString()}`);
    return { name: 'Customer', product: 'our product', business: 'Your Business Name', phone: '', email: '' }; // Return safe defaults
  }
}

// --- NEW FUNCTION: recordFollowUp ---
// FIXED: recordFollowUp function with robust WhatsApp link generation
function recordFollowUp(leadId, followUpData) {
  try {
    if (!leadId || !followUpData || !followUpData.type || !followUpData.message) {
      return { success: false, error: 'Lead ID, follow-up type, and message are required.' };
    }

    const leadsSheet = getSheet('leads');
    const headers = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];

    // Find column indices
    const leadIDColIndex = headers.indexOf('LeadID');
    const statusColIndex = headers.indexOf('Status');
    const followUpCountColIndex = headers.indexOf('FollowUpCount');
    const nextFollowUpColIndex = headers.indexOf('NextFollowUp');
    const lastContactColIndex = headers.indexOf('LastContact');
    const phoneColIndex = headers.indexOf('Phone');
    const nameColIndex = headers.indexOf('Name');

    if ([leadIDColIndex, statusColIndex, followUpCountColIndex, nextFollowUpColIndex, lastContactColIndex, phoneColIndex, nameColIndex].some(idx => idx === -1)) {
      throw new Error('Required headers not found in Leads sheet.');
    }

    const data = leadsSheet.getDataRange().getValues();
    let leadFound = false;
    let currentFollowUpCount = 0;
    let leadPhone = '';
    let leadName = '';
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      // Use string comparison for safety
      if (String(data[i][leadIDColIndex]) === String(leadId)) {
        leadFound = true;
        rowIndex = i;
        leadName = data[i][nameColIndex] || 'Unknown Lead';

        // SAFELY get phone number - handle any data type
        let phoneValue = data[i][phoneColIndex];
        if (phoneValue === null || phoneValue === undefined || phoneValue === '') {
          leadPhone = '';
        } else if (typeof phoneValue === 'number') {
          leadPhone = phoneValue.toString();
        } else if (typeof phoneValue === 'string') {
          leadPhone = phoneValue;
        } else if (phoneValue instanceof Date) {
          leadPhone = ''; // Dates are not valid phone numbers
        } else {
          leadPhone = String(phoneValue); // Fallback for any other type
        }

        // Update Status to 'In Progress' if it's 'New'
        if (data[i][statusColIndex] === 'New') {
          data[i][statusColIndex] = 'In Progress';
        }

        // Increment FollowUpCount
        currentFollowUpCount = (parseInt(data[i][followUpCountColIndex]) || 0) + 1;
        data[i][followUpCountColIndex] = currentFollowUpCount;

        // Calculate and update NextFollowUp date
        data[i][nextFollowUpColIndex] = calculateNextFollowUp(new Date(), currentFollowUpCount);

        // Update LastContact
        data[i][lastContactColIndex] = new Date();

        // Write the updated row back to the sheet
        leadsSheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]);
        break;
      }
    }

    if (!leadFound) {
      return { success: false, error: `Lead with ID '${leadId}' not found.` };
    }

    // Add history entry
    const fullNote = followUpData.notes ? `${followUpData.notes} - Message: ${followUpData.message}` : followUpData.message;
    addFollowUpHistory(leadId, new Date(), followUpData.type, 'Completed', fullNote, getCurrentUser());

    // --- NEW: Update the main lead note ---
    if (followUpData.notes) { // Only update if there are internal notes
      updateLeadNote(leadId, followUpData.notes);
    }

    // FIXED: Generate WhatsApp links with safe phone handling
    let whatsappWebLink = '';
    let whatsappApiLink = '';
    let hasValidPhone = false;

    // Ensure leadPhone is a string and clean it
    if (leadPhone) {
        try {
            // Convert to string first, then clean
            const phoneString = String(leadPhone).trim();
            if (phoneString) {
                const cleanedPhone = phoneString.replace(/[^0-9+]/g, ''); // Keep + for international numbers
                if (cleanedPhone && cleanedPhone.length >= 10) {
                    const encodedMessage = encodeURIComponent(followUpData.message);
                    whatsappWebLink = `https://web.whatsapp.com/send?phone=${cleanedPhone}&text=${encodedMessage}`;
                    whatsappApiLink = `https://api.whatsapp.com/send?phone=${cleanedPhone}&text=${encodedMessage}`;
                    hasValidPhone = true;
                    Logger.log(`WhatsApp links generated for phone: ${cleanedPhone}`);
                } else {
                    Logger.log(`Invalid phone number format for lead ${leadId}: ${phoneString}`);
                }
            } else {
                Logger.log(`Empty phone number for lead ${leadId}`);
            }
        } catch (phoneError) {
            Logger.log(`Error processing phone number for lead ${leadId}: ${phoneError.toString()}`);
        }
    } else {
        Logger.log(`No phone number available for lead ${leadId}`);
    }

    Logger.log(`Follow-up recorded for ${leadId}. WhatsApp available: ${hasValidPhone}`);

    return {
      success: true,
      leadName: leadName,
      followUpCount: currentFollowUpCount,
      hasWhatsApp: hasValidPhone,
      whatsappLink: {
        web: whatsappWebLink,
        api: whatsappApiLink
      }
    };

  } catch (e) {
    Logger.log(`Error in recordFollowUp for LeadID '${leadId}': ${e.toString()}`);
    return { success: false, error: e.toString() };
  }
}
// --- NEW FUNCTION: markAsContacted ---
function markAsContacted(leadId, contactType, notes) {
  try {
    if (!leadId || !contactType) {
      return { success: false, error: 'Lead ID and contact type are required.' };
    }

    const leadsSheet = getSheet('leads');
    const headers = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];

    // Find column indices
    const leadIDColIndex = headers.indexOf('LeadID');
    const statusColIndex = headers.indexOf('Status');
    const followUpCountColIndex = headers.indexOf('FollowUpCount');
    const nextFollowUpColIndex = headers.indexOf('NextFollowUp');
    const lastContactColIndex = headers.indexOf('LastContact');

    if ([leadIDColIndex, statusColIndex, followUpCountColIndex, nextFollowUpColIndex, lastContactColIndex].some(idx => idx === -1)) {
      throw new Error('Required headers not found in Leads sheet. Check: LeadID, Status, FollowUpCount, NextFollowUp, LastContact');
    }

    const data = leadsSheet.getDataRange().getValues();
    let leadFound = false;
    let currentFollowUpCount = 0;

    for (let i = 1; i < data.length; i++) {
      // Use string comparison for safety
      if (String(data[i][leadIDColIndex]) === String(leadId)) {
        leadFound = true;

        // Update Status to 'In Progress' if it's 'New'
        if (data[i][statusColIndex] === 'New') {
          data[i][statusColIndex] = 'In Progress';
        }

        // Increment FollowUpCount
        currentFollowUpCount = (parseInt(data[i][followUpCountColIndex]) || 0) + 1;
        data[i][followUpCountColIndex] = currentFollowUpCount;

        // Calculate and update NextFollowUp date
        data[i][nextFollowUpColIndex] = calculateNextFollowUp(new Date(), currentFollowUpCount);

        // Update LastContact
        data[i][lastContactColIndex] = new Date();

        // Write the updated row back to the sheet
        leadsSheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]);
        break;
      }
    }

    if (!leadFound) {
      return { success: false, error: `Lead with ID '${leadId}' not found.` };
    }

    // Add history entry
    addFollowUpHistory(leadId, new Date(), contactType, 'Completed', notes, getCurrentUser());

    // --- NEW: Update the main lead note ---
    updateLeadNote(leadId, notes);

    Logger.log(`Lead ${leadId} marked as contacted by ${getCurrentUser()}. Next follow-up set.`);
    return { success: true, message: `Lead ${leadId} marked as contacted.` };

  } catch (e) {
    Logger.log(`Error in markAsContacted for LeadID '${leadId}': ${e.toString()}`);
    return { success: false, error: e.toString() };
  }
}

// --- NEW FUNCTION: checkFollowUps ---
function checkFollowUps() {
  try {
    const leads = getLeads(); // Get all leads
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let dueTodayCount = 0;
    let overdueCount = 0;
    const leadsRequiringAttention = []; // To store details for email

    leads.forEach(lead => {
      const status = lead.Status || 'New';
      if (status !== 'Won' && status !== 'Lost' && status !== 'Drop' && lead.NextFollowUp) {
        try {
          const followUpDate = new Date(lead.NextFollowUp);
          followUpDate.setHours(0, 0, 0, 0);
          const daysDiff = Math.floor((today - followUpDate) / (1000 * 60 * 60 * 24));

          if (daysDiff === 0) {
            dueTodayCount++;
            leadsRequiringAttention.push(`• Due Today: ${lead.Name} (${lead.LeadID}) - Assigned: ${lead.AssignedTo}`);
          } else if (daysDiff > 0) {
            overdueCount++;
            leadsRequiringAttention.push(`• OVERDUE (${daysDiff} days): ${lead.Name} (${lead.LeadID}) - Assigned: ${lead.AssignedTo}`);
          }
        } catch (dateError) {
          Logger.log(`Error parsing NextFollowUp date for LeadID ${lead.LeadID}: ${dateError}`);
        }
      }
    });

    // --- Optional: Send Email Notification ---
    let notificationEmail = '';
    let businessName = 'Your Business Name';
    let autoNotify = 'FALSE';

    try {
      const settingsSheet = getSheet('settings');
      const settingsData = settingsSheet.getDataRange().getValues();
      const headers = settingsData[0];
      const keyColIndex = headers.indexOf('Key');
      const valueColIndex = headers.indexOf('Value');

      if (keyColIndex !== -1 && valueColIndex !== -1) {
        for (let i = 1; i < settingsData.length; i++) {
          const key = settingsData[i][keyColIndex];
          const value = settingsData[i][valueColIndex];
          if (key === 'NOTIFICATION_EMAIL') notificationEmail = value;
          if (key === 'BUSINESS_NAME') businessName = value;
          if (key === 'AUTO_NOTIFY') autoNotify = value;
        }
      }
    } catch (settingError) {
      Logger.log('Could not load settings for notifications: ' + settingError.toString());
    }

    if (autoNotify.toUpperCase() === 'TRUE' && notificationEmail && (dueTodayCount > 0 || overdueCount > 0)) {
      let emailBody = `Hello Team,\n\nHere is your daily Lead Follow-up summary from ${businessName}:\n\n`;
      emailBody += `Leads Due Today: ${dueTodayCount}\n`;
      emailBody += `Leads Overdue: ${overdueCount}\n\n`;
      emailBody += `Details:\n${leadsRequiringAttention.join('\n')}\n\n`;
      emailBody += `Please log in to the Lead Management System to take action.\n\n`;
      emailBody += `Best regards,\nYour Automated System`;

      MailApp.sendEmail(notificationEmail, `Daily Lead Follow-up Reminder - ${businessName}`, emailBody);
      Logger.log(`Email notification sent to ${notificationEmail}`);
    } else if (autoNotify.toUpperCase() === 'TRUE' && !notificationEmail) {
      Logger.log('AUTO_NOTIFY is TRUE but NOTIFICATION_EMAIL is not set in Settings.');
    } else if (autoNotify.toUpperCase() !== 'TRUE') {
      Logger.log('Email notifications are disabled (AUTO_NOTIFY is not TRUE).');
    }
    // --- End Optional Email ---

    Logger.log(`Follow-up check complete: Due Today: ${dueTodayCount}, Overdue: ${overdueCount}`);
    return {
      success: true,
      dueToday: dueTodayCount,
      overdue: overdueCount,
      total: dueTodayCount + overdueCount
    };

  } catch (e) {
    Logger.log('Error in checkFollowUps: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// DEBUG: Test function to check lead data
function debugGetLeads() {
  try {
    Logger.log('=== DEBUG GETLEADS ===');

    const sheet = getSheet('leads');
    Logger.log('Sheet name: ' + sheet.getName());
    Logger.log('Last row: ' + sheet.getLastRow());
    Logger.log('Last column: ' + sheet.getLastColumn());

    if (sheet.getLastRow() > 1) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log('Headers: ' + JSON.stringify(headers));

      const firstDataRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log('First data row: ' + JSON.stringify(firstDataRow));
    }

    const leads = getLeads();
    Logger.log('Number of leads returned: ' + leads.length);
    if (leads.length > 0) {
      Logger.log('First lead: ' + JSON.stringify(leads[0]));
    }

    return {
      success: true,
      sheetInfo: {
        name: sheet.getName(),
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn()
      },
      leadCount: leads.length,
      sampleLead: leads.length > 0 ? leads[0] : null
    };

  } catch (e) {
    Logger.log('Debug error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// Quick phone debug function
function debugPhone(leadId) {
  google.script.run.withSuccessHandler(function(result) {
    console.log('Phone debug:', result);
    if (result.success) {
      alert(`Phone Debug for ${leadId}:\n\n` +
            `Value: ${result.phoneValue}\n` +
            `Type: ${result.phoneType}\n` +
            `As String: ${result.phoneString}\n` +
            `Is Number: ${result.isNumber}\n` +
            `Is String: ${result.isString}`);
    } else {
      alert('Debug failed: ' + result.error);
    }
  }).debugLeadPhone(leadId);
}

function recordQuotation(leadId) {
  try {
    if (!leadId) {
      return { success: false, error: 'Lead ID is required.' };
    }

    // Add to follow-up history
    addFollowUpHistory(leadId, new Date(), 'Quotation Created', 'Completed', 'Quotation created via system', getCurrentUser());

    Logger.log(`Quotation recorded for lead: ${leadId}`);
    return { success: true, message: 'Quotation recorded successfully' };

  } catch (e) {
    Logger.log(`Error recording quotation for LeadID '${leadId}': ${e.toString()}`);
    return { success: false, error: e.toString() };
  }
}

// DEBUG: Check deal values in the system
function debugDealValues() {
  try {
    const leads = getLeads();
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();

    let debugInfo = {
      totalLeads: leads.length,
      leadsWithDealValues: 0,
      monthlyDealValues: [],
      allDealValues: []
    };

    leads.forEach(lead => {
      const dealValue = parseFloat(lead.DealValue) || 0;
      const lastContact = lead.LastContact ? new Date(lead.LastContact) : new Date();
      const status = lead.Status || 'New';

      if (dealValue > 0) {
        debugInfo.leadsWithDealValues++;
        debugInfo.allDealValues.push({
          leadId: lead.LeadID,
          name: lead.Name,
          status: status,
          dealValue: dealValue,
          lastContact: lastContact,
          isCurrentMonth: (lastContact.getMonth() === currentMonth && lastContact.getFullYear() === currentYear)
        });

        if (lastContact.getMonth() === currentMonth && lastContact.getFullYear() === currentYear) {
          debugInfo.monthlyDealValues.push({
            leadId: lead.LeadID,
            name: lead.Name,
            status: status,
            dealValue: dealValue
          });
        }
      }
    });

    Logger.log('=== DEAL VALUE DEBUG ===');
    Logger.log(`Total leads: ${debugInfo.totalLeads}`);
    Logger.log(`Leads with deal values: ${debugInfo.leadsWithDealValues}`);
    Logger.log(`Monthly deal values count: ${debugInfo.monthlyDealValues.length}`);
    Logger.log('Monthly deal values:', debugInfo.monthlyDealValues);
    Logger.log('All deal values:', debugInfo.allDealValues);

    return debugInfo;

  } catch (e) {
    Logger.log('Error in debugDealValues: ' + e.toString());
    return { error: e.toString() };
  }
}


// --- HISTORICAL REPORTING AND BACKFILL LOGIC ---

/**
 * Main function to be run on a daily trigger or manually from the menu.
 * It now backfills data for any missed days since the last successful run.
 */
function dailyUpdate() {
  const scriptProperties = PropertiesService.getScriptProperties();
  try {
    Logger.log('Starting daily update with backfill logic...');

    const lastRunDateStr = scriptProperties.getProperty('lastSuccessfulRunDate');
    let startDate = new Date();
    if (lastRunDateStr) {
      startDate = new Date(lastRunDateStr);
      startDate.setDate(startDate.getDate() + 1); // Start from the day after the last run.
    } else {
      // On the very first run, start from when the first lead was created.
      const firstLeadDate = getFirstLeadDate();
      startDate = firstLeadDate ? new Date(firstLeadDate) : new Date(new Date().setDate(new Date().getDate() - 30)); // Fallback
    }
    startDate.setHours(0, 0, 0, 0);

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    if (startDate >= today) {
        Logger.log('Reports are already up-to-date.');
        SpreadsheetApp.getUi().alert('Reports are already up-to-date!');
        return;
    }

    let currentDate = new Date(startDate.getTime());
    let lastProcessedDate = null;

    Logger.log(`Backfilling reports from ${currentDate.toLocaleDateString()} until yesterday.`);
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<p>Processing historical data... This may take a few minutes.</p>').setTitle('Update in Progress'));

    // Pre-fetch all data once to avoid multiple calls inside the loop
    const allLeads = getLeads();
    const allFollowUps = getSheet('followups').getDataRange().getValues(); // Includes headers

    while (currentDate < today) {
      const dateToProcess = new Date(currentDate.getTime());
      Logger.log(`Processing reports for: ${dateToProcess.toLocaleDateString()}`);

      // Pass prefetched data to the historical functions
      updateMonthlyLeadSummaryForDate(dateToProcess, allLeads, allFollowUps);
      updateDailyDashboardStatsForDate(dateToProcess, allLeads, allFollowUps);

      lastProcessedDate = dateToProcess;

      currentDate.setDate(currentDate.getDate() + 1);
    }

    if (lastProcessedDate) {
        scriptProperties.setProperty('lastSuccessfulRunDate', lastProcessedDate.toISOString());
        Logger.log(`Successfully processed until ${lastProcessedDate.toLocaleDateString()}. Updated lastSuccessfulRunDate.`);
    }

    Logger.log('Daily update backfill completed successfully.');
    SpreadsheetApp.getUi().alert('Daily reports have been successfully updated!');

  } catch (e) {
    Logger.log(`FATAL: Daily update failed: ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`An error occurred during the daily update: ${e.message}`);
  }
}

function getInitialData() {
  try {
    const teamMembers = getTeamMembers();
    const statusOptions = getStatusOptions();
    const sourceOptions = getSourceOptions();
    const showroomOptions = getShowroomOptions();

    return {
      teamMembers: teamMembers,
      statusOptions: statusOptions,
      sourceOptions: sourceOptions,
      showroomOptions: showroomOptions
    };

  } catch (e) {
    Logger.log('Error in getInitialData: ' + e.toString());
    return {
      teamMembers: [],
      statusOptions: [],
      sourceOptions: [],
      showroomOptions: []
    };
  }
}

/**
 * Helper to get the date of the very first lead to avoid unnecessary back-filling.
 */
function getFirstLeadDate() {
    const leadsSheet = getSheet('leads');
    if (leadsSheet.getLastRow() < 2) {
        return null;
    }
    // Assuming 'CreatedAt' is column G (index 7)
    const firstDate = leadsSheet.getRange(2, 7).getValue();
    return firstDate ? new Date(firstDate) : null;
}

/**
 * Generates and logs the MonthlyLeadSummary for a specific historical date.
 * This function is idempotent and reconstructs the state of leads for the given date.
 */
function updateMonthlyLeadSummaryForDate(dateToProcess, allLeadsData, allFollowUpsData) {
  const summarySheet = getSheet('monthlyLeadSummary');
  const endOfDateToProcess = new Date(dateToProcess);
  endOfDateToProcess.setHours(23, 59, 59, 999);

  // --- Make function idempotent: Delete today's existing records ---
  const data = summarySheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const recordDate = new Date(data[i][0]);
    recordDate.setHours(0, 0, 0, 0);
    if (recordDate.getTime() === dateToProcess.getTime()) {
      summarySheet.deleteRow(i + 1);
    }
  }

  // --- Historical Reconstruction Logic ---
  const activeStatuses = ['Prospecting', 'Convincing', 'Negotiation', 'Won'];
  let rowsToAdd = [];

  for (const lead of allLeadsData) {
    const createdAt = new Date(lead.CreatedAt);
    if (createdAt > endOfDateToProcess) continue; // Skip leads created after the date we are processing.

    // Determine the lead's state at the end of `dateToProcess`
    let stateAtDate = {
      status: lead.Status,
      dealValue: lead.DealValue,
      lastContact: new Date(lead.LastContact),
      notes: lead.Notes
    };

    const relevantFollowUps = allFollowUpsData
      .slice(1) // Skip header row
      .filter(fu => fu[0] === lead.LeadID && new Date(fu[1]) <= endOfDateToProcess)
      .sort((a, b) => new Date(a[1]) - new Date(b[1]));

    let lastStatus = 'New';
    let lastNote = lead.Notes;

    for (const fu of relevantFollowUps) {
      if (fu[3]) lastStatus = fu[3]; // Status is in column 4
      if (fu[4]) lastNote = fu[4]; // Notes is in column 5
    }

    // The status from the main 'Leads' sheet is the most current one.
    // We need to find the last status *before or on* the dateToProcess.
    const lastRelevantFollowUp = relevantFollowUps[relevantFollowUps.length - 1];
    if(lastRelevantFollowUp) {
        stateAtDate.status = lastRelevantFollowUp[3] || stateAtDate.status;
        stateAtDate.notes = lastRelevantFollowUp[4] || stateAtDate.notes;
    }

    if (activeStatuses.includes(stateAtDate.status)) {
      rowsToAdd.push([
        dateToProcess,
        lead.LeadID,
        lead.AssignedTo,
        lead.Name,
        lead.Source,
        lead.Phone,
        stateAtDate.notes,
        stateAtDate.status,
        lead.DealValue, // Assume DealValue is constant, could be improved
        stateAtDate.lastContact
      ]);
    }
  }

  if (rowsToAdd.length > 0) {
    summarySheet.getRange(summarySheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    Logger.log(`Logged ${rowsToAdd.length} leads to MonthlyLeadSummary for ${dateToProcess.toLocaleDateString()}.`);
  }
}


/**
 * Generates a log of daily events (leads created, status changes) and updates the DailyDashboardStats sheet.
 */
function updateDailyDashboardStatsForDate(dateToProcess, allLeadsData, allFollowUpsData) {
    const statsSheet = getSheet('dailyDashboardStats');
    const dayStart = new Date(dateToProcess);
    dayStart.setHours(0, 0, 0, 0);
    const dayEnd = new Date(dateToProcess);
    dayEnd.setHours(23, 59, 59, 999);

    // --- Make function idempotent: Delete today's existing records ---
    const data = statsSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
        const recordDate = new Date(data[i][0]);
        recordDate.setHours(0, 0, 0, 0);
        if (recordDate.getTime() === dateToProcess.getTime()) {
            statsSheet.deleteRow(i + 1);
        }
    }

    // --- Daily Event-Based Logic ---
    const dailyEvents = {
        LeadsCreated_Count: 0, LeadsCreated_Value: 0,
        BecameInProgress_Count: 0, BecameInProgress_Value: 0,
        BecameProspecting_Count: 0, BecameProspecting_Value: 0,
        BecameConvincing_Count: 0, BecameConvincing_Value: 0,
        BecameNegotiation_Count: 0, BecameNegotiation_Value: 0,
        BecameWon_Count: 0, BecameWon_Value: 0,
        BecameLost_Count: 0, BecameLost_Value: 0,
        BecameDrop_Count: 0, BecameDrop_Value: 0,
        LeadsContacted_Count: 0
    };

    const leadsMap = allLeadsData.reduce((map, lead) => {
        map[lead.LeadID] = lead;
        return map;
    }, {});

    // 1. Count leads created on this day
    const leadsCreatedOnDate = allLeadsData.filter(lead => {
        const createdAt = new Date(lead.CreatedAt);
        return createdAt >= dayStart && createdAt <= dayEnd;
    });

    dailyEvents.LeadsCreated_Count = leadsCreatedOnDate.length;
    dailyEvents.LeadsCreated_Value = leadsCreatedOnDate.reduce((sum, lead) => sum + (parseFloat(lead.DealValue) || 0), 0);

    // 2. Count status changes on this day
    const followUpsOnDate = allFollowUpsData
        .slice(1)
        .filter(fu => {
            const fuDate = new Date(fu[1]);
            return fuDate >= dayStart && fuDate <= dayEnd && fu[3]; // Ensure it has a status change
        });

    const contactedLeads = new Set();
    for (const fu of followUpsOnDate) {
        const leadId = fu[0];
        contactedLeads.add(leadId); // Add to set for unique count

        const newStatus = fu[3];
        const lead = leadsMap[leadId];
        if (!lead) continue;

        const dealValue = parseFloat(lead.DealValue) || 0;
        const statusKey = newStatus.replace(/\s+/g, ''); // e.g., "In Progress" -> "InProgress"

        const countKey = `Became${statusKey}_Count`;
        const valueKey = `Became${statusKey}_Value`;

        if (dailyEvents.hasOwnProperty(countKey)) {
            dailyEvents[countKey]++;
            dailyEvents[valueKey] += dealValue;
        }
    }
    dailyEvents.LeadsContacted_Count = contactedLeads.size;

    const newRow = [
        dateToProcess,
        dailyEvents.LeadsCreated_Count, dailyEvents.LeadsCreated_Value,
        dailyEvents.BecameInProgress_Count, dailyEvents.BecameInProgress_Value,
        dailyEvents.BecameProspecting_Count, dailyEvents.BecameProspecting_Value,
        dailyEvents.BecameConvincing_Count, dailyEvents.BecameConvincing_Value,
        dailyEvents.BecameNegotiation_Count, dailyEvents.BecameNegotiation_Value,
        dailyEvents.BecameWon_Count, dailyEvents.BecameWon_Value,
        dailyEvents.BecameLost_Count, dailyEvents.BecameLost_Value,
        dailyEvents.BecameDrop_Count, dailyEvents.BecameDrop_Value,
        dailyEvents.LeadsContacted_Count
    ];

    statsSheet.appendRow(newRow);
    Logger.log(`Logged daily events to DailyDashboardStats for ${dateToProcess.toLocaleDateString()}.`);
}

/**
 * Generates a comprehensive, event-based summary of team performance
 * and updates the TeamPerformanceDashboard sheet.
 */
function updateTeamPerformanceDashboard() {
    const ui = SpreadsheetApp.getUi();
    ui.showSidebar(HtmlService.createHtmlOutput('<p>Generating Team Performance Dashboard... This may take a moment.</p>').setTitle('Update in Progress'));

    try {
        const allLeads = getLeads();
        if (allLeads.length === 0) {
            ui.alert('No leads found to generate a report.');
            return;
        }

        const allFollowUps = getSheet('followups').getDataRange().getValues().slice(1);
        const performanceData = {};

        // Helper to initialize a month-member object
        const initPerfData = (key, month, member) => {
            if (!performanceData[key]) {
                performanceData[key] = {
                    month: month, teamMember: member,
                    NewLeads_Count: 0, NewLeads_Value: 0, InProgress_Count: 0, InProgress_Value: 0,
                    Prospecting_Count: 0, Prospecting_Value: 0, Convincing_Count: 0, Convincing_Value: 0,
                    Negotiation_Count: 0, Negotiation_Value: 0, Won_Count: 0, Won_Value: 0,
                    Lost_Count: 0, Lost_Value: 0, Drop_Count: 0, Drop_Value: 0
                };
            }
        };

        // 1. Process Lead Creation Events
        for (const lead of allLeads) {
            const createdAt = new Date(lead.CreatedAt);
            const monthStr = `${createdAt.getFullYear()}-${String(createdAt.getMonth() + 1).padStart(2, '0')}`;
            const assignee = lead.AssignedTo || 'Unassigned';
            const key = `${monthStr}-${assignee}`;

            initPerfData(key, monthStr, assignee);

            performanceData[key].NewLeads_Count++;
            performanceData[key].NewLeads_Value += parseFloat(lead.DealValue) || 0;
        }

        // 2. Process Status Change Events from Follow-ups
        const leadsMap = allLeads.reduce((map, lead) => {
            map[lead.LeadID] = lead;
            return map;
        }, {});

        for (const fu of allFollowUps) {
            const leadId = fu[0];
            const status = fu[3];
            if (!status) continue;

            const lead = leadsMap[leadId];
            if (!lead) continue;

            const eventDate = new Date(fu[1]);
            const monthStr = `${eventDate.getFullYear()}-${String(eventDate.getMonth() + 1).padStart(2, '0')}`;
            const assignee = lead.AssignedTo || 'Unassigned';
            const key = `${monthStr}-${assignee}`;

            initPerfData(key, monthStr, assignee);

            const dealValue = parseFloat(lead.DealValue) || 0;
            const statusKey = status.replace(/\s+/g, '');
            const countKey = `${statusKey}_Count`;
            const valueKey = `${statusKey}_Value`;

            if (performanceData[key].hasOwnProperty(countKey)) {
                performanceData[key][countKey]++;
                performanceData[key][valueKey] += dealValue;
            }
        }

        const sheet = getSheet('teamPerformanceDashboard');
        if (sheet.getLastRow() > 1) {
            sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
        }

        const rows = Object.values(performanceData).map(d => [
            d.month, d.teamMember,
            d.NewLeads_Count, d.NewLeads_Value, d.InProgress_Count, d.InProgress_Value,
            d.Prospecting_Count, d.Prospecting_Value, d.Convincing_Count, d.Convincing_Value,
            d.Negotiation_Count, d.Negotiation_Value, d.Won_Count, d.Won_Value,
            d.Lost_Count, d.Lost_Value, d.Drop_Count, d.Drop_Value
        ]);

        if (rows.length > 0) {
            rows.sort((a, b) => b[0].localeCompare(a[0]) || a[1].localeCompare(b[1]));
            sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
        }

        ui.alert('Team Performance Dashboard has been successfully updated!');
        Logger.log('Successfully updated the Team Performance Dashboard.');

    } catch (e) {
        Logger.log(`Error updating Team Performance Dashboard: ${e.toString()}`);
        ui.alert(`An error occurred while updating the dashboard: ${e.message}`);
    }
}

/**
 * Creates a custom menu in the Google Sheet UI when the sheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Admin Tools')
      .addItem('Run Daily Update', 'dailyUpdate')
      .addSeparator()
      .addItem('Force Full Data Regeneration', 'forceResetAndUpdate')
      .addSeparator()
      .addItem('Update Team Dashboard', 'updateTeamPerformanceDashboard')
      .addSeparator()
      .addItem('Sync All Lead Notes', 'syncAllLeadNotes')
      .addToUi();
}

/**
 * Deletes the last run date property and re-runs the full update.
 * This allows for a complete data regeneration if sheets are manually cleared.
 */
function forceResetAndUpdate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirm Action',
    'Are you sure you want to force a full regeneration of all historical report data? This may take several minutes and will replace all existing data in the summary sheets.',
    ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    try {
      Logger.log('User confirmed. Forcing full data regeneration...');
      PropertiesService.getScriptProperties().deleteProperty('lastSuccessfulRunDate');
      Logger.log('Deleted lastSuccessfulRunDate property.');

      // Call the main update function, which will now run from the beginning.
      dailyUpdate();
    } catch (e) {
      Logger.log(`Error during forced regeneration: ${e.toString()}`);
      ui.alert(`An error occurred during the regeneration: ${e.message}`);
    }
  } else {
    Logger.log('User cancelled the force regeneration.');
  }
}

/**
 * NEW UTILITY FUNCTION
 * Scans the entire FollowUpHistory to find the last recorded note for each lead
 * and updates the 'Notes' column in the 'Leads' sheet.
 */
function syncAllLeadNotes() {
  const ui = SpreadsheetApp.getUi();
  ui.showSidebar(HtmlService.createHtmlOutput('<p>Syncing all lead notes from history... This may take a moment.</p>').setTitle('Sync in Progress'));

  try {
    Logger.log('Starting syncAllLeadNotes...');
    const leadsSheet = getSheet('leads');
    const followupsSheet = getSheet('followups');

    const leadsData = leadsSheet.getDataRange().getValues();
    const followupsData = followupsSheet.getDataRange().getValues();

    const leadHeaders = leadsData[0];
    const leadIdCol = leadHeaders.indexOf('LeadID');
    const notesCol = leadHeaders.indexOf('Notes');

    if (leadIdCol === -1 || notesCol === -1) {
      throw new Error('Could not find LeadID or Notes column in Leads sheet.');
    }

    // Create a map of the latest note for each lead ID
    const latestNotes = {};
    const followupLeadIdCol = 0; // Assuming LeadID is the first column
    const followupNotesCol = 4;   // Assuming Notes is the fifth column

    for (let i = 1; i < followupsData.length; i++) {
      const row = followupsData[i];
      const leadId = row[followupLeadIdCol];
      const note = row[followupNotesCol];
      if (leadId && note) { // Only store if there is a note
        latestNotes[leadId] = note;
      }
    }

    let updatedCount = 0;
    // Update the leads sheet with the latest notes
    for (let i = 1; i < leadsData.length; i++) {
      const leadId = leadsData[i][leadIdCol];
      if (latestNotes[leadId] && leadsData[i][notesCol] !== latestNotes[leadId]) {
        leadsSheet.getRange(i + 1, notesCol + 1).setValue(latestNotes[leadId]);
        updatedCount++;
      }
    }

    Logger.log(`Sync complete. Updated ${updatedCount} lead notes.`);
    ui.alert(`Sync Complete!`, `Successfully updated the notes for ${updatedCount} leads based on their follow-up history.`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(`Error during syncAllLeadNotes: ${e.toString()}`);
    ui.alert(`An error occurred: ${e.message}`);
  }
}