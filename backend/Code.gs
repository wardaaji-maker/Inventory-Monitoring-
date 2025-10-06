const CONFIG = {
  sheetId: '1Wvn_BAvlYHsEiYN5D_0D1BrDwa1239vjuBk8LFJ_dKA', // Your actual spreadsheet ID
  sheets: {
    leads: 'Leads',
    followups: 'FollowUpHistory',
    transactions: 'Transaction',
    settings: 'Settings',
    templates: 'MessageTemplates',
    monthlyLeadSummary: 'MonthlyLeadSummary',
    dailyDashboardStats: 'DailyDashboardStats'
  }
};

function getTeamMembers() {
  return ['AJI', 'HERU', 'FAHMY', 'HILLARY', 'DWI PUJI', 'RAHMAT', 'RIZKY', 'MELLINDA', 'CISCO'];
}

// New function to get status options
function getStatusOptions() {
  return ['New', 'In Progress', 'Prospecting', 'Convincing', 'Negotiating', 'Converted', 'Lost'];
}

// New function to get source options
function getSourceOptions() {
  return ['Website', 'Social Media', 'Referral', 'Advertisement', 'Walk In', 'Other'];
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
      ['L001', 'John Doe', '+1234567890', 'john@test.com', 'Website', 'Product A', new Date(), 'New', new Date(Date.now() + 3 * 24 * 60 * 60 * 1000), 0, 'Interested in pricing', 'Team', new Date()],
      ['L002', 'Jane Smith', '+1234567891', 'jane@test.com', 'Social Media', 'Product B', new Date(), 'In Progress', new Date(Date.now() + 1 * 24 * 60 * 60 * 1000), 1, 'Follow up tomorrow', 'Team', new Date()],
      ['L003', 'Mike Johnson', '+1234567892', 'mike@test.com', 'Referral', 'Product C', new Date(), 'New', new Date(Date.now() + 5 * 24 * 60 * 60 * 1000), 0, 'Hot lead', 'Team', new Date()]
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
          headers = ['LeadID', 'Name', 'Phone', 'Email', 'Source', 'Product', 'CreatedAt', 'Status', 'NextFollowUp', 'FollowUpCount', 'Notes', 'AssignedTo', 'LastContact', 'DealValue', 'ReceiptNumber'];
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
            headers = ['Date', 'TotalLeads', 'New', 'InProgress', 'Prospecting', 'Convincing', 'Negotiating', 'Converted', 'Lost', 'DueToday', 'Overdue', 'ProspectingValue', 'ConvincingValue', 'NegotiatingValue', 'ConvertedValue', 'TotalValue', 'LeadsContactedToday'];
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
      total: leads.length,
      new: 0,
      inProgress: 0,
      prospecting: 0,
      convincing: 0,
      negotiating: 0,
      converted: 0,
      lost: 0,
      dueToday: 0,
      overdue: 0,
      monthlyStats: {
        prospectingValue: 0,
        convincingValue: 0,
        negotiatingValue: 0,
        convertedValue: 0,
        totalValue: 0
      }
    };

    today.setHours(0, 0, 0, 0);

    leads.forEach(lead => {
      const status = lead.Status || 'New';
      const dealValue = parseFloat(lead.DealValue) || 0;
      const lastContact = lead.LastContact ? new Date(lead.LastContact) : new Date();

      // Count statuses
      if (status === 'New') stats.new++;
      else if (status === 'In Progress') stats.inProgress++;
      else if (status === 'Prospecting') stats.prospecting++;
      else if (status === 'Convincing') stats.convincing++;
      else if (status === 'Negotiating') stats.negotiating++;
      else if (status === 'Converted') stats.converted++;
      else if (status === 'Lost') stats.lost++;

      // FIXED: Monthly accumulation based on current month (not creation date)
      // Count deal values for current month regardless of when lead was created
      if (lastContact.getMonth() === currentMonth && lastContact.getFullYear() === currentYear) {
        if (status === 'Prospecting') {
          stats.monthlyStats.prospectingValue += dealValue;
        } else if (status === 'Convincing') {
          stats.monthlyStats.convincingValue += dealValue;
        } else if (status === 'Negotiating') {
          stats.monthlyStats.negotiatingValue += dealValue;
        } else if (status === 'Converted') {
          stats.monthlyStats.convertedValue += dealValue;
        }

        // Add to total if it's one of the tracked statuses
        if (['Prospecting', 'Convincing', 'Negotiating', 'Converted'].includes(status)) {
          stats.monthlyStats.totalValue += dealValue;
        }
      }

      // Check follow-up dates
      if (status !== 'Converted' && status !== 'Lost' && lead.NextFollowUp) {
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
    Logger.log(`Negotiating: ${stats.negotiating} with value: ${stats.monthlyStats.negotiatingValue}`);
    Logger.log(`Converted: ${stats.converted} with value: ${stats.monthlyStats.convertedValue}`);
    Logger.log(`Total Monthly Value: ${stats.monthlyStats.totalValue}`);

    return stats;

  } catch (e) {
    Logger.log('Error in getDashboardStats: ' + e.toString());
    return {
      total: 0, new: 0, inProgress: 0, prospecting: 0, convincing: 0,
      negotiating: 0, converted: 0, lost: 0, dueToday: 0, overdue: 0,
      monthlyStats: {
        prospectingValue: 0,
        convincingValue: 0,
        negotiatingValue: 0,
        convertedValue: 0,
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
function updateStatus(leadId, newStatus, dealValue = '', receiptNumber = '') {
  try {
    if (!leadId || !newStatus) {
      return { success: false, error: 'Lead ID and new status are required.' };
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
      if (data[i][leadIDColIndex] === leadId) {
        // Update the row
        sheet.getRange(i + 1, statusColIndex + 1).setValue(newStatus);
        sheet.getRange(i + 1, lastContactColIndex + 1).setValue(new Date());

        // Update deal value if provided
        if (dealValue && dealValueColIndex !== -1) {
          sheet.getRange(i + 1, dealValueColIndex + 1).setValue(dealValue);
        }

        // Update receipt number if provided and status is Converted
        if (receiptNumber && receiptNumberColIndex !== -1 && newStatus === 'Converted') {
          sheet.getRange(i + 1, receiptNumberColIndex + 1).setValue(receiptNumber);
        }

        // If converted or lost, clear next follow-up
        if (newStatus === 'Converted' || newStatus === 'Lost') {
          sheet.getRange(i + 1, nextFollowUpColIndex + 1).setValue('');
        }

        leadFound = true;

        // Add history entry
        let historyNotes = `Lead status changed to '${newStatus}'`;
        if (dealValue) historyNotes += ` with deal value: ${dealValue}`;
        if (receiptNumber && newStatus === 'Converted') historyNotes += ` | Receipt: ${receiptNumber}`;

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
function updateLeadStatus(leadId, newStatus) {
  return updateStatus(leadId, newStatus);
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
      if (data[i][leadIDColIndex] === leadId) {
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
    addFollowUpHistory(leadId, new Date(), followUpData.type, 'Completed', followUpData.notes + ' - Message: ' + followUpData.message, getCurrentUser());

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
      if (data[i][leadIDColIndex] === leadId) {
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
      if (status !== 'Converted' && status !== 'Lost' && lead.NextFollowUp) {
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

// --- NEW DAILY UPDATE FUNCTIONS ---

/**
 * Main function to be run on a daily trigger.
 * To set up the trigger:
 * 1. In the Apps Script editor, go to Edit > Current project's triggers.
 * 2. Click "Add Trigger".
 * 3. Choose "dailyUpdate" as the function to run.
 * 4. Select "Time-driven" as the event source.
 * 5. Select "Day timer" and choose a time (e.g., 1am to 2am).
 * 6. Click "Save".
 */
function dailyUpdate() {
  try {
    Logger.log('Starting daily update...');
    updateMonthlyLeadSummary();
    updateDailyDashboardStats();
    Logger.log('Daily update completed successfully.');
  } catch (e) {
    Logger.log(`FATAL: Daily update failed: ${e.toString()}`);
  }
}

/**
 * Updates the MonthlyLeadSummary sheet with the latest data for active leads.
 */
function updateMonthlyLeadSummary() {
  try {
    Logger.log('Updating Monthly Lead Summary...');
    const leads = getLeads();
    const followUps = getSheet('followups').getDataRange().getValues().slice(1); // Get all follow-up data, skip headers

    const summarySheet = getSheet('monthlyLeadSummary');
    const today = new Date();

    // Create a map of the latest follow-up note for each lead
    const latestNotes = {};
    followUps.forEach(row => {
      const leadId = row[0];
      const note = row[4];
      latestNotes[leadId] = note; // This will overwrite older notes, leaving the last one
    });

    const activeStatuses = ['Prospecting', 'Convincing', 'Negotiating', 'Converted'];
    const leadsToLog = leads.filter(lead => activeStatuses.includes(lead.Status));

    if (leadsToLog.length === 0) {
      Logger.log('No active leads to log today.');
      return;
    }

    const rowsToAdd = leadsToLog.map(lead => {
      return [
        today,
        lead.LeadID,
        lead.AssignedTo,
        lead.Name,
        lead.Source,
        lead.Phone,
        latestNotes[lead.LeadID] || lead.Notes, // Use latest follow-up note or original note
        lead.Status,
        lead.DealValue,
        lead.LastContact ? new Date(lead.LastContact) : ''
      ];
    });

    summarySheet.getRange(summarySheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    Logger.log(`Logged ${rowsToAdd.length} leads to MonthlyLeadSummary.`);

  } catch (e) {
    Logger.log(`Error in updateMonthlyLeadSummary: ${e.toString()}`);
  }
}

/**
 * Updates the DailyDashboardStats sheet with the current day's stats.
 */
function updateDailyDashboardStats() {
  try {
    Logger.log('Updating Daily Dashboard Stats...');
    const stats = getDashboardStats();
    const today = new Date();

    // Get leads contacted today
    const followUpsSheet = getSheet('followups');
    const followUpData = followUpsSheet.getDataRange().getValues().slice(1);
    const contactedToday = new Set();
    const todayStart = new Date(today.setHours(0, 0, 0, 0));

    followUpData.forEach(row => {
        const followUpDate = new Date(row[1]);
        if (followUpDate >= todayStart) {
            contactedToday.add(row[0]); // Add lead ID to a set to count unique leads
        }
    });

    const newRow = [
      today,
      stats.total,
      stats.new,
      stats.inProgress,
      stats.prospecting,
      stats.convincing,
      stats.negotiating,
      stats.converted,
      stats.lost,
      stats.dueToday,
      stats.overdue,
      stats.monthlyStats.prospectingValue,
      stats.monthlyStats.convincingValue,
      stats.monthlyStats.negotiatingValue,
      stats.monthlyStats.convertedValue,
      stats.monthlyStats.totalValue,
      contactedToday.size // Number of unique leads contacted today
    ];

    const statsSheet = getSheet('dailyDashboardStats');
    statsSheet.appendRow(newRow);
    Logger.log(`Logged today's stats to DailyDashboardStats. Leads contacted: ${contactedToday.size}`);

  } catch (e) {
    Logger.log(`Error in updateDailyDashboardStats: ${e.toString()}`);
  }
}