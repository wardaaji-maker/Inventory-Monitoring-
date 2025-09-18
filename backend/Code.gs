// CONFIGURATION
const CONFIG = {
  SHEETS: {
    APP_SHEET: "Inventory Monitoring System South78",
    INVENTORY: "Inventory",
    COUNTS: "PhysicalCounts"
  },
  COLUMNS: {
    APP_SHEET: {
      ID: 0,        // Column A: ItemID
      NAME: 1,      // Column B: ItemName
      STOCK: 10     // Column K: CurrentStock
    },
    INVENTORY_HEADER: ["ItemID", "ItemName", "CurrentStock"]
  },
  MAX_RETRIES: 3
};

// SIMPLIFIED doGet - JUST SERVE HTML, NO LOGIC
function doGet() {
  try {
    console.log("doGet called - serving HTML");
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('📦 Physical Count System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error("doGet error:", error);
    return HtmlService.createHtmlOutput(`
      <html>
        <body>
          <h2>❌ Application Error</h2>
          <p>Error: ${error.message}</p>
          <p>Please check the script and try again.</p>
        </body>
      </html>
    `);
  }
}

// SIMPLIFIED doPost
function doPost(e) {
  console.log("doPost called with parameters:", e ? e.parameters : 'No parameters');

  if (!e || !e.postData) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: "No request data received."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  let data;
  try {
    data = JSON.parse(e.postData.contents);
    console.log("Parsed data:", data);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: "Invalid JSON input."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const action = data?.action;
  console.log("Action:", action);

  try {
    switch (action) {
      case "getSystemStatus":
        return ContentService.createTextOutput(JSON.stringify(getSystemStatus()))
          .setMimeType(ContentService.MimeType.JSON);

      case "getAllItems":
        return ContentService.createTextOutput(JSON.stringify(getAllItems()))
          .setMimeType(ContentService.MimeType.JSON);

      case "getItemDetails":
        return ContentService.createTextOutput(JSON.stringify(getItemDetails(data.itemId)))
          .setMimeType(ContentService.MimeType.JSON);

      case "syncInventory":
        return ContentService.createTextOutput(JSON.stringify(syncInventory()))
          .setMimeType(ContentService.MimeType.JSON);

      case "initializeSheets":
        return ContentService.createTextOutput(JSON.stringify(initializeSheets()))
          .setMimeType(ContentService.MimeType.JSON);

      case "submitCount":
        return ContentService.createTextOutput(JSON.stringify(submitCount(data.payload)))
          .setMimeType(ContentService.MimeType.JSON);

      case "uploadPhoto":
        return ContentService.createTextOutput(JSON.stringify(uploadPhoto(data.payload.image, data.payload.fileName, data.payload.mimeType)))
          .setMimeType(ContentService.MimeType.JSON);

      case "getRecentSubmissions":
        return ContentService.createTextOutput(JSON.stringify(getRecentSubmissions(data.limit)))
          .setMimeType(ContentService.MimeType.JSON);

      default:
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: `Unknown action: ${action}`
        })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    console.error(`Action ${action} failed:`, error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// SYSTEM STATUS - SIMPLIFIED
function getSystemStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, ready: false, message: "Spreadsheet not found" };

    // Check if main sheet exists
    const mainSheet = ss.getSheetByName(CONFIG.SHEETS.APP_SHEET);
    if (!mainSheet) {
      return {
        success: false,
        ready: false,
        message: `Main sheet '${CONFIG.SHEETS.APP_SHEET}' not found`
      };
    }

    return {
      success: true,
      ready: true,
      message: "System ready",
      mainSheetExists: true
    };

  } catch (error) {
    return { success: false, ready: false, message: error.message };
  }
}

// GET ALL ITEMS - SIMPLIFIED
function getAllItems() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.APP_SHEET);

    if (!sheet) {
      return { success: false, error: "Sheet not found", items: [] };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, items: [] };
    }

    // Read data from columns A, B, K
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const items = data.map(row => ({
      id: String(row[CONFIG.COLUMNS.APP_SHEET.ID] || '').trim(),
      name: String(row[CONFIG.COLUMNS.APP_SHEET.NAME] || '').trim(),
      stock: Number(row[CONFIG.COLUMNS.APP_SHEET.STOCK]) || 0
    })).filter(item => item.id && item.id !== '');

    console.log(`Found ${items.length} items`);
    return { success: true, items: items };

  } catch (error) {
    console.error("Error in getAllItems:", error);
    return { success: false, error: error.message, items: [] };
  }
}

// GET ITEM DETAILS
function getItemDetails(itemId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.APP_SHEET);

    if (!sheet) {
      return { success: false, error: "Sheet not found" };
    }

    const data = sheet.getDataRange().getValues();

    // Skip header row, search for item
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const currentId = String(row[CONFIG.COLUMNS.APP_SHEET.ID] || '').trim();

      if (currentId.toLowerCase() === String(itemId).toLowerCase().trim()) {
        return {
          success: true,
          id: currentId,
          name: String(row[CONFIG.COLUMNS.APP_SHEET.NAME] || '').trim(),
          stock: Number(row[CONFIG.COLUMNS.APP_SHEET.STOCK]) || 0
        };
      }
    }

    return { success: false, error: "Item not found" };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

// INITIALIZE SHEETS
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Check if main sheet exists
    const mainSheet = ss.getSheetByName(CONFIG.SHEETS.APP_SHEET);
    if (!mainSheet) {
      return { success: false, error: "Main sheet not found" };
    }

    // Create Inventory sheet if needed
    let inventorySheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    if (!inventorySheet) {
      inventorySheet = ss.insertSheet(CONFIG.SHEETS.INVENTORY);
      inventorySheet.getRange('A1:C1').setValues([CONFIG.COLUMNS.INVENTORY_HEADER]);
    }

    // Create Counts sheet if needed
    let countsSheet = ss.getSheetByName(CONFIG.SHEETS.COUNTS);
    if (!countsSheet) {
      countsSheet = ss.insertSheet(CONFIG.SHEETS.COUNTS);
      countsSheet.getRange('A1:J1').setValues([[
        'CountID', 'Timestamp', 'User', 'ItemID', 'ItemName',
        'CurrentStock', 'PhysicalQty', 'Variance', 'PhotoURL', 'Notes'
      ]]);
    }

    return {
      success: true,
      message: "Sheets initialized successfully"
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

// SYNC INVENTORY
function syncInventory() {
  try {
    const result = getAllItems();
    if (!result.success) {
      return result;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);

    if (!targetSheet) {
      return { success: false, error: "Inventory sheet not found" };
    }

    // Clear and update target sheet
    targetSheet.clear();
    targetSheet.getRange('A1:C1').setValues([CONFIG.COLUMNS.INVENTORY_HEADER]);

    if (result.items.length > 0) {
      const data = result.items.map(item => [item.id, item.name, item.stock]);
      targetSheet.getRange(2, 1, data.length, 3).setValues(data);
    }

    return {
      success: true,
      syncedItems: result.items.length,
      message: `Synced ${result.items.length} items successfully`
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

// MENU FUNCTIONS
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📦 Inventory Tools')
      .addItem('Launch Count System', 'launchPhysicalCountSystem')
      .addSeparator()
      .addItem('Sync Inventory', 'menuSyncInventory')
      .addItem('Initialize System', 'menuInitializeSheets')
      .addToUi();
  } catch (error) {
    console.error("Error creating menu:", error);
  }
}

function launchPhysicalCountSystem() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(850)
      .setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, '📦 Physical Count System');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error launching system: ' + error.message);
  }
}

function menuSyncInventory() {
  try {
    const result = syncInventory();
    const ui = SpreadsheetApp.getUi();
    ui.alert(result.success ? '✅ Success' : '❌ Error',
             result.success ? result.message : result.error,
             ui.ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Sync error: ' + error.message);
  }
}

function menuInitializeSheets() {
  try {
    const result = initializeSheets();
    const ui = SpreadsheetApp.getUi();
    ui.alert(result.success ? '✅ Success' : '❌ Error',
             result.success ? result.message : result.error,
             ui.ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Initialization error: ' + error.message);
  }
}

function submitCount(formData) {
  try {
    console.log("Submitting count with data:", formData);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const countsSheet = ss.getSheetByName(CONFIG.SHEETS.COUNTS);

    if (!countsSheet) {
      return { success: false, error: "Counts sheet not found. Please initialize system first." };
    }

    // Calculate variance
    const currentStock = Number(formData.currentStock) || 0;
    const physicalQty = Number(formData.physicalQty) || 0;
    const variance = physicalQty - currentStock;

    // Prepare the record
    const record = [
      `COUNT-${Utilities.getUuid().slice(0, 8)}`, // CountID
      new Date(),                                  // Timestamp
      Session.getActiveUser().getEmail(),          // User
      String(formData.itemId || '').trim(),        // ItemID
      String(formData.itemName || '').trim(),      // ItemName
      currentStock,                                // CurrentStock
      physicalQty,                                 // PhysicalQty
      variance,                                    // Variance
      String(formData.photoUrl || ''),             // PhotoURL
      String(formData.notes || '').trim()          // Notes
    ];

    // Append to sheet
    countsSheet.appendRow(record);

    console.log("Count submitted successfully:", record);

    return {
      success: true,
      countId: record[0],
      variance: variance,
      message: "Count submitted successfully!"
    };

  } catch (error) {
    console.error("Submit count error:", error);
    return { success: false, error: error.message };
  }
}
// ==================== UPLOAD PHOTO FUNCTION ====================

function uploadPhoto(base64Image, fileName, mimeType) {
  try {
    console.log("Starting photo upload...");

    if (!base64Image || !fileName) {
      return { success: false, error: "Invalid image data" };
    }

    // Remove data URL prefix if present
    const base64Data = base64Image.includes(',') ? base64Image.split(',')[1] : base64Image;
    const byteString = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(byteString, mimeType, fileName);

    // Create or get the folder
    let folder;
    try {
      folder = DriveApp.getFoldersByName('Inventory Count Photos').next();
    } catch (e) {
      folder = DriveApp.createFolder('Inventory Count Photos');
      console.log("Created new folder: Inventory Count Photos");
    }

    // Create the file
    const photo = folder.createFile(blob);
    photo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    console.log("Photo uploaded successfully:", photo.getUrl());

    return {
      success: true,
      url: photo.getUrl(),
      id: photo.getId(),
      message: "Photo uploaded successfully"
    };

  } catch (error) {
    console.error("Photo upload error:", error);
    return { success: false, error: error.message };
  }
}


// ==================== FIXED GET RECENT SUBMISSIONS ====================

function getRecentSubmissions(limit = 20) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.COUNTS);

    if (!sheet) {
      console.log("Counts sheet not found");
      return { success: true, counts: [], message: "Counts sheet not found" };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log("No data in counts sheet");
      return { success: true, counts: [], message: "No counts recorded yet" };
    }

    // --- OPTIMIZATION ---
    // Instead of reading the whole sheet, read only the last ~50 rows.
    // This is much faster if the sheet has many entries.
    const rowsToRead = 50;
    const startRow = Math.max(2, lastRow - rowsToRead + 1);
    const numRows = lastRow - startRow + 1;

    const data = sheet.getRange(startRow, 1, numRows, 10).getValues();

    // Process this smaller chunk of data
    const counts = data
      .map(row => {
        let timestamp = row[1];
        if (typeof timestamp === 'string') {
          timestamp = new Date(timestamp);
        } else if (!(timestamp instanceof Date)) {
          timestamp = new Date(); // Fallback
        }

        return {
          countId: row[0] || '',
          timestamp: timestamp.toISOString(), // Convert date to string for transfer
          user: row[2] || '',
          itemId: row[3] || '',
          itemName: row[4] || '',
          currentStock: Number(row[5]) || 0,
          physicalQty: Number(row[6]) || 0,
          variance: Number(row[7]) || 0,
          photoUrl: row[8] || '',
          notes: row[9] || ''
        };
      })
      .filter(item => item.itemId && item.itemId !== '') // Filter out empty rows
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp)) // Sort by timestamp descending
      .slice(0, limit); // Get the most recent `limit`

    console.log(`Processed ${counts.length} recent counts`);
    return { success: true, counts: counts };

  } catch (error) {
    console.error("Error in getRecentSubmissions:", error);
    return { success: false, error: error.message, counts: [] };
  }
}

function debugCountsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.COUNTS);

    if (!sheet) {
      return { success: false, error: "Counts sheet not found" };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];
    const rowCount = data.length;

    // Get sample data (first 5 rows if available)
    const sampleData = rowCount > 1 ? data.slice(1, Math.min(6, rowCount)) : [];

    return {
      success: true,
      sheetExists: true,
      rowCount: rowCount,
      headers: headers,
      sampleData: sampleData,
      hasData: rowCount > 1,
      sheetName: CONFIG.SHEETS.COUNTS
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}
