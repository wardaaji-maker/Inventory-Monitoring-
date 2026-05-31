// ============================================
// DEBT MANAGER PRO - Smart Debt Management System
// Version: 2.0
// Compatible with Global Users (English/Indonesian)
// ============================================

var SHEET_ID = '1zw6V_uzHEcyLKgLpGgoxDqook0UP8Kp4QuiJyjf_SEg';
var sheet;

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Debt Manager Pro - Finance Tracker')
    .setWidth(1400)
    .setHeight(900)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function sanitizeData(data) {
  if (Array.isArray(data)) {
    return data.map(function(item) {
      return sanitizeData(item);
    });
  } else if (data instanceof Date) {
    return data.toISOString();
  } else if (data !== null && typeof data === 'object') {
    var sanitizedObj = {};
    for (var key in data) {
      sanitizedObj[key] = sanitizeData(data[key]);
    }
    return sanitizedObj;
  }
  return data;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Debt Manager Pro')
    .addItem('🚀 Open Web App', 'showWebApp')
    .addItem('🎯 Generate Repayment Strategy', 'sortDebtsByPriority')
    .addItem('📊 View Dashboard', 'showDashboard')
    .addToUi();
}

function showWebApp() {
  var html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Debt Manager Pro')
    .setWidth(1400)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, '💰 Debt Manager Pro');
}

function showDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('dashboard')
    .setTitle('Debt Dashboard')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Financial Dashboard');
}


function getOrCreateSheet() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    sheet = ss.getSheetByName('Debts_Data');

    // Create sheet if doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Debts_Data');

      // Create headers with proper formatting
      var headers = [
        ['ID', 'Timestamp', 'Debt Name', 'Debt Type', 'Principal Balance',
         'Interest Rate', 'Interest Type', 'Due Date', 'OJK Status',
         'Daily Interest %', 'Priority Score', 'Repayment Priority', 'Status', 'Note']
      ];

      sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

      // Format headers
      var headerRange = sheet.getRange(1, 1, 1, headers[0].length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4F46E5');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setHorizontalAlignment('center');

      // Set column widths
      sheet.setColumnWidths(1, 14, [120, 150, 180, 150, 150, 120, 120, 120, 120, 130, 130, 150, 120, 200]);

      // Freeze header row
      sheet.setFrozenRows(1);

      // Add data validation for Debt Type
      var debtTypes = ['Legal Fintech', 'Illegal Fintech', 'Buy Now Pay Later', 'Mortgage', 'Credit Card', 'Personal Loan', 'Student Loan'];
      var debtRule = SpreadsheetApp.newDataValidation().requireValueInList(debtTypes, true).build();
      sheet.getRange('D:D').setDataValidation(debtRule);

      // Add data validation for Interest Type
      var interestTypes = ['Per Day', 'Per Month', 'Per Year'];
      var interestRule = SpreadsheetApp.newDataValidation().requireValueInList(interestTypes, true).build();
      sheet.getRange('G:G').setDataValidation(interestRule);

      // Add data validation for OJK Status
      var ojkStatus = ['Registered', 'Not Registered', 'Pending'];
      var ojkRule = SpreadsheetApp.newDataValidation().requireValueInList(ojkStatus, true).build();
      sheet.getRange('I:I').setDataValidation(ojkRule);

      // Add data validation for Status
      var statuses = ['Active', 'Paid', 'Overdue', 'Negotiating'];
      var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).build();
      sheet.getRange('M:M').setDataValidation(statusRule);

      // Add conditional formatting for high interest (>0.1% per day)
      var range = sheet.getRange('J:J');
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0.1)
        .setBackground('#FFE5E5')
        .setFontColor('#FF0000')
        .setRanges([range])
        .build();
      sheet.setConditionalFormatRules([rule]);
    }

    var paymentsSheet = ss.getSheetByName('Payments_Data');
    if (!paymentsSheet) {
      paymentsSheet = ss.insertSheet('Payments_Data');
      var paymentHeaders = [
        ['Payment ID', 'Timestamp', 'Debt ID', 'Debt Name', 'Amount Paid', 'Proof URL']
      ];
      paymentsSheet.getRange(1, 1, 1, paymentHeaders[0].length).setValues(paymentHeaders);

      var pHeaderRange = paymentsSheet.getRange(1, 1, 1, paymentHeaders[0].length);
      pHeaderRange.setFontWeight('bold');
      pHeaderRange.setBackground('#10B981');
      pHeaderRange.setFontColor('#FFFFFF');
      pHeaderRange.setHorizontalAlignment('center');

      paymentsSheet.setColumnWidths(1, 6, [150, 150, 150, 180, 150, 250]);
      paymentsSheet.setFrozenRows(1);
    } else {
      // Ensure the existing sheet has the Proof URL column header
      var currentHeaders = paymentsSheet.getRange(1, 1, 1, paymentsSheet.getLastColumn() || 1).getValues()[0];
      if (currentHeaders.length < 6 || currentHeaders[5] !== 'Proof URL') {
        paymentsSheet.getRange(1, 6).setValue('Proof URL');
        paymentsSheet.getRange(1, 6).setFontWeight('bold')
                                    .setBackground('#10B981')
                                    .setFontColor('#FFFFFF')
                                    .setHorizontalAlignment('center');
        paymentsSheet.setColumnWidth(6, 250);
      }
    }

    var financeSheet = ss.getSheetByName('Finance_Data');
    if (!financeSheet) {
      financeSheet = ss.insertSheet('Finance_Data');
      var financeHeaders = [
        ['ID', 'Timestamp', 'Type', 'Category', 'Amount', 'Note']
      ];
      financeSheet.getRange(1, 1, 1, financeHeaders[0].length).setValues(financeHeaders);

      var fHeaderRange = financeSheet.getRange(1, 1, 1, financeHeaders[0].length);
      fHeaderRange.setFontWeight('bold');
      fHeaderRange.setBackground('#3B82F6');
      fHeaderRange.setFontColor('#FFFFFF');
      fHeaderRange.setHorizontalAlignment('center');

      financeSheet.setColumnWidths(1, 6, [150, 150, 120, 150, 150, 250]);
      financeSheet.setFrozenRows(1);
    }

    var budgetSheet = ss.getSheetByName('Budgets_Data');
    if (!budgetSheet) {
      budgetSheet = ss.insertSheet('Budgets_Data');
      var budgetHeaders = [
        ['ID', 'Timestamp', 'Budget Name', 'Budget Category', 'Amount']
      ];
      budgetSheet.getRange(1, 1, 1, budgetHeaders[0].length).setValues(budgetHeaders);

      var bHeaderRange = budgetSheet.getRange(1, 1, 1, budgetHeaders[0].length);
      bHeaderRange.setFontWeight('bold');
      bHeaderRange.setBackground('#F59E0B');
      bHeaderRange.setFontColor('#FFFFFF');
      bHeaderRange.setHorizontalAlignment('center');

      budgetSheet.setColumnWidths(1, 5, [150, 150, 200, 150, 150]);
      budgetSheet.setFrozenRows(1);
    }

    return sheet;
  } catch (error) {
    console.error('Error accessing sheet:', error);
    throw new Error('Unable to access spreadsheet. Please check permissions.');
  }
}

function saveDebt(data) {
  try {
    var sheet = getOrCreateSheet();
    var timestamp = new Date();
    var debtId = 'DEBT_' + timestamp.getTime();
    var dailyInterest = calculateDailyInterest(data.interestRate, data.interestType);

    // Check for illegal interest rate
    var warning = '';
    if (dailyInterest > 0.1 && (data.debtType === 'Legal Fintech' || data.debtType === 'Illegal Fintech')) {
      warning = '⚠️ INTEREST EXCEEDS OJK LIMIT (0.1%/day) - Potential Illegal Lender';
    }

    var row = [
      debtId,
      timestamp,
      data.debtName,
      data.debtType,
      parseFloat(data.principalBalance),
      parseFloat(data.interestRate),
      data.interestType,
      data.dueDate,
      data.ojkStatus,
      dailyInterest,
      0, // Priority score (to be calculated)
      '', // Repayment priority
      'Active',
      data.note || ''
    ];

    var lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);

    // Apply formatting for warning
    if (warning) {
      sheet.getRange(lastRow, 1, 1, 14).setBackground('#FFE5E5');
      sheet.getRange(lastRow, 3).setNote(warning);
    }

    return {
      success: true,
      message: '✅ Debt added successfully!',
      warning: warning,
      debtId: debtId
    };
  } catch (error) {
    return { success: false, message: '❌ Error: ' + error.toString() };
  }
}

function calculateDailyInterest(rate, type) {
  var rateValue = parseFloat(rate);
  if (type === 'Per Day') {
    return rateValue;
  } else if (type === 'Per Month') {
    return rateValue / 30;
  } else if (type === 'Per Year') {
    return rateValue / 365;
  }
  return 0;
}

function getAllDebts() {
  try {
    var sheet = getOrCreateSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    var debts = [];

    for (var i = 0; i < data.length; i++) {
      if (data[i][12] !== 'Paid') { // Only show active debts
        debts.push({
          id: data[i][0],
          row: i + 2,
          timestamp: data[i][1],
          debtName: data[i][2],
          debtType: data[i][3],
          principalBalance: data[i][4],
          interestRate: data[i][5],
          interestType: data[i][6],
          dueDate: data[i][7],
          ojkStatus: data[i][8],
          dailyInterest: data[i][9],
          priorityScore: data[i][10],
          repaymentPriority: data[i][11],
          status: data[i][12],
          note: data[i][13] || ''
        });
      }
    }

    return sanitizeData(debts);
  } catch (error) {
    console.error('Error in getAllDebts: ' + error.toString());
    return [];
  }
}

function sortDebtsByPriority() {
  try {
    var sheet = getOrCreateSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) return { success: false, message: 'No debts found. Please add some debts first.' };

    var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    var today = new Date();

    // Calculate priority score for active debts only
    for (var i = 0; i < data.length; i++) {
      if (data[i][12] !== 'Paid') {
        var dueDate = new Date(data[i][7]);
        var daysUntilDue = Math.ceil((dueDate - today) / (1000 * 60 * 60 * 24));
        var urgencyScore = daysUntilDue > 0 ? Math.max(0, 1 / daysUntilDue) : 100;

        // Avalanche Method: Higher weight on interest rate
        var priorityScore = (data[i][9] * 1000) + (urgencyScore * 10);
        data[i][10] = priorityScore;
      }
    }

    // Sort by priority score (descending) - active debts first
    data.sort(function(a, b) {
      if (a[12] === 'Paid' && b[12] !== 'Paid') return 1;
      if (a[12] !== 'Paid' && b[12] === 'Paid') return -1;
      return b[10] - a[10];
    });

    // Update priority numbers
    var priorityCounter = 1;
    for (var i = 0; i < data.length; i++) {
      if (data[i][12] !== 'Paid') {
        data[i][11] = 'Priority ' + priorityCounter++;
      } else {
        data[i][11] = 'Completed';
      }
    }

    // Write back sorted data
    sheet.getRange(2, 1, data.length, 14).setValues(data);

    // Apply visual formatting for top priority
    for (var i = 0; i < data.length; i++) {
      if (data[i][11] === 'Priority 1') {
        sheet.getRange(i + 2, 1, 1, 14).setBackground('#FFF3E0');
        sheet.getRange(i + 2, 1, 1, 14).setFontWeight('bold');
      } else if (data[i][9] > 0.1) {
        sheet.getRange(i + 2, 1, 1, 14).setBackground('#FFE5E5');
      } else if (data[i][11] === 'Completed') {
        sheet.getRange(i + 2, 1, 1, 14).setBackground('#E8F5E9');
      } else {
        sheet.getRange(i + 2, 1, 1, 14).setBackground('#FFFFFF');
      }
    }

    return {
      success: true,
      message: '🎯 Repayment strategy generated! Pay in order: Highest interest + Nearest due date first.'
    };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function getStatistics() {
  try {
    var sheet = getOrCreateSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return {
        totalDebt: 0,
        activeDebts: 0,
        highInterestDebts: 0,
        avgInterest: 0,
        urgentDebts: 0
      };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    var totalDebt = 0;
    var activeCount = 0;
    var highInterestCount = 0;
    var totalInterest = 0;
    var urgentCount = 0;
    var today = new Date();

    for (var i = 0; i < data.length; i++) {
      if (data[i][12] !== 'Paid') {
        totalDebt += data[i][4];
        activeCount++;
        totalInterest += data[i][9];

        if (data[i][9] > 0.1) highInterestCount++;

        var dueDate = new Date(data[i][7]);
        var daysUntilDue = Math.ceil((dueDate - today) / (1000 * 60 * 60 * 24));
        if (daysUntilDue <= 7 && daysUntilDue > 0) urgentCount++;
      }
    }

    return {
      totalDebt: totalDebt,
      activeDebts: activeCount,
      highInterestDebts: highInterestCount,
      avgInterest: activeCount > 0 ? totalInterest / activeCount : 0,
      urgentDebts: urgentCount
    };
  } catch (error) {
    return {
      totalDebt: 0,
      activeDebts: 0,
      highInterestDebts: 0,
      avgInterest: 0,
      urgentDebts: 0
    };
  }
}

function updateDebtStatus(debtId, newStatus) {
  try {
    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === debtId) {
        sheet.getRange(i + 1, 13).setValue(newStatus);
        return { success: true, message: 'Status updated successfully!' };
      }
    }
    return { success: false, message: 'Debt not found.' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function makePayment(debtId, amount, base64Proof, filename) {
  try {
    var sheet = getOrCreateSheet();
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var paymentsSheet = ss.getSheetByName('Payments_Data');

    var data = sheet.getDataRange().getValues();
    var paymentAmount = parseFloat(amount);

    if (isNaN(paymentAmount) || paymentAmount <= 0) {
      return { success: false, message: 'Invalid payment amount.' };
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === debtId) {
        var currentBalance = parseFloat(data[i][4]);
        var newBalance = currentBalance - paymentAmount;
        var debtName = data[i][2];

        if (newBalance < 0) newBalance = 0;

        sheet.getRange(i + 1, 5).setValue(newBalance);

        if (newBalance === 0) {
          sheet.getRange(i + 1, 13).setValue('Paid');
        }

        var proofUrl = '';
        if (base64Proof && filename) {
          try {
            var contentType = base64Proof.substring(5, base64Proof.indexOf(';'));
            var bytes = Utilities.base64Decode(base64Proof.substr(base64Proof.indexOf('base64,') + 7));
            var blob = Utilities.newBlob(bytes, contentType, filename);
            var folder = DriveApp.getRootFolder();
            var file = folder.createFile(blob);
            proofUrl = file.getUrl();
          } catch(e) {
            console.error('File upload error: ' + e);
          }
        }

        var timestamp = new Date();
        var paymentId = 'PAY_' + timestamp.getTime();
        var paymentRow = [
          paymentId,
          timestamp,
          debtId,
          debtName,
          paymentAmount,
          proofUrl
        ];

        paymentsSheet.appendRow(paymentRow);

        return {
          success: true,
          message: '✅ Payment of ' + paymentAmount + ' successfully recorded for ' + debtName + '!'
        };
      }
    }

    return { success: false, message: 'Debt not found.' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function getAllPayments() {
  try {
    getOrCreateSheet(); // Ensure sheets are created
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var paymentsSheet = ss.getSheetByName('Payments_Data');
    var lastRow = paymentsSheet.getLastRow();

    if (lastRow <= 1) return [];

    var data = paymentsSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var payments = [];

    for (var i = 0; i < data.length; i++) {
      payments.push({
        paymentId: data[i][0],
        timestamp: (data[i][1] instanceof Date) ? data[i][1].toISOString() : data[i][1],
        debtId: data[i][2],
        debtName: data[i][3],
        amount: data[i][4],
        proofUrl: data[i][5]
      });
    }

    // Sort by most recent first
    payments.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });

    return sanitizeData(payments);
  } catch (error) {
    console.error('Error in getAllPayments: ' + error.toString());
    return [];
  }
}

function updateDebtDetails(debtId, updatedData) {
  try {
    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === debtId) {
        var rate = parseFloat(updatedData.interestRate);
        var dailyInterest = calculateDailyInterest(rate, updatedData.interestType);

        sheet.getRange(i + 1, 3).setValue(updatedData.debtName);
        sheet.getRange(i + 1, 4).setValue(updatedData.debtType);
        sheet.getRange(i + 1, 5).setValue(parseFloat(updatedData.principalBalance));
        sheet.getRange(i + 1, 6).setValue(rate);
        sheet.getRange(i + 1, 7).setValue(updatedData.interestType);
        sheet.getRange(i + 1, 8).setValue(updatedData.dueDate);
        sheet.getRange(i + 1, 9).setValue(updatedData.ojkStatus);
        sheet.getRange(i + 1, 10).setValue(dailyInterest);
        sheet.getRange(i + 1, 14).setValue(updatedData.note || '');

        return { success: true, message: 'Debt updated successfully!' };
      }
    }
    return { success: false, message: 'Debt not found.' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function saveFinanceRecord(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var financeSheet = ss.getSheetByName('Finance_Data');
    if (!financeSheet) {
      getOrCreateSheet();
      financeSheet = ss.getSheetByName('Finance_Data');
    }

    var timestamp = new Date();
    var recordId = 'FIN_' + timestamp.getTime();

    var row = [
      recordId,
      timestamp,
      data.type,
      data.category,
      parseFloat(data.amount),
      data.note || ''
    ];

    financeSheet.appendRow(row);

    return { success: true, message: 'Finance record saved successfully!' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function getFinanceRecords() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var financeSheet = ss.getSheetByName('Finance_Data');
    if (!financeSheet) return [];

    var lastRow = financeSheet.getLastRow();
    if (lastRow <= 1) return [];

    var data = financeSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var records = [];

    for (var i = 0; i < data.length; i++) {
      records.push({
        id: data[i][0],
        timestamp: (data[i][1] instanceof Date) ? data[i][1].toISOString() : data[i][1],
        type: data[i][2],
        category: data[i][3],
        amount: data[i][4],
        note: data[i][5]
      });
    }

    records.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });

    return sanitizeData(records);
  } catch (error) {
    console.error('Error in getFinanceRecords: ' + error.toString());
    return [];
  }
}

function saveBudget(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var budgetSheet = ss.getSheetByName('Budgets_Data');
    if (!budgetSheet) {
      getOrCreateSheet();
      budgetSheet = ss.getSheetByName('Budgets_Data');
    }

    var timestamp = new Date();
    var recordId = 'BUDGET_' + timestamp.getTime();

    var row = [
      recordId,
      timestamp,
      data.budgetName,
      data.budgetCategory,
      parseFloat(data.amount)
    ];

    budgetSheet.appendRow(row);
    return { success: true, message: 'Budget allocated successfully!' };
  } catch(error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function getBudgets() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var budgetSheet = ss.getSheetByName('Budgets_Data');
    if (!budgetSheet) return [];

    var lastRow = budgetSheet.getLastRow();
    if (lastRow <= 1) return [];

    var data = budgetSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var budgets = [];

    // Get spent amounts
    var financeRecords = getFinanceRecords();
    var spentMap = {};
    financeRecords.forEach(function(r) {
      if (r.type === 'Outcome') {
        spentMap[r.category] = (spentMap[r.category] || 0) + parseFloat(r.amount);
      }
    });

    for (var i = 0; i < data.length; i++) {
      var name = data[i][2];
      var allocated = parseFloat(data[i][4]);
      var spent = spentMap[name] || 0;
      budgets.push({
        id: data[i][0],
        timestamp: data[i][1],
        budgetName: name,
        budgetCategory: data[i][3],
        allocated: allocated,
        spent: spent,
        remaining: allocated - spent
      });
    }

    return sanitizeData(budgets);
  } catch (error) {
    console.error('Error in getBudgets: ' + error.toString());
    return [];
  }
}

function getFinanceSummary() {
  try {
    var records = getFinanceRecords();
    var budgets = getBudgets();
    var totalIncome = 0;
    var totalOutcome = 0;

    records.forEach(function(record) {
      if (record.type === 'Income') {
        totalIncome += parseFloat(record.amount);
      } else if (record.type === 'Outcome') {
        totalOutcome += parseFloat(record.amount);
      }
    });

    var budgetSummary = {
      'Fixed Budget': { allocated: 0, spent: 0 },
      'Variable Budget': { allocated: 0, spent: 0 },
      'Periodic Budget': { allocated: 0, spent: 0 },
      'Saving Budget': { allocated: 0, spent: 0 }
    };

    budgets.forEach(function(b) {
      if (budgetSummary[b.budgetCategory]) {
        budgetSummary[b.budgetCategory].allocated += b.allocated;
        budgetSummary[b.budgetCategory].spent += b.spent;
      }
    });

    var netBalance = totalIncome - totalOutcome;
    var dailyCapability = netBalance > 0 ? (netBalance / 30) : 0;

    var result = {
      totalIncome: totalIncome,
      totalOutcome: totalOutcome,
      netBalance: netBalance,
      monthlyCapability: netBalance > 0 ? netBalance : 0,
      dailyCapability: dailyCapability,
      budgetSummary: budgetSummary
    };

    return sanitizeData(result);
  } catch (error) {
    console.error('Error in getFinanceSummary: ' + error.toString());
    return {
      totalIncome: 0,
      totalOutcome: 0,
      netBalance: 0,
      monthlyCapability: 0,
      dailyCapability: 0,
      budgetSummary: {}
    };
  }
}