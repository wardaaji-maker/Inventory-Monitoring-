// ============================================
// DEBT MANAGER PRO - Smart Debt Management System
// Version: 2.0
// Compatible with Global Users (English/Indonesian)
// ============================================

var SHEET_ID = '1zw6V_uzHEcyLKgLpGgoxDqook0UP8Kp4QuiJyjf_SEg';
var sheet;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Debt Manager Pro - Smart Debt Repayment System')
    .setWidth(1400)
    .setHeight(900)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
         'Daily Interest %', 'Priority Score', 'Repayment Priority', 'Status']
      ];

      sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

      // Format headers
      var headerRange = sheet.getRange(1, 1, 1, headers[0].length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4F46E5');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setHorizontalAlignment('center');

      // Set column widths
      sheet.setColumnWidths(1, 13, [120, 150, 180, 150, 150, 120, 120, 120, 120, 130, 130, 150, 120]);

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
        ['Payment ID', 'Timestamp', 'Debt ID', 'Debt Name', 'Amount Paid']
      ];
      paymentsSheet.getRange(1, 1, 1, paymentHeaders[0].length).setValues(paymentHeaders);

      var pHeaderRange = paymentsSheet.getRange(1, 1, 1, paymentHeaders[0].length);
      pHeaderRange.setFontWeight('bold');
      pHeaderRange.setBackground('#10B981');
      pHeaderRange.setFontColor('#FFFFFF');
      pHeaderRange.setHorizontalAlignment('center');

      paymentsSheet.setColumnWidths(1, 5, [150, 150, 150, 180, 150]);
      paymentsSheet.setFrozenRows(1);
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
      'Active'
    ];

    var lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);

    // Apply formatting for warning
    if (warning) {
      sheet.getRange(lastRow, 1, 1, 13).setBackground('#FFE5E5');
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

    var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    var debts = [];

    for (var i = 0; i < data.length; i++) {
      if (data[i][12] !== 'Paid') { // Only show active debts
        debts.push({
          id: data[i][0],
          row: i + 2,
          timestamp: (data[i][1] instanceof Date) ? data[i][1].toISOString() : data[i][1],
          debtName: data[i][2],
          debtType: data[i][3],
          principalBalance: data[i][4],
          interestRate: data[i][5],
          interestType: data[i][6],
          dueDate: (data[i][7] instanceof Date) ? data[i][7].toISOString() : data[i][7],
          ojkStatus: data[i][8],
          dailyInterest: data[i][9],
          priorityScore: data[i][10],
          repaymentPriority: data[i][11],
          status: data[i][12]
        });
      }
    }

    return debts;
  } catch (error) {
    return [];
  }
}

function sortDebtsByPriority() {
  try {
    var sheet = getOrCreateSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) return { success: false, message: 'No debts found. Please add some debts first.' };

    var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
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
    sheet.getRange(2, 1, data.length, 13).setValues(data);

    // Apply visual formatting for top priority
    for (var i = 0; i < data.length; i++) {
      if (data[i][11] === 'Priority 1') {
        sheet.getRange(i + 2, 1, 1, 13).setBackground('#FFF3E0');
        sheet.getRange(i + 2, 1, 1, 13).setFontWeight('bold');
      } else if (data[i][9] > 0.1) {
        sheet.getRange(i + 2, 1, 1, 13).setBackground('#FFE5E5');
      } else if (data[i][11] === 'Completed') {
        sheet.getRange(i + 2, 1, 1, 13).setBackground('#E8F5E9');
      } else {
        sheet.getRange(i + 2, 1, 1, 13).setBackground('#FFFFFF');
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

    var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
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

function makePayment(debtId, amount) {
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

        var timestamp = new Date();
        var paymentId = 'PAY_' + timestamp.getTime();
        var paymentRow = [
          paymentId,
          timestamp,
          debtId,
          debtName,
          paymentAmount
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

    var data = paymentsSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var payments = [];

    for (var i = 0; i < data.length; i++) {
      payments.push({
        paymentId: data[i][0],
        timestamp: (data[i][1] instanceof Date) ? data[i][1].toISOString() : data[i][1],
        debtId: data[i][2],
        debtName: data[i][3],
        amount: data[i][4]
      });
    }

    // Sort by most recent first
    payments.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });

    return payments;
  } catch (error) {
    return [];
  }
}