const fs = require('fs');
let code = fs.readFileSync('backend/Code.gs', 'utf-8');

const regexHeaders = /(function setupSheetHeaders\(sheetName, sheet\) \{\s*try \{\s*)(\/\/ Clear existing data if needed \(only if empty\)\s*if \(sheet\.getLastRow\(\) === 0\) \{)([\s\S]*?)(\}\s*\}\s*catch \(e\) \{)/;

const newHeadersInner = `let headers = [];

    switch(sheetName) {
      case 'Leads':
        headers = ['LeadID', 'Name', 'Phone', 'Domicile', 'LeadType', 'Source', 'Product', 'CreatedAt', 'Status', 'NextFollowUp', 'FollowUpCount', 'Notes', 'AssignedTo', 'LastContact', 'DealValue', 'ReceiptNumber'];
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
      case 'Visitors':
        headers = ['Date', 'Customer Name', 'Domicile', 'Customer Type', 'Source', 'Interest', 'PICSales', 'Notes'];
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
            'BecameNegotiating_Count', 'BecameNegotiating_Value',
            'BecameWon_Count', 'BecameWon_Value',
            'BecameLost_Count', 'BecameLost_Value',
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
            'Negotiating_Count', 'Negotiating_Value',
            'Won_Count', 'Won_Value',
            'Lost_Count', 'Lost_Value'
          ];
          break;
      default:
        return; // No headers for unknown sheets
    }

    if (headers.length > 0) {
      // Check if we need to insert columns to fit the headers
      const maxColumns = sheet.getMaxColumns();
      if (headers.length > maxColumns) {
          sheet.insertColumnsAfter(maxColumns, headers.length - maxColumns);
      }

      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');
      headerRange.setHorizontalAlignment('center');
      headerRange.setBorder(true, true, true, true, true, true);

      sheet.setFrozenRows(1);

      if (sheet.getLastRow() === 0) {
        for (let i = 1; i <= headers.length; i++) {
          sheet.autoResizeColumn(i);
        }
      }

      Logger.log(\`Headers set up for \${sheetName}\`);
    }
  `;

code = code.replace(regexHeaders, `$1$2\n      ${newHeadersInner}\n    $4`);

const regexOnOpen = /(function onOpen\(\) \{\s*SpreadsheetApp\.getUi\(\)\s*\.createMenu\('Admin Tools'\)[\s\S]*?)(\.addToUi\(\);\s*\})/;

code = code.replace(regexOnOpen, `$1.addSeparator()\n      .addItem('Fix and Format All Headers', 'setupSheets')\n      $2`);

fs.writeFileSync('backend/Code.gs', code);
