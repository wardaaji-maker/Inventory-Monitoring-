// The ID of your Google Sheet
const SHEET_ID = '1mmeGtFaIfMKavcRFU5CIYG-SeKxuaWaLiVj4IySyZq0';

// The names of your sheets
const SALES_SHEET_NAME = 'Master Data';
const TARGET_SHEET_NAME = 'Target';
const TEAM_SHEET_NAME = 'Team Member';
const ACTION_PLANS_SHEET_NAME = 'Action Plans';

// --- Main Function to Serve the Web App ---
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sales Performance Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- Analytics Functions ---
function getDashboardData(timeFilter) {
  const salesData = getFilteredSales(timeFilter);

  // Overall Performance
  const totalSales = salesData.reduce((sum, sale) => sum + (sale['AFTER TAX'] || 0), 0);
  const uniqueSOs = [...new Set(salesData.map(sale => sale['SALES ORDER NUMBER']))];
  const soCount = uniqueSOs.length;
  const qtySold = salesData.reduce((sum, sale) => sum + (sale['QTY'] || 0), 0);

  const overall = {
    totalSales: totalSales,
    soCount: soCount,
    qtySold: qtySold,
    ats: soCount > 0 ? totalSales / soCount : 0,
    basketSize: soCount > 0 ? qtySold / soCount : 0,
    asp: qtySold > 0 ? totalSales / qtySold : 0
  };

  // Category Performance
  const byCategory1 = calculateGroupPerformance(salesData, 'CATEGORY 1');
  const byCategory2 = calculateGroupPerformance(salesData, 'CATEGORY 2');
  const bySubCategory = calculateGroupPerformance(salesData, 'SUB CATEGORY');

  // Salesman Performance
  const salesmanPerformance = calculateGroupPerformance(salesData, 'SALESMAN');
  const salesmanDetails = calculateSalesmanPerformanceDetails(salesData);

  // Best Sellers
  const bestSellersFurniture = getBestSellersByCategory(salesData, 'F');
  const bestSellersAccessories = getBestSellersByCategory(salesData, 'A');

  // Performance by Customer Type and Promotion
  const byCustomerType = calculateGroupPerformance(salesData, 'CUSTOMER TYPE');
  const byPromotion = calculateGroupPerformance(salesData, 'PROMOTION');

  // Performance Trends (simplified)
  const trends = getPerformanceTrends(salesData);

  // --- Sales Target Calculation ---
  const { totalTarget, dailyTargets } = getTargetData(timeFilter);
  overall.salesTarget = totalTarget;
  overall.achievement = totalTarget > 0 ? (overall.totalSales / totalTarget) * 100 : 0;
  trends.dailyTargets = dailyTargets;
  // --- End Sales Target Calculation ---

  return {
    overall: overall,
    byCategory1: byCategory1,
    byCategory2: byCategory2,
    bySubCategory: bySubCategory,
    salesmanPerformance: salesmanPerformance,
    salesmanPerformanceDetails: salesmanDetails,
    bestSellersFurniture: bestSellersFurniture,
    bestSellersAccessories: bestSellersAccessories,
    byCustomerType: byCustomerType,
    byPromotion: byPromotion,
    trends: trends,
    summary: {
      totalTransactions: salesData.length,
      averageSale: totalSales / (salesData.length || 1),
      topCategory: Object.keys(byCategory1)[0] || 'N/A',
      topSalesman: Object.keys(salesmanPerformance)[0] || 'N/A'
    }
  };
}

function calculateGroupPerformance(salesData, groupKey) {
  const performance = {};

  salesData.forEach(sale => {
    const groupValue = sale[groupKey];
    if (!groupValue) return;

    if (!performance[groupValue]) {
      performance[groupValue] = {
        totalSales: 0,
        soCount: new Set(),
        qtySold: 0,
        transactions: 0
      };
    }

    performance[groupValue].totalSales += sale['AFTER TAX'] || 0;
    performance[groupValue].soCount.add(sale['SALES ORDER NUMBER']);
    performance[groupValue].qtySold += sale['QTY'] || 0;
    performance[groupValue].transactions += 1;
  });

  // Calculate derived metrics
  Object.keys(performance).forEach(key => {
    const perf = performance[key];
    perf.soCount = perf.soCount.size;
    perf.ats = perf.soCount > 0 ? perf.totalSales / perf.soCount : 0;
    perf.basketSize = perf.soCount > 0 ? perf.qtySold / perf.soCount : 0;
    perf.asp = perf.qtySold > 0 ? perf.totalSales / perf.qtySold : 0;
  });

  return performance;
}

function calculateSalesmanPerformanceDetails(salesData) {
  const details = {};

  salesData.forEach(sale => {
    const salesman = sale['SALESMAN'];
    const category = sale['CATEGORY 1'];
    const customerType = sale['CUSTOMER TYPE'];

    if (!salesman) return;

    if (!details[salesman]) {
      details[salesman] = {
        byCategory: {},
        byCustomerType: {}
      };
    }

    // Process by category
    if (category) {
      if (!details[salesman].byCategory[category]) {
        details[salesman].byCategory[category] = { totalSales: 0, soCount: new Set(), qtySold: 0 };
      }
      details[salesman].byCategory[category].totalSales += sale['AFTER TAX'] || 0;
      details[salesman].byCategory[category].soCount.add(sale['SALES ORDER NUMBER']);
      details[salesman].byCategory[category].qtySold += sale['QTY'] || 0;
    }

    // Process by customer type
    if (customerType) {
      if (!details[salesman].byCustomerType[customerType]) {
        details[salesman].byCustomerType[customerType] = { totalSales: 0, soCount: new Set(), qtySold: 0 };
      }
      details[salesman].byCustomerType[customerType].totalSales += sale['AFTER TAX'] || 0;
      details[salesman].byCustomerType[customerType].soCount.add(sale['SALES ORDER NUMBER']);
      details[salesman].byCustomerType[customerType].qtySold += sale['QTY'] || 0;
    }
  });

  // Finalize soCount and calculate derived metrics
  for (const salesman in details) {
    for (const category in details[salesman].byCategory) {
      const perf = details[salesman].byCategory[category];
      perf.soCount = perf.soCount.size;
      perf.ats = perf.soCount > 0 ? perf.totalSales / perf.soCount : 0;
      perf.basketSize = perf.soCount > 0 ? perf.qtySold / perf.soCount : 0;
    }
    for (const customerType in details[salesman].byCustomerType) {
      const perf = details[salesman].byCustomerType[customerType];
      perf.soCount = perf.soCount.size;
      perf.ats = perf.soCount > 0 ? perf.totalSales / perf.soCount : 0;
      perf.basketSize = perf.soCount > 0 ? perf.qtySold / perf.soCount : 0;
    }
  }

  return details;
}

function getBestSellersByCategory(salesData, categoryPrefix, limit = 10) {
  const products = {};

  salesData.forEach(sale => {
    const product = sale['DESCRIPTION'];
    const category = sale['CATEGORY 1'];
    if (!product || !category || !category.startsWith(categoryPrefix)) return;

    if (!products[product]) {
      products[product] = {
        totalSales: 0,
        qtySold: 0,
        unitPrice: sale['UNIT PRICE'] || 0
      };
    }
    products[product].totalSales += sale['AFTER TAX'] || 0;
    products[product].qtySold += sale['QTY'] || 0;
  });

  return Object.entries(products)
    .sort((a, b) => b[1].totalSales - a[1].totalSales)
    .slice(0, limit)
    .map(([product, data]) => ({ product, ...data }));
}

function getBestSellers(salesData, limit = 10) {
  const products = {};

  salesData.forEach(sale => {
    const product = sale['DESCRIPTION'];
    if (!product) return;

    if (!products[product]) {
      products[product] = {
        totalSales: 0,
        qtySold: 0,
        unitPrice: sale['UNIT PRICE'] || 0
      };
    }
    products[product].totalSales += sale['AFTER TAX'] || 0;
    products[product].qtySold += sale['QTY'] || 0;
  });

  return Object.entries(products)
    .sort((a, b) => b[1].totalSales - a[1].totalSales)
    .slice(0, limit)
    .map(([product, data]) => ({ product, ...data }));
}

function getPerformanceTrends(salesData) {
  // Simplified trend calculation - you can enhance this with actual time-series data
  const dailySales = {};

  salesData.forEach(sale => {
    const date = new Date(sale['DATE']).toDateString();
    if (!dailySales[date]) {
      dailySales[date] = 0;
    }
    dailySales[date] += sale['AFTER TAX'] || 0;
  });

  return {
    dailySales: Object.entries(dailySales)
      .sort(([a], [b]) => new Date(a) - new Date(b))
      .slice(-7) // Last 7 days
  };
}

function getFilteredSales(timeFilter) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SALES_SHEET_NAME);
  let allData = sheet.getDataRange().getValues();
  const headers = allData.shift();

  // --- Product Exclusion ---
  const excludedProductIDs = [
    'F000000028311', // SPUNBON BAG VIVERE S
    'F000000026401', // SPUNBON BAG VIVERE M
    'F000000026402'  // SPUNBON BAG VIVERE L
  ];
  const productIDColumnIndex = headers.indexOf('PRODUCT ID');
  if (productIDColumnIndex !== -1) {
    allData = allData.filter(row => !excludedProductIDs.includes(row[productIDColumnIndex]));
  }
  // --- End Product Exclusion ---

  const dateColumnIndex = headers.indexOf('DATE');

  if (dateColumnIndex === -1) {
    throw new Error('"DATE" column not found in Master Data sheet.');
  }

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  let startDate, endDate;

  switch(timeFilter) {
    case 'daily':
      // Adjust for "Today" filter to show yesterday's data as per user request
      startDate = new Date(today);
      startDate.setDate(startDate.getDate() - 1); // Yesterday
      endDate = new Date(today); // Today (so the range is all of yesterday)
      break;
    case 'weekly':
      startDate = new Date(today);
      const dayOfWeek = startDate.getDay();
      const diffToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      startDate.setDate(startDate.getDate() + diffToMonday);
      startDate.setHours(0, 0, 0, 0);

      endDate = new Date(startDate);
      endDate.setDate(endDate.getDate() + 7);
      break;
    case 'monthly':
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
      endDate = new Date(today.getFullYear(), today.getMonth() + 1, 1);
      break;
    case 'yearly':
      startDate = new Date(today.getFullYear(), 0, 1);
      endDate = new Date(today.getFullYear() + 1, 0, 1);
      break;
    default: // 'all'
      return allData.map(row => {
        const sale = {};
        headers.forEach((header, i) => {
          sale[header] = row[i];
        });
        return sale;
      });
  }

  const filteredRows = allData.filter(row => {
    const rowDate = new Date(row[dateColumnIndex]);
    if (isNaN(rowDate.getTime())) return false;

    const normalizedRowDate = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
    return normalizedRowDate >= startDate && normalizedRowDate < endDate;
  });

  return filteredRows.map(row => {
    const sale = {};
    headers.forEach((header, i) => {
      sale[header] = row[i];
    });
    return sale;
  });
}

// Export function for data export
function exportData(timeFilter) {
  const data = getDashboardData(timeFilter);
  return JSON.stringify(data);
}

function getTargetData(timeFilter) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TARGET_SHEET_NAME);
  const allData = sheet.getDataRange().getValues();
  const headers = allData.shift();
  const dateColumnIndex = headers.indexOf('DATE');
  const targetColumnIndex = headers.indexOf('TARGET');

  if (dateColumnIndex === -1 || targetColumnIndex === -1) {
    // Return zeroed data if columns are not found, to avoid breaking the dashboard
    return { totalTarget: 0, dailyTargets: [] };
  }

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  let startDate, endDate;

  // This date logic should EXACTLY match getFilteredSales to ensure alignment
  switch(timeFilter) {
    case 'daily':
      startDate = new Date(today);
      startDate.setDate(startDate.getDate() - 1);
      endDate = new Date(today);
      break;
    case 'weekly':
      startDate = new Date(today);
      const dayOfWeek = startDate.getDay();
      const diffToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      startDate.setDate(startDate.getDate() + diffToMonday);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(startDate);
      endDate.setDate(endDate.getDate() + 7);
      break;
    case 'monthly':
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
      endDate = new Date(today.getFullYear(), today.getMonth() + 1, 1);
      break;
    case 'yearly':
      startDate = new Date(today.getFullYear(), 0, 1);
      endDate = new Date(today.getFullYear() + 1, 0, 1);
      break;
    default: // 'all'
      startDate = new Date(0);
      endDate = new Date(8640000000000000);
  }

  const dailyTargets = {};
  let totalTarget = 0;

  allData.forEach(row => {
    const rowDate = new Date(row[dateColumnIndex]);
    if (isNaN(rowDate.getTime())) return;

    const normalizedRowDate = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());

    if (normalizedRowDate >= startDate && normalizedRowDate < endDate) {
      const targetValue = parseFloat(row[targetColumnIndex]) || 0;
      totalTarget += targetValue;

      const dateString = normalizedRowDate.toDateString();
      if (!dailyTargets[dateString]) {
        dailyTargets[dateString] = 0;
      }
      dailyTargets[dateString] += targetValue;
    }
  });

  const sortedDailyTargets = Object.entries(dailyTargets)
    .sort(([a], [b]) => new Date(a) - new Date(b));

  return {
    totalTarget: totalTarget,
    dailyTargets: sortedDailyTargets
  };
}
