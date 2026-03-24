// ============================================
// CLOUD KITCHEN INVENTORY MANAGEMENT SYSTEM
// ============================================

const SS = SpreadsheetApp.getActiveSpreadsheet();
const INVENTORY_SHEET = 'Inventory';
const STOCK_IN_SHEET = 'Stock In';
const STOCK_OUT_SHEET = 'Stock Out';
const SETTINGS_SHEET = 'Settings';
const DASHBOARD_SHEET = 'Dashboard';
const BATCHES_SHEET = 'Batches';

// ============================================
// MENU SETUP
// ============================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🍳 Inventory Manager')
    .addItem('➕ Add Stock (Receive)', 'showStockInDialog')
    .addItem('➖ Remove Stock (Use)', 'showStockOutDialog')
    .addItem('📦 Add New Item', 'showNewItemDialog')
    .addSeparator()
    .addItem('🔄 Update All Stock Levels', 'recalculateAllStock')
    .addItem('📊 Refresh Dashboard', 'updateDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Reports')
      .addItem('Low Stock Report', 'generateLowStockReport')
      .addItem('Expiring Items Report', 'generateExpiryReport')
      .addItem('Stock Valuation Report', 'generateValuationReport')
      .addItem('Usage Report (Last 7 Days)', 'generateUsageReport'))
    .addSeparator()
    .addItem('📧 Send Alert Email', 'sendAlertEmail')
    .addItem('⚙️ Settings', 'showSettingsDialog')
    .addToUi();
}

// ============================================
// DIALOG FUNCTIONS
// ============================================

function showStockInDialog() {
  const html = HtmlService.createHtmlOutput(getStockInHtml())
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Receive Stock');
}

function showStockOutDialog() {
  const html = HtmlService.createHtmlOutput(getStockOutHtml())
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Stock Usage');
}

function showNewItemDialog() {
  const html = HtmlService.createHtmlOutput(getNewItemHtml())
    .setWidth(550)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Inventory Item');
}

function showSettingsDialog() {
  const html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setWidth(450)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

// ============================================
// CONSTANTS
// ============================================

const COLUMN_NONPERISHABLE = 13;  // Column N
const COLUMN_STATUS = 14;          // Column O
const COLUMN_UPDATED = 15;         // Column P

// ============================================
// CORE INVENTORY FUNCTIONS
// ============================================

function normalizeBoolean(value) {
  if (typeof value === 'boolean') return value;
  if (value === 'on' || value === 'TRUE' || value === 'true') return true;
  return false;
}

function getInventoryItems() {
  const sheet = SS.getSheetByName(INVENTORY_SHEET);
  const data = sheet.getDataRange().getValues();
  const items = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      items.push({
        id: data[i][0],
        name: data[i][1],
        category: data[i][2],
        unit: data[i][3],
        currentStock: data[i][4],
        reorderLevel: data[i][5],
        nonPerishable: data[i][COLUMN_NONPERISHABLE] === true
      });
    }
  }
  return items;
}

function getSuppliers() {
  const sheet = SS.getSheetByName('Suppliers');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const suppliers = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) {
      suppliers.push(data[i][1]);
    }
  }
  return suppliers;
}

function getCategories() {
  const sheet = SS.getSheetByName('Categories');
  if (!sheet) return ['Proteins', 'Vegetables', 'Dairy', 'Grains', 'Spices', 'Sauces', 'Packaging', 'Beverages', 'Other'];
  const data = sheet.getDataRange().getValues();
  const categories = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      categories.push(data[i][0]);
    }
  }
  return categories.length > 0 ? categories : ['Proteins', 'Vegetables', 'Dairy', 'Grains', 'Spices', 'Sauces', 'Packaging', 'Beverages', 'Other'];
}

function addNewItem(itemData) {
  const sheet = SS.getSheetByName(INVENTORY_SHEET);
  const lastRow = sheet.getLastRow();
  const newId = 'INV-' + String(lastRow).padStart(4, '0');
  const lowPct = getSettingValue('Low Stock Alert %') || 20;
  const criticalPct = getSettingValue('Critical Stock Alert %') || 10;
  const nonPerishable = normalizeBoolean(itemData.nonPerishable);
  const newRow = [
    newId,                         // A
    itemData.name,                 // B
    itemData.category,             // C
    itemData.unit,                 // D
    itemData.initialStock || 0,    // E
    itemData.reorderLevel,         // F
    itemData.reorderQty,           // G
    itemData.unitCost,             // H
    '=E' + (lastRow+1) + '*H' + (lastRow+1), // I total value
    itemData.supplier,             // J
    itemData.location,             // K
    nonPerishable ? '' : (itemData.expiryDate || ''),     // L expiry (empty if non-perishable)
    nonPerishable ? '' : (itemData.expiryDate ? '=L' + (lastRow+1) + '-TODAY()' : ''), // M days (empty if non-perishable)
    nonPerishable,                 // N non-perishable flag
    '',                            // O status
    new Date()                     // P updated
  ];
  sheet.appendRow(newRow);
  // Insert checkbox for non‑perishable column
  const cell = sheet.getRange(lastRow+1, 14);
  cell.insertCheckboxes();
  // Initial status update
  updateInventoryStatusForItem(newId);
  return newId;
}

function processStockIn(data) {
  const inventorySheet = SS.getSheetByName(INVENTORY_SHEET);
  const stockInSheet = SS.getSheetByName(STOCK_IN_SHEET);
  const batchesSheet = SS.getSheetByName(BATCHES_SHEET);
  const inventoryData = inventorySheet.getDataRange().getValues();
  let itemRow = -1;
  let itemName = '';
  let nonPerishable = false;
  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0] === data.itemId) {
      itemRow = i + 1;
      itemName = inventoryData[i][1];
      nonPerishable = inventoryData[i][COLUMN_NONPERISHABLE] === true;
      break;
    }
  }
  if (itemRow === -1) throw new Error('Item not found: ' + data.itemId);
  const qty = parseFloat(data.quantity);
  if (isNaN(qty) || qty <= 0) throw new Error('Invalid quantity. Must be a positive number.');
  const expiryDate = (nonPerishable || !data.expiryDate) ? '' : new Date(data.expiryDate);
  const now = new Date();
  // --- Update Inventory aggregate ---
  const currentStock = inventorySheet.getRange(itemRow, 5).getValue();
  inventorySheet.getRange(itemRow, 5).setValue(currentStock + qty);
  inventorySheet.getRange(itemRow, COLUMN_UPDATED).setValue(now);
  // --- Record in Stock In sheet ---
  stockInSheet.appendRow([
    now,
    data.itemId,
    itemName,
    qty,
    data.unitCost || '',
    data.supplier || '',
    data.invoiceNo || '',
    expiryDate,
    data.receivedBy || '',
    data.notes || ''
  ]);
  // --- Add new Batch row ---
  const lastRow = batchesSheet.getLastRow();
  const batchId = 'BATCH-' + String(lastRow).padStart(4, '0');
  const daysToExpiry = expiryDate ? Math.floor((expiryDate - now) / (1000 * 3600 * 24)) : '';
  const statusFormula = `=IF(F${lastRow + 1}<TODAY(),"EXPIRED",IF(F${lastRow + 1}-TODAY()<=7,"EXPIRING","OK"))`;
  batchesSheet.appendRow([
    batchId,
    data.itemId,
    itemName,
    qty,
    data.unitCost || '',
    expiryDate || '',
    data.supplier || '',
    now,
    data.notes || '',
    daysToExpiry,
    '', // placeholder for status formula
  ]);
  batchesSheet.getRange(lastRow + 1, 11).setFormula(statusFormula);
  updateInventoryStatusForItem(data.itemId);
  return true;
}

function processStockOut(data) {
  const inventorySheet = SS.getSheetByName(INVENTORY_SHEET);
  const stockOutSheet = SS.getSheetByName(STOCK_OUT_SHEET);
  const batchesSheet = SS.getSheetByName('Batches');

  const inventoryData = inventorySheet.getDataRange().getValues();
  let itemRow = -1;
  let itemName = '';

  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0] === data.itemId) {
      itemRow = i + 1;
      itemName = inventoryData[i][1];
      break;
    }
  }
  if (itemRow === -1) throw new Error('Item not found.');

  const usageQty = parseFloat(data.quantity);
  if (isNaN(usageQty) || usageQty <= 0) throw new Error('Invalid quantity. Must be a positive number.');
  const today = new Date();

  // Deduct from the selected batch
  const batchesData = batchesSheet.getDataRange().getValues();
  const batchRow = batchesData.findIndex(r => r[0] === data.batchId);
  if (batchRow === -1) throw new Error('Selected batch not found.');
  const batch = batchesData[batchRow];
  const available = parseFloat(batch[3]);

  if (usageQty > available) throw new Error(`Not enough in ${data.batchId} (available: ${available}).`);
  batchesSheet.getRange(batchRow + 1, 4).setValue(available - usageQty);

  // Update inventory
  const invStock = inventorySheet.getRange(itemRow, 5).getValue();
  inventorySheet.getRange(itemRow, 5).setValue(invStock - usageQty);
  inventorySheet.getRange(itemRow, COLUMN_UPDATED).setValue(today);

  // Log usage
  stockOutSheet.appendRow([
    today,
    data.itemId,
    itemName,
    usageQty,
    data.reason || 'Production',
    data.orderNo || '',
    data.recordedBy || `Batch: ${data.batchId}`
  ]);

  return true;
}


// ============================================
// AUTO-UPDATE FUNCTIONS
// ============================================

function recalculateAllStock() {
  const inventorySheet = SS.getSheetByName(INVENTORY_SHEET);
  const batchesSheet = SS.getSheetByName('Batches');
  const inventoryData = inventorySheet.getDataRange().getValues();
  const batchesData = batchesSheet.getDataRange().getValues();

  const totals = {};

  for (let i = 1; i < batchesData.length; i++) {
    const row = batchesData[i];
    const itemId = row[1];
    const qty = parseFloat(row[3]) || 0;
    totals[itemId] = (totals[itemId] || 0) + qty;
  }

  for (let i = 1; i < inventoryData.length; i++) {
    const id = inventoryData[i][0];
    if (id) {
      const stock = totals[id] || 0;
      inventorySheet.getRange(i + 1, 5).setValue(stock);
      inventorySheet.getRange(i + 1, COLUMN_UPDATED).setValue(new Date());
      updateInventoryStatusForItem(id);
    }
  }

  applyBatchStatusColors();
  applyInventoryStatusColors();

  SpreadsheetApp.getUi().alert('Inventory recalculated from batch data.');
}


function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Auto-update timestamp when inventory is edited
  if (sheet.getName() === INVENTORY_SHEET && range.getColumn() >= 2 && range.getColumn() <= 12) {
    const row = range.getRow();
    if (row > 1) {
      sheet.getRange(row, 15).setValue(new Date());
    }
  }
}

// ============================================
// DASHBOARD FUNCTIONS
// ============================================

function updateDashboard() {
  const dashboardSheet = SS.getSheetByName(DASHBOARD_SHEET);
  const inventorySheet = SS.getSheetByName(INVENTORY_SHEET);
  
  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('Dashboard sheet not found. Please create a sheet named "Dashboard".');
    return;
  }
  
  dashboardSheet.clear();
  
  const inventoryData = inventorySheet.getDataRange().getValues();
  
  // Calculate metrics
  let totalItems = 0;
  let totalValue = 0;
  let lowStockItems = 0;
  let criticalItems = 0;
  let expiringItems = 0;
  let expiredItems = 0;
  const categoryBreakdown = {};
  
  const today = new Date();
  const warningDays = getSettingValue('Expiry Warning Days') || 7;
  
  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0]) {
      totalItems++;
      totalValue += parseFloat(inventoryData[i][8]) || 0;

      const status = inventoryData[i][COLUMN_STATUS];
      if (status === 'LOW') lowStockItems++;
      if (status === 'CRITICAL' || status === 'OUT') criticalItems++;
      if (status === 'EXPIRED') expiredItems++;

      const daysToExpiry = inventoryData[i][12];
      const isNonPerishable = inventoryData[i][COLUMN_NONPERISHABLE] === true;
      if (!isNonPerishable && daysToExpiry && daysToExpiry > 0 && daysToExpiry <= warningDays) {
        expiringItems++;
      }

      const category = inventoryData[i][2] || 'Uncategorized';
      if (!categoryBreakdown[category]) {
        categoryBreakdown[category] = { count: 0, value: 0 };
      }
      categoryBreakdown[category].count++;
      categoryBreakdown[category].value += parseFloat(inventoryData[i][8]) || 0;
    }
  }
  
  // Build dashboard
  const currency = getSettingValue('Currency') || '₹';
  
  dashboardSheet.getRange('A1').setValue('📊 INVENTORY DASHBOARD').setFontSize(18).setFontWeight('bold');
  dashboardSheet.getRange('A2').setValue('Last Updated: ' + new Date().toLocaleString());
  
  // Key Metrics
  dashboardSheet.getRange('A4').setValue('KEY METRICS').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  dashboardSheet.getRange('A5').setValue('Total Items');
  dashboardSheet.getRange('B5').setValue(totalItems);
  dashboardSheet.getRange('A6').setValue('Total Inventory Value');
  dashboardSheet.getRange('B6').setValue(currency + ' ' + totalValue.toFixed(2));
  dashboardSheet.getRange('A7').setValue('Low Stock Items');
  dashboardSheet.getRange('B7').setValue(lowStockItems).setBackground(lowStockItems > 0 ? '#fff3cd' : '#d4edda');
  dashboardSheet.getRange('A8').setValue('Critical/Out of Stock');
  dashboardSheet.getRange('B8').setValue(criticalItems).setBackground(criticalItems > 0 ? '#f8d7da' : '#d4edda');
  dashboardSheet.getRange('A9').setValue('Expiring Soon (≤' + warningDays + ' days)');
  dashboardSheet.getRange('B9').setValue(expiringItems).setBackground(expiringItems > 0 ? '#fff3cd' : '#d4edda');
  dashboardSheet.getRange('A10').setValue('Expired Items');
  dashboardSheet.getRange('B10').setValue(expiredItems).setBackground(expiredItems > 0 ? '#f8d7da' : '#d4edda');
  
  // Category Breakdown
  dashboardSheet.getRange('A12').setValue('CATEGORY BREAKDOWN').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  dashboardSheet.getRange('A13:C13').setValues([['Category', 'Items', 'Value']]).setFontWeight('bold');
  
  let row = 14;
  for (const category in categoryBreakdown) {
    dashboardSheet.getRange(row, 1).setValue(category);
    dashboardSheet.getRange(row, 2).setValue(categoryBreakdown[category].count);
    dashboardSheet.getRange(row, 3).setValue(currency + ' ' + categoryBreakdown[category].value.toFixed(2));
    row++;
  }
  
  // Format columns
  dashboardSheet.setColumnWidth(1, 200);
  dashboardSheet.setColumnWidth(2, 150);
  dashboardSheet.setColumnWidth(3, 150);
  
  SpreadsheetApp.getUi().alert('Dashboard updated successfully!');
}

// ============================================
// REPORT FUNCTIONS
// ============================================

function generateLowStockReport() {
  const inventorySheet = SS.getSheetByName(INVENTORY_SHEET);
  const data = inventorySheet.getDataRange().getValues();
  
  let report = 'LOW STOCK REPORT\n';
  report += 'Generated: ' + new Date().toLocaleString() + '\n';
  report += '='.repeat(50) + '\n\n';
  
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const status = data[i][COLUMN_STATUS];
    if (status === 'LOW' || status === 'CRITICAL' || status === 'OUT') {
      count++;
      report += `${data[i][1]} (${data[i][0]})\n`;
      report += `  Status: ${status}\n`;
      report += `  Current: ${data[i][4]} ${data[i][3]}\n`;
      report += `  Reorder Level: ${data[i][5]} ${data[i][3]}\n`;
      report += `  Suggested Order: ${data[i][6]} ${data[i][3]}\n`;
      report += `  Supplier: ${data[i][9] || 'Not specified'}\n\n`;
    }
  }
  
  if (count === 0) {
    report += 'All items are adequately stocked! ✓\n';
  } else {
    report = `Found ${count} items needing attention:\n\n` + report.split('\n\n').slice(1).join('\n\n');
  }
  
  SpreadsheetApp.getUi().alert(report);
}

function generateExpiryReport() {
  const inventorySheet = SS.getSheetByName(INVENTORY_SHEET);
  const data = inventorySheet.getDataRange().getValues();
  const warningDays = getSettingValue('Expiry Warning Days') || 7;
  
  let report = 'EXPIRY REPORT\n';
  report += 'Generated: ' + new Date().toLocaleString() + '\n';
  report += '='.repeat(50) + '\n\n';
  
  const expired = [];
  const expiringSoon = [];
  
  for (let i = 1; i < data.length; i++) {
    const isNonPerishable = data[i][COLUMN_NONPERISHABLE] === true;
    if (isNonPerishable) continue; // Skip non-perishable items
    const daysToExpiry = data[i][12];
    if (daysToExpiry !== '' && daysToExpiry !== null) {
      if (daysToExpiry < 0) {
        expired.push({ name: data[i][1], days: Math.abs(daysToExpiry), stock: data[i][4], unit: data[i][3] });
      } else if (daysToExpiry <= warningDays) {
        expiringSoon.push({ name: data[i][1], days: daysToExpiry, stock: data[i][4], unit: data[i][3] });
      }
    }
  }
  
  if (expired.length > 0) {
    report += '🚨 EXPIRED ITEMS:\n';
    expired.forEach(item => {
      report += `  • ${item.name}: Expired ${item.days} days ago (${item.stock} ${item.unit} remaining)\n`;
    });
    report += '\n';
  }
  
  if (expiringSoon.length > 0) {
    report += '⚠️ EXPIRING SOON:\n';
    expiringSoon.sort((a, b) => a.days - b.days);
    expiringSoon.forEach(item => {
      report += `  • ${item.name}: ${item.days} days left (${item.stock} ${item.unit})\n`;
    });
  }
  
  if (expired.length === 0 && expiringSoon.length === 0) {
    report += 'No expiring items found! ✓\n';
  }
  
  SpreadsheetApp.getUi().alert(report);
}

function generateValuationReport() {
  const inventorySheet = SS.getSheetByName(INVENTORY_SHEET);
  const data = inventorySheet.getDataRange().getValues();
  const currency = getSettingValue('Currency') || '₹';
  
  let totalValue = 0;
  const categoryValues = {};
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const value = parseFloat(data[i][8]) || 0;
      totalValue += value;
      
      const category = data[i][2] || 'Uncategorized';
      categoryValues[category] = (categoryValues[category] || 0) + value;
    }
  }
  
  let report = 'STOCK VALUATION REPORT\n';
  report += 'Generated: ' + new Date().toLocaleString() + '\n';
  report += '='.repeat(50) + '\n\n';
  report += `TOTAL INVENTORY VALUE: ${currency} ${totalValue.toFixed(2)}\n\n`;
  report += 'BY CATEGORY:\n';
  
  Object.keys(categoryValues).sort((a, b) => categoryValues[b] - categoryValues[a]).forEach(category => {
    const percentage = ((categoryValues[category] / totalValue) * 100).toFixed(1);
    report += `  ${category}: ${currency} ${categoryValues[category].toFixed(2)} (${percentage}%)\n`;
  });
  
  SpreadsheetApp.getUi().alert(report);
}

function generateUsageReport() {
  const stockOutSheet = SS.getSheetByName(STOCK_OUT_SHEET);
  const data = stockOutSheet.getDataRange().getValues();
  const currency = getSettingValue('Currency') || '₹';
  
  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  
  const usage = {};
  const reasons = {};
  
  for (let i = 1; i < data.length; i++) {
    const date = new Date(data[i][0]);
    if (date >= sevenDaysAgo) {
      const itemName = data[i][2];
      const quantity = parseFloat(data[i][3]) || 0;
      const reason = data[i][4] || 'Unknown';
      
      usage[itemName] = (usage[itemName] || 0) + quantity;
      reasons[reason] = (reasons[reason] || 0) + 1;
    }
  }
  
  let report = 'USAGE REPORT (Last 7 Days)\n';
  report += 'Generated: ' + new Date().toLocaleString() + '\n';
  report += '='.repeat(50) + '\n\n';
  
  report += 'TOP USED ITEMS:\n';
  Object.keys(usage).sort((a, b) => usage[b] - usage[a]).slice(0, 10).forEach(item => {
    report += `  • ${item}: ${usage[item]}\n`;
  });
  
  report += '\nUSAGE BY REASON:\n';
  Object.keys(reasons).forEach(reason => {
    report += `  • ${reason}: ${reasons[reason]} transactions\n`;
  });
  
  SpreadsheetApp.getUi().alert(report);
}

// ============================================
// EMAIL ALERTS
// ============================================

function sendAlertEmail() {
  const email = getSettingValue('Email Alerts');
  if (!email) {
    SpreadsheetApp.getUi().alert('No email configured in Settings. Please add your email address.');
    return;
  }
  
  const inventorySheet = SS.getSheetByName(INVENTORY_SHEET);
  const data = inventorySheet.getDataRange().getValues();
  const warningDays = getSettingValue('Expiry Warning Days') || 7;
  const kitchenName = getSettingValue('Kitchen Name') || 'Cloud Kitchen';
  
  const lowStock = [];
  const critical = [];
  const expiring = [];
  const expired = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const status = data[i][COLUMN_STATUS];
      const daysToExpiry = data[i][12];
      const isNonPerishable = data[i][COLUMN_NONPERISHABLE] === true;

      if (status === 'LOW') lowStock.push(data[i][1]);
      if (status === 'CRITICAL' || status === 'OUT') critical.push(data[i][1]);
      if (status === 'EXPIRED') expired.push(data[i][1]);
      if (!isNonPerishable && daysToExpiry > 0 && daysToExpiry <= warningDays) expiring.push(`${data[i][1]} (${daysToExpiry} days)`);
    }
  }
  
  if (lowStock.length === 0 && critical.length === 0 && expiring.length === 0 && expired.length === 0) {
    SpreadsheetApp.getUi().alert('No alerts to send. All inventory levels are good!');
    return;
  }
  
  let body = `<h2>${kitchenName} - Inventory Alert</h2>`;
  body += `<p>Generated: ${new Date().toLocaleString()}</p>`;
  
  if (critical.length > 0) {
    body += `<h3 style="color: red;">🚨 Critical/Out of Stock (${critical.length})</h3><ul>`;
    critical.forEach(item => body += `<li>${item}</li>`);
    body += '</ul>';
  }
  
  if (expired.length > 0) {
    body += `<h3 style="color: red;">⛔ Expired Items (${expired.length})</h3><ul>`;
    expired.forEach(item => body += `<li>${item}</li>`);
    body += '</ul>';
  }
  
  if (lowStock.length > 0) {
    body += `<h3 style="color: orange;">⚠️ Low Stock (${lowStock.length})</h3><ul>`;
    lowStock.forEach(item => body += `<li>${item}</li>`);
    body += '</ul>';
  }
  
  if (expiring.length > 0) {
    body += `<h3 style="color: orange;">📅 Expiring Soon (${expiring.length})</h3><ul>`;
    expiring.forEach(item => body += `<li>${item}</li>`);
    body += '</ul>';
  }
  
  body += `<p><a href="${SS.getUrl()}">Open Inventory Spreadsheet</a></p>`;
  
  GmailApp.sendEmail(email, `[${kitchenName}] Inventory Alert`, '', { htmlBody: body });
  SpreadsheetApp.getUi().alert('Alert email sent to ' + email);
}

// Daily trigger for automatic alerts
function dailyAlertCheck() {
  const email = getSettingValue('Email Alerts');
  if (email) {
    sendAlertEmail();
  }
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

function getSettingValue(settingName) {
  const sheet = SS.getSheetByName(SETTINGS_SHEET);
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === settingName) {
      return data[i][1];
    }
  }
  return null;
}

function setSettingValue(settingName, value) {
  const sheet = SS.getSheetByName(SETTINGS_SHEET);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === settingName) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  // If not found, add new row
  sheet.appendRow([settingName, value]);
}

function saveSettings(settings) {
  setSettingValue('Kitchen Name', settings.kitchenName);
  setSettingValue('Currency', settings.currency);
  setSettingValue('Low Stock Alert %', settings.lowStockPct);
  setSettingValue('Expiry Warning Days', settings.expiryDays);
  setSettingValue('Email Alerts', settings.email);
  return true;
}

function getSettings() {
  return {
    kitchenName: getSettingValue('Kitchen Name') || '',
    currency: getSettingValue('Currency') || '₹',
    lowStockPct: getSettingValue('Low Stock Alert %') || 20,
    expiryDays: getSettingValue('Expiry Warning Days') || 7,
    email: getSettingValue('Email Alerts') || ''
  };
}

function updateInventoryStatusForItem(itemId) {
  const invSheet = SS.getSheetByName(INVENTORY_SHEET);
  const batchesSheet = SS.getSheetByName('Batches');
  const invData = invSheet.getDataRange().getValues();
  const batches = batchesSheet.getDataRange().getValues();
  const today = new Date();

  let invRow = -1, nonPerishable = false;
  for (let i = 1; i < invData.length; i++) {
    if (invData[i][0] === itemId) {
      invRow = i + 1;
      nonPerishable = invData[i][COLUMN_NONPERISHABLE] === true;
      break;
    }
  }
  if (invRow === -1) return;

  if (nonPerishable) {
    invSheet.getRange(invRow, COLUMN_NONPERISHABLE + 1).setValue(true);  // Column N checkbox
    invSheet.getRange(invRow, COLUMN_STATUS + 1).setValue('N/A');        // Column O status
    return;
  }

  // find active batches for perishable items
  const itemBatches = batches.filter((r, i) => {
    if (i === 0) return false; // Skip header
    if (r[1] !== itemId) return false;
    const qty = parseFloat(r[3]);
    if (isNaN(qty) || qty <= 0) return false;
    if (!r[5]) return false; // No expiry date
    return true;
  });

  if (itemBatches.length === 0) {
    invSheet.getRange(invRow, 12).setValue('');
    invSheet.getRange(invRow, 13).setValue('');
    invSheet.getRange(invRow, COLUMN_STATUS + 1).setValue('OUT');
    return;
  }

  // earliest non-expired expiry among batches
  const expiryDates = itemBatches
    .map(b => new Date(b[5]))
    .filter(d => d instanceof Date && !isNaN(d))
    .sort((a, b) => a - b);

  const earliest = expiryDates[0];
  const days = Math.floor((earliest - today) / (1000 * 3600 * 24));

  let status = 'OK';
  if (days < 0) status = 'EXPIRED';
  else if (days <= 7) status = 'EXPIRING SOON';

  invSheet.getRange(invRow, 12).setValue(earliest);
  invSheet.getRange(invRow, 13).setValue(days);
  invSheet.getRange(invRow, COLUMN_STATUS + 1).setValue(status);
}



function applyBatchStatusColors() {
  const sheet = SS.getSheetByName('Batches');
  const statusCol = 11; // column K
  const range = sheet.getRange(2, statusCol, sheet.getLastRow() - 1);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('EXPIRED').setBackground('#f8d7da').setFontColor('#721c24')
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('EXPIRING').setBackground('#fff3cd').setFontColor('#856404')
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OK').setBackground('#d4edda').setFontColor('#155724')
      .setRanges([range]).build()
  ];

  sheet.setConditionalFormatRules(rules);
}

function applyInventoryStatusColors() {
  const sheet = SS.getSheetByName('Inventory');
  const statusCol = COLUMN_STATUS + 1;  // 1-indexed for Google Sheets
  const range = sheet.getRange(2, statusCol, sheet.getLastRow() - 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OK')
      .setBackground('#d4edda').setFontColor('#155724').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('EXPIRING')
      .setBackground('#fff3cd').setFontColor('#856404').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('EXPIRED')
      .setBackground('#f8d7da').setFontColor('#721c24').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OUT')
      .setBackground('#e2e3e5').setFontColor('#383d41').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('N/A')
      .setBackground('#f0f0f0').setFontColor('#666666').setRanges([range]).build(),
  ];
  sheet.setConditionalFormatRules(rules);
}



// ============================================
// HTML TEMPLATES
// ============================================

function getStockInHtml() {
  const items = getInventoryItems();
  const suppliers = getSuppliers();

  let itemOptions = items.map(item =>
    `<option value="${item.id}" data-non-perishable="${item.nonPerishable}">${item.name} (${item.currentStock} ${item.unit})</option>`
  ).join('');

  let supplierOptions = suppliers.map(s => `<option value="${s}">${s}</option>`).join('');
  let itemsData = JSON.stringify(items.reduce((acc, item) => { acc[item.id] = item.nonPerishable; return acc; }, {}));

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input, select { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }
        button { background: #4285f4; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px; }
        button:hover { background: #3367d6; }
        .cancel { background: #ccc; color: #333; }
        .required::after { content: " *"; color: red; }
        #expiryDateGroup { transition: opacity 0.3s; }
        #expiryDateGroup.hidden { opacity: 0.5; pointer-events: none; }
      </style>
    </head>
    <body>
      <form id="stockInForm">
        <div class="form-group">
          <label class="required">Item</label>
          <select name="itemId" id="itemSelect" required>
            <option value="">-- Select Item --</option>
            ${itemOptions}
          </select>
        </div>
        <div class="form-group">
          <label class="required">Quantity Received</label>
          <input type="number" name="quantity" step="0.01" min="0.01" required>
        </div>
        <div class="form-group">
          <label>Unit Cost</label>
          <input type="number" name="unitCost" step="0.01" min="0">
        </div>
        <div class="form-group">
          <label>Supplier</label>
          <select name="supplier">
            <option value="">-- Select Supplier --</option>
            ${supplierOptions}
          </select>
        </div>
        <div class="form-group">
          <label>Invoice #</label>
          <input type="text" name="invoiceNo">
        </div>
        <div class="form-group" id="expiryDateGroup">
          <label>Expiry Date</label>
          <input type="date" name="expiryDate" id="expiryDateInput">
        </div>
        <div class="form-group">
          <label>Received By</label>
          <input type="text" name="receivedBy">
        </div>
        <div class="form-group">
          <label>Notes</label>
          <input type="text" name="notes">
        </div>
        <button type="submit">Add Stock</button>
        <button type="button" class="cancel" onclick="google.script.host.close()">Cancel</button>
      </form>
      <script>
        const itemsData = ${itemsData};
        const itemSelect = document.getElementById('itemSelect');
        const expiryDateGroup = document.getElementById('expiryDateGroup');
        const expiryDateInput = document.getElementById('expiryDateInput');

        // Handle item selection
        itemSelect.addEventListener('change', function() {
          const selectedItemId = this.value;
          const isNonPerishable = itemsData[selectedItemId] || false;

          if (isNonPerishable) {
            expiryDateGroup.classList.add('hidden');
            expiryDateInput.value = '';
          } else {
            expiryDateGroup.classList.remove('hidden');
          }
        });

        document.getElementById('stockInForm').addEventListener('submit', function(e) {
          e.preventDefault();
          const selectedItemId = itemSelect.value;
          const isNonPerishable = itemsData[selectedItemId] || false;

          const formData = new FormData(this);
          // Clear expiry date if item is non-perishable
          if (isNonPerishable) {
            formData.set('expiryDate', '');
          }

          const data = Object.fromEntries(formData.entries());
          google.script.run
            .withSuccessHandler(() => {
              alert('Stock added successfully!');
              google.script.host.close();
            })
            .withFailureHandler(err => alert('Error: ' + err.message))
            .processStockIn(data);
        });
      </script>
    </body>
    </html>
  `;
}

function getStockOutHtml() {
  const items = getInventoryItems();
  const batchesSheet = SS.getSheetByName('Batches');
  const data = batchesSheet.getDataRange().getValues();

  // Group active batches by Item ID
  const today = new Date();
  const batchMap = {};
  for (let i = 1; i < data.length; i++) {
    const [batchId, itemId, itemName, qty, , expiry] = data[i];
    const remaining = parseFloat(qty) || 0;
    if (remaining > 0 && itemId) {
      if (!batchMap[itemId]) batchMap[itemId] = [];
      const daysLeft = expiry ? Math.floor((new Date(expiry) - today) / (1000 * 3600 * 24)) : '';
      let expStr = expiry ? Utilities.formatDate(new Date(expiry), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
      batchMap[itemId].push(`${batchId} | Qty: ${remaining} | Exp: ${expStr}${daysLeft < 0 ? ' (EXPIRED!)' : daysLeft <= 7 ? ' (Expiring)' : ''}`);
    }
  }

  // Create item and batch dropdown HTML
  const itemOptions = items.map(i => `<option value="${i.id}">${i.name}</option>`).join('');

  return `
  <!DOCTYPE html>
  <html>
  <head>
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display:block; margin-bottom:5px; font-weight:bold; }
      select,input { width:100%; padding:8px; border:1px solid #ddd; border-radius:4px; }
      button { background:#ea4335; color:white; padding:10px 20px; border:none; border-radius:4px; cursor:pointer; }
      button:hover { background:#c5221f; }
      .cancel { background:#ccc; color:#333; }
    </style>
  </head>
  <body>
    <form id="stockOutForm">
      <div class="form-group">
        <label>Item</label>
        <select id="itemSelect" name="itemId" required>
          <option value="">--Select Item--</option>
          ${itemOptions}
        </select>
      </div>
      <div class="form-group" id="batchDiv" style="display:none;">
        <label>Batch</label>
        <select id="batchSelect" name="batchId" required></select>
      </div>
      <div class="form-group">
        <label>Quantity Used</label>
        <input type="number" name="quantity" step="0.01" min="0.01" required>
      </div>
      <div class="form-group">
        <label>Reason</label>
        <select name="reason" required>
          <option value="Production">Production</option>
          <option value="Waste">Waste</option>
          <option value="Transfer">Transfer</option>
          <option value="Adjustment">Inventory Adjustment</option>
          <option value="Sample">Sample</option>
          <option value="Other">Other</option>
        </select>
      </div>
      <div class="form-group">
        <label>Order/Batch #</label>
        <input type="text" name="orderNo">
      </div>
      <div class="form-group">
        <label>Recorded By</label>
        <input type="text" name="recordedBy">
      </div>
      <button type="submit">Record Usage</button>
      <button type="button" class="cancel" onclick="google.script.host.close()">Cancel</button>
    </form>

    <script>
      const batchData = ${JSON.stringify(batchMap)};
      document.getElementById('itemSelect').addEventListener('change', function() {
        const batches = batchData[this.value] || [];
        const batchSelect = document.getElementById('batchSelect');
        const batchDiv = document.getElementById('batchDiv');

        // Clear and rebuild options using safe DOM methods
        batchSelect.innerHTML = '';
        batches.forEach(b => {
          const batchId = b.split(' | ')[0];
          const option = document.createElement('option');
          option.value = batchId;
          option.textContent = b;
          batchSelect.appendChild(option);
        });

        if (batches.length === 0) {
          const option = document.createElement('option');
          option.value = '';
          option.textContent = '(No available batch)';
          batchSelect.appendChild(option);
        }

        batchDiv.style.display = batches.length ? 'block' : 'none';
      });

      document.getElementById('stockOutForm').addEventListener('submit', function(e) {
        e.preventDefault();
        const formData = new FormData(this);
        const data = Object.fromEntries(formData.entries());
        google.script.run
          .withSuccessHandler(() => { alert('Stock usage recorded.'); google.script.host.close(); })
          .withFailureHandler(err => alert('Error: ' + err.message))
          .processStockOut(data);
      });
    </script>
  </body>
  </html>`;
}


function getNewItemHtml() {
  const categories = getCategories();
  const suppliers = getSuppliers();

  let categoryOptions = categories.map(c => `<option value="${c}">${c}</option>`).join('');
  let supplierOptions = suppliers.map(s => `<option value="${s}">${s}</option>`).join('');

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input, select { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }
        button { background: #34a853; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px; }
        button:hover { background: #2d8e47; }
        .cancel { background: #ccc; color: #333; }
        .required::after { content: " *"; color: red; }
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        .checkbox-group { display: flex; align-items: center; }
        .checkbox-group input[type="checkbox"] { width: auto; margin-right: 8px; cursor: pointer; }
        .checkbox-group label { display: inline; margin-bottom: 0; }
        #expiryDateGroup { transition: opacity 0.3s; }
        #expiryDateGroup.hidden { opacity: 0.5; pointer-events: none; }
      </style>
    </head>
    <body>
      <form id="newItemForm">
        <div class="form-group">
          <label class="required">Item Name</label>
          <input type="text" name="name" required>
        </div>
        <div class="two-col">
          <div class="form-group">
            <label class="required">Category</label>
            <select name="category" required>
              ${categoryOptions}
            </select>
          </div>
          <div class="form-group">
            <label class="required">Unit</label>
            <select name="unit" required>
              <option value="kg">kg</option>
              <option value="g">g</option>
              <option value="L">L</option>
              <option value="mL">mL</option>
              <option value="pcs">pcs</option>
              <option value="dozen">dozen</option>
              <option value="pack">pack</option>
              <option value="box">box</option>
            </select>
          </div>
        </div>
        <div class="two-col">
          <div class="form-group">
            <label>Initial Stock</label>
            <input type="number" name="initialStock" step="0.01" min="0" value="0">
          </div>
          <div class="form-group">
            <label class="required">Unit Cost</label>
            <input type="number" name="unitCost" step="0.01" min="0" required>
          </div>
        </div>
        <div class="two-col">
          <div class="form-group">
            <label class="required">Reorder Level</label>
            <input type="number" name="reorderLevel" step="0.01" min="0" required>
          </div>
          <div class="form-group">
            <label class="required">Reorder Quantity</label>
            <input type="number" name="reorderQty" step="0.01" min="0" required>
          </div>
        </div>
        <div class="form-group">
          <label>Supplier</label>
          <select name="supplier">
            <option value="">-- Select Supplier --</option>
            ${supplierOptions}
          </select>
        </div>
        <div class="form-group">
          <label>Storage Location</label>
          <input type="text" name="location" placeholder="e.g., Fridge 1, Shelf A3">
        </div>
        <div class="checkbox-group">
          <input type="checkbox" id="nonPerishableCheckbox" name="nonPerishable">
          <label for="nonPerishableCheckbox">Non-Perishable Item</label>
        </div>
        <div class="form-group" id="expiryDateGroup">
          <label>Expiry Date</label>
          <input type="date" name="expiryDate">
        </div>
        <button type="submit">Add Item</button>
        <button type="button" class="cancel" onclick="google.script.host.close()">Cancel</button>
      </form>
      <script>
        const nonPerishableCheckbox = document.getElementById('nonPerishableCheckbox');
        const expiryDateGroup = document.getElementById('expiryDateGroup');

        // Handle checkbox change
        nonPerishableCheckbox.addEventListener('change', function() {
          if (this.checked) {
            expiryDateGroup.classList.add('hidden');
            document.querySelector('input[name="expiryDate"]').value = '';
          } else {
            expiryDateGroup.classList.remove('hidden');
          }
        });

        document.getElementById('newItemForm').addEventListener('submit', function(e) {
          e.preventDefault();
          const formData = new FormData(this);
          const data = Object.fromEntries(formData.entries());
          google.script.run
            .withSuccessHandler((id) => {
              alert('Item added successfully! ID: ' + id);
              google.script.host.close();
            })
            .withFailureHandler(err => alert('Error: ' + err.message))
            .addNewItem(data);
        });
      </script>
    </body>
    </html>
  `;
}

function getSettingsHtml() {
  const settings = getSettings();
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input, select { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }
        button { background: #4285f4; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px; }
        button:hover { background: #3367d6; }
        .cancel { background: #ccc; color: #333; }
      </style>
    </head>
    <body>
      <form id="settingsForm">
        <div class="form-group">
          <label>Kitchen Name</label>
          <input type="text" name="kitchenName" value="${settings.kitchenName}">
        </div>
        <div class="form-group">
          <label>Currency Symbol</label>
          <input type="text" name="currency" value="${settings.currency}">
        </div>
        <div class="form-group">
          <label>Low Stock Alert %</label>
          <input type="number" name="lowStockPct" value="${settings.lowStockPct}" min="1" max="100">
        </div>
        <div class="form-group">
          <label>Expiry Warning (days)</label>
          <input type="number" name="expiryDays" value="${settings.expiryDays}" min="1">
        </div>
        <div class="form-group">
          <label>Alert Email Address</label>
          <input type="email" name="email" value="${settings.email}">
        </div>
        <button type="submit">Save Settings</button>
        <button type="button" class="cancel" onclick="google.script.host.close()">Cancel</button>
      </form>
      <script>
        document.getElementById('settingsForm').addEventListener('submit', function(e) {
          e.preventDefault();
          const formData = new FormData(this);
          const data = Object.fromEntries(formData.entries());
          google.script.run
            .withSuccessHandler(() => {
              alert('Settings saved!');
              google.script.host.close();
            })
            .withFailureHandler(err => alert('Error: ' + err.message))
            .saveSettings(data);
        });
      </script>
    </body>
    </html>
  `;
}

// ============================================
// TRIGGER SETUP
// ============================================

function setupDailyTrigger() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyAlertCheck') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new daily trigger at 8 AM
  ScriptApp.newTrigger('dailyAlertCheck')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
    
  SpreadsheetApp.getUi().alert('Daily alert trigger set for 8:00 AM');
}
