// ============================================================================
// SHEET STYLING & LAYOUT UPDATE FUNCTIONS (PRESERVES DATA)
// ============================================================================

/**
 * Master function to update all sheets in the spreadsheet with enterprise styling
 * Preserves all existing data while updating headers, styling, and layout
 */
function updateAllSheetsFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    'Update Sheet Formatting',
    'This will update headers, styling, and layout for all configured sheets while preserving your data. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  const sheets = ss.getSheets();
  let updated = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const config = getSheetConfig_(sheetName);
    if (config) {
      applySheetFormatting_(sheet, config);
      updated++;
    }
  });

  ui.alert('Formatting Complete', `Updated ${updated} sheet(s) with enterprise styling.`, ui.ButtonSet.OK);
}

/**
 * Update a single sheet's formatting by name
 * @param {string} sheetName - Name of the sheet to update
 * @returns {Object} Result with success status and message
 */
function updateSheetFormatting(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  const config = getSheetConfig_(sheetName);
  if (!config) {
    Logger.log('No configuration found for sheet: ' + sheetName);
    return { success: false, message: 'No configuration found for: ' + sheetName };
  }

  applySheetFormatting_(sheet, config);
  return { success: true, message: 'Sheet formatted successfully: ' + sheetName };
}

/**
 * Refresh formatting for the active sheet
 */
function refreshActiveSheetFormatting() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  const config = getSheetConfig_(sheetName);
  if (!config) {
    SpreadsheetApp.getUi().alert('No configuration found for sheet: ' + sheetName);
    return;
  }

  applySheetFormatting_(sheet, config);
  SpreadsheetApp.getUi().alert('Formatting refreshed for: ' + config.displayName);
}

// ============================================================================
// SHEET CONFIGURATIONS
// ============================================================================

/**
 * Get sheet configuration by name
 * @param {string} sheetName - Name of the sheet
 * @returns {Object|null} Configuration object or null
 */
function getSheetConfig_(sheetName) {
  const configs = {
    'Master_Log': {
      displayName: 'Master Log',
      headers: ['Date_Received', 'Warehouse', 'Project', 'Customer_PO_Number', 'BOL_Number', 'Push #', 'Asset Type', 'Manufacturer', 'FBPN', 'MFPN', 'Qty_Received', 'UOM', 'Carrier', 'Notes'],
      theme: 'inbound',
      frozenRows: 1,
      frozenCols: 0,
      columnWidths: { 1: 100, 2: 80, 3: 120, 4: 130, 5: 110, 6: 70, 7: 100, 8: 130, 9: 120, 10: 120, 11: 90, 12: 60, 13: 100, 14: 200 },
      filters: true,
      conditionalFormatting: true
    },
    'Outbound_Log': {
      displayName: 'Outbound Log',
      headers: ['Date', 'Company', 'Task_Number', 'Order_Number', 'Project', 'FBPN', 'Qty', 'Manufacturer', 'Notes'],
      theme: 'outbound',
      frozenRows: 1,
      frozenCols: 0,
      columnWidths: { 1: 100, 2: 150, 3: 110, 4: 130, 5: 120, 6: 120, 7: 80, 8: 130, 9: 200 },
      filters: true,
      conditionalFormatting: true
    },
    'PO_Master': {
      displayName: 'PO Master',
      headers: ['Customer_PO', 'Project', 'Vendor', 'Status', 'Created_Date', 'Expected_Date', 'Total_Value', 'Notes'],
      theme: 'neutral',
      frozenRows: 1,
      frozenCols: 1,
      columnWidths: { 1: 130, 2: 120, 3: 150, 4: 90, 5: 100, 6: 100, 7: 100, 8: 200 },
      filters: true,
      conditionalFormatting: false
    },
    'Item_Master': {
      displayName: 'Item Master',
      headers: ['FBPN', 'Description', 'Manufacturer', 'MFPN', 'Asset_Type', 'Category', 'UOM', 'Unit_Cost', 'Notes'],
      theme: 'neutral',
      frozenRows: 1,
      frozenCols: 1,
      columnWidths: { 1: 120, 2: 200, 3: 130, 4: 130, 5: 100, 6: 100, 7: 60, 8: 90, 9: 200 },
      filters: true,
      conditionalFormatting: false
    },
    'Inventory': {
      displayName: 'Inventory',
      headers: ['FBPN', 'Description', 'Warehouse', 'Location', 'Qty_On_Hand', 'Qty_Reserved', 'Qty_Available', 'Last_Updated'],
      theme: 'inventory',
      frozenRows: 1,
      frozenCols: 1,
      columnWidths: { 1: 120, 2: 200, 3: 100, 4: 100, 5: 100, 6: 100, 7: 100, 8: 130 },
      filters: true,
      conditionalFormatting: true
    }
  };

  return configs[sheetName] || null;
}

/**
 * Register a custom sheet configuration
 * @param {string} sheetName - Name of the sheet
 * @param {Object} config - Configuration object
 */
function registerSheetConfig(sheetName, config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();

  const customConfigs = JSON.parse(props.getProperty('customSheetConfigs') || '{}');
  customConfigs[sheetName] = config;
  props.setProperty('customSheetConfigs', JSON.stringify(customConfigs));

  return { success: true, message: 'Configuration registered for: ' + sheetName };
}

// ============================================================================
// THEME COLORS
// ============================================================================

/**
 * Get theme colors by theme name
 * @param {string} themeName - Theme name
 * @returns {Object} Theme colors
 */
function getThemeColors_(themeName) {
  const themes = {
    inbound: {
      headerBg: '#1e3a5f',
      headerText: '#ffffff',
      accentBg: '#dbeafe',
      accentText: '#1e40af',
      altRowBg: '#f8fafc',
      borderColor: '#cbd5e1',
      highlightBg: '#ecfdf5',
      highlightText: '#059669'
    },
    outbound: {
      headerBg: '#065f46',
      headerText: '#ffffff',
      accentBg: '#d1fae5',
      accentText: '#065f46',
      altRowBg: '#f0fdf4',
      borderColor: '#a7f3d0',
      highlightBg: '#fff7ed',
      highlightText: '#ea580c'
    },
    inventory: {
      headerBg: '#7c3aed',
      headerText: '#ffffff',
      accentBg: '#ede9fe',
      accentText: '#6d28d9',
      altRowBg: '#faf5ff',
      borderColor: '#c4b5fd',
      highlightBg: '#fef3c7',
      highlightText: '#d97706'
    },
    neutral: {
      headerBg: '#374151',
      headerText: '#ffffff',
      accentBg: '#f3f4f6',
      accentText: '#1f2937',
      altRowBg: '#f9fafb',
      borderColor: '#d1d5db',
      highlightBg: '#fef3c7',
      highlightText: '#92400e'
    },
    danger: {
      headerBg: '#991b1b',
      headerText: '#ffffff',
      accentBg: '#fee2e2',
      accentText: '#991b1b',
      altRowBg: '#fef2f2',
      borderColor: '#fecaca',
      highlightBg: '#fef3c7',
      highlightText: '#92400e'
    },
    info: {
      headerBg: '#1e40af',
      headerText: '#ffffff',
      accentBg: '#dbeafe',
      accentText: '#1e40af',
      altRowBg: '#eff6ff',
      borderColor: '#bfdbfe',
      highlightBg: '#ecfdf5',
      highlightText: '#059669'
    }
  };

  return themes[themeName] || themes.neutral;
}

// ============================================================================
// CORE FORMATTING FUNCTIONS
// ============================================================================

/**
 * Apply formatting to a sheet based on configuration
 * @param {Sheet} sheet - Google Sheet object
 * @param {Object} config - Configuration object
 */
function applySheetFormatting_(sheet, config) {
  const theme = getThemeColors_(config.theme);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = Math.max(sheet.getLastColumn(), config.headers.length);

  // Step 1: Update headers (preserving data below)
  updateSheetHeaders_(sheet, config.headers, theme);

  // Step 2: Apply column widths
  if (config.columnWidths) {
    Object.entries(config.columnWidths).forEach(([col, width]) => {
      sheet.setColumnWidth(parseInt(col), width);
    });
  }

  // Step 3: Apply frozen rows/columns
  if (config.frozenRows !== undefined) sheet.setFrozenRows(config.frozenRows);
  if (config.frozenCols !== undefined) sheet.setFrozenColumns(config.frozenCols);

  // Step 4: Style header row
  styleHeaderRow_(sheet, config.headers.length, theme);

  // Step 5: Style data rows (alternating colors)
  if (lastRow > 1) {
    styleDataRows_(sheet, lastRow, config.headers.length, theme);
  }

  // Step 6: Apply filters
  if (config.filters && lastRow > 1) {
    applyFilters_(sheet, lastRow, config.headers.length);
  }

  // Step 7: Apply conditional formatting if enabled
  if (config.conditionalFormatting) {
    applyConditionalFormatting_(sheet, config, theme);
  }

  // Step 8: Apply borders
  applyBorders_(sheet, lastRow, config.headers.length, theme);

  Logger.log('Applied formatting to sheet: ' + config.displayName);
}

/**
 * Update headers without affecting data below row 1
 * @param {Sheet} sheet - Google Sheet object
 * @param {Array} headers - Array of header names
 * @param {Object} theme - Theme colors
 */
function updateSheetHeaders_(sheet, headers, theme) {
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
}

/**
 * Style the header row
 * @param {Sheet} sheet - Google Sheet object
 * @param {number} numCols - Number of columns
 * @param {Object} theme - Theme colors
 */
function styleHeaderRow_(sheet, numCols, theme) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);

  headerRange
    .setBackground(theme.headerBg)
    .setFontColor(theme.headerText)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // Set row height for header
  sheet.setRowHeight(1, 35);
}

/**
 * Style data rows with alternating colors
 * @param {Sheet} sheet - Google Sheet object
 * @param {number} lastRow - Last row with data
 * @param {number} numCols - Number of columns
 * @param {Object} theme - Theme colors
 */
function styleDataRows_(sheet, lastRow, numCols, theme) {
  if (lastRow < 2) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, numCols);

  // Reset background first
  dataRange.setBackground('#ffffff');

  // Apply alternating row colors
  for (let row = 2; row <= lastRow; row++) {
    if (row % 2 === 0) {
      sheet.getRange(row, 1, 1, numCols).setBackground(theme.altRowBg);
    }
  }

  // Apply general data styling
  dataRange
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(false);
}

/**
 * Apply or update filters
 * @param {Sheet} sheet - Google Sheet object
 * @param {number} lastRow - Last row with data
 * @param {number} numCols - Number of columns
 */
function applyFilters_(sheet, lastRow, numCols) {
  // Remove existing filter
  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  // Apply new filter
  const filterRange = sheet.getRange(1, 1, lastRow, numCols);
  filterRange.createFilter();
}

/**
 * Apply conditional formatting rules
 * @param {Sheet} sheet - Google Sheet object
 * @param {Object} config - Sheet configuration
 * @param {Object} theme - Theme colors
 */
function applyConditionalFormatting_(sheet, config, theme) {
  // Clear existing rules
  sheet.clearConditionalFormatRules();

  const rules = [];
  const lastRow = Math.max(sheet.getLastRow(), 100);

  // Find quantity column for highlighting
  const qtyColIndex = config.headers.findIndex(h =>
    h.toLowerCase().includes('qty') || h.toLowerCase().includes('quantity')
  );

  if (qtyColIndex !== -1) {
    const qtyRange = sheet.getRange(2, qtyColIndex + 1, lastRow - 1, 1);

    // Highlight high quantities (green)
    const highQtyRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(100)
      .setBackground(theme.highlightBg)
      .setFontColor(theme.highlightText)
      .setBold(true)
      .setRanges([qtyRange])
      .build();
    rules.push(highQtyRule);

    // Highlight zero quantities (yellow warning)
    const zeroQtyRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#fef3c7')
      .setFontColor('#92400e')
      .setRanges([qtyRange])
      .build();
    rules.push(zeroQtyRule);
  }

  // Find status column for highlighting
  const statusColIndex = config.headers.findIndex(h =>
    h.toLowerCase().includes('status')
  );

  if (statusColIndex !== -1) {
    const statusRange = sheet.getRange(2, statusColIndex + 1, lastRow - 1, 1);

    // Green for complete/active statuses
    const completeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Complete')
      .setBackground('#d1fae5')
      .setFontColor('#065f46')
      .setRanges([statusRange])
      .build();
    rules.push(completeRule);

    // Yellow for pending statuses
    const pendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Pending')
      .setBackground('#fef3c7')
      .setFontColor('#92400e')
      .setRanges([statusRange])
      .build();
    rules.push(pendingRule);

    // Red for cancelled/error statuses
    const cancelledRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Cancel')
      .setBackground('#fee2e2')
      .setFontColor('#dc2626')
      .setRanges([statusRange])
      .build();
    rules.push(cancelledRule);
  }

  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
  }
}

/**
 * Apply borders to the data range
 * @param {Sheet} sheet - Google Sheet object
 * @param {number} lastRow - Last row with data
 * @param {number} numCols - Number of columns
 * @param {Object} theme - Theme colors
 */
function applyBorders_(sheet, lastRow, numCols, theme) {
  const dataRange = sheet.getRange(1, 1, lastRow, numCols);

  dataRange.setBorder(
    true, true, true, true, true, true,
    theme.borderColor,
    SpreadsheetApp.BorderStyle.SOLID
  );

  // Thicker border under header
  const headerBottomRange = sheet.getRange(1, 1, 1, numCols);
  headerBottomRange.setBorder(
    null, null, true, null, null, null,
    theme.headerBg,
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
}

// ============================================================================
// TAB & SHEET MANAGEMENT
// ============================================================================

/**
 * Rename a sheet tab with validation
 * @param {string} oldName - Current sheet name
 * @param {string} newName - New sheet name
 * @returns {Object} Result with success status and message
 */
function renameSheetTab(oldName, newName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(oldName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + oldName };
  }

  // Check if new name already exists
  if (ss.getSheetByName(newName)) {
    return { success: false, message: 'A sheet with name "' + newName + '" already exists' };
  }

  sheet.setName(newName);
  return { success: true, message: 'Sheet renamed from "' + oldName + '" to "' + newName + '"' };
}

/**
 * Add a new configured sheet with enterprise styling
 * @param {string} sheetType - Type of sheet to create (from getSheetConfig_)
 * @returns {Object} Result with success status and message
 */
function addConfiguredSheet(sheetType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getSheetConfig_(sheetType);

  if (!config) {
    return { success: false, message: 'Unknown sheet type: ' + sheetType };
  }

  // Check if sheet already exists
  if (ss.getSheetByName(sheetType)) {
    return { success: false, message: 'Sheet already exists: ' + sheetType };
  }

  // Create new sheet
  const newSheet = ss.insertSheet(sheetType);

  // Apply formatting
  applySheetFormatting_(newSheet, config);

  return { success: true, message: 'Created and formatted sheet: ' + config.displayName };
}

/**
 * Set the tab color for a sheet
 * @param {string} sheetName - Name of the sheet
 * @param {string} color - Hex color code
 * @returns {Object} Result with success status and message
 */
function setSheetTabColor(sheetName, color) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  sheet.setTabColor(color);
  return { success: true, message: 'Tab color set for: ' + sheetName };
}

// ============================================================================
// HEADER & DATA UTILITIES
// ============================================================================

/**
 * Update headers only (preserves all styling and data)
 * @param {string} sheetName - Name of the sheet
 * @param {Array} newHeaders - New header names
 * @returns {Object} Result with success status and message
 */
function updateHeadersOnly(sheetName, newHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  const headerRange = sheet.getRange(1, 1, 1, newHeaders.length);
  headerRange.setValues([newHeaders]);

  return { success: true, message: 'Headers updated for: ' + sheetName };
}

/**
 * Get current headers from a sheet
 * @param {string} sheetName - Name of the sheet
 * @returns {Array|null} Array of header names or null
 */
function getSheetHeaders(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) return null;

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];

  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

/**
 * Add a new column to a sheet
 * @param {string} sheetName - Name of the sheet
 * @param {string} headerName - Name for the new column header
 * @param {number} position - Column position (1-indexed), or null for end
 * @returns {Object} Result with success status and message
 */
function addColumn(sheetName, headerName, position) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  const lastCol = sheet.getLastColumn();
  const insertPos = position || lastCol + 1;

  if (insertPos <= lastCol) {
    sheet.insertColumnBefore(insertPos);
  }

  sheet.getRange(1, insertPos).setValue(headerName);

  return { success: true, message: 'Column "' + headerName + '" added at position ' + insertPos };
}

// ============================================================================
// CUSTOM STYLING
// ============================================================================

/**
 * Apply enterprise styling to any sheet (custom config)
 * @param {string} sheetName - Name of the sheet
 * @param {Object} customConfig - Custom configuration object
 * @returns {Object} Result with success status and message
 */
function applyCustomStyling(sheetName, customConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  const config = {
    displayName: customConfig.displayName || sheetName,
    headers: customConfig.headers || [],
    theme: customConfig.theme || 'neutral',
    frozenRows: customConfig.frozenRows !== undefined ? customConfig.frozenRows : 1,
    frozenCols: customConfig.frozenCols || 0,
    columnWidths: customConfig.columnWidths || {},
    filters: customConfig.filters !== false,
    conditionalFormatting: customConfig.conditionalFormatting || false
  };

  applySheetFormatting_(sheet, config);

  return { success: true, message: 'Custom styling applied to: ' + sheetName };
}

/**
 * Apply styling to sheet without changing headers (preserves existing headers)
 * @param {string} sheetName - Name of the sheet
 * @param {string} themeName - Theme to apply
 * @returns {Object} Result with success status and message
 */
function applyThemeOnly(sheetName, themeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  const theme = getThemeColors_(themeName);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  // Style header row
  styleHeaderRow_(sheet, lastCol, theme);

  // Style data rows
  if (lastRow > 1) {
    styleDataRows_(sheet, lastRow, lastCol, theme);
  }

  // Apply borders
  applyBorders_(sheet, lastRow, lastCol, theme);

  return { success: true, message: 'Theme "' + themeName + '" applied to: ' + sheetName };
}

/**
 * Clear all formatting from a sheet (preserves data)
 * @param {string} sheetName - Name of the sheet
 * @returns {Object} Result with success status and message
 */
function clearSheetFormatting(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  const range = sheet.getRange(1, 1, lastRow, lastCol);

  // Clear formatting only
  range.clearFormat();

  // Remove filters
  const filter = sheet.getFilter();
  if (filter) filter.remove();

  // Clear conditional formatting
  sheet.clearConditionalFormatRules();

  // Reset frozen rows/cols
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  return { success: true, message: 'Formatting cleared for: ' + sheetName };
}

// ============================================================================
// BATCH OPERATIONS
// ============================================================================

/**
 * Apply the same theme to multiple sheets
 * @param {Array} sheetNames - Array of sheet names
 * @param {string} themeName - Theme to apply
 * @returns {Object} Results summary
 */
function applyThemeToMultipleSheets(sheetNames, themeName) {
  const results = {
    success: [],
    failed: []
  };

  sheetNames.forEach(name => {
    const result = applyThemeOnly(name, themeName);
    if (result.success) {
      results.success.push(name);
    } else {
      results.failed.push({ name, error: result.message });
    }
  });

  return {
    success: results.failed.length === 0,
    message: `Applied theme to ${results.success.length} sheets. ${results.failed.length} failed.`,
    details: results
  };
}

/**
 * Auto-resize all columns in a sheet to fit content
 * @param {string} sheetName - Name of the sheet
 * @returns {Object} Result with success status and message
 */
function autoResizeAllColumns(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    sheet.autoResizeColumns(1, lastCol);
  }

  return { success: true, message: 'Columns auto-resized for: ' + sheetName };
}
