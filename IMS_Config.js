// ============================================================================
// CONFIG.GS - Shared Configuration and Universal Utility Functions
// ============================================================================
const SHEET_IDS = {
  // The "Database" (Inventory, Master Logs)
  IMS_SOURCE: '1YGde5-R06qFcY5KGcmOP0-0Fbre-HVwKKZeljFFkgnI',

  // The "Interface" (New Sheet for Supervisors)
  ORDER_MANAGEMENT: '1ab6TvvDgh7SyiZDUBb96CUdsk5u0nZ7saR33Q0m95WE'
};
const FOLDERS = {
  IMS_ROOT: '1Fvr0WsyAfejmTiVRgYoA7zeLiwwnhxJ7',
  INBOUND_UPLOADS: '1pj5rVUwP7drH_vpigdJRePdCu9mj9jYF',
  CUSTOMER_ORDERS: '1G97y64fxlq6rBd8RItHREmNxRMRrK-VV',
  SOURCE_NEW_ORDERS: '1L3mjeQizzjVU5uTqGxv1sOUOuq25I2pM',
  PROCESSED_ORDERS: '1s28aomu1Th2_yNZOCkTyHdLF3Cq2dNep',
  IMS_TEMPLATES: '1cQatcc-vJLgx89_XLWoQ7XlIL_k6by33',
  PDF_TEMPLATES: '1Pi6u2Nt-WI5m6UAi5k9x-AM3srJ-GAGb',
  LABEL_TEMPLATES: '1DcfLwokIy2S9ldMMrNmWBdBSoKJ0XxgG',
  TEMP_WORK: '1s3hlgrOQ1KR4kgGTnOOgsOlLaj08ygD6',
  SOURCE_RECEIPTS: 'ID_FOR_SOURCE_RECEIPTS_FOLDER',
  PROCESSED_RECEIPTS: 'ID_FOR_PROCESSED_RECEIPTS_FOLDER',
  IMS_Reports: '1NxJRIcraNfpGk8ByvauVBFWkBtwp4y9D'
};
var REPORT_TEMPLATES = {
  INBOUND: '1wNz7SXXL-5Df15BBMg7wAE9imumgukLi3fnPYJ54Byk',
  OUTBOUND: '1yN0k4mYC4Cvr-Eon9WQDNccWbgmzwetRm3n6jtbBvgc'
};
const TEMPLATES = {
  TOC_Template: '1S6W9u_iLNrKzVL51rXLpXRY7d7i3RXRXp6v2FdFdmC4',
  Packing_List_Template: '1Z-InQHf8dUxQzsvF7dmbtwZ1SotmWabU2FY64ZNi-Zs',
  Backorder_Template: '1n8ZkctRKibi98yv-n8-ybWOPD1djVpVWiAICoT-dBTM',
  Pickticket_Template: '1J75-dFbotHgn9R5xLZbMMpSeeDM1mvz5SVsB_mNC0lY', // UPDATED ID
  Inbound_Label: '1a_XK-UdV-dPVnlDvHg7kK5zNuVCRyuev',
  Masterskid_Label: '1N5EXHfrvIT1-LyghWzthT_wsvteaFKAa'
};
const TABS = {
  CUSTOMER_ORDERS: "Customer_Orders",
  STOCK_TOTALS: "Stock_Totals",
  MASTER_LOG: "Master_Log",
  OUTBOUNDLOG: "OutboundLog",
  INBOUND_SKIDS: "Inbound_Skids",
  BACKORDERS: "Backorders",
  VERIFICATION_LOG: "Verification_Log",
  INBOUND_STAGING: "Inbound_Staging",
  TXN_RECON_20260107_082711: "TXN_Recon_20260107_082711",
  SHEET56: "Sheet56",
  SHEET53: "Sheet53",
  INBOUND_QTY_AUDIT: "Inbound_Qty_Audit",
  BIN_STOCK: "Bin_Stock",
  FLOOR_STOCK_LEVELS: "Floor_Stock_Levels",
  SHEET69: "Sheet69",
  CYCLE_COUNT: "Cycle_Count",
  REQUESTED_ITEMS: "Requested_Items",
  PICK_LOG: "Pick_Log",
  SHEET71: "Sheet71",
  CONSOLIDATIONSUGGESTIONS: "ConsolidationSuggestions",
  PICKED_ITEM_LOG: "Picked_Item_Log",
  LOCATIONLOG: "LocationLog",
  CYCLE_COUNT_BATCHES: "Cycle_Count_Batches",
  CYCLE_COUNT_LINES: "Cycle_Count_Lines",
  REPLENISHMENT_TASKS: "Replenishment_Tasks",
  DOCK_SCHEDULE: "Dock_Schedule",
  RACKING_AUDIT: "Racking_Audit",
  BREAKDOWN_LOG: "Breakdown_Log",
  NVSCRIPTSPROPERTIES: "NVScriptsProperties",
  BACKORDERFULFILLMENT_LOG: "BackorderFulfillment_Log",
  TRUCK_SCHEDULE: "Truck_Schedule",
  ALLOCATION_LOG: "Allocation_Log",
  PROJECT_MASTER: "Project_Master",
  SUPPORT_SHEET: "Support_Sheet",
  AUDIT_TRAIL: "Audit_Trail",
  ITEM_MASTER: "Item_Master",
  PO_MASTER: "PO_Master",
  CUSTOMER_ACCESS: "Customer_Access",
  MANUFACTURER_MASTER: "Manufacturer_Master"
};

const HEADERS = {
  "Customer_Orders": [
    "Order_ID",
    "Task_Number",
    "Project",
    "NBD",
    "Request_Status",
    "Stock_Status",
    "Company",
    "Order_Title",
    "Deliver_To",
    "Name",
    "Phone_Number",
    "Original_Order",
    "Order_Folder",
    "Pick_Ticket_PDF",
    "TOC_PDF",
    "Packing_Lists",
    "Created_TS",
    "Created_By"
  ],
  "Stock_Totals": [
    "SKU",
    "Asset_Type",
    "Manufacturer",
    "MFPN",
    "FBPN",
    "UOM",
    "Qty_Available",
    "Qty_In_Racking",
    "Qty_On_Floor",
    "Qty_Inbound_Staging",
    "Qty_Allocated",
    "Qty_Backordered",
    "Qty_Shipped",
    "Qty_Received",
    "Qty_In_Stock"
  ],
  "Master_Log": [
    "Txn_ID",
    "Date_Received",
    "Inbound_Files",
    "Transaction_Type",
    "Warehouse",
    "Project",
    "Push #",
    "FBPN",
    "Qty_Received",
    "UOM",
    "Total_Skid_Count",
    "Carrier",
    "BOL_Number",
    "Customer_PO_Number",
    "Manufacturer",
    "MFPN",
    "Description",
    "Received_By",
    "SKU"
  ],
  "OutboundLog": [
    "Date",
    "Order_Number",
    "Task_Number",
    "Transaction Type",
    "Warehouse",
    "Company",
    "Project",
    "FBPN",
    "Manufacturer",
    "Qty",
    "UOM",
    "Skid_ID",
    "SKU"
  ],
  "Inbound_Skids": [
    "Skid_ID",
    "TXN_ID",
    "Date",
    "Asset_Type",
    "FBPN",
    "MFPN",
    "Project",
    "Qty_on_Skid",
    "UOM",
    "Skid_Sequence",
    "Is_Mixed",
    "Timestamp",
    "SKU"
  ],
  "Backorders": [
    "Order_ID",
    "NBD",
    "Status",
    "Task_Number",
    "Stock_Status",
    "Asset_Type",
    "FBPN",
    "UOM",
    "Qty_Requested",
    "Qty_Backordered",
    "Qty_Fulfilled",
    "Date_Logged",
    "Date_Closed",
    "Notes",
    "Backorder_ID",
    "SKU"
  ],
  "Verification_Log": [
    "Timestamp",
    "BOL_Number",
    "PO_Number",
    "Asset_Type",
    "Manufacturer",
    "MFPN",
    "FBPN",
    "UOM",
    "Expected_Qty",
    "Actual_Qty",
    "Variance",
    "Box_Labels",
    "Verified_By",
    "Skid_ID",
    "TXN_ID"
  ],
  "Inbound_Staging": [
    "Bin_Code",
    "Bin_Name",
    "Push_Number",
    "Project",
    "Manufacturer",
    "FBPN",
    "UOM",
    "Initial_Quantity",
    "Current_Quantity",
    "Stock_Percentage",
    "AUDIT NEEDED",
    "Skid_ID",
    "SKU",
    "Last_Updated"
  ],
  "TXN_Recon_20260107_082711": [
    "TXN_ID Reconciliation: Master_Log vs Inbound_Skids"
  ],
  "Sheet56": [
    "Push #"
  ],
  "Sheet53": [],
  "Inbound_Qty_Audit": [
    "Txn_ID",
    "FBPN",
    "Master_Log Qty_Received",
    "Inbound_Skids Qty_on_Skid",
    "Difference (Master - Skids)",
    "Issue",
    "Push #",
    "BOL_Number",
    "Customer_PO_Number",
    "Project (from Skids)",
    "Skid_IDs"
  ],
  "Bin_Stock": [
    "Bin_Code",
    "Bin_Name",
    "Push_Number",
    "Project",
    "Manufacturer",
    "FBPN",
    "UOM",
    "Initial_Quantity",
    "Current_Quantity",
    "Stock_Percentage",
    "AUDIT NEEDED",
    "Skid_ID",
    "SKU"
  ],
  "Floor_Stock_Levels": [
    "Bin_Code",
    "Bin_Name",
    "Push_Number",
    "Project",
    "Manufacturer",
    "FBPN",
    "UOM",
    "Initial_Quantity",
    "Current_Quantity",
    "Stock_Percentage",
    "AUDIT NEEDED",
    "Skid_ID",
    "SKU"
  ],
  "Sheet69": [
    "Bin_Code",
    "Bin_Name",
    "Push_Number",
    "FBPN",
    "Manufacturer",
    "Project",
    "Initial_Quantity",
    "Current_Quantity",
    "Stock_Percentage",
    "AUDIT NEEDED",
    "Skid_ID",
    "SKU"
  ],
  "Cycle_Count": [
    "Batch_ID",
    "Status",
    "Created_At",
    "Created_By",
    "Bin_Code",
    "FBPN",
    "Manufacturer",
    "Project",
    "Current_Qty",
    "Counted_Qty",
    "Variance",
    "Notes",
    "Counted_At",
    "Counted_By"
  ],
  "Requested_Items": [
    "Order_ID",
    "Asset_Type",
    "Manufacturer",
    "FBPN",
    "Description",
    "UOM",
    "Qty_Requested",
    "Stock_Status",
    "Qty_Backordered",
    "Qty_Allocated",
    "Qty_Shipped",
    "Backorder_ID",
    "Allocation_ID",
    "SKU"
  ],
  "Pick_Log": [
    "PIK_ID",
    "NBD",
    "Order_Number",
    "Task_Number",
    "Company",
    "Project",
    "Asset_Type",
    "Manufacturer",
    "FBPN",
    "Description",
    "UOM",
    "Qty_Requested",
    "Qty_To_Pick",
    "Bin_Code",
    "Qty_Picked",
    "Status",
    "Picked_By",
    "Shipped_Date",
    "Timestamp",
    "SKU"
  ],
  "Sheet71": [],
  "ConsolidationSuggestions": [
    "SKU",
    "FBPN",
    "Bin_To_Empty",
    "Qty_In_Donor_Bin",
    "Bin_To_Fill",
    "Original_Qty",
    "Qty_To_Move",
    "Final_Qty_In_Bin"
  ],
  "Picked_Item_Log": [],
  "LocationLog": [
    "Timestamp",
    "Action",
    "FBPN",
    "Manufacturer",
    "Bin_Code",
    "Qty_Changed",
    "Resulting_Qty",
    "Description",
    "User_Email",
    "SKU"
  ],
  "Cycle_Count_Batches": [
    "Batch_ID",
    "Zone",
    "Created_TS",
    "Created_By",
    "Status",
    "Notes"
  ],
  "Cycle_Count_Lines": [
    "Batch_ID",
    "Bin_Code",
    "FBPN",
    "SKU",
    "Expected_Qty",
    "Counted_Qty",
    "Variance",
    "Status"
  ],
  "Replenishment_Tasks": [
    "Task_ID",
    "Created_TS",
    "Priority",
    "Source_Bin",
    "Target_Bin",
    "FBPN",
    "SKU",
    "Qty",
    "Status",
    "Assigned_To",
    "Completed_TS",
    "Notes"
  ],
  "Dock_Schedule": [
    "Dock_ID",
    "Type",
    "Date",
    "Time",
    "Carrier",
    "BOL",
    "Order_Number",
    "Task_Number",
    "Status",
    "Notes"
  ],
  "Racking_Audit": [
    "FBPN",
    "Manufacturer",
    "Bin_Stock_Total",
    "Stock_Totals_Qty_In_Racking",
    "Difference"
  ],
  "Breakdown_Log": [
    "Breakdown_ID",
    "Skid_ID",
    "Item_ID",
    "FBPN",
    "Qty_Before",
    "Qty_Removed",
    "Qty_Remaining",
    "MasterSkid_ID",
    "Qty_On_ChildSkid",
    "Bin_ID_Child",
    "Date_Broken_Down",
    "Broken_Down_By",
    "Notes"
  ],
  "NVScriptsProperties": [
    "autocratn",
    "autocratp"
  ],
  "BackorderFulfillment_Log": [
    "Order_ID",
    "Task_Number",
    "FBPN",
    "Qty_Fulfilled",
    "Fulfillment_Date",
    "Status_After",
    "Fulfilled_By",
    "Notes",
    "Backorder_ID",
    "Txn_ID",
    "Fulfillment_ID",
    "Timestamp",
    "SKU"
  ],
  "Truck_Schedule": [
    "Schedule_ID",
    "Delivery_Date",
    "Time",
    "Carrier",
    "BOL_Number",
    "PO",
    "Project",
    "Total_Skids",
    "FBPN",
    "Total_Qty",
    "Notes",
    "BOL_Packing_List"
  ],
  "Allocation_Log": [
    "Order_ID",
    "Timestamp",
    "Allocation_Status",
    "Asset_Type",
    "Manufacturer",
    "FBPN",
    "UOM",
    "Qty_Requested",
    "Qty_Allocated",
    "Qty_Backordered",
    "Allocated_By",
    "Backorder_ID",
    "Allocation_ID",
    "SKU"
  ],
  "Project_Master": [
    "Customer_PO",
    "FBPN",
    "MFPN",
    "Description",
    "Manufacturer",
    "Project",
    "SKU",
    "Qty_Ordered",
    "Qty_Received"
  ],
  "Support_Sheet": [
    "Company",
    "Project",
    "Name",
    "Phone_Number"
  ],
  "Audit_Trail": [
    "Audit_ID",
    "Timestamp",
    "User",
    "Action",
    "Entity_Type",
    "Entity_ID",
    "Field",
    "Old_Value",
    "New_Value",
    "Notes"
  ],
  "Item_Master": [
    "FBPN",
    "Manufacturer",
    "MFPN",
    "Asset_Type",
    "Description",
    "UOM",
    "SKU"
  ],
  "PO_Master": [
    "Customer_PO",
    "Project"
  ],
  "Customer_Access": [
    "Email",
    "Name",
    "Company Name",
    "Email Domain",
    "Access_Level",
    "Project_Access",
    "Active"
  ]
};




function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  return sheet;
}

function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}


function appendRow(sheetName, rowData) {
  const sheet = getSheet(sheetName);
  sheet.appendRow(rowData);
}

function appendRows(sheetName, rowsData) {
  if (!rowsData || rowsData.length === 0) return;
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rowsData.length, rowsData[0].length).setValues(rowsData);
}

function generateTxnId() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let id = 'TXN_';
  for (let i = 0; i < 6; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

function generateSkidId() {
  const sheet = getSheet(TABS.INBOUND_SKIDS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return 'SKID_000001';
  }

  const lastSkidId = sheet.getRange(lastRow, 1).getValue();
  const numPart = parseInt(lastSkidId.split('_')[1]) || 0;
  const newNum = numPart + 1;

  return 'SKID_' + String(newNum).padStart(6, '0');
}

function formatDate(date) {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  return `${month}/${day}/${year}`;
}

function formatMonthYear(date) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${months[date.getMonth()]} ${date.getFullYear()}`;
}

function getTimestamp() {
  const now = new Date();
  const month = now.getMonth() + 1;
  const day = now.getDate();
  const year = now.getFullYear();
  const hours = now.getHours();
  const minutes = now.getMinutes();
  const seconds = now.getSeconds();
  return `${month}/${day}/${year} ${hours}:${minutes}:${seconds}`;
}

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

function setCache(key, value, expirationInSeconds = 600) {
  const cache = CacheService.getScriptCache();
  const stringValue = JSON.stringify(value);
  cache.put(key, stringValue, expirationInSeconds);
}


function getCache(key) {
  const cache = CacheService.getScriptCache();
  const value = cache.get(key);
  return value ? JSON.parse(value) : null;
}


function removeCache(key) {
  const cache = CacheService.getScriptCache();
  cache.remove(key);
}

function clearAllCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(cache.getKeys());
}

function lookupProjectMaster(fbpn, customerPO, manufacturer) {
  const cacheKey = `PM_${fbpn}_${customerPO}_${manufacturer}`;
  const cached = getCache(cacheKey);
  if (cached) return cached;

  const sheet = getSheet(TABS.PROJECT_MASTER);
  const data = sheet.getDataRange().getValues();
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowPO = row[0];
    const rowFBPN = row[1];
    const rowManufacturer = row[4];
    if (rowFBPN === fbpn && rowPO === customerPO && rowManufacturer === manufacturer) {
      const result = {
        mfpn: row[2] ||
          '',
        description: row[3] ||
          '',
        project: row[5] ||
          '',
        sku: row[6] || ''
      };
      setCache(cacheKey, result, 1800); // Cache for 30 minutes
      return result;
    }
  }

  // Not found
  const emptyResult = { mfpn: '', description: '', project: '', sku: '' };
  setCache(cacheKey, emptyResult, 1800);
  return emptyResult;
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}


function createInboundFolder(date, bolNumber) {
  const rootFolder = DriveApp.getFolderById(FOLDERS.INBOUND_UPLOADS);

  // Create or get month folder (e.g., "Jan 2024")
  const monthFolderName = formatMonthYear(date);
  const monthFolder = getOrCreateFolder(rootFolder, monthFolderName);

  // Create or get day folder (e.g., "15")
  const dayFolderName = String(date.getDate()).padStart(2, '0');
  const dayFolder = getOrCreateFolder(monthFolder, dayFolderName);

  // Create or get BOL folder
  const bolFolder = getOrCreateFolder(dayFolder, bolNumber);

  return bolFolder;
}


function uploadFileToFolder(folder, fileBlob, fileName) {
  return folder.createFile(fileBlob.setName(fileName));
}

function getNextStagingLocation(stagingArea) {
  const areaPrefix = stagingArea.replace('Inbound Staging ', 'IS');
  const sheet = getSheet(TABS.INBOUND_STAGING);
  const data = sheet.getDataRange().getValues();

  // Find all bins in the specified area that are empty
  for (let i = 1; i < data.length; i++) {
    const binCode = data[i][0];
    const binName = data[i][1];
    const fbpn = data[i][2];

    // Check if this bin is in the correct area and is empty
    if (binCode.startsWith(areaPrefix) && !fbpn) {
      return { binCode: binCode, binName: binName, rowIndex: i + 1 };
    }
  }

  return null; // No available locations
}


function allocateStagingLocations(stagingArea, numberOfSkids) {
  const locations = [];
  const areaPrefix = stagingArea.replace('Inbound Staging ', 'IS');
  const sheet = getSheet(TABS.INBOUND_STAGING);
  const data = sheet.getDataRange().getValues();

  let allocated = 0;
  for (let i = 1; i < data.length && allocated < numberOfSkids; i++) {
    const binCode = data[i][0];
    const binName = data[i][1];
    const fbpn = data[i][2];

    if (binCode.startsWith(areaPrefix) && !fbpn) {
      locations.push({ binCode: binCode, binName: binName, rowIndex: i + 1 });
      allocated++;
    }
  }

  if (allocated < numberOfSkids) {
    throw new Error(`Only ${allocated} of ${numberOfSkids} locations available in ${stagingArea}`);
  }

  return locations;
}
const REPORT_SETTINGS = {
  // Date ranges for each period type
  PERIODS: {
    DAILY: 1,    // Last 1 day
    WEEKLY: 7,   // Last 7 days
    MONTHLY: 30  // Last 30 days
  },

  // Filename format
  DATE_FORMAT: 'MMddyy',           // e.g., 111325
  MONTH_FORMAT: 'MMMyy',           // e.g., Nov25

  // File naming pattern: {Period}_{Type}_{Date}.pdf
  // Examples: Daily_Inbound_111325.pdf, Weekly_Outbound_111325.pdf
};

// ═══════════════════════════════════════════════════════════════════════════
// UI SETTINGS
// ═══════════════════════════════════════════════════════════════════════════

const UI_CONFIG = {
  MODAL_WIDTH: 600,
  MODAL_HEIGHT: 400,
  THEME: {
    PRIMARY_COLOR: '#1976D2',
    SUCCESS_COLOR: '#4CAF50',
    ERROR_COLOR: '#F44336',
    BACKGROUND: '#F5F5F5'
  }
};
// ═══════════════════════════════════════════════════════════════════════════
// SYSTEM SETTINGS
// ═══════════════════════════════════════════════════════════════════════════

const SYSTEM_CONFIG = {
  TIMEZONE: Session.getScriptTimeZone(),
  MAX_RETRIES: 3,           // Number of retry attempts for Drive operations
  RETRY_DELAY: 2000,        // Milliseconds between retries
  LOG_LEVEL: 'INFO'         // INFO, WARNING, ERROR
};
// ═══════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Get column index by name
 */
function getColumnIndex(headers, columnName) {
  const index = headers.indexOf(columnName);
  if (index === -1) {
    Logger.log(`Warning: Column "${columnName}" not found`);
  }
  return index;
}
var SHEET_NAMES = {
  MASTER_LOG: 'Master_Log',           // Inbound transactions
  OUTBOUND_LOG: 'OutboundLog',        // Outbound transactions
  BACKORDERS: 'Backorders',           // Backorder tracking
  CUSTOMER_ORDERS: 'Customer_Orders', // Order data
  BIN_STOCK: 'Bin_Stock',            // Bin-level inventory
  STOCK_TOTALS: 'Stock_Totals'       // Aggregated stock levels
};
// ═══════════════════════════════════════════════════════════════════════════
// COLUMN MAPPINGS
// ═══════════════════════════════════════════════════════════════════════════

// Master_Log columns (Inbound)
var MASTER_LOG_COLUMNS = {
  TXN_ID: 'Txn_ID',
  DATE_RECEIVED: 'Date_Received',
  TRANSACTION_TYPE: 'Transaction_Type',
  WAREHOUSE: 'Warehouse',
  PUSH_NUM: 'Push #',
  FBPN: 'FBPN',
  QTY_RECEIVED: 'Qty_Received',
  TOTAL_SKID_COUNT: 'Total_Skid_Count',
  CARRIER: 'Carrier',
  BOL_NUMBER: 'BOL_Number',
  CUSTOMER_PO_NUMBER: 'Customer_PO_Number',
  MANUFACTURER: 'Manufacturer',
  MFPN: 'MFPN',
  DESCRIPTION: 'Description',
  RECEIVED_BY: 'Received_By',
  SKU: 'SKU'
};
// OutboundLog columns
var OUTBOUND_LOG_COLUMNS = {
  DATE: 'Date',
  ORDER_NUMBER: 'Order_Number',
  TASK_NUMBER: 'Task_Number',
  TRANSACTION_TYPE: 'Transaction Type',
  WAREHOUSE: 'Warehouse',
  COMPANY: 'Company',
  PROJECT: 'Project',
  FBPN: 'FBPN',
  MANUFACTURER: 'Manufacturer',
  QTY: 'Qty',
  TOC: 'TOC',
  PO_NUMBER: 'PO_Number',
  SKID_ID: 'Skid_ID',
  SKU: 'SKU'
};
/**
 * Validate configuration
 */
function validateConfig() {
  const errors = [];
  // Check template IDs
  if (REPORT_TEMPLATES.INBOUND === '1ABC...') {
    errors.push('Inbound template ID not set');
  }
  if (REPORT_TEMPLATES.OUTBOUND === '1DEF...') {
    errors.push('Outbound template ID not set');
  }

  // Check sheet names exist
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEET_NAMES).forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      errors.push(`Sheet "${sheetName}" not found`);
    }
  });
  if (errors.length > 0) {
    return {
      valid: false,
      errors: errors
    };
  }

  return {
    valid: true,
    message: 'Configuration is valid'
  };
}
const DRIVE_CONFIG = {
  REPORTS_FOLDER: 'IMS Reports',      // Main folder for all reports
  CREATE_SUBFOLDERS: true,            // Create Inbound/Outbound subfolders
  SUBFOLDER_NAMES: {
    INBOUND: 'Inbound Reports',
    OUTBOUND: 'Outbound Reports'
  }
};
/**
 * Show configuration status
 */
function showConfigStatus() {
  const validation = validateConfig();
  if (validation.valid) {
    SpreadsheetApp.getUi().alert('✓ Configuration Valid', validation.message, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    const errorMessage = 'Configuration Errors:\n\n' + validation.errors.join('\n');
    SpreadsheetApp.getUi().alert('✗ Configuration Errors', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// SHARED CONFIG ACCESSOR (used by Shipping_Docs / Outbound / other modules)
// ─────────────────────────────────────────────────────────────────────────────
/**
 * Canonical config accessor for the IMS system.
* Shipping_Docs expects these keys:
 * - TOC_TEMPLATE_ID
 * - PACKING_LIST_TEMPLATE_ID
 * - TOC_PACKING_OUTPUT_FOLDER_ID
 * - CUSTOMER_ORDERS_FOLDER_ID
 * - SOURCE_NEW_ORDERS_FOLDER_ID (Added for Automation)
 */
function getIMSConfig() {
  return {
    // Templates
    TOC_TEMPLATE_ID: (typeof TEMPLATES !== 'undefined' && TEMPLATES.TOC_Template) ?
      TEMPLATES.TOC_Template : '',
    PACKING_LIST_TEMPLATE_ID: (typeof TEMPLATES !== 'undefined' && TEMPLATES.Packing_List_Template) ?
      TEMPLATES.Packing_List_Template : '',
    // ADDED PICK TICKET ID EXPLICITLY HERE FOR SAFETY
    PICK_TICKET_TEMPLATE_ID: (typeof TEMPLATES !== 'undefined' && TEMPLATES.Pickticket_Template) ?
      TEMPLATES.Pickticket_Template : '',

    // Folders (default: save PDFs under PROCESSED_ORDERS)
    TOC_PACKING_OUTPUT_FOLDER_ID: (typeof FOLDERS !== 'undefined' && FOLDERS.PROCESSED_ORDERS) ?
      FOLDERS.PROCESSED_ORDERS : ((typeof FOLDERS !== 'undefined' && FOLDERS.IMS_ROOT) ? FOLDERS.IMS_ROOT : ''),

    CUSTOMER_ORDERS_FOLDER_ID: (typeof FOLDERS !== 'undefined' && FOLDERS.CUSTOMER_ORDERS) ?
      FOLDERS.CUSTOMER_ORDERS : ((typeof FOLDERS !== 'undefined' && FOLDERS.IMS_ROOT) ? FOLDERS.IMS_ROOT : ''),

    // Added Source Folder for Auto-Processing
    SOURCE_NEW_ORDERS_FOLDER_ID: (typeof FOLDERS !== 'undefined' && FOLDERS.SOURCE_NEW_ORDERS) ?
      FOLDERS.SOURCE_NEW_ORDERS : ''
  };
}

// ============================================================================
// MAIN CONTROLLER / SHELL FUNCTIONS
// ============================================================================

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  const customerOrdersMenu = ui.createMenu('Customer Orders')
    .addItem('Order Entry Portal', 'showOrderEntryPortal')
    .addItem('Create Customer Order', 'showCustomerOrderModal')
    .addItem('Generate Pick Ticket', 'openPickTicketGenerator')
    .addItem('Process Outbound', 'openPackingTOCGenerator')
    
  const inboundMenu = ui.createMenu('Inbound Receiving')
    .addItem('Inbound Delivery Form', 'openInboundModal')
    .addItem('Inbound Count Verification', 'openInboundVerificationModal')
    .addItem('Create Skid Label', 'openInboundSkidLabelModal')
    .addItem('Reprint Past Delivery Labels', 'openPastDeliverySkidLabelModal')
    .addItem('Batch Label Generator', 'openBatchLabelGeneratorModal')

  const inventoryToolsMenu = ui.createMenu('Inventory Tools')
    .addSeparator()
    .addItem('Create Skid Label', 'openInboundSkidLabelModal')
    .addItem('Add New Item (Item_Master)', 'openAddItemModal')
    .addItem('Add New PO / Project (PO_Master)', 'openAddPOModal')
    .addSeparator()
    .addItem('Bin Stock Put-Away', 'openStockToolsModal')
    .addItem('Cycle Count', 'openCycleCountModal')
    .addItem('Create Empty Bin Location PDF', 'createZeroStockPutAwayLog');
    

  ui.createMenu('IMS')
    .addSubMenu(inboundMenu)
    .addSeparator()
    .addSubMenu(customerOrdersMenu)
    .addSeparator()
    .addSubMenu(inventoryToolsMenu)
    .addSeparator()
    .addItem('Generate Reports', 'showReportGeneratorModal')
    .addSeparator()
    .addItem('🔧 Debug: Master_Log Structure', 'showDiagnostics')
    .addItem('🔧 Test: Batch Label Function', 'test_getBolsForLabelGeneration')
    .addToUi();
}

function onEdit(e) {
  // onEdit trigger - add custom logic here if needed
}

function doGet(e) {
  const context = getUserContext();

  const template = HtmlService.createTemplateFromFile('IMSWebApp');
  template.userContext = JSON.stringify(context);

  // Deep link param: ?skid=SKID_000123
  const skidParam = (e && e.parameter && e.parameter.skid) ? String(e.parameter.skid).trim() : '';
  template.deepLinkSkidId = JSON.stringify(skidParam); // keep JSON-safe

  const html = template.evaluate()
    .setTitle('IMS - Warehouse Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  return html;
}

// ----------------------------------------------------------------------------
// MODAL OPENERS
// ----------------------------------------------------------------------------
function openInboundManagerModal() {
  const html = HtmlService.createTemplateFromFile('InboundManagerModal')
    .evaluate()
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Inbound Manager');
}

function openInboundVerificationModal() {
  const html = HtmlService.createTemplateFromFile('InboundVerificationModal')
    .evaluate()
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Inbound Count Verification');
}

function openInboundModal() {
  const html = HtmlService.createTemplateFromFile('InboundModal')
    .evaluate().setWidth(950).setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Inbound Form');
}

function showCustomerOrderModal() {
  const html = HtmlService.createTemplateFromFile('CustomerOrderModal')
    .evaluate().setWidth(900).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Customer Order');
}

function openCancelOrderModal() {
  const html = HtmlService.createTemplateFromFile('CancelOrderModal')
    .evaluate().setWidth(600).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cancel Customer Order');
}

function openPickTicketGenerator() {
  const html = HtmlService.createTemplateFromFile('PickTicketModal')
    .evaluate().setWidth(950).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Pick Ticket');
}

function openPackingTOCGenerator() {
  const html = HtmlService.createTemplateFromFile('PackingTOCModal')
    .evaluate().setWidth(1050).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Customer Order Outbound');
}

function showReportGeneratorModal() {
  const html = HtmlService.createTemplateFromFile('ReportGeneratorModal')
    .evaluate()
    .setWidth(950)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Reports');
}

function openStockToolsModal() {
  const html = HtmlService.createTemplateFromFile('StockToolsModal')
    .evaluate().setWidth(1050).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bin Stock Put-Away');
}

function openBinLookupModal() {
  const html = HtmlService.createTemplateFromFile('BinLookupModal')
    .evaluate().setWidth(1000).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bin & Item Lookup');
}

function openBinUpdateModal() {
  const html = HtmlService.createTemplateFromFile('BinUpdateModal')
    .evaluate().setWidth(900).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bin Update');
}

function openCycleCountModal() {
  const html = HtmlService.createTemplateFromFile('CycleCountModal')
    .evaluate().setWidth(1100).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cycle Count');
}

function openCurrentItemsModal() {
  const html = HtmlService.createTemplateFromFile('CurrentItemsModal')
    .evaluate().setWidth(1100).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Current Items');
}

function showCustomerPortal() {
  const html = HtmlService.createTemplateFromFile('CustomerPortalUI')
    .evaluate().setTitle('Inventory Portal');
  SpreadsheetApp.getUi().showSidebar(html);
}

function openGeneratePastInboundLabelsByDateModal() {
  const html = HtmlService.createTemplateFromFile('LabelDatePickerModal')
    .evaluate()
    .setWidth(520)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Past Inbound Labels');
}

// Manual Skid Label Modal Functions
function openInboundSkidLabelModal() {
  const html = HtmlService.createTemplateFromFile('InboundSkidLabelModal')
    .evaluate()
    .setWidth(850)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Inbound Skid Label');
}

// Data Entry Modal Functions
function openAddItemModal() {
  const html = HtmlService.createTemplateFromFile('AddItemModal')
    .evaluate()
    .setWidth(500)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Item');
}

function openAddPOModal() {
  const html = HtmlService.createTemplateFromFile('AddPOModal')
    .evaluate()
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Customer PO & Project');
}

function openBatchLabelGeneratorModal() {
  const html = HtmlService.createTemplateFromFile('BatchLabelGenerator')
    .evaluate()
    .setWidth(1100)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Batch Label Generator');
}

function openPastDeliverySkidLabelModal() {
  const html = HtmlService.createTemplateFromFile('PastDeliverySkidLabelModal')
    .evaluate()
    .setWidth(860)
    .setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, 'Past Delivery Labels');
}

// ----------------------------------------------------------------------------
// SHELL WRAPPER FUNCTIONS
// Note: Most functions are implemented in their respective modules:
// - IMS_Inbound.js: Inbound processing, labels, manufacturers, FBPNs
// - CustomerOrderBackend.js: Orders, companies, projects
// - SHIPPING_DOCS.js: Pick tickets, TOC, packing lists
// - BinLookup.js: Bin search, details, history
// - BinUpdate.js: Inventory batch operations
// - CycleCount.js: Cycle count functions
// - IMS_Inbound_Manager.js: Inbound undo, search, label regeneration
// - ReportGenerator.js: Report generation
// ----------------------------------------------------------------------------

function shell_generateLabelsForAllPastInbounds(startDate, endDate) {
  if (!startDate) throw new Error('startDate is required');
  return generateLabelsForAllPastInbounds(startDate, endDate || startDate);
}

// ----------------------------------------------------------------------------
// BATCH LABEL GENERATION FUNCTIONS
// ----------------------------------------------------------------------------

/**
 * findHeaderRow_
 * Scans the first N rows for a row containing all required column names.
 * Useful when sheets have title rows or spacer rows above the actual header.
 */
function findHeaderRow_(data, requiredNames, maxRows = 10) {
  if (!data || data.length === 0) return -1;
  const limit = Math.min(maxRows, data.length);
  for (let r = 0; r < limit; r++) {
    const row = data[r] || [];
    const set = new Set(row.map(v => String(v || '').trim()));
    let allFound = true;
    for (const name of requiredNames) {
      if (!set.has(name)) {
        allFound = false;
        break;
      }
    }
    if (allFound) return r;
  }
  return -1;
}

/**
 * Gets all BOLs from Master_Log for the batch generator modal.
 */
function shell_getBolsForLabelGeneration() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName('Master_Log');
    var skidsSheet = ss.getSheetByName('Inbound_Skids');

    if (!masterSheet) {
      return { success: false, message: 'Master_Log sheet not found.' };
    }

    var mData = masterSheet.getDataRange().getValues();
    var sData = skidsSheet ? skidsSheet.getDataRange().getValues() : [];

    if (!mData || mData.length === 0) {
      return { success: true, bols: [] };
    }

    var mHead = mData[0];
    var mTxn = mHead.indexOf('Txn_ID');
    var mDate = mHead.indexOf('Date_Received');
    var mBol = mHead.indexOf('BOL_Number');
    var mFbpn = mHead.indexOf('FBPN');
    var mMan = mHead.indexOf('Manufacturer');
    var mProj = mHead.indexOf('Project');

    if (mTxn < 0 || mBol < 0) {
      return { success: false, message: 'Master_Log missing required columns (Txn_ID, BOL_Number).' };
    }

    // Count skids per TXN_ID
    var skidCounts = {};
    if (sData.length > 1) {
      var sHead = sData[0];
      var sTxn = sHead.indexOf('TXN_ID');
      if (sTxn >= 0) {
        for (var j = 1; j < sData.length; j++) {
          var tid = String(sData[j][sTxn] || '').trim();
          if (tid) {
            skidCounts[tid] = (skidCounts[tid] || 0) + 1;
          }
        }
      }
    }

    // Build unique BOL list
    var bolMap = {};
    for (var i = 1; i < mData.length; i++) {
      var txnId = String(mData[i][mTxn] || '').trim();
      var bol = String(mData[i][mBol] || '').trim();

      if (!txnId || !bol) continue;

      var key = txnId + '|' + bol;
      if (bolMap[key]) continue;

      var dateVal = mData[i][mDate];
      var dateStr = '';
      if (dateVal) {
        try {
          var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
          if (!isNaN(d.getTime())) {
            dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
          }
        } catch (e) {
          dateStr = String(dateVal);
        }
      }

      bolMap[key] = {
        key: key,
        txnId: txnId,
        bol: bol,
        dateStr: dateStr,
        dateVal: dateVal,
        fbpn: mFbpn >= 0 ? String(mData[i][mFbpn] || '') : '',
        manufacturer: mMan >= 0 ? String(mData[i][mMan] || '') : '',
        project: mProj >= 0 ? String(mData[i][mProj] || '') : '',
        skidCount: skidCounts[txnId] || 0,
        hasLabels: false
      };
    }

    // Convert to array and sort
    var bols = [];
    for (var k in bolMap) {
      if (bolMap.hasOwnProperty(k)) {
        bols.push(bolMap[k]);
      }
    }

    bols.sort(function (a, b) {
      var da = a.dateVal ? new Date(a.dateVal) : new Date(0);
      var db = b.dateVal ? new Date(b.dateVal) : new Date(0);
      return db - da;
    });

    // Limit to 250
    if (bols.length > 250) {
      bols = bols.slice(0, 250);
    }

    return { success: true, bols: bols };

  } catch (err) {
    Logger.log('shell_getBolsForLabelGeneration error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.message };
  }
}

/**
 * Cache key for label existence checks.
 * Uses the same folder path rules as checkLabelsExist_.
 */
function buildLabelCacheKey_(dateVal, bol) {
  if (!dateVal || !bol) return '';
  try {
    const d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return '';
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthYear = months[d.getMonth()] + ' ' + d.getFullYear();
    const day = String(d.getDate()).padStart(2, '0');
    const safeBol = String(bol).trim().replace(/[\/\\?%*:|"<>\.]/g, '_');
    return 'lbl|' + monthYear + '|' + day + '|' + safeBol;
  } catch (e) {
    return '';
  }
}

/**
 * Checks if label files exist in the BOL folder.
 */
function checkLabelsExist_(rootFolder, dateVal, bol) {
  if (!dateVal || !bol) return false;

  try {
    const d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return false;

    // Build folder path: {MonthYear}/{Day}/{BOL}
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthYear = months[d.getMonth()] + ' ' + d.getFullYear();
    const day = String(d.getDate()).padStart(2, '0');
    const safeBol = String(bol).trim().replace(/[\/\\?%*:|"<>\.]/g, '_');

    // Navigate to month folder
    const monthFolders = rootFolder.getFoldersByName(monthYear);
    if (!monthFolders.hasNext()) return false;
    const monthFolder = monthFolders.next();

    // Navigate to day folder
    const dayFolders = monthFolder.getFoldersByName(day);
    if (!dayFolders.hasNext()) return false;
    const dayFolder = dayFolders.next();

    // Navigate to BOL folder
    const bolFolders = dayFolder.getFoldersByName(safeBol);
    if (!bolFolders.hasNext()) return false;
    const bolFolder = bolFolders.next();

    // Check for label files (Labels_*.html or Labels_*.pdf)
    const files = bolFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName().toLowerCase();
      if (name.startsWith('labels_') && (name.endsWith('.html') || name.endsWith('.pdf'))) {
        return true;
      }
    }

    return false;
  } catch (e) {
    return false;
  }
}

/**
 * Cache key for label existence checks.
 * Keeps Drive traversal fast for the batch label generator modal.
 */
function buildLabelCacheKey_(dateVal, bol) {
  try {
    if (!dateVal || !bol) return null;
    const d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return null;

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthYear = months[d.getMonth()] + ' ' + d.getFullYear();
    const day = String(d.getDate()).padStart(2, '0');
    const safeBol = String(bol).trim().replace(/[\/\\?%*:|"<>\.]/g, '_');
    return `lbl:${monthYear}:${day}:${safeBol}`;
  } catch (e) {
    return null;
  }
}

/**
 * Generates labels for a specific BOL/Transaction.
 */
function shell_generateLabelsForBol(txnId, bolNumber) {
  try {
    if (!txnId) {
      return { success: false, message: 'Transaction ID is required.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('Master_Log');
    const skidsSheet = ss.getSheetByName('Inbound_Skids');

    if (!masterSheet || !skidsSheet) {
      return { success: false, message: 'Required sheets not found.' };
    }

    const mData = masterSheet.getDataRange().getValues();
    const sData = skidsSheet.getDataRange().getValues();

    // Get Master_Log headers and columns
    const mHeaderRow = findHeaderRow_(mData, ['Txn_ID', 'BOL_Number']);
    if (mHeaderRow < 0) return { success: false, message: 'Master_Log headers not found.' };

    const mHead = mData[mHeaderRow];
    const mCol = (n) => mHead.indexOf(n);
    const mTxn = mCol('Txn_ID');
    const mDate = mCol('Date_Received');
    const mBol = mCol('BOL_Number');
    const mPush = mCol('Push #');
    const mMan = mCol('Manufacturer');

    // Get Inbound_Skids headers and columns
    const sHeaderRow = findHeaderRow_(sData, ['TXN_ID', 'Skid_ID']);
    if (sHeaderRow < 0) return { success: false, message: 'Inbound_Skids headers not found.' };

    const sHead = sData[sHeaderRow];
    const sCol = (n) => sHead.indexOf(n);
    const sTxn = sCol('TXN_ID');
    const sSkid = sCol('Skid_ID');
    const sFbpn = sCol('FBPN');
    const sQty = sCol('Qty_on_Skid');
    const sSku = sCol('SKU');
    const sProj = sCol('Project');
    const sSeq = sCol('Skid_Sequence');

    // Find the master record for this transaction
    let masterRecord = null;
    for (let i = mHeaderRow + 1; i < mData.length; i++) {
      const rowTxn = String(mData[i][mTxn] || '').trim();
      if (rowTxn === txnId) {
        masterRecord = {
          date: mData[i][mDate],
          bol: mBol >= 0 ? mData[i][mBol] : bolNumber,
          push: mPush >= 0 ? mData[i][mPush] : '',
          manufacturer: mMan >= 0 ? mData[i][mMan] : ''
        };
        break;
      }
    }

    if (!masterRecord) {
      return { success: false, message: 'Transaction not found in Master_Log.' };
    }

    // Get all skids for this transaction
    const skids = [];
    for (let i = sHeaderRow + 1; i < sData.length; i++) {
      if (String(sData[i][sTxn] || '').trim() === txnId) {
        skids.push({
          skidId: sSkid >= 0 ? sData[i][sSkid] : '',
          fbpn: sFbpn >= 0 ? sData[i][sFbpn] : '',
          qty: sQty >= 0 ? sData[i][sQty] : 0,
          sku: sSku >= 0 ? sData[i][sSku] : '',
          project: sProj >= 0 ? sData[i][sProj] : '',
          skidSeq: sSeq >= 0 ? (sData[i][sSeq] || 1) : 1
        });
      }
    }

    if (skids.length === 0) {
      return { success: false, message: 'No skids found for this transaction in Inbound_Skids.' };
    }

    // Build label data
    const totalSkids = skids.length;
    const labelData = skids.map(skid => ({
      skidId: skid.skidId,
      fbpn: skid.fbpn,
      quantity: skid.qty,
      sku: skid.sku,
      manufacturer: masterRecord.manufacturer,
      project: skid.project,
      pushNumber: masterRecord.push,
      dateReceived: formatLabelDate_(masterRecord.date),
      skidNumber: skid.skidSeq,
      totalSkids: totalSkids
    }));

    // Create target folder
    let targetFolder = null;
    try {
      const d = (masterRecord.date instanceof Date) ? masterRecord.date : new Date(masterRecord.date);
      if (!isNaN(d.getTime())) {
        targetFolder = createInboundFolder_(d, masterRecord.bol);
      }
    } catch (e) {
      Logger.log('Could not create target folder: ' + e);
    }

    // Generate labels
    const result = generateSkidLabels(labelData, {
      bolNumber: masterRecord.bol,
      targetFolder: targetFolder
    });

    if (result && result.success) {
      return {
        success: true,
        message: `Generated ${labelData.length} label(s) for BOL ${masterRecord.bol}`,
        pdfUrl: result.pdfUrl,
        htmlUrl: result.htmlUrl,
        labelCount: labelData.length
      };
    } else {
      return { success: false, message: result.message || 'Label generation failed.' };
    }

  } catch (err) {
    Logger.log('shell_generateLabelsForBol error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.message };
  }
}

/**
 * Helper to format date for labels.
 */
function formatLabelDate_(dateVal) {
  if (!dateVal) return '';
  try {
    const d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return String(dateVal);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  } catch (e) {
    return String(dateVal);
  }
}

/**
 * Creates the inbound folder structure: Inbound_Uploads/{MonthYear}/{Day}/{BOL}
 */
function createInboundFolder_(date, bolNumber) {
  const rootId = (typeof FOLDERS !== 'undefined' && FOLDERS.INBOUND_UPLOADS)
    ? FOLDERS.INBOUND_UPLOADS
    : FOLDERS.IMS_ROOT;

  const rootFolder = DriveApp.getFolderById(rootId);

  // Month folder (e.g., "Jan 2024")
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthYear = months[date.getMonth()] + ' ' + date.getFullYear();
  const monthFolder = getOrCreateSubfolder_(rootFolder, monthYear);

  // Day folder (e.g., "15")
  const day = String(date.getDate()).padStart(2, '0');
  const dayFolder = getOrCreateSubfolder_(monthFolder, day);

  // BOL folder
  const safeBol = String(bolNumber || 'NO_BOL').trim().replace(/[\/\\?%*:|"<>\.]/g, '_');
  const bolFolder = getOrCreateSubfolder_(dayFolder, safeBol);

  return bolFolder;
}

/**
 * Gets or creates a subfolder.
 */
function getOrCreateSubfolder_(parentFolder, name) {
  const folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(name);
}
function escapeHtml_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ----------------------------------------------------------------------------
// BOL LOOKUP FOR MANUAL LABEL MODAL
// ----------------------------------------------------------------------------
/**
 * Looks up BOL data from Master_Log to pre-populate the manual label form.
 * Returns the first matching record's data (Manufacturer, Project, Push #, Total Skids).
 */
function shell_lookupBOLData(bolNumber) {
  try {
    if (!bolNumber || !String(bolNumber).trim()) {
      return { success: false, message: 'BOL number is required.' };
    }

    const bol = String(bolNumber).trim().toUpperCase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master_Log');

    if (!sheet) {
      return { success: false, message: 'Master_Log sheet not found.' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: false, message: 'No data in Master_Log.' };
    }

    const headerRow = findHeaderRow_(data, ['BOL_Number', 'Manufacturer']);
    if (headerRow < 0) {
      return { success: false, message: 'BOL_Number/Manufacturer columns not found in Master_Log.' };
    }

    const headers = data[headerRow];
    const colIdx = (name) => headers.indexOf(name);

    const bolCol = colIdx('BOL_Number');
    const mfgCol = colIdx('Manufacturer');
    const projCol = colIdx('Project');
    const pushCol = colIdx('Push #');
    const totalSkidsCol = colIdx('Total_Skid_Count');
    const fbpnCol = colIdx('FBPN');
    const warehouseCol = colIdx('Warehouse');

    if (bolCol < 0) {
      return { success: false, message: 'BOL_Number column not found in Master_Log.' };
    }

    // Find first matching BOL record
    for (let i = headerRow + 1; i < data.length; i++) {
      const rowBol = String(data[i][bolCol] || '').trim().toUpperCase();
      if (rowBol === bol) {
        return {
          success: true,
          data: {
            manufacturer: mfgCol >= 0 ? String(data[i][mfgCol] || '') : '',
            project: projCol >= 0 ? String(data[i][projCol] || '') : '',
            pushNumber: pushCol >= 0 ? String(data[i][pushCol] || '') : '',
            totalSkids: totalSkidsCol >= 0 ? (parseInt(data[i][totalSkidsCol]) || 1) : 1,
            fbpn: fbpnCol >= 0 ? String(data[i][fbpnCol] || '') : '',
            warehouse: warehouseCol >= 0 ? String(data[i][warehouseCol] || '') : ''
          }
        };
      }
    }

    return { success: false, message: 'BOL not found in Master_Log.' };

  } catch (err) {
    Logger.log('shell_lookupBOLData error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.message };
  }
}

// ----------------------------------------------------------------------------
// CUSTOMER PORTAL FUNCTIONS
// ----------------------------------------------------------------------------
function authenticateUser(email) {
  // Stub - implement customer authentication if needed
  return { success: false, message: 'authenticateUser not implemented' };
}

function searchInventoryForCustomer(email, criteria) {
  // Use getStockTotalsForWebApp for inventory search
  return getStockTotalsForWebApp({ email: email }, criteria);
}

function getAvailableFBPNsForOrder(email) {
  // Return FBPN list for ordering
  return getFBPNList();
}

function submitCustomerOrderFromPortal(email, data) {
  // Use processCustomerOrder for order submission
  return processCustomerOrder(data);
}

function getUserProjectAccess(email) {
  // Get user project access from Customer_Access sheet
  const context = getUserContextDirect_();
  return context.projectAccess || [];
}

function getProjectsForPortal() {
  return getProjects();
}

// ----------------------------------------------------------------------------
// MANUAL SKID LABEL GENERATION
// ----------------------------------------------------------------------------
function shell_generateManualSkidLabel(data) {
  return generateManualSkidLabelFromModal(data);
}

function generateManualSkidLabelFromModal(data) {
  try {
    if (!data.fbpn) throw new Error('FBPN is required.');
    if (!data.qty || data.qty <= 0) throw new Error('Quantity must be greater than 0.');
    if (!data.manufacturer) throw new Error('Manufacturer is required.');
    if (!data.project) throw new Error('Project is required.');

    const copies = Math.min(50, Math.max(1, parseInt(data.copies) || 1));
    const skidNumber = parseInt(data.skidNumber) || 1;
    const totalSkids = parseInt(data.totalSkids) || 1;
    const now = new Date();
    const dateStr = formatDateISO_(now);

    const labelData = [];
    for (let i = 0; i < copies; i++) {
      const skidId = generateRandomId_('SKD-', 8);
      const sku = generateSKU(data.fbpn, data.manufacturer);

      labelData.push({
        skidId: skidId,
        fbpn: String(data.fbpn).toUpperCase().trim(),
        quantity: data.qty,
        sku: sku,
        manufacturer: String(data.manufacturer).trim(),
        project: String(data.project).trim(),
        pushNumber: String(data.push || '').trim(),
        dateReceived: dateStr,
        skidNumber: skidNumber,
        totalSkids: totalSkids,
        notes: String(data.notes || '').trim()
      });
    }

    const bolNumber = String(data.bol || 'MANUAL').trim();
    const result = generateSkidLabels(labelData, { bolNumber: bolNumber });

    if (!result || !result.success) {
      return { success: false, message: (result && result.message) ? result.message : 'Label generation failed.' };
    }

    return {
      success: true,
      pdfUrl: result.pdfUrl || '',
      htmlUrl: result.htmlUrl || '',
      labelCount: labelData.length,
      message: 'Successfully generated ' + labelData.length + ' label(s).'
    };
  } catch (err) {
    Logger.log('generateManualSkidLabelFromModal error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.message };
  }
}

function generateRandomId_(prefix, length) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return prefix + result;
}

// ----------------------------------------------------------------------------
// PAST DELIVERY SKID LABEL GENERATION
// ----------------------------------------------------------------------------

/**
 * Gets skids for a specific TXN_ID directly from Inbound_Skids sheet.
 * Used by the PastDeliverySkidLabelModal View button.
 */
function shell_getSkidsByTxnId(txnId) {
  try {
    if (!txnId || !String(txnId).trim()) {
      return { success: false, message: 'TXN_ID is required.' };
    }

    const txn = String(txnId).trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const skidsSheet = ss.getSheetByName('Inbound_Skids');
    const masterSheet = ss.getSheetByName('Master_Log');

    if (!skidsSheet) {
      return { success: false, message: 'Inbound_Skids sheet not found.' };
    }

    const sData = skidsSheet.getDataRange().getValues();
    const sHeaderRow = findHeaderRow_(sData, ['TXN_ID', 'Skid_ID']);
    if (sHeaderRow < 0) return { success: false, message: 'Inbound_Skids headers not found.' };

    const sHead = sData[sHeaderRow];
    const sCol = (n) => sHead.indexOf(n);

    // Get additional info from Master_Log if available
    let masterInfo = { manufacturer: '', pushNumber: '', bolNumber: '' };
    if (masterSheet) {
      const mData = masterSheet.getDataRange().getValues();
      const mHeaderRow = findHeaderRow_(mData, ['Txn_ID', 'BOL_Number']);
      if (mHeaderRow >= 0) {
        const mHead = mData[mHeaderRow];
        const mCol = (n) => mHead.indexOf(n);
        for (let i = mHeaderRow + 1; i < mData.length; i++) {
          if (String(mData[i][mCol('Txn_ID')] || '').trim() === txn) {
            masterInfo.manufacturer = mData[i][mCol('Manufacturer')] || '';
            masterInfo.pushNumber = mData[i][mCol('Push #')] || '';
            masterInfo.bolNumber = mData[i][mCol('BOL_Number')] || '';
            break;
          }
        }
      }
    }

    // Get skids for this TXN_ID
    const skids = [];
    for (let i = sHeaderRow + 1; i < sData.length; i++) {
      const rowTxn = String(sData[i][sCol('TXN_ID')] || '').trim();
      if (rowTxn === txn) {
        const dateVal = sData[i][sCol('Date')];
        let dateStr = '';
        if (dateVal) {
          try {
            const d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            if (!isNaN(d.getTime())) {
              dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            }
          } catch (e) {
            dateStr = String(dateVal);
          }
        }

        skids.push({
          txnId: rowTxn,
          skidId: sData[i][sCol('Skid_ID')] || '',
          fbpn: sData[i][sCol('FBPN')] || '',
          qty: sData[i][sCol('Qty_on_Skid')] || 0,
          sku: sData[i][sCol('SKU')] || '',
          project: sData[i][sCol('Project')] || '',
          skidSeq: sData[i][sCol('Skid_Sequence')] || 1,
          manufacturer: masterInfo.manufacturer,
          pushNumber: masterInfo.pushNumber,
          dateReceived: dateStr
        });
      }
    }

    if (skids.length === 0) {
      return { success: false, message: `No skids found for TXN_ID: ${txnId}` };
    }

    return { success: true, skids: skids };

  } catch (err) {
    Logger.log('shell_getSkidsByTxnId error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.message };
  }
}

/**
 * Gets skids for a specific BOL number from Inbound_Skids sheet.
 * Used by the PastDeliverySkidLabelModal.
 */
function shell_getSkidsForBOL(bolNumber) {
  try {
    if (!bolNumber || !String(bolNumber).trim()) {
      return { success: false, message: 'BOL number is required.' };
    }

    const bol = String(bolNumber).trim().toUpperCase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('Master_Log');
    const skidsSheet = ss.getSheetByName('Inbound_Skids');

    if (!masterSheet || !skidsSheet) {
      return { success: false, message: 'Required sheets not found.' };
    }

    const mData = masterSheet.getDataRange().getValues();
    const sData = skidsSheet.getDataRange().getValues();

    // Find header rows
    const mHeaderRow = findHeaderRow_(mData, ['Txn_ID', 'BOL_Number']);
    if (mHeaderRow < 0) return { success: false, message: 'Master_Log headers not found.' };

    const sHeaderRow = findHeaderRow_(sData, ['TXN_ID', 'Skid_ID']);
    if (sHeaderRow < 0) return { success: false, message: 'Inbound_Skids headers not found.' };

    const mHead = mData[mHeaderRow];
    const mCol = (n) => mHead.indexOf(n);

    const sHead = sData[sHeaderRow];
    const sCol = (n) => sHead.indexOf(n);

    // Find TXN IDs matching this BOL
    const txnMap = {};
    for (let i = mHeaderRow + 1; i < mData.length; i++) {
      const rowBol = String(mData[i][mCol('BOL_Number')] || '').trim().toUpperCase();
      if (rowBol === bol) {
        const txnId = mData[i][mCol('Txn_ID')];
        if (txnId && !txnMap[txnId]) {
          txnMap[txnId] = {
            manufacturer: mData[i][mCol('Manufacturer')] || '',
            pushNumber: mData[i][mCol('Push #')] || '',
            date: mData[i][mCol('Date_Received')]
          };
        }
      }
    }

    const txnIds = Object.keys(txnMap);
    if (txnIds.length === 0) {
      return { success: false, message: `No transactions found for BOL: ${bolNumber}` };
    }

    // Get skids for matching TXN IDs
    const skids = [];
    for (let i = sHeaderRow + 1; i < sData.length; i++) {
      const rowTxn = String(sData[i][sCol('TXN_ID')] || '').trim();
      if (txnMap[rowTxn]) {
        const info = txnMap[rowTxn];
        const dateVal = sData[i][sCol('Date')] || info.date;
        let dateStr = '';
        if (dateVal) {
          try {
            const d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            if (!isNaN(d.getTime())) {
              dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            }
          } catch (e) {
            dateStr = String(dateVal);
          }
        }

        skids.push({
          txnId: rowTxn,
          skidId: sData[i][sCol('Skid_ID')] || '',
          fbpn: sData[i][sCol('FBPN')] || '',
          qty: sData[i][sCol('Qty_on_Skid')] || 0,
          sku: sData[i][sCol('SKU')] || '',
          project: sData[i][sCol('Project')] || '',
          skidSeq: sData[i][sCol('Skid_Sequence')] || 1,
          manufacturer: info.manufacturer,
          pushNumber: info.pushNumber,
          dateReceived: dateStr
        });
      }
    }

    if (skids.length === 0) {
      return { success: false, message: `No skids found for BOL: ${bolNumber}` };
    }

    return { success: true, skids: skids };

  } catch (err) {
    Logger.log('shell_getSkidsForBOL error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.message };
  }
}

/**
 * Generates skid labels for selected skids from a past delivery.
 * Used by the PastDeliverySkidLabelModal.
 */
function shell_generatePastDeliverySkidLabels(payload) {
  try {
    if (!payload || !payload.skids || payload.skids.length === 0) {
      return { success: false, message: 'No skids selected.' };
    }

    const copies = Math.min(10, Math.max(1, parseInt(payload.copies) || 1));
    const bolNumber = String(payload.bolNumber || 'PAST-REPRINT').trim();

    const labelData = [];
    const totalSkids = payload.skids.length;

    payload.skids.forEach((skid, idx) => {
      for (let c = 0; c < copies; c++) {
        labelData.push({
          skidId: skid.skidId,
          fbpn: String(skid.fbpn || '').toUpperCase(),
          quantity: skid.qty || 0,
          sku: skid.sku || generateSKU(skid.fbpn, skid.manufacturer),
          manufacturer: skid.manufacturer || '',
          project: skid.project || '',
          pushNumber: skid.pushNumber || '',
          dateReceived: skid.dateReceived || formatLabelDate_(new Date()),
          skidNumber: skid.skidSeq || (idx + 1),
          totalSkids: totalSkids
        });
      }
    });

    if (labelData.length === 0) {
      return { success: false, message: 'No label data to generate.' };
    }

    const result = generateSkidLabels(labelData, { bolNumber: bolNumber });

    if (!result || !result.success) {
      return { success: false, message: (result && result.message) ? result.message : 'Label generation failed.' };
    }

    return {
      success: true,
      pdfUrl: result.pdfUrl || '',
      htmlUrl: result.htmlUrl || '',
      labelCount: labelData.length,
      message: `Successfully generated ${labelData.length} label(s) for ${payload.skids.length} skid(s).`
    };

  } catch (err) {
    Logger.log('shell_generatePastDeliverySkidLabels error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.message };
  }
}

// ----------------------------------------------------------------------------
// WEBAPP FUNCTIONS
// ----------------------------------------------------------------------------
function getUserContext() {
  return getUserContextDirect_();
}

function getUserContextDirect_() {
  try {
    const email = Session.getActiveUser().getEmail();

    if (!email) {
      return {
        authenticated: false,
        error: 'Unable to retrieve user email. Please ensure you are signed in.'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Customer_Access');

    if (!sheet) {
      Logger.log('Customer_Access sheet not found');
      return {
        authenticated: false,
        error: 'System configuration error: Customer_Access sheet not found.'
      };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const emailCol = headers.indexOf('Email');
    const nameCol = headers.indexOf('Name');
    const companyCol = headers.indexOf('Company Name');
    const accessCol = headers.indexOf('Access_Level');
    const projectCol = headers.indexOf('Project_Access');
    const activeCol = headers.indexOf('Active');

    const normalizedEmail = email.toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][emailCol] || '').toLowerCase().trim();

      if (rowEmail === normalizedEmail) {
        const projectAccessRaw = String(data[i][projectCol] || '').trim();
        const projectAccess = projectAccessRaw.toUpperCase() === 'ALL'
          ? ['ALL']
          : projectAccessRaw.split(',').map(p => p.trim()).filter(p => p);

        const activeRaw = data[i][activeCol];
        const isActive = activeRaw === true ||
          String(activeRaw).toUpperCase() === 'TRUE' ||
          String(activeRaw).toUpperCase() === 'YES' ||
          String(activeRaw).toUpperCase() === 'Y' ||
          String(activeRaw).toUpperCase() === 'ACTIVE' ||
          activeRaw === 1 ||
          String(activeRaw) === '1' ||
          (activeCol < 0);

        if (!isActive) {
          return {
            authenticated: false,
            error: 'Your account has been deactivated.'
          };
        }

        const accessLevel = data[i][accessCol] || 'Standard';

        return {
          authenticated: true,
          email: email,
          name: data[i][nameCol] || email.split('@')[0],
          company: data[i][companyCol] || '',
          accessLevel: accessLevel,
          projectAccess: projectAccess,
          isActive: true,
          permissions: buildPermissionsFromLevel_(accessLevel),
          timestamp: new Date().toISOString()
        };
      }
    }

    Logger.log('Access denied for unregistered user: ' + email);
    return {
      authenticated: false,
      error: 'Access denied. Your email is not registered in the system.'
    };

  } catch (error) {
    Logger.log('getUserContext error: ' + error.toString());
    return {
      authenticated: false,
      error: 'Authentication error: ' + error.message
    };
  }
}

function buildPermissionsFromLevel_(accessLevel) {
  const level = String(accessLevel || '').toUpperCase();
  const isMEI = level === 'MEI';
  const isTurner = level === 'TURNER';

  return {
    isAdmin: isMEI,
    canViewAllOrders: isMEI || isTurner,
    canCreateOrders: true,
    canAccessInventoryOps: isMEI,
    canAccessReports: isMEI || isTurner,
    canGenerateDocs: isMEI,
    canAccessInbound: isMEI,
    canAccessCycleCount: isMEI
  };
}

// ============================================================================
// WEBAPP FORWARDERS - DASHBOARD
// ============================================================================

function getDashboardMetrics() {
  return getDashboardMetricsDirect_();
}

function getDashboardMetricsDirect_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metrics = {
    orders: { pending: 0, processing: 0, shipped: 0 },
    inventory: { totalSKUs: 0, lowStock: 0, outOfStock: 0 },
    inbound: { scheduled: 0, received: 0 }
  };

  try {
    const ordersSheet = ss.getSheetByName('Customer_Orders');
    if (ordersSheet) {
      const ordersData = ordersSheet.getDataRange().getValues();
      const statusCol = ordersData[0].indexOf('Request_Status');
      if (statusCol >= 0) {
        for (let i = 1; i < ordersData.length; i++) {
          const status = String(ordersData[i][statusCol] || '').toLowerCase();
          if (status.includes('pending')) metrics.orders.pending++;
          else if (status.includes('processing') || status.includes('picking')) metrics.orders.processing++;
          else if (status.includes('shipped') || status.includes('delivered')) metrics.orders.shipped++;
        }
      }
    }

    const stockSheet = ss.getSheetByName('Stock_Totals');
    if (stockSheet) {
      metrics.inventory.totalSKUs = Math.max(0, stockSheet.getLastRow() - 1);
    }
  } catch (e) {
    Logger.log('getDashboardMetrics error: ' + e.toString());
  }

  return { success: true, metrics: metrics };
}

// ============================================================================
// WEBAPP FORWARDERS - CUSTOMER ORDERS
// ============================================================================

function getCustomerOrders(context) {
  return getCustomerOrdersDirect_(context);
}

function getCustomerOrdersDirect_(context) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Customer_Orders');

    if (!sheet) {
      return { success: false, message: 'Customer_Orders sheet not found', orders: [] };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, orders: [] };
    }

    const headers = data[0];
    const colMap = {};
    ['Order_ID', 'Task_Number', 'Project', 'NBD', 'Company', 'Order_Title',
      'Deliver_To', 'Request_Status', 'Stock_Status', 'Created_TS'].forEach(h => {
        colMap[h] = headers.indexOf(h);
      });

    const orders = [];
    const userCompany = (context && context.accessLevel === 'Standard' && context.company)
      ? context.company.toLowerCase() : '';

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const orderId = colMap['Order_ID'] >= 0 ? row[colMap['Order_ID']] : '';
      if (!orderId) continue;

      const rowCompany = colMap['Company'] >= 0 ? String(row[colMap['Company']] || '').toLowerCase() : '';
      if (userCompany && rowCompany !== userCompany) continue;

      orders.push({
        orderId: String(orderId),
        taskNumber: colMap['Task_Number'] >= 0 ? String(row[colMap['Task_Number']] || '') : '',
        project: colMap['Project'] >= 0 ? String(row[colMap['Project']] || '') : '',
        nbd: colMap['NBD'] >= 0 ? formatDateISO_(row[colMap['NBD']]) : '',
        company: colMap['Company'] >= 0 ? String(row[colMap['Company']] || '') : '',
        orderTitle: colMap['Order_Title'] >= 0 ? String(row[colMap['Order_Title']] || '') : '',
        status: colMap['Request_Status'] >= 0 ? String(row[colMap['Request_Status']] || 'Pending') : 'Pending',
        stockStatus: colMap['Stock_Status'] >= 0 ? String(row[colMap['Stock_Status']] || '') : '',
        createdAt: colMap['Created_TS'] >= 0 ? formatDateISO_(row[colMap['Created_TS']]) : ''
      });
    }

    orders.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));

    return { success: true, orders: orders, totalCount: orders.length };
  } catch (error) {
    return { success: false, message: error.toString(), orders: [] };
  }
}

/**
 * Formats date as ISO (yyyy-MM-dd) for web app use
 */
function formatDateISO_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

// ============================================================================
// WEBAPP FORWARDERS - DOCUMENT GENERATION
// ============================================================================

function regenerateOrderDoc(orderId, docType) {
  try {
    if (!orderId) {
      return { success: false, message: 'Order ID is required' };
    }

    const orderData = getFullOrderData_(orderId);

    if (!orderData) {
      return { success: false, message: 'Order not found: ' + orderId };
    }

    if (!orderData.items || orderData.items.length === 0) {
      return { success: false, message: 'No items found for order: ' + orderId };
    }

    let result;

    switch (docType) {
      case 'PICK':
        const pickData = {
          orderNumber: orderData.orderNumber,
          orderId: orderData.orderId,
          taskNumber: orderData.taskNumber,
          company: orderData.company,
          project: orderData.project,
          orderTitle: orderData.orderTitle,
          date: new Date().toLocaleDateString(),
          items: orderData.items.map(item => ({
            fbpn: item.fbpn,
            description: item.description,
            qtyRequested: item.qtyRequested,
            qtyToPick: item.qtyRequested,
            qty: item.qtyRequested
          }))
        };

        result = generatePickTicket(pickData);
        break;

      case 'PACKING':
      case 'TOC':
        const skids = [{
          skidNumber: 1,
          items: orderData.items.map(item => ({
            fbpn: item.fbpn,
            description: item.description,
            qtyRequested: item.qtyRequested,
            qtyOnSkid: item.qtyRequested,
            qty: item.qtyRequested
          }))
        }];

        const docData = {
          orderNumber: orderData.orderNumber,
          orderId: orderData.orderId,
          taskNumber: orderData.taskNumber,
          company: orderData.company,
          project: orderData.project,
          orderTitle: orderData.orderTitle,
          deliverTo: orderData.deliverTo || '',
          name: orderData.name || '',
          phoneNumber: orderData.phoneNumber || '',
          shipDate: orderData.shipDate || new Date().toLocaleDateString(),
          date: new Date().toLocaleDateString(),
          totalSkids: '1',
          skids: skids,
          items: orderData.items
        };

        if (docType === 'TOC') {
          result = generateTOC(docData);
        } else {
          result = generatePackingLists(docData);
        }
        break;

      default:
        return { success: false, message: 'Invalid document type: ' + docType };
    }

    if (result && (result.success || result.pdfUrl || result.url)) {
      return {
        success: true,
        url: result.pdfUrl || result.url,
        docType: docType,
        message: docType + ' generated successfully'
      };
    } else {
      return { success: false, message: result ? result.message : 'Document generation failed' };
    }

  } catch (e) {
    Logger.log('regenerateOrderDoc error: ' + e.toString());
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function getFullOrderData_(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('Customer_Orders');
  if (!orderSheet) return null;

  const orderData = orderSheet.getDataRange().getValues();
  const orderHeaders = orderData[0];

  const findCol = (names) => {
    const normalized = orderHeaders.map(h => String(h || '').toLowerCase().trim().replace(/[_\s]+/g, '_'));
    for (const name of names) {
      const idx = normalized.indexOf(name.toLowerCase().replace(/[_\s]+/g, '_'));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const col = {
    orderId: findCol(['order_id', 'order_number', 'orderid']),
    taskNumber: findCol(['task_number', 'task_number', 'task']),
    project: findCol(['project']),
    nbd: findCol(['nbd', 'need_by_date']),
    company: findCol(['company']),
    orderTitle: findCol(['order_title', 'title']),
    deliverTo: findCol(['deliver_to', 'delivery_address']),
    name: findCol(['name', 'contact_name']),
    phoneNumber: findCol(['phone_number', 'phone']),
    status: findCol(['request_status', 'status']),
    shipDate: findCol(['ship_date', 'shipped_date'])
  };

  let orderRow = null;
  for (let i = 1; i < orderData.length; i++) {
    const rowOrderId = col.orderId >= 0 ? String(orderData[i][col.orderId]) : '';
    if (rowOrderId === String(orderId)) {
      orderRow = orderData[i];
      break;
    }
  }

  if (!orderRow) return null;

  const order = {
    orderNumber: String(orderId),
    orderId: String(orderId),
    taskNumber: col.taskNumber >= 0 ? String(orderRow[col.taskNumber] || '') : '',
    project: col.project >= 0 ? String(orderRow[col.project] || '') : '',
    nbd: col.nbd >= 0 ? formatDateISO_(orderRow[col.nbd]) : '',
    company: col.company >= 0 ? String(orderRow[col.company] || '') : '',
    orderTitle: col.orderTitle >= 0 ? String(orderRow[col.orderTitle] || '') : '',
    deliverTo: col.deliverTo >= 0 ? String(orderRow[col.deliverTo] || '') : '',
    name: col.name >= 0 ? String(orderRow[col.name] || '') : '',
    phoneNumber: col.phoneNumber >= 0 ? String(orderRow[col.phoneNumber] || '') : '',
    status: col.status >= 0 ? String(orderRow[col.status] || '') : '',
    shipDate: col.shipDate >= 0 ? formatDateISO_(orderRow[col.shipDate]) : '',
    date: new Date().toLocaleDateString()
  };

  const itemsSheet = ss.getSheetByName('Requested_Items');
  order.items = [];

  if (itemsSheet) {
    const itemsData = itemsSheet.getDataRange().getValues();
    const itemHeaders = itemsData[0];

    const findItemCol = (names) => {
      const normalized = itemHeaders.map(h => String(h || '').toLowerCase().trim().replace(/[_\s]+/g, '_'));
      for (const name of names) {
        const idx = normalized.indexOf(name.toLowerCase().replace(/[_\s]+/g, '_'));
        if (idx >= 0) return idx;
      }
      return -1;
    };

    const itemCol = {
      orderId: findItemCol(['order_id', 'order_number']),
      fbpn: findItemCol(['fbpn']),
      description: findItemCol(['description', 'desc']),
      qtyRequested: findItemCol(['qty_requested', 'qty', 'quantity']),
      sku: findItemCol(['sku'])
    };

    for (let i = 1; i < itemsData.length; i++) {
      const rowOrderId = itemCol.orderId >= 0 ? String(itemsData[i][itemCol.orderId]) : '';

      if (rowOrderId === String(orderId) || rowOrderId === String(Math.trunc(Number(orderId)))) {
        const fbpn = itemCol.fbpn >= 0 ? String(itemsData[i][itemCol.fbpn] || '').trim() : '';
        if (!fbpn) continue;

        order.items.push({
          fbpn: fbpn,
          description: itemCol.description >= 0 ? String(itemsData[i][itemCol.description] || '') : '',
          qtyRequested: itemCol.qtyRequested >= 0 ? Number(itemsData[i][itemCol.qtyRequested] || 0) : 0,
          qty: itemCol.qtyRequested >= 0 ? Number(itemsData[i][itemCol.qtyRequested] || 0) : 0,
          sku: itemCol.sku >= 0 ? String(itemsData[i][itemCol.sku] || '') : ''
        });
      }
    }
  }

  return order;
}

// ============================================================================
// WEBAPP FORWARDERS - ORDER DATA FOR SHIPPING MODAL
// ============================================================================

function getOrderDataForShipping(orderId) {
  return getFullOrderData_(orderId);
}

function getOrderByTaskNumber(taskNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('Customer_Orders');
  const itemsSheet = ss.getSheetByName('Requested_Items');

  if (!orderSheet || !itemsSheet) return null;

  const orderData = orderSheet.getDataRange().getValues();
  const orderHeaders = orderData[0];

  const findCol = (headers, names) => {
    const normalized = headers.map(h => String(h || '').toLowerCase().trim().replace(/[_\s]+/g, '_'));
    for (const name of names) {
      const idx = normalized.indexOf(name.toLowerCase().replace(/[_\s]+/g, '_'));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const cTask = findCol(orderHeaders, ['task_number', 'task']);
  const cOrder = findCol(orderHeaders, ['order_id', 'order_number']);
  const cProj = findCol(orderHeaders, ['project']);
  const cComp = findCol(orderHeaders, ['company']);
  const cTitle = findCol(orderHeaders, ['order_title', 'title']);
  const cDeliver = findCol(orderHeaders, ['deliver_to', 'delivery_address']);
  const cName = findCol(orderHeaders, ['name', 'contact_name']);
  const cPhone = findCol(orderHeaders, ['phone_number', 'phone']);

  const key = String(taskNumber).trim();
  let orderRow = null;

  for (let r = 1; r < orderData.length; r++) {
    const row = orderData[r];
    const vTask = cTask >= 0 ? String(row[cTask] || '').trim() : '';
    const vOrder = cOrder >= 0 ? String(row[cOrder] || '').trim() : '';

    if (vTask === key || vOrder === key ||
      String(Math.trunc(Number(vTask))) === key ||
      String(Math.trunc(Number(vOrder))) === key) {
      orderRow = row;
      break;
    }
  }

  if (!orderRow) return null;

  const orderId = (cOrder >= 0 ? String(orderRow[cOrder] || '') : '') || key;

  const itemsData = itemsSheet.getDataRange().getValues();
  const itemHeaders = itemsData[0];

  const iOrder = findCol(itemHeaders, ['order_id', 'order_number', 'task_number']);
  const iFbpn = findCol(itemHeaders, ['fbpn']);
  const iDesc = findCol(itemHeaders, ['description', 'desc']);
  const iQty = findCol(itemHeaders, ['qty_requested', 'qty']);

  const items = [];
  const matchKeys = [key, orderId, String(Math.trunc(Number(key))), String(Math.trunc(Number(orderId)))];

  for (let j = 1; j < itemsData.length; j++) {
    const row = itemsData[j];
    const ok = iOrder >= 0 ? String(row[iOrder] || '').trim() : '';

    if (!matchKeys.includes(ok)) continue;

    const fbpn = iFbpn >= 0 ? String(row[iFbpn] || '').trim() : '';
    if (!fbpn) continue;

    items.push({
      fbpn: fbpn,
      description: iDesc >= 0 ? String(row[iDesc] || '') : '',
      qtyRequested: iQty >= 0 ? Number(row[iQty] || 0) : 0
    });
  }

  const combined = {};
  items.forEach(item => {
    if (!combined[item.fbpn]) {
      combined[item.fbpn] = item;
    } else {
      combined[item.fbpn].qtyRequested += item.qtyRequested;
    }
  });

  return {
    orderId: orderId,
    orderNumber: orderId,
    orderTitle: cTitle >= 0 ? String(orderRow[cTitle] || '') : '',
    company: cComp >= 0 ? String(orderRow[cComp] || '') : '',
    project: cProj >= 0 ? String(orderRow[cProj] || '') : '',
    deliverTo: cDeliver >= 0 ? String(orderRow[cDeliver] || '') : '',
    name: cName >= 0 ? String(orderRow[cName] || '') : '',
    phoneNumber: cPhone >= 0 ? String(orderRow[cPhone] || '') : '',
    items: Object.values(combined)
  };
}

// ============================================================================
// WEBAPP FORWARDERS - FILE UPLOAD
// ============================================================================

const AUTOMATION_FOLDER_ID = '1L3mjeQizzjVU5uTqGxv1sOUOuq25I2pM';

function uploadToAutomationFolder(fileData) {
  try {
    if (!fileData || !fileData.content || !fileData.fileName) {
      return { success: false, message: 'Missing file data' };
    }

    const folder = DriveApp.getFolderById(AUTOMATION_FOLDER_ID);
    const decoded = Utilities.base64Decode(fileData.content);
    const blob = Utilities.newBlob(decoded,
      fileData.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      fileData.fileName);

    const file = folder.createFile(blob);

    return {
      success: true,
      message: 'File uploaded. It will be processed automatically.',
      fileUrl: file.getUrl()
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// WEBAPP FORWARDERS - FORM HELPERS
// ============================================================================

function getCompaniesFiltered(context) {
  if (context && context.accessLevel === 'Standard' && context.company) {
    return [context.company];
  }
  return getCompaniesDirect_();
}

function getCompaniesDirect_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Support_Sheet');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const companies = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) companies.add(data[i][0]);
  }
  return Array.from(companies).sort();
}

function getProjectsFiltered(company) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Support_Sheet');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const projects = new Set();

  for (let i = 1; i < data.length; i++) {
    const rowCompany = String(data[i][0] || '').trim();
    const rowProject = String(data[i][1] || '').trim();

    if (rowProject) {
      if (company && rowCompany.toLowerCase() === company.toLowerCase()) {
        projects.add(rowProject);
      } else if (!company) {
        projects.add(rowProject);
      }
    }
  }
  return Array.from(projects).sort();
}

function getNextTaskNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) return '1001';

  const data = sheet.getDataRange().getValues();
  const taskCol = data[0].indexOf('Task_Number');
  if (taskCol < 0) return '1001';

  let maxNum = 1000;
  for (let i = 1; i < data.length; i++) {
    const num = parseInt(String(data[i][taskCol]).replace(/\D/g, ''), 10);
    if (!isNaN(num) && num > maxNum) maxNum = num;
  }
  return String(maxNum + 1);
}

function validateFBPN(fbpn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Project_Master');
  if (!sheet) return { valid: false, message: 'Project_Master not found' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const fbpnCol = headers.indexOf('FBPN');
  const descCol = headers.indexOf('Description');

  if (fbpnCol < 0) return { valid: false, message: 'FBPN column not found' };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][fbpnCol]).toLowerCase() === fbpn.toLowerCase()) {
      return {
        valid: true,
        fbpn: data[i][fbpnCol],
        description: descCol >= 0 ? String(data[i][descCol] || '') : ''
      };
    }
  }
  return { valid: false, message: 'FBPN not found' };
}

function api_getSkidById(skidId) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Inbound_Skids');
  if (!sh) return { success: false, message: 'Inbound_Skids not found' };

  const id = String(skidId || '').trim().toUpperCase();
  const values = sh.getDataRange().getValues();
  const headers = values[0];

  const c = name => headers.indexOf(name);
  const idxSkid = c('Skid_ID');
  if (idxSkid === -1) return { success: false, message: 'Skid_ID column missing' };

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idxSkid]).toUpperCase() === id) {
      return {
        success: true,
        skid: {
          Skid_ID: id,
          FBPN: values[r][c('FBPN')],
          MFPN: values[r][c('MFPN')],
          Project: values[r][c('Project')],
          Qty: values[r][c('Qty_on_Skid')],
          SKU: values[r][c('SKU')],
          TXN_ID: values[r][c('TXN_ID')],
          Timestamp: values[r][c('Timestamp')]
        }
      };
    }
  }
  return { success: false, message: `Skid not found: ${id}` };
}

// ============================================================================
// CYCLE COUNT WRAPPER FUNCTIONS AND MODAL
// ============================================================================

/**
 * Opens the Cycle Count modal
 */
function openCycleCountModal() {
  const html = HtmlService.createTemplateFromFile('CycleCountModal')
    .evaluate()
    .setWidth(1200)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cycle Count');
}

/**
 * Wrapper function for frontend compatibility
 * Maps getBinsForCycleCount() to imsGetCycleCountBins()
 * Auto-initializes Cycle_Count sheet if missing
 */
function getBinsForCycleCount(filters) {
  try {
    // Check if Cycle_Count sheet exists, initialize if not
    const ss = SpreadsheetApp.getActive();
    let cycleSheet = ss.getSheetByName('Cycle_Count');

    if (!cycleSheet) {
      Logger.log('Cycle_Count sheet not found, initializing...');
      const initResult = initializeCycleCountSheet();
      if (!initResult.success) {
        return {
          success: false,
          message: 'Failed to initialize Cycle_Count sheet: ' + (initResult.message || 'Unknown error')
        };
      }
    }

    // Call the actual function
    const results = imsGetCycleCountBins(filters || {});

    // Return results in compatible format
    if (Array.isArray(results)) {
      return results;
    }

    return results;

  } catch (err) {
    Logger.log('getBinsForCycleCount error: ' + err.toString());
    return {
      success: false,
      message: 'Error loading bins: ' + err.message
    };
  }
}

/**
 * NOTE: imsGenerateCycleCountReportPdf is defined in CycleCount.js (lines 823-852)
 * It uses HTML-to-PDF conversion for better formatting.
 * No wrapper needed - the function is called directly from the modal.
 */

/**
 * Initialize Cycle Count sheet with proper headers
 */
function initializeCycleCountSheet() {
  const ss = SpreadsheetApp.getActive();
  let cycleSheet = ss.getSheetByName('Cycle_Count');

  if (!cycleSheet) {
    cycleSheet = ss.insertSheet('Cycle_Count');
  }

  // Check if headers exist
  if (cycleSheet.getLastRow() === 0) {
    const headers = [
      'Batch_ID',        // A
      'Status',          // B  (Open, In Progress, Completed, Canceled)
      'Created_At',      // C
      'Created_By',      // D
      'Bin_Code',        // E
      'FBPN',            // F
      'Manufacturer',    // G
      'Project',         // H
      'Current_Qty',     // I (snapshot at batch creation)
      'Counted_Qty',     // J
      'Variance',        // K (Counted - Current)
      'Notes',           // L
      'Counted_At',      // M
      'Counted_By'       // N
    ];

    cycleSheet.appendRow(headers);

    // Format header row
    const headerRange = cycleSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4a5568');
    headerRange.setFontColor('#ffffff');

    // Freeze header row
    cycleSheet.setFrozenRows(1);

    Logger.log('Cycle_Count sheet initialized with headers');
  }

  return { success: true, message: 'Cycle_Count sheet ready' };
}
// ============================================================================
// FBPN LOOKUP BACKEND FUNCTIONS - Add to IMS_Config.js or BinLookup.js
// ============================================================================

/**
 * Get comprehensive FBPN details including inbounds, outbounds, stock, and locations
 * @param {string} fbpn - The FBPN to lookup
 * @returns {Object} Comprehensive FBPN data
 */
function getFbpnDetails(fbpn) {
  try {
    const ss = SpreadsheetApp.getActive();
    const searchFbpn = (fbpn || '').toString().trim().toUpperCase();

    if (!searchFbpn) {
      return { success: false, message: 'FBPN is required' };
    }

    // Get all sheets
    const masterLog = ss.getSheetByName('Master_Log');
    const inboundSkids = ss.getSheetByName('Inbound_Skids');
    const outboundLog = ss.getSheetByName('OutboundLog');
    const stockTotals = ss.getSheetByName('Stock_Totals');
    const binStock = ss.getSheetByName('Bin_Stock');

    if (!masterLog || !stockTotals || !binStock) {
      return { success: false, message: 'Required sheets not found' };
    }

    // =====================================================
    // 1. GET INBOUND HISTORY
    // =====================================================
    const inbounds = [];

    if (inboundSkids) {
      const skidsData = inboundSkids.getDataRange().getValues();
      const skidsHeaders = skidsData[0];

      const skidCol = (name) => skidsHeaders.indexOf(name);
      const sTxn = skidCol('TXN_ID');
      const sFbpn = skidCol('FBPN');
      const sQty = skidCol('Qty_on_Skid');
      const sMfr = skidCol('Manufacturer');
      const sProj = skidCol('Project');
      const sTime = skidCol('Timestamp');
      const sSkidId = skidCol('Skid_ID');

      // Get Master_Log data for BOL numbers
      const masterData = masterLog.getDataRange().getValues();
      const masterHeaders = masterData[0];
      const mCol = (name) => masterHeaders.indexOf(name);
      const mTxn = mCol('Txn_ID');
      const mBol = mCol('BOL_Number');
      const mDate = mCol('Date_Received');

      // Create TXN to BOL map
      const txnBolMap = {};
      for (let i = 1; i < masterData.length; i++) {
        const txn = masterData[i][mTxn];
        if (txn && !txnBolMap[txn]) {
          txnBolMap[txn] = {
            bol: masterData[i][mBol],
            date: masterData[i][mDate]
          };
        }
      }

      // Find all inbound skids for this FBPN
      for (let i = 1; i < skidsData.length; i++) {
        const row = skidsData[i];
        const rowFbpn = (row[sFbpn] || '').toString().toUpperCase();

        if (rowFbpn === searchFbpn) {
          const txnId = row[sTxn];
          const bolInfo = txnBolMap[txnId] || { bol: txnId, date: null };

          inbounds.push({
            date: bolInfo.date ? formatDate(bolInfo.date) : 'N/A',
            bol: bolInfo.bol || txnId,
            txnId: txnId,
            skidId: row[sSkidId] || '',
            qty: row[sQty] || 0,
            manufacturer: row[sMfr] || '',
            project: row[sProj] || '',
            timestamp: row[sTime] ? formatDate(row[sTime]) : ''
          });
        }
      }
    }

    // =====================================================
    // 2. GET OUTBOUND HISTORY
    // =====================================================
    const outbounds = [];

    if (outboundLog) {
      const outData = outboundLog.getDataRange().getValues();
      const outHeaders = outData[0];

      const oCol = (name) => outHeaders.indexOf(name);
      const oFbpn = oCol('FBPN');
      const oTime = oCol('Timestamp');
      const oAction = oCol('Action');
      const oQty = oCol('Qty_Changed');
      const oBin = oCol('Bin_Code');
      const oTaskNum = oCol('Task_Number');
      const oOrderNum = oCol('Order_Number');
      const oUser = oCol('User_Email');

      for (let i = 1; i < outData.length; i++) {
        const row = outData[i];
        const rowFbpn = (row[oFbpn] || '').toString().toUpperCase();

        if (rowFbpn === searchFbpn) {
          outbounds.push({
            date: row[oTime] ? formatDate(row[oTime]) : '',
            action: row[oAction] || '',
            qty: Math.abs(row[oQty] || 0),
            binCode: row[oBin] || '',
            taskNumber: row[oTaskNum] || '',
            orderNumber: row[oOrderNum] || '',
            user: row[oUser] || ''
          });
        }
      }
    }

    // =====================================================
    // 3. GET STOCK BY MANUFACTURER
    // =====================================================
    const stockByMfr = [];
    const stockData = stockTotals.getDataRange().getValues();
    const stockHeaders = stockData[0];

    const stCol = (name) => stockHeaders.indexOf(name);
    const stFbpn = stCol('FBPN');
    const stMfr = stCol('Manufacturer');
    const stQty = stCol('Current_Stock');
    const stSku = stCol('SKU');
    const stDesc = stCol('Description');

    for (let i = 1; i < stockData.length; i++) {
      const row = stockData[i];
      const rowFbpn = (row[stFbpn] || '').toString().toUpperCase();

      if (rowFbpn === searchFbpn) {
        stockByMfr.push({
          manufacturer: row[stMfr] || 'Unknown',
          sku: row[stSku] || '',
          totalQty: row[stQty] || 0,
          description: row[stDesc] || ''
        });
      }
    }

    // =====================================================
    // 4. GET BIN LOCATIONS
    // =====================================================
    const binLocations = [];
    const binData = binStock.getDataRange().getValues();
    const binHeaders = binData[0];

    const bCol = (name) => binHeaders.indexOf(name);
    const bFbpn = bCol('FBPN');
    const bBin = bCol('Bin_Code');
    const bMfr = bCol('Manufacturer');
    const bProj = bCol('Project');
    const bQty = bCol('Current_Quantity');
    const bPct = bCol('Stock_Percentage');

    for (let i = 1; i < binData.length; i++) {
      const row = binData[i];
      const rowFbpn = (row[bFbpn] || '').toString().toUpperCase();

      if (rowFbpn === searchFbpn && (row[bQty] || 0) > 0) {
        binLocations.push({
          binCode: row[bBin] || '',
          qty: row[bQty] || 0,
          manufacturer: row[bMfr] || '',
          project: row[bProj] || '',
          stockPercentage: row[bPct] || 0
        });
      }
    }

    // Calculate totals
    const totalInbound = inbounds.reduce((sum, item) => sum + (item.qty || 0), 0);
    const totalOutbound = outbounds.reduce((sum, item) => sum + (item.qty || 0), 0);
    const totalStock = stockByMfr.reduce((sum, item) => sum + (item.totalQty || 0), 0);
    const totalBins = binLocations.length;

    return {
      success: true,
      fbpn: searchFbpn,
      summary: {
        totalInbound: totalInbound,
        totalOutbound: totalOutbound,
        currentStock: totalStock,
        totalBins: totalBins,
        manufacturerCount: stockByMfr.length
      },
      inbounds: inbounds.sort((a, b) => new Date(b.date) - new Date(a.date)),
      outbounds: outbounds.sort((a, b) => new Date(b.date) - new Date(a.date)),
      stockByManufacturer: stockByMfr.sort((a, b) => b.totalQty - a.totalQty),
      binLocations: binLocations.sort((a, b) => b.qty - a.qty)
    };

  } catch (err) {
    Logger.log('getFbpnDetails error: ' + err.toString());
    return {
      success: false,
      message: 'Error retrieving FBPN details: ' + err.message
    };
  }
}

/**
 * Helper function to format dates consistently
 */
function formatDate(date) {
  if (!date) return '';
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

/**
 * Generate PDF report for FBPN lookup
 */
function generateFbpnLookupReport(fbpn) {
  try {
    const details = getFbpnDetails(fbpn);

    if (!details.success) {
      return details;
    }

    const html = buildFbpnReportHtml_(details);

    // Create PDF
    const fileName = `FBPN_Report_${details.fbpn}_${new Date().getTime()}`;
    const htmlBlob = Utilities.newBlob(html, 'text/html', fileName + '.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName(fileName + '.pdf');

    // Save to reports folder
    const reportsFolder = DriveApp.getFolderById(FOLDERS.IMS_Reports);
    const pdfFile = reportsFolder.createFile(pdfBlob);

    return {
      success: true,
      pdfUrl: pdfFile.getUrl(),
      name: pdfFile.getName(),
      message: 'FBPN report generated successfully'
    };

  } catch (err) {
    Logger.log('generateFbpnLookupReport error: ' + err.toString());
    return {
      success: false,
      message: 'Error generating PDF: ' + err.message
    };
  }
}

/**
 * Build HTML for FBPN report PDF
 */
function buildFbpnReportHtml_(details) {
  const s = details.summary;
  const inboundRows = details.inbounds.map(item => `
    <tr>
      <td>${item.date}</td>
      <td>${item.bol}</td>
      <td>${item.skidId}</td>
      <td style="text-align:right">${item.qty}</td>
      <td>${item.manufacturer}</td>
      <td>${item.project}</td>
    </tr>
  `).join('');

  const outboundRows = details.outbounds.slice(0, 50).map(item => `
    <tr>
      <td>${item.date}</td>
      <td>${item.action}</td>
      <td style="text-align:right">${item.qty}</td>
      <td>${item.binCode}</td>
      <td>${item.taskNumber || item.orderNumber || '-'}</td>
    </tr>
  `).join('');

  const stockRows = details.stockByManufacturer.map(item => `
    <tr>
      <td>${item.manufacturer}</td>
      <td>${item.sku}</td>
      <td style="text-align:right">${item.totalQty}</td>
      <td>${item.description}</td>
    </tr>
  `).join('');

  const binRows = details.binLocations.map(item => `
    <tr>
      <td>${item.binCode}</td>
      <td style="text-align:right">${item.qty}</td>
      <td>${item.manufacturer}</td>
      <td>${item.project}</td>
      <td style="text-align:right">${item.stockPercentage.toFixed(1)}%</td>
    </tr>
  `).join('');

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>FBPN Report - ${details.fbpn}</title>
  <style>
    @page { size: landscape; margin: 0.5in; }
    body { font-family: Arial, sans-serif; font-size: 10pt; color: #333; margin: 0; padding: 20px; }
    h1 { font-size: 18pt; color: #1a1a1a; margin-bottom: 5px; }
    h2 { font-size: 14pt; color: #333; margin-top: 25px; margin-bottom: 10px; border-bottom: 2px solid #2563eb; padding-bottom: 5px; }
    .summary-grid { display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 25px; }
    .summary-card { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 6px; padding: 12px 16px; min-width: 140px; }
    .summary-card .label { font-size: 9pt; color: #666; text-transform: uppercase; margin-bottom: 4px; }
    .summary-card .value { font-size: 18pt; font-weight: 700; color: #1a1a1a; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    th { background: #2563eb; color: white; padding: 8px 10px; text-align: left; font-size: 9pt; text-transform: uppercase; }
    td { padding: 6px 10px; border-bottom: 1px solid #e0e0e0; font-size: 9pt; }
    tr:nth-child(even) { background: #f8f9fa; }
    .page-break { page-break-before: always; }
  </style>
</head>
<body>
  <h1>FBPN Comprehensive Report</h1>
  <div style="color: #666; margin-bottom: 20px;">FBPN: <strong>${details.fbpn}</strong> | Generated: ${new Date().toLocaleString()}</div>

  <div class="summary-grid">
    <div class="summary-card">
      <div class="label">Total Inbound</div>
      <div class="value">${s.totalInbound}</div>
    </div>
    <div class="summary-card">
      <div class="label">Total Outbound</div>
      <div class="value">${s.totalOutbound}</div>
    </div>
    <div class="summary-card">
      <div class="label">Current Stock</div>
      <div class="value">${s.currentStock}</div>
    </div>
    <div class="summary-card">
      <div class="label">Bin Locations</div>
      <div class="value">${s.totalBins}</div>
    </div>
    <div class="summary-card">
      <div class="label">Manufacturers</div>
      <div class="value">${s.manufacturerCount}</div>
    </div>
  </div>

  <h2>Current Stock by Manufacturer</h2>
  <table>
    <thead>
      <tr>
        <th>Manufacturer</th>
        <th>SKU</th>
        <th>Quantity</th>
        <th>Description</th>
      </tr>
    </thead>
    <tbody>${stockRows || '<tr><td colspan="4" style="text-align:center;">No stock data</td></tr>'}</tbody>
  </table>

  <h2>Current Bin Locations</h2>
  <table>
    <thead>
      <tr>
        <th>Bin Code</th>
        <th>Quantity</th>
        <th>Manufacturer</th>
        <th>Project</th>
        <th>Stock %</th>
      </tr>
    </thead>
    <tbody>${binRows || '<tr><td colspan="5" style="text-align:center;">No bins found</td></tr>'}</tbody>
  </table>

  <div class="page-break"></div>

  <h2>Inbound History</h2>
  <table>
    <thead>
      <tr>
        <th>Date</th>
        <th>BOL #</th>
        <th>Skid ID</th>
        <th>Qty</th>
        <th>Manufacturer</th>
        <th>Project</th>
      </tr>
    </thead>
    <tbody>${inboundRows || '<tr><td colspan="6" style="text-align:center;">No inbound history</td></tr>'}</tbody>
  </table>

  <h2>Outbound History (Last 50)</h2>
  <table>
    <thead>
      <tr>
        <th>Date</th>
        <th>Action</th>
        <th>Qty</th>
        <th>Bin</th>
        <th>Order/Task</th>
      </tr>
    </thead>
    <tbody>${outboundRows || '<tr><td colspan="5" style="text-align:center;">No outbound history</td></tr>'}</tbody>
  </table>
</body>
</html>`;
}

// ============================================================================
// INBOUND VERIFICATION FUNCTIONS
// ============================================================================

/**
 * Look up a skid by ID for verification
 * BOL and Manufacturer come from Master_Log (linked by TXN_ID)
 * @param {string} skidId - The skid ID to look up
 * @returns {Object} Result with skid data
 */
function getSkidForVerification(skidId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const skidsSheet = ss.getSheetByName('Inbound_Skids');
    const masterSheet = ss.getSheetByName('Master_Log');

    if (!skidsSheet) {
      return { success: false, message: 'Inbound_Skids sheet not found' };
    }

    // Get Inbound_Skids data
    const sData = skidsSheet.getDataRange().getValues();
    if (sData.length < 2) {
      return { success: false, message: 'No skid data found' };
    }

    const sHeaders = sData[0];
    const sCol = (name) => {
      const idx = sHeaders.findIndex(h => String(h).toUpperCase() === name.toUpperCase());
      return idx >= 0 ? idx : sHeaders.findIndex(h => String(h).toUpperCase().includes(name.toUpperCase()));
    };

    const skidIdCol = sCol('Skid_ID');
    const txnIdCol = sCol('TXN_ID');
    const fbpnCol = sCol('FBPN');
    const qtyCol = sCol('Qty_on_Skid');
    const skuCol = sCol('SKU');
    const projectCol = sCol('Project');
    const dateCol = sCol('Date');
    const mfpnCol = sCol('MFPN');

    // Find all rows matching the skid ID
    const searchId = String(skidId).trim().toUpperCase();
    const matchingRows = [];

    for (let i = 1; i < sData.length; i++) {
      const rowSkidId = skidIdCol >= 0 ? String(sData[i][skidIdCol] || '').trim().toUpperCase() : '';
      if (rowSkidId === searchId) {
        matchingRows.push(sData[i]);
      }
    }

    if (matchingRows.length === 0) {
      return { success: false, message: 'Skid not found: ' + skidId };
    }

    // Get TXN_ID from first matching row
    const firstRow = matchingRows[0];
    const txnId = txnIdCol >= 0 ? String(firstRow[txnIdCol] || '').trim() : '';

    // Count total unique skids in this transaction
    const uniqueSkidsInTxn = new Set();
    if (txnId && txnIdCol >= 0 && skidIdCol >= 0) {
      for (let i = 1; i < sData.length; i++) {
        const rowTxn = String(sData[i][txnIdCol] || '').trim();
        const rowSkid = String(sData[i][skidIdCol] || '').trim();
        if (rowTxn === txnId && rowSkid) {
          uniqueSkidsInTxn.add(rowSkid.toUpperCase());
        }
      }
    }

    // Build skid data
    const skidData = {
      skidId: skidId,
      txnId: txnId,
      bolNumber: '',
      dateReceived: dateCol >= 0 ? formatDate(firstRow[dateCol]) : '',
      manufacturer: '',
      project: projectCol >= 0 ? String(firstRow[projectCol] || '') : '',
      pushNumber: '',
      poNumber: '',
      totalSkidsInTxn: uniqueSkidsInTxn.size || 1,
      items: []
    };

    // Lookup BOL, Manufacturer, and Push # from Master_Log using TXN_ID
    if (txnId && masterSheet) {
      const mData = masterSheet.getDataRange().getValues();
      const mHeaders = mData[0];
      const mCol = (name) => {
        const idx = mHeaders.findIndex(h => String(h).toUpperCase() === name.toUpperCase());
        return idx >= 0 ? idx : mHeaders.findIndex(h => String(h).toUpperCase().includes(name.toUpperCase()));
      };

      const mTxnCol = mCol('Txn_ID');
      const mBolCol = mCol('BOL_Number');
      const mMfrCol = mCol('Manufacturer');
      const mPushCol = mCol('Push #');
      const mPoCol = mCol('PO_Number');

      if (mTxnCol >= 0) {
        for (let i = 1; i < mData.length; i++) {
          const rowTxn = String(mData[i][mTxnCol] || '').trim();
          if (rowTxn === txnId) {
            skidData.bolNumber = mBolCol >= 0 ? String(mData[i][mBolCol] || '') : '';
            skidData.manufacturer = mMfrCol >= 0 ? String(mData[i][mMfrCol] || '') : '';
            skidData.pushNumber = mPushCol >= 0 ? String(mData[i][mPushCol] || '') : '';
            skidData.poNumber = mPoCol >= 0 ? String(mData[i][mPoCol] || '') : '';
            break;
          }
        }
      }
    }

    // Aggregate items by FBPN
    const itemsByFbpn = {};
    matchingRows.forEach(row => {
      const fbpn = fbpnCol >= 0 ? String(row[fbpnCol] || '').trim() : '';
      const qty = qtyCol >= 0 ? (Number(row[qtyCol]) || 0) : 0;
      const sku = skuCol >= 0 ? String(row[skuCol] || '') : '';
      const mfpn = mfpnCol >= 0 ? String(row[mfpnCol] || '') : '';

      if (fbpn) {
        if (!itemsByFbpn[fbpn]) {
          itemsByFbpn[fbpn] = { qty: 0, sku: '', mfpn: '' };
        }
        itemsByFbpn[fbpn].qty += qty;
        if (!itemsByFbpn[fbpn].sku && sku) {
          itemsByFbpn[fbpn].sku = sku;
        }
        if (!itemsByFbpn[fbpn].mfpn && mfpn) {
          itemsByFbpn[fbpn].mfpn = mfpn;
        }
      }
    });

    for (const fbpn in itemsByFbpn) {
      skidData.items.push({
        fbpn: fbpn,
        expectedQty: itemsByFbpn[fbpn].qty,
        sku: itemsByFbpn[fbpn].sku,
        mfpn: itemsByFbpn[fbpn].mfpn
      });
    }

    return { success: true, data: skidData };

  } catch (error) {
    Logger.log('Error in getSkidForVerification: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Submit verification results
 * @param {Object} verificationData - The verification data
 * @returns {Object} Result
 */
function submitSkidVerification(verificationData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Verification_Log');

    // Create sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet('Verification_Log');
      logSheet.getRange(1, 1, 1, 14).setValues([[
        'Timestamp', 'BOL_Number', 'PO_Number', 'Manufacturer', 'MFPN', 'FBPN',
        'Expected_Qty', 'Actual_Qty', 'Variance', 'Box_Labels', 'Verified_By',
        'Skid_ID', 'TXN_ID', 'UOM'
      ]]);
      logSheet.getRange(1, 1, 1, 14).setFontWeight('bold');
    }

    // Build UOM lookup map from Item_Master
    const uomMap = {};
    const itemSheet = ss.getSheetByName('Item_Master');
    if (itemSheet && itemSheet.getLastRow() > 1) {
      const itemData = itemSheet.getDataRange().getValues();
      const itemHeaders = itemData[0].map(h => String(h).trim());
      const fbpnIdx = itemHeaders.indexOf('FBPN');
      const uomIdx = itemHeaders.indexOf('UOM');
      if (fbpnIdx > -1 && uomIdx > -1) {
        for (let i = 1; i < itemData.length; i++) {
          const fbpn = String(itemData[i][fbpnIdx] || '').toUpperCase().trim();
          if (fbpn) uomMap[fbpn] = String(itemData[i][uomIdx] || '');
        }
      }
    }

    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail();
    const rows = [];
    let startRow = 2; // Default if inserting at top

    (verificationData.items || []).forEach(item => {
      const fbpnKey = String(item.fbpn || '').toUpperCase().trim();
      const uom = uomMap[fbpnKey] || '';

      rows.push([
        timestamp,
        verificationData.bolNumber || '',
        verificationData.poNumber || '',
        verificationData.manufacturer || '',
        item.mfpn || '',
        item.fbpn || '',
        item.expectedQty || 0,
        item.actualQty || 0,
        item.variance || 0,
        '', // Placeholder for Box_Labels URL
        userEmail,
        verificationData.skidId || '',
        verificationData.txnId || '',
        uom
      ]);
    });

    if (rows.length > 0) {
      // Insert at top (after header)
      logSheet.insertRowsAfter(1, rows.length);
      logSheet.getRange(2, 1, rows.length, 14).setValues(rows);
    }

    return {
      success: true,
      message: 'Verification logged successfully',
      logRow: 2 // Since we insert at the top, the newest is always at row 2
    };

  } catch (error) {
    Logger.log('Error in submitSkidVerification: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Get recent verification log entries
 * @returns {Object} Result with log entries
 */
function getVerificationLog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Verification_Log');

    if (!logSheet) {
      return { success: true, entries: [] };
    }

    const data = logSheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, entries: [] };
    }

    const entries = [];
    const maxEntries = Math.min(100, data.length - 1);

    for (let i = 1; i <= maxEntries; i++) {
      const row = data[i];
      entries.push({
        timestamp: row[0] ? formatDate(row[0]) + ' ' + formatTime_(row[0]) : '',
        verifiedBy: row[1] || '',
        skidId: row[2] || '',
        txnId: row[3] || '',
        bolNumber: row[4] || '',
        fbpn: row[5] || '',
        expectedQty: row[6] || 0,
        actualQty: row[7] || 0,
        variance: row[8] || 0
      });
    }

    return { success: true, entries: entries };

  } catch (error) {
    Logger.log('Error in getVerificationLog: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Format time helper
 */
function formatTime_(date) {
  if (!date) return '';
  try {
    const d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
  } catch (e) {
    return '';
  }
}

/**
 * Generate box labels after verification
 * Creates 4x2 labels - one per box based on boxCount
 * Layout matches inbound skid labels (scaled for 4x2)
 * Barcode is Skid ID for scanning
 * @param {Object} labelRequest - Contains skidId, items with boxCount/qtyPerBox, and metadata
 * @returns {Object} Result with PDF URL
 */
function generateBoxLabelsAfterVerification(labelRequest) {
  try {
    const skidId = labelRequest.skidId || '';
    const txnId = labelRequest.txnId || '';
    const bolNumber = labelRequest.bolNumber || '';
    const manufacturer = labelRequest.manufacturer || '';
    const project = labelRequest.project || '';
    const pushNumber = labelRequest.pushNumber || '';
    const poNumber = labelRequest.poNumber || '';
    const items = labelRequest.items || [];

    if (items.length === 0) {
      return { success: false, message: 'No items to generate labels for' };
    }

    // Calculate total labels to generate (one per box)
    let totalLabels = 0;
    items.forEach(item => {
      const bCount = (item.boxes && item.boxes.length > 0) ? item.boxes.length : (Number(item.boxCount) || 0);
      totalLabels += bCount;
    });

    if (totalLabels === 0) {
      return { success: false, message: 'No boxes to label (box count is 0)' };
    }

    // Generate Skid ID barcode (same for all boxes from this skid)
    let skidBarcodeUri = '';
    try {
      skidBarcodeUri = bwipPngDataUri_('code128', skidId, { scale: 3, height: 10 });
    } catch (e) {
      Logger.log('Barcode generation failed for skidId: ' + e);
    }

    // Build HTML for 4x2 labels - matching inbound layout but scaled
    let boxCounter = 0;
    let labelsHtml = '';

    items.forEach(item => {
      const fbpn = item.fbpn || '';
      const sku = item.sku || '';
      const mfpn = item.mfpn || '';

      // Support either boxes array (with individual qtys) or boxCount/qtyPerBox
      const boxes = item.boxes || [];
      const boxCount = boxes.length > 0 ? boxes.length : (Number(item.boxCount) || 0);
      const defaultQty = Number(item.qtyPerBox) || 0;

      // Generate SKU barcode for top-right
      let skuBarcodeUri = '';
      try {
        skuBarcodeUri = bwipPngDataUri_('code128', sku || fbpn, { scale: 2, height: 8 });
      } catch (e) {
        Logger.log('SKU Barcode generation failed: ' + e);
      }

      // Create one label per box
      for (let i = 0; i < boxCount; i++) {
        boxCounter++;
        // Get qty for this specific box (from boxes array) or use default
        const boxQty = boxes.length > i ? (Number(boxes[i]) || defaultQty) : defaultQty;

        labelsHtml += `
    <div class="label">
      <div class="top-row">
        <div class="top-left">
          <div class="manufacturer">${escapeHtml_(manufacturer)}</div>
          <div class="push-line">PUSH #: ${escapeHtml_(pushNumber)}</div>
          <div class="po-line">PO #: ${escapeHtml_(poNumber)}</div>
          <div class="box-count-line">BOX ${boxCounter} of ${totalLabels}</div>
        </div>
        <div class="top-right">
          <div class="top-barcode"><img src="${skuBarcodeUri}" class="top-barcode-img"></div>
          <div class="sku-text">${escapeHtml_(sku || fbpn)}</div>
        </div>
      </div>
      <div class="middle-block">
        <div class="line-fbpn">FBPN: ${escapeHtml_(fbpn)}</div>
        <div class="line-qty">Qty: ${boxQty}</div>
        <div class="line-project">PROJECT: ${escapeHtml_(project)}</div>
        <div class="line-mfpn">MFPN: ${escapeHtml_(mfpn)}</div>
      </div>
      <div class="bottom-block">
        <div class="scan-text">Scan Skid ID</div>
        <div class="bottom-barcode"><img src="${skidBarcodeUri}" class="bottom-barcode-img"></div>
        <div class="skid-id-text">${escapeHtml_(skidId)}</div>
      </div>
    </div>`;
      }
    });

    // Full HTML document - 4x2 label format matching inbound style
    const fullHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Box Labels - ${escapeHtml_(skidId)}</title>
  <style>
    @page { size: 4in 2in; margin: 0; }
    * { box-sizing: border-box; }
    body { margin: 0; padding: 0; font-family: Arial, sans-serif; font-weight: bold; }
    
    .label { 
      width: 4in; 
      height: 2in; 
      padding: 0.06in 0.12in 0.05in 0.12in; 
      page-break-after: always; 
      border: 1px solid #000; 
      display: flex; 
      flex-direction: column; 
    }
    .label:last-child { page-break-after: auto; }
    
    .top-row { display: flex; justify-content: space-between; align-items: flex-start; }
    .top-left { flex: 1.4; }
    .manufacturer { font-size: 9pt; line-height: 1.05; text-transform: uppercase; }
    .push-line { margin-top: 0.01in; font-size: 10pt; }
    .po-line { margin-top: 0.01in; font-size: 8pt; }
    .box-count-line { margin-top: 0.01in; font-size: 8pt; color: #0066cc; }
    
    .top-right { flex: 1; text-align: center; }
    .top-barcode { width: 100%; height: 0.35in; margin-bottom: 0.01in; }
    .top-barcode-img { width: 100%; height: 100%; object-fit: contain; display: block; }
    .sku-text { font-size: 7pt; text-align: center; }
    
    .middle-block { margin-top: 0.03in; }
    .line-fbpn { font-size: 18pt; line-height: 1.0; }
    .line-qty { font-size: 14pt; line-height: 1.0; margin-top: 0.02in; }
    .line-project { font-size: 9pt; line-height: 1.0; margin-top: 0.02in; }
    .line-mfpn { font-size: 8pt; line-height: 1.0; margin-top: 0.01in; color: #555; }
    
    .bottom-block { margin-top: auto; text-align: center; padding-top: 0.02in; width: 100%; }
    .scan-text { font-size: 6pt; margin-bottom: 0.02in; }
    .bottom-barcode { width: 100%; height: 0.32in; padding: 2px 15px; overflow: hidden; display: flex; justify-content: center; }
    .bottom-barcode-img { width: 100%; height: 100%; object-fit: contain; display: block; }
    .skid-id-text { font-size: 6pt; margin-top: 1px; }
  </style>
</head>
<body>
  ${labelsHtml}
</body>
</html>`;

    // Get or create destination folder - same structure as inbound labels
    let targetFolder;
    try {
      if (bolNumber) {
        targetFolder = createInboundFolder_(new Date(), bolNumber);
      }
    } catch (e) {
      Logger.log('Could not get inbound folder: ' + e);
    }

    if (!targetFolder) {
      const rootFolders = DriveApp.getFoldersByName('Inbound_Labels');
      if (rootFolders.hasNext()) {
        targetFolder = rootFolders.next();
      } else {
        targetFolder = DriveApp.createFolder('Inbound_Labels');
      }
    }

    // Save like other inbound labels
    const safeBol = String(bolNumber || 'NO_BOL').trim().replace(/[\/\\?%*:|"<>\.]/g, '_');
    const safeSkid = String(skidId || 'SKID').trim().replace(/[\/\\?%*:|"<>\.]/g, '_');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const baseName = `BoxLabels_${safeSkid}_${timestamp}`;

    // Create HTML file
    const htmlBlob = Utilities.newBlob(fullHtml, 'text/html', `${baseName}.html`);
    const htmlFile = targetFolder.createFile(htmlBlob);

    // Create PDF file
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName(`${baseName}.pdf`);
    const pdfFile = targetFolder.createFile(pdfBlob);
    const downloadUrl = pdfFile.getUrl();

    // If a log row index was passed, update the Verification_Log with the URL
    // The Box_Labels column is index 10 (J column)
    if (labelRequest.logRow) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Verification_Log');
        if (sheet) {
          // If we verified multiple items, we created multiple rows.
          // But usually we just update the first row corresponding to this skid/verification call.
          // Or update all rows created? The row index passed back was the *first* row.
          // Let's assume we just update the row specified.
          sheet.getRange(labelRequest.logRow, 10).setValue(downloadUrl);
        }
      } catch (e) {
        Logger.log('Failed to update log with URL: ' + e);
      }
    }

    return {
      success: true,
      pdfUrl: downloadUrl,
      htmlUrl: htmlFile.getUrl(),
      totalLabels: totalLabels,
      message: `Generated ${totalLabels} box labels`
    };

  } catch (error) {
    Logger.log('Error in generateBoxLabelsAfterVerification: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ----------------------------------------------------------------------------
// DATA ENTRY FUNCTIONS - Item_Master and PO_Master
// ----------------------------------------------------------------------------

/**
 * Adds a new item to Item_Master sheet.
 * Expected columns: FBPN, MFPN, Description, UOM, Manufacturer
 */
function addItemToItemMaster(data) {
  try {
    if (!data || !data.fbpn) {
      return { success: false, message: 'FBPN is required.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TABS.ITEM_MASTER || 'Item_Master');

    if (!sheet) {
      return { success: false, message: 'Item_Master sheet not found.' };
    }

    // Get headers
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = (name) => {
      const idx = headers.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
      return idx >= 0 ? idx : -1;
    };

    // Check for duplicate FBPN
    const fbpnCol = colIndex('FBPN');
    if (fbpnCol >= 0 && sheet.getLastRow() > 1) {
      const existingFBPNs = sheet.getRange(2, fbpnCol + 1, sheet.getLastRow() - 1, 1).getValues();
      const fbpnUpper = String(data.fbpn).toUpperCase().trim();
      for (let i = 0; i < existingFBPNs.length; i++) {
        if (String(existingFBPNs[i][0]).toUpperCase().trim() === fbpnUpper) {
          return { success: false, message: 'FBPN already exists in Item_Master.' };
        }
      }
    }

    // Build row based on headers
    const newRow = new Array(headers.length).fill('');

    const colMap = {
      'fbpn': data.fbpn,
      'mfpn': data.mfpn || '',
      'description': data.description || '',
      'uom': data.uom || '',
      'manufacturer': data.manufacturer || ''
    };

    Object.keys(colMap).forEach(key => {
      const idx = colIndex(key);
      if (idx >= 0) {
        newRow[idx] = colMap[key];
      }
    });

    // Add the row
    sheet.appendRow(newRow);
    const newRowNum = sheet.getLastRow();

    Logger.log('Added new item to Item_Master: ' + data.fbpn + ' at row ' + newRowNum);

    return { success: true, row: newRowNum, message: 'Item added successfully.' };

  } catch (err) {
    Logger.log('Error in addItemToItemMaster: ' + err.toString());
    return { success: false, message: err.message };
  }
}

/**
 * Adds a new PO and Project to PO_Master sheet.
 * Expected columns: Customer_PO (A), Project (B), Customer (C), Notes (D)
 */
function addPOToPOMaster(data) {
  try {
    if (!data || !data.customerPO) {
      return { success: false, message: 'Customer PO Number is required.' };
    }
    if (!data.project) {
      return { success: false, message: 'Project is required.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TABS.PO_MASTER || 'PO_Master');

    if (!sheet) {
      return { success: false, message: 'PO_Master sheet not found.' };
    }

    // Check for duplicate PO
    if (sheet.getLastRow() > 1) {
      const existingPOs = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      const poUpper = String(data.customerPO).toUpperCase().trim();
      for (let i = 0; i < existingPOs.length; i++) {
        if (String(existingPOs[i][0]).toUpperCase().trim() === poUpper) {
          return { success: false, message: 'Customer PO already exists in PO_Master.' };
        }
      }
    }

    // Build row: [Customer_PO, Project, Customer, Notes]
    const newRow = [
      data.customerPO,
      data.project,
      data.customer || '',
      data.notes || ''
    ];

    // Add the row
    sheet.appendRow(newRow);
    const newRowNum = sheet.getLastRow();

    Logger.log('Added new PO to PO_Master: ' + data.customerPO + ' at row ' + newRowNum);

    return { success: true, row: newRowNum, message: 'PO added successfully.' };

  } catch (err) {
    Logger.log('Error in addPOToPOMaster: ' + err.toString());
    return { success: false, message: err.message };
  }
}

/**
 * Gets recent POs from PO_Master for display in the modal.
 */
function getRecentPOs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TABS.PO_MASTER || 'PO_Master');

    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

    return data.map(row => ({
      po: row[0] || '',
      project: row[1] || ''
    }));

  } catch (err) {
    Logger.log('Error in getRecentPOs: ' + err.toString());
    return [];
  }
}
