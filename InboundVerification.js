/**
 * ============================================================================
 * INBOUND VERIFICATION SYSTEM
 * ============================================================================
 * 
 * Handles skid content verification, box counting, variance logging,
 * and 4x2 label generation.
 */

/* ----------------------------------------------------------------------------
 * 1. DATA LOOKUP & RETRIEVAL
 * ------------------------------------------------------------------------- */

function getVerificationSkidData(skidId) {
    if (!skidId) throw new Error("Skid ID is required.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const skidsSheet = ss.getSheetByName(TABS.INBOUND_SKIDS);
    const masterSheet = ss.getSheetByName(TABS.MASTER_LOG);
    const stagingSheet = ss.getSheetByName(TABS.INBOUND_STAGING);

    if (!skidsSheet) throw new Error("Inbound_Skids sheet not found.");

    const sData = skidsSheet.getDataRange().getValues();
    const sHeaders = sData[0];
    const sIdx = (h) => sHeaders.indexOf(h);

    // Find skid row
    let skidRow = null;
    const skidIdIdx = sIdx('Skid_ID');
    if (skidIdIdx === -1) throw new Error("Skid_ID column not found in Inbound_Skids.");

    for (let i = 1; i < sData.length; i++) {
        if (String(sData[i][skidIdIdx]).trim().toUpperCase() === String(skidId).trim().toUpperCase()) {
            skidRow = sData[i];
            break;
        }
    }

    if (!skidRow) throw new Error(`Skid ID '${skidId}' not found.`);

    // Extract info
    const skidInfo = {
        skidId: skidRow[skidIdIdx],
        txnId: sIdx('TXN_ID') > -1 ? skidRow[sIdx('TXN_ID')] : '',
        fbpn: sIdx('FBPN') > -1 ? skidRow[sIdx('FBPN')] : '',
        project: sIdx('Project') > -1 ? skidRow[sIdx('Project')] : '',
        expectedQty: sIdx('Qty_on_Skid') > -1 ? Number(skidRow[sIdx('Qty_on_Skid')]) : 0,
        sku: sIdx('SKU') > -1 ? skidRow[sIdx('SKU')] : '',
        manufacturer: '', // Will fetch below
        pushNumber: '',   // Will fetch below
        uom: 'EA',        // Will fetch from Item_Master
        assetType: '',    // Will fetch from Item_Master
        mfpn: sIdx('MFPN') > -1 ? skidRow[sIdx('MFPN')] : ''
    };

    // Fetch Manufacturer (from Staging or Item Master, but Staging is linked to SkidID or Push)
    // Fetch Push Number (from Staging is easiest if SkidId is there)

    // Try finding in Staging by Skid_ID first
    if (stagingSheet) {
        const stData = stagingSheet.getDataRange().getValues();
        const stHeaders = stData[0];
        const stSkidIdIdx = stHeaders.indexOf('Skid_ID');
        const stPushIdx = stHeaders.indexOf('Push_Number');
        const stManIdx = stHeaders.indexOf('Manufacturer');

        if (stSkidIdIdx > -1) {
            for (let i = 1; i < stData.length; i++) {
                if (String(stData[i][stSkidIdIdx]).trim().toUpperCase() === String(skidId).trim().toUpperCase()) {
                    if (stPushIdx > -1) skidInfo.pushNumber = stData[i][stPushIdx];
                    if (stManIdx > -1) skidInfo.manufacturer = stData[i][stManIdx];
                    break;
                }
            }
        }
    }

    // If Manufacturer still missing, look in Master Log via TXN_ID
    if ((!skidInfo.manufacturer || !skidInfo.pushNumber) && masterSheet && skidInfo.txnId) {
        const mData = masterSheet.getDataRange().getValues();
        const mHeaders = mData[0];
        const mTxnIdx = mHeaders.indexOf('Txn_ID');
        const mManIdx = mHeaders.indexOf('Manufacturer');
        const mPushIdx = mHeaders.indexOf('Push #');

        if (mTxnIdx > -1) {
            for (let i = 1; i < mData.length; i++) {
                if (String(mData[i][mTxnIdx]) === String(skidInfo.txnId)) {
                    if (!skidInfo.manufacturer && mManIdx > -1) skidInfo.manufacturer = mData[i][mManIdx];
                    if (!skidInfo.pushNumber && mPushIdx > -1) skidInfo.pushNumber = mData[i][mPushIdx];
                    break;
                }
            }
        }
    }

    // Look up UOM and Asset_Type from Item_Master based on SKU
    // Use getItemMasterDetails function from IMS_Inbound.js if available, otherwise inline lookup
    if (typeof getItemMasterDetails === 'function') {
        const itemDetails = getItemMasterDetails(skidInfo.sku, skidInfo.fbpn);
        skidInfo.uom = itemDetails.uom || 'EA';
        skidInfo.assetType = itemDetails.assetType || '';
    } else {
        // Inline lookup from Item_Master
        try {
            const itemSheet = ss.getSheetByName(TABS.ITEM_MASTER || 'Item_Master');
            if (itemSheet) {
                const itemData = itemSheet.getDataRange().getValues();
                const itemHeaders = itemData[0].map(h => String(h).trim());
                const skuIdx = itemHeaders.indexOf('SKU');
                const fbpnIdx = itemHeaders.indexOf('FBPN');
                const uomIdx = itemHeaders.indexOf('UOM');
                const assetTypeIdx = itemHeaders.indexOf('Asset_Type');

                // Try to match by SKU first
                let found = false;
                if (skuIdx > -1 && skidInfo.sku) {
                    const skuUpper = String(skidInfo.sku).toUpperCase().trim();
                    for (let i = 1; i < itemData.length; i++) {
                        const rowSku = String(itemData[i][skuIdx] || '').toUpperCase().trim();
                        if (rowSku === skuUpper) {
                            skidInfo.uom = uomIdx > -1 ? String(itemData[i][uomIdx] || 'EA') : 'EA';
                            skidInfo.assetType = assetTypeIdx > -1 ? String(itemData[i][assetTypeIdx] || '') : '';
                            found = true;
                            break;
                        }
                    }
                }
                // Fallback: match by FBPN
                if (!found && fbpnIdx > -1 && skidInfo.fbpn) {
                    const fbpnUpper = String(skidInfo.fbpn).toUpperCase().trim();
                    for (let i = 1; i < itemData.length; i++) {
                        const rowFbpn = String(itemData[i][fbpnIdx] || '').toUpperCase().trim();
                        if (rowFbpn === fbpnUpper) {
                            skidInfo.uom = uomIdx > -1 ? String(itemData[i][uomIdx] || 'EA') : 'EA';
                            skidInfo.assetType = assetTypeIdx > -1 ? String(itemData[i][assetTypeIdx] || '') : '';
                            break;
                        }
                    }
                }
            }
        } catch (e) {
            Logger.log('Error looking up Item_Master details: ' + e.toString());
        }
    }

    return skidInfo;
}

/* ----------------------------------------------------------------------------
 * 2. SUBMISSION & LOGGING
 * ------------------------------------------------------------------------- */

function submitVerification_backend(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Ensure Verification_Log sheet exists
    const logSheetName = (typeof TABS !== 'undefined' && TABS.VERIFICATION_LOG) ? TABS.VERIFICATION_LOG : 'Verification_Log';
    let logSheet = ss.getSheetByName(logSheetName);

    if (!logSheet) {
        logSheet = ss.insertSheet(logSheetName);
        // Use HEADERS from IMS_Config if available
        const verificationHeaders = (typeof HEADERS !== 'undefined' && HEADERS['Verification_Log']) 
            ? HEADERS['Verification_Log'] 
            : [
                'Timestamp', 'BOL_Number', 'PO_Number', 'Asset_Type', 'Manufacturer',
                'MFPN', 'FBPN', 'UOM', 'Expected_Qty', 'Actual_Qty',
                'Variance', 'Box_Labels', 'Verified_By', 'Skid_ID', 'TXN_ID'
              ];
        logSheet.appendRow(verificationHeaders);
        logSheet.setFrozenRows(1);
    }

    const user = Session.getActiveUser().getEmail();
    const timestamp = new Date();

    const variance = payload.actualQty - payload.expectedQty;

    // Look up UOM and Asset_Type from Item_Master based on SKU if not provided in payload
    let resolvedUom = payload.uom || 'EA';
    let resolvedAssetType = payload.assetType || '';
    
    if ((!payload.uom || !payload.assetType) && (payload.sku || payload.fbpn)) {
        // Use getItemMasterDetails if available, otherwise inline lookup
        if (typeof getItemMasterDetails === 'function') {
            const itemDetails = getItemMasterDetails(payload.sku, payload.fbpn);
            resolvedUom = payload.uom || itemDetails.uom || 'EA';
            resolvedAssetType = payload.assetType || itemDetails.assetType || '';
        } else {
            // Inline lookup from Item_Master
            try {
                const itemSheet = ss.getSheetByName(TABS.ITEM_MASTER || 'Item_Master');
                if (itemSheet) {
                    const itemData = itemSheet.getDataRange().getValues();
                    const itemHeaders = itemData[0].map(h => String(h).trim());
                    const skuIdx = itemHeaders.indexOf('SKU');
                    const fbpnIdx = itemHeaders.indexOf('FBPN');
                    const uomIdx = itemHeaders.indexOf('UOM');
                    const assetTypeIdx = itemHeaders.indexOf('Asset_Type');

                    let found = false;
                    // Try to match by SKU first
                    if (skuIdx > -1 && payload.sku) {
                        const skuUpper = String(payload.sku).toUpperCase().trim();
                        for (let i = 1; i < itemData.length; i++) {
                            const rowSku = String(itemData[i][skuIdx] || '').toUpperCase().trim();
                            if (rowSku === skuUpper) {
                                resolvedUom = payload.uom || (uomIdx > -1 ? String(itemData[i][uomIdx] || 'EA') : 'EA');
                                resolvedAssetType = payload.assetType || (assetTypeIdx > -1 ? String(itemData[i][assetTypeIdx] || '') : '');
                                found = true;
                                break;
                            }
                        }
                    }
                    // Fallback: match by FBPN
                    if (!found && fbpnIdx > -1 && payload.fbpn) {
                        const fbpnUpper = String(payload.fbpn).toUpperCase().trim();
                        for (let i = 1; i < itemData.length; i++) {
                            const rowFbpn = String(itemData[i][fbpnIdx] || '').toUpperCase().trim();
                            if (rowFbpn === fbpnUpper) {
                                resolvedUom = payload.uom || (uomIdx > -1 ? String(itemData[i][uomIdx] || 'EA') : 'EA');
                                resolvedAssetType = payload.assetType || (assetTypeIdx > -1 ? String(itemData[i][assetTypeIdx] || '') : '');
                                break;
                            }
                        }
                    }
                }
            } catch (e) {
                Logger.log('Error looking up Item_Master details in submitVerification: ' + e.toString());
            }
        }
    }

    // Get existing headers from the sheet to build row dynamically
    const existingHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    const hIdx = (name) => existingHeaders.indexOf(name);
    
    // Build row based on actual headers in sheet
    const row = new Array(existingHeaders.length).fill('');
    
    if (hIdx('Timestamp') > -1) row[hIdx('Timestamp')] = timestamp;
    if (hIdx('BOL_Number') > -1) row[hIdx('BOL_Number')] = payload.bolNumber || '';
    if (hIdx('PO_Number') > -1) row[hIdx('PO_Number')] = payload.poNumber || '';
    if (hIdx('Asset_Type') > -1) row[hIdx('Asset_Type')] = resolvedAssetType;
    if (hIdx('Manufacturer') > -1) row[hIdx('Manufacturer')] = payload.manufacturer || '';
    if (hIdx('MFPN') > -1) row[hIdx('MFPN')] = payload.mfpn || '';
    if (hIdx('FBPN') > -1) row[hIdx('FBPN')] = payload.fbpn || '';
    if (hIdx('UOM') > -1) row[hIdx('UOM')] = resolvedUom;
    if (hIdx('Expected_Qty') > -1) row[hIdx('Expected_Qty')] = payload.expectedQty || 0;
    if (hIdx('Actual_Qty') > -1) row[hIdx('Actual_Qty')] = payload.actualQty || 0;
    if (hIdx('Variance') > -1) row[hIdx('Variance')] = variance;
    if (hIdx('Box_Labels') > -1) row[hIdx('Box_Labels')] = payload.boxCount || 0;
    if (hIdx('Verified_By') > -1) row[hIdx('Verified_By')] = user;
    if (hIdx('Skid_ID') > -1) row[hIdx('Skid_ID')] = payload.skidId || '';
    if (hIdx('TXN_ID') > -1) row[hIdx('TXN_ID')] = payload.txnId || '';

    logSheet.appendRow(row);

    // Generate Labels if requested
    let labelUrl = '';
    let labelHtmlUrl = '';

    if (payload.generateLabels && payload.boxes && payload.boxes.length > 0) {
        try {
            const labelRes = generateVerificationBoxLabels(payload.boxes);
            if (labelRes.success) {
                labelUrl = labelRes.pdfUrl;
                labelHtmlUrl = labelRes.htmlUrl;
            }
        } catch (e) {
            Logger.log("Error generating labels: " + e.toString());
        }
    }

    return {
        success: true,
        variance: variance,
        status: status,
        labelUrl: labelUrl,
        labelHtmlUrl: labelHtmlUrl
    };
}

/* ----------------------------------------------------------------------------
 * 3. LABEL GENERATION (4x2)
 * ------------------------------------------------------------------------- */

function generateVerificationBoxLabels(boxesData) {
    // boxesData = array of { skidId, fbpn, manufacturer, project, qty, containerId, pushNumber, uom }

    try {
        const html = generateBoxLabelsHtml_(boxesData);

        // Save to "Verification Labels" folder
        const parentId = (typeof FOLDERS !== 'undefined' && FOLDERS.INBOUND_UPLOADS) ? FOLDERS.INBOUND_UPLOADS : DriveApp.getRootFolder().getId();
        let rootFolder;
        try { rootFolder = DriveApp.getFolderById(parentId); } catch (e) { rootFolder = DriveApp.getRootFolder(); }

        const folder = getOrCreateSubfolder_(rootFolder, "Verification_Labels");
        const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
        const fileName = `Labels_${boxesData[0].skidId}_${dateStr}`;

        const res = saveLabelsToDrive(html, folder, fileName); // Assuming saveLabelsToDrive exists and creates PDF

        return { success: true, pdfUrl: res.pdfFile.getUrl(), htmlUrl: res.htmlFile.getUrl() };

    } catch (e) {
        return { success: false, error: e.toString() };
    }
}

function generateBoxLabelsHtml_(boxes) {
    // Uses BWIP-JS for barcodes via Google API or pure JS if included. 
    // Assuming the existing label generator uses a helper `bwipPngDataUri_`.

    let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    @page { size: 4in 2in; margin: 0; }
    * { box-sizing: border-box; }
    body { margin: 0; padding: 0; font-family: 'Arial', sans-serif; }
    
    .label { 
      width: 4in; 
      height: 2in; 
      page-break-after: always; 
      position: relative;
      padding: 0.1in;
      display: flex;
      flex-direction: column;
      border: 1px dotted #ccc; /* Helper border, remove for production if needed */
    }
    
    .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      height: 0.5in;
      border-bottom: 1px solid #000;
      padding-bottom: 2px;
    }
    
    .top-barcode {
      width: 1.8in;
      height: 0.45in;
    }
    .top-barcode img { width: 100%; height: 100%; object-fit: contain; }
    
    .container-id-text {
      font-size: 10pt;
      font-weight: bold;
      text-align: right;
    }
    
    .main-content {
      flex: 1;
      display: flex;
      flex-direction: column;
      justify-content: center;
      padding: 0.05in 0; 
    }
    
    .row { display: flex; justify-content: space-between; margin-bottom: 2px; }
    .label-key { font-size: 8pt; color: #444; }
    .label-val { font-size: 11pt; font-weight: bold; }
    .fbpn { font-size: 16pt; font-weight: 900; }
    .qty { font-size: 18pt; font-weight: bold; }
    
    .footer {
      height: 0.5in;
      border-top: 1px solid #000;
      padding-top: 2px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    
    .bottom-barcode {
      width: 2.0in;
      height: 0.45in;
    }
    .bottom-barcode img { width: 100%; height: 100%; object-fit: contain; }
    
    .skid-id-text {
      font-size: 9pt;
      font-weight: bold;
    }
  </style></head><body>`;

    boxes.forEach(box => {
        // Top Barcode: Container ID
        const topUri = bwipPngDataUri_('code128', box.containerId, { scale: 2, height: 10, includeText: false });

        // Bottom Barcode: Skid ID
        const bottomUri = bwipPngDataUri_('code128', box.skidId, { scale: 2, height: 10, includeText: false });

        html += `
    <div class="label">
      <div class="header">
        <div class="top-barcode"><img src="${topUri}"></div>
        <div class="container-id-text">${escapeHtml_(box.containerId)}</div>
      </div>
      
      <div class="main-content">
        <div class="row">
           <div class="fbpn">${escapeHtml_(box.fbpn)}</div>
        </div>
        <div class="row">
           <div><span class="label-key">MFR:</span> <span class="label-val">${escapeHtml_(box.manufacturer || '')}</span></div>
           <div><span class="label-key">PUSH:</span> <span class="label-val">${escapeHtml_(box.pushNumber || '')}</span></div>
        </div>
        <div class="row">
           <div><span class="label-key">PROJ:</span> <span class="label-val">${escapeHtml_(box.project || '')}</span></div>
        </div>
        <div class="row" style="margin-top: 5px;">
           <div class="qty">QTY: ${escapeHtml_(box.qty)} <span style="font-size:12pt">${escapeHtml_(box.uom || 'EA')}</span></div>
        </div>
      </div>
      
      <div class="footer">
        <div class="skid-id-text">Ref Skid: ${escapeHtml_(box.skidId)}</div>
        <div class="bottom-barcode"><img src="${bottomUri}"></div>
      </div>
    </div>`;
    });

    html += `</body></html>`;
    return html;
}

// Reusing helper from IMS_Inbound if available, otherwise redefine here
function bwipPngDataUri_(symbology, text, opts) {
    // Use existing helper if available, else simple shim (this assumes the helper is global)
    // If not global, we would need to duplicate the logic.
    // Assuming 'bwipPngDataUri_' exists in the project scope (it's likely in IMS_Config or similar).
    if (typeof bwipPngDataUri_ !== 'undefined') {
        return bwipPngDataUri_(symbology, text, opts);
    }
    // Fallback if needed (unlikely based on project structure)
    return `https://bwipjs-api.metafloor.com/?bcid=${symbology}&text=${encodeURIComponent(text)}&scale=2`;
}


/* ----------------------------------------------------------------------------
 * 4. REPORTING
 * ------------------------------------------------------------------------- */

function generateVerificationReport(params) {
    // params: { startDate, endDate, frequency }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(TABS.VERIFICATION_LOG || 'Verification_Log');

    if (!logSheet) return { success: false, message: 'Verification_Log sheet not found.' };

    const data = logSheet.getDataRange().getValues();
    const headers = data[0];

    const idxDate = headers.indexOf('Timestamp');
    const idxSkid = headers.indexOf('Skid_ID');
    const idxFbpn = headers.indexOf('FBPN');
    const idxUser = headers.indexOf('User');
    const idxExp = headers.indexOf('Expected_Qty');
    const idxAct = headers.indexOf('Actual_Qty_Total');
    const idxVar = headers.indexOf('Variance');
    const idxStatus = headers.indexOf('Status');

    const start = new Date(params.startDate);
    start.setHours(0, 0, 0, 0);
    const end = new Date(params.endDate);
    end.setHours(23, 59, 59, 999);

    const rows = [];

    for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const d = new Date(r[idxDate]);
        if (d >= start && d <= end) {
            rows.push({
                date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm'),
                skid: r[idxSkid],
                fbpn: r[idxFbpn],
                user: r[idxUser],
                exp: r[idxExp],
                act: r[idxAct],
                variance: r[idxVar],
                status: r[idxStatus]
            });
        }
    }

    if (rows.length === 0) return { success: false, message: 'No records found for this period.' };

    // Build HTML Report
    let html = `
  <!DOCTYPE html>
  <html>
  <head>
    <style>
      body { font-family: Arial, sans-serif; font-size: 10pt; }
      h1 { color: #333; }
      .meta { margin-bottom: 20px; color: #666; }
      table { width: 100%; border-collapse: collapse; margin-top: 10px; }
      th { background: #f4f4f4; text-align: left; padding: 6px; border-bottom: 2px solid #ddd; }
      td { padding: 6px; border-bottom: 1px solid #eee; }
      .status-MATCH { color: green; font-weight: bold; }
      .status-OVERAGE { color: orange; font-weight: bold; }
      .status-SHORTAGE { color: red; font-weight: bold; }
      .variance-positive { color: orange; }
      .variance-negative { color: red; }
      .total-row { font-weight: bold; background: #fafafa; }
    </style>
  </head>
  <body>
    <h1>Inbound Verification Report</h1>
    <div class='meta'>
       Period: ${params.startDate} to ${params.endDate}<br>
       Generated: ${new Date().toLocaleString()}
    </div>
    
    <table>
      <thead>
        <tr>
          <th>Date</th>
          <th>User</th>
          <th>Skid ID</th>
          <th>FBPN</th>
          <th>Expected</th>
          <th>Actual</th>
          <th>Variance</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>`;

    rows.forEach(r => {
        let varClass = '';
        if (r.variance > 0) varClass = 'variance-positive';
        if (r.variance < 0) varClass = 'variance-negative';

        html += `<tr>
       <td>${r.date}</td>
       <td>${r.user}</td>
       <td>${r.skid}</td>
       <td>${r.fbpn}</td>
       <td>${r.exp}</td>
       <td>${r.act}</td>
       <td class='${varClass}'>${r.variance > 0 ? '+' : ''}${r.variance}</td>
       <td class='status-${r.status}'>${r.status}</td>
    </tr>`;
    });

    html += `</tbody></table></body></html>`;

    try {
        const blob = Utilities.newBlob(html, 'text/html', 'Verification_Report.html');
        const pdf = blob.getAs('application/pdf').setName(`Verification_Report_${params.endDate}.pdf`);

        // Save to Reports Folder (using helper if available, else root)
        let file;
        if (typeof saveReportToFolder === 'function') {
            file = saveReportToFolder(pdf, 'Verification', params.frequency);
        } else {
            file = DriveApp.createFile(pdf);
        }

        return {
            success: true,
            url: file.getUrl(),
            name: file.getName(),
            rowCount: rows.length
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

