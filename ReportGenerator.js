// ============================================================================
// PROJECT & ASSET RESOLUTION
// ============================================================================

function buildPoToProjectMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('PO_Master');
  const map = {};
  if (!sh) return map;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return map;

  const headers = values[0].map(h => (h || '').toString().trim());
  const poIdx = headers.indexOf('Customer_PO');
  const projIdx = headers.indexOf('Project');
  if (poIdx === -1 || projIdx === -1) return map;

  for (let r = 2; r < values.length; r++) { 
    const po = (values[r][poIdx] || '').toString().trim();
    const proj = (values[r][projIdx] || '').toString().trim();
    if (po) map[po] = proj;
  }
  return map;
}

function resolveProject_(rowProject, poNumber, poMap) {
  const p = (rowProject || '').toString().trim();
  if (p) return p;
  const po = (poNumber || '').toString().trim();
  if (!po) return '';
  return (poMap && poMap[po]) ? poMap[po] : '';
}

function buildItemAssetTypeMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Item_Master');
  const map = {};
  if (!sh) return map;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return map;

  const headers = values[0].map(h => (h || '').toString().trim());
  const fbpnIdx = headers.indexOf('FBPN');
  const assetIdx = headers.indexOf('Asset Type') !== -1 ? headers.indexOf('Asset Type') : headers.indexOf('Asset_Type');
  if (fbpnIdx === -1 || assetIdx === -1) return map;

  for (let r = 2; r < values.length; r++) {
    const fbpn = (values[r][fbpnIdx] || '').toString().trim();
    const asset = (values[r][assetIdx] || '').toString().trim();
    if (fbpn) map[fbpn] = asset;
  }
  return map;
}

function resolveAssetType_(rowAssetType, fbpn, assetMap) {
  const a = (rowAssetType || '').toString().trim();
  if (a) return a;
  const key = (fbpn || '').toString().trim();
  if (!key) return '';
  return (assetMap && assetMap[key]) ? assetMap[key] : '';
}

// ============================================================================
// MAIN GENERATION FUNCTIONS
// ============================================================================

function generateInboundReport(params) {
  try {
    const reportData = getInboundReportData(params);
    if (!reportData.rows || reportData.rows.length === 0) {
      return { success: false, message: 'No inbound data found.' };
    }

    const html = buildInboundReportHtml_(params, reportData);
    const ts = formatDateForFilename(new Date());
    const filenameBase = `${params.frequency}_Inbound_${ts}`;

    const pdfFile = saveHtmlAsPdfToReports_(html, `${filenameBase}.pdf`, 'Inbound', params.frequency);
    const xlsxFile = saveRowsAsXlsxToReports_(reportData.rows, inboundXlsxHeaders_(), `${filenameBase}.xlsx`, 'Inbound', params.frequency);

    return {
      success: true,
      reportType: 'Inbound',
      name: pdfFile.getName(),
      rowCount: reportData.rows.length,
      pdfUrl: pdfFile.getUrl(),
      xlsxUrl: xlsxFile ? xlsxFile.getUrl() : ''
    };
  } catch (error) {
    Logger.log('Error generating inbound report: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function generateOutboundReport(params) {
  try {
    const reportData = getOutboundReportData(params);
    if (!reportData.rows || reportData.rows.length === 0) {
      return { success: false, message: 'No outbound data found.' };
    }

    const html = buildOutboundReportHtml_(params, reportData);
    const ts = formatDateForFilename(new Date());
    const filenameBase = `${params.frequency}_Outbound_${ts}`;

    const pdfFile = saveHtmlAsPdfToReports_(html, `${filenameBase}.pdf`, 'Outbound', params.frequency);
    const xlsxFile = saveRowsAsXlsxToReports_(reportData.rows, outboundXlsxHeaders_(), `${filenameBase}.xlsx`, 'Outbound', params.frequency);

    return {
      success: true,
      reportType: 'Outbound',
      name: pdfFile.getName(),
      rowCount: reportData.rows.length,
      pdfUrl: pdfFile.getUrl(),
      xlsxUrl: xlsxFile ? xlsxFile.getUrl() : ''
    };
  } catch (error) {
    Logger.log('Error generating outbound report: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============================================================================
// DATA FETCHING & SORTING
// ============================================================================

function getInboundReportData(params) {
  const sheet = getSheet(TABS.MASTER_LOG);
  const data = sheet.getDataRange().getValues();
  const poMap = buildPoToProjectMap_();
  const assetMap = buildItemAssetTypeMap_();
  const headers = data[0].map(h => (h || '').toString().trim());

  const dateIndex = headers.indexOf('Date_Received');
  const warehouseIndex = headers.indexOf('Warehouse');
  const projectIndex = headers.indexOf('Project');
  const fbpnIndex = headers.indexOf('FBPN');
  const qtyIndex = headers.indexOf('Qty_Received');
  const manufacturerIndex = headers.indexOf('Manufacturer');
  const mfpnIndex = headers.indexOf('MFPN');
  const carrierIndex = headers.indexOf('Carrier');
  const poIndex = headers.indexOf('Customer_PO_Number');
  const bolIndex = headers.indexOf('BOL_Number');
  const pushIndex = headers.indexOf('Push #');
  const assetTypeIndex = headers.indexOf('Asset Type') !== -1 ? headers.indexOf('Asset Type') : headers.indexOf('Asset_Type');
  const uomIndex = headers.indexOf('UOM') !== -1 ? headers.indexOf('UOM') : -1;

  const startDate = new Date(params.startDate);
  const endDate = new Date(params.endDate);
  endDate.setHours(23, 59, 59, 999);
  
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawDate = (dateIndex !== -1 ? row[dateIndex] : null);
    const rowDate = rawDate ? new Date(rawDate) : null;
    if (!rowDate || isNaN(rowDate.getTime())) continue;

    if (rowDate >= startDate && rowDate <= endDate) {
      rows.push({
        _dateObj: rowDate,
        dateReceived: formatDate(rowDate),
        warehouse: row[warehouseIndex] || '',
        project: resolveProject_(row[projectIndex], row[poIndex] || '', poMap),
        poNumber: row[poIndex] || '',
        bol: row[bolIndex] || '',
        push: row[pushIndex] || '',
        manufacturer: row[manufacturerIndex] || '',
        fbpn: row[fbpnIndex] || '',
        mfpn: (mfpnIndex !== -1 ? (row[mfpnIndex] || '') : ''),
        qty: row[qtyIndex] || 0,
        uom: (uomIndex !== -1 ? (row[uomIndex] || '') : ''),
        carrier: row[carrierIndex] || '',
        assetType: resolveAssetType_((assetTypeIndex !== -1 ? (row[assetTypeIndex] || '') : ''), row[fbpnIndex] || '', assetMap)
      });
    }
  }
  
  rows.sort((a, b) => {
    const da = a._dateObj.getTime();
    const db = b._dateObj.getTime();
    if (da !== db) return da - db;
    if (a.project !== b.project) return a.project.localeCompare(b.project);
    if (a.poNumber !== b.poNumber) return a.poNumber.localeCompare(b.poNumber);
    return a.bol.localeCompare(b.bol);
  });
  
  return { rows: rows, dateRange: formatDate(startDate) + ' - ' + formatDate(endDate) };
}

function getOutboundReportData(params) {
  const sheet = getSheet(TABS.OUTBOUNDLOG);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const dateIndex = headers.indexOf('Date');
  const companyIndex = headers.indexOf('Company');
  const projectIndex = headers.indexOf('Project');
  const fbpnIndex = headers.indexOf('FBPN');
  const qtyIndex = headers.indexOf('Qty');
  const manufacturerIndex = headers.indexOf('Manufacturer');
  const taskNumberIndex = headers.indexOf('Task_Number');
  const orderNumberIndex = headers.indexOf('Order_Number');
  
  const startDate = new Date(params.startDate);
  const endDate = new Date(params.endDate);
  endDate.setHours(23, 59, 59, 999);
  
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = new Date(row[dateIndex]);
    if (rowDate >= startDate && rowDate <= endDate) {
      rows.push({
        dateObj: rowDate,
        date: formatDate(rowDate),
        company: row[companyIndex] || '',
        project: row[projectIndex] || '',
        fbpn: row[fbpnIndex] || '',
        qty: row[qtyIndex] || 0,
        manufacturer: row[manufacturerIndex] || '',
        taskNumber: row[taskNumberIndex] || '',
        orderNumber: row[orderNumberIndex] || ''
      });
    }
  }
  
  rows.sort((a, b) => {
    if (a.dateObj.getTime() !== b.dateObj.getTime()) return a.dateObj.getTime() - b.dateObj.getTime();
    return a.orderNumber.localeCompare(b.orderNumber);
  });
  
  return { rows: rows, dateRange: formatDate(startDate) + ' - ' + formatDate(endDate) };
}

// ============================================================================
// HTML PORTRAIT GENERATION
// ============================================================================

function buildInboundReportHtml_(params, reportData) {
  const rows = reportData.rows || [];
  const title = `${params.frequency} Inbound Report`;
  const dateRange = reportData.dateRange || `${params.startDate} - ${params.endDate}`;
  const generatedAt = new Date().toLocaleString();

  const groups = {};
  rows.forEach(r => {
    const d = r.dateReceived || '';
    const proj = (r.project || '').toString().trim() || 'Unassigned Project';
    const po = (r.poNumber || '').toString().trim() || 'No PO';
    const bol = (r.bol || '').toString().trim() || 'No BOL';
    groups[d] = groups[d] || {};
    groups[d][proj] = groups[d][proj] || {};
    groups[d][proj][po] = groups[d][proj][po] || {};
    groups[d][proj][po][bol] = groups[d][proj][po][bol] || [];
    groups[d][proj][po][bol].push(r);
  });

  const esc = (s) => String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  let bodyHtml = '';
  Object.keys(groups).sort().forEach(dateKey => {
    bodyHtml += `<div class="lvl-date"><div class="lvl-date-h">Date: ${esc(dateKey)}</div>`;
    const projObj = groups[dateKey];
    Object.keys(projObj).sort().forEach(projKey => {
      bodyHtml += `<div class="lvl-project"><div class="lvl-project-h">Project: ${esc(projKey)}</div>`;
      const poObj = projObj[projKey];
      Object.keys(poObj).sort().forEach(poKey => {
        bodyHtml += `<div class="lvl-po"><div class="lvl-po-h">Customer PO: ${esc(poKey)}</div>`;
        const bolObj = poObj[poKey];
        Object.keys(bolObj).sort().forEach(bolKey => {
          const items = bolObj[bolKey];
          let bolTotal = 0;
          items.forEach(x => bolTotal += Number(x.qty || 0) || 0);
          bodyHtml += `
            <div class="lvl-bol">
              <div class="lvl-bol-h">
                <div>BOL: ${esc(bolKey)}</div>
                <div>Total: ${esc(bolTotal)} • Lines: ${esc(items.length)}</div>
              </div>
              <table>
                <thead>
                  <tr>
                    <th>Push #</th>
                    <th>Asset Type</th>
                    <th>FBPN</th>
                    <th>Qty</th>
                    <th>Carrier</th>
                  </tr>
                </thead>
                <tbody>
                  ${items.map(it => `
                    <tr>
                      <td>${esc(it.push)}</td>
                      <td>${esc(it.assetType || '')}</td>
                      <td class="mono">${esc(it.fbpn)}</td>
                      <td>${esc(it.qty)}</td>
                      <td>${esc(it.carrier)}</td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>`;
        });
        bodyHtml += `</div>`; 
      });
      bodyHtml += `</div>`;
    });
    bodyHtml += `</div>`;
  });

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    @page { size: letter; margin: 0.5in; }
    body { font-family: Arial, sans-serif; font-size: 9pt; line-height: 1.3; color: #111827; margin: 0; }
    .header { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 2px solid #111827; padding-bottom: 10px; margin-bottom: 20px; }
    .title { font-size: 16pt; font-weight: 800; }
    .lvl-date-h { font-size: 11pt; font-weight: 800; padding: 8px; background: #374151; color: #fff; border-radius: 4px; margin-bottom: 10px; page-break-after: avoid; }
    .lvl-project-h { font-weight: 700; padding: 4px 8px; border-left: 4px solid #111827; margin-bottom: 6px; page-break-after: avoid; }
    .lvl-po { page-break-inside: avoid; margin-bottom: 10px; }
    .lvl-po-h { font-weight: 700; padding: 6px 8px; background: #4b5563; color: #fff; border-radius: 4px; margin-bottom: 2px; }
    .lvl-bol { page-break-inside: avoid; margin-bottom: 10px; }
    .lvl-bol-h { display: flex; justify-content: space-between; padding: 4px 8px; font-size: 8pt; font-weight: 700; border-bottom: 1px solid #111827; }
    table { width: 100%; border-collapse: collapse; table-layout: fixed; }
    th { text-align: center; font-size: 7.5pt; background: #d1d5db; padding: 6px; border: 1px solid #9ca3af; }
    td { padding: 6px; border: 1px solid #e5e7eb; text-align: center; font-weight: bold; word-wrap: break-word; }
    tr:nth-child(even) td { background: #f9fafb; }
    .mono { font-family: monospace; }
  </style></head><body>
  <div class="header">
    <div class="title">${esc(title)}</div>
    <div style="text-align:right; font-size:8pt;"><span>Range:</span> ${esc(dateRange)}</div>
  </div>
  ${bodyHtml}
</body></html>`;
}

function buildOutboundReportHtml_(params, reportData) {
  const rows = reportData.rows || [];
  const title = `${params.frequency} Outbound Report`;
  const dateRange = reportData.dateRange || `${params.startDate} - ${params.endDate}`;

  const groups = {};
  rows.forEach(r => {
    const d = r.date || '';
    const company = (r.company || '').toString().trim() || 'Unknown Company';
    const task = (r.taskNumber || '').toString().trim() || 'No Task';
    groups[d] = groups[d] || {};
    groups[d][company] = groups[d][company] || {};
    groups[d][company][task] = groups[d][company][task] || [];
    groups[d][company][task].push(r);
  });

  const esc = (s) => String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  let bodyHtml = '';
  Object.keys(groups).sort().forEach(dateKey => {
    bodyHtml += `<div class="lvl1"><div class="lvl1-h">Date: ${esc(dateKey)}</div>`;
    const companyObj = groups[dateKey];
    Object.keys(companyObj).sort().forEach(companyKey => {
      bodyHtml += `<div class="lvl2"><div class="lvl2-h">Company: ${esc(companyKey)}</div>`;
      const taskObj = companyObj[companyKey];
      Object.keys(taskObj).sort().forEach(taskKey => {
        const items = taskObj[taskKey];
        let taskTotal = 0;
        items.forEach(x => taskTotal += Number(x.qty || 0) || 0);
        bodyHtml += `
          <div class="lvl3">
            <div class="lvl3-h">
              <div>Task #: ${esc(taskKey)}</div>
              <div style="font-size:8pt; opacity:0.9;">Lines: ${items.length} • Qty: ${taskTotal}</div>
            </div>
            <table class="tbl">
              <thead>
                <tr>
                  <th>Order #</th>
                  <th>FBPN</th>
                  <th>Qty</th>
                  <th>Project</th>
                </tr>
              </thead>
              <tbody>
                ${items.map(it => `
                  <tr>
                    <td class="mono">${esc(it.orderNumber)}</td>
                    <td class="mono">${esc(it.fbpn)}</td>
                    <td>${esc(it.qty)}</td>
                    <td>${esc(it.project)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>`;
      });
      bodyHtml += `</div>`;
    });
    bodyHtml += `</div>`;
  });

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    @page { size: letter; margin: 0.5in; }
    body { font-family: Arial, sans-serif; color: #111827; font-size: 9pt; margin: 0; }
    .hdr { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 2px solid #111827; padding-bottom: 10px; margin-bottom: 20px; }
    .h1 { font-size: 16pt; font-weight: 800; }
    .lvl1-h { font-size: 11pt; font-weight: 800; padding: 10px; background: #374151; color: #fff; border-radius: 4px; margin-bottom: 10px; page-break-after: avoid; }
    .lvl2 { page-break-inside: avoid; margin-bottom: 15px; }
    .lvl2-h { font-weight: 700; padding: 5px 10px; border-left: 4px solid #111827; margin-bottom: 8px; page-break-after: avoid; }
    .lvl3 { page-break-inside: avoid; margin-bottom: 10px; }
    .lvl3-h { display: flex; justify-content: space-between; padding: 8px 10px; background: #4b5563; color: #fff; border-radius: 4px; font-weight: 700; }
    .tbl { width: 100%; border-collapse: collapse; table-layout: fixed; }
    .tbl th { text-align: center; font-size: 7.5pt; background: #d1d5db; padding: 8px; border: 1px solid #9ca3af; }
    .tbl td { padding: 8px; border: 1px solid #e5e7eb; text-align: center; font-weight: bold; word-wrap: break-word; }
    .tbl tr:nth-child(even) td { background: #f9fafb; }
    .mono { font-family: monospace; }
  </style></head><body>
  <div class="hdr">
    <div class="h1">${esc(title)}</div>
    <div style="text-align:right; font-size:8pt;">Range: ${esc(dateRange)}</div>
  </div>
  ${bodyHtml}
</body></html>`;
}

// ============================================================================
// EXCEL GENERATION & STYLING
// ============================================================================

function saveHtmlAsPdfToReports_(html, filename, reportType, frequency) {
  const blob = HtmlService.createHtmlOutput(html).getBlob().setName(filename).getAs(MimeType.PDF);
  return saveReportToFolder(blob, reportType, frequency);
}

function inboundXlsxHeaders_() {
  return ['Date_Received','Project','Customer_PO_Number','BOL_Number','Warehouse','Push #','Asset Type','Manufacturer','FBPN','MFPN','Qty_Received','UOM','Carrier'];
}

function outboundXlsxHeaders_() {
  return ['Date','Company','Task_Number','Order_Number','Project','FBPN','Qty','Manufacturer'];
}

function saveRowsAsXlsxToReports_(rows, headers, filename, reportType, frequency) {
  try {
    const tmp = SpreadsheetApp.create('TEMP_' + filename.replace(/\.xlsx$/i, ''));
    const sh = tmp.getSheets()[0];
    sh.setName('Report');
    sh.getRange(1,1,1,headers.length).setValues([headers]);

    applyReportXlsxStyles_(sh, headers, reportType);

    const values = [];
    if (reportType === 'Inbound') {
      rows.forEach(r => values.push([r.dateReceived||'',r.project||'',r.poNumber||'',r.bol||'',r.warehouse||'',r.push||'',r.assetType||'',r.manufacturer||'',r.fbpn||'',r.mfpn||'',r.qty||0,r.uom||'',r.carrier||'']));
    } else {
      rows.forEach(r => values.push([r.date||'', r.company||'', r.taskNumber||'', r.orderNumber||'', r.project||'', r.fbpn||'', r.qty||0, r.manufacturer||'']));
    }
    
    if (values.length) sh.getRange(2,1,values.length,headers.length).setValues(values);
    sh.autoResizeColumns(1, headers.length);

    const xlsxBlob = exportSpreadsheetAsXlsx_(tmp.getId(), filename);
    const xlsxFile = saveReportXlsxToFolder_(xlsxBlob, reportType, frequency);

    DriveApp.getFileById(tmp.getId()).setTrashed(true);
    return xlsxFile;
  } catch (e) {
    Logger.log('Warning: XLSX export failed: ' + e.toString());
    return null;
  }
}

function applyReportXlsxStyles_(sheet, headers, reportType) {
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setFontWeight('bold')
    .setFontColor('#111827')
    .setBackground('#d1d5db') 
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const bodyRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
  bodyRange
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('bold');
  
  bodyRange.setBorder(true, true, true, true, true, true, '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, headers.length);
}

function exportSpreadsheetAsXlsx_(spreadsheetId, filename) {
  const file = DriveApp.getFileById(spreadsheetId);
  const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const blob = file.getBlob().getAs(mime);
  blob.setName(filename);
  return blob;
}

function saveReportXlsxToFolder_(xlsxBlob, reportType, frequency) {
  const rootFolder = DriveApp.getFolderById(FOLDERS.IMS_Reports);
  const typeFolder = getOrCreateFolder(rootFolder, reportType + ' Reports');
  const frequencyFolder = getOrCreateFolder(typeFolder, frequency);
  const yearFolder = getOrCreateFolder(frequencyFolder, String(new Date().getFullYear()));
  return yearFolder.createFile(xlsxBlob);
}