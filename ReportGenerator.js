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
// DATE RANGE PRESETS
// ============================================================================

function formatDateForFilename(date) {
  const d = date instanceof Date ? date : new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getDateRangePreset(preset) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let startDate, endDate;

  switch (preset) {
    case 'today':
      startDate = new Date(today);
      endDate = new Date(today);
      break;
    case 'yesterday':
      startDate = new Date(today);
      startDate.setDate(startDate.getDate() - 1);
      endDate = new Date(startDate);
      break;
    case 'thisWeek':
      startDate = new Date(today);
      startDate.setDate(today.getDate() - today.getDay()); // Sunday
      endDate = new Date(today);
      break;
    case 'lastWeek':
      startDate = new Date(today);
      startDate.setDate(today.getDate() - today.getDay() - 7); // Last Sunday
      endDate = new Date(startDate);
      endDate.setDate(endDate.getDate() + 6); // Last Saturday
      break;
    case 'thisMonth':
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
      endDate = new Date(today);
      break;
    case 'lastMonth':
      startDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      endDate = new Date(today.getFullYear(), today.getMonth(), 0); // Last day of previous month
      break;
    default:
      startDate = new Date(today);
      endDate = new Date(today);
  }

  // Format as mm/dd/yyyy for the modal's convertToInputDate function
  const formatDate = (d) => {
    const month = (d.getMonth() + 1).toString();
    const day = d.getDate().toString();
    const year = d.getFullYear().toString();
    return `${month}/${day}/${year}`;
  };

  return {
    startDate: formatDate(startDate),
    endDate: formatDate(endDate)
  };
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
    const projA = String(a.project || '');
    const projB = String(b.project || '');
    if (projA !== projB) return projA.localeCompare(projB);
    const poA = String(a.poNumber || '');
    const poB = String(b.poNumber || '');
    if (poA !== poB) return poA.localeCompare(poB);
    return String(a.bol || '').localeCompare(String(b.bol || ''));
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
    return String(a.orderNumber || '').localeCompare(String(b.orderNumber || ''));
  });
  
  return { rows: rows, dateRange: formatDate(startDate) + ' - ' + formatDate(endDate) };
}

// ============================================================================
// HTML LANDSCAPE A4 GENERATION
// ============================================================================

function buildInboundReportHtml_(params, reportData) {
  const rows = reportData.rows || [];
  const dateRange = reportData.dateRange || `${params.startDate} - ${params.endDate}`;
  const generatedAt = new Date().toLocaleString();
  const reportId = `${new Date().getFullYear()}-INB-${String(Math.floor(Math.random() * 1000)).padStart(3, '0')}`;

  // Calculate summary statistics
  let totalQty = 0;
  const uniqueProjects = new Set();
  const uniquePOs = new Set();
  const uniqueBOLs = new Set();
  const uniqueFBPNs = new Set();
  rows.forEach(r => {
    totalQty += Number(r.qty || 0) || 0;
    if (r.project) uniqueProjects.add(r.project);
    if (r.poNumber) uniquePOs.add(r.poNumber);
    if (r.bol) uniqueBOLs.add(r.bol);
    if (r.fbpn) uniqueFBPNs.add(r.fbpn);
  });

  // Group by project -> PO
  const projectGroups = {};
  rows.forEach(r => {
    const proj = (r.project || '').toString().trim() || 'Unassigned';
    const po = (r.poNumber || '').toString().trim() || 'No PO';
    projectGroups[proj] = projectGroups[proj] || {};
    projectGroups[proj][po] = projectGroups[proj][po] || [];
    projectGroups[proj][po].push(r);
  });

  const esc = (s) => String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  // Build hierarchy HTML (Page 1)
  let hierarchyHtml = '';
  Object.keys(projectGroups).sort().forEach(projKey => {
    hierarchyHtml += `<div class="h-proj">${esc(projKey)}</div>`;
    const poObj = projectGroups[projKey];
    Object.keys(poObj).sort().forEach(poKey => {
      const items = poObj[poKey];
      hierarchyHtml += `<div class="h-po-wrap"><div class="h-po-head">PO: ${esc(poKey)}</div>`;
      items.forEach(item => {
        hierarchyHtml += `
          <div class="h-item">
            <div><span class="lbl">MFR:</span>${esc(item.manufacturer)}</div>
            <div><span class="lbl">FBPN:</span>${esc(item.fbpn)}</div>
            <div><span class="lbl">Asset:</span>${esc(item.assetType || '')}</div>
            <div style="text-align:right"><strong>${Number(item.qty || 0).toLocaleString()}</strong> <span class="lbl">QTY</span></div>
          </div>`;
      });
      hierarchyHtml += `</div>`;
    });
  });

  // Build detailed tables HTML (Page 2+)
  let detailHtml = '';
  Object.keys(projectGroups).sort().forEach(projKey => {
    const allItems = [];
    Object.values(projectGroups[projKey]).forEach(items => allItems.push(...items));
    if (allItems.length === 0) return;

    detailHtml += `
      <div class="detail-block">
        <div class="detail-head">PROJECT: ${esc(projKey)}</div>
        <table>
          <thead>
            <tr>
              <th style="width:9%">Date</th>
              <th style="width:7%">Whse</th>
              <th style="width:8%">Asset</th>
              <th style="width:12%">Manufacturer</th>
              <th style="width:12%">FBPN</th>
              <th style="width:10%">PO#</th>
              <th style="width:10%">BOL#</th>
              <th style="width:6%">Qty</th>
              <th style="width:8%">Carrier</th>
              <th>Push #</th>
            </tr>
          </thead>
          <tbody>
            ${allItems.map(item => `
              <tr>
                <td>${esc(item.dateReceived)}</td>
                <td>${esc(item.warehouse)}</td>
                <td>${esc(item.assetType || '')}</td>
                <td>${esc(item.manufacturer)}</td>
                <td>${esc(item.fbpn)}</td>
                <td>${esc(item.poNumber)}</td>
                <td>${esc(item.bol)}</td>
                <td class="qty-cell">${Number(item.qty || 0).toLocaleString()}</td>
                <td>${esc(item.carrier)}</td>
                <td>${esc(item.push)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>`;
  });

  return `<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Inbound Report</title>
<style>
  @page { size: A4 landscape; margin: 15mm; }
  :root { --primary: #0f172a; --accent: #2563eb; --success: #15803d; --border: #cbd5e1; --bg-odd: #f8fafc; }
  body { font-family: 'Segoe UI', 'Helvetica', sans-serif; background: white; margin: 0; padding: 10mm; color: #1e293b; -webkit-print-color-adjust: exact; font-size: 11px; }

  .print-header { position: fixed; top: 0; right: 0; text-align: right; font-size: 10px; color: #64748b; width: 100%; }
  .print-footer { position: fixed; bottom: 0; left: 0; width: 100%; text-align: center; font-size: 9px; color: #94a3b8; border-top: 1px solid #cbd5e1; padding-top: 5px; background: white; }

  .report-section { width: 100%; margin: 0 auto 20px auto; background: white; position: relative; }
  .page-break-after { page-break-after: always; }

  h1 { margin: 0; font-size: 24px; text-transform: uppercase; letter-spacing: 1px; color: var(--primary); }
  h2 { font-size: 16px; color: var(--primary); border-bottom: 2px solid var(--primary); padding-bottom: 5px; margin-top: 0; margin-bottom: 15px; }
  .meta-info { text-align: right; font-size: 11px; color: #475569; }

  .kpi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin: 20px 0; }
  .kpi-card { border: 2px solid var(--success); background: #f0fdf4; padding: 15px; text-align: center; border-radius: 6px; }
  .kpi-num { display: block; font-size: 28px; font-weight: 900; color: var(--success); }
  .kpi-lbl { font-size: 11px; font-weight: 700; text-transform: uppercase; color: #166534; }

  .hierarchy-box { border: 1px solid var(--border); border-radius: 4px; overflow: hidden; margin-bottom: 20px; }
  .h-proj { background: var(--primary); color: white; padding: 8px 12px; font-weight: 700; font-size: 13px; }
  .h-po-wrap { background: #fff; border-bottom: 1px solid var(--border); }
  .h-po-wrap:last-child { border-bottom: none; }
  .h-po-head { background: #f1f5f9; color: var(--accent); padding: 6px 12px; font-size: 11px; font-weight: 700; border-bottom: 1px solid #e2e8f0; }
  .h-item { display: grid; grid-template-columns: 2fr 1.5fr 1fr 1fr; padding: 5px 12px 5px 25px; font-size: 10px; border-bottom: 1px dashed #e2e8f0; align-items: center; }
  .h-item:last-child { border-bottom: none; }
  .lbl { font-size: 8px; text-transform: uppercase; color: #64748b; margin-right: 4px; font-weight: 700; }

  .detail-block { margin-bottom: 25px; border: 1px solid var(--border); page-break-inside: avoid; }
  .detail-head { background: var(--primary); color: white; padding: 8px 12px; font-weight: 700; font-size: 13px; }

  table { width: 100%; border-collapse: collapse; font-size: 10px; }
  thead { display: table-header-group; }
  tr { page-break-inside: avoid; }
  th { background: #f1f5f9; text-align: left; padding: 7px 5px; border-bottom: 2px solid #94a3b8; font-size: 10px; font-weight: 800; color: var(--primary); border-right: 1px solid #cbd5e1; }
  td { padding: 6px 5px; border-bottom: 1px solid #cbd5e1; border-right: 1px solid #f1f5f9; vertical-align: top; }
  tr:nth-child(even) td { background: var(--bg-odd); }
  .qty-cell { background: #eff6ff; text-align: center; font-weight: 700; color: var(--primary); }

  @media print {
    body { background: white; margin: 0; padding: 0; }
    .report-section { box-shadow: none; margin: 0; width: 100%; max-width: none; border: none; padding: 0 0 15px 0; }
    .print-footer { position: fixed; bottom: 0; }
    body { margin-top: 15px; margin-bottom: 25px; }
  }
</style></head><body>

<div class="print-header">${esc(params.frequency)} Report &nbsp;|&nbsp; ${esc(dateRange)}</div>
<div class="print-footer">CONFIDENTIAL — Generated: ${esc(generatedAt)}</div>

<div class="report-section page-break-after">
  <div style="display:flex; justify-content:space-between; align-items:flex-end; border-bottom: 3px solid #0f172a; padding-bottom:10px; margin-bottom:15px;">
    <div>
      <h1>Inbound Logistics Summary</h1>
      <div style="color:#2563eb; font-weight:700; font-size:12px; margin-top:5px;">${esc(params.frequency)} Report • ${esc(dateRange)}</div>
    </div>
    <div class="meta-info">
      <strong>Report ID:</strong> ${esc(reportId)}<br>
      <strong>Generated:</strong> ${esc(generatedAt)}
    </div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card"><span class="kpi-num">${uniqueBOLs.size}</span><span class="kpi-lbl">Total BOLs</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniqueFBPNs.size}</span><span class="kpi-lbl">Unique FBPNs</span></div>
    <div class="kpi-card"><span class="kpi-num">${totalQty.toLocaleString()}</span><span class="kpi-lbl">Total Qty</span></div>
  </div>

  <h3 style="font-size:13px; text-transform:uppercase; margin-bottom:10px; color:#0f172a;">Project Breakdown Structure</h3>
  <div class="hierarchy-box">${hierarchyHtml}</div>
</div>

<div class="report-section">
  <h2>Detailed Delivery Overview</h2>
  ${detailHtml}
</div>

</body></html>`;
}

function buildOutboundReportHtml_(params, reportData) {
  const rows = reportData.rows || [];
  const dateRange = reportData.dateRange || `${params.startDate} - ${params.endDate}`;
  const generatedAt = new Date().toLocaleString();
  const reportId = `${new Date().getFullYear()}-OUT-${String(Math.floor(Math.random() * 1000)).padStart(3, '0')}`;

  // Calculate summary statistics
  let totalQty = 0;
  const uniqueCompanies = new Set();
  const uniqueTasks = new Set();
  const uniqueOrders = new Set();
  const uniqueFBPNs = new Set();
  rows.forEach(r => {
    totalQty += Number(r.qty || 0) || 0;
    if (r.company) uniqueCompanies.add(r.company);
    if (r.taskNumber) uniqueTasks.add(r.taskNumber);
    if (r.orderNumber) uniqueOrders.add(r.orderNumber);
    if (r.fbpn) uniqueFBPNs.add(r.fbpn);
  });

  // Group by company -> task
  const companyGroups = {};
  rows.forEach(r => {
    const company = (r.company || '').toString().trim() || 'Unknown';
    const task = (r.taskNumber || '').toString().trim() || 'No Task';
    companyGroups[company] = companyGroups[company] || {};
    companyGroups[company][task] = companyGroups[company][task] || [];
    companyGroups[company][task].push(r);
  });

  const esc = (s) => String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  // Build hierarchy HTML (Page 1)
  let hierarchyHtml = '';
  Object.keys(companyGroups).sort().forEach(companyKey => {
    hierarchyHtml += `<div class="h-proj">${esc(companyKey)}</div>`;
    const taskObj = companyGroups[companyKey];
    Object.keys(taskObj).sort().forEach(taskKey => {
      const items = taskObj[taskKey];
      hierarchyHtml += `<div class="h-po-wrap"><div class="h-po-head">Task #: ${esc(taskKey)}</div>`;
      items.forEach(item => {
        hierarchyHtml += `
          <div class="h-item">
            <div><span class="lbl">Order:</span>${esc(item.orderNumber)}</div>
            <div><span class="lbl">FBPN:</span>${esc(item.fbpn)}</div>
            <div><span class="lbl">Project:</span>${esc(item.project)}</div>
            <div style="text-align:right"><strong>${Number(item.qty || 0).toLocaleString()}</strong> <span class="lbl">QTY</span></div>
          </div>`;
      });
      hierarchyHtml += `</div>`;
    });
  });

  // Build detailed tables HTML (Page 2+)
  let detailHtml = '';
  Object.keys(companyGroups).sort().forEach(companyKey => {
    const allItems = [];
    Object.values(companyGroups[companyKey]).forEach(items => allItems.push(...items));
    if (allItems.length === 0) return;

    detailHtml += `
      <div class="detail-block">
        <div class="detail-head">COMPANY: ${esc(companyKey)}</div>
        <table>
          <thead>
            <tr>
              <th style="width:10%">Date</th>
              <th style="width:10%">Task #</th>
              <th style="width:15%">Order #</th>
              <th style="width:15%">FBPN</th>
              <th style="width:8%">Qty</th>
              <th style="width:15%">Project</th>
              <th>Manufacturer</th>
            </tr>
          </thead>
          <tbody>
            ${allItems.map(item => `
              <tr>
                <td>${esc(item.date)}</td>
                <td>${esc(item.taskNumber)}</td>
                <td>${esc(item.orderNumber)}</td>
                <td>${esc(item.fbpn)}</td>
                <td class="qty-cell">${Number(item.qty || 0).toLocaleString()}</td>
                <td>${esc(item.project)}</td>
                <td>${esc(item.manufacturer)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>`;
  });

  return `<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Outbound Report</title>
<style>
  @page { size: A4 landscape; margin: 15mm; }
  :root { --primary: #065f46; --accent: #10b981; --success: #f97316; --border: #a7f3d0; --bg-odd: #f0fdf4; }
  body { font-family: 'Segoe UI', 'Helvetica', sans-serif; background: white; margin: 0; padding: 10mm; color: #1e293b; -webkit-print-color-adjust: exact; font-size: 11px; }

  .print-header { position: fixed; top: 0; right: 0; text-align: right; font-size: 10px; color: #64748b; width: 100%; }
  .print-footer { position: fixed; bottom: 0; left: 0; width: 100%; text-align: center; font-size: 9px; color: #94a3b8; border-top: 1px solid #a7f3d0; padding-top: 5px; background: white; }

  .report-section { width: 100%; margin: 0 auto 20px auto; background: white; position: relative; }
  .page-break-after { page-break-after: always; }

  h1 { margin: 0; font-size: 24px; text-transform: uppercase; letter-spacing: 1px; color: var(--primary); }
  h2 { font-size: 16px; color: var(--primary); border-bottom: 2px solid var(--primary); padding-bottom: 5px; margin-top: 0; margin-bottom: 15px; }
  .meta-info { text-align: right; font-size: 11px; color: #475569; }

  .kpi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin: 20px 0; }
  .kpi-card { border: 2px solid var(--success); background: #fff7ed; padding: 15px; text-align: center; border-radius: 6px; }
  .kpi-num { display: block; font-size: 28px; font-weight: 900; color: var(--success); }
  .kpi-lbl { font-size: 11px; font-weight: 700; text-transform: uppercase; color: #c2410c; }

  .hierarchy-box { border: 1px solid var(--border); border-radius: 4px; overflow: hidden; margin-bottom: 20px; }
  .h-proj { background: var(--primary); color: white; padding: 8px 12px; font-weight: 700; font-size: 13px; }
  .h-po-wrap { background: #fff; border-bottom: 1px solid var(--border); }
  .h-po-wrap:last-child { border-bottom: none; }
  .h-po-head { background: #f0fdf4; color: var(--success); padding: 6px 12px; font-size: 11px; font-weight: 700; border-bottom: 1px solid #d1fae5; }
  .h-item { display: grid; grid-template-columns: 1.5fr 1.5fr 1.5fr 1fr; padding: 5px 12px 5px 25px; font-size: 10px; border-bottom: 1px dashed #e2e8f0; align-items: center; }
  .h-item:last-child { border-bottom: none; }
  .lbl { font-size: 8px; text-transform: uppercase; color: #64748b; margin-right: 4px; font-weight: 700; }

  .detail-block { margin-bottom: 25px; border: 1px solid var(--border); page-break-inside: avoid; }
  .detail-head { background: var(--primary); color: white; padding: 8px 12px; font-weight: 700; font-size: 13px; }

  table { width: 100%; border-collapse: collapse; font-size: 10px; }
  thead { display: table-header-group; }
  tr { page-break-inside: avoid; }
  th { background: #f0fdf4; text-align: left; padding: 7px 5px; border-bottom: 2px solid #10b981; font-size: 10px; font-weight: 800; color: var(--primary); border-right: 1px solid #a7f3d0; }
  td { padding: 6px 5px; border-bottom: 1px solid #d1fae5; border-right: 1px solid #f0fdf4; vertical-align: top; }
  tr:nth-child(even) td { background: var(--bg-odd); }
  .qty-cell { background: #fff7ed; text-align: center; font-weight: 700; color: var(--success); }

  @media print {
    body { background: white; margin: 0; padding: 0; }
    .report-section { box-shadow: none; margin: 0; width: 100%; max-width: none; border: none; padding: 0 0 15px 0; }
    .print-footer { position: fixed; bottom: 0; }
    body { margin-top: 15px; margin-bottom: 25px; }
  }
</style></head><body>

<div class="print-header">${esc(params.frequency)} Report &nbsp;|&nbsp; ${esc(dateRange)}</div>
<div class="print-footer">CONFIDENTIAL — Generated: ${esc(generatedAt)}</div>

<div class="report-section page-break-after">
  <div style="display:flex; justify-content:space-between; align-items:flex-end; border-bottom: 3px solid #065f46; padding-bottom:10px; margin-bottom:15px;">
    <div>
      <h1>Outbound Logistics Summary</h1>
      <div style="color:#10b981; font-weight:700; font-size:12px; margin-top:5px;">${esc(params.frequency)} Report • ${esc(dateRange)}</div>
    </div>
    <div class="meta-info">
      <strong>Report ID:</strong> ${esc(reportId)}<br>
      <strong>Generated:</strong> ${esc(generatedAt)}
    </div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card"><span class="kpi-num">${uniqueOrders.size}</span><span class="kpi-lbl">Total Orders</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniqueFBPNs.size}</span><span class="kpi-lbl">Unique FBPNs</span></div>
    <div class="kpi-card"><span class="kpi-num">${totalQty.toLocaleString()}</span><span class="kpi-lbl">Total Qty</span></div>
  </div>

  <h3 style="font-size:13px; text-transform:uppercase; margin-bottom:10px; color:#065f46;">Company Breakdown Structure</h3>
  <div class="hierarchy-box">${hierarchyHtml}</div>
</div>

<div class="report-section">
  <h2>Detailed Shipment Overview</h2>
  ${detailHtml}
</div>

</body></html>`;
}

// ============================================================================
// EXCEL GENERATION & STYLING
// ============================================================================

function saveReportToFolder(blob, reportType, frequency) {
  const rootFolder = DriveApp.getFolderById(FOLDERS.IMS_Reports);
  const typeFolder = getOrCreateFolder(rootFolder, reportType + ' Reports');
  const frequencyFolder = getOrCreateFolder(typeFolder, frequency);
  const yearFolder = getOrCreateFolder(frequencyFolder, String(new Date().getFullYear()));
  return yearFolder.createFile(blob);
}

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