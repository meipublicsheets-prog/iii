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
  const reportId = `INB-${new Date().getFullYear()}-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`;

  // Calculate summary statistics
  let totalQty = 0;
  const uniqueProjects = new Set();
  const uniquePOs = new Set();
  const uniqueBOLs = new Set();
  const uniqueFBPNs = new Set();
  const uniqueCarriers = new Set();
  const uniqueWarehouses = new Set();
  const dailyData = {};
  const projectQty = {};
  const carrierQty = {};

  rows.forEach(r => {
    const qty = Number(r.qty || 0) || 0;
    totalQty += qty;
    if (r.project) { uniqueProjects.add(r.project); projectQty[r.project] = (projectQty[r.project] || 0) + qty; }
    if (r.poNumber) uniquePOs.add(r.poNumber);
    if (r.bol) uniqueBOLs.add(r.bol);
    if (r.fbpn) uniqueFBPNs.add(r.fbpn);
    if (r.carrier) { uniqueCarriers.add(r.carrier); carrierQty[r.carrier] = (carrierQty[r.carrier] || 0) + qty; }
    if (r.warehouse) uniqueWarehouses.add(r.warehouse);
    const dateKey = r.dateReceived || 'Unknown';
    dailyData[dateKey] = dailyData[dateKey] || { bols: new Set(), qty: 0 };
    dailyData[dateKey].bols.add(r.bol);
    dailyData[dateKey].qty += qty;
  });

  // Top projects and carriers for insights
  const topProjects = Object.entries(projectQty).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const topCarriers = Object.entries(carrierQty).sort((a, b) => b[1] - a[1]).slice(0, 3);
  const avgDailyQty = Object.keys(dailyData).length > 0 ? Math.round(totalQty / Object.keys(dailyData).length) : 0;

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

  // Build daily trend table
  const dailyTrendHtml = Object.entries(dailyData).sort((a, b) => a[0].localeCompare(b[0])).map(([date, data]) =>
    `<tr><td>${esc(date)}</td><td class="num">${data.bols.size}</td><td class="num">${data.qty.toLocaleString()}</td></tr>`
  ).join('');

  // Build project summary table
  const projectSummaryHtml = topProjects.map(([proj, qty], i) =>
    `<tr><td>${i + 1}</td><td>${esc(proj)}</td><td class="num">${qty.toLocaleString()}</td><td class="num">${totalQty > 0 ? ((qty / totalQty) * 100).toFixed(1) : 0}%</td></tr>`
  ).join('');

  // Build detailed tables HTML
  let detailHtml = '';
  Object.keys(projectGroups).sort().forEach(projKey => {
    const allItems = [];
    Object.values(projectGroups[projKey]).forEach(items => allItems.push(...items));
    if (allItems.length === 0) return;
    const projTotal = allItems.reduce((sum, item) => sum + (Number(item.qty) || 0), 0);

    detailHtml += `
      <div class="detail-block">
        <div class="detail-head">
          <span>${esc(projKey)}</span>
          <span class="detail-head-stat">${allItems.length} items • ${projTotal.toLocaleString()} units</span>
        </div>
        <table>
          <thead>
            <tr>
              <th style="width:9%">Date</th>
              <th style="width:6%">Whse</th>
              <th style="width:10%">Manufacturer</th>
              <th style="width:12%">FBPN</th>
              <th style="width:10%">PO#</th>
              <th style="width:10%">BOL#</th>
              <th style="width:7%">Qty</th>
              <th style="width:8%">Carrier</th>
              <th>Asset Type</th>
            </tr>
          </thead>
          <tbody>
            ${allItems.map(item => `
              <tr>
                <td>${esc(item.dateReceived)}</td>
                <td>${esc(item.warehouse)}</td>
                <td>${esc(item.manufacturer)}</td>
                <td class="mono">${esc(item.fbpn)}</td>
                <td class="mono">${esc(item.poNumber)}</td>
                <td class="mono">${esc(item.bol)}</td>
                <td class="qty-cell">${Number(item.qty || 0).toLocaleString()}</td>
                <td>${esc(item.carrier)}</td>
                <td>${esc(item.assetType || '')}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>`;
  });

  return `<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Inbound Logistics Report</title>
<style>
  @page { size: A4 landscape; margin: 12mm 15mm; }
  :root { --primary: #1e3a5f; --accent: #0ea5e9; --success: #059669; --warn: #d97706; --border: #e2e8f0; --bg-light: #f8fafc; --text: #334155; --text-muted: #64748b; }
  * { box-sizing: border-box; }
  body { font-family: 'Inter', 'Segoe UI', system-ui, sans-serif; background: #fff; margin: 0; padding: 8mm; color: var(--text); font-size: 10px; line-height: 1.4; -webkit-print-color-adjust: exact; print-color-adjust: exact; }

  /* Header Bar */
  .header-bar { background: linear-gradient(135deg, var(--primary) 0%, #0f172a 100%); color: white; padding: 20px 25px; margin: -8mm -8mm 20px -8mm; display: flex; justify-content: space-between; align-items: center; }
  .header-bar h1 { margin: 0; font-size: 22px; font-weight: 700; letter-spacing: -0.5px; }
  .header-bar .subtitle { opacity: 0.85; font-size: 12px; margin-top: 4px; }
  .header-bar .meta { text-align: right; font-size: 10px; opacity: 0.9; }
  .header-bar .meta strong { display: block; font-size: 11px; }

  /* Executive Summary Card */
  .exec-summary { background: white; border: 1px solid var(--border); border-radius: 8px; padding: 20px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
  .exec-title { font-size: 14px; font-weight: 700; color: var(--primary); margin: 0 0 15px 0; padding-bottom: 10px; border-bottom: 2px solid var(--accent); display: flex; align-items: center; gap: 8px; }
  .exec-title::before { content: ''; width: 4px; height: 18px; background: var(--accent); border-radius: 2px; }

  /* KPI Grid */
  .kpi-grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 15px; margin-bottom: 20px; }
  .kpi-card { background: white; border: 1px solid var(--border); border-radius: 6px; padding: 15px; text-align: center; position: relative; overflow: hidden; }
  .kpi-card::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: var(--accent); }
  .kpi-card.highlight::before { background: var(--success); }
  .kpi-num { display: block; font-size: 26px; font-weight: 800; color: var(--primary); margin-bottom: 4px; }
  .kpi-lbl { font-size: 9px; font-weight: 600; text-transform: uppercase; color: var(--text-muted); letter-spacing: 0.5px; }

  /* Insights Panel */
  .insights-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
  .insight-card { background: var(--bg-light); border-radius: 6px; padding: 15px; }
  .insight-card h4 { margin: 0 0 12px 0; font-size: 11px; font-weight: 700; color: var(--primary); text-transform: uppercase; letter-spacing: 0.5px; }
  .insight-list { margin: 0; padding: 0; list-style: none; }
  .insight-list li { padding: 6px 0; border-bottom: 1px solid var(--border); font-size: 10px; display: flex; justify-content: space-between; }
  .insight-list li:last-child { border-bottom: none; }
  .insight-list .label { color: var(--text-muted); }
  .insight-list .value { font-weight: 700; color: var(--primary); }

  /* Trend Table */
  .trend-section { margin-bottom: 20px; }
  .trend-section h3 { font-size: 12px; font-weight: 700; color: var(--primary); margin: 0 0 10px 0; }
  .trend-table { width: 100%; border-collapse: collapse; font-size: 9px; }
  .trend-table th { background: var(--primary); color: white; padding: 8px 10px; text-align: left; font-weight: 600; }
  .trend-table td { padding: 7px 10px; border-bottom: 1px solid var(--border); }
  .trend-table tr:nth-child(even) td { background: var(--bg-light); }
  .trend-table .num { text-align: right; font-weight: 600; font-family: 'SF Mono', 'Consolas', monospace; }

  /* Page Break */
  .page-break { page-break-after: always; margin-bottom: 0; }

  /* Detail Section */
  .detail-section { margin-top: 20px; }
  .detail-section h2 { font-size: 16px; font-weight: 700; color: var(--primary); margin: 0 0 15px 0; padding-bottom: 8px; border-bottom: 2px solid var(--primary); }
  .detail-block { margin-bottom: 20px; border: 1px solid var(--border); border-radius: 6px; overflow: hidden; page-break-inside: avoid; }
  .detail-head { background: var(--primary); color: white; padding: 10px 15px; font-weight: 700; font-size: 12px; display: flex; justify-content: space-between; align-items: center; }
  .detail-head-stat { font-size: 10px; opacity: 0.85; font-weight: 500; }

  /* Data Table */
  table { width: 100%; border-collapse: collapse; font-size: 9px; }
  thead { display: table-header-group; }
  th { background: #f1f5f9; text-align: left; padding: 8px 6px; font-size: 9px; font-weight: 700; color: var(--primary); border-bottom: 2px solid var(--border); text-transform: uppercase; letter-spacing: 0.3px; }
  td { padding: 7px 6px; border-bottom: 1px solid #f1f5f9; }
  tr:nth-child(even) td { background: var(--bg-light); }
  .qty-cell { background: #ecfdf5 !important; text-align: center; font-weight: 700; color: var(--success); }
  .mono { font-family: 'SF Mono', 'Consolas', monospace; font-size: 9px; }

  /* Footer */
  .page-footer { position: fixed; bottom: 0; left: 0; right: 0; padding: 8px 15mm; font-size: 8px; color: var(--text-muted); border-top: 1px solid var(--border); background: white; display: flex; justify-content: space-between; }

  @media print {
    body { padding: 0; margin: 0; }
    .header-bar { margin: 0 0 20px 0; }
    .page-footer { position: fixed; bottom: 0; }
  }
</style></head><body>

<div class="header-bar">
  <div>
    <h1>Inbound Logistics Report</h1>
    <div class="subtitle">${esc(params.frequency)} Analysis • ${esc(dateRange)}</div>
  </div>
  <div class="meta">
    <strong>${esc(reportId)}</strong>
    Generated: ${esc(generatedAt)}
  </div>
</div>

<div class="exec-summary">
  <div class="exec-title">Executive Summary</div>
  <div class="kpi-grid">
    <div class="kpi-card highlight"><span class="kpi-num">${totalQty.toLocaleString()}</span><span class="kpi-lbl">Total Units Received</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniqueBOLs.size}</span><span class="kpi-lbl">Shipments (BOLs)</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniquePOs.size}</span><span class="kpi-lbl">Purchase Orders</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniqueFBPNs.size}</span><span class="kpi-lbl">Unique SKUs</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniqueProjects.size}</span><span class="kpi-lbl">Active Projects</span></div>
    <div class="kpi-card"><span class="kpi-num">${avgDailyQty.toLocaleString()}</span><span class="kpi-lbl">Avg Daily Volume</span></div>
  </div>

  <div class="insights-grid">
    <div class="insight-card">
      <h4>Top Projects by Volume</h4>
      <ul class="insight-list">
        ${topProjects.map(([proj, qty]) => `<li><span class="label">${esc(proj)}</span><span class="value">${qty.toLocaleString()} units</span></li>`).join('')}
        ${topProjects.length === 0 ? '<li><span class="label">No data</span></li>' : ''}
      </ul>
    </div>
    <div class="insight-card">
      <h4>Top Carriers</h4>
      <ul class="insight-list">
        ${topCarriers.map(([carrier, qty]) => `<li><span class="label">${esc(carrier)}</span><span class="value">${qty.toLocaleString()} units</span></li>`).join('')}
        ${topCarriers.length === 0 ? '<li><span class="label">No data</span></li>' : ''}
      </ul>
    </div>
  </div>

  <div class="trend-section">
    <h3>Daily Receiving Trend</h3>
    <table class="trend-table">
      <thead><tr><th>Date</th><th>Shipments</th><th>Quantity</th></tr></thead>
      <tbody>${dailyTrendHtml || '<tr><td colspan="3">No daily data</td></tr>'}</tbody>
    </table>
  </div>
</div>

<div class="page-break"></div>

<div class="detail-section">
  <h2>Detailed Receiving Log</h2>
  ${detailHtml}
</div>

<div class="page-footer">
  <span>CONFIDENTIAL — Internal Use Only</span>
  <span>Report: ${esc(reportId)} | Generated: ${esc(generatedAt)}</span>
</div>

</body></html>`;
}

function buildOutboundReportHtml_(params, reportData) {
  const rows = reportData.rows || [];
  const dateRange = reportData.dateRange || `${params.startDate} - ${params.endDate}`;
  const generatedAt = new Date().toLocaleString();
  const reportId = `OUT-${new Date().getFullYear()}-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`;

  // Calculate summary statistics
  let totalQty = 0;
  const uniqueCompanies = new Set();
  const uniqueTasks = new Set();
  const uniqueOrders = new Set();
  const uniqueFBPNs = new Set();
  const uniqueProjects = new Set();
  const dailyData = {};
  const companyQty = {};
  const projectQty = {};

  rows.forEach(r => {
    const qty = Number(r.qty || 0) || 0;
    totalQty += qty;
    if (r.company) { uniqueCompanies.add(r.company); companyQty[r.company] = (companyQty[r.company] || 0) + qty; }
    if (r.taskNumber) uniqueTasks.add(r.taskNumber);
    if (r.orderNumber) uniqueOrders.add(r.orderNumber);
    if (r.fbpn) uniqueFBPNs.add(r.fbpn);
    if (r.project) { uniqueProjects.add(r.project); projectQty[r.project] = (projectQty[r.project] || 0) + qty; }
    const dateKey = r.date || 'Unknown';
    dailyData[dateKey] = dailyData[dateKey] || { orders: new Set(), qty: 0 };
    dailyData[dateKey].orders.add(r.orderNumber);
    dailyData[dateKey].qty += qty;
  });

  // Top companies and projects for insights
  const topCompanies = Object.entries(companyQty).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const topProjects = Object.entries(projectQty).sort((a, b) => b[1] - a[1]).slice(0, 3);
  const avgDailyQty = Object.keys(dailyData).length > 0 ? Math.round(totalQty / Object.keys(dailyData).length) : 0;

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

  // Build daily trend table
  const dailyTrendHtml = Object.entries(dailyData).sort((a, b) => a[0].localeCompare(b[0])).map(([date, data]) =>
    `<tr><td>${esc(date)}</td><td class="num">${data.orders.size}</td><td class="num">${data.qty.toLocaleString()}</td></tr>`
  ).join('');

  // Build company summary table
  const companySummaryHtml = topCompanies.map(([company, qty], i) =>
    `<tr><td>${i + 1}</td><td>${esc(company)}</td><td class="num">${qty.toLocaleString()}</td><td class="num">${totalQty > 0 ? ((qty / totalQty) * 100).toFixed(1) : 0}%</td></tr>`
  ).join('');

  // Build detailed tables HTML
  let detailHtml = '';
  Object.keys(companyGroups).sort().forEach(companyKey => {
    const allItems = [];
    Object.values(companyGroups[companyKey]).forEach(items => allItems.push(...items));
    if (allItems.length === 0) return;
    const compTotal = allItems.reduce((sum, item) => sum + (Number(item.qty) || 0), 0);

    detailHtml += `
      <div class="detail-block">
        <div class="detail-head">
          <span>${esc(companyKey)}</span>
          <span class="detail-head-stat">${allItems.length} items • ${compTotal.toLocaleString()} units</span>
        </div>
        <table>
          <thead>
            <tr>
              <th style="width:10%">Date</th>
              <th style="width:10%">Task #</th>
              <th style="width:14%">Order #</th>
              <th style="width:14%">FBPN</th>
              <th style="width:8%">Qty</th>
              <th style="width:14%">Project</th>
              <th>Manufacturer</th>
            </tr>
          </thead>
          <tbody>
            ${allItems.map(item => `
              <tr>
                <td>${esc(item.date)}</td>
                <td class="mono">${esc(item.taskNumber)}</td>
                <td class="mono">${esc(item.orderNumber)}</td>
                <td class="mono">${esc(item.fbpn)}</td>
                <td class="qty-cell">${Number(item.qty || 0).toLocaleString()}</td>
                <td>${esc(item.project)}</td>
                <td>${esc(item.manufacturer)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>`;
  });

  return `<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Outbound Logistics Report</title>
<style>
  @page { size: A4 landscape; margin: 12mm 15mm; }
  :root { --primary: #065f46; --accent: #10b981; --success: #f97316; --warn: #dc2626; --border: #d1fae5; --bg-light: #f0fdf4; --text: #334155; --text-muted: #64748b; }
  * { box-sizing: border-box; }
  body { font-family: 'Inter', 'Segoe UI', system-ui, sans-serif; background: #fff; margin: 0; padding: 8mm; color: var(--text); font-size: 10px; line-height: 1.4; -webkit-print-color-adjust: exact; print-color-adjust: exact; }

  /* Header Bar */
  .header-bar { background: linear-gradient(135deg, var(--primary) 0%, #064e3b 100%); color: white; padding: 20px 25px; margin: -8mm -8mm 20px -8mm; display: flex; justify-content: space-between; align-items: center; }
  .header-bar h1 { margin: 0; font-size: 22px; font-weight: 700; letter-spacing: -0.5px; }
  .header-bar .subtitle { opacity: 0.85; font-size: 12px; margin-top: 4px; }
  .header-bar .meta { text-align: right; font-size: 10px; opacity: 0.9; }
  .header-bar .meta strong { display: block; font-size: 11px; }

  /* Executive Summary Card */
  .exec-summary { background: white; border: 1px solid var(--border); border-radius: 8px; padding: 20px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
  .exec-title { font-size: 14px; font-weight: 700; color: var(--primary); margin: 0 0 15px 0; padding-bottom: 10px; border-bottom: 2px solid var(--accent); display: flex; align-items: center; gap: 8px; }
  .exec-title::before { content: ''; width: 4px; height: 18px; background: var(--accent); border-radius: 2px; }

  /* KPI Grid */
  .kpi-grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 15px; margin-bottom: 20px; }
  .kpi-card { background: white; border: 1px solid var(--border); border-radius: 6px; padding: 15px; text-align: center; position: relative; overflow: hidden; }
  .kpi-card::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: var(--accent); }
  .kpi-card.highlight::before { background: var(--success); }
  .kpi-num { display: block; font-size: 26px; font-weight: 800; color: var(--primary); margin-bottom: 4px; }
  .kpi-lbl { font-size: 9px; font-weight: 600; text-transform: uppercase; color: var(--text-muted); letter-spacing: 0.5px; }

  /* Insights Panel */
  .insights-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
  .insight-card { background: var(--bg-light); border-radius: 6px; padding: 15px; }
  .insight-card h4 { margin: 0 0 12px 0; font-size: 11px; font-weight: 700; color: var(--primary); text-transform: uppercase; letter-spacing: 0.5px; }
  .insight-list { margin: 0; padding: 0; list-style: none; }
  .insight-list li { padding: 6px 0; border-bottom: 1px solid var(--border); font-size: 10px; display: flex; justify-content: space-between; }
  .insight-list li:last-child { border-bottom: none; }
  .insight-list .label { color: var(--text-muted); }
  .insight-list .value { font-weight: 700; color: var(--primary); }

  /* Trend Table */
  .trend-section { margin-bottom: 20px; }
  .trend-section h3 { font-size: 12px; font-weight: 700; color: var(--primary); margin: 0 0 10px 0; }
  .trend-table { width: 100%; border-collapse: collapse; font-size: 9px; }
  .trend-table th { background: var(--primary); color: white; padding: 8px 10px; text-align: left; font-weight: 600; }
  .trend-table td { padding: 7px 10px; border-bottom: 1px solid var(--border); }
  .trend-table tr:nth-child(even) td { background: var(--bg-light); }
  .trend-table .num { text-align: right; font-weight: 600; font-family: 'SF Mono', 'Consolas', monospace; }

  /* Page Break */
  .page-break { page-break-after: always; margin-bottom: 0; }

  /* Detail Section */
  .detail-section { margin-top: 20px; }
  .detail-section h2 { font-size: 16px; font-weight: 700; color: var(--primary); margin: 0 0 15px 0; padding-bottom: 8px; border-bottom: 2px solid var(--primary); }
  .detail-block { margin-bottom: 20px; border: 1px solid var(--border); border-radius: 6px; overflow: hidden; page-break-inside: avoid; }
  .detail-head { background: var(--primary); color: white; padding: 10px 15px; font-weight: 700; font-size: 12px; display: flex; justify-content: space-between; align-items: center; }
  .detail-head-stat { font-size: 10px; opacity: 0.85; font-weight: 500; }

  /* Data Table */
  table { width: 100%; border-collapse: collapse; font-size: 9px; }
  thead { display: table-header-group; }
  th { background: #ecfdf5; text-align: left; padding: 8px 6px; font-size: 9px; font-weight: 700; color: var(--primary); border-bottom: 2px solid var(--border); text-transform: uppercase; letter-spacing: 0.3px; }
  td { padding: 7px 6px; border-bottom: 1px solid #f0fdf4; }
  tr:nth-child(even) td { background: var(--bg-light); }
  .qty-cell { background: #fff7ed !important; text-align: center; font-weight: 700; color: var(--success); }
  .mono { font-family: 'SF Mono', 'Consolas', monospace; font-size: 9px; }

  /* Footer */
  .page-footer { position: fixed; bottom: 0; left: 0; right: 0; padding: 8px 15mm; font-size: 8px; color: var(--text-muted); border-top: 1px solid var(--border); background: white; display: flex; justify-content: space-between; }

  @media print {
    body { padding: 0; margin: 0; }
    .header-bar { margin: 0 0 20px 0; }
    .page-footer { position: fixed; bottom: 0; }
  }
</style></head><body>

<div class="header-bar">
  <div>
    <h1>Outbound Logistics Report</h1>
    <div class="subtitle">${esc(params.frequency)} Analysis • ${esc(dateRange)}</div>
  </div>
  <div class="meta">
    <strong>${esc(reportId)}</strong>
    Generated: ${esc(generatedAt)}
  </div>
</div>

<div class="exec-summary">
  <div class="exec-title">Executive Summary</div>
  <div class="kpi-grid">
    <div class="kpi-card highlight"><span class="kpi-num">${totalQty.toLocaleString()}</span><span class="kpi-lbl">Total Units Shipped</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniqueOrders.size}</span><span class="kpi-lbl">Total Orders</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniqueTasks.size}</span><span class="kpi-lbl">Fulfillment Tasks</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniqueFBPNs.size}</span><span class="kpi-lbl">Unique SKUs</span></div>
    <div class="kpi-card"><span class="kpi-num">${uniqueCompanies.size}</span><span class="kpi-lbl">Companies Served</span></div>
    <div class="kpi-card"><span class="kpi-num">${avgDailyQty.toLocaleString()}</span><span class="kpi-lbl">Avg Daily Volume</span></div>
  </div>

  <div class="insights-grid">
    <div class="insight-card">
      <h4>Top Companies by Volume</h4>
      <ul class="insight-list">
        ${topCompanies.map(([company, qty]) => `<li><span class="label">${esc(company)}</span><span class="value">${qty.toLocaleString()} units</span></li>`).join('')}
        ${topCompanies.length === 0 ? '<li><span class="label">No data</span></li>' : ''}
      </ul>
    </div>
    <div class="insight-card">
      <h4>Top Projects</h4>
      <ul class="insight-list">
        ${topProjects.map(([proj, qty]) => `<li><span class="label">${esc(proj)}</span><span class="value">${qty.toLocaleString()} units</span></li>`).join('')}
        ${topProjects.length === 0 ? '<li><span class="label">No data</span></li>' : ''}
      </ul>
    </div>
  </div>

  <div class="trend-section">
    <h3>Daily Shipping Trend</h3>
    <table class="trend-table">
      <thead><tr><th>Date</th><th>Orders</th><th>Quantity</th></tr></thead>
      <tbody>${dailyTrendHtml || '<tr><td colspan="3">No daily data</td></tr>'}</tbody>
    </table>
  </div>
</div>

<div class="page-break"></div>

<div class="detail-section">
  <h2>Detailed Shipment Log</h2>
  ${detailHtml}
</div>

<div class="page-footer">
  <span>CONFIDENTIAL — Internal Use Only</span>
  <span>Report: ${esc(reportId)} | Generated: ${esc(generatedAt)}</span>
</div>

</body></html>`;
}

// ============================================================================
// EXCEL GENERATION & STYLING - ENTERPRISE UI/UX
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
  return ['Date Received','Project','PO Number','BOL Number','Warehouse','Push #','Asset Type','Manufacturer','FBPN','MFPN','Qty Received','UOM','Carrier'];
}

function outboundXlsxHeaders_() {
  return ['Date','Company','Task Number','Order Number','Project','FBPN','Qty','Manufacturer'];
}

function saveRowsAsXlsxToReports_(rows, headers, filename, reportType, frequency) {
  try {
    const tmp = SpreadsheetApp.create('TEMP_' + filename.replace(/\.xlsx$/i, ''));

    // Create enterprise multi-sheet workbook
    if (reportType === 'Inbound') {
      createInboundEnterpriseWorkbook_(tmp, rows, headers);
    } else {
      createOutboundEnterpriseWorkbook_(tmp, rows, headers);
    }

    const xlsxBlob = exportSpreadsheetAsXlsx_(tmp.getId(), filename);
    const xlsxFile = saveReportXlsxToFolder_(xlsxBlob, reportType, frequency);

    DriveApp.getFileById(tmp.getId()).setTrashed(true);
    return xlsxFile;
  } catch (e) {
    Logger.log('Warning: XLSX export failed: ' + e.toString());
    return null;
  }
}

function createInboundEnterpriseWorkbook_(workbook, rows, headers) {
  // Calculate metrics
  let totalQty = 0;
  const uniqueProjects = new Set();
  const uniquePOs = new Set();
  const uniqueBOLs = new Set();
  const uniqueFBPNs = new Set();
  const uniqueCarriers = new Set();
  const dailyData = {};
  const projectData = {};
  const carrierData = {};

  rows.forEach(r => {
    const qty = Number(r.qty || 0) || 0;
    totalQty += qty;
    if (r.project) { uniqueProjects.add(r.project); projectData[r.project] = (projectData[r.project] || 0) + qty; }
    if (r.poNumber) uniquePOs.add(r.poNumber);
    if (r.bol) uniqueBOLs.add(r.bol);
    if (r.fbpn) uniqueFBPNs.add(r.fbpn);
    if (r.carrier) { uniqueCarriers.add(r.carrier); carrierData[r.carrier] = (carrierData[r.carrier] || 0) + qty; }
    const dateKey = r.dateReceived || 'Unknown';
    dailyData[dateKey] = dailyData[dateKey] || { bols: 0, qty: 0 };
    dailyData[dateKey].bols += 1;
    dailyData[dateKey].qty += qty;
  });

  // Sheet 1: Executive Summary
  const summarySheet = workbook.getSheets()[0];
  summarySheet.setName('Executive Summary');
  createExecutiveSummarySheet_(summarySheet, 'Inbound', {
    'Total Units Received': totalQty,
    'Total Shipments (BOLs)': uniqueBOLs.size,
    'Unique Purchase Orders': uniquePOs.size,
    'Unique SKUs (FBPNs)': uniqueFBPNs.size,
    'Active Projects': uniqueProjects.size,
    'Carriers Used': uniqueCarriers.size
  }, rows.length);

  // Sheet 2: Daily Trends
  const trendsSheet = workbook.insertSheet('Daily Trends');
  createDailyTrendsSheet_(trendsSheet, dailyData, 'Inbound');

  // Sheet 3: Project Analysis
  const projectSheet = workbook.insertSheet('Project Analysis');
  createAnalysisSheet_(projectSheet, projectData, 'Project', totalQty);

  // Sheet 4: Carrier Analysis
  const carrierSheet = workbook.insertSheet('Carrier Analysis');
  createAnalysisSheet_(carrierSheet, carrierData, 'Carrier', totalQty);

  // Sheet 5: Detailed Data
  const dataSheet = workbook.insertSheet('Detailed Data');
  createDetailedDataSheet_(dataSheet, rows, headers, 'Inbound');
}

function createOutboundEnterpriseWorkbook_(workbook, rows, headers) {
  // Calculate metrics
  let totalQty = 0;
  const uniqueCompanies = new Set();
  const uniqueTasks = new Set();
  const uniqueOrders = new Set();
  const uniqueFBPNs = new Set();
  const uniqueProjects = new Set();
  const dailyData = {};
  const companyData = {};
  const projectData = {};

  rows.forEach(r => {
    const qty = Number(r.qty || 0) || 0;
    totalQty += qty;
    if (r.company) { uniqueCompanies.add(r.company); companyData[r.company] = (companyData[r.company] || 0) + qty; }
    if (r.taskNumber) uniqueTasks.add(r.taskNumber);
    if (r.orderNumber) uniqueOrders.add(r.orderNumber);
    if (r.fbpn) uniqueFBPNs.add(r.fbpn);
    if (r.project) { uniqueProjects.add(r.project); projectData[r.project] = (projectData[r.project] || 0) + qty; }
    const dateKey = r.date || 'Unknown';
    dailyData[dateKey] = dailyData[dateKey] || { orders: 0, qty: 0 };
    dailyData[dateKey].orders += 1;
    dailyData[dateKey].qty += qty;
  });

  // Sheet 1: Executive Summary
  const summarySheet = workbook.getSheets()[0];
  summarySheet.setName('Executive Summary');
  createExecutiveSummarySheet_(summarySheet, 'Outbound', {
    'Total Units Shipped': totalQty,
    'Total Orders': uniqueOrders.size,
    'Fulfillment Tasks': uniqueTasks.size,
    'Unique SKUs (FBPNs)': uniqueFBPNs.size,
    'Companies Served': uniqueCompanies.size,
    'Projects': uniqueProjects.size
  }, rows.length);

  // Sheet 2: Daily Trends
  const trendsSheet = workbook.insertSheet('Daily Trends');
  createDailyTrendsSheet_(trendsSheet, dailyData, 'Outbound');

  // Sheet 3: Company Analysis
  const companySheet = workbook.insertSheet('Company Analysis');
  createAnalysisSheet_(companySheet, companyData, 'Company', totalQty);

  // Sheet 4: Project Analysis
  const projectSheet = workbook.insertSheet('Project Analysis');
  createAnalysisSheet_(projectSheet, projectData, 'Project', totalQty);

  // Sheet 5: Detailed Data
  const dataSheet = workbook.insertSheet('Detailed Data');
  createDetailedDataSheet_(dataSheet, rows, headers, 'Outbound');
}

function createExecutiveSummarySheet_(sheet, reportType, metrics, rowCount) {
  const isInbound = reportType === 'Inbound';
  const primaryColor = isInbound ? '#1e3a5f' : '#065f46';
  const accentColor = isInbound ? '#0ea5e9' : '#10b981';
  const lightBg = isInbound ? '#eff6ff' : '#f0fdf4';

  // Title Section
  sheet.getRange('A1:F1').merge().setValue(reportType.toUpperCase() + ' LOGISTICS REPORT').setFontSize(18).setFontWeight('bold').setFontColor(primaryColor).setBackground('#f8fafc');
  sheet.getRange('A2:F2').merge().setValue('Executive Summary').setFontSize(12).setFontColor('#64748b').setBackground('#f8fafc');
  sheet.getRange('A3:F3').merge().setValue('Generated: ' + new Date().toLocaleString()).setFontSize(9).setFontColor('#94a3b8').setBackground('#f8fafc');

  // KPI Section Header
  sheet.getRange('A5:F5').merge().setValue('KEY PERFORMANCE INDICATORS').setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground(primaryColor).setHorizontalAlignment('center');

  // KPI Cards
  let row = 6;
  const metricEntries = Object.entries(metrics);
  for (let i = 0; i < metricEntries.length; i += 3) {
    const cols = ['A', 'C', 'E'];
    for (let j = 0; j < 3 && i + j < metricEntries.length; j++) {
      const [label, value] = metricEntries[i + j];
      const col = cols[j];
      const colNum = col.charCodeAt(0) - 64;
      sheet.getRange(row, colNum).setValue(typeof value === 'number' ? value.toLocaleString() : value).setFontSize(22).setFontWeight('bold').setFontColor(primaryColor).setHorizontalAlignment('center');
      sheet.getRange(row + 1, colNum).setValue(label).setFontSize(9).setFontColor('#64748b').setHorizontalAlignment('center');
    }
    row += 3;
  }

  // Report Info Section
  row += 1;
  sheet.getRange(row, 1, 1, 6).merge().setValue('REPORT INFORMATION').setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground(primaryColor).setHorizontalAlignment('center');
  row++;
  sheet.getRange(row, 1).setValue('Report Type:').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(reportType + ' Logistics');
  sheet.getRange(row, 4).setValue('Total Records:').setFontWeight('bold');
  sheet.getRange(row, 5).setValue(rowCount.toLocaleString());
  row++;
  sheet.getRange(row, 1).setValue('Generated:').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(new Date().toLocaleString());
  sheet.getRange(row, 4).setValue('Status:').setFontWeight('bold');
  sheet.getRange(row, 5).setValue('Complete').setFontColor('#059669').setFontWeight('bold');

  // Styling
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);
  sheet.setFrozenRows(4);

  // Add borders
  sheet.getRange('A5:F' + row).setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
}

function createDailyTrendsSheet_(sheet, dailyData, reportType) {
  const isInbound = reportType === 'Inbound';
  const primaryColor = isInbound ? '#1e3a5f' : '#065f46';
  const headerColor = isInbound ? '#dbeafe' : '#d1fae5';
  const label = isInbound ? 'Shipments' : 'Orders';

  // Headers
  sheet.getRange('A1:D1').setValues([['Date', label, 'Quantity', '% of Total']]).setFontWeight('bold').setFontColor('#ffffff').setBackground(primaryColor).setHorizontalAlignment('center');

  // Calculate total
  let totalQty = 0;
  Object.values(dailyData).forEach(d => totalQty += d.qty);

  // Data
  const sortedDates = Object.entries(dailyData).sort((a, b) => a[0].localeCompare(b[0]));
  const values = sortedDates.map(([date, data]) => {
    const count = isInbound ? data.bols : data.orders;
    const pct = totalQty > 0 ? ((data.qty / totalQty) * 100).toFixed(1) + '%' : '0%';
    return [date, count, data.qty, pct];
  });

  if (values.length > 0) {
    sheet.getRange(2, 1, values.length, 4).setValues(values);

    // Alternate row colors
    for (let i = 0; i < values.length; i++) {
      if (i % 2 === 0) {
        sheet.getRange(i + 2, 1, 1, 4).setBackground(headerColor);
      }
    }

    // Number formatting
    sheet.getRange(2, 2, values.length, 1).setNumberFormat('#,##0');
    sheet.getRange(2, 3, values.length, 1).setNumberFormat('#,##0');
  }

  // Totals row
  const totalRow = values.length + 2;
  sheet.getRange(totalRow, 1).setValue('TOTAL').setFontWeight('bold').setBackground('#f1f5f9');
  sheet.getRange(totalRow, 2).setFormula('=SUM(B2:B' + (totalRow - 1) + ')').setFontWeight('bold').setBackground('#f1f5f9').setNumberFormat('#,##0');
  sheet.getRange(totalRow, 3).setFormula('=SUM(C2:C' + (totalRow - 1) + ')').setFontWeight('bold').setBackground('#f1f5f9').setNumberFormat('#,##0');
  sheet.getRange(totalRow, 4).setValue('100%').setFontWeight('bold').setBackground('#f1f5f9');

  // Styling
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.getRange('A1:D' + totalRow).setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('B:D').setHorizontalAlignment('center');
}

function createAnalysisSheet_(sheet, data, label, totalQty) {
  const primaryColor = '#1e3a5f';
  const headerColor = '#f1f5f9';

  // Headers
  sheet.getRange('A1:E1').setValues([['Rank', label, 'Quantity', '% of Total', 'Cumulative %']]).setFontWeight('bold').setFontColor('#ffffff').setBackground(primaryColor).setHorizontalAlignment('center');

  // Sort and prepare data
  const sorted = Object.entries(data).sort((a, b) => b[1] - a[1]);
  let cumulative = 0;
  const values = sorted.map(([name, qty], i) => {
    const pct = totalQty > 0 ? (qty / totalQty) * 100 : 0;
    cumulative += pct;
    return [i + 1, name, qty, pct.toFixed(1) + '%', cumulative.toFixed(1) + '%'];
  });

  if (values.length > 0) {
    sheet.getRange(2, 1, values.length, 5).setValues(values);

    // Alternate row colors
    for (let i = 0; i < values.length; i++) {
      if (i % 2 === 0) {
        sheet.getRange(i + 2, 1, 1, 5).setBackground(headerColor);
      }
    }

    // Highlight top 3
    for (let i = 0; i < Math.min(3, values.length); i++) {
      sheet.getRange(i + 2, 1, 1, 5).setFontWeight('bold').setBackground('#fef3c7');
    }

    // Number formatting
    sheet.getRange(2, 3, values.length, 1).setNumberFormat('#,##0');
  }

  // Totals row
  const totalRow = values.length + 2;
  sheet.getRange(totalRow, 1).setValue('').setBackground('#e2e8f0');
  sheet.getRange(totalRow, 2).setValue('TOTAL').setFontWeight('bold').setBackground('#e2e8f0');
  sheet.getRange(totalRow, 3).setFormula('=SUM(C2:C' + (totalRow - 1) + ')').setFontWeight('bold').setBackground('#e2e8f0').setNumberFormat('#,##0');
  sheet.getRange(totalRow, 4).setValue('100%').setFontWeight('bold').setBackground('#e2e8f0');
  sheet.getRange(totalRow, 5).setValue('').setBackground('#e2e8f0');

  // Styling
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 110);
  sheet.getRange('A1:E' + totalRow).setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A:A').setHorizontalAlignment('center');
  sheet.getRange('C:E').setHorizontalAlignment('center');
}

function createDetailedDataSheet_(sheet, rows, headers, reportType) {
  const isInbound = reportType === 'Inbound';
  const primaryColor = isInbound ? '#1e3a5f' : '#065f46';
  const headerColor = isInbound ? '#dbeafe' : '#d1fae5';

  // Headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setFontColor('#ffffff').setBackground(primaryColor).setHorizontalAlignment('center');

  // Data
  const values = [];
  if (isInbound) {
    rows.forEach(r => values.push([
      r.dateReceived || '', r.project || '', r.poNumber || '', r.bol || '',
      r.warehouse || '', r.push || '', r.assetType || '', r.manufacturer || '',
      r.fbpn || '', r.mfpn || '', r.qty || 0, r.uom || '', r.carrier || ''
    ]));
  } else {
    rows.forEach(r => values.push([
      r.date || '', r.company || '', r.taskNumber || '', r.orderNumber || '',
      r.project || '', r.fbpn || '', r.qty || 0, r.manufacturer || ''
    ]));
  }

  if (values.length > 0) {
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);

    // Alternate row colors
    for (let i = 0; i < values.length; i++) {
      if (i % 2 === 0) {
        sheet.getRange(i + 2, 1, 1, headers.length).setBackground(headerColor);
      }
    }

    // Number formatting for quantity column
    const qtyCol = isInbound ? 11 : 7;
    sheet.getRange(2, qtyCol, values.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
  }

  // Styling
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  // Add filter
  if (values.length > 0) {
    sheet.getRange(1, 1, values.length + 1, headers.length).createFilter();
  }

  // Borders
  const lastRow = Math.max(2, values.length + 1);
  sheet.getRange(1, 1, lastRow, headers.length).setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
}

function exportSpreadsheetAsXlsx_(spreadsheetId, filename) {
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  const blob = response.getBlob().setName(filename);
  return blob;
}

function saveReportXlsxToFolder_(xlsxBlob, reportType, frequency) {
  const rootFolder = DriveApp.getFolderById(FOLDERS.IMS_Reports);
  const typeFolder = getOrCreateFolder(rootFolder, reportType + ' Reports');
  const frequencyFolder = getOrCreateFolder(typeFolder, frequency);
  const yearFolder = getOrCreateFolder(frequencyFolder, String(new Date().getFullYear()));
  return yearFolder.createFile(xlsxBlob);
}