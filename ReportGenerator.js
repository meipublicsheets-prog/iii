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
// HTML PORTRAIT GENERATION
// ============================================================================

function buildInboundReportHtml_(params, reportData) {
  const rows = reportData.rows || [];
  const dateRange = reportData.dateRange || `${params.startDate} - ${params.endDate}`;
  const generatedAt = new Date().toLocaleString();

  // Calculate summary statistics
  let totalQty = 0;
  const uniqueProjects = new Set();
  const uniquePOs = new Set();
  const uniqueBOLs = new Set();
  rows.forEach(r => {
    totalQty += Number(r.qty || 0) || 0;
    if (r.project) uniqueProjects.add(r.project);
    if (r.poNumber) uniquePOs.add(r.poNumber);
    if (r.bol) uniqueBOLs.add(r.bol);
  });

  // Group by date -> project -> PO -> BOL
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

  // Build daily summary data
  const dailySummary = [];
  Object.keys(groups).sort().forEach(dateKey => {
    let dateTotal = 0, dateLines = 0;
    const dateProjects = new Set(), datePOs = new Set(), dateBOLs = new Set();
    Object.keys(groups[dateKey]).forEach(proj => {
      dateProjects.add(proj);
      Object.keys(groups[dateKey][proj]).forEach(po => {
        datePOs.add(po);
        Object.keys(groups[dateKey][proj][po]).forEach(bol => {
          dateBOLs.add(bol);
          const items = groups[dateKey][proj][po][bol];
          dateLines += items.length;
          items.forEach(x => dateTotal += Number(x.qty || 0) || 0);
        });
      });
    });
    dailySummary.push({ date: dateKey, lines: dateLines, qty: dateTotal, projects: dateProjects.size, pos: datePOs.size, bols: dateBOLs.size });
  });

  // Build daily summary table
  let dailyTableHtml = `
    <table class="daily-table">
      <thead>
        <tr><th>Date</th><th>Lines</th><th>Units</th><th>Projects</th><th>POs</th><th>BOLs</th></tr>
      </thead>
      <tbody>
        ${dailySummary.map(d => `
          <tr>
            <td>${esc(d.date)}</td>
            <td>${d.lines.toLocaleString()}</td>
            <td>${d.qty.toLocaleString()}</td>
            <td>${d.projects}</td>
            <td>${d.pos}</td>
            <td>${d.bols}</td>
          </tr>
        `).join('')}
        <tr class="total-row">
          <td>TOTAL</td>
          <td>${rows.length.toLocaleString()}</td>
          <td>${totalQty.toLocaleString()}</td>
          <td>${uniqueProjects.size}</td>
          <td>${uniquePOs.size}</td>
          <td>${uniqueBOLs.size}</td>
        </tr>
      </tbody>
    </table>`;

  // Build detailed breakdown by project
  let detailHtml = '';
  Object.keys(groups).sort().forEach(dateKey => {
    Object.keys(groups[dateKey]).sort().forEach(projKey => {
      detailHtml += `<div class="project-section"><div class="project-title">PROJECT: ${esc(projKey)}</div>`;
      const poObj = groups[dateKey][projKey];
      Object.keys(poObj).sort().forEach(poKey => {
        detailHtml += `<div class="po-container"><div class="po-title">PO: ${esc(poKey)}</div>`;
        const bolObj = poObj[poKey];
        Object.keys(bolObj).sort().forEach(bolKey => {
          const items = bolObj[bolKey];
          let bolTotal = 0;
          items.forEach(x => bolTotal += Number(x.qty || 0) || 0);
          detailHtml += `
            <div class="bol-section">
              <div class="bol-title">BOL: ${esc(bolKey)} <span class="bol-summary">${items.length} lines | ${bolTotal.toLocaleString()} units</span></div>
              <table class="data-table">
                <thead>
                  <tr><th>Push #</th><th>Asset Type</th><th>FBPN</th><th>Qty</th><th>Carrier</th></tr>
                </thead>
                <tbody>
                  ${items.map(it => `
                    <tr>
                      <td>${esc(it.push)}</td>
                      <td>${esc(it.assetType || '')}</td>
                      <td><span class="pn-style">${esc(it.fbpn)}</span></td>
                      <td class="qty-val">${esc(it.qty)}</td>
                      <td>${esc(it.carrier)}</td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>`;
        });
        detailHtml += `</div>`;
      });
      detailHtml += `</div>`;
    });
  });

  return `<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Inbound Report</title>
<style>
  :root { --primary: #1e293b; --accent: #2563eb; --success: #059669; --bg: #f1f5f9; --border: #cbd5e1; }
  @page { size: letter; margin: 0.5in; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; color: #334155; margin: 0; padding: 20px; font-size: 10pt; }

  header { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 4px solid var(--primary); padding-bottom: 15px; margin-bottom: 25px; }
  h1 { margin: 0; font-size: 22px; color: var(--primary); text-transform: uppercase; letter-spacing: 1px; }
  .subtitle { margin: 5px 0 0 0; color: var(--accent); font-weight: 600; font-size: 12px; }
  .meta { text-align: right; font-size: 11px; color: #64748b; }

  .stat-banner { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
  .stat-card { background: #f8fafc; padding: 12px; border: 1px solid var(--border); text-align: center; border-radius: 4px; }
  .stat-val { display: block; font-size: 20px; font-weight: 800; color: var(--accent); }
  .stat-label { font-size: 10px; text-transform: uppercase; font-weight: 600; color: #64748b; }

  .section-title { font-size: 14px; font-weight: 700; color: var(--primary); margin: 25px 0 15px 0; padding-bottom: 8px; border-bottom: 2px solid var(--primary); }

  .daily-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
  .daily-table th { background: #f8fafc; padding: 10px; border: 1px solid var(--border); font-size: 10px; text-transform: uppercase; color: #64748b; }
  .daily-table td { padding: 10px; border: 1px solid var(--border); text-align: center; font-size: 11px; }
  .daily-table .total-row { background: #f1f5f9; font-weight: 800; }

  .project-section { margin-bottom: 25px; page-break-inside: avoid; }
  .project-title { background: var(--primary); color: white; padding: 10px 15px; font-size: 13px; font-weight: 700; border-radius: 4px 4px 0 0; }
  .po-container { border-left: 3px solid var(--accent); margin-left: 10px; padding-left: 15px; margin-top: 12px; }
  .po-title { font-weight: 700; color: var(--accent); font-size: 12px; margin-bottom: 8px; }
  .bol-section { margin-bottom: 12px; }
  .bol-title { font-weight: 600; font-size: 11px; color: #475569; margin-bottom: 6px; padding: 6px 10px; background: #f8fafc; border-radius: 4px; }
  .bol-summary { float: right; font-weight: 500; color: #64748b; }

  .data-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
  .data-table th { text-align: left; font-size: 9px; text-transform: uppercase; color: #64748b; padding: 8px; border-bottom: 1px solid var(--border); background: #fafafa; }
  .data-table td { padding: 8px; font-size: 11px; border-bottom: 1px dashed #e2e8f0; }
  .pn-style { font-family: 'Consolas', monospace; background: #f1f5f9; padding: 2px 5px; border-radius: 3px; font-size: 10px; }
  .qty-val { font-weight: 700; color: var(--success); }

  @media print {
    body { padding: 0; }
    .stat-card, .project-title, .daily-table th, .daily-table .total-row { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  }
</style></head><body>

<header>
  <div>
    <h1>Inbound Report</h1>
    <p class="subtitle">${esc(params.frequency)} Summary - ${esc(dateRange)}</p>
  </div>
  <div class="meta">Generated: ${esc(generatedAt)}<br>Source: Master_Log</div>
</header>

<div class="stat-banner">
  <div class="stat-card"><span class="stat-val">${rows.length.toLocaleString()}</span><span class="stat-label">Total Lines</span></div>
  <div class="stat-card"><span class="stat-val">${totalQty.toLocaleString()}</span><span class="stat-label">Total Units</span></div>
  <div class="stat-card"><span class="stat-val">${uniqueBOLs.size}</span><span class="stat-label">Total BOLs</span></div>
</div>

<div class="section-title">Daily Activity Summary</div>
${dailyTableHtml}

<div class="section-title">Detailed Breakdown by Project</div>
${detailHtml}

</body></html>`;
}

function buildOutboundReportHtml_(params, reportData) {
  const rows = reportData.rows || [];
  const dateRange = reportData.dateRange || `${params.startDate} - ${params.endDate}`;
  const generatedAt = new Date().toLocaleString();

  // Calculate summary statistics
  let totalQty = 0;
  const uniqueCompanies = new Set();
  const uniqueTasks = new Set();
  const uniqueOrders = new Set();
  rows.forEach(r => {
    totalQty += Number(r.qty || 0) || 0;
    if (r.company) uniqueCompanies.add(r.company);
    if (r.taskNumber) uniqueTasks.add(r.taskNumber);
    if (r.orderNumber) uniqueOrders.add(r.orderNumber);
  });

  // Group by date -> company -> task
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

  // Build daily summary data
  const dailySummary = [];
  Object.keys(groups).sort().forEach(dateKey => {
    let dateTotal = 0, dateLines = 0;
    const dateCompanies = new Set(), dateTasks = new Set(), dateOrders = new Set();
    Object.keys(groups[dateKey]).forEach(company => {
      dateCompanies.add(company);
      Object.keys(groups[dateKey][company]).forEach(task => {
        dateTasks.add(task);
        const items = groups[dateKey][company][task];
        dateLines += items.length;
        items.forEach(x => {
          dateTotal += Number(x.qty || 0) || 0;
          if (x.orderNumber) dateOrders.add(x.orderNumber);
        });
      });
    });
    dailySummary.push({ date: dateKey, lines: dateLines, qty: dateTotal, companies: dateCompanies.size, tasks: dateTasks.size, orders: dateOrders.size });
  });

  // Build daily summary table
  let dailyTableHtml = `
    <table class="daily-table">
      <thead>
        <tr><th>Date</th><th>Lines</th><th>Units</th><th>Companies</th><th>Tasks</th><th>Orders</th></tr>
      </thead>
      <tbody>
        ${dailySummary.map(d => `
          <tr>
            <td>${esc(d.date)}</td>
            <td>${d.lines.toLocaleString()}</td>
            <td>${d.qty.toLocaleString()}</td>
            <td>${d.companies}</td>
            <td>${d.tasks}</td>
            <td>${d.orders}</td>
          </tr>
        `).join('')}
        <tr class="total-row">
          <td>TOTAL</td>
          <td>${rows.length.toLocaleString()}</td>
          <td>${totalQty.toLocaleString()}</td>
          <td>${uniqueCompanies.size}</td>
          <td>${uniqueTasks.size}</td>
          <td>${uniqueOrders.size}</td>
        </tr>
      </tbody>
    </table>`;

  // Build detailed breakdown by company
  let detailHtml = '';
  Object.keys(groups).sort().forEach(dateKey => {
    Object.keys(groups[dateKey]).sort().forEach(companyKey => {
      detailHtml += `<div class="company-section"><div class="company-title">COMPANY: ${esc(companyKey)}</div>`;
      const taskObj = groups[dateKey][companyKey];
      Object.keys(taskObj).sort().forEach(taskKey => {
        const items = taskObj[taskKey];
        let taskTotal = 0;
        items.forEach(x => taskTotal += Number(x.qty || 0) || 0);
        detailHtml += `
          <div class="task-container">
            <div class="task-title">Task #${esc(taskKey)} <span class="task-summary">${items.length} lines | ${taskTotal.toLocaleString()} units</span></div>
            <table class="data-table">
              <thead>
                <tr><th>Order #</th><th>FBPN</th><th>Qty</th><th>Project</th></tr>
              </thead>
              <tbody>
                ${items.map(it => `
                  <tr>
                    <td><span class="pn-style">${esc(it.orderNumber)}</span></td>
                    <td><span class="pn-style">${esc(it.fbpn)}</span></td>
                    <td class="qty-val">${esc(it.qty)}</td>
                    <td>${esc(it.project)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>`;
      });
      detailHtml += `</div>`;
    });
  });

  return `<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Outbound Report</title>
<style>
  :root { --primary: #065f46; --accent: #10b981; --warning: #f97316; --bg: #f0fdf4; --border: #a7f3d0; }
  @page { size: letter; margin: 0.5in; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; color: #334155; margin: 0; padding: 20px; font-size: 10pt; }

  header { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 4px solid var(--primary); padding-bottom: 15px; margin-bottom: 25px; }
  h1 { margin: 0; font-size: 22px; color: var(--primary); text-transform: uppercase; letter-spacing: 1px; }
  .subtitle { margin: 5px 0 0 0; color: var(--accent); font-weight: 600; font-size: 12px; }
  .meta { text-align: right; font-size: 11px; color: #64748b; }

  .stat-banner { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
  .stat-card { background: var(--bg); padding: 12px; border: 1px solid var(--border); text-align: center; border-radius: 4px; }
  .stat-val { display: block; font-size: 20px; font-weight: 800; color: var(--primary); }
  .stat-label { font-size: 10px; text-transform: uppercase; font-weight: 600; color: #64748b; }

  .section-title { font-size: 14px; font-weight: 700; color: var(--primary); margin: 25px 0 15px 0; padding-bottom: 8px; border-bottom: 2px solid var(--primary); }

  .daily-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
  .daily-table th { background: #f8fafc; padding: 10px; border: 1px solid #cbd5e1; font-size: 10px; text-transform: uppercase; color: #64748b; }
  .daily-table td { padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-size: 11px; }
  .daily-table .total-row { background: var(--bg); font-weight: 800; }

  .company-section { margin-bottom: 25px; page-break-inside: avoid; }
  .company-title { background: var(--primary); color: white; padding: 10px 15px; font-size: 13px; font-weight: 700; border-radius: 4px 4px 0 0; }
  .task-container { border-left: 3px solid var(--warning); margin-left: 10px; padding-left: 15px; margin-top: 12px; margin-bottom: 15px; }
  .task-title { font-weight: 700; color: var(--warning); font-size: 12px; margin-bottom: 8px; }
  .task-summary { float: right; font-weight: 500; color: #64748b; font-size: 11px; }

  .data-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
  .data-table th { text-align: left; font-size: 9px; text-transform: uppercase; color: #64748b; padding: 8px; border-bottom: 1px solid #cbd5e1; background: #fafafa; }
  .data-table td { padding: 8px; font-size: 11px; border-bottom: 1px dashed #e2e8f0; }
  .pn-style { font-family: 'Consolas', monospace; background: #f1f5f9; padding: 2px 5px; border-radius: 3px; font-size: 10px; }
  .qty-val { font-weight: 700; color: var(--warning); }

  @media print {
    body { padding: 0; }
    .stat-card, .company-title, .daily-table th, .daily-table .total-row { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  }
</style></head><body>

<header>
  <div>
    <h1>Outbound Report</h1>
    <p class="subtitle">${esc(params.frequency)} Summary - ${esc(dateRange)}</p>
  </div>
  <div class="meta">Generated: ${esc(generatedAt)}<br>Source: OutboundLog</div>
</header>

<div class="stat-banner">
  <div class="stat-card"><span class="stat-val">${rows.length.toLocaleString()}</span><span class="stat-label">Total Lines</span></div>
  <div class="stat-card"><span class="stat-val">${totalQty.toLocaleString()}</span><span class="stat-label">Total Units</span></div>
  <div class="stat-card"><span class="stat-val">${uniqueOrders.size}</span><span class="stat-label">Total Orders</span></div>
</div>

<div class="section-title">Daily Activity Summary</div>
${dailyTableHtml}

<div class="section-title">Detailed Breakdown by Company</div>
${detailHtml}

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