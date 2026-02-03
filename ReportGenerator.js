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
  const title = `${params.frequency} Inbound Report`;
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

  // Build daily summary table data
  const dailySummary = [];
  Object.keys(groups).sort().forEach(dateKey => {
    let dateTotal = 0;
    let dateLines = 0;
    const dateProjects = new Set();
    const datePOs = new Set();
    const dateBOLs = new Set();
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

  // Build summary table HTML
  let summaryTableHtml = `
    <div class="summary-section">
      <div class="summary-title">Daily Summary</div>
      <table class="summary-table">
        <thead>
          <tr>
            <th>Date</th>
            <th>Lines</th>
            <th>Units</th>
            <th>Projects</th>
            <th>POs</th>
            <th>BOLs</th>
          </tr>
        </thead>
        <tbody>
          ${dailySummary.map((d, idx) => `
            <tr class="${idx % 2 === 0 ? 'row-even' : 'row-odd'}">
              <td class="date-cell">${esc(d.date)}</td>
              <td>${d.lines.toLocaleString()}</td>
              <td class="qty-cell">${d.qty.toLocaleString()}</td>
              <td>${d.projects}</td>
              <td>${d.pos}</td>
              <td>${d.bols}</td>
            </tr>
          `).join('')}
          <tr class="total-row">
            <td><strong>TOTAL</strong></td>
            <td><strong>${rows.length.toLocaleString()}</strong></td>
            <td class="qty-cell"><strong>${totalQty.toLocaleString()}</strong></td>
            <td><strong>${uniqueProjects.size}</strong></td>
            <td><strong>${uniquePOs.size}</strong></td>
            <td><strong>${uniqueBOLs.size}</strong></td>
          </tr>
        </tbody>
      </table>
    </div>
  `;

  // Build detailed breakdown by day
  let bodyHtml = '<div class="detail-section"><div class="detail-title">Detailed Breakdown by Day</div>';
  Object.keys(groups).sort().forEach(dateKey => {
    let dateTotal = 0;
    let dateLines = 0;
    Object.values(groups[dateKey]).forEach(proj => {
      Object.values(proj).forEach(po => {
        Object.values(po).forEach(items => {
          dateLines += items.length;
          items.forEach(x => dateTotal += Number(x.qty || 0) || 0);
        });
      });
    });

    bodyHtml += `<div class="lvl-date"><div class="lvl-date-h"><span class="date-icon">&#128197;</span> ${esc(dateKey)}<span class="date-summary">${dateLines} lines | ${dateTotal.toLocaleString()} units</span></div>`;
    const projObj = groups[dateKey];
    Object.keys(projObj).sort().forEach(projKey => {
      let projTotal = 0;
      let projLines = 0;
      Object.values(projObj[projKey]).forEach(po => {
        Object.values(po).forEach(items => {
          projLines += items.length;
          items.forEach(x => projTotal += Number(x.qty || 0) || 0);
        });
      });

      bodyHtml += `<div class="lvl-project"><div class="lvl-project-h"><span class="project-label">PROJECT</span> ${esc(projKey)}<span class="project-summary">${projLines} lines | ${projTotal.toLocaleString()} units</span></div>`;
      const poObj = projObj[projKey];
      Object.keys(poObj).sort().forEach(poKey => {
        bodyHtml += `<div class="lvl-po"><div class="lvl-po-h"><span class="po-label">PO#</span> ${esc(poKey)}</div>`;
        const bolObj = poObj[poKey];
        Object.keys(bolObj).sort().forEach(bolKey => {
          const items = bolObj[bolKey];
          let bolTotal = 0;
          items.forEach(x => bolTotal += Number(x.qty || 0) || 0);
          bodyHtml += `
            <div class="lvl-bol">
              <div class="lvl-bol-h">
                <div><span class="bol-label">BOL</span> ${esc(bolKey)}</div>
                <div class="bol-stats"><span class="stat-badge">${items.length} lines</span><span class="stat-badge qty-badge">${bolTotal.toLocaleString()} units</span></div>
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
                  ${items.map((it, idx) => `
                    <tr class="${idx % 2 === 0 ? 'row-even' : 'row-odd'}">
                      <td>${esc(it.push)}</td>
                      <td>${esc(it.assetType || '')}</td>
                      <td class="mono">${esc(it.fbpn)}</td>
                      <td class="qty-cell">${esc(it.qty)}</td>
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
  bodyHtml += '</div>';

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    @page { size: letter; margin: 0.5in; }
    body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 9pt; line-height: 1.4; color: #1f2937; margin: 0; background: #fff; }

    /* Header Styles */
    .report-header { background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%); color: #fff; padding: 20px; margin: -0.5in -0.5in 20px -0.5in; }
    .header-content { display: flex; justify-content: space-between; align-items: center; max-width: 100%; }
    .title { font-size: 20pt; font-weight: 700; letter-spacing: -0.5px; }
    .subtitle { font-size: 9pt; opacity: 0.9; margin-top: 4px; }
    .header-right { text-align: right; }
    .date-range { font-size: 10pt; font-weight: 600; }
    .generated { font-size: 8pt; opacity: 0.8; margin-top: 4px; }

    /* Summary Section */
    .summary-section { margin-bottom: 24px; page-break-after: auto; }
    .summary-title { font-size: 14pt; font-weight: 700; color: #1e40af; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #1e40af; }
    .summary-table { width: 100%; border-collapse: collapse; margin-bottom: 16px; }
    .summary-table th { text-align: center; font-size: 9pt; background: #1e40af; color: #fff; padding: 10px 8px; border: 1px solid #1e3a8a; font-weight: 600; text-transform: uppercase; }
    .summary-table td { padding: 10px 8px; border: 1px solid #e2e8f0; text-align: center; font-weight: 500; }
    .summary-table .date-cell { font-weight: 600; color: #1e3a8a; text-align: left; padding-left: 12px; }
    .summary-table .row-even td { background: #f8fafc; }
    .summary-table .row-odd td { background: #ffffff; }
    .summary-table .total-row td { background: #e0e7ff; border-top: 2px solid #1e40af; }
    .summary-table .qty-cell { color: #166534; font-weight: 700; }

    /* Detail Section */
    .detail-section { margin-top: 20px; }
    .detail-title { font-size: 14pt; font-weight: 700; color: #1e40af; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #1e40af; }

    /* Date Level - Primary grouping */
    .lvl-date { margin-bottom: 20px; }
    .lvl-date-h { font-size: 12pt; font-weight: 700; padding: 12px 16px; background: linear-gradient(135deg, #1e3a5f 0%, #2563eb 100%); color: #fff; border-radius: 8px 8px 0 0; display: flex; justify-content: space-between; align-items: center; page-break-after: avoid; }
    .date-icon { margin-right: 8px; }
    .date-summary { font-size: 9pt; font-weight: 500; opacity: 0.9; background: rgba(255,255,255,0.2); padding: 4px 10px; border-radius: 12px; }

    /* Project Level */
    .lvl-project { margin: 0 0 12px 0; border-left: 4px solid #2563eb; background: #f8fafc; }
    .lvl-project-h { font-weight: 600; padding: 10px 16px; font-size: 10pt; display: flex; align-items: center; gap: 8px; background: #e0e7ff; color: #1e3a8a; page-break-after: avoid; }
    .project-label { background: #1e40af; color: #fff; font-size: 7pt; padding: 2px 6px; border-radius: 4px; font-weight: 700; }
    .project-summary { margin-left: auto; font-size: 8pt; color: #4b5563; font-weight: 500; }

    /* PO Level */
    .lvl-po { page-break-inside: avoid; margin: 8px 16px 12px 16px; }
    .lvl-po-h { font-weight: 600; padding: 8px 12px; background: #7c3aed; color: #fff; border-radius: 6px 6px 0 0; font-size: 9pt; display: flex; align-items: center; gap: 8px; }
    .po-label { background: rgba(255,255,255,0.25); font-size: 7pt; padding: 2px 6px; border-radius: 4px; font-weight: 700; }

    /* BOL Level */
    .lvl-bol { page-break-inside: avoid; margin: 0 0 8px 0; border: 1px solid #e5e7eb; border-radius: 0 0 6px 6px; overflow: hidden; }
    .lvl-bol-h { display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; font-size: 9pt; font-weight: 600; background: #faf5ff; color: #6b21a8; border-bottom: 1px solid #e9d5ff; }
    .bol-label { background: #a855f7; color: #fff; font-size: 7pt; padding: 2px 6px; border-radius: 4px; font-weight: 700; margin-right: 6px; }
    .bol-stats { display: flex; gap: 6px; }
    .stat-badge { font-size: 8pt; padding: 3px 8px; background: #f3e8ff; border-radius: 10px; color: #7c3aed; font-weight: 600; }
    .qty-badge { background: #dcfce7; color: #166534; }

    /* Table Styles */
    table { width: 100%; border-collapse: collapse; table-layout: fixed; }
    th { text-align: center; font-size: 8pt; background: #f1f5f9; color: #475569; padding: 8px 6px; border: 1px solid #cbd5e1; font-weight: 600; text-transform: uppercase; letter-spacing: 0.3px; }
    td { padding: 8px 6px; border: 1px solid #e2e8f0; text-align: center; font-weight: 500; word-wrap: break-word; }
    .row-even td { background: #ffffff; }
    .row-odd td { background: #f8fafc; }
    .mono { font-family: 'Consolas', 'Monaco', monospace; font-size: 8.5pt; color: #0f172a; }
    .qty-cell { font-weight: 700; color: #166534; }

    /* Print optimizations */
    @media print {
      .report-header { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .summary-table th, .summary-table .total-row td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .lvl-date-h, .lvl-project-h, .lvl-po-h, .lvl-bol-h, th { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style></head><body>
  <div class="report-header">
    <div class="header-content">
      <div>
        <div class="title">INBOUND REPORT</div>
        <div class="subtitle">${esc(params.frequency)} Summary</div>
      </div>
      <div class="header-right">
        <div class="date-range">${esc(dateRange)}</div>
        <div class="generated">Generated: ${esc(generatedAt)}</div>
      </div>
    </div>
  </div>
  ${summaryTableHtml}
  ${bodyHtml}
</body></html>`;
}

function buildOutboundReportHtml_(params, reportData) {
  const rows = reportData.rows || [];
  const title = `${params.frequency} Outbound Report`;
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

  // Build daily summary table data
  const dailySummary = [];
  Object.keys(groups).sort().forEach(dateKey => {
    let dateTotal = 0;
    let dateLines = 0;
    const dateCompanies = new Set();
    const dateTasks = new Set();
    const dateOrders = new Set();
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

  // Build summary table HTML
  let summaryTableHtml = `
    <div class="summary-section">
      <div class="summary-title">Daily Summary</div>
      <table class="summary-table">
        <thead>
          <tr>
            <th>Date</th>
            <th>Lines</th>
            <th>Units</th>
            <th>Companies</th>
            <th>Tasks</th>
            <th>Orders</th>
          </tr>
        </thead>
        <tbody>
          ${dailySummary.map((d, idx) => `
            <tr class="${idx % 2 === 0 ? 'row-even' : 'row-odd'}">
              <td class="date-cell">${esc(d.date)}</td>
              <td>${d.lines.toLocaleString()}</td>
              <td class="qty-cell">${d.qty.toLocaleString()}</td>
              <td>${d.companies}</td>
              <td>${d.tasks}</td>
              <td>${d.orders}</td>
            </tr>
          `).join('')}
          <tr class="total-row">
            <td><strong>TOTAL</strong></td>
            <td><strong>${rows.length.toLocaleString()}</strong></td>
            <td class="qty-cell"><strong>${totalQty.toLocaleString()}</strong></td>
            <td><strong>${uniqueCompanies.size}</strong></td>
            <td><strong>${uniqueTasks.size}</strong></td>
            <td><strong>${uniqueOrders.size}</strong></td>
          </tr>
        </tbody>
      </table>
    </div>
  `;

  // Build detailed breakdown by day
  let bodyHtml = '<div class="detail-section"><div class="detail-title">Detailed Breakdown by Day</div>';
  Object.keys(groups).sort().forEach(dateKey => {
    let dateTotal = 0;
    let dateLines = 0;
    Object.values(groups[dateKey]).forEach(company => {
      Object.values(company).forEach(items => {
        dateLines += items.length;
        items.forEach(x => dateTotal += Number(x.qty || 0) || 0);
      });
    });

    bodyHtml += `<div class="lvl-date"><div class="lvl-date-h"><span class="date-icon">&#128197;</span> ${esc(dateKey)}<span class="date-summary">${dateLines} lines | ${dateTotal.toLocaleString()} units</span></div>`;
    const companyObj = groups[dateKey];
    Object.keys(companyObj).sort().forEach(companyKey => {
      let companyTotal = 0;
      let companyLines = 0;
      Object.values(companyObj[companyKey]).forEach(items => {
        companyLines += items.length;
        items.forEach(x => companyTotal += Number(x.qty || 0) || 0);
      });

      bodyHtml += `<div class="lvl-company"><div class="lvl-company-h"><span class="company-label">COMPANY</span> ${esc(companyKey)}<span class="company-summary">${companyLines} lines | ${companyTotal.toLocaleString()} units</span></div>`;
      const taskObj = companyObj[companyKey];
      Object.keys(taskObj).sort().forEach(taskKey => {
        const items = taskObj[taskKey];
        let taskTotal = 0;
        items.forEach(x => taskTotal += Number(x.qty || 0) || 0);
        bodyHtml += `
          <div class="lvl-task">
            <div class="lvl-task-h">
              <div><span class="task-label">TASK#</span> ${esc(taskKey)}</div>
              <div class="task-stats"><span class="stat-badge">${items.length} lines</span><span class="stat-badge qty-badge">${taskTotal.toLocaleString()} units</span></div>
            </div>
            <table>
              <thead>
                <tr>
                  <th>Order #</th>
                  <th>FBPN</th>
                  <th>Qty</th>
                  <th>Project</th>
                </tr>
              </thead>
              <tbody>
                ${items.map((it, idx) => `
                  <tr class="${idx % 2 === 0 ? 'row-even' : 'row-odd'}">
                    <td class="mono">${esc(it.orderNumber)}</td>
                    <td class="mono">${esc(it.fbpn)}</td>
                    <td class="qty-cell">${esc(it.qty)}</td>
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
  bodyHtml += '</div>';

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    @page { size: letter; margin: 0.5in; }
    body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 9pt; line-height: 1.4; color: #1f2937; margin: 0; background: #fff; }

    /* Header Styles - Teal/Green theme for Outbound */
    .report-header { background: linear-gradient(135deg, #065f46 0%, #10b981 100%); color: #fff; padding: 20px; margin: -0.5in -0.5in 20px -0.5in; }
    .header-content { display: flex; justify-content: space-between; align-items: center; max-width: 100%; }
    .title { font-size: 20pt; font-weight: 700; letter-spacing: -0.5px; }
    .subtitle { font-size: 9pt; opacity: 0.9; margin-top: 4px; }
    .header-right { text-align: right; }
    .date-range { font-size: 10pt; font-weight: 600; }
    .generated { font-size: 8pt; opacity: 0.8; margin-top: 4px; }

    /* Summary Section */
    .summary-section { margin-bottom: 24px; page-break-after: auto; }
    .summary-title { font-size: 14pt; font-weight: 700; color: #065f46; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #065f46; }
    .summary-table { width: 100%; border-collapse: collapse; margin-bottom: 16px; }
    .summary-table th { text-align: center; font-size: 9pt; background: #065f46; color: #fff; padding: 10px 8px; border: 1px solid #064e3b; font-weight: 600; text-transform: uppercase; }
    .summary-table td { padding: 10px 8px; border: 1px solid #e2e8f0; text-align: center; font-weight: 500; }
    .summary-table .date-cell { font-weight: 600; color: #064e3b; text-align: left; padding-left: 12px; }
    .summary-table .row-even td { background: #f0fdf4; }
    .summary-table .row-odd td { background: #ffffff; }
    .summary-table .total-row td { background: #d1fae5; border-top: 2px solid #065f46; }
    .summary-table .qty-cell { color: #c2410c; font-weight: 700; }

    /* Detail Section */
    .detail-section { margin-top: 20px; }
    .detail-title { font-size: 14pt; font-weight: 700; color: #065f46; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #065f46; }

    /* Date Level - Primary grouping */
    .lvl-date { margin-bottom: 20px; }
    .lvl-date-h { font-size: 12pt; font-weight: 700; padding: 12px 16px; background: linear-gradient(135deg, #064e3b 0%, #059669 100%); color: #fff; border-radius: 8px 8px 0 0; display: flex; justify-content: space-between; align-items: center; page-break-after: avoid; }
    .date-icon { margin-right: 8px; }
    .date-summary { font-size: 9pt; font-weight: 500; opacity: 0.9; background: rgba(255,255,255,0.2); padding: 4px 10px; border-radius: 12px; }

    /* Company Level */
    .lvl-company { margin: 0 0 12px 0; border-left: 4px solid #10b981; background: #f0fdf4; }
    .lvl-company-h { font-weight: 600; padding: 10px 16px; font-size: 10pt; display: flex; align-items: center; gap: 8px; background: #d1fae5; color: #064e3b; page-break-after: avoid; }
    .company-label { background: #065f46; color: #fff; font-size: 7pt; padding: 2px 6px; border-radius: 4px; font-weight: 700; }
    .company-summary { margin-left: auto; font-size: 8pt; color: #4b5563; font-weight: 500; }

    /* Task Level */
    .lvl-task { page-break-inside: avoid; margin: 8px 16px 12px 16px; border: 1px solid #d1fae5; border-radius: 6px; overflow: hidden; }
    .lvl-task-h { display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; font-size: 9pt; font-weight: 600; background: #f97316; color: #fff; }
    .task-label { background: rgba(255,255,255,0.25); font-size: 7pt; padding: 2px 6px; border-radius: 4px; font-weight: 700; margin-right: 6px; }
    .task-stats { display: flex; gap: 6px; }
    .stat-badge { font-size: 8pt; padding: 3px 8px; background: rgba(255,255,255,0.25); border-radius: 10px; color: #fff; font-weight: 600; }
    .qty-badge { background: rgba(255,255,255,0.35); }

    /* Table Styles */
    table { width: 100%; border-collapse: collapse; table-layout: fixed; }
    th { text-align: center; font-size: 8pt; background: #f1f5f9; color: #475569; padding: 8px 6px; border: 1px solid #cbd5e1; font-weight: 600; text-transform: uppercase; letter-spacing: 0.3px; }
    td { padding: 8px 6px; border: 1px solid #e2e8f0; text-align: center; font-weight: 500; word-wrap: break-word; }
    .row-even td { background: #ffffff; }
    .row-odd td { background: #f0fdf4; }
    .mono { font-family: 'Consolas', 'Monaco', monospace; font-size: 8.5pt; color: #0f172a; }
    .qty-cell { font-weight: 700; color: #c2410c; }

    /* Print optimizations */
    @media print {
      .report-header { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .summary-table th, .summary-table .total-row td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .lvl-date-h, .lvl-company-h, .lvl-task-h, th { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style></head><body>
  <div class="report-header">
    <div class="header-content">
      <div>
        <div class="title">OUTBOUND REPORT</div>
        <div class="subtitle">${esc(params.frequency)} Summary</div>
      </div>
      <div class="header-right">
        <div class="date-range">${esc(dateRange)}</div>
        <div class="generated">Generated: ${esc(generatedAt)}</div>
      </div>
    </div>
  </div>
  ${summaryTableHtml}
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