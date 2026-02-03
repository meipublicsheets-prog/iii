function backfillMasterLogProjectFromPOMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shML = ss.getSheetByName('Master_Log');
  if (!shML) throw new Error('Master_Log sheet not found');

  const shPO = ss.getSheetByName('PO_Master');
  if (!shPO) throw new Error('PO_Master sheet not found');

  // Headers are on row 1, spacer row 2, data starts row 3
  const HEADER_ROW = 1;
  const DATA_START_ROW = 3;

  const mlLastRow = shML.getLastRow();
  if (mlLastRow < DATA_START_ROW) return;

  const mlHeaders = shML.getRange(HEADER_ROW, 1, 1, shML.getLastColumn()).getValues()[0].map(h => (h || '').toString().trim());
  const mlIdxProject = mlHeaders.indexOf('Project');
  const mlIdxPO = mlHeaders.indexOf('Customer_PO_Number');
  if (mlIdxProject === -1) throw new Error('Master_Log missing header: Project');
  if (mlIdxPO === -1) throw new Error('Master_Log missing header: Customer_PO_Number');

  const poHeaders = shPO.getRange(HEADER_ROW, 1, 1, shPO.getLastColumn()).getValues()[0].map(h => (h || '').toString().trim());
  const poIdxPO = poHeaders.indexOf('Customer_PO');
  const poIdxProject = poHeaders.indexOf('Project');
  if (poIdxPO === -1) throw new Error('PO_Master missing header: Customer_PO');
  if (poIdxProject === -1) throw new Error('PO_Master missing header: Project');

  // Build PO -> Project map
  const poLastRow = shPO.getLastRow();
  const poMap = new Map();
  if (poLastRow >= 2) {
    const poData = shPO.getRange(2, 1, poLastRow - 1, shPO.getLastColumn()).getValues();
    for (let i = 0; i < poData.length; i++) {
      const po = (poData[i][poIdxPO] || '').toString().trim();
      if (!po) continue;
      const proj = (poData[i][poIdxProject] || '').toString().trim();
      // last write wins if duplicates exist
      poMap.set(po, proj);
    }
  }

  // Read ML data
  const mlNumRows = mlLastRow - DATA_START_ROW + 1;
  const mlNumCols = shML.getLastColumn();
  const mlDataRange = shML.getRange(DATA_START_ROW, 1, mlNumRows, mlNumCols);
  const mlData = mlDataRange.getValues();

  let updated = 0;
  let missingPO = 0;
  let notFound = 0;

  for (let r = 0; r < mlData.length; r++) {
    const poVal = (mlData[r][mlIdxPO] || '').toString().trim();
    if (!poVal) {
      missingPO++;
      continue;
    }

    const currentProject = (mlData[r][mlIdxProject] || '').toString().trim();
    if (currentProject) continue; // only backfill blanks

    const mappedProject = poMap.get(poVal);
    if (!mappedProject) {
      notFound++;
      continue;
    }

    mlData[r][mlIdxProject] = mappedProject;
    updated++;
  }

  if (updated > 0) mlDataRange.setValues(mlData);

  Logger.log(
    JSON.stringify(
      {
        updated_rows: updated,
        rows_missing_po: missingPO,
        rows_po_not_found_in_po_master: notFound,
        po_master_entries: poMap.size
      },
      null,
      2
    )
  );
}

