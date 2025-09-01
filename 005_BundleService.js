/**
 * @file This file contains the service logic for handling device bundles.
 * It includes validation and data grouping functions.
 */


/**
 * Validates a bundle's integrity by checking for consecutive rows and matching term/quantity.
 * REFACTORED: The core loop now finds ALL rows with a given bundle number before
 * checking for consecutiveness, correctly identifying gaps.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {number} editedRowNum The row number that was just edited.
 * @param {string|number} bundleNumber The bundle ID to validate.
 * @returns {{isValid: boolean, errorMessage: string|null, startRow: number|null, endRow: number|null, errorCode?: string, expected?: {term: any, quantity: any}}} An object with the validation result and bundle boundaries.
 */
function validateBundle(sheet, editedRowNum, bundleNumber) {
  const sourceFile = "BundleService_gs";
  ExecutionTimer.start('validateBundle_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'validateBundle_start' });
  Log[sourceFile](`[${sourceFile} - validateBundle] Start: Validating bundle #${bundleNumber} for row ${editedRowNum}.`);

  if (!bundleNumber || String(bundleNumber).trim() === '') {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'validateBundle_noBundleNumber' });
    Log[sourceFile](`[${sourceFile} - validateBundle] No bundle number provided. Validation trivially passes.`);
    ExecutionTimer.end('validateBundle_total');
    return { isValid: true, errorMessage: null, startRow: null, endRow: null };
  }

  const dataStartRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = sheet.getLastRow();

  if (lastRow < dataStartRow) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'validateBundle_noDataRows' });
    ExecutionTimer.end('validateBundle_total');
    return { isValid: true, errorMessage: null, startRow: null, endRow: null };
  }

  const bundleNumCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;

  ExecutionTimer.start('validateBundle_readColumn');
  const bundleColumnValues = sheet.getRange(dataStartRow, bundleNumCol, lastRow - dataStartRow + 1, 1).getValues();
  ExecutionTimer.end('validateBundle_readColumn');
  
  ExecutionTimer.start('validateBundle_findIndexes');
  const bundleRowIndices = []; 

  for (let i = 0; i < bundleColumnValues.length; i++) {
    const sheetValue = bundleColumnValues[i][0];
    const cleanSheetValue = (typeof sheetValue === 'string') ? sheetValue.trim() : sheetValue;
    if (cleanSheetValue != "" && cleanSheetValue == bundleNumber) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'validateBundle_bundleMatch' });
      bundleRowIndices.push(i);
    }
  }
  ExecutionTimer.end('validateBundle_findIndexes');
  Log[sourceFile](`[${sourceFile} - validateBundle] Found ${bundleRowIndices.length} total items for bundle #${bundleNumber}.`);

  if (bundleRowIndices.length <= 1) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'validateBundle_singleItem' });
    const startRowAbs = bundleRowIndices.length > 0 ? dataStartRow + bundleRowIndices[0] : null;
    ExecutionTimer.end('validateBundle_total');
    return { isValid: true, errorMessage: null, startRow: startRowAbs, endRow: startRowAbs };
  }
  
  const bundleStartRow = dataStartRow + bundleRowIndices[0];
  const bundleEndRow = dataStartRow + bundleRowIndices[bundleRowIndices.length - 1];

  ExecutionTimer.start('validateBundle_consecutiveCheck');
  for (let i = 1; i < bundleRowIndices.length; i++) {
    if (bundleRowIndices[i] !== bundleRowIndices[i - 1] + 1) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'validateBundle_fail_nonConsecutive' });
      const errorMsg = `Bundle items must be in consecutive rows. A gap was detected for bundle #${bundleNumber}.`;
      ExecutionTimer.end('validateBundle_consecutiveCheck');
      ExecutionTimer.end('validateBundle_total');
      return { isValid: false, errorMessage: errorMsg, startRow: null, endRow: null, errorCode: 'GAP_DETECTED' };
    }
  }
  ExecutionTimer.end('validateBundle_consecutiveCheck');
  
  ExecutionTimer.start('validateBundle_readAndValidate');
  const termCol = CONFIG.approvalWorkflow.columnIndices.aeTerm;
  const quantityCol = CONFIG.approvalWorkflow.columnIndices.aeQuantity;
  const bundleSize = bundleEndRow - bundleStartRow + 1;
  const termAndQtyValues = sheet.getRange(bundleStartRow, Math.min(termCol, quantityCol), bundleSize, Math.abs(termCol - quantityCol) + 1).getValues();
  const termColIndexInArray = termCol < quantityCol ? 0 : 1;
  const qtyColIndexInArray = termCol < quantityCol ? 1 : 0; // Corrected logic

  const firstTerm = termAndQtyValues[0][termColIndexInArray];
  const firstQuantity = termAndQtyValues[0][qtyColIndexInArray];
  Log[sourceFile](`[${sourceFile} - validateBundle] Checking against base values: Term=${firstTerm}, Quantity=${firstQuantity}.`);

  for (let i = 1; i < termAndQtyValues.length; i++) {
    const currentQuantity = termAndQtyValues[i][qtyColIndexInArray];
    const currentTerm = termAndQtyValues[i][termColIndexInArray];

    if (String(currentTerm) !== String(firstTerm) || String(currentQuantity) !== String(firstQuantity)) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'validateBundle_fail_mismatch' });
      const errorMsg = `All items in bundle #${bundleNumber} must have the same Quantity and Term. Row ${bundleStartRow + i} has mismatched values.`;
      ExecutionTimer.end('validateBundle_readAndValidate');
      ExecutionTimer.end('validateBundle_total');
      // --- NEW: Return the expected values ---
      return { 
          isValid: false, 
          errorMessage: errorMsg, 
          startRow: null, 
          endRow: null, 
          errorCode: 'MISMATCH',
          expected: { term: firstTerm, quantity: firstQuantity }
      };
    }
  }
  ExecutionTimer.end('validateBundle_readAndValidate');

  Log.TestCoverage_gs({ file: sourceFile, coverage: 'validateBundle_pass' });
  Log[sourceFile](`[${sourceFile} - validateBundle] End: Validation successful for bundle #${bundleNumber}.`);
  ExecutionTimer.end('validateBundle_total');
  return { isValid: true, errorMessage: null, startRow: bundleStartRow, endRow: bundleEndRow };
}

function groupApprovedItems(allDataRows, startCol) {
  const sourceFile = "BundleService_gs";
  ExecutionTimer.start('groupApprovedItems_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_start' });
  Log[sourceFile](`[${sourceFile} - groupApprovedItems] START. Processing ${allDataRows.length} total rows with startCol ${startCol}.`);

  const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };

  const approvedStatuses = [CONFIG.approvalWorkflow.statusStrings.approvedOriginal, CONFIG.approvalWorkflow.statusStrings.approvedNew];
  Log[sourceFile](`[${sourceFile} - groupApprovedItems] Will filter for statuses: ${JSON.stringify(approvedStatuses)}`);
  
  const bundlesMap = new Map(); 

  ExecutionTimer.start('groupApprovedItems_mapAllRows');
  allDataRows.forEach((row) => {
    const bundleNumber = String(row[c.bundleNumber - startCol] || '').trim();
    if (bundleNumber) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_rowHasBundleNumber' });
      if (!bundlesMap.has(bundleNumber)) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_newBundleInMap' });
        bundlesMap.set(bundleNumber, { totalItems: 0 });
      }
      bundlesMap.get(bundleNumber).totalItems++;
    } else {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_rowHasNoBundleNumber' });
    }
  });
  ExecutionTimer.end('groupApprovedItems_mapAllRows');
  Log[sourceFile](`[${sourceFile} - groupApprovedItems] Mapped bundle counts. Found ${bundlesMap.size} unique bundle numbers.`);

  const approvedRows = allDataRows.filter(row => approvedStatuses.includes(row[c.status - startCol]));
  Log[sourceFile](`[${sourceFile} - groupApprovedItems] Filtered to ${approvedRows.length} approved rows.`);

  const processedBundleNumbers = new Set();
  const result = [];

  ExecutionTimer.start('groupApprovedItems_mainLoop');
  approvedRows.forEach((row) => {
    const bundleNumber = String(row[c.bundleNumber - startCol] || '').trim();
    
    if (!bundleNumber) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_isIndividual' });
      result.push({ isBundle: false, row: row });
    } else {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_isBundleItem' });
      if (processedBundleNumbers.has(bundleNumber)) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_bundleAlreadyProcessed' });
        return;
      }
      
      processedBundleNumbers.add(bundleNumber);
      const bundleInfo = bundlesMap.get(bundleNumber);
      const approvedItemsInBundle = approvedRows.filter(r => String(r[c.bundleNumber - startCol] || '').trim() == bundleNumber);
      Log[sourceFile](`[${sourceFile} - groupApprovedItems] Bundle #${bundleNumber} Check: Total items expected: ${bundleInfo ? bundleInfo.totalItems : 'N/A'}. Approved items found: ${approvedItemsInBundle.length}.`);

      if (!bundleInfo || approvedItemsInBundle.length !== bundleInfo.totalItems) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_bundleIncomplete' });
        Log[sourceFile](`[${sourceFile} - groupApprovedItems] Decision: Bundle #${bundleNumber} is INCOMPLETE. Skipping.`);
        return;
      }
      
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_bundleComplete' });
      Log[sourceFile](`[${sourceFile} - groupApprovedItems] Decision: Bundle #${bundleNumber} is COMPLETE. Consolidating and adding to results.`);
      
      let totalNetMonthlyPrice = 0;
      let modelsWithPrices = [];

      approvedItemsInBundle.forEach(bundleItem => {
        const price = getNumericValue(bundleItem[c.financeApprovedPrice - startCol]);
        totalNetMonthlyPrice += price;
        modelsWithPrices.push({ name: bundleItem[c.model - startCol], price: price });
      });
      
      modelsWithPrices.sort((a, b) => b.price - a.price);
      const sortedModelNames = modelsWithPrices.map(m => m.name).join(',\n');

      result.push({
        isBundle: true,
        models: sortedModelNames,
        quantity: approvedItemsInBundle[0][c.aeQuantity - startCol],
        term: approvedItemsInBundle[0][c.aeTerm - startCol],
        totalNetMonthlyPrice: totalNetMonthlyPrice
      });
    }
  });
  ExecutionTimer.end('groupApprovedItems_mainLoop');

  Log.TestCoverage_gs({ file: sourceFile, coverage: 'groupApprovedItems_end' });
  Log[sourceFile](`[${sourceFile} - groupApprovedItems] END. Processed into ${result.length} renderable items.`);
  ExecutionTimer.end('groupApprovedItems_total');
  return result;
}


/**
 * --- NEW ---
 * A lightweight, publicly callable function for the sidebar to check if a previously
 * detected bundle error has been manually corrected by the user in the sheet.
 * @param {string|number} bundleNumber The bundle ID to re-validate.
 * @returns {boolean} True if the bundle is still invalid, false if it is now valid.
 */
function isBundleStillInvalid(bundleNumber) {
    const sourceFile = "BundleService_gs";
    Log[sourceFile](`[${sourceFile} - isBundleStillInvalid] Start: Re-validating bundle #${bundleNumber} for sidebar.`);
    
    // We can re-use the powerful, existing validateBundle function.
    // We don't need the editedRowNum for this check, so we can pass a placeholder like 0.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const validationResult = validateBundle(sheet, 0, bundleNumber);

    // The logic is simple: if the validation result is NOT valid, the bundle is still broken.
    if (!validationResult.isValid) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'isBundleStillInvalid_isInvalid' });
        Log[sourceFile](`[${sourceFile} - isBundleStillInvalid] Result for bundle #${bundleNumber}: Still Invalid.`);
        return true;
    }

    Log.TestCoverage_gs({ file: sourceFile, coverage: 'isBundleStillInvalid_isValid' });
    Log[sourceFile](`[${sourceFile} - isBundleStillInvalid] Result for bundle #${bundleNumber}: Now Valid.`);
    return false;
}

/**
 * Scans the entire sheet to find all bundle-related errors, including
 * mismatched data and non-consecutive rows.
 * REFACTORED FOR PERFORMANCE: This version reads the entire data block once
 * and performs all validation in memory to minimize sheet interactions.
 *
 * @returns {Array<Object>} An array of error objects. Each object contains
 *   the bundleNumber and details about the specific error.
 */
function findAllBundleErrors() {
  const sourceFile = "BundleService_gs";
  ExecutionTimer.start('findAllBundleErrors_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'findAllBundleErrors_start' });
  Log[sourceFile](`[${sourceFile} - findAllBundleErrors] Start: Beginning full-sheet bundle health check (Optimized).`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataStartRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = sheet.getLastRow();
  const allErrors = [];

  if (lastRow < dataStartRow) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'findAllBundleErrors_noData' });
    ExecutionTimer.end('findAllBundleErrors_total');
    return allErrors;
  }

  // 1. Single Bulk Read of the entire data block
  ExecutionTimer.start('findAllBundleErrors_readSheet');
  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - dataBlockStartCol + 1;
  const allData = sheet.getRange(dataStartRow, dataBlockStartCol, lastRow - dataStartRow + 1, numCols).getValues();
  ExecutionTimer.end('findAllBundleErrors_readSheet');

  // Pre-calculate 0-based array indices from 1-based config indices
  const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };
  const bundleNumColIndex = c.bundleNumber - dataBlockStartCol;
  const termColIndex = c.aeTerm - dataBlockStartCol;
  const quantityColIndex = c.aeQuantity - dataBlockStartCol;

  // 2. Group rows by bundle number in memory
  ExecutionTimer.start('findAllBundleErrors_groupInMemory');
  const bundlesMap = new Map();
  for (let i = 0; i < allData.length; i++) {
    const rowData = allData[i];
    const bundleNum = String(rowData[bundleNumColIndex] || '').trim();
    if (bundleNum) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'findAllBundleErrors_rowHasBundleNum' });
      if (!bundlesMap.has(bundleNum)) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'findAllBundleErrors_newBundleInMap' });
        bundlesMap.set(bundleNum, []);
      }
      bundlesMap.get(bundleNum).push({
        rowData: rowData,
        rowIndex: dataStartRow + i // Store original sheet row index
      });
    } else {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'findAllBundleErrors_rowHasNoBundleNum' });
    }
  }
  ExecutionTimer.end('findAllBundleErrors_groupInMemory');
  Log[sourceFile](`[${sourceFile} - findAllBundleErrors] Grouped ${bundlesMap.size} unique bundles in memory.`);

  // 3. Validate each bundle group in memory
  ExecutionTimer.start('findAllBundleErrors_validateInMemory');
  for (const [bundleNum, rows] of bundlesMap.entries()) {
    if (rows.length <= 1) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'findAllBundleErrors_singleItemBundle' });
      continue;
    }

    // A. Check for non-consecutive rows (gaps)
    rows.sort((a, b) => a.rowIndex - b.rowIndex); // Ensure sorted by original row index
    let hasGap = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i].rowIndex !== rows[i - 1].rowIndex + 1) {
        allErrors.push({
          bundleNumber: bundleNum,
          errorCode: 'GAP_DETECTED',
          errorMessage: `Bundle items must be in consecutive rows. A gap was detected for bundle #${bundleNum}.`,
          expected: null
        });
        hasGap = true;
        break; // Found a gap, no need to check for mismatch
      }
    }
    if (hasGap) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'findAllBundleErrors_gapFound' });
      continue;
    }

    // B. Check for mismatched Term or Quantity
    const expectedTerm = rows[0].rowData[termColIndex];
    const expectedQuantity = rows[0].rowData[quantityColIndex];

    for (let i = 1; i < rows.length; i++) {
      const currentTerm = rows[i].rowData[termColIndex];
      const currentQuantity = rows[i].rowData[quantityColIndex];
      if (String(currentTerm) !== String(expectedTerm) || String(currentQuantity) !== String(expectedQuantity)) {
        allErrors.push({
          bundleNumber: bundleNum,
          errorCode: 'MISMATCH',
          errorMessage: `All items in bundle #${bundleNum} must have the same Quantity and Term.`,
          expected: { term: expectedTerm, quantity: expectedQuantity }
        });
        break; // Found a mismatch, move to the next bundle
      }
    }
  }
  ExecutionTimer.end('findAllBundleErrors_validateInMemory');

  Log[sourceFile](`[${sourceFile} - findAllBundleErrors] End: Found a total of ${allErrors.length} bundle errors after in-memory scan.`);
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'findAllBundleErrors_end' });
  ExecutionTimer.end('findAllBundleErrors_total');
  
  return allErrors;
}

