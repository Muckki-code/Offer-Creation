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
  Log.TestCoverage_gs({ file: 'BundleService.gs', coverage: 'groupApprovedItems_start' });
  Log[sourceFile](`[${sourceFile} - groupApprovedItems] START. Processing ${allDataRows.length} total rows with startCol ${startCol}.`);

  const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };

  const approvedStatuses = [CONFIG.approvalWorkflow.statusStrings.approvedOriginal, CONFIG.approvalWorkflow.statusStrings.approvedNew];
  Log[sourceFile](`[${sourceFile} - groupApprovedItems] Will filter for statuses: ${JSON.stringify(approvedStatuses)}`);
  
  const bundlesMap = new Map(); 

  ExecutionTimer.start('groupApprovedItems_mapAllRows');
  allDataRows.forEach((row) => {
    const bundleNumber = String(row[c.bundleNumber - startCol] || '').trim();
    if (bundleNumber) {
      if (!bundlesMap.has(bundleNumber)) {
        bundlesMap.set(bundleNumber, { totalItems: 0 });
      }
      bundlesMap.get(bundleNumber).totalItems++;
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
      Log.TestCoverage_gs({ file: 'BundleService.gs', coverage: 'groupApprovedItems_isIndividual' });
      result.push({ isBundle: false, row: row });
    } else {
      Log.TestCoverage_gs({ file: 'BundleService.gs', coverage: 'groupApprovedItems_isBundleItem' });
      if (processedBundleNumbers.has(bundleNumber)) {
        Log.TestCoverage_gs({ file: 'BundleService.gs', coverage: 'groupApprovedItems_bundleAlreadyProcessed' });
        return;
      }
      
      processedBundleNumbers.add(bundleNumber);
      const bundleInfo = bundlesMap.get(bundleNumber);
      const approvedItemsInBundle = approvedRows.filter(r => String(r[c.bundleNumber - startCol] || '').trim() == bundleNumber);
      Log[sourceFile](`[${sourceFile} - groupApprovedItems] Bundle #${bundleNumber} Check: Total items expected: ${bundleInfo ? bundleInfo.totalItems : 'N/A'}. Approved items found: ${approvedItemsInBundle.length}.`);

      if (!bundleInfo || approvedItemsInBundle.length !== bundleInfo.totalItems) {
        Log.TestCoverage_gs({ file: 'BundleService.gs', coverage: 'groupApprovedItems_bundleIncomplete' });
        Log[sourceFile](`[${sourceFile} - groupApprovedItems] Decision: Bundle #${bundleNumber} is INCOMPLETE. Skipping.`);
        return;
      }
      
      Log.TestCoverage_gs({ file: 'BundleService.gs', coverage: 'groupApprovedItems_bundleComplete' });
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

  Log.TestCoverage_gs({ file: 'BundleService.gs', coverage: 'groupApprovedItems_end' });
  Log[sourceFile](`[${sourceFile} - groupApprovedItems] END. Processed into ${result.length} renderable items.`);
  ExecutionTimer.end('groupApprovedItems_total');
  return result;
}