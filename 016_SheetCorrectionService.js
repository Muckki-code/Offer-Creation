/**
 * @file This file contains functions for actively correcting data and structure
 * in the sheet, typically in response to a user-confirmed dialog.
 */

/**
 * Applies the correct term and quantity to ALL rows within a given bundle.
 */
function applyBundleCorrection(bundleNumber, term, quantity) {
  const sourceFile = "SheetCorrectionService_gs";
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'applyBundleCorrection_start' });
  Log[sourceFile](
    `[${sourceFile} - applyBundleCorrection] Start. Applying Term=${term}, Qty=${quantity} to ALL of bundle #${bundleNumber}.`
  );
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    const bundleRange = _findBundleRange(sheet, bundleNumber);
    if (!bundleRange.startRow || !bundleRange.endRow) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'applyBundleCorrection_noBundleRange' });
      throw new Error(`Could not find the range for bundle #${bundleNumber}.`);
    }

    const termCol = CONFIG.approvalWorkflow.columnIndices.aeTerm;
    const quantityCol = CONFIG.approvalWorkflow.columnIndices.aeQuantity;
    const numRows = bundleRange.endRow - bundleRange.startRow + 1;

    sheet.getRange(bundleRange.startRow, termCol, numRows).setValue(term);
    sheet
      .getRange(bundleRange.startRow, quantityCol, numRows)
      .setValue(quantity);

    // --- THIS IS THE FIX for robustness ---
    SpreadsheetApp.flush();
    // Request a UI refresh because the bundle is now valid and needs its border drawn.
    recalculateAllRows({ refreshUx: true });
    // --- END FIX ---

    SpreadsheetApp.getActive().toast(
      `Bundle #${bundleNumber} has been corrected.`,
      "Success",
      3
    );
    Log[sourceFile](
      `[${sourceFile} - applyBundleCorrection] End. Correction successful.`
    );
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'applyBundleCorrection_end' });
  } catch (e) {
    Log[sourceFile](
      `[${sourceFile} - applyBundleCorrection] ERROR: ${e.message}`
    );
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'applyBundleCorrection_error' });
    SpreadsheetApp.getActive().toast(
      `Failed to apply bundle correction: ${e.message}`,
      "Error",
      5
    );
  }
}


/**
 * Fixes non-consecutive bundle rows by moving them together.
 * @param {string|number} bundleNumber The bundle ID to fix.
 */
function fixBundleGaps(bundleNumber) {
  // This function remains unchanged, it is already correct.
  const sourceFile = "SheetCorrectionService_gs";
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'fixBundleGaps_start' });
  Log[sourceFile](
    `[${sourceFile} - fixBundleGaps] Start. Fixing gaps for bundle #${bundleNumber}.`
  );
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'fixBundleGaps_lockTimeout' });
    SpreadsheetApp.getActive().toast(
      "Sheet is busy, please try again.",
      "Error",
      5
    );
    return;
  }

  try {
    const dataStartRow = CONFIG.approvalWorkflow.startDataRow;
    const lastRow = sheet.getLastRow();
    const bundleNumCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;
    const bundleColumnValues = sheet
      .getRange(dataStartRow, bundleNumCol, lastRow - dataStartRow + 1, 1)
      .getValues();
    const bundleRows = [];
    bundleColumnValues.forEach((val, i) => {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'fixBundleGaps_findRowsLoop' });
      if (String(val[0]).trim() == String(bundleNumber)) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'fixBundleGaps_bundleMatch' });
        bundleRows.push(dataStartRow + i);
      }
    });
    if (bundleRows.length <= 1) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'fixBundleGaps_singleItemBundle' });
      return;
    }
    const targetStartRow = bundleRows[0] + 1;
    const rowsToMove = bundleRows.slice(1).reverse();
    rowsToMove.forEach((sourceRow, i) => {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'fixBundleGaps_moveRowsLoop' });
      const destinationRow = targetStartRow + (bundleRows.length - 2 - i);
      if (sourceRow !== destinationRow) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'fixBundleGaps_moveRow' });
        sheet.moveRows(
          sheet.getRange(sourceRow + ":" + sourceRow),
          destinationRow
        );
      }
    });
    SpreadsheetApp.flush();
    const mockEvent = {
      range: sheet.getRange(bundleRows[0], bundleNumCol),
      value: bundleNumber,
      oldValue: bundleNumber,
    };
    handleSheetAutomations(mockEvent);
    SpreadsheetApp.getActive().toast(
      `Bundle #${bundleNumber} has been re-ordered.`,
      "Success",
      3
    );
  } catch (e) {
    Log[sourceFile](`[${sourceFile} - fixBundleGaps] ERROR: ${e.message}`);
    SpreadsheetApp.getActive().toast(
      `Failed to fix bundle gaps: ${e.message}`,
      "Error",
      5
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * A private helper function to find the start and end row of a bundle
 * by scanning the bundle column, without performing data validation.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {string|number} bundleNumber The bundle ID to find.
 * @returns {{startRow: number|null, endRow: number|null}} An object with the start and end row numbers, or nulls if not found.
 */
function _findBundleRange(sheet, bundleNumber) {
  const sourceFile = "SheetCorrectionService_gs";
  Log.TestCoverage_gs({ file: sourceFile, coverage: '_findBundleRange_start' });
  Log[sourceFile](`[${sourceFile} - _findBundleRange] Start. Searching for bundle #${bundleNumber}.`);

  const dataStartRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: '_findBundleRange_noDataRows' });
    return { startRow: null, endRow: null };
  }

  const bundleNumCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;
  const bundleColumnValues = sheet.getRange(dataStartRow, bundleNumCol, lastRow - dataStartRow + 1, 1).getValues();
  
  let firstRow = null;
  let lastRowFound = null;

  bundleColumnValues.forEach((val, i) => {
    Log.TestCoverage_gs({ file: sourceFile, coverage: '_findBundleRange_loop_iteration' });
    if (String(val[0]).trim() == String(bundleNumber)) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: '_findBundleRange_bundleMatch' });
      const currentRow = dataStartRow + i;
      if (firstRow === null) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: '_findBundleRange_firstRow' });
        firstRow = currentRow;
      }
      lastRowFound = currentRow;
    }
  });

  Log[sourceFile](`[${sourceFile} - _findBundleRange] End. Found range: ${firstRow}-${lastRowFound}.`);
  Log.TestCoverage_gs({ file: sourceFile, coverage: '_findBundleRange_end' });
  return { startRow: firstRow, endRow: lastRowFound };
}

/**
 * --- NEW ---
 * Finds all rows belonging to a given bundle number and clears the bundle
 * number from them, effectively dissolving the bundle.
 */
function dissolveBundle(bundleNumber) {
  const sourceFile = "SheetCorrectionService_gs";
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'dissolveBundle_start' });
  Log[sourceFile](`[${sourceFile} - dissolveBundle] Start. Dissolving bundle #${bundleNumber}.`);
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startRow = CONFIG.approvalWorkflow.startDataRow;
    const lastRow = sheet.getLastRow();
    
    if (lastRow < startRow) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'dissolveBundle_noDataRows' });
      return;
    }

    const bundleNumCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;
    const allBundleColValues = sheet.getRange(startRow, bundleNumCol, lastRow - startRow + 1, 1).getValues();

    const rangesToClear = [];
    allBundleColValues.forEach((val, index) => {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'dissolveBundle_loop_iteration' });
      if (String(val[0]).trim() == String(bundleNumber)) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'dissolveBundle_bundleMatch' });
        const row = startRow + index;
        rangesToClear.push(sheet.getRange(row, bundleNumCol).getA1Notation());
      }
    });

    if (rangesToClear.length > 0) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'dissolveBundle_rangesToClear' });
      const rangeList = sheet.getRangeList(rangesToClear);
      rangeList.clearContent();
      
      const firstRowToClear = sheet.getRange(rangesToClear[0]).getRow();
      const lastRowToClear = sheet.getRange(rangesToClear[rangesToClear.length - 1]).getRow();
      const numRowsToClear = lastRowToClear - firstRowToClear + 1;
      const fullBundleRange = sheet.getRange(firstRowToClear, CONFIG.documentDeviceData.columnIndices.sku, numRowsToClear, CONFIG.maxDataColumn);
      fullBundleRange.setBorder(null, null, null, null, null, null);

      // --- THIS IS THE FIX for robustness ---
      SpreadsheetApp.flush();
      // Request a UI refresh because the bundle has been dissolved and borders must be cleared.
      recalculateAllRows({ refreshUx: true });
      // --- END FIX ---
      
      SpreadsheetApp.getActive().toast(`Bundle #${bundleNumber} has been dissolved.`, "Success", 3);
    } else {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'dissolveBundle_noRangesToClear' });
      SpreadsheetApp.getActive().toast(`Could not find any items for bundle #${bundleNumber}.`, "Info", 3);
    }

    Log[sourceFile](`[${sourceFile} - dissolveBundle] End. Dissolved bundle #${bundleNumber}.`);
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'dissolveBundle_end' });
  } catch (e) {
    Log[sourceFile](`[${sourceFile} - dissolveBundle] ERROR: ${e.message}`);
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'dissolveBundle_error' });
    SpreadsheetApp.getActive().toast(`Failed to dissolve bundle: ${e.message}`, "Error", 5);
  }
}

/**
 * A periodic, self-healing function that brings all bundle borders into alignment
 * with the current data state. It is robust and does not rely on metadata.
 */
function repairAllBundleBorders() {
  const sourceFile = "SheetCorrectionService_gs";
  ExecutionTimer.start('repairAllBundleBorders_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'repairAllBundleBorders_start' });
  Log[sourceFile]("[repairAllBundleBorders] Starting periodic bundle border repair job.");
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataStartRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'repairAllBundleBorders_noData' });
    ExecutionTimer.end('repairAllBundleBorders_total');
    return;
  }

  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - dataBlockStartCol + 1;
  const fullDataRange = sheet.getRange(dataStartRow, dataBlockStartCol, lastRow - dataStartRow + 1, numCols);

  ExecutionTimer.start('repairAllBundleBorders_clearAll');
  fullDataRange.setBorder(null, null, null, null, null, null);
  ExecutionTimer.end('repairAllBundleBorders_clearAll');

  // Find all bundle numbers from the sheet data
  const bundleNumCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;
  const allBundleNumbers = sheet.getRange(dataStartRow, bundleNumCol, lastRow - dataStartRow + 1, 1).getValues()
                                .flat()
                                .map(String)
                                .filter(val => val.trim() !== '');
  const uniqueBundleNumbers = [...new Set(allBundleNumbers)];

  ExecutionTimer.start('repairAllBundleBorders_validateAndApply');
  uniqueBundleNumbers.forEach(bundleNum => {
    // We use the robust validateBundle function as our single source of truth.
    // We pass the first row of the sheet as a placeholder for editedRowNum.
    const validationResult = validateBundle(sheet, dataStartRow, bundleNum);
    
    if (validationResult.startRow && validationResult.endRow && validationResult.endRow > validationResult.startRow) {
       const range = sheet.getRange(validationResult.startRow, dataBlockStartCol, validationResult.endRow - validationResult.startRow + 1, numCols);
       if (validationResult.isValid) {
         _clearAndApplyBundleBorder(range); // Your standard border function
       } else {
         // Apply a distinct border for invalid bundles
         range.setBorder(true, true, true, true, false, false, '#ff0000', SpreadsheetApp.BorderStyle.DASHED);
       }
    }
  });
  ExecutionTimer.end('repairAllBundleBorders_validateAndApply');
  
  Log[sourceFile]("[repairAllBundleBorders] Border repair job complete.");
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'repairAllBundleBorders_end' });
  ExecutionTimer.end('repairAllBundleBorders_total');
}