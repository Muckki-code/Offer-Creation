/**
* @file This file contains core sheet automation functions,
* including the main onEdit handler and general row calculations.
*/

// --- SPRINT 2 PERFORMANCE REFACTOR: EXECUTION-SCOPED CACHE ---
let _staticValuesCache = null;

// The global getColumnIndexByLetter function from Config.gs is used.

/**
* A robust and performant function to convert a potential currency string or number into a clean numeric value.
*/
function getNumericValue(value) {
  Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'getNumericValue_start' });
  if (typeof value === 'number' && !isNaN(value)) {
    return value;
  }
  if (typeof value !== 'string' || value.trim() === '') {
    return 0;
  }
  let numberString = value.replace(/[€$]/g, '').trim();
  numberString = numberString.replace(/,/g, ''); // Remove all commas.
  const validNumericRegex = /^-?\d*\.?\d*$/;
  if (!validNumericRegex.test(numberString)) {
    return 0;
  }
  const result = parseFloat(numberString);
  return isNaN(result) ? 0 : result;
}

// In SheetCoreAutomations.gs

/**
* SPRINT 2 PERFORMANCE REFACTOR: Caching helper function.
* OPTIMIZED: This version reads all required header/config cells in a single batch
* operation (.getValues()) to significantly reduce API call overhead.
*/
function _getStaticSheetValues(sheet) {
  Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: '_getStaticSheetValues_start' });
  const sourceFile = "SheetCoreAutomations_gs";
  if (_staticValuesCache) {
    Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: '_getStaticSheetValues_fromCache' });
    Log[sourceFile](`[_getStaticSheetValues] Returning values from cache.`);
    return _staticValuesCache;
  }
  Log[sourceFile](`[SheetCoreAutomations_gs - _getStaticSheetValues] Cache empty. Reading static values from sheet.`);
  ExecutionTimer.start('_getStaticSheetValues_read');

  // Define a single range that encompasses all the static cells we need.
  // This reads from I1 to L4.
  const staticCellsRange = sheet.getRange("I1:L4");
  const staticCellValues = staticCellsRange.getValues();

  ExecutionTimer.end('_getStaticSheetValues_read');
  ExecutionTimer.start('_getStaticSheetValues_parse');

  // Extract values from the 2D array based on their relative positions.
  // getRange("I1:L4") means:
  // I1 is at [0][0], J1 is [0][1], K1 is [0][2], L1 is [0][3]
  // I2 is at [1][0], J2 is [1][1], K2 is [1][2], L2 is [1][3]
  // etc.

  const languageValue = staticCellValues[0][0]; // I1
  const telekomDealValue = staticCellValues[0][3]; // L1

  const staticValues = {
    isTelekomDeal: String(telekomDealValue || "").toLowerCase() === 'yes',
    docLanguage: String(languageValue || 'german').trim().toLowerCase()
  };

  ExecutionTimer.end('_getStaticSheetValues_parse');
  Log[sourceFile](`[SheetCoreAutomations_gs - _getStaticSheetValues] Caching and returning: ${JSON.stringify(staticValues)}`);
  _staticValuesCache = staticValues;
  return _staticValuesCache;
}

/**
* OPTIMIZED: Finds the maximum existing index and returns the next available index.
* This version performs a single, efficient read of only the necessary data.
*/
function getNextAvailableIndex(sheet) {
  ExecutionTimer.start('getNextAvailableIndex_total');
  Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'getNextAvailableIndex_start' });
  const indexColIndex = CONFIG.documentDeviceData.columnIndices.index;
  const startRow = CONFIG.approvalWorkflow.startDataRow;
  let maxIndex = 0;

  ExecutionTimer.start('getNextAvailableIndex_getValues');
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'getNextAvailableIndex_hasDataRows' });
    // Read the entire index column from the data start row to the end in one operation.
    const indexValues = sheet.getRange(startRow, indexColIndex, lastRow - startRow + 1, 1).getValues();
    ExecutionTimer.end('getNextAvailableIndex_getValues');

    ExecutionTimer.start('getNextAvailableIndex_loop');
    // Find the max index from the in-memory array.
    maxIndex = indexValues.reduce((max, row) => {
      const value = parseFloat(row[0]);
      return !isNaN(value) && value > max ? value : max;
    }, 0);
    ExecutionTimer.end('getNextAvailableIndex_loop');
  } else {
    ExecutionTimer.end('getNextAvailableIndex_getValues');
  }

  Log.SheetCoreAutomations_gs(`[SheetCoreAutomations_gs - getNextAvailableIndex] Found max index ${maxIndex}. Next available will be ${maxIndex + 1}.`);
  ExecutionTimer.end('getNextAvailableIndex_total');
  return maxIndex + 1;
}

/**
* OPTIMIZED: Recalculates all data rows in the active sheet.
* This version determines the next available index once from the in-memory array
* to avoid repeated, slow calls back to the sheet.
*/
function recalculateAllRows(options = {}) {
  const sourceFile = "SheetCoreAutomations_gs";
  ExecutionTimer.start('recalculateAllRows_total');
  Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'recalculateAllRows_start' });
  _staticValuesCache = null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = getLastLastRow(sheet);
  if (lastRow < startRow) {
    Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'recalculateAllRows_noData' });
    Log[sourceFile](`[${sourceFile} - recalculateAllRows] No data rows found (lastRow ${lastRow} < startRow ${startRow}). Exiting.`);
    ExecutionTimer.end('recalculateAllRows_total');
    return;
  }
  const numRows = lastRow - startRow + 1;
  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - dataBlockStartCol + 1;
  ExecutionTimer.start('recalculateAllRows_readSheet');
  const allValuesBefore = sheet.getRange(startRow, dataBlockStartCol, numRows, numCols).getValues();
  const allValuesAfter = JSON.parse(JSON.stringify(allValuesBefore));
  ExecutionTimer.end('recalculateAllRows_readSheet');
  Log[sourceFile](`[${sourceFile} - recalculateAllRows] Read ${numRows} rows from sheet.`);

  const staticValues = _getStaticSheetValues(sheet);
  const combinedIndexes = { ...CONFIG.approvalWorkflow.columnIndices, ...CONFIG.documentDeviceData.columnIndices };
  const statusStrings = CONFIG.approvalWorkflow.statusStrings;
  let nextIndex = null; // Initialize to null

  ExecutionTimer.start('recalculateAllRows_mainLoop');
  for (let i = 0; i < numRows; i++) {
    Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'recalculateAllRows_loop_iteration' });
    const currentRowNum = startRow + i;
    const inMemoryRowValues = allValuesAfter[i];
    const originalRowValuesForThisRow = allValuesBefore[i];
    Log[sourceFile](`[${sourceFile} - recalculateAllRows] Processing row ${currentRowNum}.`);

    const modelName = inMemoryRowValues[combinedIndexes.model - dataBlockStartCol];
    if (modelName && !inMemoryRowValues[combinedIndexes.index - dataBlockStartCol]) {
      Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'recalculateAllRows_assignIndex' });
      if (nextIndex === null) {
        const allCurrentIndices = allValuesAfter.map(row => parseFloat(row[combinedIndexes.index - dataBlockStartCol])).filter(val => !isNaN(val));
        const maxCurrentIndex = allCurrentIndices.length > 0 ? Math.max(...allCurrentIndices) : 0;
        nextIndex = maxCurrentIndex + 1;
      }
      inMemoryRowValues[combinedIndexes.index - dataBlockStartCol] = nextIndex++;
      Log[sourceFile](`[${sourceFile} - recalculateAllRows] Row ${currentRowNum}: Assigned new index ${inMemoryRowValues[combinedIndexes.index - dataBlockStartCol]}.`);
    }

    if (modelName && !inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol]) {
      Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'recalculateAllRows_assignApproverAction' });
      inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol] = "Choose Action";
      Log[sourceFile](`[${sourceFile} - recalculateAllRows] Row ${currentRowNum}: Assigned default Approver Action.`);
    }

    updateCalculationsForRow(sheet, currentRowNum, inMemoryRowValues, staticValues.isTelekomDeal, combinedIndexes, CONFIG.approvalWorkflow, dataBlockStartCol);

    // --- THIS IS THE FIX ---
    const statusUpdateOptions = { forceRevisionOfFinalizedItems: true };
    const initialStatus = originalRowValuesForThisRow[combinedIndexes.status - dataBlockStartCol] || "";
    
    // Corrected function call with the right signature
    const newStatus = updateStatusForRow(
      inMemoryRowValues,
      originalRowValuesForThisRow,
      staticValues.isTelekomDeal,
      statusUpdateOptions,
      dataBlockStartCol,
      combinedIndexes
    );

    // Logic to handle the returned status, copied from handleSheetAutomations
    if (newStatus !== initialStatus) {
        if (newStatus === null) { 
            inMemoryRowValues[combinedIndexes.status - dataBlockStartCol] = ""; 
        } else {
            inMemoryRowValues[combinedIndexes.status - dataBlockStartCol] = newStatus;
            if ([statusStrings.pending, statusStrings.draft, statusStrings.revisedByAE].includes(newStatus)) {
                inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol] = "Choose Action";
            }
            const approvedStatuses = [statusStrings.approvedOriginal, statusStrings.approvedNew, statusStrings.rejected];
            if (approvedStatuses.includes(initialStatus) && !approvedStatuses.includes(newStatus)) {
                inMemoryRowValues[combinedIndexes.financeApprovedPrice - dataBlockStartCol] = "";
                inMemoryRowValues[combinedIndexes.approvedBy - dataBlockStartCol] = "";
                inMemoryRowValues[combinedIndexes.approvalDate - dataBlockStartCol] = "";
            }
        }
    }
  }
  ExecutionTimer.end('recalculateAllRows_mainLoop');

  ExecutionTimer.start('recalculateAllRows_writeSheet');
  sheet.getRange(startRow, dataBlockStartCol, numRows, numCols).setValues(allValuesAfter);
  ExecutionTimer.end('recalculateAllRows_writeSheet');
  Log[sourceFile](`[${sourceFile} - recalculateAllRows] Wrote all recalculated data back to the sheet.`);
  ExecutionTimer.end('recalculateAllRows_total');
}



/**
* Main onEdit trigger handler.
* REVISED: This version is robust against inconsistent onEdit event objects,
* uses e.oldValue for single edits, accepts an injected array for testing,
* and contains comprehensive logging. The BQ data wipe logic now correctly
* handles both single edits and partial pastes by inspecting the edit range.
* Includes SpreadsheetApp.flush() to prevent race conditions from rapid edits.
* Bundle integrity validation and formatting is now handled efficiently
* using ROW-LEVEL Developer Metadata for immediate feedback.
* Business logic simplified to no longer wipe AE data on model changes.
*/
function handleSheetAutomations(e, trueOriginalValuesForTest = null) {
  const sourceFile = "SheetCoreAutomations_gs";
  ExecutionTimer.start('handleSheetAutomations_total');
  Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_start' });
  _staticValuesCache = null;
  const range = e.range;

  if (range.getRow() < CONFIG.approvalWorkflow.startDataRow && range.getA1Notation() !== CONFIG.offerDetailsCells.telekomDeal) {
    Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_exit_notDataRow' });
    ExecutionTimer.end('handleSheetAutomations_total');
    return;
  }
  if (range.getA1Notation() === CONFIG.offerDetailsCells.telekomDeal) {
    Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_recalcTriggered' });
    Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Detected edit in global 'Telekom Deal' cell. Triggering recalculateAllRows.`);
    recalculateAllRows({ oldValueIsTelekomDeal: (e.oldValue || 'no').toLowerCase() === 'yes' });
    ExecutionTimer.end('handleSheetAutomations_total');
    return;
  }

  ExecutionTimer.start('handleSheetAutomations_lock');
  const lock = LockService.getScriptLock();
  try {
    Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_lock_wait' });
    lock.waitLock(30000);
  } catch (err) {
    Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_lock_fail' });
    Log[sourceFile](`[${sourceFile} - handleSheetAutomations] WARNING: Could not obtain lock. Error: ${err.message}.`);
    SpreadsheetApp.getActive().toast("The sheet is busy, please try your edit again in a moment.", "Busy", 3);
    ExecutionTimer.end('handleSheetAutomations_lock');
    ExecutionTimer.end('handleSheetAutomations_total');
    return;
  }
  ExecutionTimer.end('handleSheetAutomations_lock');

  try {
    ExecutionTimer.start('handleSheetAutomations_flush');
    SpreadsheetApp.flush();
    Log[sourceFile]("[handleSheetAutomations] Lock acquired. Flushed all pending spreadsheet changes to prevent race conditions.");
    ExecutionTimer.end('handleSheetAutomations_flush');
    
    const sheet = range.getSheet();
    const combinedIndexes = { ...CONFIG.approvalWorkflow.columnIndices, ...CONFIG.documentDeviceData.columnIndices };
    const editedRowStart = range.getRow();
    const numEditedRows = range.getNumRows();
    const editedColStart = range.getColumn();
    const isSingleCellEdit = (numEditedRows === 1 && range.getNumColumns() === 1);
    const editedCol = editedColStart;
    const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
    const numColsInDataBlock = CONFIG.maxDataColumn - dataBlockStartCol + 1;
    
    const finalValuesToWrite = trueOriginalValuesForTest ? JSON.parse(JSON.stringify(trueOriginalValuesForTest)) : sheet.getRange(editedRowStart, dataBlockStartCol, numEditedRows, numColsInDataBlock).getValues();
    const originalValuesForComparison = JSON.parse(JSON.stringify(finalValuesToWrite));
    
    Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Data source: ${trueOriginalValuesForTest ? 'Injected by Test' : 'Read from Sheet'}.`);
    if (isSingleCellEdit) {
      Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_reconstruct_before_state' });
      const editedColIndexInArray = editedCol - dataBlockStartCol;
      if (editedColIndexInArray >= 0 && editedColIndexInArray < originalValuesForComparison[0].length) {
          originalValuesForComparison[0][editedColIndexInArray] = e.oldValue;
          Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Reconstructed 'before' state for single edit in col ${editedCol} with oldValue: '${e.oldValue}'`);
      }
    }
    const staticValues = _getStaticSheetValues(sheet);
    let nextIndex = null;

    ExecutionTimer.start('handleSheetAutomations_mainLoop');
    for (let i = 0; i < numEditedRows; i++) {
      Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_loop_start' });
      const currentRowNumInSheet = editedRowStart + i;
      const inMemoryRowValues = finalValuesToWrite[i]; 
      const originalRowValues = originalValuesForComparison[i];

      Log[sourceFile](`[${sourceFile} - handleSheetAutomations] ---- STARTING ROW ${currentRowNumInSheet} ----`);
      
      let wipeBqData = false;
      if (isSingleCellEdit) {
        const skuChanged = String(inMemoryRowValues[combinedIndexes.sku - dataBlockStartCol] || "") !== String(originalRowValues[combinedIndexes.sku - dataBlockStartCol] || "");
        const modelChanged = String(inMemoryRowValues[combinedIndexes.model - dataBlockStartCol] || "") !== String(originalRowValues[combinedIndexes.model - dataBlockStartCol] || "");
        if (skuChanged !== modelChanged) {
          wipeBqData = true;
          Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Flagging BQ data for wipe because SKU/Model changed independently (SKU changed: ${skuChanged}, Model changed: ${modelChanged}).`);
        }
      } else {
        const pasteStartCol = range.getColumn();
        const pasteEndCol = pasteStartCol + range.getNumColumns() - 1;
        const skuCol = combinedIndexes.sku;
        const modelCol = combinedIndexes.model;
        const skuWasPasted = (skuCol >= pasteStartCol && skuCol <= pasteEndCol);
        const modelWasPasted = (modelCol >= pasteStartCol && modelCol <= pasteEndCol);
        if (skuWasPasted !== modelWasPasted) {
          wipeBqData = true;
          Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Flagging BQ data for wipe due to desynchronizing paste (SKU pasted: ${skuWasPasted}, Model pasted: ${modelWasPasted}).`);
        }
      }
      if (wipeBqData) {
        Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_wipeBqData' });
        inMemoryRowValues[combinedIndexes.epCapexRaw - dataBlockStartCol] = "";
        inMemoryRowValues[combinedIndexes.tkCapexRaw - dataBlockStartCol] = "";
        inMemoryRowValues[combinedIndexes.rentalTargetRaw - dataBlockStartCol] = "";
        inMemoryRowValues[combinedIndexes.rentalLimitRaw - dataBlockStartCol] = "";
      }
      
      ExecutionTimer.start('handleSheetAutomations_sanitizeBlock');
      if (!isSingleCellEdit) { 
        Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_pasteSanitization' });
        const fieldsToWipe = [
          'index', 'lrfPreview', 'contractValuePreview', 'status', 'financeApprovedPrice', 
          'approvedBy', 'approvalDate', 'approverComments', 'approverPriceProposal'
        ];
        fieldsToWipe.forEach(key => { 
          if(combinedIndexes[key]) {
            inMemoryRowValues[combinedIndexes[key] - dataBlockStartCol] = ""; 
          }
        });
        inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol] = "Choose Action";
        originalRowValues[combinedIndexes.status - dataBlockStartCol] = "";
      } else { 
        if (CONFIG.protectedColumnIndices.includes(editedCol)) {
          Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_protectedColumnRevert' });
             inMemoryRowValues[editedCol - dataBlockStartCol] = e.oldValue;
        }
      }
      ExecutionTimer.end('handleSheetAutomations_sanitizeBlock');

      ExecutionTimer.start('handleSheetAutomations_downstreamLogic');
      updateCalculationsForRow(sheet, currentRowNumInSheet, inMemoryRowValues, staticValues.isTelekomDeal, combinedIndexes, CONFIG.approvalWorkflow, dataBlockStartCol);
      const approverActionValue = inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol];
      const isApprovalAction = isSingleCellEdit && editedCol === combinedIndexes.approverAction && approverActionValue && approverActionValue !== "Choose Action";
      
      if (isApprovalAction) {
        const mockEventForApproval = { ...e, value: approverActionValue, oldValue: originalRowValues[editedCol - dataBlockStartCol]};
        processSingleApprovalAction(sheet, currentRowNumInSheet, mockEventForApproval, inMemoryRowValues, combinedIndexes, originalRowValues, dataBlockStartCol);
      } else {
        const initialStatus = originalRowValues[combinedIndexes.status - dataBlockStartCol] || "";
        const newStatus = updateStatusForRow(inMemoryRowValues, originalRowValues, staticValues.isTelekomDeal, {}, dataBlockStartCol, combinedIndexes);
        if (newStatus !== initialStatus) {
            logTableActivity({ mainSheet: sheet, rowNum: currentRowNumInSheet, oldStatus: initialStatus, newStatus: newStatus, currentFullRowValues: inMemoryRowValues, originalFullRowValues: originalRowValues, startCol: dataBlockStartCol });
            if (newStatus === null) { 
                inMemoryRowValues[combinedIndexes.status - dataBlockStartCol] = ""; 
            } else {
                inMemoryRowValues[combinedIndexes.status - dataBlockStartCol] = newStatus;
                if ([CONFIG.approvalWorkflow.statusStrings.pending, CONFIG.approvalWorkflow.statusStrings.draft, CONFIG.approvalWorkflow.statusStrings.revisedByAE].includes(newStatus)) {
                    inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol] = "Choose Action";
                }
                const approvedStatuses = [CONFIG.approvalWorkflow.statusStrings.approvedOriginal, CONFIG.approvalWorkflow.statusStrings.approvedNew];
                if (approvedStatuses.includes(initialStatus) && !approvedStatuses.includes(newStatus)) {
                    inMemoryRowValues[combinedIndexes.financeApprovedPrice - dataBlockStartCol] = "";
                    inMemoryRowValues[combinedIndexes.approvedBy - dataBlockStartCol] = "";
                    inMemoryRowValues[combinedIndexes.approvalDate - dataBlockStartCol] = "";
                }
            }
        }
      }
      ExecutionTimer.end('handleSheetAutomations_downstreamLogic');
      
      ExecutionTimer.start('handleSheetAutomations_initializationChecks');
      const modelName = inMemoryRowValues[combinedIndexes.model - dataBlockStartCol];
      if (modelName && !inMemoryRowValues[combinedIndexes.index - dataBlockStartCol]) {
        if (nextIndex === null) { nextIndex = getNextAvailableIndex(sheet); }
        inMemoryRowValues[combinedIndexes.index - dataBlockStartCol] = nextIndex++;
      }
      if (modelName && !inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol]) {
        inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol] = "Choose Action";
      }
      ExecutionTimer.end('handleSheetAutomations_initializationChecks');
      
      ExecutionTimer.start('handleSheetAutomations_surgicalWrite');
      finalValuesToWrite[i].forEach((finalCellValue, colIndexInArray) => {
        const colIndexInSheet = dataBlockStartCol + colIndexInArray;
        const currentSheetValue = sheet.getRange(currentRowNumInSheet, colIndexInSheet).getValue();
        if(String(currentSheetValue) !== String(finalCellValue)) {
            sheet.getRange(currentRowNumInSheet, colIndexInSheet).setValue(finalCellValue);
        }
      });
      ExecutionTimer.end('handleSheetAutomations_surgicalWrite');
    }
    ExecutionTimer.end('handleSheetAutomations_mainLoop');
    
    // --- METADATA-DRIVEN BUNDLE VALIDATION & FORMATTING ---
    if (CONFIG.featureFlags.highlightBundlesWithBorders && isSingleCellEdit) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'handleSheetAutomations_bundleFlag_enabled' });
      ExecutionTimer.start('handleSheetAutomations_bundleLogic');
      
      const integrityCols = [combinedIndexes.bundleNumber, combinedIndexes.aeQuantity, combinedIndexes.aeTerm];
      if (integrityCols.includes(editedCol)) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'handleSheetAutomations_bundleIntegrityEdit' });
        
        const oldBundleInfo = _getBundleInfoFromRange(range);
        const currentBundleNum = sheet.getRange(editedRowStart, combinedIndexes.bundleNumber).getValue();

        if (oldBundleInfo && (!currentBundleNum || String(currentBundleNum) !== String(oldBundleInfo.bundleId))) {
          Log.TestCoverage_gs({ file: sourceFile, coverage: 'handleSheetAutomations_bundleBroken' });
          _clearMetadataFromRowRange(sheet, oldBundleInfo.startRow, oldBundleInfo.endRow);
          const oldRange = sheet.getRange(oldBundleInfo.startRow, dataBlockStartCol, oldBundleInfo.endRow - oldBundleInfo.startRow + 1, numColsInDataBlock);
          oldRange.setBorder(null, null, null, null, null, null);
        }

        if (currentBundleNum) {
            Log.TestCoverage_gs({ file: sourceFile, coverage: 'handleSheetAutomations_bundleValidationStart' });
            const validationResult = validateBundle(sheet, editedRowStart, currentBundleNum);
            
             if (validationResult.isValid) {
              Log.TestCoverage_gs({ file: sourceFile, coverage: 'handleSheetAutomations_bundleValid' });
              if (validationResult.startRow && validationResult.endRow && validationResult.startRow !== validationResult.endRow) {
                const newBundleInfo = { bundleId: String(currentBundleNum), startRow: validationResult.startRow, endRow: validationResult.endRow };
                if (JSON.stringify(oldBundleInfo) !== JSON.stringify(newBundleInfo)) {
                  _setMetadataForRowRange(sheet, newBundleInfo);
                  const bundleRangeToFormat = sheet.getRange(newBundleInfo.startRow, dataBlockStartCol, newBundleInfo.endRow - newBundleInfo.startRow + 1, numColsInDataBlock);
                  _clearAndApplyBundleBorder(bundleRangeToFormat);
                }
              }
            } else { 
              // --- FINAL, CORRECTED LOGIC FOR HANDLING INVALID BUNDLES ---
              ExecutionTimer.start('handleSheetAutomations_bundleInvalidDialog');
              Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Bundle validation FAILED. Error Code: '${validationResult.errorCode}'. Showing corrective dialog.`);
              
              // We DO NOT revert the edit here. The user will decide in the dialog.
              
              switch (validationResult.errorCode) {
                case 'MISMATCH':
                  Log.TestCoverage_gs({ file: sourceFile, coverage: 'handleSheetAutomations_bundleInvalid_mismatch' });
                  const currentInvalidValues = {
                    term: sheet.getRange(editedRowStart, combinedIndexes.aeTerm).getValue(),
                    quantity: sheet.getRange(editedRowStart, combinedIndexes.aeQuantity).getValue()
                  };
                   Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Calling Mismatch Dialog. Row=${editedRowStart}, Bundle=${currentBundleNum}, Current=${JSON.stringify(currentInvalidValues)}, Expected=${JSON.stringify(validationResult.expected)}`);
                  
                  // Call the dialog with all required parameters, including the bundleNumber
                  showBundleMismatchDialog(editedRowStart, currentBundleNum, currentInvalidValues, validationResult.expected);
                  break;

                case 'GAP_DETECTED':
                   Log.TestCoverage_gs({ file: sourceFile, coverage: 'handleSheetAutomations_bundleInvalid_gap' });
                   Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Calling Gap Dialog for bundle #${currentBundleNum}`);
                   showBundleGapDialog(currentBundleNum);
                  break;

                default:
                  Log.TestCoverage_gs({ file: sourceFile, coverage: 'handleSheetAutomations_bundleInvalid_defaultToast' });
                  // MODIFICATION: The line that reverted the edit has been removed.
                  SpreadsheetApp.getActive().toast(validationResult.errorMessage, "Validation Error", 10);
                  break;
              }
              ExecutionTimer.end('handleSheetAutomations_bundleInvalidDialog');
            }
        }
      }
      ExecutionTimer.end('handleSheetAutomations_bundleLogic');
    }

  } finally {
    Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_lock_released' });
    lock.releaseLock();
    Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Lock released.`);
  }
  Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'handleSheetAutomations_end' });
  ExecutionTimer.end('handleSheetAutomations_total');
}


// In SheetCoreAutomations.gs

/**
* Calculates and updates BOTH the LRF and Contract Value for a specific row's in-memory data,
* AND applies the correct number format to the corresponding cells.
*/
function updateCalculationsForRow(sheet, rowNum, rowValues, isTelekomDeal, colIndexes, approvalWorkflowConfig, dataBlockStartCol) {
  const sourceFile = "SheetCoreAutomations_gs";
  ExecutionTimer.start('updateCalculationsForRow_total');
  Log.TestCoverage_gs({ file: 'SheetCoreAutomations.gs', coverage: 'updateCalculationsForRow_start' });
  const statusStrings = approvalWorkflowConfig.statusStrings;
  let rentalPrice = 0;
  const status = rowValues[colIndexes.status - dataBlockStartCol];
  const approvedStatuses = [statusStrings.approvedOriginal, statusStrings.approvedNew];
  ExecutionTimer.start('updateCalculationsForRow_getPrice');
  if (approvedStatuses.includes(status)) {
    rentalPrice = getNumericValue(rowValues[colIndexes.financeApprovedPrice - dataBlockStartCol]);
  } else {
    const approverPrice = getNumericValue(rowValues[colIndexes.approverPriceProposal - dataBlockStartCol]);
    const aeSalesAskPrice = getNumericValue(rowValues[colIndexes.aeSalesAskPrice - dataBlockStartCol]);
    rentalPrice = (approverPrice > 0) ? approverPrice : aeSalesAskPrice;
  }
  ExecutionTimer.end('updateCalculationsForRow_getPrice');
  
  Log[sourceFile](`[${sourceFile} - updateCalculationsForRow] Row ${rowNum}: Inputs: rentalPrice=${rentalPrice}, isTelekomDeal=${isTelekomDeal}`);
  const epCapex = getNumericValue(rowValues[colIndexes.aeEpCapex - dataBlockStartCol]);
  const tkCapex = getNumericValue(rowValues[colIndexes.aeTkCapex - dataBlockStartCol]);
  let chosenCapex = isTelekomDeal ? tkCapex : epCapex;
  Log[sourceFile](`[${sourceFile} - updateCalculationsForRow] Row ${rowNum}: Capex values: EP=${epCapex}, TK=${tkCapex}. Chosen Capex: ${chosenCapex}`);
  
  ExecutionTimer.start('updateCalculationsForRow_calcLrf');
  const lrfCell = sheet.getRange(rowNum, colIndexes.lrfPreview);
  const contractValueCell = sheet.getRange(rowNum, colIndexes.contractValuePreview);
  const formats = CONFIG.numberFormats;

  if (rentalPrice === 0 && (!chosenCapex || chosenCapex === 0)) {
    rowValues[colIndexes.lrfPreview - dataBlockStartCol] = "";
    rowValues[colIndexes.contractValuePreview - dataBlockStartCol] = "";
    Log[sourceFile](`[${sourceFile} - updateCalculationsForRow] Row ${rowNum}: Clearing LRF and Contract Value due to zero price/capex.`);
  } else {
    if (!chosenCapex || chosenCapex <= 0) {
      rowValues[colIndexes.lrfPreview - dataBlockStartCol] = `Missing\n${isTelekomDeal ? 'TK' : 'EP'} CAPEX`;
    } else {
      const term = getNumericValue(rowValues[colIndexes.aeTerm - dataBlockStartCol]);
      if (chosenCapex > 0 && rentalPrice > 0 && term > 0) {
        rowValues[colIndexes.lrfPreview - dataBlockStartCol] = (rentalPrice * term) / chosenCapex;
      } else {
        rowValues[colIndexes.lrfPreview - dataBlockStartCol] = "";
      }
    }
    const quantity = getNumericValue(rowValues[colIndexes.aeQuantity - dataBlockStartCol]);
    const term = getNumericValue(rowValues[colIndexes.aeTerm - dataBlockStartCol]);
    if (rentalPrice > 0 && term > 0 && quantity > 0) {
      rowValues[colIndexes.contractValuePreview - dataBlockStartCol] = rentalPrice * term * quantity;
    } else {
      rowValues[colIndexes.contractValuePreview - dataBlockStartCol] = "";
    }
  }
  ExecutionTimer.end('updateCalculationsForRow_calcLrf');

  // --- NEW: Apply Number Formatting Directly ---
  ExecutionTimer.start('updateCalculationsForRow_setFormats');
  lrfCell.setNumberFormat(formats.percentage);
  contractValueCell.setNumberFormat(formats.currency);
  ExecutionTimer.end('updateCalculationsForRow_setFormats');

  Log[sourceFile](`[${sourceFile} - updateCalculationsForRow] Row ${rowNum}: Outputs: LRF=${rowValues[colIndexes.lrfPreview - dataBlockStartCol]}, ContractValue=${rowValues[colIndexes.contractValuePreview - dataBlockStartCol]}`);
  ExecutionTimer.end('updateCalculationsForRow_total');
}


