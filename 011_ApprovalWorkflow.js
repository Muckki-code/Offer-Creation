/**
 * @file This file contains the logic for handling approver actions.
 * It has been refactored to support a direct, single-action approval workflow.
 */

/**
 * Processes a single approval action immediately upon an approver changing the dropdown.
 * This is the new core function for the approval workflow.
 */
function processSingleApprovalAction(sheet, rowNum, e, inMemoryRowValues, allColIndexes, originalFullRowValues, startCol) {
    const sourceFile = "ApprovalWorkflow_gs";
    ExecutionTimer.start('processSingleApprovalAction_total');
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_start' });
    Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] START. rowNum=${rowNum}, startCol=${startCol}, newValue='${e.value}', oldValue='${e.oldValue}'`);

    const config = CONFIG.approvalWorkflow;
    const statusStrings = config.statusStrings;
    const approverAction = e.value;
    const oldApproverAction = e.oldValue;

    if (!approverAction || approverAction === "Choose Action") {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_noAction' });
        Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] Condition: No valid action selected. Exiting.`);
        ExecutionTimer.end('processSingleApprovalAction_total');
        return false;
    }

    const bundleNumber = inMemoryRowValues[allColIndexes.bundleNumber - startCol];
    if (bundleNumber) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_isBundle' });
        Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] Row ${rowNum} is part of bundle #${bundleNumber}. Performing bundle validation before processing.`);
        const validationResult = validateBundle(sheet, rowNum, bundleNumber);
        if (!validationResult.isValid) {
            Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_bundleInvalid' });
            Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] BUNDLE VALIDATION FAIL: Bundle #${bundleNumber} is invalid. Reason: ${validationResult.errorMessage}. Reverting action.`);
            
            inMemoryRowValues[allColIndexes.approverAction - startCol] = oldApproverAction; 
            SpreadsheetApp.getActive().toast(`Action blocked for row ${rowNum}. Bundle #${bundleNumber} has an error that must be fixed first.`, "Bundle Invalid", 8);
            logGeneralActivity({ action: "Approval Action Blocked", details: `Row ${rowNum}: Bundle #${bundleNumber} is invalid.`, sheetName: sheet.getName(), row: rowNum });
            
            ExecutionTimer.end('processSingleApprovalAction_total');
            return false; 
        }
        Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] Bundle #${bundleNumber} is valid. Proceeding with action processing.`);
    }

    ExecutionTimer.start('processSingleApprovalAction_validation');
    const currentStatus = inMemoryRowValues[allColIndexes.status - startCol];
    if (currentStatus !== statusStrings.pending && currentStatus !== statusStrings.revisedByAE) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_invalidStatus' });
        Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] VALIDATION FAIL: Row ${rowNum} has status '${currentStatus}', which is not processable. Reverting action.`);
        SpreadsheetApp.getActive().toast(`Row ${rowNum} cannot be processed because its status is '${currentStatus}'.`);
        inMemoryRowValues[allColIndexes.approverAction - startCol] = oldApproverAction;
        logGeneralActivity({ action: "Approval Action Failed", details: `Row ${rowNum}: Invalid status '${currentStatus}'.`, sheetName: sheet.getName(), row: rowNum });
        ExecutionTimer.end('processSingleApprovalAction_validation');
        ExecutionTimer.end('processSingleApprovalAction_total');
        return false;
    }

    Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] Validation Passed: Row ${rowNum} is in a processable status ('${currentStatus}').`);
    let newStatus = currentStatus;
    let finalPrice = null;
    let validationError = "";
    const lrfPreview = getNumericValue(inMemoryRowValues[allColIndexes.lrfPreview - startCol]);
    const isApprovalAction = approverAction.includes("Approve");

    if (isApprovalAction && (isNaN(lrfPreview) || lrfPreview <= 0)) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_invalidLrf' });
        validationError = `Row ${rowNum}: Cannot approve with an invalid or missing LRF.`;
        Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] VALIDATION FAIL: LRF check failed. LRF value: '${lrfPreview}'.`);
    } else {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_validLrf' });
        Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] LRF check passed for action '${approverAction}'. LRF: ${lrfPreview}`);
        switch (approverAction) {
            case "Approve Original Price":
                Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_approveOriginal' });
                const originalPrice = getNumericValue(inMemoryRowValues[allColIndexes.aeSalesAskPrice - startCol]);
                if (!originalPrice || originalPrice <= 0) {
                    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_approveOriginal_invalidPrice' });
                    validationError = `Row ${rowNum}: Cannot 'Approve Original Price' without a valid 'AE Sales Ask Price'.`;
                } else {
                    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_approveOriginal_validPrice' });
                    finalPrice = originalPrice;
                    newStatus = statusStrings.approvedOriginal;
                }
                break;
            case "Approve New Price":
                Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_approveNew' });
                const proposedPrice = getNumericValue(inMemoryRowValues[allColIndexes.approverPriceProposal - startCol]);
                if (!proposedPrice || proposedPrice <= 0) {
                    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_approveNew_invalidPrice' });
                    validationError = `Row ${rowNum}: Cannot 'Approve New Price' without a valid 'Approver Price Proposal'.`;
                } else {
                    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_approveNew_validPrice' });
                    finalPrice = proposedPrice;
                    newStatus = statusStrings.approvedNew;
                }
                break;
            case "Reject with Comment":
                Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_reject' });
                if (!inMemoryRowValues[allColIndexes.approverComments - startCol] || String(inMemoryRowValues[allColIndexes.approverComments - startCol]).trim() === '') {
                    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_reject_noComment' });
                    validationError = `Row ${rowNum}: Cannot 'Reject with Comment' without adding a comment.`;
                } else {
                    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_reject_hasComment' });
                    newStatus = statusStrings.rejected;
                }
                break;
        }
    }
    ExecutionTimer.end('processSingleApprovalAction_validation');

    ExecutionTimer.start('processSingleApprovalAction_applyChanges');
    if (validationError) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_validationFailed' });
        Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] FINAL: Validation Error Found: '${validationError}'. Reverting action.`);
        SpreadsheetApp.getActive().toast(validationError, "Action Blocked", 6);
        inMemoryRowValues[allColIndexes.approverAction - startCol] = oldApproverAction;
        logGeneralActivity({ action: "Approval Action Failed", details: `Row ${rowNum}: Validation error - '${validationError}'.`, sheetName: sheet.getName(), row: rowNum });
        ExecutionTimer.end('processSingleApprovalAction_applyChanges');
        ExecutionTimer.end('processSingleApprovalAction_total');
        return false;
    } else {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_success' });
        Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] FINAL: All validation passed. Applying changes to row ${rowNum}.`);
        const approverEmail = Session.getActiveUser().getEmail();
        const timestamp = new Date();
        inMemoryRowValues[allColIndexes.status - startCol] = newStatus;
        inMemoryRowValues[allColIndexes.approvedBy - startCol] = approverEmail;
        inMemoryRowValues[allColIndexes.approvalDate - startCol] = timestamp;
        if (finalPrice !== null) {
            Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_setFinalPrice' });
            inMemoryRowValues[allColIndexes.financeApprovedPrice - startCol] = finalPrice;
        } else {
            Log.TestCoverage_gs({ file: sourceFile, coverage: 'processSingleApprovalAction_noFinalPrice' });
        }
        // --- THIS IS THE FIX ---
        logTableActivity({ mainSheet: sheet, rowNum: rowNum, oldStatus: currentStatus, newStatus: newStatus, currentFullRowValues: inMemoryRowValues, originalFullRowValues: originalFullRowValues, startCol: startCol });
        Log[sourceFile](`[${sourceFile} - processSingleApprovalAction] Successfully processed row ${rowNum}. New status: '${newStatus}'.`);
        ExecutionTimer.end('processSingleApprovalAction_applyChanges');
        ExecutionTimer.end('processSingleApprovalAction_total');
        return true;
    }
}



/**
 * Performs a health check on the sheet to find and fix data inconsistencies.
 */
function runSheetHealthCheck() {
  const sourceFile = 'ApprovalWorkflow_gs';
  ExecutionTimer.start('runSheetHealthCheck_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'runSheetHealthCheck_start' });
  Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Start: Running new sheet data health check.`);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = getLastLastRow(sheet);
  Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Sheet: '${sheet.getName()}', Data Start Row: ${startRow}, Last Row: ${lastRow}.`);
  if (lastRow < startRow) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'runSheetHealthCheck_noData' });
    Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] No data rows to check. Exiting.`);
    ExecutionTimer.end('runSheetHealthCheck_total');
    return;
  }
  const startCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - startCol + 1;
  Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Reading data from startCol=${startCol} with numCols=${numCols}.`);
  ExecutionTimer.start('runSheetHealthCheck_readData');
  const range = sheet.getRange(startRow, startCol, lastRow - startRow + 1, numCols);
  const allValues = range.getValues();
  ExecutionTimer.end('runSheetHealthCheck_readData');
  Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Read ${allValues.length} rows from range ${range.getA1Notation()}.`);
  const colIndexes = CONFIG.approvalWorkflow.columnIndices;
  const statusStrings = CONFIG.approvalWorkflow.statusStrings;
  const finalizedStatuses = [
    statusStrings.approvedOriginal,
    statusStrings.approvedNew,
    statusStrings.rejected
  ];
  let fixesMade = 0;
  let changesToLog = [];
  ExecutionTimer.start('runSheetHealthCheck_mainLoop');
  Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Starting main health check loop.`);
  for (let i = 0; i < allValues.length; i++) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'runSheetHealthCheck_loop_iteration' });
    const row = allValues[i];
    const currentRowNum = startRow + i;
    // CRAZY VERBOSE LOGGING
    Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Evaluating Row ${currentRowNum}: ${JSON.stringify(row)}`);
    const status = row[colIndexes.status - startCol];
    const approvalDate = row[colIndexes.approvalDate - startCol];
    if (finalizedStatuses.includes(status) && !approvalDate) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'runSheetHealthCheck_fixFound' });
      Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] FIX FOUND on row ${currentRowNum}: Status is '${status}' but 'Approval Date' is empty.`);
      const originalStatus = status;
      const originalRowForLog = JSON.parse(JSON.stringify(row));
      // CRAZY VERBOSE LOGGING
      Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Row ${currentRowNum} BEFORE fix: ${JSON.stringify(originalRowForLog)}`);
      
      row[colIndexes.status - startCol] = statusStrings.pending;
      row[colIndexes.financeApprovedPrice - startCol] = "";
      row[colIndexes.approvedBy - startCol] = "";
      row[colIndexes.approverAction - startCol] = "Choose Action";
      
      changesToLog.push({
          rowNum: currentRowNum, oldStatus: originalStatus, newStatus: statusStrings.pending,
          currentFullRowValues: row, originalFullRowValues: originalRowForLog, startCol: startCol
      });
      fixesMade++;
      // CRAZY VERBOSE LOGGING
      Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Row ${currentRowNum} AFTER fix: ${JSON.stringify(row)}`);
    } else {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'runSheetHealthCheck_noFixNeeded' });
    }
  }
  ExecutionTimer.end('runSheetHealthCheck_mainLoop');
  Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Health check loop finished. Found ${fixesMade} inconsistencies.`);
  if (fixesMade > 0) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'runSheetHealthCheck_writingFixes' });
    ExecutionTimer.start('runSheetHealthCheck_writeData');
    Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Writing ${fixesMade} fixes back to the sheet.`);
    range.setValues(allValues);
    changesToLog.forEach(logInfo => logTableActivity(logInfo));
    SpreadsheetApp.getActive().toast(`Health Check Complete: Reverted ${fixesMade} rows with inconsistent status back to 'Pending Approval'.`, "Sheet Repaired", 6);
    ExecutionTimer.end('runSheetHealthCheck_writeData');
  } else {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'runSheetHealthCheck_noFixesWritten' });
    Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] Health check complete. No inconsistencies found.`);
    SpreadsheetApp.getActive().toast('Health Check Complete: No inconsistencies found.', "Sheet Health Check");
  }
  ExecutionTimer.end('runSheetHealthCheck_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'runSheetHealthCheck_end' });
  Log[sourceFile](`[${sourceFile} - runSheetHealthCheck] END.`);
}


// In ApprovalWorkflow.gs

/**
 * Processes all approvable items in the sheet in a single batch operation.
 * Called from the "Process All Set Actions" button in the sidebar.
 * - Ignores 'Draft' and 'Rejected' items.
 * - Approves with new price if a valid proposal exists.
 * - Approves with original price otherwise.
 */
function processAllApproverActions() {
  const sourceFile = "ApprovalWorkflow_gs";
  ExecutionTimer.start('processAllApproverActions_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'processAllApproverActions_start' });
  Log[sourceFile](`[${sourceFile} - processAllApproverActions] Start.`);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const config = CONFIG.approvalWorkflow;
  const startRow = config.startDataRow;
  const lastRow = getLastLastRow(sheet);

  if (lastRow < startRow) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processAllApproverActions_noDataRows' });
    ui.alert("No data rows found to process.");
    ExecutionTimer.end('processAllApproverActions_total');
    return;
  }

  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numColsInDataBlock = CONFIG.maxDataColumn - dataBlockStartCol + 1;
  const dataRange = sheet.getRange(startRow, dataBlockStartCol, lastRow - startRow + 1, numColsInDataBlock);

  ExecutionTimer.start('processAllApproverActions_readSheet');
  const allValues = dataRange.getValues();
  const originalValues = JSON.parse(JSON.stringify(allValues)); // Deep copy for logging
  ExecutionTimer.end('processAllApproverActions_readSheet');

  const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };
  const statusStrings = config.statusStrings;
  const processableStatuses = [statusStrings.pending, statusStrings.revisedByAE];
  const approverEmail = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  let itemsProcessed = 0;

  ExecutionTimer.start('processAllApproverActions_mainLoop');
  allValues.forEach((row, index) => {
    const originalRow = originalValues[index];
    const initialStatus = originalRow[c.status - dataBlockStartCol];
    
    if (processableStatuses.includes(initialStatus)) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'processAllApproverActions_rowIsProcessable' });
      const proposedPrice = getNumericValue(row[c.approverPriceProposal - dataBlockStartCol]);
      let newStatus = "";
      let finalPrice = 0;

      if (proposedPrice > 0) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'processAllApproverActions_approveNewPrice' });
        newStatus = statusStrings.approvedNew;
        finalPrice = proposedPrice;
      } else {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'processAllApproverActions_approveOriginalPrice' });
        newStatus = statusStrings.approvedOriginal;
        finalPrice = getNumericValue(row[c.aeSalesAskPrice - dataBlockStartCol]);
      }

      // Apply the changes to the in-memory array
      row[c.status - dataBlockStartCol] = newStatus;
      row[c.financeApprovedPrice - dataBlockStartCol] = finalPrice;
      row[c.approvedBy - dataBlockStartCol] = approverEmail;
      row[c.approvalDate - dataBlockStartCol] = timestamp;
      
      logTableActivity({
        mainSheet: sheet,
        rowNum: startRow + index,
        oldStatus: initialStatus,
        newStatus: newStatus,
        currentFullRowValues: row,
        originalFullRowValues: originalRow,
        startCol: dataBlockStartCol
      });
      itemsProcessed++;
    } else {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'processAllApproverActions_rowIsNotProcessable' });
    }
  });
  ExecutionTimer.end('processAllApproverActions_mainLoop');

  if (itemsProcessed > 0) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processAllApproverActions_writingChanges' });
    ExecutionTimer.start('processAllApproverActions_writeSheet');
    dataRange.setValues(allValues);
    ExecutionTimer.end('processAllApproverActions_writeSheet');
    ui.alert(`Bulk approval complete. ${itemsProcessed} item(s) have been processed.`);
    Log[sourceFile](`[${sourceFile} - processAllApproverActions] Wrote changes for ${itemsProcessed} items to the sheet.`);
  } else {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processAllApproverActions_noItemsFound' });
    ui.alert("No items with status 'Pending Approval' or 'Revised by AE' were found to process.");
    Log[sourceFile](`[${sourceFile} - processAllApproverActions] No processable items were found.`);
  }
  
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'processAllApproverActions_end' });
  Log[sourceFile](`[${sourceFile} - processAllApproverActions] End.`);
  ExecutionTimer.end('processAllApproverActions_total');
}