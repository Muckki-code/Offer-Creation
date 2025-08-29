/**
* @file This file handles the logic for managing row statuses
* within the approval workflow, including AE input and revision handling.
*/

/**
 * Checks if a row contains all the required data for its status to be "Pending Approval".
 * This function is the definitive source for this business rule.
 * REVISED: Now includes a check for the relevant Capex based on the deal type.
 * @private
 * @param {Array<any>} rowValues The in-memory array of values for the row.
 * @param {Object} colIndexes A map of column names to their 1-based index.
 * @param {number} startCol The 1-based index of the starting column for the rowValues array.
 * @param {boolean} isTelekomDeal Whether the current sheet is marked as a Telekom Deal.
 * @returns {boolean} True if all required data is present, false otherwise.
 */
function _isRowDataComplete(rowValues, colIndexes, startCol, isTelekomDeal) {
 const sourceFile = 'SheetStatusLogic_gs';
 ExecutionTimer.start('_isRowDataComplete_total');
 Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: '_isRowDataComplete_start' });
 Log[sourceFile](`[${sourceFile} - _isRowDataComplete] Start. Checking row for completeness. isTelekomDeal=${isTelekomDeal}`);

 const requiredFields = {
  "Model": { index: colIndexes.model, value: rowValues[colIndexes.model - startCol] },
  "Sales Ask Price": { index: colIndexes.aeSalesAskPrice, value: getNumericValue(rowValues[colIndexes.aeSalesAskPrice - startCol]) },
  "Quantity": { index: colIndexes.aeQuantity, value: getNumericValue(rowValues[colIndexes.aeQuantity - startCol]) },
  "Term": { index: colIndexes.aeTerm, value: getNumericValue(rowValues[colIndexes.aeTerm - startCol]) }
 };

 // Conditionally add the required Capex field to the check
 if (isTelekomDeal) {
    requiredFields["Telekom Capex"] = { index: colIndexes.aeTkCapex, value: getNumericValue(rowValues[colIndexes.aeTkCapex - startCol]) };
 } else {
    requiredFields["EP Capex"] = { index: colIndexes.aeEpCapex, value: getNumericValue(rowValues[colIndexes.aeEpCapex - startCol]) };
 }

 ExecutionTimer.start('_isRowDataComplete_loop');
 for (const fieldName in requiredFields) {
  Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: '_isRowDataComplete_loop_iteration' });
  const field = requiredFields[fieldName];
  if (!field.value || (typeof field.value === 'number' && field.value <= 0)) {
   Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: '_isRowDataComplete_fail' });
   Log[sourceFile](`[${sourceFile} - _isRowDataComplete] FAILED: Row is incomplete. Missing or invalid required field: '${fieldName}'. Value: '${field.value}'.`);
   ExecutionTimer.end('_isRowDataComplete_loop');
   ExecutionTimer.end('_isRowDataComplete_total');
   return false;
  }
 }
 ExecutionTimer.end('_isRowDataComplete_loop');

 Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: '_isRowDataComplete_pass' });
 Log[sourceFile](`[${sourceFile} - _isRowDataComplete] PASSED: All required fields are present.`);
 ExecutionTimer.end('_isRowDataComplete_total');
 return true;
}


// In SheetStatusLogic.gs

/**
 * Analyzes a row's state and determines what its new status should be.
 * REFACTORED: This is now a "pure" function. It is READ-ONLY and does not
 * modify any data arrays. It only returns the calculated new status string (or null).
 * This prevents destructive side-effect bugs from race conditions.
 *
 * @param {Array<any>} inMemoryRowValues The current, in-memory array of values for the row.
 * @param {Array<any>} originalFullRowValuesFromCaller The array of values for the row before the edit.
 * @param {boolean} isTelekomDeal Whether the current sheet is marked as a Telekom Deal.
 * @param {Object} options An optional options object. Can contain { forceRevisionOfFinalizedItems: boolean }.
 * @param {number} startCol The 1-based index of the starting column for the rowValues array.
 * @param {Object} allColIndexes A map of column names to their 1-based index.
 * @returns {string|null} The calculated new status string, or null if the row should be completely cleared.
 */
function updateStatusForRow(inMemoryRowValues, originalFullRowValuesFromCaller, isTelekomDeal, options, startCol, allColIndexes) {
 const sourceFile = 'SheetStatusLogic_gs';
 ExecutionTimer.start('updateStatusForRow_total');
 Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'updateStatusForRow_start' });
 
 const wfConfig = CONFIG.approvalWorkflow;
 const statusStrings = wfConfig.statusStrings;
 const colIndexes = allColIndexes;

 const initialStatus = originalFullRowValuesFromCaller[colIndexes.status - startCol] || "";
 const originalModel = originalFullRowValuesFromCaller[colIndexes.model - startCol];
 const currentModel = inMemoryRowValues[colIndexes.model - startCol];
 
 // Start with the assumption that the status will not change.
 let newStatus = initialStatus;
 Log[sourceFile](`[${sourceFile} - updateStatusForRow] Analyze Start: initialStatus='${initialStatus}'.`);

 ExecutionTimer.start('updateStatusForRow_stateMachine');
 const finalizedStatuses = [statusStrings.approvedOriginal, statusStrings.approvedNew, statusStrings.rejected];
 
 if (originalModel && !currentModel) {
    Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'updateStatusForRow_rule_modelDeleted' });
    Log[sourceFile](`[${sourceFile} - updateStatusForRow] Rule #1 Result: Model deleted. Returning null to signal a full clear.`);
    newStatus = null;
 } 
 else if (finalizedStatuses.includes(initialStatus) && (wasKeyFieldEdited(inMemoryRowValues, originalFullRowValuesFromCaller, colIndexes, startCol) || options.forceRevisionOfFinalizedItems)) {
    Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'updateStatusForRow_rule_keyFieldEditOnFinalized' });
    const reason = options.forceRevisionOfFinalizedItems ? "bulk recalculation" : "a key field was edited";
    Log[sourceFile](`[${sourceFile} - updateStatusForRow] Rule #2 Result: Finalized item revised due to ${reason}. Setting status to 'Revised by AE'.`);
    newStatus = statusStrings.revisedByAE;
 }
 else if (wasKeyFieldEdited(inMemoryRowValues, originalFullRowValuesFromCaller, colIndexes, startCol) || (!originalModel && currentModel)) {
    Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'updateStatusForRow_rule_nonFinalizedEdit' });
    Log[sourceFile](`[${sourceFile} - updateStatusForRow] Rule #3 Result: Non-finalized item edit or new model. Setting status to 'Draft'.`);
    newStatus = statusStrings.draft;
 }
 ExecutionTimer.end('updateStatusForRow_stateMachine');
 
 ExecutionTimer.start('updateStatusForRow_finalStateDetermination');
 if (newStatus !== null) {
    const hasRequiredData = _isRowDataComplete(inMemoryRowValues, colIndexes, startCol, isTelekomDeal);
    if (hasRequiredData) {
        Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'updateStatusForRow_final_hasData' });
        if (newStatus === statusStrings.draft || newStatus === "") {
            Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'updateStatusForRow_final_draftToPending' });
            newStatus = statusStrings.pending;
            Log[sourceFile](`[${sourceFile} - updateStatusForRow] Final State Check: Row has required data. Promoting status to '${newStatus}'.`);
        }
    } else {
        Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'updateStatusForRow_final_missingData' });
        if (currentModel) {
            Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'updateStatusForRow_final_forceToDraft' });
            newStatus = statusStrings.draft;
            Log[sourceFile](`[${sourceFile} - updateStatusForRow] Final State Check: Row is missing required data. Forcing status to '${newStatus}'.`);
        } else if (newStatus !== initialStatus && !finalizedStatuses.includes(initialStatus)) {
            // This handles a paste/delete that results in a totally blank row. It should have a blank status.
            Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'updateStatusForRow_final_forceToBlank' });
            Log[sourceFile](`[${sourceFile} - updateStatusForRow] Final State Check: Row has no model. Forcing status to blank.`);
            newStatus = "";
        }
    }
 }
 ExecutionTimer.end('updateStatusForRow_finalStateDetermination');

 Log[sourceFile](`[${sourceFile} - updateStatusForRow] Analyze End. Final calculated status is '${newStatus}'.`);
 ExecutionTimer.end('updateStatusForRow_total');
 
 return newStatus;
}


/**
* Helper to determine if a key data field was edited by an AE.
* @private
* @param {Array<any>} currentRow The current in-memory row data.
* @param {Array<any>} originalRow The original row data.
* @param {Object} colIndexes A map of column names to their 1-based index.
* @param {number} startCol The 1-based index of the starting column for the arrays.
* @returns {boolean} True if a key field was changed, false otherwise.
*/
function wasKeyFieldEdited(currentRow, originalRow, colIndexes, startCol) {
  const sourceFile = 'SheetStatusLogic_gs';
  ExecutionTimer.start('wasKeyFieldEdited_total');
  Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'wasKeyFieldEdited_start' });

  const revisionTriggerCols = [
    colIndexes.model, colIndexes.aeSalesAskPrice, colIndexes.aeQuantity, colIndexes.aeTerm,
    colIndexes.aeEpCapex, colIndexes.aeTkCapex
  ];

  ExecutionTimer.start('wasKeyFieldEdited_loop');
  for (const colIndex of revisionTriggerCols) {
    if (String(currentRow[colIndex - startCol] || "") !== String(originalRow[colIndex - startCol] || "")) {
      Log.TestCoverage_gs({ file: 'SheetStatusLogic.gs', coverage: 'wasKeyFieldEdited_fallbackPath_isChanged' });
      Log[sourceFile](`[${sourceFile} - wasKeyFieldEdited] Result: true (change detected in column ${colIndex}).`);
      ExecutionTimer.end('wasKeyFieldEdited_loop');
      ExecutionTimer.end('wasKeyFieldEdited_total');
      return true;
    }
  }
  ExecutionTimer.end('wasKeyFieldEdited_loop');

  Log[sourceFile](`[${sourceFile} - wasKeyFieldEdited] Result: false. No key fields were changed.`);
  ExecutionTimer.end('wasKeyFieldEdited_total');
  return false;
}