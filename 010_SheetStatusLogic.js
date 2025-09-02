/**
 * @file This file handles the logic for managing row statuses
 * within the approval workflow, including AE input and revision handling.
 */

/**
 * Checks if a row contains all the required data for its status to be "Pending Approval".
 * This function is the definitive source for this business rule.
 * REVISED: Now checks for the new single aeCapex column, removing the old conditional logic.
 * @private
 * @param {Array<any>} rowValues The in-memory array of values for the row.
 * @param {Object} colIndexes A map of column names to their 1-based index.
 * @param {number} startCol The 1-based index of the starting column for the rowValues array.
 * @param {boolean} isTelekomDeal Whether the current sheet is marked as a Telekom Deal (parameter kept for signature compatibility).
 * @returns {boolean} True if all required data is present, false otherwise.
 */
function _isRowDataComplete(rowValues, colIndexes, startCol, isTelekomDeal) {
  const sourceFile = "SheetStatusLogic_gs";
  ExecutionTimer.start("_isRowDataComplete_total");
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "_isRowDataComplete_start",
  });
  Log[sourceFile](
    `[${sourceFile} - _isRowDataComplete] Start. Checking row for completeness.`
  );

  const requiredFields = {
    Model: {
      index: colIndexes.model,
      value: rowValues[colIndexes.model - startCol],
    },
    "AE Capex": {
      index: colIndexes.aeCapex,
      value: getNumericValue(rowValues[colIndexes.aeCapex - startCol]),
    },
    "Sales Ask Price": {
      index: colIndexes.aeSalesAskPrice,
      value: getNumericValue(rowValues[colIndexes.aeSalesAskPrice - startCol]),
    },
    Quantity: {
      index: colIndexes.aeQuantity,
      value: getNumericValue(rowValues[colIndexes.aeQuantity - startCol]),
    },
    Term: {
      index: colIndexes.aeTerm,
      value: getNumericValue(rowValues[colIndexes.aeTerm - startCol]),
    },
  };

  ExecutionTimer.start("_isRowDataComplete_loop");
  for (const fieldName in requiredFields) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "_isRowDataComplete_loop_iteration",
    });
    const field = requiredFields[fieldName];
    const fieldValue =
      typeof field.value === "string" ? `"${field.value}"` : field.value;
    Log[sourceFile](
      `[${sourceFile} - _isRowDataComplete] CRAZY VERBOSE: Validating field '${fieldName}', Value: ${fieldValue}`
    );

    if (!field.value || (typeof field.value === "number" && field.value <= 0)) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "_isRowDataComplete_fail",
      });
      Log[sourceFile](
        `[${sourceFile} - _isRowDataComplete] FAILED: Row is incomplete. Missing or invalid required field: '${fieldName}'. Value: '${field.value}'.`
      );
      ExecutionTimer.end("_isRowDataComplete_loop");
      ExecutionTimer.end("_isRowDataComplete_total");
      return false;
    }
  }
  ExecutionTimer.end("_isRowDataComplete_loop");

  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "_isRowDataComplete_pass",
  });
  Log[sourceFile](
    `[${sourceFile} - _isRowDataComplete] PASSED: All required fields are present.`
  );
  ExecutionTimer.end("_isRowDataComplete_total");
  return true;
}

/**
 * Analyzes a row's state and determines what its new status should be.
 * REFACTORED: This is now a "pure" function. It is READ-ONLY and does not
 * modify any data arrays. It only returns the calculated new status string (or null).
 * This prevents destructive side-effect bugs from race conditions.
 * REVISED: Now includes context-aware logic to differentiate between a direct user edit
 * and a programmatic recalculation via the 'options' parameter.
 * @param {Array<any>} inMemoryRowValues The current, in-memory array of values for the row.
 * @param {Array<any>} originalFullRowValuesFromCaller The array of values for the row before the edit.
 * @param {boolean} isTelekomDeal Whether the current sheet is marked as a Telekom Deal.
 * @param {Object} options An optional options object. Can contain { forceRevisionOfFinalizedItems: boolean }.
 * @param {number} startCol The 1-based index of the starting column for the rowValues array.
 * @param {Object} allColIndexes A map of column names to their 1-based index.
 * @returns {string|null} The calculated new status string, or null if the row should be completely cleared.
 */
function updateStatusForRow(
  inMemoryRowValues,
  originalFullRowValuesFromCaller,
  isTelekomDeal,
  options,
  startCol,
  allColIndexes
) {
  const sourceFile = "SheetStatusLogic_gs";
  ExecutionTimer.start("updateStatusForRow_total");
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "updateStatusForRow_start",
  });

  const wfConfig = CONFIG.approvalWorkflow;
  const statusStrings = wfConfig.statusStrings;
  const colIndexes = allColIndexes;

  const initialStatus =
    originalFullRowValuesFromCaller[colIndexes.status - startCol] || "";
  const originalModel =
    originalFullRowValuesFromCaller[colIndexes.model - startCol];
  const currentModel = inMemoryRowValues[colIndexes.model - startCol];

  let newStatus = initialStatus;
  Log[sourceFile](
    `[${sourceFile} - updateStatusForRow] Analyze Start: initialStatus='${initialStatus}', options=${JSON.stringify(
      options
    )}.`
  );

  ExecutionTimer.start("updateStatusForRow_stateMachine");
  const finalizedStatuses = [
    statusStrings.approvedOriginal,
    statusStrings.approvedNew,
    statusStrings.rejected,
  ];
  const keyFieldWasEdited = wasKeyFieldEdited(
    inMemoryRowValues,
    originalFullRowValuesFromCaller,
    colIndexes,
    startCol
  );

  // --- REVISED STATE MACHINE ---
  if (originalModel && !currentModel) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateStatusForRow_rule_modelDeleted",
    });
    Log[sourceFile](
      `[${sourceFile} - updateStatusForRow] Rule #1: Model was deleted. Returning null to signal a full row clear.`
    );
    newStatus = null;
  } else if (finalizedStatuses.includes(initialStatus) && keyFieldWasEdited) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateStatusForRow_rule_keyFieldEditOnFinalized",
    });
    Log[sourceFile](
      `[${sourceFile} - updateStatusForRow] Rule #2: A finalized item was edited. Setting status to 'Revised by AE'.`
    );
    newStatus = statusStrings.revisedByAE;
  } else if (
    (!originalModel && currentModel) ||
    (!options.forceRevisionOfFinalizedItems && keyFieldWasEdited)
  ) {
    // --- THIS IS THE FIX ---
    // This block now only runs for a NEW row OR a "real" user edit (when forceRevision is false).
    // It is skipped during a bulk recalculation, which prevents the incorrect "Draft" reset.
    // --- END FIX ---
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateStatusForRow_rule_nonFinalizedEdit",
    });
    Log[sourceFile](
      `[${sourceFile} - updateStatusForRow] Rule #3: A non-finalized item was edited by a user, or it's a new row. Resetting status to 'Draft' for re-evaluation.`
    );
    newStatus = statusStrings.draft;
  }
  ExecutionTimer.end("updateStatusForRow_stateMachine");

  ExecutionTimer.start("updateStatusForRow_finalStateDetermination");
  if (newStatus !== null) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateStatusForRow_statusIsNotNull",
    });
    const hasRequiredData = _isRowDataComplete(
      inMemoryRowValues,
      colIndexes,
      startCol,
      isTelekomDeal
    );
    if (hasRequiredData) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "updateStatusForRow_final_hasData",
      });
      if (
        newStatus === statusStrings.draft ||
        newStatus === "" ||
        newStatus === statusStrings.revisedByAE
      ) {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "updateStatusForRow_final_promoteToPending",
        });
        newStatus = statusStrings.pending;
        Log[sourceFile](
          `[${sourceFile} - updateStatusForRow] Final State Check: Row has required data. Promoting/setting status to '${newStatus}'.`
        );
      }
    } else {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "updateStatusForRow_final_missingData",
      });
      if (currentModel) {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "updateStatusForRow_final_forceToDraft",
        });
        newStatus = statusStrings.draft;
        Log[sourceFile](
          `[${sourceFile} - updateStatusForRow] Final State Check: Row is missing required data. Forcing status to '${newStatus}'.`
        );
      } else if (
        newStatus !== initialStatus &&
        !finalizedStatuses.includes(initialStatus)
      ) {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "updateStatusForRow_final_forceToBlank",
        });
        Log[sourceFile](
          `[${sourceFile} - updateStatusForRow] Final State Check: Row has no model. Forcing status to blank.`
        );
        newStatus = "";
      }
    }
  } else {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateStatusForRow_statusIsNull",
    });
  }
  ExecutionTimer.end("updateStatusForRow_finalStateDetermination");

  Log[sourceFile](
    `[${sourceFile} - updateStatusForRow] Analyze End. Final calculated status is '${newStatus}'.`
  );
  ExecutionTimer.end("updateStatusForRow_total");

  return newStatus;
}

/**
 * Helper to determine if a key data field was edited by an AE.
 * REVISED: Now checks the single aeCapex column.
 * @private
 * @param {Array<any>} currentRow The current in-memory row data.
 * @param {Array<any>} originalRow The original row data.
 * @param {Object} colIndexes A map of column names to their 1-based index.
 * @param {number} startCol The 1-based index of the starting column for the arrays.
 * @returns {boolean} True if a key field was changed, false otherwise.
 */
function wasKeyFieldEdited(currentRow, originalRow, colIndexes, startCol) {
  const sourceFile = "SheetStatusLogic_gs";
  ExecutionTimer.start("wasKeyFieldEdited_total");
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "wasKeyFieldEdited_start",
  });
  Log[sourceFile](
    `[${sourceFile} - wasKeyFieldEdited] Start: Comparing current row against original to detect key field edits.`
  );

  // REFACTORED: Use the single aeCapex column, removing the old separate ones.
  const revisionTriggerCols = [
    colIndexes.model,
    colIndexes.aeCapex,
    colIndexes.aeSalesAskPrice,
    colIndexes.aeQuantity,
    colIndexes.aeTerm,
  ];

  ExecutionTimer.start("wasKeyFieldEdited_loop");
  for (const colIndex of revisionTriggerCols) {
    const originalValue = String(originalRow[colIndex - startCol] || "");
    const currentValue = String(currentRow[colIndex - startCol] || "");
    Log[sourceFile](
      `[${sourceFile} - wasKeyFieldEdited] CRAZY VERBOSE: Checking column ${colIndex}. Original: '${originalValue}', Current: '${currentValue}'.`
    );

    if (currentValue !== originalValue) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "wasKeyFieldEdited_fallbackPath_isChanged",
      });
      Log[sourceFile](
        `[${sourceFile} - wasKeyFieldEdited] CHANGE DETECTED: Column index ${colIndex} changed from '${originalValue}' to '${currentValue}'. Returning true.`
      );
      ExecutionTimer.end("wasKeyFieldEdited_loop");
      ExecutionTimer.end("wasKeyFieldEdited_total");
      return true;
    }
  }
  ExecutionTimer.end("wasKeyFieldEdited_loop");

  Log[sourceFile](
    `[${sourceFile} - wasKeyFieldEdited] Result: false. No key fields were changed.`
  );
  ExecutionTimer.end("wasKeyFieldEdited_total");
  return false;
}
