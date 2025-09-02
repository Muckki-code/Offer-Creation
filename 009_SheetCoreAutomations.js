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
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "getNumericValue_start",
  });
  if (typeof value === "number" && !isNaN(value)) {
    return value;
  }
  if (typeof value !== "string" || value.trim() === "") {
    return 0;
  }
  let numberString = value.replace(/[€$]/g, "").trim();
  numberString = numberString.replace(/,/g, ""); // Remove all commas.
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
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "_getStaticSheetValues_start",
  });
  const sourceFile = "SheetCoreAutomations_gs";
  if (_staticValuesCache) {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "_getStaticSheetValues_fromCache",
    });
    Log[sourceFile](`[_getStaticSheetValues] Returning values from cache.`);
    return _staticValuesCache;
  }
  Log[sourceFile](
    `[SheetCoreAutomations_gs - _getStaticSheetValues] Cache empty. Reading static values from sheet.`
  );
  ExecutionTimer.start("_getStaticSheetValues_read");

  // Define a single range that encompasses all the static cells we need.
  // This reads from I1 to L4.
  const staticCellsRange = sheet.getRange("I1:L4");
  const staticCellValues = staticCellsRange.getValues();

  ExecutionTimer.end("_getStaticSheetValues_read");
  ExecutionTimer.start("_getStaticSheetValues_parse");

  // Extract values from the 2D array based on their relative positions.
  // getRange("I1:L4") means:
  // I1 is at [0][0], J1 is [0][1], K1 is [0][2], L1 is [0][3]
  // I2 is at [1][0], J2 is [1][1], K2 is [1][2], L2 is [1][3]
  // etc.

  const languageValue = staticCellValues[0][0]; // I1
  const telekomDealValue = staticCellValues[0][3]; // L1

  const staticValues = {
    isTelekomDeal: String(telekomDealValue || "").toLowerCase() === "yes",
    docLanguage: String(languageValue || "german")
      .trim()
      .toLowerCase(),
  };

  ExecutionTimer.end("_getStaticSheetValues_parse");
  Log[sourceFile](
    `[SheetCoreAutomations_gs - _getStaticSheetValues] Caching and returning: ${JSON.stringify(
      staticValues
    )}`
  );
  _staticValuesCache = staticValues;
  return _staticValuesCache;
}

/**
 * OPTIMIZED: Finds the maximum existing index and returns the next available index.
 * This version performs a single, efficient read of only the necessary data.
 */
function getNextAvailableIndex(sheet) {
  ExecutionTimer.start("getNextAvailableIndex_total");
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "getNextAvailableIndex_start",
  });
  const indexColIndex = CONFIG.documentDeviceData.columnIndices.index;
  const startRow = CONFIG.approvalWorkflow.startDataRow;
  let maxIndex = 0;

  ExecutionTimer.start("getNextAvailableIndex_getValues");
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "getNextAvailableIndex_hasDataRows",
    });
    // Read the entire index column from the data start row to the end in one operation.
    const indexValues = sheet
      .getRange(startRow, indexColIndex, lastRow - startRow + 1, 1)
      .getValues();
    ExecutionTimer.end("getNextAvailableIndex_getValues");

    ExecutionTimer.start("getNextAvailableIndex_loop");
    // Find the max index from the in-memory array.
    maxIndex = indexValues.reduce((max, row) => {
      const value = parseFloat(row[0]);
      return !isNaN(value) && value > max ? value : max;
    }, 0);
    ExecutionTimer.end("getNextAvailableIndex_loop");
  } else {
    ExecutionTimer.end("getNextAvailableIndex_getValues");
  }

  Log.SheetCoreAutomations_gs(
    `[SheetCoreAutomations_gs - getNextAvailableIndex] Found max index ${maxIndex}. Next available will be ${maxIndex +
      1}.`
  );
  ExecutionTimer.end("getNextAvailableIndex_total");
  return maxIndex + 1;
}

/**
 * OPTIMIZED: Recalculates all data rows in the active sheet.
 * This version determines the next available index once from the in-memory array
 * to avoid repeated, slow calls back to the sheet.
 */
function recalculateAllRows(options = {}) {
  const sourceFile = "SheetCoreAutomations_gs";
  ExecutionTimer.start("recalculateAllRows_total");
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "recalculateAllRows_start",
  });
  _staticValuesCache = null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = getLastLastRow(sheet);
  if (lastRow < startRow) {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "recalculateAllRows_noData",
    });
    Log[sourceFile](
      `[${sourceFile} - recalculateAllRows] No data rows found (lastRow ${lastRow} < startRow ${startRow}). Exiting.`
    );
    ExecutionTimer.end("recalculateAllRows_total");
    return;
  }
  const numRows = lastRow - startRow + 1;
  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku; // FIXED: Added space
  const numCols = CONFIG.maxDataColumn - dataBlockStartCol + 1;
  ExecutionTimer.start("recalculateAllRows_readSheet");
  const allValuesBefore = sheet
    .getRange(startRow, dataBlockStartCol, numRows, numCols)
    .getValues();
  const allValuesAfter = JSON.parse(JSON.stringify(allValuesBefore));
  ExecutionTimer.end("recalculateAllRows_readSheet");
  Log[sourceFile](
    `[${sourceFile} - recalculateAllRows] Read ${numRows} rows from sheet.`
  );

  const staticValues = _getStaticSheetValues(sheet);
  const combinedIndexes = {
    ...CONFIG.approvalWorkflow.columnIndices,
    ...CONFIG.documentDeviceData.columnIndices,
  };
  let nextIndex = null; // Initialize to null

  ExecutionTimer.start("recalculateAllRows_mainLoop");
  for (let i = 0; i < numRows; i++) {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "recalculateAllRows_loop_iteration",
    });
    const currentRowNum = startRow + i;
    const inMemoryRowValues = allValuesAfter[i];
    const originalRowValuesForThisRow = allValuesBefore[i];
    Log[sourceFile](
      `[${sourceFile} - recalculateAllRows] Processing row ${currentRowNum}.`
    );

    const modelName =
      inMemoryRowValues[combinedIndexes.model - dataBlockStartCol];
    if (
      modelName &&
      !inMemoryRowValues[combinedIndexes.index - dataBlockStartCol]
    ) {
      Log.TestCoverage_gs({
        file: "SheetCoreAutomations.gs",
        coverage: "recalculateAllRows_assignIndex",
      });
      // If this is the first new row we've encountered, calculate the starting index ONCE.
      if (nextIndex === null) {
        const allCurrentIndices = allValuesAfter
          .map((row) =>
            parseFloat(row[combinedIndexes.index - dataBlockStartCol])
          )
          .filter((val) => !isNaN(val));
        const maxCurrentIndex =
          allCurrentIndices.length > 0 ? Math.max(...allCurrentIndices) : 0;
        nextIndex = maxCurrentIndex + 1;
      }
      inMemoryRowValues[
        combinedIndexes.index - dataBlockStartCol
      ] = nextIndex++; // Assign the index and increment for the next one
      Log[sourceFile](
        `[${sourceFile} - recalculateAllRows] Row ${currentRowNum}: Assigned new index ${
          inMemoryRowValues[combinedIndexes.index - dataBlockStartCol]
        }.`
      );
    }

    if (
      modelName &&
      !inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol]
    ) {
      Log.TestCoverage_gs({
        file: "SheetCoreAutomations.gs",
        coverage: "recalculateAllRows_assignApproverAction",
      });
      inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol] =
        "Choose Action";
      Log[sourceFile](
        `[${sourceFile} - recalculateAllRows] Row ${currentRowNum}: Assigned default Approver Action.`
      );
    }

    // --- THIS IS THE FIX ---
    updateCalculationsForRow(
      sheet,
      currentRowNum,
      inMemoryRowValues,
      combinedIndexes,
      CONFIG.approvalWorkflow,
      dataBlockStartCol
    );
    // --- END FIX ---

    const statusUpdateOptions = { forceRevisionOfFinalizedItems: true };
    const newStatus = updateStatusForRow(
      inMemoryRowValues,
      originalRowValuesForThisRow,
      staticValues.isTelekomDeal,
      statusUpdateOptions,
      dataBlockStartCol,
      combinedIndexes
    );
    if (newStatus !== null) {
      inMemoryRowValues[combinedIndexes.status - dataBlockStartCol] = newStatus;
    }
  }
  ExecutionTimer.end("recalculateAllRows_mainLoop");

  ExecutionTimer.start("recalculateAllRows_writeSheet");
  sheet
    .getRange(startRow, dataBlockStartCol, numRows, numCols)
    .setValues(allValuesAfter);
  ExecutionTimer.end("recalculateAllRows_writeSheet");
  Log[sourceFile](
    `[${sourceFile} - recalculateAllRows] Wrote all recalculated data back to the sheet.`
  );

  if (options.refreshUx) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "recalculateAllRows_refreshUx",
    });
    applyUxRules(true);
  }

  ExecutionTimer.end("recalculateAllRows_total");
}

// In SheetCoreAutomations.gs

// FILE: 009_SheetCoreAutomations.js

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
  ExecutionTimer.start("handleSheetAutomations_total");
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "handleSheetAutomations_start",
  });
  _staticValuesCache = null;
  const range = e.range;

  if (
    range.getRow() < CONFIG.approvalWorkflow.startDataRow &&
    range.getA1Notation() !== CONFIG.offerDetailsCells.telekomDeal
  ) {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "handleSheetAutomations_exit_notDataRow",
    });
    ExecutionTimer.end("handleSheetAutomations_total");
    return;
  }
  if (range.getA1Notation() === CONFIG.offerDetailsCells.telekomDeal) {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "handleSheetAutomations_recalcTriggered",
    });
    Log[sourceFile](
      `[${sourceFile} - handleSheetAutomations] Detected edit in global 'Telekom Deal' cell. Triggering recalculateAllRows.`
    );
    recalculateAllRows({
      oldValueIsTelekomDeal: (e.oldValue || "no").toLowerCase() === "yes",
    });
    ExecutionTimer.end("handleSheetAutomations_total");
    return;
  }

  ExecutionTimer.start("handleSheetAutomations_lock");
  const lock = LockService.getScriptLock();
  try {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "handleSheetAutomations_lock_wait",
    });
    lock.waitLock(30000);
  } catch (err) {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "handleSheetAutomations_lock_fail",
    });
    Log[sourceFile](
      `[${sourceFile} - handleSheetAutomations] WARNING: Could not obtain lock. Error: ${err.message}.`
    );
    SpreadsheetApp.getActive().toast(
      "The sheet is busy, please try your edit again in a moment.",
      "Busy",
      3
    );
    ExecutionTimer.end("handleSheetAutomations_lock");
    ExecutionTimer.end("handleSheetAutomations_total");
    return;
  }
  ExecutionTimer.end("handleSheetAutomations_lock");

  let needsRecalculation = false; // --- NEW: Initialize flag ---

  try {
    ExecutionTimer.start("handleSheetAutomations_flush");
    SpreadsheetApp.flush();
    Log[sourceFile](
      "[handleSheetAutomations] Lock acquired. Flushed all pending spreadsheet changes to prevent race conditions."
    );
    ExecutionTimer.end("handleSheetAutomations_flush");

    const sheet = range.getSheet();
    const combinedIndexes = {
      ...CONFIG.approvalWorkflow.columnIndices,
      ...CONFIG.documentDeviceData.columnIndices,
    };
    const editedRowStart = range.getRow();
    const numEditedRows = range.getNumRows();
    const editedColStart = range.getColumn();
    const isSingleCellEdit = numEditedRows === 1 && range.getNumColumns() === 1;
    const editedCol = editedColStart;
    const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
    const numColsInDataBlock = CONFIG.maxDataColumn - dataBlockStartCol + 1;
    const targetRange = sheet.getRange(
      editedRowStart,
      dataBlockStartCol,
      numEditedRows,
      numColsInDataBlock
    );
    const finalValuesToWrite = trueOriginalValuesForTest
      ? JSON.parse(JSON.stringify(trueOriginalValuesForTest))
      : targetRange.getValues();
    const originalValuesForComparison = JSON.parse(
      JSON.stringify(finalValuesToWrite)
    );
    Log[sourceFile](
      `[${sourceFile} - handleSheetAutomations] Data source: ${
        trueOriginalValuesForTest ? "Injected by Test" : "Read from Sheet"
      }.`
    );
    if (isSingleCellEdit) {
      Log.TestCoverage_gs({
        file: "SheetCoreAutomations.gs",
        coverage: "handleSheetAutomations_reconstruct_before_state",
      });
      const editedColIndexInArray = editedCol - dataBlockStartCol;
      if (
        editedColIndexInArray >= 0 &&
        editedColIndexInArray < originalValuesForComparison[0].length
      ) {
        originalValuesForComparison[0][editedColIndexInArray] = e.oldValue;
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Reconstructed 'before' state for single edit in col ${editedCol} with oldValue: '${e.oldValue}'`
        );
      }
    }
    const staticValues = _getStaticSheetValues(sheet);
    let nextIndex = null;

    ExecutionTimer.start("handleSheetAutomations_mainLoop");
    for (let i = 0; i < numEditedRows; i++) {
      Log.TestCoverage_gs({
        file: "SheetCoreAutomations.gs",
        coverage: "handleSheetAutomations_loop_start",
      });
      const currentRowNumInSheet = editedRowStart + i;
      const inMemoryRowValues = finalValuesToWrite[i];
      const originalRowValues = originalValuesForComparison[i];

      Log[sourceFile](
        `[${sourceFile} - handleSheetAutomations] ---- STARTING ROW ${currentRowNumInSheet} ----`
      );

      let wasBundleValidBeforeEdit = false;
      const oldBundleInfo = _getBundleInfoFromRange(range);
      if (oldBundleInfo) {
        wasBundleValidBeforeEdit = true;
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] CRAZY VERBOSE: Row ${currentRowNumInSheet} was part of valid bundle #${oldBundleInfo.bundleId} before this edit.`
        );
      }

      // --- BQ Data wipe logic ---
      let wipeBqData = false;
      if (isSingleCellEdit) {
        const skuChanged =
          String(
            inMemoryRowValues[combinedIndexes.sku - dataBlockStartCol] || ""
          ) !==
          String(
            originalRowValues[combinedIndexes.sku - dataBlockStartCol] || ""
          );
        const modelChanged =
          String(
            inMemoryRowValues[combinedIndexes.model - dataBlockStartCol] || ""
          ) !==
          String(
            originalRowValues[combinedIndexes.model - dataBlockStartCol] || ""
          );
        if (skuChanged !== modelChanged) {
          wipeBqData = true;
          Log[sourceFile](
            `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Flagging BQ data for wipe because SKU/Model changed independently (SKU changed: ${skuChanged}, Model changed: ${modelChanged}).`
          );
        }
      } else {
        const pasteStartCol = range.getColumn();
        const pasteEndCol = pasteStartCol + range.getNumColumns() - 1;
        const skuCol = combinedIndexes.sku;
        const modelCol = combinedIndexes.model;
        const skuWasPasted = skuCol >= pasteStartCol && skuCol <= pasteEndCol;
        const modelWasPasted =
          modelCol >= pasteStartCol && modelCol <= pasteEndCol;
        if (skuWasPasted !== modelWasPasted) {
          wipeBqData = true;
          Log[sourceFile](
            `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Flagging BQ data for wipe due to desynchronizing paste (SKU pasted: ${skuWasPasted}, Model pasted: ${modelWasPasted}).`
          );
        }
      }
      if (wipeBqData) {
        Log.TestCoverage_gs({
          file: "SheetCoreAutomations.gs",
          coverage: "handleSheetAutomations_wipeBqData",
        });
        inMemoryRowValues[combinedIndexes.epCapexRaw - dataBlockStartCol] = "";
        inMemoryRowValues[combinedIndexes.tkCapexRaw - dataBlockStartCol] = "";
        inMemoryRowValues[combinedIndexes.rentalTargetRaw - dataBlockStartCol] =
          "";
        inMemoryRowValues[combinedIndexes.rentalLimitRaw - dataBlockStartCol] =
          "";
      }
      // --- Sanitization logic ---
      ExecutionTimer.start("handleSheetAutomations_sanitizeBlock");
      if (!isSingleCellEdit) {
        Log.TestCoverage_gs({
          file: "SheetCoreAutomations.gs",
          coverage: "handleSheetAutomations_pasteSanitization",
        });
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Detected Paste operation. Sanitizing calculated/approval fields.`
        );
        const fieldsToWipe = [
          combinedIndexes.index,
          combinedIndexes.lrfPreview,
          combinedIndexes.contractValuePreview,
          combinedIndexes.status,
          combinedIndexes.financeApprovedPrice,
          combinedIndexes.approvedBy,
          combinedIndexes.approvalDate,
          combinedIndexes.approverComments,
          combinedIndexes.approverPriceProposal,
        ];
        fieldsToWipe.forEach((key) => {
          inMemoryRowValues[combinedIndexes[key] - dataBlockStartCol] = "";
        });
        inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol] =
          "Choose Action";
        originalRowValues[combinedIndexes.status - dataBlockStartCol] = "";
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Wiped original status for paste re-evaluation.`
        );
      } else {
        Log.TestCoverage_gs({
          file: "SheetCoreAutomations.gs",
          coverage: "handleSheetAutomations_singleCellSanitization",
        });
        if (CONFIG.protectedColumnIndices.includes(editedCol)) {
          Log.TestCoverage_gs({
            file: "SheetCoreAutomations.gs",
            coverage: "handleSheetAutomations_protectedColumnRevert",
          });
          inMemoryRowValues[editedCol - dataBlockStartCol] = e.oldValue;
          Log[sourceFile](
            `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: PROTECTED COLUMN. Reverted illegal edit on col ${editedCol} back to original value: '${e.oldValue}'.`
          );
        }
      }
      Log[sourceFile](
        `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Data AFTER sanitization: ${JSON.stringify(
          inMemoryRowValues
        )}`
      );
      ExecutionTimer.end("handleSheetAutomations_sanitizeBlock");

      // --- Downstream logic ---
      ExecutionTimer.start("handleSheetAutomations_downstreamLogic");
      updateCalculationsForRow(
        sheet,
        currentRowNumInSheet,
        inMemoryRowValues,
        combinedIndexes,
        CONFIG.approvalWorkflow,
        dataBlockStartCol
      );
      const approverActionValue =
        inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol];
      const isApprovalAction =
        isSingleCellEdit &&
        editedCol === combinedIndexes.approverAction &&
        approverActionValue &&
        approverActionValue !== "Choose Action";
      Log[sourceFile](
        `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Is approval action? ${isApprovalAction}.`
      );
      if (isApprovalAction) {
        Log.TestCoverage_gs({
          file: "SheetCoreAutomations.gs",
          coverage: "handleSheetAutomations_approvalActionPath",
        });
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Detected approver action. Calling processSingleApprovalAction.`
        );
        const mockEventForApproval = {
          ...e,
          value: approverActionValue,
          oldValue: originalRowValues[editedCol - dataBlockStartCol],
        };
        processSingleApprovalAction(
          sheet,
          currentRowNumInSheet,
          mockEventForApproval,
          inMemoryRowValues,
          combinedIndexes,
          originalRowValues,
          dataBlockStartCol
        );
      } else {
        Log.TestCoverage_gs({
          file: "SheetCoreAutomations.gs",
          coverage: "handleSheetAutomations_statusUpdatePath",
        });
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Not an approval action. Calling status determination logic.`
        );

        const initialStatus =
          originalRowValues[combinedIndexes.status - dataBlockStartCol] || "";
        const newStatus = updateStatusForRow(
          inMemoryRowValues,
          originalRowValues,
          staticValues.isTelekomDeal,
          {},
          dataBlockStartCol,
          combinedIndexes
        );

        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Status determination complete. Initial='${initialStatus}', New='${newStatus}'.`
        );

        if (newStatus !== initialStatus) {
          Log.TestCoverage_gs({
            file: "SheetCoreAutomations.gs",
            coverage: "handleSheetAutomations_statusDidChange",
          });
          Log[sourceFile](
            `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Status changed. Applying updates.`
          );
          logTableActivity({
            mainSheet: sheet,
            rowNum: currentRowNumInSheet,
            oldStatus: initialStatus,
            newStatus: newStatus,
            currentFullRowValues: inMemoryRowValues,
            originalFullRowValues: originalRowValues,
            startCol: dataBlockStartCol,
          });

          if (newStatus === null) {
            Log.TestCoverage_gs({
              file: "SheetCoreAutomations.gs",
              coverage: "handleSheetAutomations_clearStatusOnly",
            });
            inMemoryRowValues[combinedIndexes.status - dataBlockStartCol] = "";
          } else {
            inMemoryRowValues[
              combinedIndexes.status - dataBlockStartCol
            ] = newStatus;
            if (
              [
                CONFIG.approvalWorkflow.statusStrings.pending,
                CONFIG.approvalWorkflow.statusStrings.draft,
                CONFIG.approvalWorkflow.statusStrings.revisedByAE,
              ].includes(newStatus)
            ) {
              inMemoryRowValues[
                combinedIndexes.approverAction - dataBlockStartCol
              ] = "Choose Action";
            }
            const approvedStatuses = [
              CONFIG.approvalWorkflow.statusStrings.approvedOriginal,
              CONFIG.approvalWorkflow.statusStrings.approvedNew,
            ];
            if (
              approvedStatuses.includes(initialStatus) &&
              !approvedStatuses.includes(newStatus)
            ) {
              inMemoryRowValues[
                combinedIndexes.financeApprovedPrice - dataBlockStartCol
              ] = "";
              inMemoryRowValues[
                combinedIndexes.approvedBy - dataBlockStartCol
              ] = "";
              inMemoryRowValues[
                combinedIndexes.approvalDate - dataBlockStartCol
              ] = "";
            }
          }
        } else {
          Log.TestCoverage_gs({
            file: "SheetCoreAutomations.gs",
            coverage: "handleSheetAutomations_statusUnchanged",
          });
          Log[sourceFile](
            `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Status did not change. No updates needed.`
          );
        }
      }
      ExecutionTimer.end("handleSheetAutomations_downstreamLogic");

      // --- Initialization checks ---
      ExecutionTimer.start("handleSheetAutomations_initializationChecks");
      const modelName =
        inMemoryRowValues[combinedIndexes.model - dataBlockStartCol];
      if (
        modelName &&
        !inMemoryRowValues[combinedIndexes.index - dataBlockStartCol]
      ) {
        Log.TestCoverage_gs({
          file: "SheetCoreAutomations.gs",
          coverage: "handleSheetAutomations_setNewIndex",
        });
        if (nextIndex === null) {
          nextIndex = getNextAvailableIndex(sheet);
        }
        inMemoryRowValues[
          combinedIndexes.index - dataBlockStartCol
        ] = nextIndex++;
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Initialized with new Index ${
            inMemoryRowValues[combinedIndexes.index - dataBlockStartCol]
          }.`
        );
      }
      if (
        modelName &&
        !inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol]
      ) {
        Log.TestCoverage_gs({
          file: "SheetCoreAutomations.gs",
          coverage: "handleSheetAutomations_setDefaultAction",
        });
        inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol] =
          "Choose Action";
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Initialized with default Approver Action 'Choose Action'.`
        );
      }
      ExecutionTimer.end("handleSheetAutomations_initializationChecks");

      // --- METADATA-DRIVEN BUNDLE VALIDATION & FORMATTING ---
      if (
        CONFIG.featureFlags.enforceBundleIntegrityOnEdit &&
        isSingleCellEdit
      ) {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "handleSheetAutomations_bundleFlag_enabled",
        });
        ExecutionTimer.start("handleSheetAutomations_bundleMetadataLogic");

        const integrityCols = [
          combinedIndexes.bundleNumber,
          combinedIndexes.aeQuantity,
          combinedIndexes.aeTerm,
        ];
        if (integrityCols.includes(editedCol)) {
          Log.TestCoverage_gs({
            file: sourceFile,
            coverage: "handleSheetAutomations_bundleIntegrityEdit",
          });
          Log[sourceFile](
            `[${sourceFile} - handleSheetAutomations] CRAZY VERBOSE: Integrity column ${editedCol} was edited. Running bundle integrity check.`
          );

          const oldBundleNumber = wasBundleValidBeforeEdit
            ? oldBundleInfo.bundleId
            : null;
          const currentBundleNum =
            inMemoryRowValues[combinedIndexes.bundleNumber - dataBlockStartCol];

          if (
            wasBundleValidBeforeEdit &&
            String(currentBundleNum).trim() !== String(oldBundleNumber).trim()
          ) {
            Log.TestCoverage_gs({
              file: sourceFile,
              coverage: "handleSheetAutomations_bundleBrokenOrLeft",
            });
            Log[sourceFile](
              `[${sourceFile} - handleSheetAutomations] Bundle #${oldBundleNumber} has been broken or left by row ${currentRowNumInSheet}. Clearing old metadata and borders.`
            );
            _clearMetadataFromRowRange(
              sheet,
              oldBundleInfo.startRow,
              oldBundleInfo.endRow
            );
            const oldRange = sheet.getRange(
              oldBundleInfo.startRow,
              dataBlockStartCol,
              oldBundleInfo.endRow - oldBundleInfo.startRow + 1,
              numColsInDataBlock
            );
            oldRange.setBorder(null, null, null, null, null, null);

            const remainingBundleRows = _findRowsForBundle(
              sheet,
              oldBundleNumber
            );
            if (remainingBundleRows.length > 1) {
              const revalidated = validateBundle(
                sheet,
                remainingBundleRows[0],
                oldBundleNumber
              );
              if (revalidated.isValid) {
                const newBundleInfo = {
                  bundleId: String(oldBundleNumber),
                  startRow: revalidated.startRow,
                  endRow: revalidated.endRow,
                };
                _setMetadataForRowRange(sheet, newBundleInfo);
                const bundleRangeToFormat = sheet.getRange(
                  newBundleInfo.startRow,
                  dataBlockStartCol,
                  newBundleInfo.endRow - newBundleInfo.startRow + 1,
                  numColsInDataBlock
                );
                _clearAndApplyBundleBorder(bundleRangeToFormat);
              }
            }
          }

          if (currentBundleNum) {
            Log.TestCoverage_gs({
              file: sourceFile,
              coverage: "handleSheetAutomations_bundleValidationStart",
            });
            const validationResult = validateBundle(
              sheet,
              currentRowNumInSheet,
              currentBundleNum
            );

            if (validationResult.isValid) {
              Log.TestCoverage_gs({
                file: sourceFile,
                coverage: "handleSheetAutomations_bundleValid",
              });
              if (
                validationResult.startRow &&
                validationResult.endRow &&
                validationResult.startRow !== validationResult.endRow
              ) {
                Log.TestCoverage_gs({
                  file: sourceFile,
                  coverage: "handleSheetAutomations_bundleValid_isMultiRow",
                });
                const newBundleInfo = {
                  bundleId: String(currentBundleNum),
                  startRow: validationResult.startRow,
                  endRow: validationResult.endRow,
                };

                if (
                  JSON.stringify(oldBundleInfo) !==
                  JSON.stringify(newBundleInfo)
                ) {
                  Log.TestCoverage_gs({
                    file: sourceFile,
                    coverage: "handleSheetAutomations_bundleStructureChanged",
                  });
                  Log[sourceFile](
                    `[${sourceFile} - handleSheetAutomations] Bundle structure changed for #${currentBundleNum}. Updating metadata and borders.`
                  );
                  _setMetadataForRowRange(sheet, newBundleInfo);
                  const bundleRangeToFormat = sheet.getRange(
                    newBundleInfo.startRow,
                    dataBlockStartCol,
                    newBundleInfo.endRow - newBundleInfo.startRow + 1,
                    numColsInDataBlock
                  );
                  _clearAndApplyBundleBorder(bundleRangeToFormat);

                  // --- THIS IS THE FIX ---
                  if (wasBundleValidBeforeEdit === false) {
                    Log[sourceFile](
                      `[${sourceFile} - handleSheetAutomations] A bundle was just fixed or created. Setting flag for full recalculation.`
                    );
                    needsRecalculation = true;
                  }
                  // --- END FIX ---
                }
              }
            } else {
              Log.TestCoverage_gs({
                file: sourceFile,
                coverage:
                  "handleSheetAutomations_bundleInvalid_NON_DESTRUCTIVE",
              });
              Log[sourceFile](
                `[${sourceFile} - handleSheetAutomations] CRAZY VERBOSE: BUNDLE INVALID: Bundle #${currentBundleNum} failed validation. Reason: ${validationResult.errorCode}. Preserving user edit.`
              );

              // --- THIS IS THE FIX: Destructive multi-row write is REMOVED ---

              if (validationResult.errorCode === "MISMATCH") {
                Log.TestCoverage_gs({
                  file: sourceFile,
                  coverage: "handleSheetAutomations_showMismatchDialog",
                });
                const currentValues = {
                  term:
                    inMemoryRowValues[
                      combinedIndexes.aeTerm - dataBlockStartCol
                    ],
                  quantity:
                    inMemoryRowValues[
                      combinedIndexes.aeQuantity - dataBlockStartCol
                    ],
                };
                showBundleMismatchDialog(
                  currentRowNumInSheet,
                  currentBundleNum,
                  currentValues,
                  validationResult.expected
                );
                Log[sourceFile](
                  `[${sourceFile} - handleSheetAutomations] CRAZY VERBOSE: Called showBundleMismatchDialog for bundle #${currentBundleNum}.`
                );
              } else if (validationResult.errorCode === "GAP_DETECTED") {
                Log.TestCoverage_gs({
                  file: sourceFile,
                  coverage: "handleSheetAutomations_showGapDialog",
                });
                showBundleGapDialog(currentBundleNum);
                Log[sourceFile](
                  `[${sourceFile} - handleSheetAutomations] CRAZY VERBOSE: Called showBundleGapDialog for bundle #${currentBundleNum}.`
                );
              }

              inMemoryRowValues[combinedIndexes.status - dataBlockStartCol] =
                CONFIG.approvalWorkflow.statusStrings.draft;
            }
          }
        }
        ExecutionTimer.end("handleSheetAutomations_bundleMetadataLogic");
      }
      Log[sourceFile](
        `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: FINAL in-memory data before write: ${JSON.stringify(
          inMemoryRowValues
        )}`
      );
    }
    ExecutionTimer.end("handleSheetAutomations_mainLoop");

    if (trueOriginalValuesForTest !== null) {
      Log[sourceFile](
        `[${sourceFile} - handleSheetAutomations] Data was injected by test. Proceeding to write results to sheet.`
      );
    }
    targetRange.setValues(finalValuesToWrite);
    Log[sourceFile](
      `[${sourceFile} - handleSheetAutomations] Wrote final values back to range ${targetRange.getA1Notation()}.`
    );

    // --- NEW: Trigger recalculation AFTER writing changes ---
    if (needsRecalculation) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "handleSheetAutomations_triggeringRecalc",
      });
      Log[sourceFile](
        `[${sourceFile} - handleSheetAutomations] Executing deferred recalculation now.`
      );
      recalculateAllRows({ refreshUx: true });
    }
  } finally {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "handleSheetAutomations_lock_released",
    });
    lock.releaseLock();
    Log[sourceFile](`[${sourceFile} - handleSheetAutomations] Lock released.`);
  }
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "handleSheetAutomations_end",
  });
  ExecutionTimer.end("handleSheetAutomations_total");
}

/**
 * Calculates and updates BOTH the LRF and Contract Value for a specific row's in-memory data,
 * AND applies the correct number format to the corresponding cells.
 * REVISED: Now correctly uses the single 'aeCapex' column.
 */
function updateCalculationsForRow(
  sheet,
  rowNum,
  rowValues,
  colIndexes,
  approvalWorkflowConfig,
  dataBlockStartCol
) {
  const sourceFile = "SheetCoreAutomations_gs";
  ExecutionTimer.start("updateCalculationsForRow_total");
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "updateCalculationsForRow_start",
  });
  const statusStrings = approvalWorkflowConfig.statusStrings;
  let rentalPrice = 0;
  const status = rowValues[colIndexes.status - dataBlockStartCol];
  const approvedStatuses = [
    statusStrings.approvedOriginal,
    statusStrings.approvedNew,
  ];
  ExecutionTimer.start("updateCalculationsForRow_getPrice");
  if (approvedStatuses.includes(status)) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateCalculationsForRow_getPrice_approved",
    });
    rentalPrice = getNumericValue(
      rowValues[colIndexes.financeApprovedPrice - dataBlockStartCol]
    );
  } else {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateCalculationsForRow_getPrice_pending",
    });
    const approverPrice = getNumericValue(
      rowValues[colIndexes.approverPriceProposal - dataBlockStartCol]
    );
    const aeSalesAskPrice = getNumericValue(
      rowValues[colIndexes.aeSalesAskPrice - dataBlockStartCol]
    );
    rentalPrice = approverPrice > 0 ? approverPrice : aeSalesAskPrice;
  }
  ExecutionTimer.end("updateCalculationsForRow_getPrice");

  // --- THIS IS THE FIX ---
  const chosenCapex = getNumericValue(
    rowValues[colIndexes.aeCapex - dataBlockStartCol]
  );
  Log[sourceFile](
    `[${sourceFile} - updateCalculationsForRow] Row ${rowNum}: Inputs: rentalPrice=${rentalPrice}, chosenCapex=${chosenCapex}`
  );
  // --- END FIX ---

  ExecutionTimer.start("updateCalculationsForRow_calcLrf");
  const lrfCell = sheet.getRange(rowNum, colIndexes.lrfPreview);
  const contractValueCell = sheet.getRange(
    rowNum,
    colIndexes.contractValuePreview
  );
  const formats = CONFIG.numberFormats;

  if (rentalPrice === 0 && (!chosenCapex || chosenCapex === 0)) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateCalculationsForRow_clearValues",
    });
    rowValues[colIndexes.lrfPreview - dataBlockStartCol] = "";
    rowValues[colIndexes.contractValuePreview - dataBlockStartCol] = "";
    Log[sourceFile](
      `[${sourceFile} - updateCalculationsForRow] Row ${rowNum}: Clearing LRF and Contract Value due to zero price/capex.`
    );
  } else {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateCalculationsForRow_calculateValues",
    });
    if (!chosenCapex || chosenCapex <= 0) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "updateCalculationsForRow_missingCapex",
      });
      rowValues[colIndexes.lrfPreview - dataBlockStartCol] = "Missing\nCAPEX";
    } else {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "updateCalculationsForRow_calculateLrf",
      });
      const term = getNumericValue(
        rowValues[colIndexes.aeTerm - dataBlockStartCol]
      );
      if (chosenCapex > 0 && rentalPrice > 0 && term > 0) {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "updateCalculationsForRow_lrfSuccess",
        });
        rowValues[colIndexes.lrfPreview - dataBlockStartCol] =
          (rentalPrice * term) / chosenCapex;
      } else {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "updateCalculationsForRow_lrfZero",
        });
        rowValues[colIndexes.lrfPreview - dataBlockStartCol] = "";
      }
    }
    const quantity = getNumericValue(
      rowValues[colIndexes.aeQuantity - dataBlockStartCol]
    );
    const term = getNumericValue(
      rowValues[colIndexes.aeTerm - dataBlockStartCol]
    );
    if (rentalPrice > 0 && term > 0 && quantity > 0) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "updateCalculationsForRow_cvSuccess",
      });
      rowValues[colIndexes.contractValuePreview - dataBlockStartCol] =
        rentalPrice * term * quantity;
    } else {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "updateCalculationsForRow_cvZero",
      });
      rowValues[colIndexes.contractValuePreview - dataBlockStartCol] = "";
    }
  }
  ExecutionTimer.end("updateCalculationsForRow_calcLrf");

  ExecutionTimer.start("updateCalculationsForRow_setFormats");
  lrfCell.setNumberFormat(formats.percentage);
  contractValueCell.setNumberFormat(formats.currency);
  ExecutionTimer.end("updateCalculationsForRow_setFormats");

  Log[sourceFile](
    `[${sourceFile} - updateCalculationsForRow] Row ${rowNum}: Outputs: LRF=${
      rowValues[colIndexes.lrfPreview - dataBlockStartCol]
    }, ContractValue=${
      rowValues[colIndexes.contractValuePreview - dataBlockStartCol]
    }`
  );
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "updateCalculationsForRow_end",
  });
  ExecutionTimer.end("updateCalculationsForRow_total");
}

/**
 * Finds all row numbers for a given bundle number.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {string|number} bundleNumber The bundle ID to find.
 * @returns {Array<number>} An array of row numbers.
 */
function _findRowsForBundle(sheet, bundleNumber) {
  const sourceFile = "SheetCoreAutomations_gs";
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "_findRowsForBundle_start",
  });
  Log[sourceFile](
    `[${sourceFile} - _findRowsForBundle] Start: Searching for all rows of bundle #${bundleNumber}.`
  );

  const dataStartRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = sheet.getLastRow();
  const bundleNumCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;

  if (lastRow < dataStartRow) {
    return [];
  }

  const bundleColumnValues = sheet
    .getRange(dataStartRow, bundleNumCol, lastRow - dataStartRow + 1, 1)
    .getValues();
  const matchingRows = [];

  bundleColumnValues.forEach((val, i) => {
    if (String(val[0]).trim() == String(bundleNumber)) {
      matchingRows.push(dataStartRow + i);
    }
  });

  Log[sourceFile](
    `[${sourceFile} - _findRowsForBundle] End: Found ${
      matchingRows.length
    } rows for bundle #${bundleNumber}: [${matchingRows.join(", ")}].`
  );
  return matchingRows;
}
