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

  const brokenBundleErrors = findAllBundleErrors();
  const brokenBundleIds = new Set(
    brokenBundleErrors.map((e) => String(e.bundleNumber))
  );
  Log[sourceFile](
    `[${sourceFile} - recalculateAllRows] Found ${brokenBundleIds.size} broken bundles to enforce 'Draft' status on.`
  );

  const numRows = lastRow - startRow + 1;
  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
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
  const statusStrings = CONFIG.approvalWorkflow.statusStrings;
  let nextIndex = null;

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
      ] = nextIndex++;
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

    updateCalculationsForRow(
      sheet,
      currentRowNum,
      inMemoryRowValues,
      staticValues.isTelekomDeal,
      combinedIndexes,
      CONFIG.approvalWorkflow,
      dataBlockStartCol
    );

    const statusUpdateOptions = {
      forceRevisionOfFinalizedItems: true,
      brokenBundleIds: brokenBundleIds,
    };
    const initialStatus =
      originalRowValuesForThisRow[combinedIndexes.status - dataBlockStartCol] ||
      "";

    // --- THIS IS THE FIX for the STATUS REGRESSION ---
    const newStatus = updateStatusForRow(
      inMemoryRowValues,
      originalRowValuesForThisRow,
      staticValues.isTelekomDeal,
      statusUpdateOptions,
      dataBlockStartCol,
      combinedIndexes
    );

    if (newStatus !== initialStatus) {
      if (newStatus === null) {
        inMemoryRowValues[combinedIndexes.status - dataBlockStartCol] = "";
      } else {
        inMemoryRowValues[
          combinedIndexes.status - dataBlockStartCol
        ] = newStatus;
        if (
          [
            statusStrings.pending,
            statusStrings.draft,
            statusStrings.revisedByAE,
          ].includes(newStatus)
        ) {
          inMemoryRowValues[
            combinedIndexes.approverAction - dataBlockStartCol
          ] = "Choose Action";
        }
        const approvedStatuses = [
          statusStrings.approvedOriginal,
          statusStrings.approvedNew,
          statusStrings.rejected,
        ];
        if (
          approvedStatuses.includes(initialStatus) &&
          !approvedStatuses.includes(newStatus)
        ) {
          inMemoryRowValues[
            combinedIndexes.financeApprovedPrice - dataBlockStartCol
          ] = "";
          inMemoryRowValues[combinedIndexes.approvedBy - dataBlockStartCol] =
            "";
          inMemoryRowValues[combinedIndexes.approvalDate - dataBlockStartCol] =
            "";
        }
      }
    }
    // --- END FIX ---
  }
  ExecutionTimer.end("recalculateAllRows_mainLoop");

  ExecutionTimer.start("recalculateAllRows_writeSheet");
  sheet
    .getRange(startRow, dataBlockStartCol, numRows, numCols)
    .setValues(allValuesAfter);
  ExecutionTimer.end("recalculateAllRows_writeSheet");

  // --- THIS IS THE FIX for the BORDER REFRESH ---
  if (options.refreshUx) {
    applyUxRules(true);
  }
  // --- END FIX ---

  Log[sourceFile](
    `[${sourceFile} - recalculateAllRows] Wrote all recalculated data. Finished.`
  );
  ExecutionTimer.end("recalculateAllRows_total");
}

/**
 * Main onEdit trigger handler.
 * FINAL MERGED VERSION: Restores all critical safety and sanitization logic
 * from the robust older version, while using the "surgical write" mechanism to
 * prevent data deletion and correctly handling bundle metadata scanning.
 */
function handleSheetAutomations(e, trueOriginalValuesForTest = null) {
  const sourceFile = "SheetCoreAutomations_gs";
  ExecutionTimer.start("handleSheetAutomations_total");
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "handleSheetAutomations_start",
  });
  _staticValuesCache = null;
  const range = e.range;

  if (
    range.getRow() < CONFIG.approvalWorkflow.startDataRow &&
    range.getA1Notation() !== CONFIG.offerDetailsCells.telekomDeal
  ) {
    ExecutionTimer.end("handleSheetAutomations_total");
    return;
  }
  if (range.getA1Notation() === CONFIG.offerDetailsCells.telekomDeal) {
    Log[sourceFile](
      "[handleSheetAutomations] Telekom Deal cell changed. Triggering full recalculation."
    );
    recalculateAllRows({ refreshUx: true });
    ExecutionTimer.end("handleSheetAutomations_total");
    return;
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    SpreadsheetApp.getActive().toast(
      "The sheet is busy, please try your edit again in a moment.",
      "Busy",
      3
    );
    ExecutionTimer.end("handleSheetAutomations_total");
    return;
  }

  try {
    const sheet = range.getSheet();
    const c = {
      ...CONFIG.approvalWorkflow.columnIndices,
      ...CONFIG.documentDeviceData.columnIndices,
    };
    const editedRowStart = range.getRow();
    const numEditedRows = range.getNumRows();
    const editedColStart = range.getColumn();
    const isSingleCellEdit = numEditedRows === 1 && range.getNumColumns() === 1;
    const dataBlockStartCol = c.sku;
    const numColsInDataBlock = CONFIG.maxDataColumn - dataBlockStartCol + 1;
    Log[sourceFile](
      `[handleSheetAutomations] START: Edit detected at ${range.getA1Notation()}. isSingleCellEdit=${isSingleCellEdit}`
    );

    // 1. Capture Pre-Edit State for logical comparisons.
    ExecutionTimer.start("handleSheetAutomations_read_before");
    const originalSheetValues = trueOriginalValuesForTest
      ? trueOriginalValuesForTest
      : sheet
          .getRange(
            editedRowStart,
            dataBlockStartCol,
            numEditedRows,
            numColsInDataBlock
          )
          .getValues();
    ExecutionTimer.end("handleSheetAutomations_read_before");
    Log[sourceFile](
      `[handleSheetAutomations] CRAZY VERBOSE: Captured 'before' state for ${numEditedRows} row(s).`
    );

    // Ensure the user's edit is fully written to the sheet before we read it back.
    SpreadsheetApp.flush();

    // 2. Capture Post-Edit State. This is the user's true intent and our baseline for processing.
    ExecutionTimer.start("handleSheetAutomations_read_after");
    const postEditValues = sheet
      .getRange(
        editedRowStart,
        dataBlockStartCol,
        numEditedRows,
        numColsInDataBlock
      )
      .getValues();
    ExecutionTimer.end("handleSheetAutomations_read_after");
    const valuesToProcess = JSON.parse(JSON.stringify(postEditValues));
    Log[sourceFile](
      `[handleSheetAutomations] CRAZY VERBOSE: Captured 'post-edit' state. Beginning processing loop.`
    );

    const staticValues = _getStaticSheetValues(sheet);
    let nextIndex = null;

    // 3. Main processing loop: Apply script logic to the post-edit data.
    ExecutionTimer.start("handleSheetAutomations_main_loop");
    for (let i = 0; i < numEditedRows; i++) {
      const currentRowNum = editedRowStart + i;
      const inMemoryRow = valuesToProcess[i];
      const originalRowForLogic = originalSheetValues[i];
      Log[sourceFile](
        `[handleSheetAutomations] Processing row ${currentRowNum}...`
      );

      let wipeBqData = false;
      if (isSingleCellEdit) {
        if (CONFIG.protectedColumnIndices.includes(editedColStart)) {
          Log[sourceFile](
            `[handleSheetAutomations] Row ${currentRowNum}: Edit was on protected column ${editedColStart}. Reverting value.`
          );
          inMemoryRow[editedColStart - dataBlockStartCol] = e.oldValue;
        }
        const skuChanged =
          String(inMemoryRow[c.sku - dataBlockStartCol] || "") !==
          String(originalRowForLogic[c.sku - dataBlockStartCol] || "");
        const modelChanged =
          String(inMemoryRow[c.model - dataBlockStartCol] || "") !==
          String(originalRowForLogic[c.model - dataBlockStartCol] || "");
        if ((skuChanged && !modelChanged) || (modelChanged && !skuChanged)) {
          Log[sourceFile](
            `[handleSheetAutomations] Row ${currentRowNum}: SKU/Model desynchronized. Flagging BQ data for wipe.`
          );
          wipeBqData = true;
        }
      } else {
        // This is a paste
        Log[sourceFile](
          `[handleSheetAutomations] Row ${currentRowNum}: Paste detected. Wiping script-managed fields.`
        );
        const fieldsToWipe = [
          c.index,
          c.lrfPreview,
          c.contractValuePreview,
          c.status,
          c.financeApprovedPrice,
          c.approvedBy,
          c.approvalDate,
          c.approverComments,
          c.approverPriceProposal,
        ];
        fieldsToWipe.forEach((col) => {
          inMemoryRow[col - dataBlockStartCol] = "";
        });
        inMemoryRow[c.approverAction - dataBlockStartCol] = "Choose Action";
        const pasteStartCol = range.getColumn();
        const pasteEndCol = pasteStartCol + range.getNumColumns() - 1;
        if (
          (c.sku >= pasteStartCol && c.sku <= pasteEndCol) !==
          (c.model >= pasteStartCol && c.model <= pasteEndCol)
        ) {
          Log[sourceFile](
            `[handleSheetAutomations] Row ${currentRowNum}: Paste desynchronized SKU and Model. Flagging BQ data for wipe.`
          );
          wipeBqData = true;
        }
      }
      if (wipeBqData) {
        Log[sourceFile](
          `[handleSheetAutomations] Row ${currentRowNum}: Executing BQ data wipe.`
        );
        inMemoryRow[c.epCapexRaw - dataBlockStartCol] = "";
        inMemoryRow[c.tkCapexRaw - dataBlockStartCol] = "";
        inMemoryRow[c.rentalTargetRaw - dataBlockStartCol] = "";
        inMemoryRow[c.rentalLimitRaw - dataBlockStartCol] = "";
      }

      const modelName = inMemoryRow[c.model - dataBlockStartCol];
      if (modelName && !inMemoryRow[c.index - dataBlockStartCol]) {
        if (nextIndex === null) nextIndex = getNextAvailableIndex(sheet);
        inMemoryRow[c.index - dataBlockStartCol] = nextIndex++;
        Log[sourceFile](
          `[handleSheetAutomations] Row ${currentRowNum}: Assigned new index ${
            inMemoryRow[c.index - dataBlockStartCol]
          }.`
        );
      }
      if (modelName && !inMemoryRow[c.approverAction - dataBlockStartCol]) {
        inMemoryRow[c.approverAction - dataBlockStartCol] = "Choose Action";
        Log[sourceFile](
          `[handleSheetAutomations] Row ${currentRowNum}: Assigned default 'Choose Action'.`
        );
      }

      updateCalculationsForRow(
        sheet,
        currentRowNum,
        inMemoryRow,
        staticValues.isTelekomDeal,
        c,
        CONFIG.approvalWorkflow,
        dataBlockStartCol
      );

      const isApprovalAction =
        isSingleCellEdit &&
        editedColStart === c.approverAction &&
        e.value &&
        e.value !== "Choose Action";
      if (isApprovalAction) {
        Log[sourceFile](
          `[handleSheetAutomations] Row ${currentRowNum}: Approval action detected. Passing to processSingleApprovalAction.`
        );
        processSingleApprovalAction(
          sheet,
          currentRowNum,
          e,
          inMemoryRow,
          c,
          originalRowForLogic,
          dataBlockStartCol
        );
      } else {
        const initialStatus =
          originalRowForLogic[c.status - dataBlockStartCol] || "";
        const newStatus = updateStatusForRow(
          inMemoryRow,
          originalRowForLogic,
          staticValues.isTelekomDeal,
          {},
          dataBlockStartCol,
          c
        );
        Log[sourceFile](
          `[handleSheetAutomations] Row ${currentRowNum}: Status logic determined new status should be '${newStatus}' (was '${initialStatus}').`
        );

        if (newStatus !== initialStatus) {
          logTableActivity({
            mainSheet: sheet,
            rowNum: currentRowNum,
            oldStatus: initialStatus,
            newStatus: newStatus,
            currentFullRowValues: inMemoryRow,
            originalFullRowValues: originalRowForLogic,
            startCol: dataBlockStartCol,
          });
          if (newStatus === null) {
            inMemoryRow[c.status - dataBlockStartCol] = "";
          } else {
            inMemoryRow[c.status - dataBlockStartCol] = newStatus;
            if (
              [
                CONFIG.approvalWorkflow.statusStrings.pending,
                CONFIG.approvalWorkflow.statusStrings.draft,
                CONFIG.approvalWorkflow.statusStrings.revisedByAE,
              ].includes(newStatus)
            ) {
              inMemoryRow[c.approverAction - dataBlockStartCol] =
                "Choose Action";
            }
            const finalizedStatuses = [
              CONFIG.approvalWorkflow.statusStrings.approvedOriginal,
              CONFIG.approvalWorkflow.statusStrings.approvedNew,
              CONFIG.approvalWorkflow.statusStrings.rejected,
            ];
            if (
              finalizedStatuses.includes(initialStatus) &&
              !finalizedStatuses.includes(newStatus)
            ) {
              Log[sourceFile](
                `[handleSheetAutomations] Row ${currentRowNum}: Status reverted from finalized state. Wiping approval fields.`
              );
              inMemoryRow[c.financeApprovedPrice - dataBlockStartCol] = "";
              inMemoryRow[c.approvedBy - dataBlockStartCol] = "";
              inMemoryRow[c.approvalDate - dataBlockStartCol] = "";
            }
          }
        }
      }
    }
    ExecutionTimer.end("handleSheetAutomations_main_loop");

    // 4. Write the processed data back to the sheet.
    ExecutionTimer.start("handleSheetAutomations_write_main");
    sheet
      .getRange(
        editedRowStart,
        dataBlockStartCol,
        numEditedRows,
        numColsInDataBlock
      )
      .setValues(valuesToProcess);
    ExecutionTimer.end("handleSheetAutomations_write_main");
    Log[sourceFile](`[handleSheetAutomations] Batch write complete.`);

    // 5. Post-write bundle and UX logic
    const integrityCols = [c.bundleNumber, c.aeQuantity, c.aeTerm];
    if (isSingleCellEdit && integrityCols.includes(editedColStart)) {
      Log[sourceFile](
        `[handleSheetAutomations] Bundle integrity column ${editedColStart} was edited. Performing bundle validation.`
      );
      const bundleNumber = String(e.value || e.oldValue || "").trim();
      if (bundleNumber) {
        const validationResult = validateBundle(
          sheet,
          editedRowStart,
          bundleNumber
        );
        Log[sourceFile](
          `[handleSheetAutomations] Validation result for bundle #${bundleNumber}: isValid=${validationResult.isValid}, errorCode=${validationResult.errorCode}`
        );
        if (!validationResult.isValid) {
          Log[sourceFile](
            `[handleSheetAutomations] Bundle #${bundleNumber} is INVALID. Triggering UI and forcing Draft status.`
          );
          if (validationResult.errorCode === "GAP_DETECTED") {
            showBundleGapDialog(bundleNumber);
          } else if (validationResult.errorCode === "MISMATCH") {
            const currentValues = {
              term: range.offset(0, c.aeTerm - editedColStart).getValue(),
              quantity: range
                .offset(0, c.aeQuantity - editedColStart)
                .getValue(),
            };
            showBundleMismatchDialog(
              editedRowStart,
              bundleNumber,
              currentValues,
              validationResult.expected
            );
          }
          const bundleRangeInfo = _findBundleRange(sheet, bundleNumber);
          if (bundleRangeInfo.startRow && bundleRangeInfo.endRow) {
            const numBundleRows =
              bundleRangeInfo.endRow - bundleRangeInfo.startRow + 1;
            Log[sourceFile](
              `[handleSheetAutomations] Forcing 'Draft' status on rows ${bundleRangeInfo.startRow}-${bundleRangeInfo.endRow}.`
            );
            const bundleStatusRange = sheet.getRange(
              bundleRangeInfo.startRow,
              c.status,
              numBundleRows,
              1
            );
            const statusesToSet = Array(numBundleRows).fill([
              CONFIG.approvalWorkflow.statusStrings.draft,
            ]);
            bundleStatusRange.setValues(statusesToSet);
          }
        }
      }
      Log[sourceFile](
        `[handleSheetAutomations] Running full metadata scan and UI refresh due to bundle edit.`
      );
      scanAndSetAllBundleMetadata();
      applyUxRules(true);
    }
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
      Log[sourceFile](`[handleSheetAutomations] Script lock released.`);
    }
  }
  ExecutionTimer.end("handleSheetAutomations_total");
}

// In SheetCoreAutomations.gs

/**
 * Calculates and updates BOTH the LRF and Contract Value for a specific row's in-memory data,
 * AND applies the correct number format to the corresponding cells.
 * REFACTORED: Now uses the single aeCapex column, removing the old Telekom Deal logic.
 */
function updateCalculationsForRow(
  sheet,
  rowNum,
  rowValues,
  isTelekomDeal, // Kept for signature compatibility, but no longer used in logic
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
    rentalPrice = getNumericValue(
      rowValues[colIndexes.financeApprovedPrice - dataBlockStartCol]
    );
    Log[sourceFile](
      `[${sourceFile} - updateCalculationsForRow] CRAZY VERBOSE: Row ${rowNum}: Status is finalized ('${status}'). Using Finance Approved Price: ${rentalPrice}`
    );
  } else {
    const approverPrice = getNumericValue(
      rowValues[colIndexes.approverPriceProposal - dataBlockStartCol]
    );
    const aeSalesAskPrice = getNumericValue(
      rowValues[colIndexes.aeSalesAskPrice - dataBlockStartCol]
    );
    rentalPrice = approverPrice > 0 ? approverPrice : aeSalesAskPrice;
    Log[sourceFile](
      `[${sourceFile} - updateCalculationsForRow] CRAZY VERBOSE: Row ${rowNum}: Status is pending. Using Approver Price (${approverPrice}) or AE Ask Price (${aeSalesAskPrice}). Final Price: ${rentalPrice}`
    );
  }
  ExecutionTimer.end("updateCalculationsForRow_getPrice");

  const chosenCapex = getNumericValue(
    rowValues[colIndexes.aeCapex - dataBlockStartCol]
  );
  Log[sourceFile](
    `[${sourceFile} - updateCalculationsForRow] Row ${rowNum}: Inputs: rentalPrice=${rentalPrice}, chosenCapex=${chosenCapex}`
  );

  ExecutionTimer.start("updateCalculationsForRow_calcLrf");
  const lrfCell = sheet.getRange(rowNum, colIndexes.lrfPreview);
  const contractValueCell = sheet.getRange(
    rowNum,
    colIndexes.contractValuePreview
  );
  const formats = CONFIG.numberFormats;

  if (rentalPrice === 0 || chosenCapex <= 0) {
    rowValues[colIndexes.lrfPreview - dataBlockStartCol] = "";
    rowValues[colIndexes.contractValuePreview - dataBlockStartCol] = "";
    Log[sourceFile](
      `[${sourceFile} - updateCalculationsForRow] Row ${rowNum}: Clearing LRF and Contract Value due to zero/invalid price or capex.`
    );
  } else {
    const term = getNumericValue(
      rowValues[colIndexes.aeTerm - dataBlockStartCol]
    );
    if (term > 0) {
      rowValues[colIndexes.lrfPreview - dataBlockStartCol] =
        (rentalPrice * term) / chosenCapex;
    } else {
      rowValues[colIndexes.lrfPreview - dataBlockStartCol] = "";
    }

    const quantity = getNumericValue(
      rowValues[colIndexes.aeQuantity - dataBlockStartCol]
    );
    if (term > 0 && quantity > 0) {
      rowValues[colIndexes.contractValuePreview - dataBlockStartCol] =
        rentalPrice * term * quantity;
    } else {
      rowValues[colIndexes.contractValuePreview - dataBlockStartCol] = "";
    }
  }
  ExecutionTimer.end("updateCalculationsForRow_calcLrf");

  // --- NEW: Apply Number Formatting Directly ---
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
  ExecutionTimer.end("updateCalculationsForRow_total");
}
