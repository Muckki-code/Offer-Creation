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
/**
 * SPRINT 2 PERFORMANCE REFACTOR: Caching helper function.
 * OPTIMIZED: This version reads all required header/config cells in a single batch
 * operation (.getValues()) to significantly reduce API call overhead.
 */
function _getStaticSheetValues(sheet) {
  const sourceFile = "SheetCoreAutomations_gs";
  ExecutionTimer.start("_getStaticSheetValues_total");
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "_getStaticSheetValues_start",
  });

  // --- LAYER 1: Check Execution Cache (Fastest) ---
  if (_staticValuesCache) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "_getStaticSheetValues_fromExecutionCache",
    });
    ExecutionTimer.end("_getStaticSheetValues_total");
    return _staticValuesCache;
  }

  // --- LAYER 2: Check Script Cache (Fast) ---
  const cache = CacheService.getScriptCache();
  const cacheKey = "staticSheetValues";
  const cachedJSON = cache.get(cacheKey);

  if (cachedJSON) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "_getStaticSheetValues_fromScriptCache",
    });
    const staticValues = JSON.parse(cachedJSON);
    _staticValuesCache = staticValues; // Populate execution cache
    ExecutionTimer.end("_getStaticSheetValues_total");
    return staticValues;
  }

  // --- LAYER 3: Read from Sheet (Slowest) ---
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "_getStaticSheetValues_fromSheet",
  });
  ExecutionTimer.start("_getStaticSheetValues_read");
  const staticCellsRange = sheet.getRange("F1:O4"); // EXPANDED RANGE
  const staticCellValues = staticCellsRange.getValues();
  ExecutionTimer.end("_getStaticSheetValues_read");

  ExecutionTimer.start("_getStaticSheetValues_parse");
  // Sanitize ONLY the fields that are intended to be strings.
  const staticValues = {
    customerCompany: (staticCellValues[0][1] || "").toString().trim(),
    language: (staticCellValues[0][3] || "german")
      .toString()
      .trim()
      .toLowerCase(),
    telekomDeal: (staticCellValues[0][6] || "").toString().trim(),
    isTelekomDeal:
      (staticCellValues[0][6] || "").toString().toLowerCase() === "yes",
    approver: (staticCellValues[0][9] || "").toString().trim(),
    customerContactName: (staticCellValues[1][1] || "").toString().trim(),
    offerType: (staticCellValues[1][3] || "")
      .toString()
      .trim()
      .toLowerCase(),
    yourName: (staticCellValues[1][6] || "").toString().trim(),
    companyAddress: (staticCellValues[2][1] || "").toString().trim(),
    contractTerm: staticCellValues[2][3] || "", // Keep raw, formatter handles it
    yourPosition: (staticCellValues[2][6] || "").toString().trim(),
    specialAgreements: (staticCellValues[3][1] || "").toString().trim(),
    offerValidUntil: staticCellValues[3][3], // CRITICAL: Preserve as Date object
    documentName: (staticCellValues[3][6] || "").toString().trim(),
  };
  ExecutionTimer.end("_getStaticSheetValues_parse");

  cache.put(cacheKey, JSON.stringify(staticValues), 21600);
  _staticValuesCache = staticValues;

  Log[sourceFile](
    `[${sourceFile}] Caching and returning: ${JSON.stringify(staticValues)}`
  );
  ExecutionTimer.end("_getStaticSheetValues_total");
  return staticValues;
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
    ExecutionTimer.end("recalculateAllRows_total");
    return;
  }

  const numRows = lastRow - startRow + 1;
  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - dataBlockStartCol + 1;

  ExecutionTimer.start("recalculateAllRows_readSheet");
  const allValuesBefore = sheet
    .getRange(startRow, dataBlockStartCol, numRows, numCols)
    .getValues();
  // Create a deep copy for modifications
  const allValuesAfter = JSON.parse(JSON.stringify(allValuesBefore));
  ExecutionTimer.end("recalculateAllRows_readSheet");

  const staticValues = _getStaticSheetValues(sheet);
  const combinedIndexes = {
    ...CONFIG.approvalWorkflow.columnIndices,
    ...CONFIG.documentDeviceData.columnIndices,
  };
  let nextIndex = null;

  ExecutionTimer.start("recalculateAllRows_mainLoop");
  for (let i = 0; i < numRows; i++) {
    const currentRowNum = startRow + i;
    const inMemoryRowValues = allValuesAfter[i];
    const originalRowValuesForThisRow = allValuesBefore[i];
    let changesToWrite = {};

    // --- LOGIC MOVED FROM handleSheetAutomations - adapted for recalculation ---
    const modelName =
      inMemoryRowValues[combinedIndexes.model - dataBlockStartCol];
    if (
      modelName &&
      !inMemoryRowValues[combinedIndexes.index - dataBlockStartCol]
    ) {
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
      changesToWrite[combinedIndexes.index] = nextIndex++;
    }

    if (
      modelName &&
      !inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol]
    ) {
      changesToWrite[combinedIndexes.approverAction] = "Choose Action";
    }

    // --- THIS IS THE FIX ---
    // Correctly call the pure function and apply its results to the changes object
    const calculatedValues = updateCalculationsForRow(
      inMemoryRowValues,
      combinedIndexes,
      CONFIG.approvalWorkflow,
      dataBlockStartCol
    );
    changesToWrite[combinedIndexes.lrfPreview] = calculatedValues.lrfPreview;
    changesToWrite[combinedIndexes.contractValuePreview] =
      calculatedValues.contractValuePreview;
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

    // Also add status to changesToWrite
    if (
      newStatus !==
      (inMemoryRowValues[combinedIndexes.status - dataBlockStartCol] || "")
    ) {
      changesToWrite[combinedIndexes.status] =
        newStatus === null ? "" : newStatus;
    }

    // Apply all calculated changes for this row
    for (const colKey in changesToWrite) {
      inMemoryRowValues[colKey - dataBlockStartCol] = changesToWrite[colKey];
    }
  }
  ExecutionTimer.end("recalculateAllRows_mainLoop");

  ExecutionTimer.start("recalculateAllRows_writeSheet");
  sheet
    .getRange(startRow, dataBlockStartCol, numRows, numCols)
    .setValues(allValuesAfter);
  ExecutionTimer.end("recalculateAllRows_writeSheet");

  if (options && options.refreshUx) {
    applyUxRules(true);
  }

  ExecutionTimer.end("recalculateAllRows_total");
}

/**
 * Checks if a given range overlaps with a container range.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Range} editedRange The range that was edited.
 * @param {GoogleAppsScript.Spreadsheet.Range} containerRange The range to check against.
 * @returns {boolean} True if the ranges overlap, false otherwise.
 */
function _isRangeWithin(editedRange, containerRange) {
  const editedStartRow = editedRange.getRow();
  const editedEndRow = editedRange.getLastRow();
  const editedStartCol = editedRange.getColumn();
  const editedEndCol = editedRange.getLastColumn();

  const containerStartRow = containerRange.getRow();
  const containerEndRow = containerRange.getLastRow();
  const containerStartCol = containerRange.getColumn();
  const containerEndCol = containerRange.getLastColumn();

  // Check for no overlap. If any of these are true, the ranges do not intersect.
  const noOverlap =
    editedEndRow < containerStartRow ||
    editedStartRow > containerEndRow ||
    editedEndCol < containerStartCol ||
    editedStartCol > containerEndCol;

  return !noOverlap; // If there is "no 'noOverlap'", then there IS an overlap.
}

// In SheetCoreAutomations.gs

// FILE: 009_SheetCoreAutomations.js

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

/**
 * --- NEW ---
 * Applies a set of changes to a given row in the most performant way possible
 * by grouping contiguous cell updates into a single .setValues() call.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to write to.
 * @param {number} rowNum The 1-based row number to apply changes to.
 * @param {Object.<number, any>} changes An object where keys are 1-based column indices and values are the new cell values.
 */
function _applyTargetedWrites(sheet, rowNum, changes) {
  const sourceFile = "SheetCoreAutomations_gs";
  ExecutionTimer.start("_applyTargetedWrites_total");
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "_applyTargetedWrites_start",
  });

  const columnKeys = Object.keys(changes).map(Number);
  if (columnKeys.length === 0) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "_applyTargetedWrites_noChanges",
    });
    ExecutionTimer.end("_applyTargetedWrites_total");
    return;
  }

  columnKeys.sort((a, b) => a - b);
  Log[sourceFile](
    `[${sourceFile} - _applyTargetedWrites] Row ${rowNum}: Applying changes to columns: ${columnKeys.join(
      ", "
    )}.`
  );

  ExecutionTimer.start("_applyTargetedWrites_grouping");
  let startCol = columnKeys[0];
  let currentGroup = [changes[startCol]];

  for (let i = 1; i < columnKeys.length; i++) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "_applyTargetedWrites_loop_iteration",
    });
    const currentCol = columnKeys[i];
    const prevCol = columnKeys[i - 1];

    if (currentCol === prevCol + 1) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "_applyTargetedWrites_isContiguous",
      });
      currentGroup.push(changes[currentCol]);
    } else {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "_applyTargetedWrites_isNotContiguous",
      });
      // Write the previous group
      const range = sheet.getRange(rowNum, startCol, 1, currentGroup.length);
      Log[sourceFile](
        `[${sourceFile} - _applyTargetedWrites] Row ${rowNum}: Writing group to range ${range.getA1Notation()}.`
      );
      if (currentGroup.length === 1) {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "_applyTargetedWrites_writeSingle",
        });
        range.setValue(currentGroup[0]);
      } else {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "_applyTargetedWrites_writeBatch",
        });
        range.setValues([currentGroup]);
      }
      // Start a new group
      startCol = currentCol;
      currentGroup = [changes[currentCol]];
    }
  }
  ExecutionTimer.end("_applyTargetedWrites_grouping");

  // Write the last remaining group
  const lastRange = sheet.getRange(rowNum, startCol, 1, currentGroup.length);
  Log[sourceFile](
    `[${sourceFile} - _applyTargetedWrites] Row ${rowNum}: Writing final group to range ${lastRange.getA1Notation()}.`
  );
  if (currentGroup.length === 1) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "_applyTargetedWrites_writeFinalSingle",
    });
    lastRange.setValue(currentGroup[0]);
  } else {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "_applyTargetedWrites_writeFinalBatch",
    });
    lastRange.setValues([currentGroup]);
  }

  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "_applyTargetedWrites_end",
  });
  ExecutionTimer.end("_applyTargetedWrites_total");
}

/**
 * Main onEdit trigger handler.
 * REFACTORED FOR CONCURRENCY: This version uses a targeted write model to prevent
 * race conditions from rapid user edits. It reads the full row state, calculates
 * all necessary changes in memory, and then writes only the specific, automated
 * cells that need to be updated, leaving user-input columns untouched by the
 * final write operation. This ensures concurrent edits are not overwritten.
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
  const sheet = e.source.getActiveSheet();

  // --- EAGER CACHING & EFFICIENT EXIT LOGIC ---
  const configHeaderRange = sheet.getRange(
    CONFIG.offerDetailsCells.cachedHeaderRangeA1
  );
  if (_isRangeWithin(range, configHeaderRange)) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "handleSheetAutomations_isConfigEdit",
    });
    Log[sourceFile](
      "[handleSheetAutomations] A cached config cell was edited. Eagerly refreshing cache."
    );
    const cache = CacheService.getScriptCache();
    cache.remove("staticSheetValues");
    _getStaticSheetValues(sheet);
    ExecutionTimer.end("handleSheetAutomations_total");
    return;
  }

  if (range.getRow() < CONFIG.approvalWorkflow.startDataRow) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "handleSheetAutomations_exit_headerEdit",
    });
    ExecutionTimer.end("handleSheetAutomations_total");
    return;
  }

  ExecutionTimer.start("handleSheetAutomations_lock");
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);
  } catch (err) {
    Log.TestCoverage_gs({
      file: "SheetCoreAutomations.gs",
      coverage: "handleSheetAutomations_lock_fail",
    });
    Log[sourceFile](
      `[${sourceFile}] WARNING: Could not obtain lock. Error: ${err.message}.`
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

  let deferredTasks = []; // Initialize the deferred task queue
  let needsRecalculation = false;

  try {
    ExecutionTimer.start("handleSheetAutomations_flush");
    SpreadsheetApp.flush();
    Log[sourceFile](
      "[handleSheetAutomations] Lock acquired. Flushed all pending changes."
    );
    ExecutionTimer.end("handleSheetAutomations_flush");

    const combinedIndexes = {
      ...CONFIG.approvalWorkflow.columnIndices,
      ...CONFIG.documentDeviceData.columnIndices,
    };
    const editedRowStart = range.getRow();
    const numEditedRows = range.getNumRows();
    const editedCol = range.getColumn();
    const isSingleCellEdit = numEditedRows === 1 && range.getNumColumns() === 1;
    const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;

    const targetRange = sheet.getRange(
      editedRowStart,
      dataBlockStartCol,
      numEditedRows,
      CONFIG.maxDataColumn - dataBlockStartCol + 1
    );
    const inMemoryRowValuesForReading = targetRange.getValues();
    const originalValuesForComparison = JSON.parse(
      JSON.stringify(inMemoryRowValuesForReading)
    );

    if (isSingleCellEdit) {
      const editedColIndexInArray = editedCol - dataBlockStartCol;
      if (
        editedColIndexInArray >= 0 &&
        editedColIndexInArray < originalValuesForComparison[0].length
      ) {
        originalValuesForComparison[0][editedColIndexInArray] = e.oldValue;
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
      const inMemoryRowValues = inMemoryRowValuesForReading[i];
      const originalRowValues = originalValuesForComparison[i];
      let changesToWrite = {};

      const oldBundleInfo = _getBundleInfoFromRange(range);
      const wasBundleValidBeforeEdit = !!oldBundleInfo;

      Log[sourceFile](
        `[${sourceFile} - handleSheetAutomations] ---- STARTING ROW ${currentRowNumInSheet} ----`
      );

      if (oldBundleInfo) {
        wasBundleValidBeforeEdit = true;
      }

      // --- BQ Data wipe logic (Simplified and safer) ---
      const skuInRow = String(
        inMemoryRowValues[combinedIndexes.sku - dataBlockStartCol] || ""
      ).trim();
      const skuInRowBefore = String(
        originalRowValues[combinedIndexes.sku - dataBlockStartCol] || ""
      ).trim();
      if (skuInRow === "" && skuInRowBefore !== "") {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "handleSheetAutomations_wipeBqData_skuCleared",
        });
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: SKU was cleared. Queuing wipe of BQ-derived data.`
        );
        changesToWrite[combinedIndexes.epCapex] = "";
        changesToWrite[combinedIndexes.ep24PriceTarget] = "";
        changesToWrite[combinedIndexes.ep36PriceTarget] = "";
        changesToWrite[combinedIndexes.tkCapex] = "";
        changesToWrite[combinedIndexes.tk24PriceTarget] = "";
        changesToWrite[combinedIndexes.tk36PriceTarget] = "";
        changesToWrite[combinedIndexes.model] = "";
      }

      // --- Sanitization logic for Paste Operations ---
      if (!isSingleCellEdit) {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "handleSheetAutomations_pasteSanitization",
        });
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: Detected Paste. Queuing sanitization of automated fields.`
        );
        changesToWrite[combinedIndexes.index] = "";
        changesToWrite[combinedIndexes.lrfPreview] = "";
        changesToWrite[combinedIndexes.contractValuePreview] = "";
        changesToWrite[combinedIndexes.status] = "";
        changesToWrite[combinedIndexes.financeApprovedPrice] = "";
        changesToWrite[combinedIndexes.approvedBy] = "";
        changesToWrite[combinedIndexes.approvalDate] = "";
        changesToWrite[combinedIndexes.approverComments] = "";
        changesToWrite[combinedIndexes.approverPriceProposal] = "";
        changesToWrite[combinedIndexes.approverAction] = "Choose Action";
        originalRowValues[combinedIndexes.status - dataBlockStartCol] = "";
      }

      // --- Protected Column Revert (for single edits) ---
      if (
        isSingleCellEdit &&
        CONFIG.protectedColumnIndices.includes(editedCol)
      ) {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "handleSheetAutomations_protectedColumnRevert",
        });
        changesToWrite[editedCol] = e.oldValue;
        Log[sourceFile](
          `[${sourceFile} - handleSheetAutomations] Row ${currentRowNumInSheet}: PROTECTED COLUMN. Queuing revert on col ${editedCol} back to original value: '${e.oldValue}'.`
        );
      }

      // --- Downstream logic ---
      const calculatedValues = updateCalculationsForRow(
        inMemoryRowValues,
        combinedIndexes,
        CONFIG.approvalWorkflow,
        dataBlockStartCol
      );
      changesToWrite[combinedIndexes.lrfPreview] = calculatedValues.lrfPreview;
      changesToWrite[combinedIndexes.contractValuePreview] =
        calculatedValues.contractValuePreview;

      if (typeof calculatedValues.lrfPreview === "number") {
        deferredTasks.push({
          type: "SET_FORMAT",
          row: currentRowNumInSheet,
          col: combinedIndexes.lrfPreview,
          format: CONFIG.numberFormats.percentage,
        });
      }
      if (typeof calculatedValues.contractValuePreview === "number") {
        deferredTasks.push({
          type: "SET_FORMAT",
          row: currentRowNumInSheet,
          col: combinedIndexes.contractValuePreview,
          format: CONFIG.numberFormats.currency,
        });
      }

      const approverActionValue =
        inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol];
      const isApprovalAction =
        isSingleCellEdit &&
        editedCol === combinedIndexes.approverAction &&
        approverActionValue &&
        approverActionValue !== "Choose Action";

      if (isApprovalAction) {
        // Approval logic is a special case that modifies the row in-place and returns true/false.
        // We will adapt it fully in a subsequent step. For now, its existing logic flow is preserved.
        // The key is that `processSingleApprovalAction` will need to be refactored to populate `changesToWrite` instead of `inMemoryRowValues`.
        // To maintain stability, we'll let it modify the array for now and then extract the changes.
        const tempRowForApproval = JSON.parse(
          JSON.stringify(inMemoryRowValues)
        );
        const approvalSuccess = processSingleApprovalAction(
          sheet,
          currentRowNumInSheet,
          {
            ...e,
            value: approverActionValue,
            oldValue: originalRowValues[editedCol - dataBlockStartCol],
          },
          tempRowForApproval,
          combinedIndexes,
          originalRowValues,
          dataBlockStartCol
        );

        if (approvalSuccess) {
          // Compare the modified temp array to the original to see what changed
          tempRowForApproval.forEach((value, idx) => {
            if (value !== inMemoryRowValues[idx]) {
              changesToWrite[dataBlockStartCol + idx] = value;
            }
          });
        } else {
          // Revert the action in the sheet
          changesToWrite[combinedIndexes.approverAction] =
            originalRowValues[
              combinedIndexes.approverAction - dataBlockStartCol
            ];
        }
      } else {
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

        if (newStatus !== initialStatus) {
          logTableActivity({
            mainSheet: sheet,
            rowNum: currentRowNumInSheet,
            oldStatus: initialStatus,
            newStatus: newStatus,
            currentFullRowValues: inMemoryRowValues,
            originalFullRowValues: originalRowValues,
            startCol: dataBlockStartCol,
          });
          changesToWrite[combinedIndexes.status] =
            newStatus === null ? "" : newStatus;

          // REPLACE WITH THIS BLOCK
          // Use the centralized function to get side-effect changes
          const sideEffectChanges = _getSideEffectChanges(
            initialStatus,
            newStatus,
            combinedIndexes
          );
          Object.assign(changesToWrite, sideEffectChanges); // Merge the changes

          if (newStatus !== null) {
            // Logic to reset the approver dropdown if the new status is a pending one
            const pendingStatuses = [
              CONFIG.approvalWorkflow.statusStrings.pending,
              CONFIG.approvalWorkflow.statusStrings.draft,
            ];
            if (pendingStatuses.includes(newStatus)) {
              changesToWrite[combinedIndexes.approverAction] = "Choose Action";
            }
          }
        }
      }

      // --- Initialization checks ---
      const modelName =
        inMemoryRowValues[combinedIndexes.model - dataBlockStartCol];
      if (
        modelName &&
        !inMemoryRowValues[combinedIndexes.index - dataBlockStartCol]
      ) {
        if (nextIndex === null) {
          nextIndex = getNextAvailableIndex(sheet);
        }
        changesToWrite[combinedIndexes.index] = nextIndex++;
      }
      if (
        modelName &&
        !inMemoryRowValues[combinedIndexes.approverAction - dataBlockStartCol]
      ) {
        changesToWrite[combinedIndexes.approverAction] = "Choose Action";
      }

      // --- REFACTORED BUNDLE LOGIC ---
      if (
        CONFIG.featureFlags.enforceBundleIntegrityOnEdit &&
        isSingleCellEdit
      ) {
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
          const currentBundleNum =
            inMemoryRowValues[combinedIndexes.bundleNumber - dataBlockStartCol];

          // A. Handle cleaning up old bundle state (if it was part of one before)
          if (
            wasBundleValidBeforeEdit &&
            String(currentBundleNum).trim() !==
              String(oldBundleInfo.bundleId).trim()
          ) {
            Log.TestCoverage_gs({
              file: sourceFile,
              coverage: "handleSheetAutomations_bundleBrokenOrLeft",
            });
            deferredTasks.push({
              type: "CLEAR_BUNDLE_FORMATTING",
              bundleInfo: oldBundleInfo,
            });
          }

          // B. Validate the new bundle state
          if (currentBundleNum) {
            const validationResult = validateBundle(
              sheet,
              currentRowNumInSheet,
              currentBundleNum
            );
            const bundleRows = _findRowsForBundle(sheet, currentBundleNum);

            if (!validationResult.isValid) {
              Log.TestCoverage_gs({
                file: sourceFile,
                coverage: "handleSheetAutomations_bundleIsInvalid",
              });
              Log[sourceFile](
                "[handleSheetAutomations] Bundle is invalid. Showing dialog and queuing invalid formatting."
              );

              // Queue 'invalid' formatting task
              if (bundleRows.length > 0) {
                deferredTasks.push({
                  type: "APPLY_INVALID_BUNDLE_FORMATTING",
                  bundleInfo: {
                    bundleId: String(currentBundleNum),
                    startRow: bundleRows[0],
                    endRow: bundleRows[bundleRows.length - 1],
                  },
                });
              }
              // Show the appropriate non-destructive dialog
              if (validationResult.errorCode === "MISMATCH") {
                showBundleMismatchDialog(
                  currentRowNumInSheet,
                  currentBundleNum,
                  {},
                  validationResult.expected
                );
              } else if (validationResult.errorCode === "GAP_DETECTED") {
                showBundleGapDialog(currentBundleNum);
              }
            } else {
              Log.TestCoverage_gs({
                file: sourceFile,
                coverage: "handleSheetAutomations_bundleIsValid",
              });
              // Bundle is VALID: Queue 'valid' formatting
              if (
                validationResult.startRow &&
                validationResult.endRow &&
                validationResult.startRow !== validationResult.endRow
              ) {
                deferredTasks.push({
                  type: "APPLY_BUNDLE_FORMATTING",
                  bundleInfo: {
                    bundleId: String(currentBundleNum),
                    startRow: validationResult.startRow,
                    endRow: validationResult.endRow,
                  },
                });
              }
            }
          }
        }
      }

      // --- Final Write and Loop Continuation ---
      _applyTargetedWrites(sheet, currentRowNumInSheet, changesToWrite);
    }
    ExecutionTimer.end("handleSheetAutomations_mainLoop");

    if (deferredTasks.length > 0) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "handleSheetAutomations_executingDeferredTasks",
      });
      _executeDeferredTasks(sheet, deferredTasks);
    }
  } finally {
    lock.releaseLock();
  }
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "handleSheetAutomations_end",
  });
  ExecutionTimer.end("handleSheetAutomations_total");
}

/**
 * Calculates the LRF and Contract Value for a specific row's in-memory data.
 * It is now a PURE function: it calculates and RETURNS the values, rather than
 * directly modifying the rowValues array or setting cell formats.
 * @param {Array<any>} rowValues The in-memory array of values for the row (READ-ONLY).
 * @param {Object} colIndexes A map of column names to their 1-based index.
 * @param {Object} approvalWorkflowConfig The CONFIG.approvalWorkflow object.
 * @param {number} dataBlockStartCol The 1-based index of the starting column for the rowValues array.
 * @returns {{lrfPreview: (number|string|null), contractValuePreview: (number|null)}} An object containing the calculated values.
 */
function updateCalculationsForRow(
  rowValues, // No longer accepts 'sheet' or 'rowNum' as it's pure
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

  const chosenCapex = getNumericValue(
    rowValues[colIndexes.aeCapex - dataBlockStartCol]
  );
  Log[sourceFile](
    `[${sourceFile} - updateCalculationsForRow] Inputs: rentalPrice=${rentalPrice}, chosenCapex=${chosenCapex}`
  );

  let lrfPreviewValue = null;
  let contractValuePreviewValue = null;

  ExecutionTimer.start("updateCalculationsForRow_calcLrf");
  if (rentalPrice === 0 && (!chosenCapex || chosenCapex === 0)) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "updateCalculationsForRow_clearValues",
    });
    lrfPreviewValue = ""; // No longer setting to null to match existing empty string behavior
    contractValuePreviewValue = ""; // No longer setting to null to match existing empty string behavior
    Log[sourceFile](
      `[${sourceFile} - updateCalculationsForRow] Clearing LRF and Contract Value due to zero price/capex.`
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
      lrfPreviewValue = "Missing\nCAPEX";
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
        lrfPreviewValue = (rentalPrice * term) / chosenCapex;
      } else {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "updateCalculationsForRow_lrfZero",
        });
        lrfPreviewValue = "";
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
      contractValuePreviewValue = rentalPrice * term * quantity;
    } else {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "updateCalculationsForRow_cvZero",
      });
      contractValuePreviewValue = "";
    }
  }
  ExecutionTimer.end("updateCalculationsForRow_calcLrf");

  // Removed direct setNumberFormat calls as this is now a pure function.
  // Formatting will be handled when values are written to the sheet.

  Log[sourceFile](
    `[${sourceFile} - updateCalculationsForRow] Outputs: LRF=${lrfPreviewValue}, ContractValue=${contractValuePreviewValue}`
  );
  Log.TestCoverage_gs({
    file: "SheetCoreAutomations.gs",
    coverage: "updateCalculationsForRow_end",
  });
  ExecutionTimer.end("updateCalculationsForRow_total");

  return {
    lrfPreview: lrfPreviewValue,
    contractValuePreview: contractValuePreviewValue,
  };
}

/**
 * Processes a queue of deferred UI/Metadata tasks after the main data write is complete.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {Array<Object>} tasks An array of task objects to execute.
 */
function _executeDeferredTasks(sheet, tasks) {
  const sourceFile = "SheetCoreAutomations_gs";
  ExecutionTimer.start("_executeDeferredTasks_total");
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "_executeDeferredTasks_start",
  });

  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numColsInDataBlock = CONFIG.maxDataColumn - dataBlockStartCol + 1;

  const uniqueTasks = tasks.filter(
    (task, index, self) =>
      index ===
      self.findIndex((t) => JSON.stringify(t) === JSON.stringify(task))
  );

  Log[sourceFile](
    `[${sourceFile}] Processing ${uniqueTasks.length} unique deferred tasks.`
  );

  uniqueTasks.forEach((task) => {
    try {
      switch (task.type) {
        case "SET_FORMAT":
          Log.TestCoverage_gs({
            file: sourceFile,
            coverage: "_executeDeferredTasks_setFormat",
          });
          sheet.getRange(task.row, task.col).setNumberFormat(task.format);
          break;

        case "APPLY_BUNDLE_FORMATTING":
          Log.TestCoverage_gs({
            file: sourceFile,
            coverage: "_executeDeferredTasks_applyBundleFormatting",
          });
          const bundleRange = sheet.getRange(
            task.bundleInfo.startRow,
            dataBlockStartCol,
            task.bundleInfo.endRow - task.bundleInfo.startRow + 1,
            numColsInDataBlock
          );
          _setMetadataForRowRange(sheet, task.bundleInfo);
          _clearAndApplyBundleBorder(bundleRange);
          break;

        case "APPLY_INVALID_BUNDLE_FORMATTING":
          Log.TestCoverage_gs({
            file: sourceFile,
            coverage: "_executeDeferredTasks_applyInvalidBundleFormatting",
          });
          const invalidBundleRange = sheet.getRange(
            task.bundleInfo.startRow,
            dataBlockStartCol,
            task.bundleInfo.endRow - task.bundleInfo.startRow + 1,
            numColsInDataBlock
          );
          invalidBundleRange.setBorder(
            true,
            true,
            true,
            true,
            false,
            false,
            "#ff0000",
            SpreadsheetApp.BorderStyle.DASHED
          );
          break;

        case "CLEAR_BUNDLE_FORMATTING":
          Log.TestCoverage_gs({
            file: sourceFile,
            coverage: "_executeDeferredTasks_clearBundleFormatting",
          });
          const oldRange = sheet.getRange(
            task.bundleInfo.startRow,
            dataBlockStartCol,
            task.bundleInfo.endRow - task.bundleInfo.startRow + 1,
            numColsInDataBlock
          );
          _clearMetadataFromRowRange(
            sheet,
            task.bundleInfo.startRow,
            task.bundleInfo.endRow
          );
          oldRange.setBorder(null, null, null, null, null, null);
          break;
      }
    } catch (err) {
      Log[sourceFile](
        `[${sourceFile}] ERROR processing deferred task: ${JSON.stringify(
          task
        )}. Error: ${err.message}`
      );
    }
  });
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "_executeDeferredTasks_end",
  });
  ExecutionTimer.end("_executeDeferredTasks_total");
}

/**
 * Determines side-effect changes based on a status transition.
 * Specifically handles wiping approval data when a finalized item is reverted.
 * @param {string} initialStatus The status before the change.
 * @param {string} newStatus The status after the change.
 * @param {Object} allColIndexes A map of all column indices.
 * @returns {Object} An object of changes to be written to the sheet.
 */
function _getSideEffectChanges(initialStatus, newStatus, allColIndexes) {
  const statusStrings = CONFIG.approvalWorkflow.statusStrings;
  const finalizedStatuses = [
    statusStrings.approvedOriginal,
    statusStrings.approvedNew,
    statusStrings.rejected,
  ];

  let changes = {};

  // RULE 1: If a row was finalized and is now pending/draft, wipe all approval data.
  if (
    finalizedStatuses.includes(initialStatus) &&
    !finalizedStatuses.includes(newStatus)
  ) {
    changes[allColIndexes.financeApprovedPrice] = "";
    changes[allColIndexes.approvedBy] = "";
    changes[allColIndexes.approvalDate] = "";
  }

  // RULE 2: If a row's new status is "Rejected", ensure the approved price is always blank.
  if (newStatus === statusStrings.rejected) {
    changes[allColIndexes.financeApprovedPrice] = "";
  }

  return changes;
}

/**
 * Calculates and updates BOTH the LRF and Contract Value for a specific row's in-memory data,
 * AND applies the correct number format to the corresponding cells.
 * REVISED: Now correctly uses the single 'aeCapex' column.

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
 * Main onEdit trigger handler.
 * REVISED: This version is robust against inconsistent onEdit event objects,
 * uses e.oldValue for single edits, accepts an injected array for testing,
 * and contains comprehensive logging. The BQ data wipe logic now correctly
 * handles both single edits and partial pastes by inspecting the edit range.
 * Includes SpreadsheetApp.flush() to prevent race conditions from rapid edits.
 * Bundle integrity validation and formatting is now handled efficiently
 * using ROW-LEVEL Developer Metadata for immediate feedback.
 * Business logic simplified to no longer wipe AE data on model changes.

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

*/
