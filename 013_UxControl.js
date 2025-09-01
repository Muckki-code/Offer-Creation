/**
 * @file This file contains functions related to User Experience (UX) controls
 * within the Google Sheet, including conditional formatting and data validation.
 */

// In UxControl.gs

/**
 * Applies all UX rules to the sheet.
 * REVISED: Now includes logic for the new dynamic approver dropdown.
 */
function applyUxRules(formatIncluded = true) {
  const sourceFile = "UxControl_gs";
  ExecutionTimer.start("applyUxRules_total");
  Log[sourceFile](
    `[${sourceFile} - applyUxRules] Start. formatIncluded = ${formatIncluded}`
  );
  Log[sourceFile](
    `[${sourceFile} - applyUxRules] CRAZY VERBOSE: Beginning full UX rule application.`
  );

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = getLastLastRow(sheet);

  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numColsInDataBlock = CONFIG.maxDataColumn - dataBlockStartCol + 1;

  const numRows = lastRow >= startRow ? lastRow - startRow + 1 : 1;
  const specificDataRange = sheet.getRange(
    startRow,
    dataBlockStartCol,
    numRows,
    numColsInDataBlock
  );

  const openEndedRangeA1 = specificDataRange
    .getA1Notation()
    .replace(/\d+$/, "");
  const openEndedRange = sheet.getRange(openEndedRangeA1);
  Log[sourceFile](
    `[${sourceFile} - applyUxRules] Determined specific range: ${specificDataRange.getA1Notation()} and open-ended range for CF: ${openEndedRange.getA1Notation()}`
  );

  if (formatIncluded) {
    ExecutionTimer.start("applyUxRules_conditionalFormatting");
    Log[sourceFile](
      "[${sourceFile} - applyUxRules] Applying conditional formatting for row colors."
    );
    sheet.clearConditionalFormatRules();

    let conditionalRules = [];
    const statusColLetter = CONFIG.approvalWorkflow.columns.status;
    const indexColLetter = CONFIG.documentDeviceData.columns.index;
    const statusStrings = CONFIG.approvalWorkflow.statusStrings;
    const colors = CONFIG.conditionalFormatColors;

    const isRowNotBlank = `NOT(ISBLANK($${indexColLetter}${startRow}))`;

    conditionalRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND(${isRowNotBlank}, OR($${statusColLetter}${startRow}="${statusStrings.approvedOriginal}", $${statusColLetter}${startRow}="${statusStrings.approvedNew}"))`
        )
        .setBackground(colors.approved.background)
        .setRanges([openEndedRange])
        .build()
    );
    conditionalRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND(${isRowNotBlank}, OR($${statusColLetter}${startRow}="${statusStrings.pending}", $${statusColLetter}${startRow}="${statusStrings.revisedByAE}"))`
        )
        .setBackground(colors.pending.background)
        .setRanges([openEndedRange])
        .build()
    );
    conditionalRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND(${isRowNotBlank}, $${statusColLetter}${startRow}="${statusStrings.rejected}")`
        )
        .setBackground(colors.rejected.background)
        .setRanges([openEndedRange])
        .build()
    );
    conditionalRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND(${isRowNotBlank}, $${statusColLetter}${startRow}="${statusStrings.draft}")`
        )
        .setBackground(colors.draft.background)
        .setRanges([openEndedRange])
        .build()
    );
    conditionalRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND(${isRowNotBlank}, ISBLANK($${statusColLetter}${startRow}))`
        )
        .setBackground("#ffffff")
        .setRanges([openEndedRange])
        .build()
    );

    sheet.setConditionalFormatRules(conditionalRules);
    ExecutionTimer.end("applyUxRules_conditionalFormatting");

    if (CONFIG.featureFlags.highlightBundlesWithBorders) {
      Log[sourceFile](
        `[${sourceFile} - applyUxRules] Feature flag for bundle borders is ON. Calling refreshBundleBorders.`
      );
      refreshBundleBorders();
    }

    ExecutionTimer.start("applyUxRules_numberFormatting");
    const formats = CONFIG.numberFormats;
    const allColIndexes = {
      ...CONFIG.documentDeviceData.columnIndices,
      ...CONFIG.approvalWorkflow.columnIndices,
    };
    const currencyCols = [
      "epCapexRaw",
      "tkCapexRaw",
      "rentalTargetRaw",
      "rentalLimitRaw",
      "aeCapex", // UPDATED
      "aeSalesAskPrice",
      "approverPriceProposal",
      "contractValuePreview",
      "financeApprovedPrice",
    ];
    currencyCols.forEach((key) => {
      const colIndex = allColIndexes[key];
      if (colIndex) {
        sheet
          .getRange(startRow, colIndex, numRows)
          .setNumberFormat(formats.currency);
      }
    });
    const lrfCol = allColIndexes["lrfPreview"];
    if (lrfCol) {
      const lrfOpenEndedRange = sheet.getRange(
        startRow,
        lrfCol,
        sheet.getMaxRows() - startRow + 1,
        1
      );
      lrfOpenEndedRange.setNumberFormat(formats.percentage);
    }
    ExecutionTimer.end("applyUxRules_numberFormatting");
  }

  ExecutionTimer.start("applyUxRules_dataValidation");
  // --- Approver Action Dropdown ---
  const approverActionCol =
    CONFIG.approvalWorkflow.columnIndices.approverAction;
  const dropdownTargetRange = sheet.getRange(
    startRow,
    approverActionCol,
    numRows
  );
  const approverActions = Object.keys(
    CONFIG.approvalWorkflow.approverActionColors
  );
  const dropdownValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(approverActions, true)
    .setAllowInvalid(false)
    .build();
  dropdownTargetRange.setDataValidation(dropdownValidation);

  // --- NEW: Dynamic Approver Dropdown ---
  Log[sourceFile](
    `[${sourceFile} - applyUxRules] CRAZY VERBOSE: Setting up dynamic approver dropdown.`
  );
  const approverCellA1 = CONFIG.offerDetailsCells.approverCell;
  const approverList = CONFIG.approvalWorkflow.approverList;
  if (approverCellA1 && approverList && approverList.length > 0) {
    Log[sourceFile](
      `[${sourceFile} - applyUxRules] Applying approver list validation to cell ${approverCellA1}.`
    );
    const approverCell = sheet.getRange(approverCellA1);
    const approverValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(approverList, true)
      .setAllowInvalid(false)
      .build();
    approverCell.setDataValidation(approverValidation);
  } else {
    Log[sourceFile](
      `[${sourceFile} - applyUxRules] CRAZY VERBOSE: Skipping dynamic approver dropdown setup (missing config).`
    );
  }
  // --- END NEW ---

  ExecutionTimer.end("applyUxRules_dataValidation");

  Log[sourceFile](`[${sourceFile} - applyUxRules] Finished successfully.`);
  ExecutionTimer.end("applyUxRules_total");
}

/**
 * A helper function that clears all borders from a range and then applies the standard bundle border.
 * This is now the single point of control for drawing borders.
 * REVISED: Now intelligently handles the top border to avoid clearing the header row's bottom border.
 * @private
 */
function _clearAndApplyBundleBorder(range) {
  const sourceFile = "UxControl_gs";
  Log[sourceFile](
    `[${sourceFile} - _clearAndApplyBundleBorder] Applying thick border to range ${range.getA1Notation()}.`
  );

  // --- THIS IS THE FIX for BUG 2 ---
  // Check if the range starts on the very first data row.
  const isFirstDataRow =
    range.getRow() === CONFIG.approvalWorkflow.startDataRow;

  range.setBorder(
    isFirstDataRow ? false : true, // top: Don't apply a top border if it's the first row
    true, // left
    true, // bottom
    true, // right
    false, // vertical
    false, // horizontal
    "#666666",
    SpreadsheetApp.BorderStyle.SOLID_THICK
  );
}

/**
 * A helper function to apply borders directly to bundle groups.
 * REFACTORED: This version is now highly performant. It reads all developer
 * metadata in a single call and applies borders only to the ranges defined therein,
 * avoiding any slow column scans.
 * REVISED: Reads column count dynamically from CONFIG to prevent regressions.
 * @private
 */
function _applyBundleBordersDirectly(sheet, startRow, numRows) {
  const sourceFile = "UxControl_gs";
  Log[sourceFile](
    "[${sourceFile} - _applyBundleBordersDirectly] Start (Metadata Version)."
  );

  // --- REGRESSION FIX ---
  const startCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - startCol + 1;
  Log[sourceFile](
    `[${sourceFile} - _applyBundleBordersDirectly] CRAZY VERBOSE: Using dynamic startCol=${startCol} and numCols=${numCols}.`
  );
  // --- END FIX ---

  const fullDataRange = sheet.getRange(startRow, startCol, numRows, numCols);

  // 1. Clear all existing borders first
  fullDataRange.setBorder(null, null, null, null, null, null);
  Log[sourceFile](
    "[${sourceFile} - _applyBundleBordersDirectly] Cleared all existing borders in data range."
  );

  // 2. Get all bundle metadata from the sheet's data range in one call.
  const allMetadata = sheet
    .getRange(startRow, 1, numRows, 1)
    .getDeveloperMetadata();
  const bundleMetadata = allMetadata.filter(
    (m) => m.getKey() === METADATA_KEY_BUNDLE
  );

  if (bundleMetadata.length === 0) {
    Log[sourceFile](
      "[${sourceFile} - _applyBundleBordersDirectly] No bundle metadata found. Exiting."
    );
    return;
  }

  // 3. Use a Set to only process each unique bundle range once.
  const processedRanges = new Set();

  bundleMetadata.forEach((meta) => {
    try {
      const bundleInfo = JSON.parse(meta.getValue());
      const bundleRange = sheet.getRange(
        bundleInfo.startRow,
        startCol,
        bundleInfo.endRow - bundleInfo.startRow + 1,
        numCols
      );
      const rangeA1 = bundleRange.getA1Notation();

      if (
        !processedRanges.has(rangeA1) &&
        bundleInfo.endRow - bundleInfo.startRow > 0
      ) {
        Log[sourceFile](
          `[${sourceFile} - _applyBundleBordersDirectly] Found bundle #${bundleInfo.bundleId}. Applying border to range ${rangeA1}.`
        );
        _clearAndApplyBundleBorder(bundleRange);
        processedRanges.add(rangeA1);
      }
    } catch (e) {
      Log[sourceFile](
        "[${sourceFile} - _applyBundleBordersDirectly] WARNING: Could not parse bundle metadata. Skipping. Value: " +
          meta.getValue()
      );
    }
  });

  Log[sourceFile](
    "[${sourceFile} - _applyBundleBordersDirectly] Finished applying bundle borders."
  );
}

/**
 * --- NEW, PUBLIC FUNCTION ---
 * Clears all borders in the data range and redraws borders ONLY for valid
 * bundles based on the current row metadata. This is the definitive function
 * for ensuring bundle borders are visually correct.
 * REVISED: Reads column count dynamically from CONFIG to prevent regressions.
 */
function refreshBundleBorders() {
  const sourceFile = "UxControl_gs";
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "refreshBundleBorders_start",
  });
  ExecutionTimer.start("refreshBundleBorders_total");
  Log[sourceFile](
    `[${sourceFile} - refreshBundleBorders] CRAZY VERBOSE: Starting full refresh of bundle borders.`
  );

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = getLastLastRow(sheet);

  if (lastRow < startRow) {
    Log[sourceFile](
      `[${sourceFile} - refreshBundleBorders] No data rows found. Exiting.`
    );
    ExecutionTimer.end("refreshBundleBorders_total");
    return;
  }

  // --- REGRESSION FIX ---
  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - dataBlockStartCol + 1;
  Log[sourceFile](
    `[${sourceFile} - refreshBundleBorders] CRAZY VERBOSE: Using dynamic startCol=${dataBlockStartCol} and numCols=${numCols}.`
  );
  // --- END FIX ---

  const fullDataRange = sheet.getRange(
    startRow,
    dataBlockStartCol,
    lastRow - startRow + 1,
    numCols
  );

  // 1. Clear all existing borders first
  fullDataRange.setBorder(null, null, null, null, null, null);
  Log[sourceFile](
    "[refreshBundleBorders] Cleared all existing borders in data range."
  );

  // 2. Get all metadata and redraw valid borders
  const allMetadata = sheet
    .getRange(startRow, 1, lastRow - startRow + 1, 1)
    .getDeveloperMetadata();
  const bundleMetadata = allMetadata.filter(
    (m) => m.getKey() === METADATA_KEY_BUNDLE
  );

  if (bundleMetadata.length === 0) {
    Log[sourceFile](
      "[refreshBundleBorders] No bundle metadata found. Exiting."
    );
    ExecutionTimer.end("refreshBundleBorders_total");
    return;
  }

  const processedRanges = new Set();
  bundleMetadata.forEach((meta) => {
    try {
      const bundleInfo = JSON.parse(meta.getValue());
      if (
        bundleInfo &&
        bundleInfo.startRow &&
        bundleInfo.endRow &&
        bundleInfo.startRow < bundleInfo.endRow
      ) {
        const bundleRange = sheet.getRange(
          bundleInfo.startRow,
          dataBlockStartCol,
          bundleInfo.endRow - bundleInfo.startRow + 1,
          numCols
        );
        const rangeA1 = bundleRange.getA1Notation();

        if (!processedRanges.has(rangeA1)) {
          Log[sourceFile](
            `[${sourceFile} - refreshBundleBorders] CRAZY VERBOSE: Redrawing border for bundle #${bundleInfo.bundleId} on range ${rangeA1}.`
          );
          _clearAndApplyBundleBorder(bundleRange);
          processedRanges.add(rangeA1);
        }
      }
    } catch (e) {
      // Ignore parse errors
    }
  });

  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "refreshBundleBorders_end",
  });
  ExecutionTimer.end("refreshBundleBorders_total");
}
