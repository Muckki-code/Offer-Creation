/**
 * @file This file contains functions related to User Experience (UX) controls
 * within the Google Sheet, including conditional formatting and data validation.
 */

// In UxControl.gs

/**
 * Applies all UX rules to the sheet.
 * REVISED: Conditional formatting formulas are now more robust, checking for a non-blank
 * row identifier to prevent "FALSE" values from appearing in blank rows.
 */
function applyUxRules(formatIncluded = true) {
  const sourceFile = 'UxControl_gs';
  ExecutionTimer.start('applyUxRules_total');
  Log[sourceFile](`[${sourceFile} - applyUxRules] Start. formatIncluded = ${formatIncluded}`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = getLastLastRow(sheet);

  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numColsInDataBlock = CONFIG.maxDataColumn - dataBlockStartCol + 1;
  
  const numRows = lastRow >= startRow ? lastRow - startRow + 1 : 1;
  const specificDataRange = sheet.getRange(startRow, dataBlockStartCol, numRows, numColsInDataBlock);
  
  const openEndedRangeA1 = specificDataRange.getA1Notation().replace(/\d+$/, "");
  const openEndedRange = sheet.getRange(openEndedRangeA1);
  Log[sourceFile](`[${sourceFile} - applyUxRules] Determined specific range: ${specificDataRange.getA1Notation()} and open-ended range for CF: ${openEndedRange.getA1Notation()}`);

  if (formatIncluded) {
    ExecutionTimer.start('applyUxRules_conditionalFormatting');
    Log[sourceFile]("[${sourceFile} - applyUxRules] Applying conditional formatting for row colors.");
    sheet.clearConditionalFormatRules();
    
    let conditionalRules = [];
    const statusColLetter = CONFIG.approvalWorkflow.columns.status;
    const indexColLetter = CONFIG.documentDeviceData.columns.index; // Using Index as the row identifier
    const statusStrings = CONFIG.approvalWorkflow.statusStrings;
    const colors = CONFIG.conditionalFormatColors;

    // --- REVISED AND SAFER FORMULAS ---
    const isRowNotBlank = `NOT(ISBLANK($${indexColLetter}${startRow}))`;

    conditionalRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${isRowNotBlank}, OR($${statusColLetter}${startRow}="${statusStrings.approvedOriginal}", $${statusColLetter}${startRow}="${statusStrings.approvedNew}"))`).setBackground(colors.approved.background).setRanges([openEndedRange]).build());
    conditionalRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${isRowNotBlank}, OR($${statusColLetter}${startRow}="${statusStrings.pending}", $${statusColLetter}${startRow}="${statusStrings.revisedByAE}"))`).setBackground(colors.pending.background).setRanges([openEndedRange]).build());
    conditionalRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${isRowNotBlank}, $${statusColLetter}${startRow}="${statusStrings.rejected}")`).setBackground(colors.rejected.background).setRanges([openEndedRange]).build());
    conditionalRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${isRowNotBlank}, $${statusColLetter}${startRow}="${statusStrings.draft}")`).setBackground(colors.draft.background).setRanges([openEndedRange]).build());
    conditionalRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${isRowNotBlank}, ISBLANK($${statusColLetter}${startRow}))`).setBackground("#ffffff").setRanges([openEndedRange]).build());
    
    sheet.setConditionalFormatRules(conditionalRules);
    ExecutionTimer.end('applyUxRules_conditionalFormatting');
    
    // ... (Rest of the function is unchanged)
    if (CONFIG.featureFlags.highlightBundlesWithBorders) {
      ExecutionTimer.start('applyUxRules_bundleBorders');
      _applyBundleBordersDirectly(sheet, startRow, numRows, dataBlockStartCol, numColsInDataBlock);
      ExecutionTimer.end('applyUxRules_bundleBorders');
    }
    
    ExecutionTimer.start('applyUxRules_numberFormatting');
    const formats = CONFIG.numberFormats;
    const allColIndexes = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };
    const currencyCols = ['epCapexRaw', 'tkCapexRaw', 'rentalTargetRaw', 'rentalLimitRaw', 'aeEpCapex', 'aeTkCapex', 'aeSalesAskPrice', 'approverPriceProposal', 'contractValuePreview', 'financeApprovedPrice'];
    currencyCols.forEach(key => {
      const colIndex = allColIndexes[key];
      if (colIndex) { sheet.getRange(startRow, colIndex, numRows).setNumberFormat(formats.currency); }
    });
    const lrfCol = allColIndexes['lrfPreview'];
    if (lrfCol) {
      const lrfOpenEndedRange = sheet.getRange(startRow, lrfCol, sheet.getMaxRows() - startRow + 1, 1);
      lrfOpenEndedRange.setNumberFormat(formats.percentage);
    }
    ExecutionTimer.end('applyUxRules_numberFormatting');
  }

  ExecutionTimer.start('applyUxRules_dataValidation');
  const approverActionCol = CONFIG.approvalWorkflow.columnIndices.approverAction;
  const dropdownTargetRange = sheet.getRange(startRow, approverActionCol, numRows);
  const approverActions = Object.keys(CONFIG.approvalWorkflow.approverActionColors);
  const dropdownValidation = SpreadsheetApp.newDataValidation().requireValueInList(approverActions, true).setAllowInvalid(false).build();
  dropdownTargetRange.setDataValidation(dropdownValidation);
  ExecutionTimer.end('applyUxRules_dataValidation');
  
  Log[sourceFile](`[${sourceFile} - applyUxRules] Finished successfully.`);
  ExecutionTimer.end('applyUxRules_total');
}

/**
 * A helper function that clears all borders from a range and then applies the standard bundle border.
 * This is now the single point of control for drawing borders.
 * @private
 */
function _clearAndApplyBundleBorder(range) {
  const sourceFile = 'UxControl_gs';
  Log[sourceFile] (`[${sourceFile} - _clearAndApplyBundleBorder] Applying thick border to range ${range.getA1Notation()}.`);
  range.setBorder(
    true, true, true, true, // top, left, bottom, right
    false, false, // vertical, horizontal
    '#666666', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
}
/**
 * A helper function to apply borders directly to bundle groups.
 * REFACTORED: This version is now highly performant. It reads all developer
 * metadata in a single call and applies borders only to the ranges defined therein,
 * avoiding any slow column scans.
 * @private
 */
function _applyBundleBordersDirectly(sheet, startRow, numRows, startCol, numCols) {
  const sourceFile = 'UxControl_gs';
  Log[sourceFile]("[${sourceFile} - _applyBundleBordersDirectly] Start (Metadata Version).");
  const fullDataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  
  // 1. Clear all existing borders first
  fullDataRange.setBorder(null, null, null, null, null, null);
  Log[sourceFile]("[${sourceFile} - _applyBundleBordersDirectly] Cleared all existing borders in data range.");

  // 2. Get all bundle metadata from the sheet's data range in one call.
  // We only need to check the first column as metadata is attached to the row.
  const allMetadata = sheet.getRange(startRow, 1, numRows, 1).getDeveloperMetadata();
  const bundleMetadata = allMetadata.filter(m => m.getKey() === METADATA_KEY_BUNDLE);

  if (bundleMetadata.length === 0) {
    Log[sourceFile]("[${sourceFile} - _applyBundleBordersDirectly] No bundle metadata found. Exiting.");
    return;
  }
  
  // 3. Use a Set to only process each unique bundle range once.
  const processedRanges = new Set();

  bundleMetadata.forEach(meta => {
    try {
      const bundleInfo = JSON.parse(meta.getValue());
      const bundleRange = sheet.getRange(bundleInfo.startRow, startCol, bundleInfo.endRow - bundleInfo.startRow + 1, numCols);
      const rangeA1 = bundleRange.getA1Notation();

      if (!processedRanges.has(rangeA1) && (bundleInfo.endRow - bundleInfo.startRow > 0)) {
        Log[sourceFile](`[${sourceFile} - _applyBundleBordersDirectly] Found bundle #${bundleInfo.bundleId}. Applying border to range ${rangeA1}.`);
        _clearAndApplyBundleBorder(bundleRange);
        processedRanges.add(rangeA1);
      }
    } catch (e) {
      Log[sourceFile]("[${sourceFile} - _applyBundleBordersDirectly] WARNING: Could not parse bundle metadata. Skipping. Value: " + meta.getValue());
    }
  });

  Log[sourceFile]("[${sourceFile} - _applyBundleBordersDirectly] Finished applying bundle borders.");
}
