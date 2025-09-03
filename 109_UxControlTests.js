/**
 * @file This file contains the test suite for functions in UxControl.gs.
 */

/**
 * Runs all INTEGRATION tests for UxControl.gs.
 */
function runUxControl_IntegrationTests() {
  Log.TestResults_gs("--- Starting UxControl Integration Test Suite ---");

  test_applyUxRules_modes();
  test_refreshBundleBorders_onEdit(); // <-- ADDED THIS NEW TEST

  Log.TestResults_gs("--- UxControl Integration Test Suite Finished ---");
}

/**
 * --- NEW TEST (NOW CORRECTED) ---
 * Integration test to verify that bundle borders are correctly redrawn after
 * an edit creates a new bundle.
 */
function test_refreshBundleBorders_onEdit() {
  const testName = "Integration Test: refreshBundleBorders on Edit";

  // --- THIS IS THE FIX: The CSV data now correctly has 21 columns ---
  const csvData = `SKU,EP CAPEX,Telekom CAPEX,Target,Limit,Index,Bundle Number,Device,AE CAPEX,AE SALES ASK,QUANTITY,TERM,APPROVER_ACTION,APPROVER_COMMENTS,APPROVER_PRICE_PROPOSAL,LRF_PREVIEW,CONTRACT_VALUE,STATUS,FINANCE_APPROVED_PRICE,APPROVED_BY,APPROVAL_DATE
,,,,,"1",,"Item A",1000,100,5,24,"Choose Action","","","","",Draft,"","",""
,,,,,"2",,"Item B",1000,100,5,24,"Choose Action","","","","",Draft,"","",""
`;
  // --- END FIX ---

  withTestConfig(function () {
    withTestSheet(csvData, function (sheet) {
      const startRow = CONFIG.approvalWorkflow.startDataRow; // 2
      const bundleCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;
      const skuCol = CONFIG.documentDeviceData.columnIndices.sku;
      const lastDataCol = CONFIG.maxDataColumn;

      // --- ACTION: Simulate user edit to create a bundle ---
      sheet.getRange(startRow, bundleCol).setValue(707);
      sheet.getRange(startRow + 1, bundleCol).setValue(707);
      
      // Before applying rules, we must set the metadata, as this is what refreshBundleBorders relies on.
      scanAndSetAllBundleMetadata();
      SpreadsheetApp.flush();

      // Trigger the UX rules which should include the border refresh
      applyUxRules(true);
      SpreadsheetApp.flush();

      // --- VERIFICATION ---
      const bundleRange = sheet.getRange(startRow, skuCol, 2, lastDataCol - skuCol + 1);
      const border = bundleRange.getBorder();
      
      _assertNotNull(border, `${testName} - Border object should not be null.`);

      const topBorder = border.getTop();
      const bottomBorder = border.getBottom();
      
      _assertNotNull(topBorder, `${testName} - Top border line should not be null.`);
      _assertNotNull(bottomBorder, `${testName} - Bottom border line should not be null.`);

      // The border style should be the thick one used for bundles.
      const expectedBorderStyle = SpreadsheetApp.BorderStyle.SOLID_THICK;
      
      if (topBorder) {
        _assertEqual(topBorder.getBorderStyle(), expectedBorderStyle, `${testName} - Top border should be SOLID_THICK.`);
      }
      if (bottomBorder) {
        _assertEqual(bottomBorder.getBorderStyle(), expectedBorderStyle, `${testName} - Bottom border should be SOLID_THICK.`);
      }
    });
  });
}


// =================================================================
// --- INTEGRATION TESTS ---
// =================================================================

/**
 * Integration test to verify that the entire row is colored based on status
 * and that data validation and number formatting rules are applied correctly.
 */
function test_applyUxRules_modes() {
  const testName = "Integration Test: applyUxRules Modes";

  withTestConfig(function () {
    setUp();
    withTestSheet(MOCK_DATA_INTEGRATION.csvForUxControlTests, function (sheet) {
      const startRow = CONFIG.approvalWorkflow.startDataRow;

      const statusCol = CONFIG.approvalWorkflow.columnIndices.status;
      const approverActionCol = CONFIG.approvalWorkflow.columnIndices.approverAction;

      // --- SCENARIO 1: Test applyUxRules(true) - Full UX Reset ---
      Log.TestDebug_gs(`[${testName}] SCENARIO 1: Testing applyUxRules(true)`);
      applyUxRules(true);
      SpreadsheetApp.flush();

      // --- Verification 1A: Check conditional formatting ---
      const colors = CONFIG.conditionalFormatColors;
      const defaultColor = "#FFFFFF";

      function _verifyRowColor(rowIndex, expectedColor, statusName) {
        const cell1 = sheet.getRange(rowIndex, statusCol);
        const cell2 = sheet.getRange(rowIndex, approverActionCol);
        _assertEqual(cell1.getBackground().toUpperCase(), expectedColor.toUpperCase(), `${testName} - BG Color for '${statusName}'`);
      }

      _verifyRowColor(startRow, colors.draft.background, "Draft");
      _verifyRowColor(startRow + 1, colors.pending.background, "Pending");
      _verifyRowColor(startRow + 2, colors.approved.background, "Approved");
      _verifyRowColor(startRow + 3, defaultColor, "No Status");
      _verifyRowColor(startRow + 4, colors.rejected.background, "Rejected");
      _verifyRowColor(startRow + 5, colors.pending.background, "Pending");

      // --- Verification 1B: Check Data Validation ---
      const dropdownCell = sheet.getRange(startRow, approverActionCol);
      _assertNotNull(dropdownCell.getDataValidation(), `${testName} (true) - Dropdown validation should exist.`);

      // --- Verification 1C: Check Number Formatting (German locale by default) ---
      const formats = CONFIG.numberFormats;
      const lrfCell = sheet.getRange(startRow, CONFIG.approvalWorkflow.columnIndices.lrfPreview);
      const priceCell = sheet.getRange(startRow, CONFIG.approvalWorkflow.columnIndices.aeSalesAskPrice);
      _assertEqual(lrfCell.getNumberFormat(), formats.percentage, `${testName} - LRF should have the correct percentage format.`);
      _assertEqual(priceCell.getNumberFormat(), formats.currency, `${testName} - Price should have the correct currency format.`);


      // --- SCENARIO 2: Test applyUxRules(false) - Validation Only ---
      Log.TestDebug_gs(`[${testName}] SCENARIO 2: Testing applyUxRules(false)`);

      dropdownCell.clearDataValidations();
      SpreadsheetApp.flush();
      _assertEqual(dropdownCell.getDataValidation(), null, `${testName} - Validation should be null after clearing.`);

      applyUxRules(false);
      SpreadsheetApp.flush();

      _assertNotNull(dropdownCell.getDataValidation(), `${testName} (false) - Dropdown validation should be restored.`);
    });
    tearDown();
  });
}