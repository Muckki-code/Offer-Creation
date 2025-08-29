/**
 * @file This file contains the test suite for functions in UxControl.gs.
 */

/**
 * Runs all INTEGRATION tests for UxControl.gs.
 */
function runUxControl_IntegrationTests() {
  Log.TestResults_gs("--- Starting UxControl Integration Test Suite ---");

  test_applyUxRules_modes();

  Log.TestResults_gs("--- UxControl Integration Test Suite Finished ---");
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
      _verifyRowColor(startRow + 5, colors.pending.background, "Revised by AE");

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