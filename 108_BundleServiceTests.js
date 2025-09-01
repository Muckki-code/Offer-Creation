/**
 * @file This file contains the test suite for functions in BundleService.gs.
 * These are integration tests that operate on a live temporary sheet.
 */

// --- Test-Suite Specific Mock Setup ---
var _originalShowBundleMismatchDialog;
var _showBundleMismatchDialogCalled = false;

function _setUpBundleServiceTests() {
    _originalShowBundleMismatchDialog = showBundleMismatchDialog;
    _showBundleMismatchDialogCalled = false; // Reset before each run

    // Replace the real function with a mock that just sets a flag
    showBundleMismatchDialog = function(rowNum, bundleNumber, currentValues, expectedValues) {
        _showBundleMismatchDialogCalled = true;
        // Log for debugging test runs
        Log.TestDebug_gs(`[MOCK] showBundleMismatchDialog was called with: rowNum=${rowNum}, bundleNumber=${bundleNumber}`);
    };
    
    setUp(); // Call the general setUp from 200_Tests.js
}

function _tearDownBundleServiceTests() {
    // Restore the original function
    if (_originalShowBundleMismatchDialog) {
        showBundleMismatchDialog = _originalShowBundleMismatchDialog;
    }
    tearDown(); // Call the general tearDown
}


/**
 * Runs all INTEGRATION tests for BundleService.gs.
 */
function runBundleService_IntegrationTests() {
    Log.TestResults_gs("--- Starting BundleService Integration Test Suite ---");

    _setUpBundleServiceTests(); // Use our specific setup

    test_validateBundle_Standalone_Integration();
    test_handleSheetAutomations_BundleValidation_Integration(); // REFACTORED TEST
    test_groupApprovedItems_Integration(); 
    test_findAllBundleErrors_Integration(); // <-- ADD THIS LINE

    _tearDownBundleServiceTests(); // Use our specific teardown

    Log.TestResults_gs("--- BundleService Integration Test Suite Finished ---");
}




// =================================================================
// --- INTEGRATION TESTS ---
// =================================================================

function test_validateBundle_Standalone_Integration() {
    const testName = "Integration Test: validateBundle Scenarios (Standalone)";

    withTestConfig(function() {
      withTestSheet(MOCK_DATA_INTEGRATION.csvForBundleValidationTests, function(sheet) {
          let result;
          // Note: withTestConfig sets startRow to 2. The rows in the CSV are 1-indexed from the start of the data.
          // Row 2 = first data row.
          result = validateBundle(sheet, 2, 101); // Row 2 of sheet is first data row
          _assertTrue(result.isValid, `${testName} - A valid 2-item bundle should pass validation.`);
          
          result = validateBundle(sheet, 5, 202); // Row 5 of sheet is fourth data row
          _assertTrue(result.isValid, `${testName} - A valid 3-item bundle should pass validation.`);

          result = validateBundle(sheet, 8, 404);
          _assertEqual(result.isValid, false, `${testName} - A non-consecutive bundle should fail validation.`);

          result = validateBundle(sheet, 11, 505);
          _assertEqual(result.isValid, false, `${testName} - A bundle with mismatched quantity should fail validation.`);

          result = validateBundle(sheet, 13, 606);
          _assertEqual(result.isValid, false, `${testName} - A bundle with mismatched term should fail validation.`);
      });
    });
}

/**
 * --- REFACTORED TEST ---
 * This test now verifies the NEW non-blocking UI behavior.
 */
function test_handleSheetAutomations_BundleValidation_Integration() {
    const testName = "Integration Test: handleSheetAutomations Bundle Validation (Non-Blocking UI)";

    withTestConfig(function() {
      withTestSheet(MOCK_DATA_INTEGRATION.csvForBundleValidationTests, function(sheet) {
          const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };

          // --- SCENARIO 1: An invalid edit (mismatched terms) should be KEPT and trigger a dialog ---
          const invalidCell = sheet.getRange(13, c.aeTerm); // Editing Term on "Mismatch Term A"
          const oldValueInvalid = invalidCell.getValue(); // This is 24
          const newValueInvalid = 99; // A new term that causes the mismatch
          
          _showBundleMismatchDialogCalled = false; // Reset our mock flag before the action

          invalidCell.setValue(newValueInvalid);
          SpreadsheetApp.flush(); // Ensure the value is written before the event handler runs
          const mockEventInvalid = { range: invalidCell, value: newValueInvalid, oldValue: oldValueInvalid };
          handleSheetAutomations(mockEventInvalid);
          
          // --- VERIFY NEW BEHAVIOR ---
          _assertEqual(invalidCell.getValue(), newValueInvalid, `${testName} - The invalid user edit should be KEPT in the sheet.`);
          _assertTrue(_showBundleMismatchDialogCalled, `${testName} - showBundleMismatchDialog SHOULD be called for an invalid bundle edit.`);
          

          // --- SCENARIO 2: A valid edit should be kept and NOT trigger a dialog ---
          // Make the bundle valid first by aligning quantity and term
          sheet.getRange(3, c.aeQuantity).setValue(10); 
          sheet.getRange(3, c.aeTerm).setValue(24);
          SpreadsheetApp.flush();

          const validCell = sheet.getRange(2, c.bundleNumber);
          const oldValidValue = validCell.getValue();
          const newValidValue = 888;
          
          _showBundleMismatchDialogCalled = false; // Reset mock flag

          validCell.setValue(newValidValue);
          const mockEventValid = { range: validCell, value: newValidValue, oldValue: oldValidValue };
          handleSheetAutomations(mockEventValid);
          
          _assertEqual(validCell.getValue(), newValidValue, `${testName} - A valid change should be kept.`);
          _assertEqual(_showBundleMismatchDialogCalled, false, `${testName} - showBundleMismatchDialog should NOT be called for a valid edit.`);
      });
    });
}


function test_groupApprovedItems_Integration() {
    const testName = "Integration Test: groupApprovedItems Logic (Refactored)";

    withTestConfig(function() {
      withTestSheet(MOCK_DATA_INTEGRATION.groupingTestsAsArray, function(sheet) {
          const startCol = CONFIG.documentDeviceData.columnIndices.sku;
          const startRow = CONFIG.approvalWorkflow.startDataRow; // Will be 2 from withTestConfig
          const lastRow = sheet.getLastRow();
          const numCols = CONFIG.maxDataColumn - startCol + 1;
          
          const allDataRows = sheet.getRange(startRow, startCol, lastRow - startRow + 1, numCols).getValues();
          const result = groupApprovedItems(allDataRows, startCol);

          _assertEqual(result.length, 3, `${testName} - Should return 3 renderable items (2 individual, 1 bundle).`);

          const bundle = result.find(item => item.isBundle);
          _assertNotNull(bundle, `${testName} - The result should contain a consolidated bundle object.`);
          
          if (bundle) {
              _assertEqual(bundle.isBundle, true, `${testName} - Bundle flag should be true.`);
              _assertEqual(bundle.quantity, 10, `${testName} - Bundle quantity should be correct.`);
              _assertEqual(bundle.term, 24, `${testName} - Bundle term should be correct.`);
              _assertWithinTolerance(bundle.totalNetMonthlyPrice, 55.50, 0.001, `${testName} - Bundle price should be the correct sum.`);
              const expectedModelString = "Complete Bundle B (Pricier),\nComplete Bundle A (Cheaper)";
              _assertEqual(bundle.models, expectedModelString, `${testName} - Models should be sorted by price and newline-separated.`);
          }

          const individual = result.find(item => !item.isBundle && item.row[CONFIG.documentDeviceData.columnIndices.index - startCol] == 1);
           _assertNotNull(individual, `${testName} - The result should contain individual item with index 1.`);
           if(individual) {
                _assertEqual(individual.row[CONFIG.documentDeviceData.columnIndices.model - startCol], "Individual Approved A", `${testName} - Individual item should have correct model name.`);
           }
      });
    });
}

/**
 * --- NEW TEST ---
 * Integration test to verify that the new findAllBundleErrors function
 * correctly scans the sheet and returns all existing bundle errors.
 */
function test_findAllBundleErrors_Integration() {
    const testName = "Integration Test: findAllBundleErrors (Proactive Scan)";

    withTestConfig(function() {
      // This specific CSV from our mock data contains 3 invalid bundles:
      // - Bundle 404 has a gap.
      // - Bundle 505 has a quantity mismatch.
      // - Bundle 606 has a term mismatch.
      withTestSheet(MOCK_DATA_INTEGRATION.csvForBundleValidationTests, function(sheet) {
          
          // --- EXECUTE ---
          const errors = findAllBundleErrors();

          // --- VERIFY ---
          _assertNotNull(errors, `${testName} - The function should return an array.`);
          _assertEqual(errors.length, 3, `${testName} - Should find exactly 3 errors in the mock data.`);

          // Verify the GAP error
          const gapError = errors.find(e => e.bundleNumber == "404");
          _assertNotNull(gapError, `${testName} - An error for bundle #404 should be found.`);
          if (gapError) {
            _assertEqual(gapError.errorCode, 'GAP_DETECTED', `${testName} - Error code for #404 should be GAP_DETECTED.`);
          }

          // Verify the QUANTITY mismatch error
          const quantityError = errors.find(e => e.bundleNumber == "505");
          _assertNotNull(quantityError, `${testName} - An error for bundle #505 should be found.`);
          if (quantityError) {
            _assertEqual(quantityError.errorCode, 'MISMATCH', `${testName} - Error code for #505 should be MISMATCH.`);
          }

          // Verify the TERM mismatch error
          const termError = errors.find(e => e.bundleNumber == "606");
          _assertNotNull(termError, `${testName} - An error for bundle #606 should be found.`);
          if (termError) {
            _assertEqual(termError.errorCode, 'MISMATCH', `${testName} - Error code for #606 should be MISMATCH.`);
          }
      });
    });
}
