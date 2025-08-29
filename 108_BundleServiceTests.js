/**
 * @file This file contains the test suite for functions in BundleService.gs.
 * These are integration tests that operate on a live temporary sheet.
 */

/**
 * Runs all INTEGRATION tests for BundleService.gs.
 */
function runBundleService_IntegrationTests() {
    Log.TestResults_gs("--- Starting BundleService Integration Test Suite ---");

    setUp();
    test_validateBundle_Standalone_Integration();
    test_handleSheetAutomations_BundleValidation_Integration();
    test_groupApprovedItems_Integration(); 
    tearDown(); 

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

function test_handleSheetAutomations_BundleValidation_Integration() {
    const testName = "Integration Test: handleSheetAutomations Bundle Validation";

    withTestConfig(function() {
      withTestSheet(MOCK_DATA_INTEGRATION.csvForBundleValidationTests, function(sheet) {
          const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };

          // --- SCENARIO 1: An invalid edit (mismatched terms) should be reverted ---
          const invalidCell = sheet.getRange(13, c.bundleNumber); 
          const oldValueInvalid = invalidCell.getValue();
          const newValueInvalid = 606; 
          
          invalidCell.setValue(newValueInvalid);
          const mockEventInvalid = { range: invalidCell, value: newValueInvalid, oldValue: oldValueInvalid };
          handleSheetAutomations(mockEventInvalid);
          
          _assertNotNull(TestMocks.MOCK_TOAST_MESSAGE, `${testName} - An alert should be shown for an invalid bundle.`);
          _assertTrue(TestMocks.MOCK_TOAST_MESSAGE.includes("must have the same Quantity and Term"), `${testName} - The alert message should be correct.`);
          _assertEqual(invalidCell.getValue(), oldValueInvalid, `${testName} - The invalid change should be reverted in the sheet.`);

          // --- SCENARIO 2: A valid edit should be kept ---
          sheet.getRange(3, c.aeQuantity).setValue(10);
          sheet.getRange(3, c.aeTerm).setValue(24);
          SpreadsheetApp.flush();

          const validCell1 = sheet.getRange(2, c.bundleNumber);
          const oldValue1 = validCell1.getValue();
          const newValueValid = 888;
          TestMocks.MOCK_TOAST_MESSAGE = null; 

          validCell1.setValue(newValueValid);
          const mockEvent1 = { range: validCell1, value: newValueValid, oldValue: oldValue1 };
          handleSheetAutomations(mockEvent1);
          
          _assertEqual(TestMocks.MOCK_TOAST_MESSAGE, null, `${testName} - No alert should be shown for the first item of a valid bundle.`);
          _assertEqual(validCell1.getValue(), newValueValid, `${testName} - The first valid change should be kept.`);

          const validCell2 = sheet.getRange(3, c.bundleNumber);
          const oldValue2 = validCell2.getValue();
          validCell2.setValue(newValueValid);
          const mockEvent2 = { range: validCell2, value: newValueValid, oldValue: oldValue2 };
          handleSheetAutomations(mockEvent2);

          _assertEqual(TestMocks.MOCK_TOAST_MESSAGE, null, `${testName} - No alert should be shown for the second item of a valid bundle.`);
          _assertEqual(validCell2.getValue(), newValueValid, `${testName} - The second valid change should be kept.`);
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