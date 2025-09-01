/**
 * @file This file contains the test suite for functions in DocumentDataService.gs.
 */

/**
 * Runs all INTEGRATION tests for DocumentDataService.gs.
 */
function runDocumentDataService_IntegrationTests() {
    Log.TestResults_gs("--- Starting DocumentDataService Integration Test Suite ---");

    setUp();
    test_prepareDocumentData_Integration();
    tearDown(); 

    Log.TestResults_gs("--- DocumentDataService Integration Test Suite Finished ---");
}


// =================================================================
// --- INTEGRATION TESTS ---
// =================================================================

/**
 * Integration test for the main prepareDocumentData function.
 * This test verifies the entire data processing pipeline, including the call
 * to the BundleService and the final data package structure.
 * It uses the refactored 'groupingTestsAsArray' data.
 */
function test_prepareDocumentData_Integration() {
    const testName = "Integration Test: prepareDocumentData (Refactored for Single Capex)";

    withTestConfig(function() {
      withTestSheet(MOCK_DATA_INTEGRATION.groupingTestsAsArray, function(sheet) {
          
          // --- SETUP ---
          const mockFormData = {
            docLanguage: "german",
            offerCreatedDate: "2025-07-25",
          };
          
          sheet.getRange(CONFIG.offerDetailsCells.customerCompany).setValue("Test Customer GmbH");
          SpreadsheetApp.flush();


          // --- EXECUTE ---
          const dataPackage = prepareDocumentData(mockFormData);


          // --- VERIFY ---
          _assertNotNull(dataPackage, `${testName} - The function should return a data package object.`);
          
          _assertEqual(dataPackage.customerCompanyName, "Test Customer GmbH", `${testName} - Customer company name should be correct.`);
          _assertEqual(dataPackage.docLanguage, "german", `${testName} - Document language should be correct.`);
          
          _assertEqual(dataPackage.devicesData.length, 3, `${testName} - devicesData should contain 3 renderable items.`);
          
          const bundle = dataPackage.devicesData.find(d => d.model.includes("\n")); 
          _assertNotNull(bundle, `${testName} - A consolidated bundle should be present in devicesData.`);
          if (bundle) {
            _assertEqual(bundle.quantity, 10, `${testName} - Bundle quantity should be correct.`);
            _assertEqual(bundle.term, 24, `${testName} - Bundle term should be correct.`);
            _assertWithinTolerance(bundle.netMonthlyRentalPrice, 55.50, 0.001, `${testName} - Bundle's net monthly price should be the sum.`);
            _assertWithinTolerance(bundle.totalNetMonthlyRentalPrice, 555.00, 0.001, `${testName} - Bundle's total price should be price * quantity.`);
          }
          
          const individualItem = dataPackage.devicesData.find(d => d.model === "Individual Approved A");
          _assertNotNull(individualItem, `${testName} - An individual approved item should be present.`);
          if (individualItem) {
            _assertEqual(individualItem.quantity, 1, `${testName} - Individual item quantity should be correct.`);
            _assertEqual(individualItem.netMonthlyRentalPrice, 50.00, `${testName} - Individual item price should be correct.`);
          }
          
          // Calculation based on updated MOCK_DATA_INTEGRATION.groupingTestsAsArray (21-column)
          // Bundle: (25.50 + 30.00) * 10 = 555.00
          // Individual A: 50.00 * 1 = 50.00
          // Individual B: 100.00 * 2 = 200.00
          // Total = 555 + 50 + 200 = 805.00
          const expectedGrandTotal = 805.00; 
          _assertWithinTolerance(dataPackage.grandTotal, expectedGrandTotal, 0.001, `${testName} - The grand total should be calculated correctly across all items.`);
      });
    });
}