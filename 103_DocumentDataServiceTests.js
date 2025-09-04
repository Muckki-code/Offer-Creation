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
          
          // There should be one individual item (Lenovo) and one bundled item.
          _assertEqual(dataPackage.devicesData.length, 2, `${testName} - devicesData should contain 2 renderable items.`);
          
          // Verify the bundle
          const bundle = dataPackage.devicesData.find(d => d.bundleId === "101");
          _assertNotNull(bundle, `${testName} - A consolidated bundle (ID 101) should be present in devicesData.`);
          if (bundle) {
            _assertEqual(bundle.quantity, 5, `${testName} - Bundle quantity should be 5.`);
            _assertEqual(bundle.term, 24, `${testName} - Bundle term should be 24.`);
            // Net monthly price is the sum of the prices of the items in the bundle
            _assertWithinTolerance(bundle.netMonthlyRentalPrice, 138.48, 0.001, `${testName} - Bundle's net monthly price should be the sum (64.49 + 73.99).`);
            // Total net monthly price is the net monthly price * quantity
            _assertWithinTolerance(bundle.totalNetMonthlyRentalPrice, 692.40, 0.001, `${testName} - Bundle's total price should be price * quantity (138.48 * 5).`);
          }
          
          // Verify the individual item
          const individualItem = dataPackage.devicesData.find(d => d.model === "Lenovo ThinkPad T14 G5");
          _assertNotNull(individualItem, `${testName} - The individual Lenovo item should be present.`);
          if (individualItem) {
            _assertEqual(individualItem.quantity, 10, `${testName} - Individual item quantity should be 10.`);
            _assertWithinTolerance(individualItem.netMonthlyRentalPrice, 61.99, 0.001, `${testName} - Individual item price should be correct.`);
          }
          
          // Calculation based on updated MOCK_DATA_INTEGRATION.groupingTestsAsArray (23-column)
          // Lenovo: 61.99 * 10 = 619.90
          // Bundle 101 (iPhone 512GB + iPhone 1TB): (64.49 + 73.99) * 5 = 138.48 * 5 = 692.40
          // Total = 619.90 + 692.40 = 1312.30
          const expectedGrandTotal = 1312.30;
          _assertWithinTolerance(dataPackage.grandTotal, expectedGrandTotal, 0.001, `${testName} - The grand total should be calculated correctly across all items.`);
      });
    });
}