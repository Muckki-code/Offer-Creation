/**
 * @file This file contains the integration test suite for the document
 * generation process, verifying the final output in a Google Doc.
 */

/**
 * Runs all INTEGRATION tests for DocGenerator.gs.
 */
function runDocGenerator_IntegrationTests() {
    Log.TestResults_gs("--- Starting DocGenerator Integration Test Suite ---");

    setUp();
    test_documentGeneration_withBundle_Integration();
    tearDown();

    Log.TestResults_gs("--- DocGenerator Integration Test Suite Finished ---");
}


// =================================================================
// --- INTEGRATION TESTS ---
// =================================================================

/**
 * A full, end-to-end integration test that simulates the user action of
 * creating a document and verifies the content of the generated Google Doc,
 * paying special attention to how a consolidated bundle is rendered.
 */
function test_documentGeneration_withBundle_Integration() {
    const testName = "End-to-End Test: Document Generation with a Bundle (Refactored for Single Capex)";
    let generatedDocId = null;

    try {
        withTestConfig(function() {
            withTestSheet(MOCK_DATA_INTEGRATION.groupingTestsAsArray, function(sheet) {
                
                // --- SETUP ---
                sheet.getRange(CONFIG.offerDetailsCells.customerCompany).setValue("Bundle Test Corp");
                SpreadsheetApp.flush();

                const mockFormData = {
                    docLanguage: "german",
                    offerCreatedDate: "2025-07-28",
                    customDocName: `[TEST] Bundle Offer ${new Date().getTime()}`
                };

                // --- EXECUTE ---
                const dataPackage = prepareDocumentData(mockFormData);
                const newFileName = mockFormData.customDocName;
                const newDocFile = createDocument(newFileName, dataPackage.docLanguage);
                populateDocContent(newDocFile, dataPackage);

                // --- VERIFY ---
                const docUrl = newDocFile.getUrl();
                _assertNotNull(docUrl, `${testName} - A document URL should have been generated.`);
                if (!docUrl) return;

                const match = docUrl.match(/[-\w]{25,}/);
                generatedDocId = match ? match[0] : null;
                _assertNotNull(generatedDocId, `${testName} - Could not extract a valid Doc ID from the URL.`);
                if (!generatedDocId) return;

                const doc = DocumentApp.openById(generatedDocId);
                const table = doc.getBody().getTables()[0];
                _assertNotNull(table, `${testName} - The generated document should contain a table.`);
                
                Log.TestDebug_gs(`[${testName}] DIAGNOSTIC: Found ${doc.getBody().getTables().length} table(s). Table has ${table.getNumRows()} rows.`);

                const numRows = table.getNumRows();
                let bundleRowFound = false;

                for (let i = 1; i < numRows - 1; i++) {
                    const row = table.getRow(i);
                    const modelCellText = row.getCell(0).getText().trim();
                    
                    Log.TestDebug_gs(`[${testName}] DIAGNOSTIC: Reading Row ${i+1}, Model Cell Text: "${modelCellText}" (Length: ${modelCellText.length})`);

                    if (modelCellText.includes('\n')) {
                        bundleRowFound = true;
                        
                        const expectedModelText = "Complete Bundle B (Pricier),\nComplete Bundle A (Cheaper)";
                        _assertEqual(modelCellText, expectedModelText, `${testName} - Model cell should contain correct text with a newline.`);
                        
                        _assertEqual(row.getCell(1).getText(), "10", `${testName} - Quantity cell should be correct.`);
                        _assertEqual(row.getCell(2).getText(), "24 Monate", `${testName} - Term cell should be correctly formatted.`);
                        
                        // The final cell (4) shows the total price (unit price * quantity).
                        // From our updated 21-column mock data, the bundle unit price is 25.50 + 30.00 = 55.50.
                        // Total price = 55.50 * 10 = 555.00
                        const expectedPrice = formatNumberForLocale(555.00, "german", true);
                        _assertEqual(row.getCell(4).getText(), expectedPrice, `${testName} - Total price cell should be correctly formatted for German locale.`);

                        break;
                    }
                }
                _assertTrue(bundleRowFound, `${testName} - A row representing the bundle (containing a newline) should be found in the table.`);
            });
        });
    } finally {
        // --- CLEANUP ---
        if (generatedDocId) {
            try {
                DriveApp.getFileById(generatedDocId).setTrashed(true);
                Log.TestDebug_gs(`[${testName}] Successfully trashed test document: ${generatedDocId}`);
            } catch (e) {
                Log.TestDebug_gs(`[${testName}] WARNING: Could not trash test document ${generatedDocId}. It may need to be deleted manually. Error: ${e.message}`);
            }
        }
    }
}