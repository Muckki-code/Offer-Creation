/**
 * @file This file contains the test suite for functions in SheetCoreAutomations.gs.
 * REFACTORED for the new 22-column layout.
 */

/**
 * Runs all UNIT tests for SheetCoreAutomations.gs.
 */
function runSheetCoreAutomations_UnitTests() {
    Log.TestResults_gs("--- Starting SheetCoreAutomations Unit Test Suite ---");
    setUp();
    test_getNumericValue();
    test_updateCalculationsForRow();
    tearDown();
    Log.TestResults_gs("--- SheetCoreAutomations Unit Test Suite Finished ---");
}

/**
 * Runs all INTEGRATION tests for SheetCoreAutomations.gs.
 */
function runSheetCoreAutomations_IntegrationTests() {
    Log.TestResults_gs("--- Starting SheetCoreAutomations Integration Test Suite ---");
    setUp();
    test_handleSheetAutomations_Initialization_Integration();
    test_handleSheetAutomations_AEDiscovery_Integration();
    test_recalculateAllRows_Integration();
    test_handleSheetAutomations_forcesBrokenBundleToDraft();
    tearDown();
    Log.TestResults_gs("--- SheetCoreAutomations Integration Test Suite Finished ---");
}


/**
 * Runs all SANITIZATION tests for SheetCoreAutomations.gs.
 */
function runSheetCoreAutomations_SanitizationTests() {
    Log.TestResults_gs("--- Starting SheetCoreAutomations Sanitization Test Suite ---");
    setUp();
    test_Scenario1_pasteFinalizedRowIntoBlank();
    test_Scenario2_editModelOnApprovedRow();
    test_Scenario3_editSkuOnApprovedRow();
    test_Scenario4_revertManualEditOnProtectedColumns();
    test_Scenario5_pasteOverlappingData();
    test_Scenario6_pasteSkuAndBqData();
    tearDown();
    Log.TestResults_gs("--- SheetCoreAutomations Sanitization Test Suite Finished ---");
}


// =================================================================
// --- UNIT TESTS ---
// =================================================================

function test_getNumericValue() {
    const testName = "Unit Test: getNumericValue";
    _assertEqual(getNumericValue(1234), 1234, `${testName} - Should handle positive integers`);
    _assertEqual(getNumericValue("1,234.56"), 1234.56, `${testName} - Should parse English-style string`);
    _assertEqual(getNumericValue(null), 0, `${testName} - Should return 0 for null input`);
}

// In SheetCoreAutomationsTests.gs

function test_updateCalculationsForRow() {
    const testName = "Unit Test: updateCalculationsForRow (Refactored for Single Capex)";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const approvalWorkflowConfig = CONFIG.approvalWorkflow;
    const startColForMockData = 1;

    // Create a mock sheet object to satisfy the function signature.
    const mockSheet = {
      getRange: () => ({
        setNumberFormat: () => {} // This is a no-op, which is perfect for a unit test.
      })
    };
    const mockRowNum = 2; // The actual number doesn't matter for this unit test.

    // Test Case 1: Standard calculation using AE Sales Ask Price
    let row1 = MOCK_DATA_UNIT.rows.get('standard');
    // Call with the new, correct signature (isTelekomDeal is now ignored by the function)
    updateCalculationsForRow(mockSheet, mockRowNum, row1, false, colIndexes, approvalWorkflowConfig, startColForMockData);
    // LRF = (Price * Term) / Capex = (100 * 12) / 1000 = 1.2
    _assertWithinTolerance(row1[colIndexes.lrfPreview - 1], 1.2, 0.001, `${testName} - Standard LRF`);
    // Contract Value = Price * Term * Quantity = 100 * 12 * 10 = 12000
    _assertEqual(row1[colIndexes.contractValuePreview - 1], 12000, `${testName} - Standard Contract Value`);

    // Test Case 2: Calculation uses Approver Price Proposal when available
    let row2 = MOCK_DATA_UNIT.rows.get('approverPrice');
    updateCalculationsForRow(mockSheet, mockRowNum, row2, false, colIndexes, approvalWorkflowConfig, startColForMockData);
    // LRF = (Approver Price * Term) / Capex = (96 * 12) / 1000 = 1.152
    _assertWithinTolerance(row2[colIndexes.lrfPreview - 1], 1.152, 0.001, `${testName} - LRF uses Approver Price`);

    // Test Case 3: Calculation uses Final Approved Price for approved rows
    let row3 = MOCK_DATA_UNIT.rows.get('approved');
    updateCalculationsForRow(mockSheet, mockRowNum, row3, false, colIndexes, approvalWorkflowConfig, startColForMockData);
    // LRF = (Final Price * Term) / Capex = (90 * 12) / 1000 = 1.08
    _assertWithinTolerance(row3[colIndexes.lrfPreview - 1], 1.08, 0.001, `${testName} - LRF uses Final Approved Price`);
}


// =================================================================
// --- INTEGRATION TESTS ---
// =================================================================

function test_handleSheetAutomations_Initialization_Integration() {
  const testName = "Integration Test: New Row Initialization (Refactored)";
  withTestConfig(function () {
    withTestSheet(MOCK_DATA_INTEGRATION.csvForSanitizationTests, function (sheet) {
      const indexCol = CONFIG.documentDeviceData.columnIndices.index;
      const actionCol = CONFIG.approvalWorkflow.columnIndices.approverAction;
      const modelCol = CONFIG.documentDeviceData.columnIndices.model;
      
      const blankRow = 4;
      const modelCell = sheet.getRange(blankRow, modelCol);
      const modelOldValue = modelCell.getValue();
      const modelNewValue = "Brand New Model";
      modelCell.setValue(modelNewValue);
      let mockEvent = { range: modelCell, value: modelNewValue, oldValue: modelOldValue };

      handleSheetAutomations(mockEvent);

      const newIndex = sheet.getRange(blankRow, indexCol).getValue();
      const newAction = sheet.getRange(blankRow, actionCol).getValue();
      _assertNotEqual(newIndex, "", `${testName} - Blank Row: Should assign a new index.`);
      _assertEqual(newAction, "Choose Action", `${testName} - Blank Row: Should set default Approver Action.`);
    });
  });
}

// --- RESTORED TEST (NOW CORRECTED AND REFACTORED) ---
function test_handleSheetAutomations_AEDiscovery_Integration() {
  const testName = "Integration Test: AE data entry (Refactored for Single Capex)";
  // MODIFIED: This is now a 21-column array, reflecting the new structure.
  const dataArray = [
    ["SKU","EP CAPEX","TK CAPEX","Target","Limit","Index","Bundle Number","Model","AE CAPEX","AE SALES ASK","Qty","Term","Action","Comments","Approver Price","LRF","Contract Value","Status","Finance Price","Approved By","Approval Date"],
    ["SKU-003","","","","","", "","Device C", "", 50, 10, 24, "", "", "", "", "", "Draft", "", "", ""]
  ];

  withTestConfig(function () {
    withTestSheet(dataArray, function (sheet) {
      const targetRow = 2;
      const c = CONFIG.approvalWorkflow.columnIndices;

      // Simulate the final edit that makes the row complete (filling in the single Capex)
      const capexCell = sheet.getRange(targetRow, c.aeCapex);
      const oldValue = capexCell.getValue();
      const newValue = 1000;
      capexCell.setValue(newValue);
      const mockEvent = { range: capexCell, value: newValue, oldValue: oldValue };
      
      handleSheetAutomations(mockEvent);

      const finalStatus = sheet.getRange(targetRow, c.status).getValue();
      const finalLrf = sheet.getRange(targetRow, c.lrfPreview).getValue();
      const finalContractValue = sheet.getRange(targetRow, c.contractValuePreview).getValue();
      
      _assertEqual(finalStatus, CONFIG.approvalWorkflow.statusStrings.pending, `${testName} should set status to Pending Approval`);
      // LRF = (50 * 24) / 1000 = 1.2
      _assertWithinTolerance(finalLrf, 1.2, 0.001, `${testName} should correctly calculate LRF`);
      // Contract Value = 50 * 24 * 10 = 12000
      _assertEqual(finalContractValue, 12000, `${testName} should correctly calculate Contract Value`);
    });
  });
}


function test_recalculateAllRows_Integration() {
    const testName = "Integration Test: recalculateAllRows (Refactored for Single Capex)";
    withTestConfig(function () {
        // MODIFIED: CSV data is now in the 21-column format.
        const csvData = `SKU,EP CAPEX,Telekom CAPEX,Target,Limit,Index,Bundle Number,Device,AE CAPEX,AE SALES ASK,QUANTITY,TERM,APPROVER_ACTION,APPROVER_COMMENTS,APPROVER_PRICE_PROPOSAL,LRF_PREVIEW,CONTRACT_VALUE,STATUS,FINANCE_APPROVED_PRICE,APPROVED_BY,APPROVAL_DATE
SKU-100,"","","","",1,,"Device R1",1000,"110","10","24","","","","","",Pending Approval,"","",""
SKU-101,"","","","",,,"Device R2",800,"95","5","24","","","","","",Draft,"","",""
SKU-103,"","","","",4,,"Device R4",1500,"150","2","36","Approve Original Price","","","","",Approved (Original Price),"150","approver@test.com","2025-07-11"`;

        withTestSheet(csvData, function (sheet) {
            const statusCol = CONFIG.approvalWorkflow.columnIndices.status;
            const indexCol = CONFIG.documentDeviceData.columnIndices.index;

            // ACTION: Manually change a key field on the approved row to make it 'dirty'.
            const salesAskCell = sheet.getRange(4, CONFIG.approvalWorkflow.columnIndices.aeSalesAskPrice);
            salesAskCell.setValue(155); // Change from 150 to 155
            SpreadsheetApp.flush();

            // Execute the function to repair the sheet state
            recalculateAllRows();

            // VERIFICATION
            const finalIndex_R2 = sheet.getRange(3, indexCol).getValue();
            const finalStatus_R3 = sheet.getRange(4, statusCol).getValue();

            // R2 was a Draft row with no index. recalculateAllRows should have assigned it the next available one.
            _assertEqual(finalIndex_R2, 5, `${testName} - R2: Should auto-assign the next available index`);
            // R3 was Approved, but we manually edited a key data field. recalculateAllRows should have reverted its status.
            _assertEqual(finalStatus_R3, CONFIG.approvalWorkflow.statusStrings.revisedByAE, `${testName} - R3: Status should revert to 'Revised by AE' after a data change.`);
        });
    });
}

// =================================================================
// --- SANITIZATION TEST SUITE (REVISED) ---
// =================================================================

function test_Scenario1_pasteFinalizedRowIntoBlank() {
  const testName = "Sanitization S1: Paste finalized row into blank row";
  withTestConfig(function () {
    withTestSheet(MOCK_DATA_INTEGRATION.csvForSanitizationTests, function (sheet) {
      const sourceRow = 2;
      const targetRow = 4;
      const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };
      const sourceRange = sheet.getRange(sourceRow, 1, 1, CONFIG.maxDataColumn);
      const targetRange = sheet.getRange(targetRow, 1, 1, CONFIG.maxDataColumn);
      sourceRange.copyTo(targetRange);
      const mockEvent = { range: targetRange };
      handleSheetAutomations(mockEvent);
      const finalStatus = sheet.getRange(targetRow, c.status).getValue();
      const finalIndex = sheet.getRange(targetRow, c.index).getValue();
      _assertEqual(finalStatus, CONFIG.approvalWorkflow.statusStrings.pending, `${testName} - Status becomes Pending Approval`);
      _assertNotEqual(finalIndex, "", `${testName} - New Index is generated`);
    });
  });
}

function test_Scenario2_editModelOnApprovedRow() {
    const testName = "Sanitization S2: Edit Model on an Approved row";
    withTestConfig(function () {
        withTestSheet(MOCK_DATA_INTEGRATION.csvForSanitizationTests, function (sheet) {
            const targetRow = 2;
            const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };
            const modelCell = sheet.getRange(targetRow, c.model);
            const oldValue = modelCell.getValue();
            const newValue = "A Brand New Model Name";
            modelCell.setValue(newValue);
            const mockEvent = { range: modelCell, value: newValue, oldValue: oldValue };
            handleSheetAutomations(mockEvent);
            const finalStatus = sheet.getRange(targetRow, c.status).getValue();
            const bqDataCell = sheet.getRange(targetRow, c.epCapexRaw).getValue();
            _assertEqual(finalStatus, CONFIG.approvalWorkflow.statusStrings.revisedByAE, `${testName} - Status becomes Revised by AE`);
            _assertEqual(bqDataCell, "", `${testName} - BQ Data is blanked`);
        });
    });
}

function test_Scenario3_editSkuOnApprovedRow() {
    const testName = "Sanitization S3: Edit SKU on an Approved row";
    withTestConfig(function () {
        withTestSheet(MOCK_DATA_INTEGRATION.csvForSanitizationTests, function (sheet) {
            const targetRow = 2;
            const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };
            const skuCell = sheet.getRange(targetRow, c.sku);
            const oldValue = skuCell.getValue();
            const newValue = "NEW-SKU-999";
            skuCell.setValue(newValue);
            const mockEvent = { range: skuCell, value: newValue, oldValue: oldValue };
            handleSheetAutomations(mockEvent);
            const finalStatus = sheet.getRange(targetRow, c.status).getValue();
            const bqDataCell = sheet.getRange(targetRow, c.epCapexRaw).getValue();
            _assertEqual(finalStatus, CONFIG.approvalWorkflow.statusStrings.approvedOriginal, `${testName} - Status remains Approved`);
            _assertEqual(bqDataCell, "", `${testName} - BQ Data is blanked`);
        });
    });
}

function test_Scenario4_revertManualEditOnProtectedColumns() {
    const testName = "Sanitization S4: Revert edits on strictly protected columns";
    const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };
    const protectedColsToTest = [ c.epCapexRaw, c.tkCapexRaw, c.lrfPreview, c.status, c.financeApprovedPrice, c.approvedBy, c.approvalDate ];
    withTestConfig(function() {
        withTestSheet(MOCK_DATA_INTEGRATION.csvForSanitizationTests, function(sheet) {
            protectedColsToTest.forEach(colIndex => {
                const targetRow = 5;
                const cell = sheet.getRange(targetRow, colIndex);
                const originalValue = cell.getValue();
                const newValue = "ILLEGAL EDIT";
                cell.setValue(newValue);
                const mockEvent = { range: cell, value: newValue, oldValue: originalValue };
                handleSheetAutomations(mockEvent);
                const finalValue = sheet.getRange(targetRow, colIndex).getValue();
                const isCalculatedCol = (colIndex === c.lrfPreview || colIndex === c.contractValuePreview);
                if (!isCalculatedCol) {
                    _assertEqual(String(finalValue), String(originalValue), `${testName} - Edit to column ${colIndex} should be reverted`);
                }
                cell.setValue(originalValue);
                SpreadsheetApp.flush();
            });
        });
    });
}

// In SheetCoreAutomationsTests.gs

function test_Scenario5_pasteOverlappingData() {
  const testName = "Sanitization S5: Paste overlapping Index/Model and AE Data (F-M)";
  withTestConfig(function () {
    withTestSheet(MOCK_DATA_INTEGRATION.csvForSanitizationTests, function (sheet) {
      const targetRow = 3;
      const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };
      const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
      const numColsInDataBlock = CONFIG.maxDataColumn - dataBlockStartCol + 1;
      
      // We are simulating a paste of columns F-M from row 6 onto row 3
      const originalTargetRowValues = sheet.getRange(targetRow, dataBlockStartCol, 1, numColsInDataBlock).getValues()[0];
      const sourceColStartForPaste = c.index;
      const sourceColEndForPaste = c.aeTerm;
      const numColsInPaste = sourceColEndForPaste - sourceColStartForPaste + 1;
      const sourceRowForPaste = 6;
      const pastedValues = sheet.getRange(sourceRowForPaste, sourceColStartForPaste, 1, numColsInPaste).getValues()[0];
      
      let postPasteRowInMemory = JSON.parse(JSON.stringify(originalTargetRowValues));
      for (let i = 0; i < pastedValues.length; i++) {
        let targetColIndexInSheet = sourceColStartForPaste + i;
        let targetIndexInArray = targetColIndexInSheet - dataBlockStartCol; 
        postPasteRowInMemory[targetIndexInArray] = pastedValues[i];
      }
      
      const targetRange = sheet.getRange(targetRow, sourceColStartForPaste, 1, numColsInPaste);
      const mockEvent = { range: targetRange }; 
      
      handleSheetAutomations(mockEvent, [postPasteRowInMemory]);
      
      const finalData = sheet.getRange(targetRow, dataBlockStartCol, 1, numColsInDataBlock).getValues()[0];

      // --- VERIFY CORE PASTE ---
      _assertNotEqual(finalData[c.index - dataBlockStartCol], "", `${testName} - A new Index should be generated.`);
      _assertEqual(finalData[c.model - dataBlockStartCol], "Pasted Model For S5", `${testName} - Model accepts pasted value.`);
      _assertEqual(finalData[c.aeSalesAskPrice - dataBlockStartCol], 35, `${testName} - AE Data (Sales Ask Price) accepts pasted value.`);
      _assertEqual(finalData[c.status - dataBlockStartCol], CONFIG.approvalWorkflow.statusStrings.pending, `${testName} - Status should become Pending Approval after paste.`);
      
      // --- NEW, CRUCIAL ASSERTION ---
      _assertEqual(finalData[c.epCapexRaw - dataBlockStartCol], "", `${testName} - BQ Data (EP Capex) must be wiped after a desynchronizing paste.`);
    });
  });
}


function test_Scenario6_pasteSkuAndBqData() {
  const testName = "Sanitization S6: Paste overlapping SKU and BQ Data (A-E)";
  withTestConfig(function () {
    withTestSheet(MOCK_DATA_INTEGRATION.csvForSanitizationTests, function (sheet) {
      const sourceRow = 2;
      const targetRow = 5;
      const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };
      const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
      const numColsInDataBlock = CONFIG.maxDataColumn - dataBlockStartCol + 1;
      const sourceColStartForPaste = c.sku;
      const sourceColEndForPaste = c.rentalLimitRaw;
      const numColsInPaste = sourceColEndForPaste - sourceColStartForPaste + 1;
      const sourceRange = sheet.getRange(sourceRow, sourceColStartForPaste, 1, numColsInPaste);
      const targetRange = sheet.getRange(targetRow, sourceColStartForPaste, 1, numColsInPaste);
      sourceRange.copyTo(targetRange);
      const mockEvent = { range: targetRange };
      handleSheetAutomations(mockEvent);
      const finalData = sheet.getRange(targetRow, dataBlockStartCol, 1, numColsInDataBlock).getValues()[0];
      _assertEqual(finalData[c.sku - dataBlockStartCol], "S1_SOURCE_APPROVED", `${testName} - SKU accepts pasted value`);
      _assertEqual(finalData[c.model - dataBlockStartCol], "Target Pending Model", `${testName} - Model remains unchanged`);
      _assertEqual(finalData[c.status - dataBlockStartCol], CONFIG.approvalWorkflow.statusStrings.pending, `${testName} - Status is correctly re-evaluated to Pending.`);
    });
  });
}

/**
 * --- NEW TEST FOR BUNDLE DRAFT RULE ---
 * Verifies that when a user edit breaks a bundle's integrity (e.g., mismatching term),
 * handleSheetAutomations forces ALL items in that bundle back to "Draft" status,
 * while critically PRESERVING the user's invalid edit.
 */
function test_handleSheetAutomations_forcesBrokenBundleToDraft() {
    const testName = "Integration Test: Broken bundle forces all items to Draft";

    // Data for a valid 2-item bundle, both pending approval.
    const csvData = `SKU,EP CAPEX,Telekom CAPEX,Target,Limit,Index,Bundle Number,Device,AE EP CAPEX,AE TK CAPEX,AE SALES ASK,QUANTITY,TERM,APPROVER_ACTION,APPROVER_COMMENTS,APPROVER_PRICE_PROPOSAL,LRF_PREVIEW,CONTRACT_VALUE,STATUS
,,,,,"1",707,"Bundle Item A",1000,1200,100,5,24,"Choose Action","","","","",Pending Approval
,,,,,"2",707,"Bundle Item B",1000,1200,100,5,24,"Choose Action","","","","",Pending Approval
`;

    withTestConfig(function () {
        withTestSheet(csvData, function (sheet) {
            const statusCol = CONFIG.approvalWorkflow.columnIndices.status;
            const termCol = CONFIG.approvalWorkflow.columnIndices.aeTerm;
            
            // --- ACTION: Edit the Term of the second item to break the bundle ---
            const targetRow = 3; // This is the second data row in the sheet
            const termCell = sheet.getRange(targetRow, termCol);
            const oldValue = termCell.getValue();
            const newValue = 99; // A new, invalid term
            
            termCell.setValue(newValue);
            const mockEvent = { range: termCell, value: newValue, oldValue: oldValue };
            
            handleSheetAutomations(mockEvent); // Execute the main automation

            // --- VERIFICATION ---
            const status_R1 = sheet.getRange(2, statusCol).getValue(); // First bundle item
            const status_R2 = sheet.getRange(3, statusCol).getValue(); // Second bundle item (the one edited)
            const finalTermValue = termCell.getValue();

            const expectedStatus = CONFIG.approvalWorkflow.statusStrings.draft;

            // --- NEW, CRUCIAL ASSERTION ---
            _assertEqual(finalTermValue, newValue, `${testName} - The user's invalid edit should be preserved.`);
            
            _assertEqual(status_R1, expectedStatus, `${testName} - The status of the FIRST item should be forced to Draft.`);
            _assertEqual(status_R2, expectedStatus, `${testName} - The status of the EDITED item should be forced to Draft.`);
        });
    });
}
