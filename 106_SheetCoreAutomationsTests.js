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
    const testName = "Unit Test: updateCalculationsForRow (Refactored)";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const approvalWorkflowConfig = CONFIG.approvalWorkflow;
    const startColForMockData = 1;

    // --- NEW: Create a mock sheet object to satisfy the new function signature ---
    const mockSheet = {
      getRange: () => ({
        setNumberFormat: () => {} // This is a no-op, it does nothing, which is perfect for a unit test.
      })
    };
    const mockRowNum = 2; // The actual number doesn't matter for this unit test.

    let row1 = MOCK_DATA_UNIT.rows.get('standard');
    // FIXED: Call with the new, correct signature
    updateCalculationsForRow(mockSheet, mockRowNum, row1, false, colIndexes, approvalWorkflowConfig, startColForMockData);
    _assertWithinTolerance(row1[colIndexes.lrfPreview - 1], 1.2, 0.001, `${testName} - Standard LRF`);
    _assertEqual(row1[colIndexes.contractValuePreview - 1], 12000, `${testName} - Standard Contract Value`);

    let row2 = MOCK_DATA_UNIT.rows.get('standard');
    // FIXED: Call with the new, correct signature
    updateCalculationsForRow(mockSheet, mockRowNum, row2, true, colIndexes, approvalWorkflowConfig, startColForMockData);
    _assertWithinTolerance(row2[colIndexes.lrfPreview - 1], 1.0, 0.001, `${testName} - Telekom Deal LRF`);

    let row4 = MOCK_DATA_UNIT.rows.get('approverPrice');
    // FIXED: Call with the new, correct signature
    updateCalculationsForRow(mockSheet, mockRowNum, row4, false, colIndexes, approvalWorkflowConfig, startColForMockData);
    _assertWithinTolerance(row4[colIndexes.lrfPreview - 1], 1.152, 0.001, `${testName} - LRF uses Approver Price`);

    let row5 = MOCK_DATA_UNIT.rows.get('approved');
    // FIXED: Call with the new, correct signature
    updateCalculationsForRow(mockSheet, mockRowNum, row5, false, colIndexes, approvalWorkflowConfig, startColForMockData);
    _assertWithinTolerance(row5[colIndexes.lrfPreview - 1], 1.08, 0.001, `${testName} - LRF uses Final Approved Price`);
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

// --- RESTORED TEST (NOW CORRECTED) ---
function test_handleSheetAutomations_AEDiscovery_Integration() {
  const testName = "Integration Test: AE data entry (Refactored)";
  // MODIFIED: This is now a full 22-column array, which is the most robust way to represent the data.
  const dataArray = [
    ["SKU","EP CAPEX","TK CAPEX","Target","Limit","Index","Bundle Number","Model","AE EP CAPEX","AE TK CAPEX","AE SALES ASK","Qty","Term","Action","Comments","Approver Price","LRF","Contract Value","Status","Finance Price","Approved By","Approval Date"],
    ["SKU-003","","","","","", "","Device C", 1000, "", 50, 10, 24, "", "", "", "", "", "Draft", "", "", ""]
  ];

  withTestConfig(function () {
    withTestSheet(dataArray, function (sheet) {
      const targetRow = 2;
      const c = CONFIG.approvalWorkflow.columnIndices;

      // Simulate the final edit that makes the row complete
      const tkCapexCell = sheet.getRange(targetRow, c.aeTkCapex);
      const oldValue = tkCapexCell.getValue();
      const newValue = 1200;
      tkCapexCell.setValue(newValue);
      const mockEvent = { range: tkCapexCell, value: newValue, oldValue: oldValue };
      
      handleSheetAutomations(mockEvent);

      const finalStatus = sheet.getRange(targetRow, c.status).getValue();
      const finalLrf = sheet.getRange(targetRow, c.lrfPreview).getValue();
      const finalContractValue = sheet.getRange(targetRow, c.contractValuePreview).getValue();
      
      _assertEqual(finalStatus, CONFIG.approvalWorkflow.statusStrings.pending, `${testName} should set status to Pending Approval`);
      _assertWithinTolerance(finalLrf, 1.2, 0.001, `${testName} should correctly calculate LRF`);
      _assertEqual(finalContractValue, 12000, `${testName} should correctly calculate Contract Value`);
    });
  });
}


function test_recalculateAllRows_Integration() {
    const testName = "Integration Test: recalculateAllRows (Refactored)";
    withTestConfig(function () {
        const csvData = `SKU,EP CAPEX,Telekom CAPEX,Target,Limit,Index,Bundle Number,Device,AE EP CAPEX,AE TK CAPEX,AE SALES ASK,QUANTITY,TERM,APPROVER_ACTION,APPROVER_COMMENTS,APPROVER_PRICE_PROPOSAL,LRF_PREVIEW,CONTRACT_VALUE,STATUS,FINANCE_APPROVED_PRICE,APPROVED_BY,APPROVAL_DATE
SKU-100,"","","","",1,,"Device R1","1000","1200","110","10","24","","","","","",Pending Approval,"","",""
SKU-101,"","","","",,,"Device R2","800","900","95","5","24","","","","","",Draft,"","",""
SKU-103,"","","","",4,,"Device R4","1500","1600","150","2","36","Approve Original Price","","","","",Approved (Original Price),"150","approver@test.com","2025-07-11"`;

        withTestSheet(csvData, function (sheet) {
            const lrfCol = CONFIG.approvalWorkflow.columnIndices.lrfPreview;
            const statusCol = CONFIG.approvalWorkflow.columnIndices.status;
            const indexCol = CONFIG.documentDeviceData.columnIndices.index;

            const telekomDealCell = sheet.getRange(CONFIG.offerDetailsCells.telekomDeal);
            telekomDealCell.setValue("Yes");
            const mockEvent = { range: telekomDealCell, value: "Yes", oldValue: "No" };
            
            handleSheetAutomations(mockEvent);

            const finalLrf_R1 = sheet.getRange(2, lrfCol).getValue();
            const finalIndex_R2 = sheet.getRange(3, indexCol).getValue();
            const finalStatus_R3 = sheet.getRange(4, statusCol).getValue();

            _assertWithinTolerance(finalLrf_R1, 2.2, 0.001, `${testName} - R1: LRF should change after switching to Telekom Deal`);
            _assertEqual(finalIndex_R2, 5, `${testName} - R2: Should auto-assign the next available index`);
            _assertEqual(finalStatus_R3, CONFIG.approvalWorkflow.statusStrings.revisedByAE, `${testName} - R3: Status should revert to 'Revised by AE'`);
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