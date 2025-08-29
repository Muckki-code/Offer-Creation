/**
 * @file This file contains the test suite for functions in ApprovalWorkflow.gs.
 * REFACTORED to test the new single-action approval workflow and updated health check.
 */

// --- Global Mocks Setup ---
var TestMocks = TestMocks || {};

// ... (Standard setUp and tearDown functions remain unchanged, they are robust) ...
if (!TestMocks.ORIGINAL_UI_SERVICE) { 
  TestMocks.ORIGINAL_UI_SERVICE = SpreadsheetApp.getUi();
  TestMocks.ORIGINAL_UI_BUTTON_SET = TestMocks.ORIGINAL_UI_SERVICE.ButtonSet;
  TestMocks.ORIGINAL_UI_BUTTON = TestMocks.ORIGINAL_UI_SERVICE.Button;
}
if (!TestMocks.ORIGINAL_SS_APP_GET_ACTIVE) {
  TestMocks.ORIGINAL_SS_APP_GET_ACTIVE = SpreadsheetApp.getActive;
}

var _originalLogTableActivity = typeof logTableActivity !== 'undefined' ? logTableActivity : null;
var _originalLogGeneralActivity = typeof logGeneralActivity !== 'undefined' ? logGeneralActivity : null;

TestMocks.MOCK_TOAST_MESSAGE = TestMocks.MOCK_TOAST_MESSAGE === undefined ? null : TestMocks.MOCK_TOAST_MESSAGE;
TestMocks.MOCK_SESSION_EMAIL = TestMocks.MOCK_SESSION_EMAIL === undefined ? "test.approver@example.com" : TestMocks.MOCK_SESSION_EMAIL;
TestMocks.LOG_TABLE_ACTIVITY_CALLS = TestMocks.LOG_TABLE_ACTIVITY_CALLS === undefined ? [] : TestMocks.LOG_TABLE_ACTIVITY_CALLS;
TestMocks.LOG_GENERAL_ACTIVITY_CALLS = TestMocks.LOG_GENERAL_ACTIVITY_CALLS === undefined ? [] : TestMocks.LOG_GENERAL_ACTIVITY_CALLS;
TestMocks.MOCK_SHEET = {
  getName: function() { return 'MockUnitTestSheet'; }
};

function setUp() {
  SpreadsheetApp.getActive = function() {
    return {
      toast: function(message, title, timeoutSeconds) {
        TestMocks.MOCK_TOAST_MESSAGE = message;
      }
    };
  };

  Session.getActiveUser = function() {
    return {
      getEmail: function() {
        return TestMocks.MOCK_SESSION_EMAIL;
      }
    };
  };

  logTableActivity = function(options) { TestMocks.LOG_TABLE_ACTIVITY_CALLS.push(options); };
  logGeneralActivity = function(options) { TestMocks.LOG_GENERAL_ACTIVITY_CALLS.push(options); };
  
  TestMocks.MOCK_TOAST_MESSAGE = null;
  TestMocks.MOCK_SESSION_EMAIL = "test.approver@example.com";
  TestMocks.LOG_TABLE_ACTIVITY_CALLS = []; 
  TestMocks.LOG_GENERAL_ACTIVITY_CALLS = []; 
}

function tearDown() {
  if (TestMocks.ORIGINAL_SS_APP_GET_ACTIVE) { SpreadsheetApp.getActive = TestMocks.ORIGINAL_SS_APP_GET_ACTIVE; }
  if (_originalLogTableActivity) logTableActivity = _originalLogTableActivity;
  if (_originalLogGeneralActivity) logGeneralActivity = _originalLogGeneralActivity;

  TestMocks.MOCK_TOAST_MESSAGE = null;
  TestMocks.MOCK_SESSION_EMAIL = "test.approver@example.com";
  TestMocks.LOG_TABLE_ACTIVITY_CALLS = [];
  TestMocks.LOG_GENERAL_ACTIVITY_CALLS = [];
}

// =================================================================
// --- TEST RUNNERS FOR THIS SUITE ---
// =================================================================

function runApprovalWorkflow_UnitTests() {
  Log.TestResults_gs("--- Starting ApprovalWorkflow Unit Test Suite (REFACTORED) ---");
  
  test_processSingleApprovalAction_approveOriginalPrice();
  test_processSingleApprovalAction_approveNewPrice();
  test_processSingleApprovalAction_rejectWithComment();
  test_processSingleApprovalAction_fail_InvalidStatus();
  test_processSingleApprovalAction_fail_ApproveOriginalNoPrice();
  test_processSingleApprovalAction_fail_ApproveNewNoPrice();
  test_processSingleApprovalAction_fail_RejectNoComment();
  test_processSingleApprovalAction_fail_InvalidLRF();

  Log.TestResults_gs("--- ApprovalWorkflow Unit Test Suite Finished ---");
}

function runApprovalWorkflow_IntegrationTests() {
  Log.TestResults_gs("--- Starting ApprovalWorkflow Integration Test Suite (REFACTORED) ---");
  
  test_handleSheetAutomations_directApprovalAction_Integration();
  test_runSheetHealthCheck_findsAndFixesInconsistencies();
  test_runSheetHealthCheck_noInconsistenciesFound();

  Log.TestResults_gs("--- ApprovalWorkflow Integration Test Suite Finished ---");
}

// =================================================================
// --- UNIT TEST HELPERS ---
// =================================================================

function _createMockEditEvent(row, column, value, oldValue) {
  return { range: { getRow: () => row, getColumn: () => column }, value: value, oldValue: oldValue };
}

function _getCleanRow(templateName) {
  const templateRow = MOCK_DATA_UNIT.rowsForApprovalTests.get(templateName);
  if (!templateRow) throw new Error(`Mock data template "${templateName}" not found.`);
  return JSON.parse(JSON.stringify(templateRow));
}

// =================================================================
// --- UNIT TESTS FOR processSingleApprovalAction() ---
// =================================================================

function test_processSingleApprovalAction_approveOriginalPrice() {
  setUp();
  const testName = "Unit Test: processSingleApprovalAction - Approve Original Price";
  const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
  const originalRow = _getCleanRow('pendingApprovedOriginal');
  const inMemoryRow = _getCleanRow('pendingApprovedOriginal');
  const mockEvent = _createMockEditEvent(10, colIndexes.approverAction, "Approve Original Price", "Choose Action");
  const result = processSingleApprovalAction(TestMocks.MOCK_SHEET, 10, mockEvent, inMemoryRow, colIndexes, originalRow, 1);
  _assertTrue(result, `${testName} - Should return true on success`);
  _assertEqual(inMemoryRow[colIndexes.status - 1], CONFIG.approvalWorkflow.statusStrings.approvedOriginal, `${testName} - Status should be Approved (Original Price)`);
  _assertEqual(inMemoryRow[colIndexes.financeApprovedPrice - 1], 100, `${testName} - Finance Approved Price should be original price`);
  _assertEqual(inMemoryRow[colIndexes.approvedBy - 1], TestMocks.MOCK_SESSION_EMAIL, `${testName} - Approved By should be set`);
  _assertNotNull(inMemoryRow[colIndexes.approvalDate - 1], `${testName} - Approval Date should be set`);
  tearDown();
}

function test_processSingleApprovalAction_approveNewPrice() {
  setUp();
  const testName = "Unit Test: processSingleApprovalAction - Approve New Price";
  const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
  const originalRow = _getCleanRow('pendingApprovedNew');
  const inMemoryRow = _getCleanRow('pendingApprovedNew');
  const mockEvent = _createMockEditEvent(10, colIndexes.approverAction, "Approve New Price", "Choose Action");
  const result = processSingleApprovalAction(TestMocks.MOCK_SHEET, 10, mockEvent, inMemoryRow, colIndexes, originalRow, 1);
  _assertTrue(result, `${testName} - Should return true on success`);
  _assertEqual(inMemoryRow[colIndexes.status - 1], CONFIG.approvalWorkflow.statusStrings.approvedNew, `${testName} - Status should be Approved (New Price)`);
  _assertEqual(inMemoryRow[colIndexes.financeApprovedPrice - 1], 90, `${testName} - Finance Approved Price should be proposed price`);
  tearDown();
}

function test_processSingleApprovalAction_rejectWithComment() {
  setUp();
  const testName = "Unit Test: processSingleApprovalAction - Reject with Comment";
  const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
  const originalRow = _getCleanRow('pendingRejected');
  const inMemoryRow = _getCleanRow('pendingRejected');
  const mockEvent = _createMockEditEvent(10, colIndexes.approverAction, "Reject with Comment", "Choose Action");
  const result = processSingleApprovalAction(TestMocks.MOCK_SHEET, 10, mockEvent, inMemoryRow, colIndexes, originalRow, 1);
  _assertTrue(result, `${testName} - Should return true on success`);
  _assertEqual(inMemoryRow[colIndexes.status - 1], CONFIG.approvalWorkflow.statusStrings.rejected, `${testName} - Status should be Rejected`);
  _assertEqual(inMemoryRow[colIndexes.financeApprovedPrice - 1], "", `${testName} - Finance Approved Price should remain empty`);
  tearDown();
}

function test_processSingleApprovalAction_fail_InvalidStatus() {
  setUp();
  const testName = "Unit Test: processSingleApprovalAction - Fail on Invalid Status";
  const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
  const originalRow = _getCleanRow('pendingInvalidStatus');
  const inMemoryRow = _getCleanRow('pendingInvalidStatus');
  const mockEvent = _createMockEditEvent(10, colIndexes.approverAction, "Approve Original Price", "Choose Action");
  const result = processSingleApprovalAction(TestMocks.MOCK_SHEET, 10, mockEvent, inMemoryRow, colIndexes, originalRow, 1);
  _assertEqual(result, false, `${testName} - Should return false on failure`);
  _assertEqual(inMemoryRow[colIndexes.status - 1], CONFIG.approvalWorkflow.statusStrings.draft, `${testName} - Status should remain Draft`);
  _assertEqual(inMemoryRow[colIndexes.approverAction - 1], "Choose Action", `${testName} - Action should be reverted`);
  _assertNotNull(TestMocks.MOCK_TOAST_MESSAGE, `${testName} - A toast message should be shown`);
  _assertTrue(TestMocks.MOCK_TOAST_MESSAGE.includes("cannot be processed"), `${testName} - Toast should explain the error`);
  tearDown();
}

function test_processSingleApprovalAction_fail_ApproveOriginalNoPrice() {
  setUp();
  const testName = "Unit Test: processSingleApprovalAction - Fail Approve Original Price with no price";
  const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
  const originalRow = _getCleanRow('pendingApproveNoPrice');
  const inMemoryRow = _getCleanRow('pendingApproveNoPrice');
  const mockEvent = _createMockEditEvent(10, colIndexes.approverAction, "Approve Original Price", "Choose Action");
  const result = processSingleApprovalAction(TestMocks.MOCK_SHEET, 10, mockEvent, inMemoryRow, colIndexes, originalRow, 1);
  _assertEqual(result, false, `${testName} - Should return false on failure`);
  _assertEqual(inMemoryRow[colIndexes.approverAction - 1], "Choose Action", `${testName} - Action should be reverted`);
  _assertTrue(TestMocks.MOCK_TOAST_MESSAGE.includes("invalid or missing LRF"), `${testName} - Toast should explain the error`);
  tearDown();
}

function test_processSingleApprovalAction_fail_ApproveNewNoPrice() {
    setUp();
    const testName = "Unit Test: processSingleApprovalAction - Fail Approve New Price with no proposal";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getCleanRow('pendingApprovedOriginal');
    originalRow[colIndexes.approverPriceProposal - 1] = ""; 
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));
    const mockEvent = _createMockEditEvent(10, colIndexes.approverAction, "Approve New Price", "Choose Action");
    const result = processSingleApprovalAction(TestMocks.MOCK_SHEET, 10, mockEvent, inMemoryRow, colIndexes, originalRow, 1);
    _assertEqual(result, false, `${testName} - Should return false on failure`);
    _assertEqual(inMemoryRow[colIndexes.approverAction - 1], "Choose Action", `${testName} - Action should be reverted`);
    _assertTrue(TestMocks.MOCK_TOAST_MESSAGE.includes("without a valid 'Approver Price Proposal'"), `${testName} - Toast should explain the error`);
    tearDown();
}

function test_processSingleApprovalAction_fail_RejectNoComment() {
  setUp();
  const testName = "Unit Test: processSingleApprovalAction - Fail Reject with no comment";
  const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
  const originalRow = _getCleanRow('pendingRejectNoComment');
  const inMemoryRow = _getCleanRow('pendingRejectNoComment');
  const mockEvent = _createMockEditEvent(10, colIndexes.approverAction, "Reject with Comment", "Choose Action");
  const result = processSingleApprovalAction(TestMocks.MOCK_SHEET, 10, mockEvent, inMemoryRow, colIndexes, originalRow, 1);
  _assertEqual(result, false, `${testName} - Should return false on failure`);
  _assertEqual(inMemoryRow[colIndexes.approverAction - 1], "Choose Action", `${testName} - Action should be reverted`);
  _assertTrue(TestMocks.MOCK_TOAST_MESSAGE.includes("without adding a comment"), `${testName} - Toast should explain the error`);
  tearDown();
}

function test_processSingleApprovalAction_fail_InvalidLRF() {
  setUp();
  const testName = "Unit Test: processSingleApprovalAction - Fail Approve with invalid LRF";
  const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
  const originalRow = _getCleanRow('pendingApprovedOriginal');
  originalRow[colIndexes.lrfPreview - 1] = "Invalid LRF"; 
  const inMemoryRow = JSON.parse(JSON.stringify(originalRow));
  const mockEvent = _createMockEditEvent(10, colIndexes.approverAction, "Approve Original Price", "Choose Action");
  const result = processSingleApprovalAction(TestMocks.MOCK_SHEET, 10, mockEvent, inMemoryRow, colIndexes, originalRow, 1);
  _assertEqual(result, false, `${testName} - Should return false on failure`);
  _assertEqual(inMemoryRow[colIndexes.approverAction - 1], "Choose Action", `${testName} - Action should be reverted`);
  _assertTrue(TestMocks.MOCK_TOAST_MESSAGE.includes("invalid or missing LRF"), `${testName} - Toast should explain the error`);
  tearDown();
}

// =================================================================
// --- INTEGRATION TESTS ---
// =================================================================

function test_handleSheetAutomations_directApprovalAction_Integration() {
    const testName = "Integration Test: Direct approval via dropdown edit";
    withTestConfig(function() {
      setUp();
      withTestSheet(MOCK_DATA_INTEGRATION.csvForApprovalWorkflowTests, function(sheet) {
          const targetRow = 2; 
          const colIndexes = CONFIG.approvalWorkflow.columnIndices;
          const actionCell = sheet.getRange(targetRow, colIndexes.approverAction);
          const oldValue = actionCell.getValue();
          const newValue = "Approve Original Price";
          actionCell.setValue(newValue);
          const mockEvent = { range: actionCell, value: newValue, oldValue: oldValue };
          handleSheetAutomations(mockEvent); 
          const finalStatus = sheet.getRange(targetRow, colIndexes.status).getValue();
          const finalPrice = sheet.getRange(targetRow, colIndexes.financeApprovedPrice).getValue();
          const finalApprover = sheet.getRange(targetRow, colIndexes.approvedBy).getValue();
          _assertEqual(finalStatus, CONFIG.approvalWorkflow.statusStrings.approvedOriginal, `${testName} - Status should be updated in the sheet`);
          _assertEqual(finalPrice, 100, `${testName} - Price should be updated in the sheet`);
          _assertEqual(finalApprover, TestMocks.MOCK_SESSION_EMAIL, `${testName} - Approver should be recorded in the sheet`);
      });
      tearDown();
    });
}

function test_runSheetHealthCheck_findsAndFixesInconsistencies() {
    const testName = "Integration Test: runSheetHealthCheck finds and fixes inconsistencies";
    withTestConfig(function() {
        setUp();
        withTestSheet(MOCK_DATA_INTEGRATION.csvForHealthCheckTests, function(sheet) {
            const statusCol = CONFIG.approvalWorkflow.columnIndices.status;
            runSheetHealthCheck();
            const status_R2 = sheet.getRange(2, statusCol).getValue();
            const status_R3 = sheet.getRange(3, statusCol).getValue();
            const status_R4 = sheet.getRange(4, statusCol).getValue();
            _assertEqual(status_R2, CONFIG.approvalWorkflow.statusStrings.pending, `${testName} - R2 (Approved without date) should be reverted to Pending.`);
            _assertEqual(status_R3, CONFIG.approvalWorkflow.statusStrings.approvedOriginal, `${testName} - R3 (Healthy Approved row) should remain Approved.`);
            _assertEqual(status_R4, CONFIG.approvalWorkflow.statusStrings.pending, `${testName} - R4 (Rejected without date) should be reverted to Pending.`);
            _assertTrue(TestMocks.MOCK_TOAST_MESSAGE.includes("Reverted 2 rows"), `${testName} - Toast should report 2 fixes.`);
        });
        tearDown();
    });
}

function test_runSheetHealthCheck_noInconsistenciesFound() {
    const testName = "Integration Test: runSheetHealthCheck finds no issues";
    
    // MODIFIED: Manipulate the array directly instead of using String.replace()
    const healthyData = JSON.parse(JSON.stringify(MOCK_DATA_INTEGRATION.csvForHealthCheckTests));
    const c = CONFIG.approvalWorkflow.columnIndices;
    // Fix the inconsistent rows in our copied data. The header is index 0. Data starts at index 1.
    healthyData[1][c.approvalDate - 1] = '2025-01-01'; // Fix data row 1 (sheet row 2)
    healthyData[3][c.approvalDate - 1] = '2025-01-01'; // Fix data row 3 (sheet row 4)
        
    withTestConfig(function() {
        setUp();
        withTestSheet(healthyData, function(sheet) {
            runSheetHealthCheck();
            _assertTrue(TestMocks.MOCK_TOAST_MESSAGE.includes("No inconsistencies found"), `${testName} - Toast should report no issues.`);
        });
        tearDown();
    });
}