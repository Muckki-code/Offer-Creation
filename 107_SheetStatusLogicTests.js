// In SheetStatusLogicTests.gs

/**
 * @file This file contains the test suite for functions in SheetStatusLogic.gs.
 * These are pure unit tests that operate on in-memory arrays.
 * REFACTORED: The entire suite is rewritten to test the "pure function" nature
 * of the new updateStatusForRow, which now returns a value instead of having side effects.
 */

/**
 * Runs all UNIT tests for SheetStatusLogic.gs.
 */
function runSheetStatusLogic_UnitTests() {
    Log.TestResults_gs("--- Starting SheetStatusLogic Unit Test Suite ---");

    setUp();

    test_updateStatusForRow_newRowToDraft();
    test_updateStatusForRow_draftToPending();
    test_updateStatusForRow_draftRemainsDraft();
    test_updateStatusForRow_pendingToDraft_onDataDeletion();
    test_updateStatusForRow_noChangeOnEdit();
    test_updateStatusForRow_approvedToRevised_onEdit();
    test_updateStatusForRow_rejectedToRevised_onEdit();
    test_updateStatusForRow_revisedByAEToPending();
    test_updateStatusForRow_revisedByAEToDraft_onDataDeletion();
    test_updateStatusForRow_returnsNull_onModelDelete();

    tearDown();

    Log.TestResults_gs("--- SheetStatusLogic Unit Test Suite Finished ---");
}

// =================================================================
// --- MOCK HELPERS (SPECIFIC TO THIS TEST FILE) ---
// =================================================================

function _getTestRow(name) {
    return JSON.parse(JSON.stringify(MOCK_DATA_UNIT.rowsForStatusLogicTests.get(name)));
}

// =================================================================
// --- UNIT TESTS FOR updateStatusForRow() ---
// =================================================================

function test_updateStatusForRow_newRowToDraft() {
    const testName = "Unit Test: updateStatusForRow - New Row to Draft";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('blank');
    const inMemoryRow = _getTestRow('blank');
    inMemoryRow[colIndexes.model - 1] = "New Test Device";

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, CONFIG.approvalWorkflow.statusStrings.draft, `${testName} - Should return 'Draft' status`);
}

function test_updateStatusForRow_draftToPending() {
    const testName = "Unit Test: updateStatusForRow - Draft to Pending Approval";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('draftIncomplete');
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));

    inMemoryRow[colIndexes.aeSalesAskPrice - 1] = 50;
    inMemoryRow[colIndexes.aeTerm - 1] = 24;
    inMemoryRow[colIndexes.aeEpCapex - 1] = 1000; // Add the required Capex

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, CONFIG.approvalWorkflow.statusStrings.pending, `${testName} - Should return 'Pending Approval' status`);
}

function test_updateStatusForRow_draftRemainsDraft() {
    const testName = "Unit Test: updateStatusForRow - Draft remains Draft";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('draftIncomplete');
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));
    inMemoryRow[colIndexes.aeTerm - 1] = 24; // Still missing price

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, CONFIG.approvalWorkflow.statusStrings.draft, `${testName} - Should return 'Draft' status`);
}

function test_updateStatusForRow_pendingToDraft_onDataDeletion() {
    const testName = "Unit Test: updateStatusForRow - Pending to Draft on required data deletion";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('pending');
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));

    inMemoryRow[colIndexes.aeQuantity - 1] = ""; // Make the row incomplete

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, CONFIG.approvalWorkflow.statusStrings.draft, `${testName} - Should return 'Draft' status`);
}

function test_updateStatusForRow_noChangeOnEdit() {
    const testName = "Unit Test: updateStatusForRow - No change on non-key-field edit";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('pending');
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));
    inMemoryRow[colIndexes.approverComments - 1] = "A new comment"; // Edit a non-key field

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, CONFIG.approvalWorkflow.statusStrings.pending, `${testName} - Status should remain 'Pending Approval'`);
}

function test_updateStatusForRow_approvedToRevised_onEdit() {
    const testName = "Unit Test: updateStatusForRow - Approved to Revised by AE";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('approved');
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));
    inMemoryRow[colIndexes.aeQuantity - 1] = 15; // Change a key field

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, CONFIG.approvalWorkflow.statusStrings.revisedByAE, `${testName} - Should return 'Revised by AE' status`);
}

function test_updateStatusForRow_rejectedToRevised_onEdit() {
    const testName = "Unit Test: updateStatusForRow - Rejected to Revised by AE on Edit";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('rejected');
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));
    inMemoryRow[colIndexes.aeQuantity - 1] = 9; // Edit a key field

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, CONFIG.approvalWorkflow.statusStrings.revisedByAE, `${testName} - Should return 'Revised by AE' status`);
}

function test_updateStatusForRow_revisedByAEToPending() {
    const testName = "Unit Test: updateStatusForRow - Revised by AE to Pending on edit";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('revisedByAE');
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));
    inMemoryRow[colIndexes.aeQuantity - 1] = 99; // An AE edits another key field.

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, CONFIG.approvalWorkflow.statusStrings.pending, `${testName} - Should return 'Pending Approval' status`);
}

function test_updateStatusForRow_revisedByAEToDraft_onDataDeletion() {
    const testName = "Unit Test: updateStatusForRow - Revised by AE to Draft on data deletion";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('revisedByAE');
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));
    inMemoryRow[colIndexes.aeSalesAskPrice - 1] = ""; // Delete a required field

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, CONFIG.approvalWorkflow.statusStrings.draft, `${testName} - Should return 'Draft' status`);
}

function test_updateStatusForRow_returnsNull_onModelDelete() {
    const testName = "Unit Test: updateStatusForRow - Deleting model returns null";
    const colIndexes = MOCK_DATA_UNIT.approvalWorkflowTestColIndexes;
    const originalRow = _getTestRow('approved');
    const inMemoryRow = JSON.parse(JSON.stringify(originalRow));
    inMemoryRow[colIndexes.model - 1] = "";

    const newStatus = updateStatusForRow(inMemoryRow, originalRow, false, {}, 1, colIndexes);

    _assertEqual(newStatus, null, `${testName} - Status should be null to signal a clear operation`);
}