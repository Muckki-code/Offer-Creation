/**
 * @file This file contains the main test runners and assertion helpers for the entire application.
 */

// =================================================================
// --- ASSERTION HELPERS ---
// =================================================================

const TestLogger = {
  log: function(message) {
    Log.TestResults_gs(message);
  }
};

let TOTAL_TESTS = 0;
let PASSED_TESTS = 0;
let FAILED_TEST_MESSAGES = [];

function _resetTestResults() {
  TOTAL_TESTS = 0;
  PASSED_TESTS = 0;
  FAILED_TEST_MESSAGES = [];
  TestLogger.log("--- Resetting Test Results ---");
}

function _assertEqual(actual, expected, message) {
  TOTAL_TESTS++;
  if (actual === expected) {
    PASSED_TESTS++;
    TestLogger.log(`PASSED: ${message}. Got expected value: '${expected}'.`);
  } else {
    FAILED_TEST_MESSAGES.push(`${message}. Expected: '${expected}', but got: '${actual}'.`);
    TestLogger.log(`FAILED: ${message}. Expected: '${expected}', but got: '${actual}'.`);
  }
}

function _assertNotEqual(actual, unexpected, message) {
  TOTAL_TESTS++;
  if (actual !== unexpected) {
    PASSED_TESTS++;
    TestLogger.log(`PASSED: ${message}. Value '${actual}' was not '${unexpected}', as expected.`);
  } else {
    FAILED_TEST_MESSAGES.push(`${message}. Expected value to not be '${unexpected}', but it was.`);
    TestLogger.log(`FAILED: ${message}. Expected value to not be '${unexpected}', but it was.`);
  }
}

function _assertNotNull(value, message) {
  TOTAL_TESTS++;
  if (value !== null && typeof value !== 'undefined') {
    PASSED_TESTS++;
    TestLogger.log(`PASSED: ${message}. Value is not null.`);
  } else {
    FAILED_TEST_MESSAGES.push(`${message}. Value was null or undefined.`);
    TestLogger.log(`FAILED: ${message}. Value was null or undefined.`);
  }
}

function _assertTrue(condition, message) {
  TOTAL_TESTS++;
  if (condition === true) {
    PASSED_TESTS++;
    TestLogger.log(`PASSED: ${message}. Condition is true.`);
  } else {
    FAILED_TEST_MESSAGES.push(`${message}. Condition was false.`);
    TestLogger.log(`FAILED: ${message}. Condition was false.`);
  }
}

function _assertWithinTolerance(actual, expected, tolerance, message) {
  TOTAL_TESTS++;
  if (Math.abs(actual - expected) <= tolerance) {
    PASSED_TESTS++;
    TestLogger.log(`PASSED: ${message}. Actual value ${actual} is within tolerance of ${expected}.`);
  } else {
    FAILED_TEST_MESSAGES.push(`${message}. Expected: ${expected} (within ${tolerance}), but got: ${actual}.`);
    TestLogger.log(`FAILED: ${message}. Expected: ${expected} (within ${tolerance}), but got: ${actual}.`);
  }
}

function _assertThrows(func, message) {
  TOTAL_TESTS++;
  try {
    func();
    FAILED_TEST_MESSAGES.push(`${message}. Expected an error to be thrown, but none was.`);
    TestLogger.log(`FAILED: ${message}. Expected an error to be thrown, but none was.`);
  } catch (e) {
    PASSED_TESTS++;
    TestLogger.log(`PASSED: ${message}. An error was thrown as expected: ${e.message}.`);
  }
}

function _reportTestSummary(suiteType) {
  TestLogger.log("========================================");
  TestLogger.log(`=== FAILED TESTS SUMMARY (${suiteType.toUpperCase()} TESTS) ===`);
  TestLogger.log("========================================");
  if (FAILED_TEST_MESSAGES.length > 0) {
    for (let i = 0; i < FAILED_TEST_MESSAGES.length; i++) {
      TestLogger.log(`(${i + 1}) ${FAILED_TEST_MESSAGES[i]}`);
    }
  } else {
    TestLogger.log("All tests passed!");
  }
  TestLogger.log("========================================");
  TestLogger.log(`Total Failed Tests: ${FAILED_TEST_MESSAGES.length}`);
  TestLogger.log("========================================");
  TestLogger.log("");
}

// =================================================================
// --- MAIN TEST RUNNERS ---
// =================================================================

function runAllTests() {
  _resetTestResults();
  const startTime = new Date().getTime();

  runUnitTests();
  runIntegrationTests();

  const endTime = new Date().getTime();
  const totalTime = (endTime - startTime) / 1000;

  TestLogger.log(`--- All Tests Completed in ${totalTime.toFixed(2)} seconds ---`);

  TestLogger.log("\n========================================");
  TestLogger.log("=== OVERALL TEST SUMMARY ===");
  TestLogger.log("========================================");
  TestLogger.log(`Total Tests Run: ${TOTAL_TESTS}`);
  TestLogger.log(`Tests Passed: ${PASSED_TESTS}`);
  TestLogger.log(`Tests Failed: ${FAILED_TEST_MESSAGES.length}`);
  if (FAILED_TEST_MESSAGES.length > 0) {
    for (let i = 0; i < FAILED_TEST_MESSAGES.length; i++) {
      TestLogger.log(`- ${FAILED_TEST_MESSAGES[i]}`);
    }
    throw new Error("Some tests failed. Check logs for details.");
  } else {
    TestLogger.log("All tests passed successfully!");
  }
  TestLogger.log("========================================");
}

function runUnitTests() {
  Log.TestResults_gs("\n========================================");
  Log.TestResults_gs("=== Starting Unit Tests           ===");
  Log.TestResults_gs("========================================");
  
  runSheetCoreAutomations_UnitTests();
  runApprovalWorkflow_UnitTests();
  runSheetStatusLogic_UnitTests(); 
  
  Log.TestResults_gs("--- Unit Tests Finished ---");
  _reportTestSummary("Unit");
}

function runIntegrationTests() {
  Log.TestResults_gs("\n========================================");
  Log.TestResults_gs("=== Starting Integration Tests      ===");
  Log.TestResults_gs("========================================");
  
  runSheetCoreAutomations_IntegrationTests();
  runSheetCoreAutomations_SanitizationTests(); // MODIFIED: Call the sanitization suite
  runApprovalWorkflow_IntegrationTests(); 
  runUxControl_IntegrationTests(); 
  runBundleService_IntegrationTests();
  runDocumentDataService_IntegrationTests();
  runDocGenerator_IntegrationTests();

  Log.TestResults_gs("--- Integration Tests Finished ---");
  _reportTestSummary("Integration");
}
// =================================================================
// --- TEST MENU SETUP ---
// =================================================================

/**
 * Adds test-related menus to the Spreadsheet UI.
 * This function is called from Main.gs:onOpen.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The Spreadsheet UI object.
 */
function _addTestMenus(ui) {
  const testMenu = ui.createMenu('Tests');
  testMenu.addItem('Run All Tests', 'runAllTests');
  testMenu.addSeparator();

  // Sub-menu for all Unit Tests
  const unitTestsSubMenu = ui.createMenu('Unit Tests');
  unitTestsSubMenu.addItem('SheetCoreAutomations', 'runSheetCoreAutomations_UnitTests');
  unitTestsSubMenu.addItem('ApprovalWorkflow', 'runApprovalWorkflow_UnitTests');
  unitTestsSubMenu.addItem('SheetStatusLogic', 'runSheetStatusLogic_UnitTests');
  testMenu.addSubMenu(unitTestsSubMenu);
  
  // Sub-menu for all Integration Tests
  const integrationTestsSubMenu = ui.createMenu('Integration Tests');
  integrationTestsSubMenu.addItem('SheetCoreAutomations', 'runSheetCoreAutomations_IntegrationTests');
  integrationTestsSubMenu.addItem('Sanitization (Core)', 'runSheetCoreAutomations_SanitizationTests');
  integrationTestsSubMenu.addItem('ApprovalWorkflow', 'runApprovalWorkflow_IntegrationTests');
  integrationTestsSubMenu.addItem('UxControl', 'runUxControl_IntegrationTests');
  integrationTestsSubMenu.addSeparator();
  integrationTestsSubMenu.addItem('BundleService', 'runBundleService_IntegrationTests');
  integrationTestsSubMenu.addItem('DocumentDataService', 'runDocumentDataService_IntegrationTests');
  integrationTestsSubMenu.addItem('DocGenerator (E2E)', 'runDocGenerator_IntegrationTests');
  testMenu.addSubMenu(integrationTestsSubMenu);

  testMenu.addToUi();
}