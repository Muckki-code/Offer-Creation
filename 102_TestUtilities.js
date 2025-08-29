/**
 * @file This file contains utility functions for creating and managing temporary test environments,
 * specifically for setting up and tearing down test sheets with mock data.
 */

/**
 * Creates a new sheet with a unique name and populates it with data.
 * The function now intelligently handles both raw CSV strings and pre-parsed 2D arrays
 * and correctly places data based on the configured start column.
 *
 * @param {string|Array<Array<any>>} data The data to populate the sheet with.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The newly created and populated sheet object.
 */
function createTestSheet(data) {
  const sourceFile = "TestUtilities_gs";
  Log[sourceFile](`[${sourceFile} - createTestSheet] Start: Creating test sheet.`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `TestSheet_${new Date().getTime()}`;
  const sheet = ss.insertSheet(sheetName);
  Log[sourceFile](`[${sourceFile} - createTestSheet] Info: Inserted new sheet with name: '${sheetName}'.`);

  // MODIFIED: This now correctly uses the configured start column for data placement.
  const configuredDataStartCol = CONFIG.documentDeviceData.columnIndices.sku;

  if (data) {
    let dataArray;

    if (Array.isArray(data)) {
      dataArray = data;
      Log[sourceFile](`[${sourceFile} - createTestSheet] Info: Input data is a pre-parsed array.`);
    } else if (typeof data === 'string') {
      try {
        dataArray = Utilities.parseCsv(data);
        Log[sourceFile](`[${sourceFile} - createTestSheet] Info: Input data is a CSV string, parsed successfully.`);
      } catch (e) {
        Log[sourceFile](`[${sourceFile} - createTestSheet] ERROR: Failed to parse CSV data. Error: ${e.message}`);
        ss.deleteSheet(sheet);
        throw new Error(`Failed to process CSV data for test sheet: ${e.message}`);
      }
    } else {
       Log[sourceFile](`[${sourceFile} - createTestSheet] ERROR: Unsupported data type for createTestSheet.`);
       ss.deleteSheet(sheet);
       throw new Error('Unsupported data type provided to createTestSheet.');
    }
    
    if (dataArray.length > 0 && dataArray[0].length > 0) {
      const numRows = dataArray.length;
      const numCols = dataArray[0].length;
      
      // The range where data will be placed.
      const range = sheet.getRange(1, configuredDataStartCol, numRows, numCols);
      range.setValues(dataArray);
      Log[sourceFile](`[${sourceFile} - createTestSheet] Info: Populated sheet with ${numRows} rows and ${numCols} columns, starting at column ${configuredDataStartCol}.`);
    } else {
      Log[sourceFile](`[${sourceFile} - createTestSheet] Warning: Data was empty or invalid. Sheet created but not populated.`);
    }
  } else {
    Log[sourceFile](`[${sourceFile} - createTestSheet] Info: No data provided. Created an empty test sheet.`);
  }

  sheet.activate();
  SpreadsheetApp.flush();
  Log[sourceFile](`[${sourceFile} - createTestSheet] End: Test sheet '${sheetName}' created and activated.`);
  return sheet;
}

/**
 * Deletes the specified test sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to delete.
 */
function deleteTestSheet(sheet) {
  const sourceFile = "TestUtilities_gs";
  if (!sheet) {
    Log[sourceFile](`[${sourceFile} - deleteTestSheet] Warning: Attempted to delete a null sheet object.`);
    return;
  }
  const sheetName = sheet.getName();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.deleteSheet(sheet);
    Log[sourceFile](`[${sourceFile} - deleteTestSheet] Info: Successfully deleted test sheet '${sheetName}'.`);
  } catch (e) {
    Log[sourceFile](`[${sourceFile} - deleteTestSheet] ERROR: Could not delete sheet '${sheetName}'. It might have been deleted already. Error: ${e.message}`);
  }
}

/**
 * A wrapper function that handles the entire lifecycle of a test sheet:
 * creation, execution of a test function, and guaranteed cleanup.
 *
 * @param {string|Array<Array<any>>} data The raw CSV data or pre-parsed array to populate the sheet.
 * @param {function(GoogleAppsScript.Spreadsheet.Sheet): void} testFunction The test function to execute. It receives the new sheet as an argument.
 */
function withTestSheet(data, testFunction) {
  const sourceFile = "TestUtilities_gs";
  Log[sourceFile](`[${sourceFile} - withTestSheet] Start: Test lifecycle wrapper initiated.`);
  let sheet = null;
  try {
    sheet = createTestSheet(data);
    if (sheet) {
      testFunction(sheet);
    } else {
      throw new Error("Test sheet could not be created.");
    }
  } catch(e) {
    Log[sourceFile](`[${sourceFile} - withTestSheet] ERROR: An error occurred during the test execution: ${e.message}. Stack: ${e.stack}`);
    throw e;
  } finally {
    if (sheet) {
      const firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
      if (firstSheet) {
        firstSheet.activate();
      }
      deleteTestSheet(sheet);
    }
    Log[sourceFile](`[${sourceFile} - withTestSheet] End: Test lifecycle wrapper finished, cleanup complete.`);
  }
}


/**
 * A wrapper function that handles the lifecycle of a temporary test configuration.
 * It overrides specified CONFIG values for a test's duration and guarantees
 * they are restored afterward, even if the test fails.
 *
 * @param {function(): void} testFunction The test function to execute.
 */
function withTestConfig(testFunction) {
  const sourceFile = "TestUtilities_gs";
  Log[sourceFile](`[${sourceFile} - withTestConfig] Start: Overriding CONFIG for test.`);

  const originalConfig = JSON.parse(JSON.stringify(CONFIG));

  try {
    // Override the configuration for testing scenarios where data starts at row 2 (after 1 header row)
    CONFIG.approvalWorkflow.startDataRow = 2;
    CONFIG.documentDeviceData.startRow = 2;
    CONFIG.bqQuerySettings.scriptStartRow = 2; // Also override this for consistency
    
    testFunction();

  } finally {
    // GUARANTEE that the original configuration is restored
    CONFIG.approvalWorkflow.startDataRow = originalConfig.approvalWorkflow.startDataRow;
    CONFIG.documentDeviceData.startRow = originalConfig.documentDeviceData.startRow;
    CONFIG.bqQuerySettings.scriptStartRow = originalConfig.bqQuerySettings.scriptStartRow;
    Log[sourceFile](`[${sourceFile} - withTestConfig] End: Restored original CONFIG start rows.`);
  }
}