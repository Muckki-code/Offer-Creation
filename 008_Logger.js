/**
 * @file This file contains the logging utility for the application,
 * managing console debug logs, execution timing, and asynchronous sheet logging.
 */

// Global Log object that will be populated with logging functions.
var Log = {};
// --- SPRINT 2: PERFORMANCE ---: New global object for timing execution.
var ExecutionTimer = {};
// --- SPRINT 2 REFACTOR ---: Declare public-facing sheet logging functions.
var logTableActivity;
var logDocumentActivity;
var logCommunicationActivity;
var logGeneralActivity;


let _loggedCoveragePoints = new Set();
let _logSpreadsheet = null;

// Self-executing anonymous function to initialize all loggers once per script execution.
(function _initializeLogger() {
  try {
    const loggingConfig = CONFIG.logging;
    const noOp = function() {};

    // --- Section 0: Global Disable Check ---
    if (!loggingConfig || loggingConfig.globalDisable) {
      const fileKeys = (typeof CONFIG !== 'undefined' && CONFIG.logging && CONFIG.logging.file) ? Object.keys(CONFIG.logging.file) : [];
      for (const key of fileKeys) { Log[key] = noOp; }
      ExecutionTimer.start = noOp;
      ExecutionTimer.end = noOp;
      logTableActivity = noOp;
      logDocumentActivity = noOp;
      logCommunicationActivity = noOp;
      logGeneralActivity = noOp;
      return;
    }

    // --- Section 1: Initialize Console and ExecutionTimer Loggers ---
    for (const sourceFile in loggingConfig.file) {
      if (Object.prototype.hasOwnProperty.call(loggingConfig.file, sourceFile)) {
        if (loggingConfig.file[sourceFile]) {
          // If logging is enabled for this source, assign the correct logger.
          switch (sourceFile) {
            case 'TestCoverage_gs':
              Log.TestCoverage_gs = function(coverageInfo) {
                const pointId = `${coverageInfo.file}::${coverageInfo.coverage}`;
                if (!_loggedCoveragePoints.has(pointId)) {
                  _loggedCoveragePoints.add(pointId);
                  Logger.log(`COVERAGE: ${pointId}`);
                }
              };
              break;
            case 'ExecutionTime_gs':
              const timers = {};
              Log.ExecutionTime_gs = function(message) { Logger.log(message); };
              ExecutionTimer.start = function(label) {
                timers[label] = Date.now();
              };
              ExecutionTimer.end = function(label) {
                if (timers[label]) {
                  const duration = Date.now() - timers[label];
                  Log.ExecutionTime_gs(`EXEC-TIME: ${label} took ${duration} ms.`);
                  delete timers[label];
                }
              };
              break;
            default:
              Log[sourceFile] = function(message) { Logger.log(message); };
          }
        } else {
          // If logging is disabled for this source, assign a no-op.
          Log[sourceFile] = noOp;
        }
      }
    }
    
    // Fallback: If ExecutionTime_gs flag was missing or false, ensure methods are no-ops.
    if (typeof ExecutionTimer.start !== 'function') {
        ExecutionTimer.start = noOp;
        ExecutionTimer.end = noOp;
    }

    // --- Section 2: Initialize Sheet Logger ---
    if (loggingConfig.sheetLogDisable) {
        logTableActivity = noOp;
        logDocumentActivity = noOp;
        logCommunicationActivity = noOp;
        logGeneralActivity = noOp;
        Log.Logger_gs("[Logger.gs - Initializer] Sheet logging is DISABLED via config.");
    } else {
        logTableActivity = _logTableActivityImpl;
        logDocumentActivity = _logDocumentActivityImpl;
        logCommunicationActivity = _logCommunicationActivityImpl;
        logGeneralActivity = _logGeneralActivityImpl;
        Log.Logger_gs("[Logger.gs - Initializer] Sheet logging is ENABLED.");
    }

  } catch (e) {
    Logger.log('CRITICAL ERROR: Logger initialization failed. Disabling all logging. Error: ' + e.message);
    const noOp = function() {};
    const fileKeys = (typeof CONFIG !== 'undefined' && CONFIG.logging && CONFIG.logging.file) ? Object.keys(CONFIG.logging.file) : [];
    for (const key of fileKeys) { Log[key] = noOp; }
    ExecutionTimer.start = noOp;
    ExecutionTimer.end = noOp;
    logTableActivity = noOp;
    logDocumentActivity = noOp;
    logCommunicationActivity = noOp;
    logGeneralActivity = noOp;
  }
})();

// --- SPRINT 2: ASYNCHRONOUS LOGGING ARCHITECTURE ---

/**
 * Retrieves or creates the dedicated log spreadsheet.
 * This function is now robust and handles cases where the file might be trashed or deleted.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The log spreadsheet object.
 */
function _getLogSpreadsheet() {
  // Use the execution-scoped cache first if available
  if (_logSpreadsheet) {
    return _logSpreadsheet;
  }

  const properties = PropertiesService.getScriptProperties();
  const logFileId = properties.getProperty('logSpreadsheetId');
  const logFileName = CONFIG.logSheets.logSpreadsheetName || 'Application Logs';
  let logFile = null;

  // --- Step 1: Try to find the file using its stored ID ---
  if (logFileId) {
    try {
      logFile = DriveApp.getFileById(logFileId);
      // Check if the file was trashed. If so, treat it as non-existent.
      if (logFile.isTrashed()) {
        Log.Logger_gs(`[Logger.gs - _getLogSpreadsheet] Log spreadsheet (ID: ${logFileId}) was found in trash. Will create a new one.`);
        logFile = null; // Invalidate it
        properties.deleteProperty('logSpreadsheetId');
      }
    } catch (e) {
      Log.Logger_gs(`[Logger.gs - _getLogSpreadsheet] Could not access log spreadsheet by ID: ${logFileId}. It may have been deleted. Error: ${e.message}`);
      logFile = null; // Invalidate it
      properties.deleteProperty('logSpreadsheetId');
    }
  }

  // --- Step 2: If no valid file yet, try to find it by name in the same folder ---
  if (!logFile) {
    const mainSheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    const parentFolders = mainSheetFile.getParents();
    if (parentFolders.hasNext()) {
      const folder = parentFolders.next();
      const files = folder.getFilesByName(logFileName);
      if (files.hasNext()) {
        logFile = files.next();
        properties.setProperty('logSpreadsheetId', logFile.getId());
        Log.Logger_gs(`[Logger.gs - _getLogSpreadsheet] Found existing log spreadsheet by name: '${logFileName}'.`);
      }
    }
  }
  
  // --- Step 3: If still no file, create a new one ---
  if (!logFile) {
    Log.Logger_gs(`[Logger.gs - _getLogSpreadsheet] No valid log spreadsheet found. Creating new file named '${logFileName}'.`);
    const newSpreadsheet = SpreadsheetApp.create(logFileName);
    const newLogFileId = newSpreadsheet.getId();
    properties.setProperty('logSpreadsheetId', newLogFileId);
    logFile = DriveApp.getFileById(newLogFileId);
    
    // Move the new log file to the same folder as the main sheet
    const mainSheetFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
    logFile.moveTo(mainSheetFolder);
    Log.Logger_gs(`[Logger.gs - _getLogSpreadsheet] New log spreadsheet created and moved to the correct folder.`);
  }

  // --- Step 4: Open the validated file as a Spreadsheet object ---
  _logSpreadsheet = SpreadsheetApp.openById(logFile.getId());
  return _logSpreadsheet;
}

function _getOrCreateLogSheet(sheetName, headerRow) {
  const ss = _getLogSpreadsheet();
  let logSheet = ss.getSheetByName(sheetName);
  if (!logSheet) {
    logSheet = ss.insertSheet(sheetName);
    logSheet.appendRow(headerRow);
    logSheet.getRange(1, 1, 1, headerRow.length).setFontWeight("bold");
  } else if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(headerRow);
    logSheet.getRange(1, 1, 1, headerRow.length).setFontWeight("bold");
  }
  return logSheet;
}

function _processLogQueue() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) { return; }
    try {
        const properties = PropertiesService.getScriptProperties();
        const logQueueJson = properties.getProperty('logQueue');
        if (!logQueueJson) { return; }
        const logQueue = JSON.parse(logQueueJson);
        if (logQueue.length === 0) { return; }
        const logsBySheet = logQueue.reduce((acc, item) => {
            const { sheetName, logEntry } = item;
            if (!acc[sheetName]) { acc[sheetName] = []; }
            acc[sheetName].push(logEntry);
            return acc;
        }, {});
        for (const sheetName in logsBySheet) {
            const config = Object.values(CONFIG.logSheets).find(s => s.sheetName === sheetName);
            if (config) {
                const logSheet = _getOrCreateLogSheet(sheetName, Object.values(config.columns));
                logSheet.getRange(logSheet.getLastRow() + 1, 1, logsBySheet[sheetName].length, logsBySheet[sheetName][0].length)
                       .setValues(logsBySheet[sheetName]);
            }
        }
        properties.deleteProperty('logQueue');
        _deleteTriggerByName('_processLogQueue');
    } catch (e) {
        Log.Logger_gs(`[Logger.gs - _processLogQueue] CRITICAL ERROR during log processing: ${e.message}. Stack: ${e.stack}`);
    } finally {
        lock.releaseLock();
    }
}

function _deleteTriggerByName(functionName) {
    const allTriggers = ScriptApp.getProjectTriggers();
    for (const trigger of allTriggers) {
        if (trigger.getHandlerFunction() === functionName) {
            ScriptApp.deleteTrigger(trigger);
        }
    }
}

function _queueLogEntry(sheetName, logEntry) {
    try {
        const properties = PropertiesService.getScriptProperties();
        const logQueueJson = properties.getProperty('logQueue');
        const logQueue = logQueueJson ? JSON.parse(logQueueJson) : [];
        logQueue.push({ sheetName, logEntry });
        properties.setProperty('logQueue', JSON.stringify(logQueue));
        _deleteTriggerByName('_processLogQueue');
        ScriptApp.newTrigger('_processLogQueue')
            .timeBased()
            .after(10 * 1000) // 10 seconds
            .create();
    } catch (e) {
        Log.Logger_gs(`[Logger.gs - _queueLogEntry] ERROR: Failed to queue log entry: ${e.message}`);
    }
}

function _getUserDisplayName() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (userEmail) {
      const userNameParts = userEmail.split('@')[0].split('.');
      return userNameParts.map(part => part.charAt(0).toUpperCase() + part.slice(1)).join(' ');
    }
  } catch(e) { /* fall through */ }
  return "Unknown User";
}

// --- "PRIVATE" IMPLEMENTATION FUNCTIONS ---

function _logTableActivityImpl(options) {
  const logSheetConfig = CONFIG.logSheets.tableLogs;
  // currentFullRowValues is the array representing the data row,
  // starting from the sheet's `startCol`.
  const mainSheetAllValues = options.currentFullRowValues;
  // Default startCol to 1 (A) if not provided, for compatibility/safety
  const startCol = options.startCol !== undefined ? options.startCol : 1; 

  const getVal = (colLetter) => {
    const indexInSheet = getColumnIndexByLetter(colLetter); // This is the 1-based column index in the sheet (e.g., 20 for 'T')
    // MODIFIED: Calculate the correct 0-based index within the `mainSheetAllValues` array
    const arrayIndex = indexInSheet - startCol; 
    
    // Ensure the arrayIndex is within the bounds of the provided array
    if (arrayIndex >= 0 && arrayIndex < mainSheetAllValues.length) {
      return mainSheetAllValues[arrayIndex];
    }
    return ""; // Return empty string if the column is outside the array's bounds
  };

  const logEntryData = {};
  for (const logColKey in logSheetConfig.columns) {
      switch (logColKey) {
          case "timestamp": logEntryData[logColKey] = new Date(); break;
          case "name": logEntryData[logColKey] = _getUserDisplayName(); break;
          // MODIFIED: Use getVal for SKU
          case "bq_info_sku": logEntryData[logColKey] = getVal(CONFIG.documentDeviceData.columns.sku); break;
          case "bq_info_prices":
              // MODIFIED: Use getVal for all BQ related columns
              logEntryData[logColKey] = `EP:${getVal(CONFIG.documentDeviceData.columns.epCapexRaw)}|TK:${getVal(CONFIG.documentDeviceData.columns.tkCapexRaw)}|Tgt:${getVal(CONFIG.documentDeviceData.columns.rentalTargetRaw)}|Limit:${getVal(CONFIG.documentDeviceData.columns.rentalLimitRaw)}`;
              break;
          default:
              let sourceColLetter = CONFIG.approvalWorkflow.columns[logColKey] || CONFIG.documentDeviceData.columns[logColKey];
              logEntryData[logColKey] = sourceColLetter ? getVal(sourceColLetter) : "";
      }
  }
  _queueLogEntry(logSheetConfig.sheetName, Object.keys(logSheetConfig.columns).map(key => logEntryData[key]));
}

function _logDocumentActivityImpl(options) {
  const logSheetConfig = CONFIG.logSheets.documentLogs;
  const logEntry = [
    new Date(), _getUserDisplayName(), options.action || "N/A", options.docName || "", options.docUrl || "",
    options.customerCompany || "", options.offerType || "", options.language || "", JSON.stringify(options.details || "")
  ];
  _queueLogEntry(logSheetConfig.sheetName, logEntry);
}

function _logCommunicationActivityImpl(options) {
  const logSheetConfig = CONFIG.logSheets.communicationLogs;
  const logEntry = [
    new Date(), _getUserDisplayName(), options.recipient || "", options.subject || "",
    options.type || "", JSON.stringify(options.details || "")
  ];
  _queueLogEntry(logSheetConfig.sheetName, logEntry);
}

function _logGeneralActivityImpl(options) {
  const logSheetConfig = CONFIG.logSheets.generalLogs;
  const logEntry = [
    new Date(), _getUserDisplayName(), options.action || "", JSON.stringify(options.details || "")
  ];
  _queueLogEntry(logSheetConfig.sheetName, logEntry);
}