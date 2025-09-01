// In Main.gs

/**
 * Shows the main action sidebar in the UI.
 */
function showActionSidebar() {
  const html = HtmlService.createTemplateFromFile('HTML/ActionSidebar').evaluate().setTitle('Action Sidebar');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Triggered when the spreadsheet is opened. Adds the menu.
 * @param {Object} e The event object.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Special Functions')
    .addItem('Show Action Sidebar', 'showActionSidebar')
    .addSeparator()
    .addItem('Setup Triggers (Run Once)', 'setupTriggers')
    .addToUi();
  _addTestMenus(ui);
  scanAndSetAllBundleMetadata();
  // --- THIS IS THE FIX for onOpen logic ---
  recalculateAllRows({ refreshUx: true });
  // --- END FIX ---
}

/**
 * ONE-TIME SETUP FUNCTION. Creates the necessary installable triggers.
 */
function setupTriggers() {
    const sourceFile = "Main_gs";
    Log[sourceFile]("[Main.gs - setupTriggers] User initiated trigger setup.");
    // First, delete any existing triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'handleEdit' || trigger.getHandlerFunction() === 'handleChange') {
            ScriptApp.deleteTrigger(trigger);
            Log[sourceFile]("[Main.gs - setupTriggers] Deleted existing trigger.");
        }
    });

    // Create the new installable triggers
    ScriptApp.newTrigger('handleEdit')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
    Log[sourceFile]("[Main.gs - setupTriggers] CREATED new installable ON_EDIT trigger.");

    ScriptApp.newTrigger('handleChange')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onChange()
        .create();
    Log[sourceFile]("[Main.gs - setupTriggers] CREATED new installable ON_CHANGE trigger.");
    
    SpreadsheetApp.getUi().alert('Success!', 'The necessary triggers have been set up for this sheet. The script will now run with the correct permissions.', SpreadsheetApp.getUi().ButtonSet.OK);
}


/**
 * Main onEdit trigger handler. Called by an INSTALLABLE trigger.
 * @param {Object} e The event object.
 */
function handleEdit(e) {
  handleSheetAutomations(e);
}

/**
 * Triggered by structural changes. Called by an INSTALLABLE trigger.
 * @param {Object} e The onChange event object.
 */
function handleChange(e) {
  if (e.changeType === 'INSERT_ROW' || e.changeType === 'REMOVE_ROW') {
    applyUxRules(false);
  }
}

// --- Other functions can remain as they are, but onInstall is no longer needed for trigger setup ---
function onInstall(e) {
  onOpen(e);
  // Log sheet setup can still be useful here
  _getOrCreateLogSheet(CONFIG.logSheets.tableLogs.sheetName, Object.values(CONFIG.logSheets.tableLogs.columns));
  _getOrCreateLogSheet(CONFIG.logSheets.documentLogs.sheetName, Object.values(CONFIG.logSheets.documentLogs.columns));
  _getOrCreateLogSheet(CONFIG.logSheets.communicationLogs.sheetName, Object.values(CONFIG.logSheets.communicationLogs.columns));
  _getOrCreateLogSheet(CONFIG.logSheets.generalLogs.sheetName, Object.values(CONFIG.logSheets.generalLogs.columns));
}

function runFullSheetRepair() {
  ExecutionTimer.start('runFullSheetRepair_total');
  Log.TestCoverage_gs({ file: 'Main.gs', coverage: 'runFullSheetRepair_start' });
  const ui = SpreadsheetApp.getUi();
  ui.alert('Starting Sheet Repair', 'This will take a moment. The script will now reset all formatting, validation, check for data inconsistencies, and verify the logging setup.', ui.ButtonSet.OK);
  Log.Main_gs("[Main.gs - runFullSheetRepair] Start: Manual repair triggered.");
  ExecutionTimer.start('runFullSheetRepair_ux');
  applyUxRules(true);
  ExecutionTimer.end('runFullSheetRepair_ux');
  ExecutionTimer.start('runFullSheetRepair_data');
  runSheetHealthCheck();
  ExecutionTimer.end('runFullSheetRepair_data');
  ExecutionTimer.start('runFullSheetRepair_logs');
  _getOrCreateLogSheet(CONFIG.logSheets.tableLogs.sheetName, Object.values(CONFIG.logSheets.tableLogs.columns));
  _getOrCreateLogSheet(CONFIG.logSheets.documentLogs.sheetName, Object.values(CONFIG.logSheets.documentLogs.columns));
  _getOrCreateLogSheet(CONFIG.logSheets.communicationLogs.sheetName, Object.values(CONFIG.logSheets.communicationLogs.columns));
  _getOrCreateLogSheet(CONFIG.logSheets.generalLogs.sheetName, Object.values(CONFIG.logSheets.generalLogs.columns));
  ExecutionTimer.end('runFullSheetRepair_logs');
  Log.Main_gs("[Main.gs - runFullSheetRepair] End: Manual repair finished.");
  ui.alert('Repair Complete', 'The sheet has been successfully repaired.', ui.ButtonSet.OK);
  ExecutionTimer.end('runFullSheetRepair_total');
}

// --- Functions to Display the Consolidated Application Overview HTML Page and Provide Content ---
/**
 * Displays the full Application Overview HTML dialog with tabs.
 */
function showApplicationOverviewDialog() {
  Log.TestCoverage_gs({ file: 'Main.gs', coverage: 'showApplicationOverviewDialog_start' });
  Log.Main_gs("[Main.gs - showApplicationOverviewDialog] Start: showApplicationOverviewDialog function started.");
  var htmlOutput = HtmlService.createTemplateFromFile('HTML/ApplicationOverview').evaluate()
      .setWidth(900)
      .setHeight(650)
      .setTitle('Application Architecture Overview');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Application Architecture Overview');
  Log.Main_gs("[Main.gs - showApplicationOverviewDialog] End: Application Overview dialog displayed.");
}

/**
 * Reads and returns the raw HTML content for the 'Overview' tab.
 */
function getOverviewContent() {
  Log.TestCoverage_gs({ file: 'Main.gs', coverage: 'getOverviewContent_start' });
  Log.Main_gs("[Main.gs - getOverviewContent] Start: getOverviewContent function started.");
  const content = `
    <h2 class="mt-4">1. Introduction to the Application</h2>
    <p>This Google Apps Script application is designed to streamline the process of generating offer documents and managing an in-sheet approval workflow. It integrates with Google Sheets for data input and workflow management, and potentially with Google Docs for document generation. The system automates calculations, status updates, conditional formatting, and email notifications to facilitate efficient offer creation and approval.</p>
  `;
  Log.Main_gs("[Main.gs - getOverviewContent] End: Overview content retrieved.");
  return content;
}


/**
 * Reads and returns the raw HTML content for the 'User Guide' tab.
 */
function getUserGuideContent() {
  Log.TestCoverage_gs({ file: 'Main.gs', coverage: 'getUserGuideContent_start' });
  Log.Main_gs("[Main.gs - getUserGuideContent] Start: getUserGuideContent function started.");
  const htmlTemplate = HtmlService.createTemplateFromFile('HTML/UserGuide');
  const content = htmlTemplate.evaluate().getContent();
  Log.Main_gs("[Main.gs - getUserGuideContent] End: User Guide content retrieved.");
  return content;
}

/**
 * Reads and returns the raw HTML content for the 'Offer Document' tab.
 */
function getOfferDocumentDetailsContent() {
  Log.TestCoverage_gs({ file: 'Main.gs', coverage: 'getOfferDocumentDetailsContent_start' });
  Log.Main_gs("[Main.gs - getOfferDocumentDetailsContent] Start: getOfferDocumentDetailsContent function started.");
  const htmlTemplate = HtmlService.createTemplateFromFile('HTML/OfferDocumentDetails');
  const content = htmlTemplate.evaluate().getContent();
  Log.Main_gs("[Main.gs - getOfferDocumentDetailsContent] End: Offer Document Details content retrieved.");
  return content;
}

/**
 * Reads and returns the raw HTML content for the 'Technical Details' tab.
 */
function getTechnicalDetailsContent() {
  Log.TestCoverage_gs({ file: 'Main.gs', coverage: 'getTechnicalDetailsContent_start' });
  Log.Main_gs("[Main.gs - getTechnicalDetailsContent] Start: getTechnicalDetailsContent function started.");
  const htmlTemplate = HtmlService.createTemplateFromFile('HTML/TechnicalDetails');
  const content = htmlTemplate.evaluate().getContent();
  Log.Main_gs("[Main.gs - getTechnicalDetailsContent] End: Technical Details content retrieved.");
  return content;
}

/**
 * Reads and returns the raw HTML content for the 'Code Analysis' tab.
 */
function getCodeAnalysisContent() {
  Log.TestCoverage_gs({ file: 'Main.gs', coverage: 'getCodeAnalysisContent_start' });
  Log.Main_gs("[Main.gs - getCodeAnalysisContent] Start: getCodeAnalysisContent function started.");
  const htmlTemplate = HtmlService.createTemplateFromFile('HTML/CodeAnalysis');
  const content = htmlTemplate.evaluate().getContent();
  Log.Main_gs("[Main.gs - getCodeAnalysisContent] End: Code Analysis content retrieved.");
  return content;
}

// NOTE: Ensure that all functions called from the sidebar (e.g., getDataFromSKU, showOfferDialog, 
// submitItemsForApproval, etc.) are defined globally in your project, just as they were for the old menus.