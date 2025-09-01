/**
 * @file This file contains functions related to managing and displaying
 * server-side generated UI elements like dialogs and sidebars.
 */

// --- SCRIPT PROPERTIES FOR STAGING UI UPDATES ---
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const PROP_KEY_UI_UPDATE = 'sidebarUiUpdate';





/**
 * --- NEW / PUBLIC (but only called by other server functions) ---
 * A helper to stage UI update information in PropertiesService.
 * REVISED: This now accumulates errors in an array (acting as a queue)
 * instead of overwriting a single value.
 * @param {string} functionName The name of the server-side function that generates the HTML for the panel.
 * @param {Object} data The data payload required by the HTML-generating function.
 */
function showSidebarPanel(functionName, data) {
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - showSidebarPanel] Staging sidebar update. Function: ${functionName}.`);
    
    const uiUpdateInfo = {
        functionName: functionName,
        data: data,
        timestamp: new Date().getTime()
    };

    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000); 

        const prop = SCRIPT_PROPS.getProperty(PROP_KEY_UI_UPDATE);
        let updates = prop ? JSON.parse(prop) : [];
        
        // Prevent adding a duplicate error for the same bundle
        const isDuplicate = updates.some(u => u.data.bundleNumber === data.bundleNumber);
        if (!isDuplicate) {
          Log.TestCoverage_gs({ file: sourceFile, coverage: 'showSidebarPanel_notDuplicate' });
          updates.push(uiUpdateInfo);
          SCRIPT_PROPS.setProperty(
              PROP_KEY_UI_UPDATE, 
              JSON.stringify(updates)
          );
          Log[sourceFile](`[${sourceFile} - showSidebarPanel] Successfully staged UI update. Queue size is now ${updates.length}.`);
        } else {
          Log[sourceFile](`[${sourceFile} - showSidebarPanel] Skipped adding duplicate error for bundle #${data.bundleNumber}.`);
        }
    } catch (e) {
        Log[sourceFile](`[${sourceFile} - showSidebarPanel] ERROR: Could not set sidebar update property. Error: ${e.message}`);
    } finally {
        lock.releaseLock();
    }
}


/**
 * --- NEW / PUBLIC ---
 * Called by the client-side sidebar on a timer (polling) to check for pending UI updates.
 * REVISED: This now retrieves and clears the entire queue of errors.
 * @returns {Array<Object>|null} The array of UI update info objects, or null if there are no updates.
 */
function getSidebarUpdate() {
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - getSidebarUpdate] Sidebar is polling for updates.`);
    
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(5000);
        const prop = SCRIPT_PROPS.getProperty(PROP_KEY_UI_UPDATE);
        
        if (prop) {
            Log.TestCoverage_gs({ file: sourceFile, coverage: 'getSidebarUpdate_foundUpdate' });
            Log[sourceFile](`[${sourceFile} - getSidebarUpdate] Found an update queue. Deleting property and returning data to client.`);
            SCRIPT_PROPS.deleteProperty(PROP_KEY_UI_UPDATE); // Clear the queue
            return JSON.parse(prop); // Return the entire array of updates
        }
        
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'getSidebarUpdate_noUpdate' });
        return null;
    } catch(e) {
        Log[sourceFile](`[${sourceFile} - getSidebarUpdate] ERROR polling for updates: ${e.message}`);
        SCRIPT_PROPS.deleteProperty(PROP_KEY_UI_UPDATE); // Clear on error to be safe
        return null;
    } finally {
      if (lock.hasLock()) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'getSidebarUpdate_releaseLock' });
        lock.releaseLock();
      }
    }
}



/**
 * Handles the user's confirmation from the mismatch dialog (now sidebar).
 * @param {string|number} bundleNumber The bundle ID to correct.
 * @param {string|number} term The correct term to apply.
 * @param {string|number} quantity The correct quantity to apply.
 */
function handleBundleMismatchConfirmation(bundleNumber, term, quantity) {
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - handleBundleMismatchConfirmation] Start. User confirmed fix for bundle #${bundleNumber}.`);
    // Correctly calls the global function from 016_SheetCorrectionService.js
    applyBundleCorrection(bundleNumber, term, quantity);
    Log[sourceFile](`[${sourceFile} - handleBundleMismatchConfirmation] End.`);
}

/**
 * --- REVISED ---
 * Handles the user choosing to dissolve a broken bundle from the sidebar.
 * This is called directly from the client-side google.script.run.
 * @param {string|number} bundleNumber The bundle ID to dissolve.
 */
function handleBundleDissolve(bundleNumber) {
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - handleBundleDissolve] Start. Relaying call to dissolve bundle #${bundleNumber}.`);
    
    // This now correctly calls the new, powerful function in 016_SheetCorrectionService.js
    dissolveBundle(bundleNumber); 
    
    Log[sourceFile](`[${sourceFile} - handleBundleDissolve] End.`);
}


/**
 * Handles the user's confirmation from the gap dialog (now sidebar).
 * @param {string|number} bundleNumber The bundle ID to fix.
 */
function handleBundleGapConfirmation(bundleNumber) {
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - handleBundleGapConfirmation] Start. User confirmed fix for bundle #${bundleNumber}.`);
    fixBundleGaps(bundleNumber);
    Log[sourceFile](`[${sourceFile} - handleBundleGapConfirmation] End.`);
}


/**
 * --- REFACTORED ---
 * Stages a UI update for the action sidebar to show the bundle mismatch correction UI.
 * This replaces the old blocking dialog.
 * @param {number} rowNum The row that was edited and caused the error.
 * @param {string|number} bundleNumber The bundle ID that is being corrected.
 * @param {{term: any, quantity: any}} currentValues The incorrect values from the edited row.
 * @param {{term: any, quantity: any}} expectedValues The correct values from the bundle.
 */
function showBundleMismatchDialog(rowNum, bundleNumber, currentValues, expectedValues) {
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - showBundleMismatchDialog] Start. Preparing sidebar correction for mismatch in bundle #${bundleNumber}.`);

    const dialogData = {
        rowNum: rowNum,
        bundleNumber: bundleNumber,
        currentValues: currentValues,
        expectedValues: expectedValues
    };
    
    // --- THIS IS THE FIX ---
    // Call the staging function with the NEW PUBLIC name of the HTML generator
    showSidebarPanel('getMismatchCorrectionHtml', dialogData);
    
    Log[sourceFile](`[${sourceFile} - showBundleMismatchDialog] End. Mismatch correction data has been staged for the sidebar.`);
}

/**
 * --- REFACTORED ---
 * Stages a UI update for the action sidebar to show the bundle gap correction UI.
 * This replaces the old blocking dialog.
 * @param {string|number} bundleNumber The ID of the bundle with gaps.
 */
function showBundleGapDialog(bundleNumber) {
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - showBundleGapDialog] Start. Preparing sidebar correction for gap in bundle #${bundleNumber}.`);

    const dialogData = { 
      bundleNumber: bundleNumber 
    };

    // --- THIS IS THE FIX ---
    // Call the staging function with the NEW PUBLIC name of the HTML generator
    showSidebarPanel('getGapCorrectionHtml', dialogData);

    Log[sourceFile](`[${sourceFile} - showBundleGapDialog] End. Gap correction data has been staged for the sidebar.`);
}

function getMismatchCorrectionHtml(data) {
  const sourceFile = "UiService_gs";
  Log[sourceFile](`[${sourceFile} - getMismatchCorrectionHtml] Start. Generating pure mismatch HTML with data attributes.`);
  
  const html = `
    <div id="mismatch-panel-content" class="p-4 bg-red-50 border-l-4 border-red-500" 
         data-bundle-number="${data.bundleNumber}" data-row-num="${data.rowNum}">
        <h2 class="text-lg font-bold text-red-800 mb-2">Bundle Mismatch</h2>
        <p class="text-sm text-gray-700 mb-4">
            The Term or Quantity for an item in bundle #${data.bundleNumber} does not match the rest.
        </p>
        <div class="space-y-3 mb-4">
            <div>
                <label class="block text-sm font-medium text-gray-600">Correct Term:</label>
                <input type="text" id="correctedTerm" value="${data.expectedValues.term}" class="mt-1 block w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm">
            </div>
            <div>
                <label class="block text-sm font-medium text-gray-600">Correct Quantity:</label>
                <input type="text" id="correctedQuantity" value="${data.expectedValues.quantity}" class="mt-1 block w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm">
            </div>
        </div>
        <div class="flex justify-end space-x-3 mt-4">
            <button class="btn-secondary" onclick="handleCancel(this)">Cancel & Revert</button>
            <button class="btn-primary" onclick="handleSubmit(this, 'mismatch')">Apply Fix</button>
        </div>
    </div>
  `;
  Log[sourceFile](`[${sourceFile} - getMismatchCorrectionHtml] End.`);
  return html;
}


function getGapCorrectionHtml(data) {
  const sourceFile = "UiService_gs";
  Log[sourceFile](`[${sourceFile} - getGapCorrectionHtml] Start. Generating pure gap HTML with data attributes.`);

  const html = `
    <div id="gap-panel-content" class="p-4 bg-yellow-50 border-l-4 border-yellow-500"
         data-bundle-number="${data.bundleNumber}">
        <h2 class="text-lg font-bold text-yellow-800 mb-2">Bundle Arrangement Error</h2>
        <p class="text-sm text-gray-700 mb-4">
            The items for bundle #<span class="font-bold">${data.bundleNumber}</span> are not in consecutive rows.
        </p>
        <div class="flex justify-end space-x-3 mt-4">
            <button class="btn-secondary" onclick="handleCancel(this)">Cancel & Dissolve</button>
            <button class="btn-primary" onclick="handleSubmit(this, 'gap')">Yes, Fix It</button>
        </div>
    </div>
  `;
  Log[sourceFile](`[${sourceFile} - getGapCorrectionHtml] End.`);
  return html;
}



/**
 * --- NEW ---
 * A publicly callable function for the sidebar to get a list of all
 * bundle errors currently present in the sheet upon loading.
 * @returns {Array<Object>} An array of error objects from the BundleService.
 */
function getInitialBundleErrors() {
  const sourceFile = "UiService_gs";
  Log[sourceFile](`[${sourceFile} - getInitialBundleErrors] Start: Sidebar is requesting initial bundle error check.`);
  
  // Directly call the new function from the BundleService
  const errors = findAllBundleErrors();
  
  Log[sourceFile](`[${sourceFile} - getInitialBundleErrors] End: Found ${errors.length} errors. Returning to client.`);
  return errors;
}



