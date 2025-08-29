/**
 * @file This file contains functions related to managing and displaying
 * server-side generated UI elements like dialogs and sidebars.
 */

const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const PROP_KEY_DIALOG_DATA = 'dialogData';

// =================================================================
// --- MISMATCH DIALOG FUNCTIONS ---
// =================================================================

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
    
    _showSidebarPanel('_getMismatchCorrectionHtml', dialogData);
    
    Log[sourceFile](`[${sourceFile} - showBundleMismatchDialog] End. Mismatch correction data has been staged for the sidebar.`);
}


/**
 * Handles the user's confirmation from the mismatch dialog.
 * @param {string|number} bundleNumber The bundle ID to correct.
 * @param {string|number} term The correct term to apply.
 * @param {string|number} quantity The correct quantity to apply.
 */
function handleBundleMismatchConfirmation(bundleNumber, term, quantity) {
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - handleBundleMismatchConfirmation] Start. User confirmed fix for bundle #${bundleNumber}.`);
    // --- UPDATED to call the refactored correction function ---
    applyBundleCorrection(bundleNumber, term, quantity);
    Log[sourceFile](`[${sourceFile} - handleBundleMismatchConfirmation] End.`);
}

/**
 * --- NEW ---
 * Handles the user cancelling the mismatch dialog.
 * @param {number} rowNum The row number of the item to remove from the bundle.
 */
function handleBundleMismatchCancellation(rowNum) {
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - handleBundleMismatchCancellation] Start. User cancelled; removing row ${rowNum} from bundle.`);
    // --- THIS IS THE FIX ---
    // It should call the function in SheetCorrectionService, not itself.
    SheetCorrectionService.handleBundleMismatchCancellation(rowNum); 
    Log[sourceFile](`[${sourceFile} - handleBundleMismatchCancellation] End.`);
}


// =================================================================
// --- GAP DIALOG FUNCTIONS ---
// =================================================================

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

    _showSidebarPanel('_getGapCorrectionHtml', dialogData);

    Log[sourceFile](`[${sourceFile} - showBundleGapDialog] End. Gap correction data has been staged for the sidebar.`);
}


/**
 * Handles the user's confirmation from the gap dialog.
 * @param {string|number} bundleNumber The bundle ID to fix.
 */
function handleBundleGapConfirmation(bundleNumber) {
    // This function remains unchanged and correct.
    const sourceFile = "UiService_gs";
    Log[sourceFile](`[${sourceFile} - handleBundleGapConfirmation] Start. User confirmed fix for bundle #${bundleNumber}.`);
    fixBundleGaps(bundleNumber);
    Log[sourceFile](`[${sourceFile} - handleBundleGapConfirmation] End.`);
}

/**
 * --- NEW / INTERNAL ---
 * Generates the HTML content for the bundle mismatch correction UI.
 * This will be injected into the ActionSidebar.
 * @param {Object} data The data object, same as what was passed to the old dialog.
 * @returns {string} The HTML content as a string.
 */
function _getMismatchCorrectionHtml(data) {
  const sourceFile = "UiService_gs";
  Log[sourceFile](`[${sourceFile} - _getMismatchCorrectionHtml] Start. Generating mismatch HTML.`);
  
  const html = `
    <div class="p-4 bg-red-50 border-l-4 border-red-500">
        <h2 class="text-lg font-bold text-red-800 mb-2">Bundle Mismatch</h2>
        <p class="text-sm text-gray-700 mb-4">
            The Term or Quantity for an item in bundle #${data.bundleNumber} does not match the rest of the bundle.
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
            <button class="btn-secondary" onclick="handleCancel()">Cancel & Revert</button>
            <button class="btn-primary" onclick="handleSubmit()">Apply Fix</button>
        </div>
    </div>
    <script>
      const dialogData = ${JSON.stringify(data)};

      function handleSubmit() {
        const term = document.getElementById('correctedTerm').value;
        const quantity = document.getElementById('correctedQuantity').value;
        const button = document.querySelector('.btn-primary');
        button.disabled = true;
        button.textContent = "Applying...";
        
        google.script.run
          .withSuccessHandler(() => closeCorrectionPanel())
          .withFailureHandler(err => alert('Error: ' + err.message))
          .handleBundleMismatchConfirmation(dialogData.bundleNumber, term, quantity);
      }

      function handleCancel() {
        const button = document.querySelector('.btn-secondary');
        button.disabled = true;
        button.textContent = "Reverting...";

        google.script.run
          .withSuccessHandler(() => closeCorrectionPanel())
          .withFailureHandler(err => alert('Error: ' + err.message))
          .handleBundleMismatchCancellation(dialogData.rowNum);
      }
    <\/script>
  `;
  Log[sourceFile](`[${sourceFile} - _getMismatchCorrectionHtml] End.`);
  return html;
}

/**
 * --- NEW / INTERNAL ---
 * Generates the HTML content for the bundle gap correction UI.
 * This will be injected into the ActionSidebar.
 * @param {Object} data The data object containing the bundleNumber.
 * @returns {string} The HTML content as a string.
 */
function _getGapCorrectionHtml(data) {
  const sourceFile = "UiService_gs";
  Log[sourceFile](`[${sourceFile} - _getGapCorrectionHtml] Start. Generating gap HTML.`);

  const html = `
    <div class="p-4 bg-yellow-50 border-l-4 border-yellow-500">
        <h2 class="text-lg font-bold text-yellow-800 mb-2">Bundle Arrangement Error</h2>
        <p class="text-sm text-gray-700 mb-4">
            The items for bundle #<span class="font-bold">${data.bundleNumber}</span> are not in consecutive rows. This can cause calculation errors.
        </p>
        <p class="text-sm text-gray-700 mb-4">
            Would you like the script to automatically move the rows together to fix this?
        </p>
        <div class="flex justify-end space-x-3 mt-4">
            <button class="btn-secondary" onclick="closeCorrectionPanel()">Cancel</button>
            <button class="btn-primary" onclick="handleSubmit()">Yes, Fix It</button>
        </div>
    </div>
    <script>
      const dialogData = ${JSON.stringify(data)};

      function handleSubmit() {
        const button = document.querySelector('.btn-primary');
        button.disabled = true;
        button.textContent = "Fixing...";
        
        google.script.run
          .withSuccessHandler(() => closeCorrectionPanel())
          .withFailureHandler((err) => {
              alert('Error: ' + err.message);
              closeCorrectionPanel();
          })
          .handleBundleGapConfirmation(dialogData.bundleNumber);
      }
    <\/script>
  `;
  Log[sourceFile](`[${sourceFile} - _getGapCorrectionHtml] End.`);
  return html;
}
