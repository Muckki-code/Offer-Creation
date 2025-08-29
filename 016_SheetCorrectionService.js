/**
 * @file This file contains functions for actively correcting data and structure
 * in the sheet, typically in response to a user-confirmed dialog.
 */

/**
 * Applies the correct term and quantity to ALL rows within a given bundle.
 * @param {string|number} bundleNumber The bundle ID to correct.
 * @param {string|number} term The correct term to apply.
 * @param {string|number} quantity The correct quantity to apply.
 */
function applyBundleCorrection(bundleNumber, term, quantity) {
    const sourceFile = "SheetCorrectionService_gs";
    Log[sourceFile](`[${sourceFile} - applyBundleCorrection] Start. Applying Term=${term}, Qty=${quantity} to ALL of bundle #${bundleNumber}.`);
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        
        // Use validateBundle to get the exact range of the bundle.
        const validationResult = validateBundle(sheet, 0, bundleNumber); // rowNum doesn't matter here
        if (!validationResult.startRow || !validationResult.endRow) {
            throw new Error(`Could not find the range for bundle #${bundleNumber}.`);
        }

        const termCol = CONFIG.approvalWorkflow.columnIndices.aeTerm;
        const quantityCol = CONFIG.approvalWorkflow.columnIndices.aeQuantity;
        const numRows = validationResult.endRow - validationResult.startRow + 1;

        // Apply the correction to the entire bundle range in two efficient calls
        sheet.getRange(validationResult.startRow, termCol, numRows).setValue(term);
        sheet.getRange(validationResult.startRow, quantityCol, numRows).setValue(quantity);
        
        SpreadsheetApp.flush(); // Ensure changes are saved before re-running automation

        // Trigger a recalculation for the whole sheet to ensure all statuses update
        recalculateAllRows();

        SpreadsheetApp.getActive().toast(`Bundle #${bundleNumber} has been corrected.`, "Success", 3);
        Log[sourceFile](`[${sourceFile} - applyBundleCorrection] End. Correction successful.`);
    } catch(e) {
        Log[sourceFile](`[${sourceFile} - applyBundleCorrection] ERROR: ${e.message}`);
        SpreadsheetApp.getActive().toast(`Failed to apply bundle correction: ${e.message}`, "Error", 5);
    }
}


/**
 * Handles the user cancelling a mismatch dialog. It removes the edited item
 * from the bundle by clearing its bundle number.
 * @param {number} rowNum The row number of the item to remove from the bundle.
 */
function handleBundleMismatchCancellation(rowNum) {
    const sourceFile = "SheetCorrectionService_gs";
    Log[sourceFile](`[${sourceFile} - handleBundleMismatchCancellation] Start. Removing row ${rowNum} from its bundle.`);
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const bundleNumCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;
        const range = sheet.getRange(rowNum, bundleNumCol);
        
        // Simulate an edit from the old bundle number to blank
        const mockEvent = {
            range: range,
            value: "",
            oldValue: range.getValue() 
        };
        
        range.setValue(""); // Clear the bundle number
        handleSheetAutomations(mockEvent); // Re-run automations to update borders, etc.

        SpreadsheetApp.getActive().toast(`Row ${rowNum} was removed from the bundle.`, "Action Cancelled", 3);
    } catch(e) {
        Log[sourceFile](`[${sourceFile} - handleBundleMismatchCancellation] ERROR: ${e.message}`);
        SpreadsheetApp.getActive().toast(`Failed to remove item from bundle: ${e.message}`, "Error", 5);
    }
}


/**
 * Fixes non-consecutive bundle rows by moving them together.
 * @param {string|number} bundleNumber The bundle ID to fix.
 */
function fixBundleGaps(bundleNumber) {
    // This function remains unchanged, it is already correct.
    const sourceFile = "SheetCorrectionService_gs";
    Log[sourceFile](`[${sourceFile} - fixBundleGaps] Start. Fixing gaps for bundle #${bundleNumber}.`);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
        SpreadsheetApp.getActive().toast("Sheet is busy, please try again.", "Error", 5);
        return;
    }
    
    try {
        const dataStartRow = CONFIG.approvalWorkflow.startDataRow;
        const lastRow = sheet.getLastRow();
        const bundleNumCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;
        const bundleColumnValues = sheet.getRange(dataStartRow, bundleNumCol, lastRow - dataStartRow + 1, 1).getValues();
        const bundleRows = [];
        bundleColumnValues.forEach((val, i) => {
            if (String(val[0]).trim() == String(bundleNumber)) {
                bundleRows.push(dataStartRow + i);
            }
        });
        if (bundleRows.length <= 1) { return; }
        const targetStartRow = bundleRows[0] + 1;
        const rowsToMove = bundleRows.slice(1).reverse();
        rowsToMove.forEach((sourceRow, i) => {
            const destinationRow = targetStartRow + (bundleRows.length - 2 - i);
            if (sourceRow !== destinationRow) {
                 sheet.moveRows(sheet.getRange(sourceRow + ":" + sourceRow), destinationRow);
            }
        });
        SpreadsheetApp.flush();
        const mockEvent = {
            range: sheet.getRange(bundleRows[0], bundleNumCol),
            value: bundleNumber,
            oldValue: bundleNumber
        };
        handleSheetAutomations(mockEvent);
        SpreadsheetApp.getActive().toast(`Bundle #${bundleNumber} has been re-ordered.`, "Success", 3);
    } catch(e) {
        Log[sourceFile](`[${sourceFile} - fixBundleGaps] ERROR: ${e.message}`);
        SpreadsheetApp.getActive().toast(`Failed to fix bundle gaps: ${e.message}`, "Error", 5);
    } finally {
        lock.releaseLock();
    }
}