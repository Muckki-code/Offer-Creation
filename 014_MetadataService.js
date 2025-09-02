// In MetadataService.gs (NEW, CORRECTED VERSION)

/**
 * @file This file contains a dedicated service for managing Developer Metadata,
 * specifically for storing and retrieving information about bundles on a ROW level.
 */

const METADATA_KEY_BUNDLE = 'bundleInfo';


// In MetadataService.gs

/**
 * Attaches bundle metadata to each row within a specified range.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {{bundleId: string, startRow: number, endRow: number}} bundleInfo An object describing the bundle.
 */
function _setMetadataForRowRange(sheet, bundleInfo) {
  const sourceFile = "MetadataService_gs";
  ExecutionTimer.start('_setMetadataForRowRange_total');
  Log[sourceFile] (`[${sourceFile} - _setMetadataForRowRange] Start. Setting metadata for bundle #${bundleInfo.bundleId} on rows ${bundleInfo.startRow}-${bundleInfo.endRow}.`);
  
  const metadataValue = JSON.stringify(bundleInfo);
  for (let i = bundleInfo.startRow; i <= bundleInfo.endRow; i++) {
    // FIXED: The range must be the ENTIRE row, specified with A1 notation like "2:2".
    sheet.getRange(i + ":" + i).addDeveloperMetadata(METADATA_KEY_BUNDLE, metadataValue, SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT);
  }
  
  ExecutionTimer.end('_setMetadataForRowRange_total');
}

/**
 * Clears bundle metadata from each row within a specified range.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {number} startRow The first row of the range to clear.
 * @param {number} endRow The last row of the range to clear.
 */
function _clearMetadataFromRowRange(sheet, startRow, endRow) {
  const sourceFile = "MetadataService_gs";
  ExecutionTimer.start('_clearMetadataFromRowRange_total');
  Log[sourceFile] (`[${sourceFile} - _clearMetadataFromRowRange] Start. Clearing metadata on rows ${startRow}-${endRow}.`);
  
  const range = sheet.getRange(startRow, 1, endRow - startRow + 1, 1);
  range.removeDeveloperMetadata(METADATA_KEY_BUNDLE); // This method works on ranges to remove metadata from all contained objects (rows).

  ExecutionTimer.end('_clearMetadataFromRowRange_total');
}

/**
 * Retrieves and parses the bundle information from a given range's ROW metadata.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range to check for metadata (typically the edited cell).
 * @returns {{bundleId: string, startRow: number, endRow: number}|null} The parsed bundle info object, or null if no metadata is found.
 */
function _getBundleInfoFromRange(range) {
  const sourceFile = "MetadataService_gs";
  ExecutionTimer.start('_getBundleInfoFromRange_total');
  
  const rowRange = range.getSheet().getRange(range.getRow(), 1);
  Log[sourceFile] (`[${sourceFile} - _getBundleInfoFromRange] Start. Reading metadata from row range ${rowRange.getA1Notation()}.`);
  
  const metadata = rowRange.getDeveloperMetadata();
  const bundleMeta = metadata.find(m => m.getKey() === METADATA_KEY_BUNDLE);

  if (bundleMeta) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: '_getBundleInfoFromRange_found' });
    try {
      const bundleInfo = JSON.parse(bundleMeta.getValue());
      Log[sourceFile] (`[${sourceFile} - _getBundleInfoFromRange] Found and parsed metadata: ${JSON.stringify(bundleInfo)}.`);
      ExecutionTimer.end('_getBundleInfoFromRange_total');
      return bundleInfo;
    } catch (e) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: '_getBundleInfoFromRange_error' });
      Log[sourceFile] (`[${sourceFile} - _getBundleInfoFromRange] ERROR: Failed to parse metadata JSON. Value: '${bundleMeta.getValue()}'. Error: ${e.message}`);
      ExecutionTimer.end('_getBundleInfoFromRange_total');
      return null;
    }
  }

  Log.TestCoverage_gs({ file: sourceFile, coverage: '_getBundleInfoFromRange_notFound' });
  Log[sourceFile] (`[${sourceFile} - _getBundleInfoFromRange] No bundle metadata found on this row.`);
  ExecutionTimer.end('_getBundleInfoFromRange_total');
  return null;
}

/**
 * --- NEW ---
 * Scans the entire sheet for valid bundles, clears all old bundle metadata,
 * and sets new metadata for all currently valid, multi-item bundles.
 * This function is designed to be called on open to initialize the sheet's state.
 * It is highly optimized to perform a single read and process data in memory.
 * REVISED: Now uses a robust, iterative method to find and remove all old metadata.
 */
function scanAndSetAllBundleMetadata() {
  const sourceFile = "MetadataService_gs";
  ExecutionTimer.start('scanAndSetAllBundleMetadata_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_start' });
  Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Start: Full sheet scan for bundle metadata initialization.`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_noDataRows' });
    Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] No data rows to process. Exiting.`);
    ExecutionTimer.end('scanAndSetAllBundleMetadata_total');
    return;
  }
  
  ExecutionTimer.start('scanAndSetAllBundleMetadata_clearMetadata');
  // --- THIS IS THE FIX ---
  // Use a robust, iterative finder-remover pattern when batch removal fails.
  const existingMetadata = sheet.createDeveloperMetadataFinder().withKey(METADATA_KEY_BUNDLE).find();
  if (existingMetadata && existingMetadata.length > 0) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_clearingExisting' });
    Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Found ${existingMetadata.length} existing metadata entries. Removing them individually.`);
    existingMetadata.forEach(meta => meta.remove());
  }
  Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Cleared all existing bundle metadata from the entire sheet.`);
  // --- END FIX ---
  ExecutionTimer.end('scanAndSetAllBundleMetadata_clearMetadata');


  ExecutionTimer.start('scanAndSetAllBundleMetadata_readSheet');
  const bundleNumCol = CONFIG.documentDeviceData.columnIndices.bundleNumber;
  const bundleColumnValues = sheet.getRange(startRow, bundleNumCol, lastRow - startRow + 1, 1).getValues();
  ExecutionTimer.end('scanAndSetAllBundleMetadata_readSheet');

  ExecutionTimer.start('scanAndSetAllBundleMetadata_groupInMemory');
  const bundlesMap = new Map();
  bundleColumnValues.forEach((val, index) => {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_loop_iteration' });
    const bundleNum = String(val[0] || '').trim();
    if (bundleNum) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_bundleNumFound' });
      if (!bundlesMap.has(bundleNum)) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_newBundleInMap' });
        bundlesMap.set(bundleNum, []);
      }
      bundlesMap.get(bundleNum).push({
        rowIndex: startRow + index
      });
    }
  });
  ExecutionTimer.end('scanAndSetAllBundleMetadata_groupInMemory');
  Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Grouped ${bundlesMap.size} unique bundle numbers in memory.`);

  ExecutionTimer.start('scanAndSetAllBundleMetadata_validateAndSet');
  for (const [bundleNum, rows] of bundlesMap.entries()) {
    if (rows.length <= 1) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_singleItemBundle' });
      continue; // Not a multi-item bundle, no metadata needed
    }

    // Since we're not validating term/quantity here (that's for onEdit), we only check for gaps.
    rows.sort((a, b) => a.rowIndex - b.rowIndex);
    let isConsecutive = true;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i].rowIndex !== rows[i - 1].rowIndex + 1) {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_gapDetected' });
        isConsecutive = false;
        break;
      }
    }

    if (isConsecutive) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_isValidConsecutive' });
      const bundleInfo = {
        bundleId: bundleNum,
        startRow: rows[0].rowIndex,
        endRow: rows[rows.length - 1].rowIndex
      };
      Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Found valid bundle #${bundleNum}. Setting metadata for rows ${bundleInfo.startRow}-${bundleInfo.endRow}.`);
      _setMetadataForRowRange(sheet, bundleInfo);
    }
  }
  ExecutionTimer.end('scanAndSetAllBundleMetadata_validateAndSet');

  Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] End: Full metadata scan and update complete.`);
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_end' });
  ExecutionTimer.end('scanAndSetAllBundleMetadata_total');
}