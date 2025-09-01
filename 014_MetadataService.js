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
 * --- NEW PUBLIC FUNCTION ---
 * Scans the entire sheet to find all valid, multi-item bundles and writes their
 * structural information (startRow, endRow, id) as Developer Metadata to each
 * row within the valid bundle. This function also clears any stale bundle metadata.
 * This is the definitive source of truth for bundle structures.
 */
function scanAndSetAllBundleMetadata() {
  const sourceFile = "MetadataService_gs";
  ExecutionTimer.start('scanAndSetAllBundleMetadata_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_start' });
  Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Start: Beginning full-sheet scan to set bundle metadata.`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataStartRow = CONFIG.approvalWorkflow.startDataRow;
  const lastRow = sheet.getLastRow();

  // 1. Clear ALL existing bundle metadata from the entire sheet first for a clean slate.
  if (lastRow >= dataStartRow) {
    Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Clearing all existing bundle metadata from rows ${dataStartRow}-${lastRow}.`);
    
    // --- THIS IS THE FIX ---
    ExecutionTimer.start('scanAndSetAllBundleMetadata_clearOldMetadata');
    // Use the DeveloperMetadataFinder to locate all metadata with the bundle key and remove them individually.
    // This is more robust than the direct range-based removal method that was causing the error.
    const rangeToClear = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, sheet.getMaxColumns());
    const existingBundleMetadata = rangeToClear.createDeveloperMetadataFinder().withKey(METADATA_KEY_BUNDLE).find();

    if (existingBundleMetadata && existingBundleMetadata.length > 0) {
      Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Found ${existingBundleMetadata.length} old metadata entries to remove.`);
      existingBundleMetadata.forEach(meta => meta.remove());
    } else {
      Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] No old metadata entries found to remove.`);
    }
    ExecutionTimer.end('scanAndSetAllBundleMetadata_clearOldMetadata');
    // --- END FIX ---
    
    SpreadsheetApp.flush(); // Ensure metadata is cleared before we add new
  }

  // 2. Read all data in one go.
  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - dataBlockStartCol + 1;
  const allData = sheet.getRange(dataStartRow, dataBlockStartCol, lastRow - dataStartRow + 1, numCols).getValues();

  // 3. Group rows by bundle number in memory.
  const bundlesMap = new Map();
  for (let i = 0; i < allData.length; i++) {
    const rowData = allData[i];
    const bundleNum = String(rowData[CONFIG.documentDeviceData.columnIndices.bundleNumber - dataBlockStartCol] || '').trim();
    if (bundleNum) {
      if (!bundlesMap.has(bundleNum)) {
        bundlesMap.set(bundleNum, []);
      }
      bundlesMap.get(bundleNum).push({
        rowData: rowData,
        rowIndex: dataStartRow + i
      });
    }
  }

  // 4. Validate each bundle group and set metadata for the valid ones.
  Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Found ${bundlesMap.size} potential bundles. Validating each...`);
  for (const [bundleNum, rows] of bundlesMap.entries()) {
    if (rows.length <= 1) continue; // Not a multi-item bundle

    // A. Check for non-consecutive rows
    rows.sort((a, b) => a.rowIndex - b.rowIndex);
    let isConsecutive = true;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i].rowIndex !== rows[i - 1].rowIndex + 1) {
        isConsecutive = false;
        break;
      }
    }
    if (!isConsecutive) {
        Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Bundle #${bundleNum} is INVALID (non-consecutive). Skipping metadata.`);
        continue;
    }

    // B. Check for mismatched Term or Quantity
    const termColIndex = CONFIG.approvalWorkflow.columnIndices.aeTerm - dataBlockStartCol;
    const qtyColIndex = CONFIG.approvalWorkflow.columnIndices.aeQuantity - dataBlockStartCol;
    const expectedTerm = rows[0].rowData[termColIndex];
    const expectedQty = rows[0].rowData[qtyColIndex];
    let hasMismatch = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i].rowData[termColIndex]) !== String(expectedTerm) || String(rows[i].rowData[qtyColIndex]) !== String(expectedQty)) {
        hasMismatch = true;
        break;
      }
    }
     if (hasMismatch) {
        Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Bundle #${bundleNum} is INVALID (mismatched term/qty). Skipping metadata.`);
        continue;
     }

    // C. If all checks pass, it's a valid bundle. Set the metadata.
    Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] Bundle #${bundleNum} is VALID. Setting metadata.`);
    const bundleInfo = {
      bundleId: bundleNum,
      startRow: rows[0].rowIndex,
      endRow: rows[rows.length - 1].rowIndex
    };
    _setMetadataForRowRange(sheet, bundleInfo);
  }

  Log[sourceFile](`[${sourceFile} - scanAndSetAllBundleMetadata] End: Scan and metadata update complete.`);
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'scanAndSetAllBundleMetadata_end' });
  ExecutionTimer.end('scanAndSetAllBundleMetadata_total');
}