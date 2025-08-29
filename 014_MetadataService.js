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