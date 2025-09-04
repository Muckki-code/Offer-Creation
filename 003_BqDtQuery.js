/**
 * @fileoverview Manages integration with Google BigQuery to fetch
 * product pricing and device details based on SKUs.
 */

// In BqDtQuery.gs

/**
 * Fetches product pricing and device details from BigQuery based on SKUs
 * entered in the sheet and updates the relevant columns.
 * This function is triggered via a custom menu item.
 */
function getDataFromSKU() {
  const sourceFile = "BqDtQuery_gs";
  ExecutionTimer.start("getDataFromSKU_total");
  Log.TestCoverage_gs({ file: sourceFile, coverage: "getDataFromSKU_start" });
  Log[sourceFile](`[${sourceFile} - getDataFromSKU] Start.`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const bqSettings = CONFIG.bqQuerySettings;
  const docDataConfig = CONFIG.documentDeviceData;

  const skuColIndex = bqSettings.skuColumnIndex;
  const dataStartRow = bqSettings.scriptStartRow;

  ExecutionTimer.start("getDataFromSKU_readSKUs");
  const lastRowWithSku = getLastPopulatedRowInColumn(sheet, skuColIndex);
  Log[sourceFile](
    `[${sourceFile} - getDataFromSKU] DEBUG: Data Start Row: ${dataStartRow}, Last Populated Row in SKU Column: ${lastRowWithSku}`
  );

  let uniqueSkusToQuery = new Set();
  if (lastRowWithSku >= dataStartRow) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "getDataFromSKU_skusFound",
    });
    const skuRange = sheet.getRange(
      dataStartRow,
      skuColIndex,
      lastRowWithSku - dataStartRow + 1,
      1
    );
    const skuValues = skuRange.getValues();
    Log[sourceFile](
      `[${sourceFile} - getDataFromSKU] Info: Identified rows ${dataStartRow} to ${lastRowWithSku} for SKU extraction.`
    );

    for (let i = 0; i < skuValues.length; i++) {
      const skuCandidate = String(skuValues[i][0] || "").trim();
      if (skuCandidate !== "") {
        uniqueSkusToQuery.add(skuCandidate);
      }
    }
  } else {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "getDataFromSKU_noSkusFound",
    });
    Log[sourceFile](
      `[${sourceFile} - getDataFromSKU] Condition: No SKUs found below data start row (${dataStartRow}).`
    );
    ui.alert(
      "No SKUs found in column " + bqSettings.skuColumnLetter + " to process."
    );
    ExecutionTimer.end("getDataFromSKU_readSKUs");
    ExecutionTimer.end("getDataFromSKU_total");
    return;
  }
  ExecutionTimer.end("getDataFromSKU_readSKUs");
  Log[sourceFile](
    `[${sourceFile} - getDataFromSKU] Info: Found ${uniqueSkusToQuery.size} unique SKUs to query.`
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Fetching data for ${uniqueSkusToQuery.size} SKUs...`,
    "BigQuery Fetch",
    -1
  );

  let queryResults;
  try {
    ExecutionTimer.start("getDataFromSKU_performBqQuery");
    queryResults = performBqQuery(Array.from(uniqueSkusToQuery));
    ExecutionTimer.end("getDataFromSKU_performBqQuery");
    Log[sourceFile](
      `[${sourceFile} - getDataFromSKU] Info: BigQuery query successful. Received ${queryResults.size} results.`
    );
  } catch (e) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "getDataFromSKU_bqError",
    });
    ui.alert("Error querying BigQuery: " + e.message);
    Log[sourceFile](
      `[${sourceFile} - getDataFromSKU] ERROR: Error querying BigQuery: ${e.message}. Stack: ${e.stack}`
    );
    return;
  } finally {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "getDataFromSKU_finallyBlock",
    });
    SpreadsheetApp.getActiveSpreadsheet().toast("", "BigQuery Fetch");
  }

  const actualLastSheetRow = sheet.getLastRow();
  const dataBlockStartCol = docDataConfig.columnIndices.sku;
  const dataBlockEndCol = CONFIG.maxDataColumn;
  const numColsInDataBlock = dataBlockEndCol - dataBlockStartCol + 1;
  Log[sourceFile](
    `[${sourceFile} - getDataFromSKU] Data block defined from column ${dataBlockStartCol} to ${dataBlockEndCol} (${numColsInDataBlock} columns).`
  );

  ExecutionTimer.start("getDataFromSKU_readSheetData");
  const dataRangeForProcessing = sheet.getRange(
    dataStartRow,
    dataBlockStartCol,
    actualLastSheetRow - dataStartRow + 1,
    numColsInDataBlock
  );
  const allDataBlockValuesBefore = dataRangeForProcessing.getValues();
  const allDataBlockValuesAfter = JSON.parse(
    JSON.stringify(allDataBlockValuesBefore)
  );
  ExecutionTimer.end("getDataFromSKU_readSheetData");
  Log[sourceFile](
    `[${sourceFile} - getDataFromSKU] Captured 'before' state of data block: ${dataRangeForProcessing.getA1Notation()}.`
  );

  const staticValues = _getStaticSheetValues(sheet);
  const colIndexes = {
    ...CONFIG.documentDeviceData.columnIndices,
    ...CONFIG.approvalWorkflow.columnIndices,
  };
  const wfConfig = CONFIG.approvalWorkflow;
  const statusStrings = wfConfig.statusStrings;

  const wasBqDerivedPopulatedBefore = (rowToCheck, currentStartCol) => {
    return (
      String(rowToCheck[colIndexes.epCapex - currentStartCol] || "").trim() !==
        "" ||
      String(
        rowToCheck[colIndexes.ep24PriceTarget - currentStartCol] || ""
      ).trim() !== "" ||
      String(
        rowToCheck[colIndexes.ep36PriceTarget - currentStartCol] || ""
      ).trim() !== "" ||
      String(rowToCheck[colIndexes.tkCapex - currentStartCol] || "").trim() !==
        "" ||
      String(
        rowToCheck[colIndexes.tk24PriceTarget - currentStartCol] || ""
      ).trim() !== "" ||
      String(
        rowToCheck[colIndexes.tk36PriceTarget - currentStartCol] || ""
      ).trim() !== "" ||
      String(rowToCheck[colIndexes.model - currentStartCol] || "").trim() !== ""
    );
  };

  let nextIndex = null; // Initialize for index assignment

  ExecutionTimer.start("getDataFromSKU_processRows");
  for (let i = 0; i < allDataBlockValuesAfter.length; i++) {
    const currentRowNum = dataStartRow + i;
    const currentRowValues = allDataBlockValuesAfter[i];
    const originalRowValuesBefore = allDataBlockValuesBefore[i];

    const skuInRow = String(
      currentRowValues[colIndexes.sku - dataBlockStartCol] || ""
    ).trim();
    const skuInRowBefore = String(
      originalRowValuesBefore[colIndexes.sku - dataBlockStartCol] || ""
    ).trim();
    const modelInRowBefore = String(
      originalRowValuesBefore[colIndexes.model - dataBlockStartCol] || ""
    ).trim();

    let needsStatusUpdate = false;

    if (skuInRow !== "" && queryResults.has(skuInRow)) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "getDataFromSKU_updateRow",
      });
      const bqData = queryResults.get(skuInRow);

      if (bqData && bqData.Model && String(bqData.Model).trim() !== "") {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "getDataFromSKU_validBqData",
        });
        currentRowValues[colIndexes.epCapex - dataBlockStartCol] = isNaN(
          parseFloat(bqData.epCapex)
        )
          ? ""
          : parseFloat(bqData.epCapex);
        currentRowValues[
          colIndexes.ep24PriceTarget - dataBlockStartCol
        ] = isNaN(parseFloat(bqData.ep24PriceTarget))
          ? ""
          : parseFloat(bqData.ep24PriceTarget);
        currentRowValues[
          colIndexes.ep36PriceTarget - dataBlockStartCol
        ] = isNaN(parseFloat(bqData.ep36PriceTarget))
          ? ""
          : parseFloat(bqData.ep36PriceTarget);
        currentRowValues[colIndexes.tkCapex - dataBlockStartCol] = isNaN(
          parseFloat(bqData.tkCapex)
        )
          ? ""
          : parseFloat(bqData.tkCapex);
        currentRowValues[
          colIndexes.tk24PriceTarget - dataBlockStartCol
        ] = isNaN(parseFloat(bqData.tk24PriceTarget))
          ? ""
          : parseFloat(bqData.tk24PriceTarget);
        currentRowValues[
          colIndexes.tk36PriceTarget - dataBlockStartCol
        ] = isNaN(parseFloat(bqData.tk36PriceTarget))
          ? ""
          : parseFloat(bqData.tk36PriceTarget);

        if (
          modelInRowBefore === "" ||
          modelInRowBefore !== String(bqData.Model).trim()
        ) {
          Log.TestCoverage_gs({
            file: sourceFile,
            coverage: "getDataFromSKU_modelNeedsUpdate",
          });
          currentRowValues[colIndexes.model - dataBlockStartCol] = bqData.Model;
          needsStatusUpdate = true;
        }
      } else {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "getDataFromSKU_invalidBqData",
        });
        if (
          wasBqDerivedPopulatedBefore(
            originalRowValuesBefore,
            dataBlockStartCol
          )
        ) {
          currentRowValues[colIndexes.model - dataBlockStartCol] = "";
          currentRowValues[colIndexes.epCapex - dataBlockStartCol] = "";
          currentRowValues[colIndexes.ep24PriceTarget - dataBlockStartCol] = "";
          currentRowValues[colIndexes.ep36PriceTarget - dataBlockStartCol] = "";
          currentRowValues[colIndexes.tkCapex - dataBlockStartCol] = "";
          currentRowValues[colIndexes.tk24PriceTarget - dataBlockStartCol] = "";
          currentRowValues[colIndexes.tk36PriceTarget - dataBlockStartCol] = "";
          currentRowValues[colIndexes.lrfPreview - dataBlockStartCol] = "";
          currentRowValues[
            colIndexes.contractValuePreview - dataBlockStartCol
          ] = "";
          needsStatusUpdate = true;
        }
      }

      const modelName = currentRowValues[colIndexes.model - dataBlockStartCol];
      const currentIndex =
        currentRowValues[colIndexes.index - dataBlockStartCol];
      if (modelName && !currentIndex) {
        if (nextIndex === null) {
          nextIndex = getNextAvailableIndex(sheet);
        }
        currentRowValues[colIndexes.index - dataBlockStartCol] = nextIndex++;
      }

      // --- THIS IS THE FIX ---
      // 1. Call the pure function with its new, correct signature.
      const calculatedValues = updateCalculationsForRow(
        currentRowValues, // The in-memory row data
        colIndexes, // The combined column indices from CONFIG
        wfConfig, // The CONFIG.approvalWorkflow object
        dataBlockStartCol // The 1-based start column of the data block
      );

      // 2. Write the returned values back into the in-memory row array.
      currentRowValues[colIndexes.lrfPreview - dataBlockStartCol] =
        calculatedValues.lrfPreview;
      currentRowValues[colIndexes.contractValuePreview - dataBlockStartCol] =
        calculatedValues.contractValuePreview;
      // --- END FIX ---
    } else if (
      skuInRow === "" &&
      skuInRowBefore !== "" &&
      wasBqDerivedPopulatedBefore(originalRowValuesBefore, dataBlockStartCol)
    ) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "getDataFromSKU_clearRow",
      });
      currentRowValues[colIndexes.model - dataBlockStartCol] = "";
      currentRowValues[colIndexes.epCapex - dataBlockStartCol] = "";
      currentRowValues[colIndexes.ep24PriceTarget - dataBlockStartCol] = "";
      currentRowValues[colIndexes.ep36PriceTarget - dataBlockStartCol] = "";
      currentRowValues[colIndexes.tkCapex - dataBlockStartCol] = "";
      currentRowValues[colIndexes.tk24PriceTarget - dataBlockStartCol] = "";
      currentRowValues[colIndexes.tk36PriceTarget - dataBlockStartCol] = "";
      currentRowValues[colIndexes.lrfPreview - dataBlockStartCol] = "";
      currentRowValues[colIndexes.contractValuePreview - dataBlockStartCol] =
        "";
      needsStatusUpdate = true;
    }

    if (needsStatusUpdate) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "getDataFromSKU_statusUpdateNeeded",
      });
      const initialStatus =
        originalRowValuesBefore[colIndexes.status - dataBlockStartCol] || "";
      const newStatus = updateStatusForRow(
        currentRowValues,
        originalRowValuesBefore,
        staticValues.isTelekomDeal,
        {},
        dataBlockStartCol,
        colIndexes
      );

      if (newStatus !== initialStatus) {
        Log.TestCoverage_gs({
          file: sourceFile,
          coverage: "getDataFromSKU_statusHasChanged",
        });
        if (newStatus === null) {
          currentRowValues[colIndexes.status - dataBlockStartCol] = "";
        } else {
          currentRowValues[colIndexes.status - dataBlockStartCol] = newStatus;
          if (
            [
              statusStrings.pending,
              statusStrings.draft,
              statusStrings.revisedByAE,
            ].includes(newStatus)
          ) {
            currentRowValues[colIndexes.approverAction - dataBlockStartCol] =
              "Choose Action";
          }
          const approvedStatuses = [
            statusStrings.approvedOriginal,
            statusStrings.approvedNew,
          ];
          if (
            approvedStatuses.includes(initialStatus) &&
            !approvedStatuses.includes(newStatus)
          ) {
            Log.TestCoverage_gs({
              file: sourceFile,
              coverage: "getDataFromSKU_clearApprovalFields",
            });
            currentRowValues[
              colIndexes.financeApprovedPrice - dataBlockStartCol
            ] = "";
            currentRowValues[colIndexes.approvedBy - dataBlockStartCol] = "";
            currentRowValues[colIndexes.approvalDate - dataBlockStartCol] = "";
          }
        }
      }
    }
  }
  ExecutionTimer.end("getDataFromSKU_processRows");

  ExecutionTimer.start("getDataFromSKU_writeSheet");
  dataRangeForProcessing.setValues(allDataBlockValuesAfter);
  ExecutionTimer.end("getDataFromSKU_writeSheet");
  Log[sourceFile](
    `[${sourceFile} - getDataFromSKU] Info: BigQuery data fetch and alignment complete. Data written to sheet.`
  );

  ExecutionTimer.start("getDataFromSKU_logChanges");
  for (let i = 0; i < allDataBlockValuesBefore.length; i++) {
    const currentRowNum = dataStartRow + i;
    const oldRow = allDataBlockValuesBefore[i];
    const newRow = allDataBlockValuesAfter[i];

    if (JSON.stringify(oldRow) !== JSON.stringify(newRow)) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "getDataFromSKU_logRowChange",
      });
      const oldStatus = oldRow[colIndexes.status - dataBlockStartCol];
      const newStatus = newRow[colIndexes.status - dataBlockStartCol];

      if (newStatus !== oldStatus) {
        logTableActivity({
          mainSheet: sheet,
          rowNum: currentRowNum,
          oldStatus: oldStatus,
          newStatus: newStatus,
          currentFullRowValues: newRow,
          originalFullRowValues: oldRow,
          startCol: dataBlockStartCol,
        });
      }
    }
  }
  ExecutionTimer.end("getDataFromSKU_logChanges");

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "BigQuery data import and sheet update complete!",
    "Success",
    5
  );
  Log.TestCoverage_gs({ file: sourceFile, coverage: "getDataFromSKU_end" });
  Log[sourceFile](`[${sourceFile} - getDataFromSKU] End.`);
  ExecutionTimer.end("getDataFromSKU_total");
}

/**
 * Performs a BigQuery query to fetch product data based on SKUs.
 */
function performBqQuery(uniqueSkusToQuery) {
  const sourceFile = "BqDtQuery_gs";
  ExecutionTimer.start("performBqQuery_total");
  Log.TestCoverage_gs({ file: sourceFile, coverage: "performBqQuery_start" });
  Log[sourceFile](
    `[${sourceFile} - performBqQuery] Start: Querying for ${uniqueSkusToQuery.length} SKUs.`
  );

  const bqResultsMap = new Map();
  if (!uniqueSkusToQuery || uniqueSkusToQuery.length === 0) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "performBqQuery_noSkus",
    });
    Log[sourceFile](
      `[${sourceFile} - performBqQuery] Condition: No unique SKUs to query. Returning empty map.`
    );
    ExecutionTimer.end("performBqQuery_total");
    return bqResultsMap;
  }

  const bqSettings = CONFIG.bqQuerySettings;
  const projectId = bqSettings.projectId;
  const tableName = bqSettings.tableName;

  const skuListForQuery = uniqueSkusToQuery
    .map((sku) => parseInt(sku, 10))
    .filter((sku) => !isNaN(sku))
    .join(",");
  if (!skuListForQuery) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "performBqQuery_noValidSkus",
    });
    throw new Error("No valid numeric SKUs to build BigQuery 'IN' clause.");
  }

  const bqQuery = `SELECT
      device_configuration_id as SKU,
      name as Model,
      ep_sourcing_price as epCapex,
      rent_target_price_EnterpriseA_24_500plus as ep24PriceTarget,
      rent_target_price_EnterpriseA_36_500plus as ep36PriceTarget,
      tk_sourcing_price as tkCapex,
      rent_target_price_TelekomGermany_24 as tk24PriceTarget,
      rent_target_price_TelekomGermany_36 as tk36PriceTarget
    FROM \`${tableName}\`
    WHERE device_configuration_id IN (${skuListForQuery})`;
  Log[sourceFile](
    `[${sourceFile} - performBqQuery] Info: Executing BigQuery Query: ${bqQuery}`
  );

  const request = { query: bqQuery, useLegacySql: false };
  let queryJob;
  try {
    queryJob = BigQuery.Jobs.query(request, projectId);
  } catch (e) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "performBqQuery_jobError",
    });
    throw new Error("Error initiating BigQuery job: " + e.message);
  }

  ExecutionTimer.start("performBqQuery_waitForJob");
  let jobComplete = false;
  let jobResults;
  const MAX_WAIT_TIME_SECONDS = 60;
  const START_TIME = new Date().getTime();
  while (
    !jobComplete &&
    (new Date().getTime() - START_TIME) / 1000 < MAX_WAIT_TIME_SECONDS
  ) {
    Utilities.sleep(1000);
    jobResults = BigQuery.Jobs.getQueryResults(
      projectId,
      queryJob.jobReference.jobId
    );
    jobComplete = jobResults.jobComplete;
  }
  ExecutionTimer.end("performBqQuery_waitForJob");

  if (!jobComplete) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "performBqQuery_jobTimeout",
    });
    throw new Error("BigQuery job timed out.");
  }

  if (jobResults.rows) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "performBqQuery_hasResults",
    });
    Log[sourceFile](
      `[${sourceFile} - performBqQuery] Info: Processing ${jobResults.rows.length} rows from BigQuery results.`
    );
    jobResults.rows.forEach((row) => {
      const fields = row.f;
      const sku = String(fields[0].v).trim();
      bqResultsMap.set(sku, {
        model: fields[1].v,
        epCapex: fields[2].v,
        ep24PriceTarget: fields[3].v,
        ep36PriceTarget: fields[4].v,
        tkCapex: fields[5].v,
        tk24PriceTarget: fields[6].v,
        tk36PriceTarget: fields[7].v,
      });
    });
  } else {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "performBqQuery_noResults",
    });
    Log[sourceFile](
      `[${sourceFile} - performBqQuery] Info: No rows returned from BigQuery job.`
    );
  }

  Log.TestCoverage_gs({ file: sourceFile, coverage: "performBqQuery_end" });
  Log[sourceFile](
    `[${sourceFile} - performBqQuery] End: Returning BQ results map with ${bqResultsMap.size} entries.`
  );
  ExecutionTimer.end("performBqQuery_total");
  return bqResultsMap;
}
