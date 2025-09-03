/**
 * @file This file contains the service for fetching and preparing all data
 * required for the document generation process.
 */

let _skippedBundlesForDocGen = [];

function showOfferDialog() {
  const sourceFile = "DocumentDataService_gs";
  ExecutionTimer.start("showOfferDialog_total");
  Log.TestCoverage_gs({ file: sourceFile, coverage: "showOfferDialog_start" });
  Log[sourceFile](`[${sourceFile} - showOfferDialog] Start.`);

  ExecutionTimer.start("showOfferDialog_readSheet");
  const htmlTemplate = HtmlService.createTemplateFromFile("HTML/OfferForm");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const today = new Date();
  const defaultCreatedDateStr =
    today.getFullYear() +
    "-" +
    ("0" + (today.getMonth() + 1)).slice(-2) +
    "-" +
    ("0" + today.getDate()).slice(-2);
  const userEmail = Session.getActiveUser().getEmail();

  const staticValues = _getStaticSheetValues(sheet);

  // 2. Use the values from the cache
  const sheetCompanyAddress = staticValues.companyAddress;
  const sheetCustomerContactName = staticValues.customerContactName;
  const sheetOfferValidUntilRaw = staticValues.offerValidUntil;
  const sheetSpecialAgreements = staticValues.specialAgreements;
  const sheetYourName = staticValues.yourName;
  const sheetYourPosition = staticValues.yourPosition;
  const sheetContractTerm = staticValues.contractTerm;
  const sheetCustomDocName = staticValues.documentName;
  let sheetDocLanguage = staticValues.language;
  let sheetOfferType = staticValues.offerType;

  // 3. Perform the validation on the cached value
  if (sheetDocLanguage !== "english" && sheetDocLanguage !== "german") {
    sheetDocLanguage = "german"; // Default to german if invalid
  }

  ExecutionTimer.end("showOfferDialog_readSheet");

  ExecutionTimer.start("showOfferDialog_prepareTemplate");
  htmlTemplate.formDataDefaults = {
    sheetCustomDocName: sheetCustomDocName,
    sheetDocLanguage: sheetDocLanguage,
    sheetOfferType: sheetOfferType,
    defaultCreatedDate: defaultCreatedDateStr,
    defaultUserEmail: userEmail,
    sheetCustomerCompanyAddress: sheetCompanyAddress,
    sheetCustomerContactName: sheetCustomerContactName,
    sheetOfferValidUntilDate: formatDateForLocale(
      sheetOfferValidUntilRaw,
      "english"
    ),
    sheetSpecialAgreements: sheetSpecialAgreements,
    sheetEverphoneContactFullName: sheetYourName,
    sheetEverphoneContactPosition: sheetYourPosition,
    sheetOverallContractTermOptions: sheetContractTerm,
  };
  Log[sourceFile](
    `[${sourceFile} - showOfferDialog] Prepared formDataDefaults for the dialog.`
  );

  const dialogTitle =
    sheetDocLanguage === "german"
      ? "Angebot erstellen"
      : "Create Offer Document";
  const htmlOutput = htmlTemplate
    .evaluate()
    .setWidth(650)
    .setHeight(780);
  ExecutionTimer.end("showOfferDialog_prepareTemplate");

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
  Log.TestCoverage_gs({ file: sourceFile, coverage: "showOfferDialog_end" });
  Log[sourceFile](`[${sourceFile} - showOfferDialog] End. Dialog displayed.`);
  ExecutionTimer.end("showOfferDialog_total");
}

function prepareDocumentData(formData) {
  const sourceFile = "DocumentDataService_gs";
  ExecutionTimer.start("prepareDocumentData_total");
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "prepareDocumentData_start",
  });
  Log[sourceFile](`[${sourceFile} - prepareDocumentData] Start.`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const START_ROW = CONFIG.documentDeviceData.startRow;
  const lastRow = getLastLastRow(sheet);
  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const dataBlockEndCol = CONFIG.maxDataColumn;
  const numColsInDataBlock = dataBlockEndCol - dataBlockStartCol + 1;
  Log[sourceFile](
    `[${sourceFile} - prepareDocumentData] Data grid definition: StartRow=${START_ROW}, StartCol=${dataBlockStartCol}, NumCols=${numColsInDataBlock}.`
  );

  let allDataRows = [];
  if (lastRow >= START_ROW) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "prepareDocumentData_hasData",
    });
    ExecutionTimer.start("prepareDocumentData_readSheet");
    const numRowsToRead = lastRow - START_ROW + 1;
    const rangeToRead = sheet.getRange(
      START_ROW,
      dataBlockStartCol,
      numRowsToRead,
      numColsInDataBlock
    );
    Log[sourceFile](
      `[${sourceFile} - prepareDocumentData] Reading data from range: ${rangeToRead.getA1Notation()}`
    );
    allDataRows = rangeToRead.getValues();
    ExecutionTimer.end("prepareDocumentData_readSheet");
  } else {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "prepareDocumentData_noData",
    });
    Log[sourceFile](
      `[${sourceFile} - prepareDocumentData] No data rows found to process.`
    );
  }
  Log[sourceFile](
    `[${sourceFile} - prepareDocumentData] Successfully read ${allDataRows.length} rows.`
  );

  ExecutionTimer.start("prepareDocumentData_groupItems");
  const groupedItems = groupApprovedItems(allDataRows, dataBlockStartCol);
  ExecutionTimer.end("prepareDocumentData_groupItems");

  const devicesData = [];
  let grandTotalNetMonthlyRentalPrice = 0;
  const c = {
    ...CONFIG.documentDeviceData.columnIndices,
    ...CONFIG.approvalWorkflow.columnIndices,
  };

  ExecutionTimer.start("prepareDocumentData_processRows");
  for (const item of groupedItems) {
    if (item.isBundle) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "prepareDocumentData_handleBundle",
      });
      const totalNetMonthlyPriceForItem =
        getNumericValue(item.quantity) *
        getNumericValue(item.totalNetMonthlyPrice);
      grandTotalNetMonthlyRentalPrice += totalNetMonthlyPriceForItem;
      devicesData.push({
        model: item.models,
        quantity: item.quantity,
        term: item.term,
        netMonthlyRentalPrice: item.totalNetMonthlyPrice,
        totalNetMonthlyRentalPrice: totalNetMonthlyPriceForItem,
      });
    } else {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "prepareDocumentData_handleIndividual",
      });
      const rowData = item.row;
      const quantity = getNumericValue(
        rowData[c.aeQuantity - dataBlockStartCol]
      );
      const approvedPrice = getNumericValue(
        rowData[c.financeApprovedPrice - dataBlockStartCol]
      );

      const totalNetMonthlyPriceForItem = quantity * approvedPrice;
      grandTotalNetMonthlyRentalPrice += totalNetMonthlyPriceForItem;
      devicesData.push({
        model: rowData[c.model - dataBlockStartCol],
        quantity: quantity,
        term: rowData[c.aeTerm - dataBlockStartCol],
        netMonthlyRentalPrice: approvedPrice,
        totalNetMonthlyRentalPrice: totalNetMonthlyPriceForItem,
      });
    }
  }
  ExecutionTimer.end("prepareDocumentData_processRows");

  const staticValues = _getStaticSheetValues(sheet);
  const customerCompanyName = staticValues.customerCompany;
  const docLanguage = (formData.docLanguage || "german")
    .toString()
    .trim()
    .toLowerCase();

  const dataPackage = {
    formData: formData,
    devicesData: devicesData,
    grandTotal: grandTotalNetMonthlyRentalPrice,
    customerCompanyName: customerCompanyName,
    docLanguage: docLanguage,
  };

  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "prepareDocumentData_end",
  });
  Log[sourceFile](
    `[${sourceFile} - prepareDocumentData] End. Data package assembled.`
  );
  ExecutionTimer.end("prepareDocumentData_total");

  return dataPackage;
}

function processFormSubmission(formData) {
  const sourceFile = "DocumentDataService_gs";
  ExecutionTimer.start("processFormSubmission_total");
  Log.TestCoverage_gs({
    file: sourceFile,
    coverage: "processFormSubmission_start",
  });
  Log[sourceFile](
    `[${sourceFile} - processFormSubmission] Start. Received form data.`
  );

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // --- NEW EFFICIENT HEADER UPDATE LOGIC ---

    // 1. Get current header values from the multi-layer cache.
    const currentValues = _getStaticSheetValues(sheet);

    // 2. Prepare an object with the new values from the form, matching the cache structure.
    const newValues = {
      ...currentValues, // Start with current values to preserve unedited fields like 'approver'
      companyAddress: formData.customerCompanyAddress || "",
      customerContactName: formData.customerContactName || "",
      offerValidUntil: formData.offerValidUntilDate || "",
      specialAgreements: formData.specialAgreementText || "",
      yourName: formData.everphoneContactFullName || "",
      yourPosition: formData.everphoneContactPosition || "",
      contractTerm: formData.overallContractTermOptions || "",
      documentName: formData.customDocName || "",
      language: formData.docLanguage || "german",
      offerType: formData.offerType || "binding",
    };

    // 3. Compare the current and new values. If no change, skip the slow sheet write.
    if (JSON.stringify(currentValues) === JSON.stringify(newValues)) {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "processFormSubmission_noHeaderChange",
      });
      Log[sourceFile](
        "[processFormSubmission] No changes detected in header data. Skipping sheet write."
      );
    } else {
      Log.TestCoverage_gs({
        file: sourceFile,
        coverage: "processFormSubmission_headerChanged",
      });
      Log[sourceFile](
        "[processFormSubmission] Header data has changed. Updating sheet in a single operation."
      );
      ExecutionTimer.start("processFormSubmission_updateSheet");

      // 4. Read the existing 2D array to preserve all formatting and non-data cells.
      const headerRange = sheet.getRange(
        CONFIG.offerDetailsCells.cachedHeaderRangeA1
      );
      const dataForSheet = headerRange.getValues();

      // 5. Update the 2D array in memory with the new values at their correct positions.
      // Note: These indices are relative to the start of the cachedHeaderRangeA1 ("F1").
      dataForSheet[0][1] = newValues.customerCompany; // G1
      dataForSheet[0][3] = newValues.language; // I1
      dataForSheet[0][6] = newValues.telekomDeal; // L1
      dataForSheet[0][9] = newValues.approver; // O1
      dataForSheet[1][1] = newValues.customerContactName; // G2
      dataForSheet[1][3] = newValues.offerType; // I2
      dataForSheet[1][6] = newValues.yourName; // L2
      dataForSheet[2][1] = newValues.companyAddress; // G3
      dataForSheet[2][3] = newValues.contractTerm; // I3
      dataForSheet[2][6] = newValues.yourPosition; // L3
      dataForSheet[3][1] = newValues.specialAgreements; // G4
      dataForSheet[3][3] = newValues.offerValidUntil; // I4
      dataForSheet[3][6] = newValues.documentName; // L4

      // 6. Perform the single, efficient write operation.
      headerRange.setValues(dataForSheet);

      // 7. Eagerly refresh both caches with the new data without another read.
      const cache = CacheService.getScriptCache();
      cache.put("staticSheetValues", JSON.stringify(newValues), 21600); // 6 hours
      _staticValuesCache = newValues; // Update the execution cache

      ExecutionTimer.end("processFormSubmission_updateSheet");
    }

    // --- ALL SUBSEQUENT LOGIC IS IDENTICAL TO YOUR ORIGINAL VERSION ---

    const dataPackage = prepareDocumentData(formData);

    const newFileName =
      formData.customDocName ||
      `Offer - ${dataPackage.customerCompanyName ||
        "Customer"} - ${formData.offerCreatedDate ||
        new Date().toISOString().slice(0, 10)}`;

    ExecutionTimer.start("processFormSubmission_createAndPopulateDoc");
    const copiedFile = createDocument(newFileName, dataPackage.docLanguage);
    populateDocContent(copiedFile, dataPackage);
    ExecutionTimer.end("processFormSubmission_createAndPopulateDoc");

    logDocumentActivity({
      action: "Document Created",
      docName: newFileName,
      docUrl: copiedFile.getUrl(),
      customerCompany: dataPackage.customerCompanyName,
      offerType: dataPackage.formData.offerType,
      language: dataPackage.docLanguage,
      details: {
        /* Redacted for brevity */
      },
    });

    showSuccessDialog(
      copiedFile.getUrl(),
      `Document '${newFileName}' created!`
    );
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "processFormSubmission_success",
    });
  } catch (error) {
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "processFormSubmission_error",
    });
    Log[sourceFile](
      `[${sourceFile} - processFormSubmission] ERROR: ${error.message}. Stack: ${error.stack}`
    );
    throw new Error("Server-side error processing form: " + error.message);
  } finally {
    ExecutionTimer.end("processFormSubmission_total");
    Log.TestCoverage_gs({
      file: sourceFile,
      coverage: "processFormSubmission_end",
    });
    Log[sourceFile](`[${sourceFile} - processFormSubmission] End.`);
  }
}
