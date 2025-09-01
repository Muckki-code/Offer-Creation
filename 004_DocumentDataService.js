/**
 * @file This file contains the service for fetching and preparing all data
 * required for the document generation process.
 */

function showOfferDialog() {
  const sourceFile = "DocumentDataService_gs";
  ExecutionTimer.start('showOfferDialog_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'showOfferDialog_start' });
  Log[sourceFile](`[${sourceFile} - showOfferDialog] Start.`);
  
  ExecutionTimer.start('showOfferDialog_readSheet');
  const htmlTemplate = HtmlService.createTemplateFromFile('HTML/OfferForm');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const today = new Date();
  const defaultCreatedDateStr = today.getFullYear() + '-' + ('0' + (today.getMonth() + 1)).slice(-2) + '-' + ('0' + today.getDate()).slice(-2);
  const userEmail = Session.getActiveUser().getEmail();

  const sheetCompanyAddress = sheet.getRange(CONFIG.offerDetailsCells.companyAddress).getValue() || "";
  const sheetCustomerContactName = sheet.getRange(CONFIG.offerDetailsCells.customerContactName).getValue() || "";
  const sheetOfferValidUntilRaw = sheet.getRange(CONFIG.offerDetailsCells.offerValidUntil).getValue();
  const sheetSpecialAgreements = sheet.getRange(CONFIG.offerDetailsCells.specialAgreements).getValue() || "";
  const sheetYourName = sheet.getRange(CONFIG.offerDetailsCells.yourName).getValue() || "";
  const sheetYourPosition = sheet.getRange(CONFIG.offerDetailsCells.yourPosition).getValue() || "";
  const sheetContractTerm = sheet.getRange(CONFIG.offerDetailsCells.contractTerm).getValue() || "";
  const sheetCustomDocName = sheet.getRange(CONFIG.offerDetailsCells.documentName).getValue() || "";
  let sheetDocLanguage = (sheet.getRange(CONFIG.offerDetailsCells.language).getValue() || "german").toString().trim().toLowerCase();
  if (sheetDocLanguage !== "english" && sheetDocLanguage !== "german") {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'showOfferDialog_invalidLanguage' });
    sheetDocLanguage = "german";
  }
  let sheetOfferType = (sheet.getRange(CONFIG.offerDetailsCells.offerType).getValue() || "").toString().trim().toLowerCase();
  ExecutionTimer.end('showOfferDialog_readSheet');

  ExecutionTimer.start('showOfferDialog_prepareTemplate');
  htmlTemplate.formDataDefaults = {
    sheetCustomDocName: sheetCustomDocName,
    sheetDocLanguage: sheetDocLanguage,
    sheetOfferType: sheetOfferType,
    defaultCreatedDate: defaultCreatedDateStr,
    defaultUserEmail: userEmail,
    sheetCustomerCompanyAddress: sheetCompanyAddress,
    sheetCustomerContactName: sheetCustomerContactName,
    sheetOfferValidUntilDate: formatDateForLocale(sheetOfferValidUntilRaw, "english"),
    sheetSpecialAgreements: sheetSpecialAgreements,
    sheetEverphoneContactFullName: sheetYourName,
    sheetEverphoneContactPosition: sheetYourPosition,
    sheetOverallContractTermOptions: sheetContractTerm
  };
  Log[sourceFile](`[${sourceFile} - showOfferDialog] Prepared formDataDefaults for the dialog.`);

  const dialogTitle = (sheetDocLanguage === "german") ? "Angebot erstellen" : "Create Offer Document";
  const htmlOutput = htmlTemplate.evaluate().setWidth(650).setHeight(780);
  ExecutionTimer.end('showOfferDialog_prepareTemplate');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'showOfferDialog_end' });
  Log[sourceFile](`[${sourceFile} - showOfferDialog] End. Dialog displayed.`);
  ExecutionTimer.end('showOfferDialog_total');
}


function prepareDocumentData(formData) {
  const sourceFile = "DocumentDataService_gs";
  ExecutionTimer.start('prepareDocumentData_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'prepareDocumentData_start' });
  Log[sourceFile](`[${sourceFile} - prepareDocumentData] Start.`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const START_ROW = CONFIG.documentDeviceData.startRow;
  const lastRow = getLastLastRow(sheet);
  const dataBlockStartCol = CONFIG.documentDeviceData.columnIndices.sku;
  const dataBlockEndCol = CONFIG.maxDataColumn;
  const numColsInDataBlock = dataBlockEndCol - dataBlockStartCol + 1;
  Log[sourceFile](`[${sourceFile} - prepareDocumentData] Data grid definition: StartRow=${START_ROW}, StartCol=${dataBlockStartCol}, NumCols=${numColsInDataBlock}.`);

  let allDataRows = [];
  if (lastRow >= START_ROW) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'prepareDocumentData_hasData' });
    ExecutionTimer.start('prepareDocumentData_readSheet');
    const numRowsToRead = lastRow - START_ROW + 1;
    const rangeToRead = sheet.getRange(START_ROW, dataBlockStartCol, numRowsToRead, numColsInDataBlock);
    Log[sourceFile](`[${sourceFile} - prepareDocumentData] Reading data from range: ${rangeToRead.getA1Notation()}`);
    allDataRows = rangeToRead.getValues();
    ExecutionTimer.end('prepareDocumentData_readSheet');
  } else {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'prepareDocumentData_noData' });
    Log[sourceFile](`[${sourceFile} - prepareDocumentData] No data rows found to process.`);
  }
  Log[sourceFile](`[${sourceFile} - prepareDocumentData] Successfully read ${allDataRows.length} rows.`);
  
  ExecutionTimer.start('prepareDocumentData_groupItems');
  const groupedItems = groupApprovedItems(allDataRows, dataBlockStartCol);
  ExecutionTimer.end('prepareDocumentData_groupItems');
  
  const devicesData = [];
  let grandTotalNetMonthlyRentalPrice = 0;
  const c = { ...CONFIG.documentDeviceData.columnIndices, ...CONFIG.approvalWorkflow.columnIndices };

  ExecutionTimer.start('prepareDocumentData_processRows');
  for (const item of groupedItems) {
    if (item.isBundle) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'prepareDocumentData_handleBundle' });
      const totalNetMonthlyPriceForItem = getNumericValue(item.quantity) * getNumericValue(item.totalNetMonthlyPrice);
      grandTotalNetMonthlyRentalPrice += totalNetMonthlyPriceForItem;
      devicesData.push({
        model: item.models,
        quantity: item.quantity,
        term: item.term,
        netMonthlyRentalPrice: item.totalNetMonthlyPrice,
        totalNetMonthlyRentalPrice: totalNetMonthlyPriceForItem
      });
    } else {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'prepareDocumentData_handleIndividual' });
      const rowData = item.row;
      const quantity = getNumericValue(rowData[c.aeQuantity - dataBlockStartCol]);
      const approvedPrice = getNumericValue(rowData[c.financeApprovedPrice - dataBlockStartCol]);
      
      const totalNetMonthlyPriceForItem = quantity * approvedPrice;
      grandTotalNetMonthlyRentalPrice += totalNetMonthlyPriceForItem;
      devicesData.push({
          model: rowData[c.model - dataBlockStartCol],
          quantity: quantity,
          term: rowData[c.aeTerm - dataBlockStartCol],
          netMonthlyRentalPrice: approvedPrice,
          totalNetMonthlyRentalPrice: totalNetMonthlyPriceForItem
      });
    }
  }
  ExecutionTimer.end('prepareDocumentData_processRows');
  
  const customerCompanyName = sheet.getRange(CONFIG.offerDetailsCells.customerCompany).getValue();
  const docLanguage = (formData.docLanguage || "german").toString().trim().toLowerCase();
  
  const dataPackage = {
    formData: formData,
    devicesData: devicesData,
    grandTotal: grandTotalNetMonthlyRentalPrice,
    customerCompanyName: customerCompanyName,
    docLanguage: docLanguage
  };

  Log.TestCoverage_gs({ file: sourceFile, coverage: 'prepareDocumentData_end' });
  Log[sourceFile](`[${sourceFile} - prepareDocumentData] End. Data package assembled.`);
  ExecutionTimer.end('prepareDocumentData_total');

  return dataPackage;
}


function processFormSubmission(formData) {
  const sourceFile = "DocumentDataService_gs";
  ExecutionTimer.start('processFormSubmission_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'processFormSubmission_start' });
  Log[sourceFile](`[${sourceFile} - processFormSubmission] Start. Received form data.`);

  try {
    ExecutionTimer.start('processFormSubmission_updateSheet');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(CONFIG.offerDetailsCells.companyAddress).setValue(formData.customerCompanyAddress || "");
    sheet.getRange(CONFIG.offerDetailsCells.customerContactName).setValue(formData.customerContactName || "");
    sheet.getRange(CONFIG.offerDetailsCells.offerValidUntil).setValue(formData.offerValidUntilDate || "");
    sheet.getRange(CONFIG.offerDetailsCells.specialAgreements).setValue(formData.specialAgreementText || "");
    sheet.getRange(CONFIG.offerDetailsCells.yourName).setValue(formData.everphoneContactFullName || "");
    sheet.getRange(CONFIG.offerDetailsCells.yourPosition).setValue(formData.everphoneContactPosition || "");
    sheet.getRange(CONFIG.offerDetailsCells.contractTerm).setValue(formData.overallContractTermOptions || "");
    sheet.getRange(CONFIG.offerDetailsCells.documentName).setValue(formData.customDocName || "");
    sheet.getRange(CONFIG.offerDetailsCells.language).setValue(formData.docLanguage || "german");
    sheet.getRange(CONFIG.offerDetailsCells.offerType).setValue(formData.offerType || "binding");
    ExecutionTimer.end('processFormSubmission_updateSheet');

    const dataPackage = prepareDocumentData(formData);

    const newFileName = formData.customDocName || `Offer - ${dataPackage.customerCompanyName || 'Customer'} - ${(formData.offerCreatedDate || new Date().toISOString().slice(0, 10))}`;
    
    ExecutionTimer.start('processFormSubmission_createAndPopulateDoc');
    const copiedFile = createDocument(newFileName, dataPackage.docLanguage);
    populateDocContent(copiedFile, dataPackage);
    ExecutionTimer.end('processFormSubmission_createAndPopulateDoc');

    logDocumentActivity({
      action: "Document Created", docName: newFileName, docUrl: copiedFile.getUrl(),
      customerCompany: dataPackage.customerCompanyName, offerType: dataPackage.formData.offerType,
      language: dataPackage.docLanguage, details: { /* Redacted for brevity */ }
    });
    
    showSuccessDialog(copiedFile.getUrl(), `Document '${newFileName}' created!`);
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processFormSubmission_success' });

  } catch (error) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processFormSubmission_error' });
    Log[sourceFile](`[${sourceFile} - processFormSubmission] ERROR: ${error.message}. Stack: ${error.stack}`);
    throw new Error('Server-side error processing form: ' + error.message);
  } finally {
    ExecutionTimer.end('processFormSubmission_total');
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'processFormSubmission_end' });
    Log[sourceFile](`[${sourceFile} - processFormSubmission] End.`);
  }
}