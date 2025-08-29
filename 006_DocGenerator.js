// In DocGenerator.gs

/**
 * Creates a new Google Doc from a template.
 * @param {string} newFileName The name for the new document.
 * @param {string} docLanguage The language ('german' or 'english') to select the template.
 * @returns {GoogleAppsScript.Drive.File} The newly created file object.
 */
function createDocument(newFileName, docLanguage) {
    const sourceFile = "DocGenerator_gs";
    ExecutionTimer.start('createDocument_total');
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'createDocument_start' });
    Log[sourceFile](`[${sourceFile} - createDocument] Start. Creating doc named '${newFileName}'.`);

    const templateId = docLanguage === "german" ? CONFIG.templates.german : CONFIG.templates.english;
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(activeSpreadsheet.getId());
    const destinationFolder = ssFile.getParents().hasNext() ? ssFile.getParents().next() : DriveApp.getRootFolder();
    
    ExecutionTimer.start('createDocument_deleteExisting');
    const existingFiles = destinationFolder.getFilesByName(newFileName);
    if (existingFiles.hasNext()) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'createDocument_existingFound' });
      existingFiles.next().setTrashed(true);
      Log[sourceFile](`[${sourceFile} - createDocument] Info: Trashed existing file with the same name.`);
    }
    ExecutionTimer.end('createDocument_deleteExisting');

    ExecutionTimer.start('createDocument_makeCopy');
    const copiedFile = DriveApp.getFileById(templateId).makeCopy(newFileName, destinationFolder);
    ExecutionTimer.end('createDocument_makeCopy');
    
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'createDocument_end' });
    Log[sourceFile](`[${sourceFile} - createDocument] End. Document created with ID: ${copiedFile.getId()}`);
    ExecutionTimer.end('createDocument_total');
    return copiedFile;
}

/**
 * Populates a Google Doc with data from the prepared data package.
 * @param {GoogleAppsScript.Drive.File} docFile The file object for the Google Doc.
 * @param {Object} dataPackage The structured data package from DocumentDataService.
 */
function populateDocContent(docFile, dataPackage) {
  const sourceFile = "DocGenerator_gs";
  ExecutionTimer.start('populateDocContent_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDocContent_start' });
  Log[sourceFile](`[${sourceFile} - populateDocContent] Start. Populating doc ID ${docFile.getId()}`);

  const doc = DocumentApp.openById(docFile.getId());
  const body = doc.getBody();
  const formData = dataPackage.formData;
  const docLanguage = dataPackage.docLanguage;
  
  ExecutionTimer.start('populateDocContent_replacePlaceholders');
  const offerTypeFromForm = formData.offerType || "binding";
  let offerValidUntilDisplay = (offerTypeFromForm === "binding") ? formatDateForLocale(formData.offerValidUntilDate, docLanguage) : ((docLanguage === "german") ? "unverbindlich" : "non-binding");
  let offerNatureDescription = (offerTypeFromForm === "binding") ? ((docLanguage === "german") ? "Ihr verbindliches Angebot" : "Your binding offer") : ((docLanguage === "german") ? "Ihr unverbindliches Angebot" : "Your non-binding offer");
  
  body.replaceText('{{OfferNatureDescription}}', offerNatureDescription);
  body.replaceText('{{CustomerCompanyName}}', dataPackage.customerCompanyName || "");
  body.replaceText('{{OfferCreatedDate}}', formatDateForLocale(formData.offerCreatedDate, docLanguage) || "");
  body.replaceText('{{OfferValidUntilDate}}', offerValidUntilDisplay || "");
  body.replaceText('{{CustomerContactName}}', formData.customerContactName || "");
  body.replaceText('{{CustomerCompanyAddress}}', formData.customerCompanyAddress || "");
  body.replaceText('{{EverphoneContactFullName}}', formData.everphoneContactFullName || "");
  body.replaceText('{{EverphoneContactPosition}}', formData.everphoneContactPosition || "");
  body.replaceText('{{EverphoneContactEmail}}', formData.everphoneContactEmail || "");
  body.replaceText('{{OverallContractTermOptions}}', formatTermForLocale(formData.overallContractTermOptions, docLanguage) || "");
  ExecutionTimer.end('populateDocContent_replacePlaceholders');

  ExecutionTimer.start('populateDocContent_specialAgreements');
  const specialAgreementText = (formData.specialAgreementText || "").trim();
  const placeholderSA = '{{SpecialAgreementText}}';
  let foundElementSA = body.findText(placeholderSA);
  if (foundElementSA) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDocContent_sa_found' });
    let el = foundElementSA.getElement(); let p = el.getParent();
    while (p && p.getType() !== DocumentApp.ElementType.PARAGRAPH) { el = p; p = el.getParent(); }
    const ap = (p && p.getType() === DocumentApp.ElementType.PARAGRAPH) ? p : null;
    if (ap) {
      if (specialAgreementText !== "") {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDocContent_sa_hasText' });
        let h = (docLanguage === "german") ? "Sondervereinbarungen:" : "Special Agreements:";
        ap.clear(); let hE = ap.appendText(h); hE.setBold(true);
        let pC = ap.getParent(); let i = pC.getChildIndex(ap);
        pC.insertParagraph(i + 1, "");
        let cP = pC.insertParagraph(i + 2, specialAgreementText); cP.setBold(false);
      } else {
        Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDocContent_sa_noText' });
        ap.removeFromParent();
      }
    } else {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDocContent_sa_noParagraph' });
      body.replaceText(placeholderSA, "");
    }
  }
  ExecutionTimer.end('populateDocContent_specialAgreements');

  populateDeviceTable(doc, dataPackage.devicesData, dataPackage.grandTotal, docLanguage);
  
  doc.saveAndClose();
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDocContent_end' });
  Log[sourceFile](`[${sourceFile} - populateDocContent] End. Content populated and doc saved.`);
  ExecutionTimer.end('populateDocContent_total');
}

function populateDeviceTable(doc, devicesData, grandTotal, language) {
  const sourceFile = "DocGenerator_gs";
  ExecutionTimer.start('populateDeviceTable_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDeviceTable_start' });

  const body = doc.getBody();
  const tables = body.getTables();
  if (tables.length === 0) { 
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDeviceTable_noTable' });
    ExecutionTimer.end('populateDeviceTable_total');
    return; 
  }
  const table = tables[0];
  const cellBorderStyle = {
    [DocumentApp.Attribute.BORDER_COLOR]: '#000000', [DocumentApp.Attribute.BORDER_WIDTH]: 1,
    [DocumentApp.Attribute.PADDING_LEFT]: 5, [DocumentApp.Attribute.PADDING_RIGHT]: 5,
    [DocumentApp.Attribute.PADDING_TOP]: 2, [DocumentApp.Attribute.PADDING_BOTTOM]: 2
  };

  const templateRowIndexInDoc = 1;
  if (table.getNumRows() <= templateRowIndexInDoc) { 
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDeviceTable_noTemplateRow' });
    ExecutionTimer.end('populateDeviceTable_total');
    return; 
  }
  const styledTemplateRow = table.getRow(templateRowIndexInDoc);

  ExecutionTimer.start('populateDeviceTable_loop');
  if (devicesData && devicesData.length > 0) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDeviceTable_hasDevices' });
    devicesData.forEach((device, i) => {
      const newRow = table.insertTableRow(templateRowIndexInDoc + i);
      newRow.appendTableCell(styledTemplateRow.getCell(0).copy()).clear().setText(device.model || '');
      newRow.appendTableCell(styledTemplateRow.getCell(1).copy()).clear().setText(String(device.quantity || ''));
      newRow.appendTableCell(styledTemplateRow.getCell(2).copy()).clear().setText(formatTermForLocale(device.term, language));
      newRow.appendTableCell(styledTemplateRow.getCell(3).copy()).clear().setText(formatNumberForLocale(device.netMonthlyRentalPrice, language, true));
      newRow.appendTableCell(styledTemplateRow.getCell(4).copy()).clear().setText(formatNumberForLocale(device.totalNetMonthlyRentalPrice, language, true));
    });
    table.removeRow(templateRowIndexInDoc + devicesData.length);
  } else {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDeviceTable_noDevices' });
    for (let c = 0; c < styledTemplateRow.getNumCells(); c++) {
      styledTemplateRow.getCell(c).clear().setText('-');
    }
  }
  ExecutionTimer.end('populateDeviceTable_loop');

  const totalRow = table.getRow(table.getNumRows() - 1);
  if (totalRow) {
    const boldWithBorderStyle = { ...cellBorderStyle, [DocumentApp.Attribute.BOLD]: true };
    const grandTotalStr = formatNumberForLocale(grandTotal, language, true);
    if (totalRow.getNumCells() >= 5) {
      totalRow.getCell(3).clear().setText("Total:").setAttributes(boldWithBorderStyle);
      totalRow.getCell(4).clear().setText(grandTotalStr).setAttributes(boldWithBorderStyle);
    }
  }
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'populateDeviceTable_end' });
  ExecutionTimer.end('populateDeviceTable_total');
}

function showSuccessDialog(docUrl, message) {
  const sourceFile = "DocGenerator_gs";
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'showSuccessDialog_start' });
  const htmlTemplate = HtmlService.createTemplateFromFile('HTML/SuccessDialog');
  htmlTemplate.docUrl = docUrl;
  htmlTemplate.message = message || "Document processed successfully!";
  const htmlOutput = htmlTemplate.evaluate().setWidth(600).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Process Complete');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'showSuccessDialog_end' });
}