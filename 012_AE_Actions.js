/**
 * @file This file contains all functions for actions taken by the Account Executive (AE).
 */

/**
 * Finds all items that are pending approval ('Pending Approval' or 'Revised by AE'),
 * and sends a summary email notification to the approver.
 * This function DOES NOT change the status of the rows.
 */
function submitItemsForApproval() {
  const sourceFile = "AE_Actions_gs";
  ExecutionTimer.start('submitItemsForApproval_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'submitItemsForApproval_start' });
  Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Start: submitItemsForApproval function started.`);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Info: Active sheet: '${sheet.getName()}'`);
  const ui = SpreadsheetApp.getUi();

  const config = CONFIG.approvalWorkflow;
  const startRow = config.startDataRow;
  
  const lastRow = getLastLastRow(sheet);
  Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Info: Data Start Row: ${startRow}, Last Row: ${lastRow}.`);

  if (lastRow < startRow) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'submitItemsForApproval_noDataRows' });
    ui.alert("No data rows to process.");
    Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Condition: No data rows found to process. Exiting.`);
    logGeneralActivity({
      action: "Submit Items for Approval Skipped",
      details: "No data rows found to process.",
      sheetName: sheet.getName()
    });
    ExecutionTimer.end('submitItemsForApproval_total');
    return;
  }

  const startCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - startCol + 1;
  const statusColIndex = config.columnIndices.status;
  Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Reading from start column ${startCol}, number of columns ${numCols}. Status is in column ${statusColIndex}.`);

  ExecutionTimer.start('submitItemsForApproval_readSheet');
  const dataRange = sheet.getRange(startRow, startCol, lastRow - startRow + 1, numCols);
  const allValues = dataRange.getValues();
  ExecutionTimer.end('submitItemsForApproval_readSheet');
  Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Info: Read ${allValues.length} rows from range ${dataRange.getA1Notation()} for status check.`);

  let itemsToNotify = 0;
  const statusesToLookFor = [
    config.statusStrings.pending,
    config.statusStrings.revisedByAE
  ];
  Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Info: Statuses to look for: ${JSON.stringify(statusesToLookFor)}.`);

  ExecutionTimer.start('submitItemsForApproval_countItems');
  for (const row of allValues) {
    const currentStatus = row[statusColIndex - startCol];
    if (statusesToLookFor.includes(currentStatus)) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'submitItemsForApproval_itemFound' });
      itemsToNotify++;
    }
  }
  ExecutionTimer.end('submitItemsForApproval_countItems');
  Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Found ${itemsToNotify} items to notify about.`);

  if (itemsToNotify > 0) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'submitItemsForApproval_showDialogPath' });
    Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Condition: ${itemsToNotify} items found to notify. Proceeding with dialog.`);
    
    ExecutionTimer.start('submitItemsForApproval_prepareDialog');
    const currentLanguage = (sheet.getRange(CONFIG.offerDetailsCells.language).getValue() || "german").toString().trim().toLowerCase();
    const initialAeName = sheet.getRange(CONFIG.offerDetailsCells.yourName).getValue() || "";
    const initialCustomerCompany = sheet.getRange(CONFIG.offerDetailsCells.customerCompany).getValue() || "";
    const initialTelekomDeal = (sheet.getRange(CONFIG.offerDetailsCells.telekomDeal).getValue() || "").toString().trim();
    
    const initialDataForDialog = {
        language: currentLanguage,
        aeName: initialAeName,
        customerCompany: initialCustomerCompany,
        telekomDealInitial: initialTelekomDeal
    };
    Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Info: Data passed to PersonalMessageDialog: ${JSON.stringify(initialDataForDialog)}.`);
    ExecutionTimer.end('submitItemsForApproval_prepareDialog');

    ExecutionTimer.start('submitItemsForApproval_renderDialog');
    const htmlTemplate = HtmlService.createTemplateFromFile('PersonalMessageDialog');
    htmlTemplate.initialData = initialDataForDialog;
    const htmlOutput = htmlTemplate.evaluate().setWidth(550).setHeight(400);

    const dialogTitle = (currentLanguage === 'german' ? "Nachricht an Genehmiger" : "Message to Approver");
    ui.showModalDialog(htmlOutput, dialogTitle);
    ExecutionTimer.end('submitItemsForApproval_renderDialog');
    Log[sourceFile](`[${sourceFile} - submitItemsForApproval] Info: Personal Message Dialog displayed with title: '${dialogTitle}'.`);

  } else {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'submitItemsForApproval_noItemsPath' });
    ui.alert("No items with status 'Pending Approval' or 'Revised by AE' were found.");
    Log[sourceFile]("[${sourceFile} - submitItemsForApproval] Condition: No items found to notify. Displayed alert.");
    logGeneralActivity({
      action: "Submit Items for Approval Skipped",
      details: "No items with relevant statuses found to notify.",
      sheetName: sheet.getName()
    });
  }
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'submitItemsForApproval_end' });
  Log[sourceFile](`[${sourceFile} - submitItemsForApproval] End: submitItemsForApproval function finished.`);
  ExecutionTimer.end('submitItemsForApproval_total');
}

/**
 * This function is called from the client-side PersonalMessageDialog.html
 * after the user enters a message or cancels.
 */
function _handlePersonalMessageSubmission(personalMessage, aeName, customerCompany, telekomDeal) {
  const sourceFile = "AE_Actions_gs";
  ExecutionTimer.start('handlePersonalMessageSubmission_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'handlePersonalMessageSubmission_start' });
  Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] Start.`);
  Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] Info: Received - personalMessage: '${personalMessage}', aeName: '${aeName}', customerCompany: '${customerCompany}', telekomDeal: '${telekomDeal}'.`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (personalMessage === null && aeName === null && customerCompany === null && telekomDeal === null) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'handlePersonalMessageSubmission_userCancelled' });
      Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] Condition: Dialog cancelled by user. Exiting.`);
      logGeneralActivity({ action: "Personal Message Dialog Cancelled", details: "User cancelled.", sheetName: sheet.getName() });
      ExecutionTimer.end('handlePersonalMessageSubmission_total');
      return;
  }
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'handlePersonalMessageSubmission_userSubmitted' });

  ExecutionTimer.start('handlePersonalMessageSubmission_updateSheet');
  if (aeName !== null && aeName.trim() !== '') {
    sheet.getRange(CONFIG.offerDetailsCells.yourName).setValue(aeName.trim());
    Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] Info: Updated AE Name in sheet to: '${aeName.trim()}'.`);
  }
  if (customerCompany !== null && customerCompany.trim() !== '') {
    sheet.getRange(CONFIG.offerDetailsCells.customerCompany).setValue(customerCompany.trim());
    Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] Info: Updated Customer Company in sheet to: '${customerCompany.trim()}'.`);
  }
  if (telekomDeal !== null && telekomDeal.trim() !== '') {
    sheet.getRange(CONFIG.offerDetailsCells.telekomDeal).setValue(telekomDeal.trim());
    Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] Info: Updated Telekom Deal in sheet to: '${telekomDeal.trim()}'.`);
  }
  ExecutionTimer.end('handlePersonalMessageSubmission_updateSheet');
  
  const config = CONFIG.approvalWorkflow;
  const startRow = config.startDataRow;
  const lastRow = getLastLastRow(sheet);
  
  const startCol = CONFIG.documentDeviceData.columnIndices.sku;
  const numCols = CONFIG.maxDataColumn - startCol + 1;
  const statusColIndex = config.columnIndices.status;
  
  ExecutionTimer.start('handlePersonalMessageSubmission_recountItems');
  const dataRange = sheet.getRange(startRow, startCol, lastRow - startRow + 1, numCols);
  const allValues = dataRange.getValues();
  let itemsToNotify = 0;
  const statusesToLookFor = [ config.statusStrings.pending, config.statusStrings.revisedByAE ];
  for (const row of allValues) {
    const currentStatus = row[statusColIndex - startCol];
    if (statusesToLookFor.includes(currentStatus)) {
      itemsToNotify++;
    }
  }
  ExecutionTimer.end('handlePersonalMessageSubmission_recountItems');
  Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] Info: Recalculated itemsToNotify: ${itemsToNotify}.`);

  ExecutionTimer.start('handlePersonalMessageSubmission_prepareEmail');
  const currentCustomerCompany = sheet.getRange(CONFIG.offerDetailsCells.customerCompany).getValue();
  const offerType = sheet.getRange(CONFIG.offerDetailsCells.offerType).getValue();
  const currentSubmitterName = sheet.getRange(CONFIG.offerDetailsCells.yourName).getValue();
  const currentTelekomDeal = sheet.getRange(CONFIG.offerDetailsCells.telekomDeal).getValue();

  const summaryData = {
      itemCount: itemsToNotify,
      customerCompany: currentCustomerCompany,
      offerType: offerType,
      submitterName: currentSubmitterName,
      personalMessage: personalMessage,
      telekomDealStatus: currentTelekomDeal
  };
  ExecutionTimer.end('handlePersonalMessageSubmission_prepareEmail');
  Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] Info: Summary data for email: ${JSON.stringify(summaryData)}.`);

  sendApprovalRequestEmail(summaryData);
  Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] Info: sendApprovalRequestEmail function called.`);

  ui.alert(`Notification Sent: The approver has been notified about ${itemsToNotify} item(s) that are pending approval.`);
  logGeneralActivity({
    action: "Approval Request Submitted",
    details: `${itemsToNotify} items submitted. Customer: ${summaryData.customerCompany}`,
    sheetName: sheet.getName()
  });

  Log.TestCoverage_gs({ file: sourceFile, coverage: 'handlePersonalMessageSubmission_end' });
  Log[sourceFile](`[${sourceFile} - _handlePersonalMessageSubmission] End.`);
  ExecutionTimer.end('handlePersonalMessageSubmission_total');
}


/**
 * Helper function to send a summary email notification to the approver.
 */
function sendApprovalRequestEmail(summaryData) {
  const sourceFile = "AE_Actions_gs";
  ExecutionTimer.start('sendApprovalRequestEmail_total');
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'sendApprovalRequestEmail_start' });
  Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] Start.`);
  Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] Info: Summary data received: ${JSON.stringify(summaryData)}.`);

  if (!summaryData || !summaryData.itemCount || summaryData.itemCount === 0) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'sendApprovalRequestEmail_noItems' });
    Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] Condition: No summary data or item count is zero. Exiting.`);
    ExecutionTimer.end('sendApprovalRequestEmail_total');
    return;
  }
  
  const approverEmail = CONFIG.approvalWorkflow.approverEmail;
  if (!approverEmail) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'sendApprovalRequestEmail_noApproverEmail' });
    Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] Condition: No approver email is set in CONFIG. Skipping email.`);
    logGeneralActivity({ action: "Email Send Failed", details: "No approver email configured.", sheetName: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() });
    ExecutionTimer.end('sendApprovalRequestEmail_total');
    return;
  }
  Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] Info: Approver Email from CONFIG: '${approverEmail}'.`);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const subject = `Price Approval Request: ${summaryData.customerCompany} (${summaryData.itemCount} items)`;
  Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] Info: Email Subject: '${subject}'.`);

  ExecutionTimer.start('sendApprovalRequestEmail_renderTemplate');
  const htmlTemplate = HtmlService.createTemplateFromFile('ApprovalEmail');
  htmlTemplate.summary = summaryData;
  htmlTemplate.spreadsheetName = spreadsheet.getName();
  htmlTemplate.sheetName = spreadsheet.getActiveSheet().getName();
  htmlTemplate.spreadsheetUrl = spreadsheet.getUrl();
  
  const approverFirstName = approverEmail.split('@')[0].split('.')[0];
  htmlTemplate.approverFirstName = approverFirstName.charAt(0).toUpperCase() + approverFirstName.slice(1);
  const htmlBody = htmlTemplate.evaluate().getContent();
  ExecutionTimer.end('sendApprovalRequestEmail_renderTemplate');
  Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] Info: HTML email body generated.`);

  try {
    ExecutionTimer.start('sendApprovalRequestEmail_mailApp');
    MailApp.sendEmail({ to: approverEmail, subject: subject, htmlBody: htmlBody });
    ExecutionTimer.end('sendApprovalRequestEmail_mailApp');
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'sendApprovalRequestEmail_success' });
    Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] Info: Email sent successfully.`);
    logCommunicationActivity({
      recipient: approverEmail,
      subject: subject,
      type: "Approval Request Email",
      details: `Success for ${summaryData.itemCount} items.`
    });
  } catch (e) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'sendApprovalRequestEmail_failure' });
    Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] ERROR: Failed to send email. Error: ${e.message}. Stack: ${e.stack}`);
    logCommunicationActivity({
      recipient: approverEmail,
      subject: subject,
      type: "Approval Request Email",
      details: `Failed to send email. Error: ${e.message}`,
      status: "Failed"
    });
  }
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'sendApprovalRequestEmail_end' });
  Log[sourceFile](`[${sourceFile} - sendApprovalRequestEmail] End.`);
  ExecutionTimer.end('sendApprovalRequestEmail_total');
}