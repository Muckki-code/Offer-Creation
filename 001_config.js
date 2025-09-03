// In Config.gs

/**
 * Global configuration object for the Offer Generation and Approval Application.
 */
var CONFIG = {
  // Changed to 'var' to allow re-assignment in tests
  // --- Document Generation ---

  featureFlags: {
    /**
     * @property {boolean} If true, runs intensive bundle integrity checks (e.g., matching term/quantity)
     * on every relevant edit. If false, these checks are skipped for performance, but data
     * integrity is not guaranteed for bundles. Default: true.
     */
    enforceBundleIntegrityOnEdit: true,
    /**
     * @property {boolean} If true, applies conditional formatting to draw borders around bundles.
     * This has minimal performance impact as it's handled by Sheets natively. Default: true.
     */
    highlightBundlesWithBorders: true,
  },

  templates: {
    english: "15JhPytfMOr1PxUldt5w8FDzZ7vi2c3slWBD3S00_EAU",
    german: "1zBZippAYWw5KOhljo5TIQhheNCdN6_uPWABuQMe7JU0",
  },

  maxDataColumn: 21, // UPDATED: Last column is now U (21)

  offerDetailsCells: {
    cachedHeaderRangeA1: "F1:O4", // <-- ADD THIS LINE
    customerCompany: "G1",
    companyAddress: "G3",
    customerContactName: "G2",
    offerValidUntil: "I4",
    specialAgreements: "G4",
    yourName: "L2",
    yourPosition: "L3",
    contractTerm: "I3",
    documentName: "L4",
    language: "I1",
    offerType: "I2",
    telekomDeal: "L1",
    approverCell: "O1", // NEW: Cell for the dynamic approver dropdown
  },

  documentDeviceData: {
    startRow: 7, // Data now starts on row 7
    columns: {
      sku: "A",
      epCapexRaw: "B",
      tkCapexRaw: "C",
      rentalTargetRaw: "D",
      rentalLimitRaw: "E",
      index: "F",
      bundleNumber: "G",
      model: "H",
    },
    columnIndices: {},
  },

  bqQuerySettings: {
    scriptStartRow: 7, // BQ script also starts on row 7
    skuColumnLetter: "A",
    outputStartColumnLetter: "B",
    numOutputColumns: 4,
    projectId: "208090676765",
    tableName: "everphone-bi-testing.superset.dt_prices_overview",
  },

  approvalWorkflow: {
    startDataRow: 7, // Approval workflow also starts on row 7
    approverList: [
      // NEW: List of approvers for the dropdown
      "alexander.muegge@everphone.de",
      "sabrina.fruehauf@everphone.de",
      "geoffrey.ochs@everphone.de"
    ],
    columns: {
      // UPDATED: Merged aeEpCapex and aeTkCapex into aeCapex and shifted all subsequent columns left.
      aeCapex: "I",
      aeSalesAskPrice: "J",
      aeQuantity: "K",
      aeTerm: "L",
      approverAction: "M",
      approverComments: "N",
      approverPriceProposal: "O",
      lrfPreview: "P",
      contractValuePreview: "Q",
      status: "R",
      financeApprovedPrice: "S",
      approvedBy: "T",
      approvalDate: "U",
    },
    columnIndices: {},
    statusStrings: {
      draft: "Draft",
      pending: "Pending Approval",
      approvedOriginal: "Approved (Original Price)",
      approvedNew: "Approved (New Price)",
      rejected: "Rejected",
    },
    friendlyColumnNames: {
      model: "Model",
      bundleNumber: "Bundle Number",
      aeSalesAskPrice: "Sales Rental Price",
      quantity: "Quantity",
      aeTerm: "Term",
      aeCapex: "AE Capex", // UPDATED from separate EP/TK Capex
    },
    approverActionColors: {
      "Choose Action": { font: "#6c757d" },
      "Approve Original Price": { font: "#1E8449" },
      "Approve New Price": { font: "#1E8449" },
      "Reject with Comment": { font: "#C0392B" },
    },
  },

  protectedColumns: [
    // Device Trader data (from BQ)
    "B",
    "C",
    "D",
    "E",
    // Row Index
    "F",
    // UPDATED: Shifted all protected calculation/approval columns left.
    "P",
    "Q",
    "R",
    "S",
    "T",
    "U",
  ],

  numberFormats: {
    currency: '#,##0.00" €"',
    number: "#,##0",
    percentage: "0.00%",
    date: "yyyy-mm-dd",
  },

  conditionalFormatColors: {
    approved: {
      background: "#D9EAD3",
    },
    pending: {
      background: "#FFF2CC",
    },
    rejected: {
      background: "#F4CCCC",
    },
    draft: {
      background: "#D9D9D9",
    },
  },

  logging: {
    globalDisable: false,
    sheetLogDisable: true,
    file: {
      config_gs: true,
      Logger_gs: false,
      Main_gs: true,
      SheetCoreAutomations_gs: true,
      BqDtQuery_gs: false,
      SheetStatusLogic_gs: true,
      AE_Actions_gs: true,
      ApprovalWorkflow_gs: true,
      DocGenerator_gs: true,
      UxControl_gs: true,
      BundleService_gs: true,
      DocumentDataService_gs: true,
      Formatter_gs: true,
      TestUtilities_gs: true,
      MetadataService_gs: true,
      SheetCorrectionService_gs: true,
      UiService_gs: true,

      TestResults_gs: true,
      TestDebug_gs: true,
      TestCoverage_gs: false,
      ExecutionTime_gs: false,
    },
  },

  logSheets: {
    logSpreadsheetName: "Offer Tool Logs",
    logSpreadsheetId: "",
    tableLogs: {
      sheetName: "TableLogs",
      columns: {
        timestamp: "Timestamp",
        name: "Name",
        index: "Index",
        model: "Model",
        bq_info_sku: "BQ_SKU",
        bq_info_prices: "BQ_Prices",
        status: "Status",
        aeCapex: "AE Capex", // UPDATED from separate EP/TK Capex
        aeSalesAskPrice: "AE Sales Ask Price",
        quantity: "Quantity",
        term: "Term",
        approverAction: "Approver Action",
        approverComments: "Approver Comments",
        approverPriceProposal: "Approver Price Proposal",
        lrfPreview: "LRF Preview",
        contractValuePreview: "Contract Value Preview",
        financeApprovedPrice: "Finance Approved Price",
        approvedBy: "Approved By",
        approvalDate: "Approval Date",
      },
    },
    documentLogs: {
      sheetName: "DocumentLogs",
      columns: {
        timestamp: "Timestamp",
        user: "User",
        action: "Action",
        docName: "Document Name",
        docUrl: "Document URL",
        customerCompany: "Customer Company",
        offerType: "Offer Type",
        language: "Language",
        details: "Details",
      },
    },
    communicationLogs: {
      sheetName: "CommunicationLogs",
      columns: {
        timestamp: "Timestamp",
        sender: "Sender",
        recipient: "Recipient",
        subject: "Subject",
        type: "Type",
        details: "Details",
      },
    },
    generalLogs: {
      sheetName: "GeneralLogs",
      columns: {
        timestamp: "Timestamp",
        user: "User",
        action: "Action",
        details: "Details",
      },
    },
  },
};

/**
 * Global helper function to convert a column letter to its 1-based numerical index.
 */
function getColumnIndexByLetter(letter) {
  if (!letter || typeof letter !== "string" || letter.length === 0) {
    return -1;
  }
  let column = 0,
    length = letter.length;
  for (let i = 0; i < length; i++) {
    const charVal = letter.toUpperCase().charCodeAt(i) - 64;
    if (charVal < 1 || charVal > 26) {
      return -1;
    }
    column += charVal * Math.pow(26, length - i - 1);
  }
  return column;
}

/**
 * Helper function to find the last row containing content in a sheet.
 */
function getLastLastRow(sheet) {
  Log.config_gs(
    `[config.gs - getLastLastRow] Start: sheet='${sheet.getName()}'`
  );
  try {
    const lastRow = sheet.getDataRange().getLastRow();
    Log.config_gs(
      `[config.gs - getLastLastRow] Result: DataRange last row found: ${lastRow}.`
    );
    return lastRow;
  } catch (e) {
    Log.config_gs(
      `[config.gs - getLastLastRow] ERROR: Error getting data range or last row: ${e.message}. Falling back to iterative check.`
    );
    const maxRows = sheet.getMaxRows();
    for (let r = maxRows; r >= 1; r--) {
      const range = sheet.getRange(r, 1, 1, sheet.getLastColumn());
      if (range.getDisplayValues()[0].some((cell) => cell !== "")) {
        Log.config_gs(
          `[config.gs - getLastLastRow] Fallback Result: Content found in row ${r}. Returning ${r}.`
        );
        return r;
      }
    }
    Log.config_gs(
      `[config.gs - getLastLastRow] Fallback Result: No content found. Returning 0.`
    );
    return 0;
  } finally {
    Log.config_gs("[config.gs - getLastLastRow] End.");
  }
}

/**
 * Helper function to find the last row containing content in a specific column.
 */
function getLastPopulatedRowInColumn(sheet, columnIndex) {
  Log.config_gs(
    `[config.gs - getLastPopulatedRowInColumn] Start: sheet='${sheet.getName()}', columnIndex=${columnIndex}`
  );
  if (columnIndex < 1 || columnIndex > sheet.getMaxColumns()) {
    Log.config_gs(
      `[config.gs - getLastPopulatedRowInColumn] Condition: Invalid columnIndex ${columnIndex}. Returning 0.`
    );
    return 0;
  }
  const maxRows = sheet.getMaxRows();
  Log.config_gs(
    `[config.gs - getLastPopulatedRowInColumn] Info: Max rows in sheet: ${maxRows}`
  );
  const columnValues = sheet.getRange(1, columnIndex, maxRows, 1).getValues();
  Log.config_gs(
    `[config.gs - getLastPopulatedRowInColumn] Info: Fetched values for column ${columnIndex}.`
  );

  for (let r = maxRows - 1; r >= 0; r--) {
    if (
      columnValues[r] &&
      columnValues[r][0] !== null &&
      String(columnValues[r][0]).trim() !== ""
    ) {
      Log.config_gs(
        `[config.gs - getLastPopulatedRowInColumn] Condition: Content found in row ${r +
          1}. Returning ${r + 1}.`
      );
      return r + 1;
    }
  }
  Log.config_gs(
    `[config.gs - getLastPopulatedRowInColumn] Condition: No content found in column ${columnIndex}. Returning 0.`
  );
  return 0;
}

/**
 * =================================================================
 * SPRINT 2 OPTIMIZATION: CONFIG INITIALIZER
 * =================================================================
 */
(function (config) {
  const calculateIndices = (columnObject) => {
    const indices = {};
    for (const key in columnObject) {
      if (Object.prototype.hasOwnProperty.call(columnObject, key)) {
        indices[key] = getColumnIndexByLetter(columnObject[key]);
      }
    }
    return indices;
  };

  const log = (message) => {
    // Check for existence of Log object and the specific logger function
    if (typeof Log !== 'undefined' && typeof Log.config_gs === 'function' && CONFIG.logging.file.config_gs) {
      Log.config_gs(message);
    } else {
      // Fallback to Logger.log if our custom logger isn't ready or is disabled
      Logger.log(message);
    }
  };

  try {
    log('[config.gs - Initializer] Pre-calculating column indices...');

    if (config.documentDeviceData && config.documentDeviceData.columns) {
      config.documentDeviceData.columnIndices = calculateIndices(config.documentDeviceData.columns);
      log(`[config.gs - Initializer] CRAZY VERBOSE: Calculated documentDeviceData.columnIndices: ${JSON.stringify(config.documentDeviceData.columnIndices)}`);
    }

    if (config.approvalWorkflow && config.approvalWorkflow.columns) {
      config.approvalWorkflow.columnIndices = calculateIndices(config.approvalWorkflow.columns);
      log(`[config.gs - Initializer] CRAZY VERBOSE: Calculated approvalWorkflow.columnIndices: ${JSON.stringify(config.approvalWorkflow.columnIndices)}`);
    }

    if (config.bqQuerySettings) {
      config.bqQuerySettings.skuColumnIndex = getColumnIndexByLetter(config.bqQuerySettings.skuColumnLetter);
      config.bqQuerySettings.outputStartColumnIndex = getColumnIndexByLetter(config.bqQuerySettings.outputStartColumnLetter);
      log(`[config.gs - Initializer] CRAZY VERBOSE: Calculated bqQuerySettings indices.`);
    }

    if (config.protectedColumns) {
      config.protectedColumnIndices = config.protectedColumns.map(letter => getColumnIndexByLetter(letter));
      log(`[config.gs - Initializer] CRAZY VERBOSE: Calculated protectedColumnIndices: ${JSON.stringify(config.protectedColumnIndices)}`);
    }

    log('[config.gs - Initializer] Column indices pre-calculation complete.');

  } catch (e) {
    log(`[config.gs - Initializer] CRITICAL ERROR during config initialization: ${e.message}`);
  }

})(CONFIG);
