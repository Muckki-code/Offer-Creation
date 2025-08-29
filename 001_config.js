// In Config.gs

/**
 * Global configuration object for the Offer Generation and Approval Application.
 */
var CONFIG = { // Changed to 'var' to allow re-assignment in tests
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
    highlightBundlesWithBorders: true
  },


  templates: {
    english: '15JhPytfMOr1PxUldt5w8FDzZ7vi2c3slWBD3S00_EAU',
    german: '1zBZippAYWw5KOhljo5TIQhheNCdN6_uPWABuQMe7JU0'
  },

  maxDataColumn: 22, // UPDATED: Last column is now V (22)

  offerDetailsCells: {
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
    telekomDeal: "L1"
    // Obsolete template cells removed
  },

  documentDeviceData: {
    startRow: 7, // UPDATED: Data now starts on row 7
    columns: {
      sku: 'A',
      epCapexRaw: 'B',
      tkCapexRaw: 'C',
      rentalTargetRaw: 'D',
      rentalLimitRaw: 'E',
      index: 'F',
      bundleNumber: 'G',
      model: 'H'
    },
    columnIndices: {}
  },

  bqQuerySettings: {
    scriptStartRow: 7, // UPDATED: BQ script also starts on row 7
    skuColumnLetter: 'A',
    outputStartColumnLetter: 'B',
    numOutputColumns: 4,
    projectId: '208090676765',
    tableName: 'everphone-bi-testing.superset.dt_prices_overview'
  },

  approvalWorkflow: {
    startDataRow: 7, // UPDATED: Approval workflow also starts on row 7
    approverEmail: "alexander.muegge@everphone.de",
    columns: {
      // UPDATED all columns, removing AccCapex and Checkbox, shifting everything left.
      aeEpCapex: 'I',
      aeTkCapex: 'J',
      // K is now available
      aeSalesAskPrice: 'K',
      aeQuantity: 'L',
      aeTerm: 'M',
      approverAction: 'N', // Was O
      approverComments: 'O', // Was Q
      approverPriceProposal: 'P', // Was R
      lrfPreview: 'Q', // Was S
      contractValuePreview: 'R', // Was T
      status: 'S', // Was U
      financeApprovedPrice: 'T', // Was V
      approvedBy: 'U', // Was W
      approvalDate: 'V'  // Was X
    },
    columnIndices: {},
    statusStrings: {
      draft: "Draft",
      pending: "Pending Approval",
      approvedOriginal: "Approved (Original Price)",
      approvedNew: "Approved (New Price)",   
      revisedByAE: "Revised by AE",
      rejected: "Rejected"
    },
    friendlyColumnNames: {
      model: "Model",
      bundleNumber: "Bundle Number",
      aeSalesAskPrice: "Sales Rental Price",
      quantity: "Quantity",
      aeTerm: "Term",
      aeEpCapex: "EP Capex",
      aeTkCapex: "TK Capex"
      // aeAccCapex removed
    },
    approverActionColors: {
      // UPDATED to new actions
      "Choose Action": { font: "#6c757d" }, // Neutral grey for default
      "Approve Original Price": { font: "#1E8449" }, // Green
      "Approve New Price": { font: "#1E8449" },      // Green
      "Reject with Comment": { font: "#C0392B" }       // Red
    }
  },

  protectedColumns: [
    // Device Trader data (from BQ)
    'B', 'C', 'D', 'E',
    // Row Index
    'F',
    // UPDATED all protected columns, removing old ones and shifting the rest left
    'Q', 'R', 'S', 'T', 'U', 'V'
    // Note: Approver Action ('N') and Comments ('O') are intentionally not protected.
  ],

   // MODIFIED: Simplified to a single, universal format
  numberFormats: {
    currency: '#,##0.00" €"',
    number: "#,##0",
    percentage: "0.00%",
    date: "yyyy-mm-dd" // Using a neutral date format for the sheet
  },

  conditionalFormatColors: {
    approved: {
      background: '#D9EAD3' // Light Green
    },
    pending: {
      background: '#FFF2CC' // Light Yellow
    },
    rejected: {
      background: '#F4CCCC' // Light Red
    },
    draft: {
      background: '#D9D9D9' // Light Grey
    }
  },


  logging: {
    globalDisable: false,
    sheetLogDisable: true,
    file: {
      config_gs: false,
      Logger_gs: false,
      Main_gs: false,
      SheetCoreAutomations_gs: true,
      BqDtQuery_gs: false,
      SheetStatusLogic_gs: true,
      AE_Actions_gs: false,
      ApprovalWorkflow_gs: false,
      DocGenerator_gs: true,
      UxControl_gs: true,
      BundleService_gs: false,
      DocumentDataService_gs:true,
      Formatters_gs: true,
      TestUtilities_gs: false,
      MetadataService_gs: true,
      SheetCorrectionService_gs: true,
      UiService_gs: true,


      TestResults_gs: true,
      TestDebug_gs: false,
      TestCoverage_gs: false,
      ExecutionTime_gs: false
    }
  },

  logSheets: {
    logSpreadsheetName: 'Offer Tool Logs',
    logSpreadsheetId: '',
    tableLogs: {
      sheetName: "TableLogs",
      columns: {
        timestamp: "Timestamp", name: "Name", index: "Index", model: "Model",
        bq_info_sku: "BQ_SKU", bq_info_prices: "BQ_Prices", status: "Status", aeEpCapex: "AE EP Capex",
        aeTkCapex: "AE TK Capex", aeSalesAskPrice: "AE Sales Ask Price", // aeAccCapex removed
        quantity: "Quantity", term: "Term", approverAction: "Approver Action", // processCheckbox removed
        approverComments: "Approver Comments", approverPriceProposal: "Approver Price Proposal",
        lrfPreview: "LRF Preview", contractValuePreview: "Contract Value Preview",
        financeApprovedPrice: "Finance Approved Price", approvedBy: "Approved By", approvalDate: "Approval Date"
      }
    },
    documentLogs: {
      sheetName: "DocumentLogs",
      columns: {
        timestamp: "Timestamp", user: "User", action: "Action", docName: "Document Name", docUrl: "Document URL",
        customerCompany: "Customer Company", offerType: "Offer Type", language: "Language", details: "Details"
      }
    },
    communicationLogs: {
      sheetName: "CommunicationLogs",
      columns: {
        timestamp: "Timestamp", sender: "Sender", recipient: "Recipient", subject: "Subject", type: "Type", details: "Details"
      }
    },
    generalLogs: {
      sheetName: "GeneralLogs",
      columns: {
        timestamp: "Timestamp", user: "User", action: "Action", details: "Details"
      }
    }
  }
};

/**
 * Global helper function to convert a column letter to its 1-based numerical index.
 */
function getColumnIndexByLetter(letter) {
  if (!letter || typeof letter !== 'string' || letter.length === 0) {
    return -1;
  }
  let column = 0, length = letter.length;
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
  Log.config_gs(`[config.gs - getLastLastRow] Start: sheet='${sheet.getName()}'`);
  try {
    const lastRow = sheet.getDataRange().getLastRow();
    Log.config_gs(`[config.gs - getLastLastRow] Result: DataRange last row found: ${lastRow}.`);
    return lastRow;
  } catch (e) {
    Log.config_gs(`[config.gs - getLastLastRow] ERROR: Error getting data range or last row: ${e.message}. Falling back to iterative check.`);
    const maxRows = sheet.getMaxRows();
    for (let r = maxRows; r >= 1; r--) {
      const range = sheet.getRange(r, 1, 1, sheet.getLastColumn());
      if (range.getDisplayValues()[0].some(cell => cell !== "")) {
        Log.config_gs(`[config.gs - getLastLastRow] Fallback Result: Content found in row ${r}. Returning ${r}.`);
        return r;
      }
    }
    Log.config_gs(`[config.gs - getLastLastRow] Fallback Result: No content found. Returning 0.`);
    return 0;
  } finally {
    Log.config_gs("[config.gs - getLastLastRow] End.");
  }
}

/**
 * Helper function to find the last row containing content in a specific column.
 */
function getLastPopulatedRowInColumn(sheet, columnIndex) {
  Log.config_gs(`[config.gs - getLastPopulatedRowInColumn] Start: sheet='${sheet.getName()}', columnIndex=${columnIndex}`);
  if (columnIndex < 1 || columnIndex > sheet.getMaxColumns()) {
    Log.config_gs(`[config.gs - getLastPopulatedRowInColumn] Condition: Invalid columnIndex ${columnIndex}. Returning 0.`);
    return 0;
  }
  const maxRows = sheet.getMaxRows();
  Log.config_gs(`[config.gs - getLastPopulatedRowInColumn] Info: Max rows in sheet: ${maxRows}`);
  const columnValues = sheet.getRange(1, columnIndex, maxRows, 1).getValues();
  Log.config_gs(`[config.gs - getLastPopulatedRowInColumn] Info: Fetched values for column ${columnIndex}.`);

  for (let r = maxRows - 1; r >= 0; r--) {
    if (columnValues[r] && columnValues[r][0] !== null && String(columnValues[r][0]).trim() !== "") {
      Log.config_gs(`[config.gs - getLastPopulatedRowInColumn] Condition: Content found in row ${r + 1}. Returning ${r + 1}.`);
      return r + 1;
    }
  }
  Log.config_gs(`[config.gs - getLastPopulatedRowInColumn] Condition: No content found in column ${columnIndex}. Returning 0.`);
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
    if (typeof Log !== 'undefined' && Log.config_gs) {
      Log.config_gs(message);
    }
  };

  try {
    log('[config.gs - Initializer] Pre-calculating column indices...');

    if (config.documentDeviceData && config.documentDeviceData.columns) {
      config.documentDeviceData.columnIndices = calculateIndices(config.documentDeviceData.columns);
    }

    if (config.approvalWorkflow && config.approvalWorkflow.columns) {
      config.approvalWorkflow.columnIndices = calculateIndices(config.approvalWorkflow.columns);
    }

    if (config.bqQuerySettings) {
      config.bqQuerySettings.skuColumnIndex = getColumnIndexByLetter(config.bqQuerySettings.skuColumnLetter);
      config.bqQuerySettings.outputStartColumnIndex = getColumnIndexByLetter(config.bqQuerySettings.outputStartColumnLetter);
    }

    // NEW: Pre-calculate indices for protected columns for fast lookups
    if (config.protectedColumns) {
      config.protectedColumnIndices = config.protectedColumns.map(letter => getColumnIndexByLetter(letter));
    }

    log('[config.gs - Initializer] Column indices pre-calculation complete.');

  } catch (e) {
    log(`[config.gs - Initializer] CRITICAL ERROR during config initialization: ${e.message}`);
  }

})(CONFIG);