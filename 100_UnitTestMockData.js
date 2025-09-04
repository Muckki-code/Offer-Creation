/**
 * @file This file contains mock data used for unit tests.
 * Unit tests are designed to run in isolation without interacting with a live sheet.
 * This data is structured to be easily consumed by the test functions.
 */

const MOCK_DATA_UNIT = {
  /**
   * A simplified set of column indexes for testing SheetCoreAutomations functions like updateCalculationsForRow.
   * These are 1-based indexes.
   * UPDATED: Reflects the new 21-column structure with a single `aeCapex` column.
   */
  colIndexes: {
    status: 18,
    financeApprovedPrice: 19,
    approverPriceProposal: 15,
    aeSalesAskPrice: 10,
    aeCapex: 9, // REPLACES aeEpCapex and aeTkCapex
    aeTerm: 12,
    lrfPreview: 16,
    contractValuePreview: 17,
    aeQuantity: 11,
  },

  /**
   * A comprehensive set of column indexes for testing the full ApprovalWorkflow.
   * These are 1-based indexes, matching the new CONFIG file after removals.
   * UPDATED: Reflects the new 21-column structure with a single `aeCapex` column.
   */
  approvalWorkflowTestColIndexes: {
    sku: 1,
    epCapex: 2,
    ep24PriceTarget: 3,
    ep36PriceTarget: 4,
    tkCapex: 5,
    tk24PriceTarget: 6,
    tk36PriceTarget: 7,
    index: 8,
    bundleNumber: 9,
    model: 10,
    aeCapex: 11,
    aeSalesAskPrice: 12,
    aeQuantity: 13,
    aeTerm: 14,
    approverAction: 15,
    approverComments: 16,
    approverPriceProposal: 17,
    lrfPreview: 18,
    contractValuePreview: 19,
    status: 20,
    financeApprovedPrice: 21,
    approvedBy: 22,
    approvalDate: 23,
  },

  /**
   * Mock row data for SheetCoreAutomations unit tests.
   * The `get` method ensures a deep copy is returned for each test.
   * UPDATED: Now uses the 21-column structure with a single aeCapex value.
   */
  rows: {
    _data: {
      standard: [
        "SKU-01", 1000, 50, 40, 900, 45, 35, 1, "", "Test Device", // BQ + identifiers
        1000, 100, 10, 12, // AE Inputs
        "Choose Action", "", "", "", "", // Approver + Calcs
        "Pending Approval", "", "", ""  // Status + Finalization
      ],
      withAccessories: [
        "SKU-02", 1100, 55, 45, 950, 50, 40, 2, "", "Test Device with Acc.", // BQ + identifiers
        1100, 110, 10, 12, // AE Inputs
        "Choose Action", "", "", "", "", // Approver + Calcs
        "Pending Approval", "", "", ""  // Status + Finalization
      ],
      approverPrice: [
        "SKU-03", 1000, 50, 40, 900, 45, 35, 3, "", "Test Device Approver", // BQ + identifiers
        1000, 100, 10, 12, // AE Inputs
        "Choose Action", "", 96, "", "", // Approver + Calcs
        "Pending Approval", "", "", ""  // Status + Finalization
      ],
      approved: [
        "SKU-04", 1000, 50, 40, 900, 45, 35, 4, "", "Test Device Approved", // BQ + identifiers
        1000, 100, 10, 12, // AE Inputs
        "Choose Action", "", "", "", "", // Approver + Calcs
        "Approved (New Price)", 90, "approver@test.com", "2023-01-01"  // Status + Finalization
      ],
    },
    get: function(name) {
      return JSON.parse(JSON.stringify(this._data[name]));
    },
  },

  /**
   * Mock row data specifically for ApprovalWorkflow unit tests.
   * The `get` method ensures a deep copy is returned for each test.
   * UPDATED: Now uses the 21-column structure with a single aeCapex value.
   */
  rowsForApprovalTests: {
    _data: {
      approvedOriginalProcessed: [
        "SKU-A1", 1000, 50, 40, 900, 45, 35, 1, "", "Test Device",
        1000, 100, 10, 24,
        "Approve Original Price", "", "", "", "",
        "Approved (Original Price)", 100, "old.approver@example.com", "2025-01-01T10:00:00.000Z"
      ],
      pendingApprovedOriginal: [
        "SKU-A2", 1000, 50, 40, 900, 45, 35, 2, "", "Test Device",
        1000, 100, 10, 24,
        "Approve Original Price", "", "", "1.2", "28800",
        "Pending Approval", "", "", ""
      ],
      pendingApprovedNew: [
        "SKU-A3", 1000, 50, 40, 900, 45, 35, 3, "", "Test Device",
        1000, 100, 10, 24,
        "Approve New Price", "", 90, "1.08", "25920",
        "Pending Approval", "", "", ""
      ],
      pendingApproveNoPrice: [
        "SKU-A4", 1000, 50, 40, 900, 45, 35, 4, "", "Test Device",
        1000, "", 10, 24,
        "Approve Original Price", "", "", "", "",
        "Pending Approval", "", "", ""
      ],
      pendingRejected: [
        "SKU-A5", 1000, 50, 40, 900, 45, 35, 5, "", "Test Device",
        1000, 100, 10, 24,
        "Reject with Comment", "Comment available", "", "1.2", "28800",
        "Pending Approval", "", "", ""
      ],
      pendingRejectNoComment: [
        "SKU-A6", 1000, 50, 40, 900, 45, 35, 6, "", "Test Device",
        1000, 100, 10, 24,
        "Reject with Comment", "", "", "", "",
        "Pending Approval", "", "", ""
      ],
      pendingRevision: [
        "SKU-A7", 1000, 50, 40, 900, 45, 35, 7, "", "Test Device",
        1000, 100, 10, 24,
        "Choose Action", "Comment available", "", "", "",
        "Pending Approval", "", "", ""
      ],
      pendingRevisionNoComment: [
        "SKU-A8", 1000, 50, 40, 900, 45, 35, 8, "", "Test Device",
        1000, 100, 10, 24,
        "Choose Action", "", "", "", "",
        "Pending Approval", "", "", ""
      ],
      pendingInvalidStatus: [
        "SKU-A9", 1000, 50, 40, 900, 45, 35, 9, "", "Test Device",
        1000, 100, 10, 24,
        "Approve Original Price", "", "", "", "",
        "Draft", "", "", ""
      ],
    },
    get: function(name) {
      return JSON.parse(JSON.stringify(this._data[name]));
    },
  },

  /**
   * Mock row data specifically for SheetStatusLogic unit tests.
   * The `get` method ensures a deep copy is returned for each test.
   * UPDATED: Now uses the 21-column structure with a single aeCapex value.
   */
  rowsForStatusLogicTests: {
    _data: {
      blank: [
        "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", "", "", "",
        "", "", "", ""
      ],
      draftIncomplete: [
        "SKU-DRAFT", 1000, 50, 40, 900, 45, 35, "1", "", "Test Device Draft",
        "", "", "10", "",
        "Choose Action", "", "", "", "",
        "Draft", "", "", ""
      ],
      pending: [
        "SKU-PENDING", 1000, 50, 40, 900, 45, 35, "4", "", "Test Device Pending",
        "1000", "50", "10", "24",
        "Choose Action", "", "", "", "",
        "Pending Approval", "", "", ""
      ],
      approved: [
        "SKU-APPROVED", 1000, 50, 40, 900, 45, 35, "2", "", "Test Device Approved",
        "1000", "100", "10", "24",
        "Approve New Price", "", "95", "1.14", "27360",
        "Approved (New Price)", "95", "approver@test.com", "2025-07-15"
      ],
      revisedByAE: [
        "SKU-REVISED", 1000, 50, 40, 900, 45, 35, "5", "", "Test Device Revised",
        "1500", "120", "12", "36",
        "Choose Action", "", "", "", "",
        "Revised by AE", "", "", ""
      ],
      rejected: [
        "SKU-REJECT", 1000, 50, 40, 900, 45, 35, "5", "", "Test Device Rejected",
        "800", "80", "5", "24",
        "Reject with Comment", "Price too high", "", "", "",
        "Rejected", "", "approver@test.com", "2025-07-16"
      ],
    },
    get: function(name) {
      // Ensure we return a deep copy to prevent tests from interfering with each other
      return JSON.parse(JSON.stringify(this._data[name]));
    },
  },
};
