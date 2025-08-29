/**
 * @file This file contains mock data used for unit tests.
 * Unit tests are designed to run in isolation without interacting with a live sheet.
 * This data is structured to be easily consumed by the test functions.
 */

const MOCK_DATA_UNIT = {

    /**
     * A simplified set of column indexes for testing SheetCoreAutomations functions like updateCalculationsForRow.
     * These are 1-based indexes.
     * UPDATED: Reflects the removal of AccCapex (col 11) and Checkbox (old col 16).
     */
    colIndexes: {
        status: 19,
        financeApprovedPrice: 20,
        approverPriceProposal: 16,
        aeSalesAskPrice: 11,
        aeEpCapex: 9,
        aeTkCapex: 10,
        // aeAccCapex (old 11) is removed
        aeTerm: 13,
        lrfPreview: 17,
        contractValuePreview: 18,
        aeQuantity: 12
    },

    /**
     * A comprehensive set of column indexes for testing the full ApprovalWorkflow.
     * These are 1-based indexes, matching the new CONFIG file after removals.
     */
    approvalWorkflowTestColIndexes: {
        sku: 1,
        epCapexRaw: 2,
        tkCapexRaw: 3,
        rentalTargetRaw: 4,
        rentalLimitRaw: 5,
        index: 6,
        bundleNumber: 7,
        model: 8,
        aeEpCapex: 9,
        aeTkCapex: 10,
        // aeAccCapex (old 11) is removed
        aeSalesAskPrice: 11,
        aeQuantity: 12,
        aeTerm: 13,
        approverAction: 14,
        // processActionCheckbox (old 16) is removed
        approverComments: 15,
        approverPriceProposal: 16,
        lrfPreview: 17,
        contractValuePreview: 18,
        status: 19,
        financeApprovedPrice: 20,
        approvedBy: 21,
        approvalDate: 22
    },

    /**
     * Mock row data for SheetCoreAutomations unit tests.
     * The `get` method ensures a deep copy is returned for each test.
     * UPDATED: Removed AccCapex (was 0) and Checkbox (was "") from all rows.
     */
    rows: {
        _data: {
            standard:        ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 12, "Choose Action", "", "", "", "", "Pending Approval", "", "", ""],
            withAccessories: ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 12, "Choose Action", "", "", "", "", "Pending Approval", "", "", ""], // Note: AccCapex is gone, so this is same as standard now
            approverPrice:   ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 12, "Choose Action", "", 96, "", "", "Pending Approval", "", "", ""],
            approved:        ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 12, "Choose Action", "", "", "", "", "Approved (New Price)", 90, "approver@test.com", "2023-01-01"]
        },
        get: function(name) {
            return JSON.parse(JSON.stringify(this._data[name]));
        }
    },

    /**
     * Mock row data specifically for ApprovalWorkflow unit tests.
     * The `get` method ensures a deep copy is returned for each test.
     * UPDATED: Removed AccCapex and Checkbox columns. Updated Approver Action values.
     */
    rowsForApprovalTests: {
        _data: {
            // This row is now invalid for the new logic, kept for legacy tests that might need it
            approvedOriginalProcessed: ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 24, "Approve Original Price", "", "", "", "", "Approved (Original Price)", 100, "old.approver@example.com", "2025-01-01T10:00:00.000Z"],
            pendingApprovedOriginal:   ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 24, "Approve Original Price", "", "", "1.2", "28800", "Pending Approval", "", "", ""],
            pendingApprovedNew:        ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 24, "Approve New Price", "", 90, "1.08", "25920", "Pending Approval", "", "", ""],
            pendingApproveNoPrice:     ["", "", "", "", "", "", "", "Test Device", 1000, 1200, "", 10, 24, "Approve Original Price", "", "", "", "", "Pending Approval", "", "", ""],
            pendingRejected:           ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 24, "Reject with Comment", "Comment available", "", "1.2", "28800", "Pending Approval", "", "", ""],
            pendingRejectNoComment:    ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 24, "Reject with Comment", "", "", "", "", "Pending Approval", "", "", ""],
            // 'Request Revision' is no longer a direct approver action. This mock is now obsolete for new tests.
            pendingRevision:           ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 24, "Choose Action", "Comment available", "", "", "", "Pending Approval", "", "", ""],
            pendingRevisionNoComment:  ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 24, "Choose Action", "", "", "", "", "Pending Approval", "", "", ""],
            pendingInvalidStatus:      ["", "", "", "", "", "", "", "Test Device", 1000, 1200, 100, 10, 24, "Approve Original Price", "", "", "", "", "Draft", "", "", ""],
        },
        get: function(name) {
            return JSON.parse(JSON.stringify(this._data[name]));
        }
    },

    /**
     * Mock row data specifically for SheetStatusLogic unit tests.
     * The `get` method ensures a deep copy is returned for each test.
     * UPDATED: Removed AccCapex and Checkbox columns. Default action is "Choose Action".
     */
    rowsForStatusLogicTests: {
        _data: {
            blank:             ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            draftIncomplete:   ["SKU-DRAFT", "", "", "", "", "1", "", "Test Device Draft", "1000", "1200", "", "10", "", "Choose Action", "", "", "", "", "Draft", "", "", ""],
            pending:           ["SKU-PENDING", "", "", "", "", "4", "", "Test Device Pending", "1000", "1200", "50", "10", "24", "Choose Action", "", "", "", "", "Pending Approval", "", "", ""],
            approved:          ["SKU-APPROVED", "", "", "", "", "2", "", "Test Device Approved", "1000", "1200", "100", "10", "24", "Approve New Price", "", "95", "1.14", "27360", "Approved (New Price)", "95", "approver@test.com", "2025-07-15"],
            revisedByAE:       ["SKU-REVISED", "", "", "", "", "5", "", "Test Device Revised", "1500", "1600", "120", "12", "36", "Choose Action", "", "", "", "", "Revised by AE", "", "", ""],
            rejected:          ["SKU-REJECT", "", "", "", "", "5", "", "Test Device Rejected", "800", "900", "80", "5", "24", "Reject with Comment", "Price too high", "", "", "", "Rejected", "", "approver@test.com", "2025-07-16"]
        },
        get: function(name) {
            // Ensure we return a deep copy to prevent tests from interfering with each other
            return JSON.parse(JSON.stringify(this._data[name]));
        }
    }
};