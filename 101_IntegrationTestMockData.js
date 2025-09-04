/**
 * @file This file contains mock data used for integration tests.
 * ALL DATA has been refactored to the new 23-column layout.
 */

const MOCK_DATA_INTEGRATION = {
  // This data is derived from the provided CSV and is the new source of truth for many tests.
  // It features approved items, pending items, draft items, and a bundle.
  groupingTestsAsArray: [
    // Header Row (23 columns)
    ["SKU", "EP CAPEX", "EP Target 24", "EP Target 36", "TK CAPEX", "TK Target 24", "TK Target 36", "Index", "Bundle Number", "Model", "AE CAPEX", "AE SALES ASK", "QUANTITY", "TERM", "APPROVER_ACTION", "APPROVER_COMMENTS", "APPROVER_PRICE_PROPOSAL", "LRF_PREVIEW", "CONTRACT_VALUE", "STATUS", "FINANCE_APPROVED_PRICE", "APPROVED_BY", "APPROVAL_DATE"],
    // Data Rows
    ["118404", 1172.00, 61.99, 47.49, "", "", "", 1, "", "Lenovo ThinkPad T14 G5", 1172.00, 61.99, 10, 24, "Choose Action", "", "", 1.2, 14877.6, "Approved (Original Price)", 61.99, "approver@test.com", "2025-09-01"],
    ["118510", 1311.00, 64.49, 50.49, 1255.00, 57.00, 43.00, 2, "101", "Apple iPhone 16 Pro Max 512GB", 1311.00, 64.49, 5, 24, "Choose Action", "", "", 1.2, 7738.8, "Approved (Original Price)", 64.49, "approver@test.com", "2025-09-01"],
    ["118511", 1522.00, 73.99, 57.49, 1450.00, 65.00, 50.00, 3, "101", "Apple iPhone 16 Pro Max 1TB", 1522.00, 73.99, 5, 24, "Choose Action", "", "", 1.2, 8878.8, "Approved (Original Price)", 73.99, "approver@test.com", "2025-09-01"],
    ["118501", 904.00, 42.99, 34.49, 864.00, 40.00, 30.00, 4, "", "Apple iPhone 16 Pro 128GB", 904.00, 50, 1, 24, "Choose Action", "", "", 1.327, 1200, "Pending Approval", "", "", ""],
    ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "Draft", "", "", ""]
  ],

  // Other mock data, now also in 23-column format
  csvForApprovalWorkflowTests: `SKU,EP CAPEX,EP T24,EP T36,TK CAPEX,TK T24,TK T36,Index,Bundle,Model,AE CAPEX,AE ASK,QTY,TERM,ACTION,COMMENTS,PROPOSAL,LRF,VALUE,STATUS,FINANCE PRICE,APPROVED BY,APPROVED DATE
118501,904,42.99,34.49,864,40,30,4,,iPhone 16 Pro,904,50,1,24,Choose Action,,,,Pending Approval,,,
`,

  csvForHealthCheckTests: [
    ["SKU", "EP CAPEX", "EP T24", "EP T36", "TK CAPEX", "TK T24", "TK T36", "Index", "Bundle", "Model", "AE CAPEX", "AE ASK", "QTY", "TERM", "ACTION", "COMMENTS", "PROPOSAL", "LRF", "VALUE", "STATUS", "FINANCE PRICE", "APPROVED BY", "APPROVED DATE"],
    ["SKU-01", 900, 40, 30, "", "", "", 1, "", "Approved but no date", 900, 40, 10, 24, "Choose Action", "", "", "", "", "Approved (Original Price)", 40, "approver@test.com", ""],
    ["SKU-02", 900, 40, 30, "", "", "", 2, "", "Healthy approved row", 900, 40, 10, 24, "Choose Action", "", "", "", "", "Approved (Original Price)", 40, "approver@test.com", "2025-01-01"],
    ["SKU-03", 600, 30, 20, "", "", "", 3, "", "Rejected but no date", 600, 30, 5, 36, "Choose Action", "", "", "", "", "Rejected", "", "approver@test.com", ""],
  ],

  csvForBundleValidationTests: `SKU,B,C,D,E,F,G,Index,Bundle,Model,K,L,M,N,O,P,Q,R,S,STATUS,U,V,W
,,,,,,,1,101,"Valid Bundle A",,,,,,,,,,,
,,,,,,,2,101,"Valid Bundle B",,,,,,,,,,,
,,,,,,,3,,Single Item,,,,,,,,,,,
,,,,,,,4,404,"Non-Consecutive A",,,,,,,,,,,
,,,,,,,5,999,INTERRUPTING,,,,,,,,,,,
,,,,,,,6,404,"Non-Consecutive B",,,,,,,,,,,
`,
};
