/**
 * @file This file contains mock data used for integration tests.
 * ALL DATA has been refactored to the new 22-column layout.
 */

const MOCK_DATA_INTEGRATION = {

  groupingTestsAsArray: [
    ['SKU','EP CAPEX','Telekom CAPEX','Target','Limit','Index','Bundle Number','Device','AE EP CAPEX','AE TK CAPEX','AE SALES ASK','QUANTITY','TERM','APPROVER_ACTION','APPROVER_COMMENTS','APPROVER_PRICE_PROPOSAL','LRF_PREVIEW','CONTRACT_VALUE','STATUS','FINANCE_APPROVED_PRICE','APPROVED_BY','APPROVAL_DATE'],
    ['SKU-A','','','','','1','','Individual Approved A','','',50.00,1,24,'Choose Action','','','','','Approved (Original Price)',50.00,'approver@test.com','2025-01-01'],
    ['SKU-B','','','','','2',"101",'Complete Bundle A (Cheaper)','','',25.50,10,24,'Choose Action','','','','','Approved (Original Price)',25.50,'approver@test.com','2025-01-01'],
    ['SKU-C','','','','','3',"101",'Complete Bundle B (Pricier)','','',30.00,10,24,'Choose Action','','','','','Approved (New Price)',30.00,'approver@test.com','2025-01-01'],
    ['SKU-D','','','','','4','','Individual Draft','','',40.00,1,24,'Choose Action','','','','','Draft','','',''],
    ['SKU-E','','','','','5','','Individual Approved B','','',100.00,2,36,'Choose Action','','','','','Approved (New Price)',100.00,'approver@test.com','2025-01-01'],
    ['SKU-F','','','','','6',"202",'Incomplete Bundle A (Approved)','','',15.00,5,12,'Choose Action','','','','','Approved (Original Price)',15.00,'approver@test.com','2025-01-01'],
    ['SKU-G','','','','','7',"202",'Incomplete Bundle B (Pending)','','',20.00,5,12,'Choose Action','','','','','Pending Approval','','','']
  ],

  csvForApprovalWorkflowTests:
`SKU,EP CAPEX,Telekom CAPEX,Target,Limit,Index,Bundle Number,Device,AE EP CAPEX,AE TK CAPEX,AE SALES ASK,QUANTITY,TERM,APPROVER_ACTION,APPROVER_COMMENTS,APPROVER_PRICE_PROPOSAL,LRF_PREVIEW,CONTRACT_VALUE,STATUS,FINANCE_APPROVED_PRICE,APPROVED_BY,APPROVAL_DATE
SKU-01,"","","","",1,,"Device A","1000","1200","100","10","24","Choose Action","","","","",Pending Approval,"","",""
SKU-02,"","","","",2,,"Device B","800","900","80","12","24","Choose Action","","90","","",Pending Approval,"","",""
`,

  csvForHealthCheckTests: [
    ['SKU','EP CAPEX','TK CAPEX','Target','Limit','Index','Bundle Number','Model','AE EP CAPEX','AE TK CAPEX','AE SALES ASK','Qty','Term','Action','Comments','Approver Price','LRF','Contract Value','Status','Finance Price','Approved By','Approval Date'],
    ['SKU-01','','','','',1,'','Approved but no date',1000,1200,100,10,24,'Choose Action','','','','','Approved (Original Price)',100,'approver@test.com',''],
    ['SKU-02','','','','',2,'','Healthy approved row',800,900,80,12,24,'Choose Action','','','','','Approved (Original Price)',80,'approver@test.com','2025-01-01'],
    ['SKU-03','','','','',3,'','Rejected but no date',500,600,50,5,36,'Choose Action','','','','','Rejected','','approver@test.com','']
  ],

  csvForUxControlTests:
`SKU,EP CAPEX,Telekom CAPEX,Target,Limit,Index,Bundle Number,Device,AE EP CAPEX,AE TK CAPEX,AE SALES ASK,QUANTITY,TERM,APPROVER_ACTION,APPROVER_COMMENTS,APPROVER_PRICE_PROPOSAL,LRF_PREVIEW,CONTRACT_VALUE,STATUS
SKU-DRAFT,,,,,"1",,"Draft Device",1000,1200,100,10,24,"Choose Action","","","","",Draft
SKU-PEND,,,,,"2",,"Pending Device",1000,1200,100,10,24,"Choose Action","","","","",Pending Approval
SKU-APP-O,,,,,"3",,"Approved (O) Device",1000,1200,100,10,24,"Choose Action","","","","",Approved (Original Price)
,,,,,"4",,"No Status Row",,,,,,,,,,,
SKU-REJ,,,,,"5",,"Rejected Device",1000,1200,100,10,24,"Choose Action","","","","",Rejected
SKU-REV-AE,,,,,"6",,"Revised AE Device",1000,1200,100,10,24,"Choose Action","","","","",Revised by AE
`,

  /**
   * REWRITTEN for clarity and logical correctness.
   * This data is used to test paste operations and sanitization logic.
   */
  csvForSanitizationTests:
`SKU,EP CAPEX,TK CAPEX,Target,Limit,Index,Bundle Number,Model,AE EP CAPEX,AE TK CAPEX,AE SALES ASK,Qty,Term,Action,Comments,Approver Price,LRF,Contract Value,Status,Finance Price,Approved By,Approval Date
S1_SOURCE_APPROVED,900,850,45,40,16,,Source Approved Model,900,850,50,20,24,"Choose Action",,"",1.2,21600,Approved (Original Price),50,approver@test.com,2025-07-21
S2_TARGET_APPROVED,1200,1100,65,60,17,,Target Approved Model,1200,1100,70,10,36,"Choose Action",,"",1.4,25200,Approved (Original Price),70,approver@test.com,2025-07-22
S1_DESTINATION_BLANK,,,,,,,,,,,,,,,,,,,,,
S4_TARGET_PENDING,500,450,25,20,18,,Target Pending Model,500,450,30,5,12,"Choose Action",,"",0.72,1800,Pending Approval,,,
S5_PASTE_SOURCE,,,,,99,,Pasted Model For S5,600,550,35,7,18,"Choose Action",,,,,,,,,
`,

  csvForBundleValidationTests:
`SKU,EP CAPEX,Telekom CAPEX,Target,Limit,Index,Bundle Number,Device,AE EP CAPEX,AE TK CAPEX,AE SALES ASK,QUANTITY,TERM,APPROVER_ACTION,APPROVER_COMMENTS,APPROVER_PRICE_PROPOSAL,LRF_PREVIEW,CONTRACT_VALUE,STATUS
,,,,,"1",101,"Valid Bundle A",,,"",10,24,,,,,,
,,,,,"2",101,"Valid Bundle B",,,"",10,24,,,,,,
,,,,,"3",,"Single Item",,,"",1,12,,,,,,
,,,,,"4",202,"Valid Bundle C1",,,"",5,36,,,,,,
,,,,,"5",202,"Valid Bundle C2",,,"",5,36,,,,,,
,,,,,"6",202,"Valid Bundle C3",,,"",5,36,,,,,,
,,,,,"7",303,"Single w/ Bundle #",,,"",1,24,,,,,,
,,,,,"8",404,"Non-Consecutive A",,,"",2,24,,,,,,
,,,,,"9",999,"INTERRUPTING ROW",,,"",99,99,,,,,,
,,,,,"10",404,"Non-Consecutive B",,,"",2,24,,,,,,
,,,,,"11",505,"Mismatch Qty A",,,"",15,24,,,,,,
,,,,,"12",505,"Mismatch Qty B",,,"",16,24,,,,,,
,,,,,"13",606,"Mismatch Term A",,,"",20,24,,,,,,
,,,,,"14",606,"Mismatch Term B",,,"",20,36,,,,,,
`,
};