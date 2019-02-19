/*CWAJEBookedType*/
ajbDefault = 0;//Get all adjustments/Set non-booked
ajbNotbooked = 1;//Set adjustments not booked in GL
ajbBooked = 2;//Set adjustments booked in GL

/*CWAccessRightType*/
artProtectionSetStaff = 0;//Protection - Set Staff
artProtectionSetGroupSetup = 1;//Protection - Set Group Setup
artProtectionSetUserSetup = 2;//Protection - Set User Setup
artProtectionSetProtection = 3;//Protection - Set Protection
artFileCopyTemplate = 4;//File - Copy Template
artFileModifyClientProfile = 5;//File - Modify Client Profile
artFileImport = 6;//File - Import
artFileExport = 7;//File - Export
artFileConsolidate = 8;//File - Consolidate
artFileYearEndCloseAndRollForward = 9;//File - Year End Close & Roll Forward
artFilePrintBatch = 10;//File - Print Batch
artFileLockAndUnlockClientFile = 11;//File - Lock & Unlock Client File
artFileAccessLockedClientFile = 12;//File - Access Locked Client File
artAccountAdjustingJournalEntries = 13;//Account - Adjusting Journal Entries
artAccountOtherJournalEntries = 14;//Account - Other Journal Entries
artAccountWorkingTrialBalance = 15;//Account - Working Trial Balance
artAccountAssignMappingNumbers = 16;//Account - Assign Mapping Numbers
artAccountMapping = 17;//Account - Mapping
artAccountSetupGroups = 18;//Account - Setup Groups
artAccountAllowAddition = 19;//Account - Allow Addition
artAccountAllowDeletion = 20;//Account - Allow Deletion
artAccountAllowChange = 21;//Account - Allow Change
artAccountAllowViewing = 22;//Account - Allow Viewing
artWorkingTrialBalanceAccount = 23;//Working Trial Balance - Account
artWorkingTrialBalanceReport = 24;//Working Trial Balance - Report
artWorkingTrialBalanceCashFlow = 25;//Working Trial Balance - Cash Flow
artWorkingTrialBalanceGrouping = 26;//Working Trial Balance - Grouping
artWorkingTrialBalanceTax = 27;//Working Trial Balance - Tax
artWorkingTrialBalanceForeignExchange = 28;//Working Trial Balance - Foreign Exchange
artWorkingTrialBalancePerformanceMeasures = 29;//Working Trial Balance - Performance Measures
artWorkingTrialBalanceSplitUpAccounts = 30;//Working Trial Balance - Split-up Accounts
artWorkingTrialBalanceReportSettings = 31;//Working Trial Balance - Report Settings
artMappingAllowModification = 32;//Mapping - Allow Modification
artMappingAllowAddition = 33;//Mapping - Allow Addition
artMappingAllowDeletion = 34;//Mapping - Allow Deletion
artMappingAllowChange = 35;//Mapping - Allow Change
artMappingAllowViewing = 36;//Mapping - Allow Viewing
artDocumentSetDocumentAssignment = 37;//Document - Set Document Assignment
artDocumentAddNewDocument = 38;//Document - Add New Document
artDocumentAllowOpenWithoutAssignment = 39;//Document - Allow Open Without Assignment
artDocumentAllowDeleteWithoutAssignment = 40;//Document - Allow Delete Without Assignment
artDocumentAllowModifyWithoutAssignment = 41;//Document - Allow Modify Without Assignment
artDocumentCanSignOffRole1 = 42;//Document - Can Sign Off Role 1
artDocumentCanSignOffRole2 = 43;//Document - Can Sign Off Role 2
artDocumentCanSignOffRole3 = 44;//Document - Can Sign Off Role 3
artDocumentCanSignOffRole4 = 45;//Document - Can Sign Off Role 4
artDocumentAllowSignOffByAnyone = 46;//Document - Allow Sign Off By Anyone
artToolsRepairFile = 47;//Tools - Repair File
artToolsDiagnostics = 48;//Tools - Diagnostics
artToolsTickmarks = 49;//Tools - Tickmarks
artToolsHelperApplications = 50;//Tools - Helper Applications
artToolsFontSettings = 51;//Tools - Font Settings
artToolsSetAnalyticalReview = 52;//Tools - Set Analytical Review
artToolsCustomize = 53;//Tools - Customize
artToolsUndoCheckOut = 54;//Tools - Maintenance - Undo Sign/Check Out
artOptionsRoles = 55;//Options - Roles
artOptionsDocument = 56;//Options - Document
artOptionsListsJournals = 57;//Options - Lists : Journals
artOptionsListsTaxCodes = 58;//Options - Lists : Tax Codes
artOptionsListsUnits = 59;//Options - Lists : Units
artOptionsAdvanced = 60;//Options - Advanced
artFileArchive = 61;//File - Archive
artDocumentCanSignOffRole5 = 62;//Document - Can Sign Off Role 5
artDocumentCanSignOffRole6 = 63;//Document - Can Sign Off Role 6
artDocumentCanSignOffRole7 = 64;//Document - Can Sign Off Role 6
artDocumentCanSignOffRole8 = 65;//Document - Can Sign Off Role 8
artProtectionAllowCopyTempSecuritySettings = 66;//Protection - Allow Copy Template Security Settings
artDocumentAllowSignOffWithoutAssignment = 67;//Document - Allow Sign Off Without Assignment
artOptionsListsCustomBalances = 68;//Options - Lists : Custom Balances
artFileModifyIssuesCreatedByOtherUsers = 69;//File - Modify issues created by other users
artProtectionAllowAssignUserToOwnGroups = 70;//Protection - Allow assignment of users to own groups
artFileModifyHistorySettings = 71;//File - Modify History Settings
artFileAllowSignInFilesSignedOutByOthers = 72;//File - Allow signing in of files signed out by others
artFileLockDown = 73;//File - Lock Down
artAccountAllowLockingUnlocking = 74;//Account - Allow locking/unlocking of accounts
artOptionsLayout = 75;//Options - Layout
artFileLockClientFile = 76;//File - Lock Client File
artFileUnlockClientFile = 77;//File - Unlock Client File
artFileIssueCompleteWithoutAssigned = 78;//File - Complete issue without being assigned to issue
artDocumentCanModifySignOffDate = 79;//Document - Can modify sign off date
artFileIssueModifyTypes = 80;//File - Can modify issue types
artToolsDeleteMilestones = 81;//Tools - Maintenance - Delete Milestones
artAccountAJAllowDeletion = 82;//Account - Allow deletion of adjusting journal entries
artToolsAbandonSyncCopies = 83;//Tools - Maintenance - Abandon Synchronized Copies
artDocumentSignoffBreakRequirements = 84;//Document - Allow role sign offs that break role requirements
artDocumentCanModifySharedFilters = 85;//Document - Can modify shared filters
artDocumentDeleteRecyclebin = 86;//Document - Delete documents from recycle bin
artFileOpenDirectly = 87;//File - Open client file directly
artFileTemplateUpdate = 88;//File - Apply updates from template
artFilePublishClientFile	;//File - Publish Client File/ Replace Server Copy	89
artOptionsLanguages = 90;//Options - Languages

/*CWAccountType*/
atBalanceSheet = 0;//Balance Sheet
atIncomeStatement = 1;//Income Statement
atCashFlow = 2;//Cash Flow
atPerformanceMeasures = 3;//Performance Measures

/*CWAccumulatorType*/
atNA	N/A	0;//
atLevel1 = 1;//Level 1
atLevel2 = 2;//Level 2
atLevel3 = 3;//Level 3
atLevel4 = 4;//Level 4
atLevel5 = 5;//Level 5
atLevel6 = 6;//Level 6
atClear = 7;//Clear Accumulators

/*CWAdjustingEntryType*/
aetNormal = 0;//Normal Adjusting
aetReclassifying = 1;//Reclassifying
aetUnrecordedProposed = 2;//Unrecorded - Proposed (Factual)
aetEliminating = 3;//Eliminating
aetTaxFederal = 4;//Tax - Federal
aetTaxState = 5;//Tax - State
aetTaxCity = 6;//Tax - City
aetPriorPeriod = 7;//Prior Period
aetOtherBasis = 8;//Other Basis
aetUnrecordedProjected = 9;//Unrecorded - Projected
aetUnrecordedJudgmental = 10;//Unrecorded - Judgmental
aetOther2 = 11;//Other adjusting 2
aetOther3 = 12;//Other adjusting 3

/*CWAdjustmentType*/
atNormal = 0;//Normal
atReclassifying = 1;//Reclassifying
atUnrecorded = 2;//Unrecorded - Proposed
atEliminating = 3;//Eliminating
atTaxFederal = 4;//Tax - Federal
atTaxState = 5;//Tax - State
atTaxCity = 6;//Tax - City
atOtherBasis = 7;//Other Basis
atUnrecProjected = 8;//Unrecorded - Projected
atUnrecJudgmental = 9;//Unrecorded - Judgmental
atOther2 = 10;//Other adjusting 2
atOther3 = 11;//Other adjusting 3

/*CWAnnotationPointType*/
apAccount = 0;//Account
apAJE = 1;//Adjusting journal entry
apAJEItem = 2;//Adjusting journal entry item
apOtherEntry = 3;//Other entry
apDocument = 4;//Document level
apOther = 5;//Other

/*CWAnnotationPointType*/
apAccount = 0;//Account
apAJE = 1;//Adjusting journal entry
apAJEItem = 2;//Adjusting journal entry item
apOtherEntry = 3;//Other entry
apDocument = 4;//Document level
apOther = 5;//Other

/*CWAnnotationType*/
anUnknown = 0;//Unknown
anNote = 1;//Note
anTickmark = 2;//Tickmark
anReference = 3;//Reference
anLegacy = 4;//Legacy
anUserText = 5;//User Text

/*CWAssignToType*/
cwaNA = 0;//N/A
cwaAssignToMe = 1;//Assign to me

/*CWAssignedIssueType*/
cwAssignedNone = 0x0000;//None
cwAssignedToMe = 0x0001;//Assigned to the current user
cwCompleted = 0x0002;//Completed
cwUncleared = 0x0003;//Uncleared
cwAny = 0x0004;//Any

/*CWAutoCVRoundingBalanceType*/
cwrConsolidated = 0;//Consolidated
cwrReport = 1;//Report
cwrAdjusted = 2;//Adjusted
cwrUnadjusted = 3;//Unadjusted
cwrTaxFederal = 4;//Tax - Federal
cwrTaxState = 5;//Tax - State
cwrTaxCity = 6;//Tax - City
cwrConsolidatedInclOBA = 7;//Consolidated incl OBA
cwrReportInclOBA = 8;//Report incl OBA
cwrAdjustedInclOBA = 9;//Adjusted incl OBA
cwrUnadjustedInclOBA = 10;//Unadjusted incl OBA
cwrTaxFederalInclOBA = 11;//Tax - Federal incl OBA
cwrTaxStateInclOBA = 12;//Tax - State incl OBA
cwrTaxCityInclOBA = 13;//Tax - City incl OBA

/*CWAutoCVRoundingCurrentIncludeAdjustmentsTo*/
cwraNotApplicable = 0;//N/A
cwraMapNumber = 1;Map Number
cwraGroup1 = 2 ;//Group 1
cwraGroup2 = 3 ;//Group 2
cwraGroup3 = 4 ;//Group 3
cwraGroup4 = 5 ;//Group 4
cwraGroup5 = 6 ;//Group 5
cwraGroup6 = 7 ;//Group 6
cwraGroup7 = 8 ;//Group 7
cwraGroup8 = 9 ;//Group 8
cwraGroup9 = 10 ;//Group 9
cwraGroup10 = 11 ;//Group 10

/*CWAutoCVRoundingPriorIncludeAdjustmentsTo*/
cwrapNotApplicable = 0 ;//N/A
cwrapSameAsCurrent = 1 ;//Same as current year

/*CWAutoCVRoundingType*/
cvrtNone = 0 ;//None
cvrtOnes = 1 ;//Ones
cvrtThousands = 2 ;//Thousands
cvrtMillions = 3 ;//Millions
cvrtBillions = 4 ;//Billions
cvrtTenThousands = 5 ;//TenThousands
cvrtHundredMillions = 6 ;//HundredMillions

/*CWAutoCompressType*/
actNo = 0 ;//No
actYes = 1 ;//Yes
actPrompt = 2 ;//Prompt

/*CWAutoDocTypeFormatType*/
adttNone = 0 ;//None
adttAccountAnalysisAccountDate = 1;//Account Analysis - Account and Date Order
adttAccountAnalysisAccountDateSubtotal = 2;//Account Analysis - Account and Date Order with Subtotals
adttAccountAnalysisMapNumberDate = 3;//Account Analysis - Map Number and Date Order
adttAccountAnalysisMapNumberDateSubtotal = 4;//Account Analysis - Map Number and Date Order with Subtotals
adttAccountReconciliationClearItems	Account  5;//Reconciliation - Clear Items for Reconciliation
adttAccountReconciliationOutstandingDebitCredit	Account 	6 ;//Reconciliation - Outstanding Debits and Credits
adttAccountReconciliationOutstandingDebit = 7 ;//Account Reconciliation - Outstanding Debits
adttAccountReconciliationOutstandingCredit = 8 ;//Account Reconciliation - Outstanding Credits
adttAccountReconciliationAutomatic = 9 ;//Account Reconciliation - Automatic
adttAccountReconciliationManual = 10 ;//Account Reconciliation - Manual
adttAnalyticalReviewRatioAnalysis = 11 ;//Analytical Review - Ratio Analysis
adttAnalyticalReviewVarianceAccount = 12 ;//Analytical Review - Variance by Account
adttAnalyticalReviewVarianceAccountGroupSubtotals = 13 ;//Analytical Review - Variance by Account with Group Subtotals
adttAnalyticalReviewVarianceGroupGroupSubtotals = 14 ;//Analytical Review - Variance by Group with Group Subtotals
adttAnalyticalReviewVarianceGroupSummary = 15 ;//Analytical Review - Variance by Group Summary
adttChartOfAccountsAlphabetical = 16 ;//Chart of Accounts - Alphabetical Order
adttChartOfAccountsNumerical = 17 ;//Chart of Accounts - Numerical Order
adttChartOfMappingNumbersAlphabetical = 18 ;//Chart of Mapping Numbers - Alphebetical Order
adttChartOfMappingNumbersNumerical = 19 ;//Chart of Mapping Numbers - Numerical Order
adttChartOfMappingNumbersReport = 20 ;//Chart of Mapping Numbers - Report Order
adttClientProfile = 21 ;//Client Profile
adttDiagnosticsMappingStructure = 22 ;//Diagnostics - Mapping Structure
adttDiagnosticsSetup = 23 ;//Diagnostics - Setup
adttDiagnosticsOutOfBalance = 24 ;//Diagnostics - Out of Balance Entries
adttDiagnosticsAccountStructure = 25 ;//Diagnostics - Account Structure/Balances
adttDiagnosticsRoleCompletion = 26 ;//Diagnostics - Role Completion
adttDiagnosticsTaxStructure = 27 ;//Diagnostics - Tax Structure
adttDocumentIndex = 28 ;//Document Index
adttDocumentManagerRoleCompletion = 29 ;//Document Manager - Role Completion Only
adttDocumentManagerRoleCompletionProperties = 30 ;//Document Manager - Role Completion and Properties
adttBalanceSheetActiveYTD = 31 ;//Financial Statements - Balance Sheet - Active YTD
adttBalanceSheetActiveYTDAC = 32 ;//Financial Statements - Balance Sheet - Active YTD A/C
adttBalanceSheetActiveBudget = 33 ;//Financial Statements - Balance Sheet - Active/Budget
adttBalanceSheetActivePrior = 34 ;//Financial Statements - Balance Sheet - Active/Prior
adttBalanceSheetActivePriorYTD = 35 ;//Financial Statements - Balance Sheet - Active/Prior YTD
adttBalanceSheetActiveBudgetPriorYTD = 36 ;//Financial Statements - Balance Sheet - Active/Budget/Prior YTD
adttBalanceSheetActiveBudgetYTD = 37 ;//Financial Statements - Balance Sheet - Active/Budget YTD
adttBalanceSheetDraftActivePriorYTD = 38 ;//Financial Statements - Balance Sheet - Draft Active/Prior YTD
adttBalanceSheetFiveYearHistory = 39 ;//Financial Statements - Balance Sheet - Five Year History
adttBalanceSheetActiveYTDPriorYTD = 40 ;//Financial Statements - Balance Sheet - Active YTD/Prior YTD
adttIncomeStatementActiveYTD = 41 ;//Financial Statements - Income Statement - Active YTD
adttIncomeStatementActiveYTDAC = 42 ;//Financial Statements - Income Statement - Active YTD A/C
adttIncomeStatementActiveYTD2 = 43 ;//Financial Statements - Income Statement - Active YTD 2
adttIncomeStatementActiveBudget = 44 ;//Financial Statements - Income Statement - Active/Budget
adttIncomeStatementActivePrior = 45 ;//Financial Statements - Income Statement - Active/Prior
adttIncomeStatementActivePriorYTD = 46 ;//Financial Statements - Income Statement - Active/Prior YTD
adttIncomeStatementActiveBudgetPriorYTD = 47 ;//Financial Statements - Income Statement - Active/Budget/Prior YTD
adttIncomeStatementActiveBudgetYTD = 48 ;//Financial Statements - Income Statement - Active/Budget YTD
adttIncomeStatementDraftActivePriorYTD = 49 ;//Financial Statements - Income Statement - Draft Active/Prior YTD
adttIncomeStatementFiveYearHistory = 50 ;//Financial Statements - Income Statement - Five Year History
adttIncomeStatementActiveYTDPriorYTD = 51 ;//Financial Statements - Income Statement - Active YTD/Prior YTD
adttStatementOfCashFlowActiveYTD = 52 ;//Financial Statements - Statement of Cash Flow - Active YTD
adttStatementOfCashFlowActiveYTDAC = 53 ;//Financial Statements - Statement of Cash Flow - Active YTD A/C
adttStatementOfCashFlowActiveBudget = 54 ;//Financial Statements - Statement of Cash Flow - Active/Budget
adttStatementOfCashFlowActivePrior = 55 ;//Financial Statements - Statement of Cash Flow - Active/Prior
adttStatementOfCashFlowActivePriorYTD = 56 ;//Financial Statements - Statement of Cash Flow - Active/Prior YTD
adttStatementOfCashFlowActiveBudgetPriorYTD = 57 ;//Financial Statements - Statement of Cash Flow - Active/Budget/Prior YTD
adttStatementOfCashFlowActiveBudgetYTD = 58 ;//Financial Statements - Statement of Cash Flow - Active/Budget YTD
adttStatementOfCashFlowDraftActivePriorYTD = 59 ;//Financial Statements - Statement of Cash Flow - Draft Active/Prior YTD
adttStatementOfCashFlowFiveYearHistory = 60 ;//Financial Statements - Statement of Cash Flow - Five Year History
adttStatementOfCashFlowActiveYTDPriorYTD = 61 ;//Financial Statements - Statement of Cash Flow - Active YTD/Prior YTD
adttGeneralLedgerAccountDateOrder = 62 ;//General Ledger - Account and Date Order
adttGeneralLedgerAccountDateOrderSubtotals = 63 ;//General Ledger - Account and Date Order with Subtotals
adttAdjustingEntriesNormal = 64 ;//Journals - Adjusting Entries - Normal
adttAdjustingEntriesReclassifying = 65 ;//Journals - Adjusting Entries - Reclassifying
adttAdjustingEntriesUnrecorded = 66 ;//Journals - Adjusting Entries - Unrecorded/Proposed
adttAdjustingEntriesEliminating = 67 ;//Journals - Adjusting Entries - Eliminating
adttAdjustingEntriesAll = 68 ;//Journals - Adjusting Entries - All
adttAdjustingEntriesTaxFederal = 69 ;//Journals - Adjusting Entries - Tax - Federal
adttAdjustingEntriesTaxState = 70 ;//Journals - Adjusting Entries - Tax - State
adttAdjustingEntriesTaxCity = 71 ;//Journals - Adjusting Entries - Tax - City
adttAdjustingEntriesPrior = 72 ;//Journals - Adjusting Entries - Prior Period
adttClosingEntries = 73 ;//Journals - Closing Entries
adttOpeningEntries = 74 ;//Journals - Opening Entries
adttAccountsPayable = 75 ;//Journals - Accounts Payable
adttCashDisbursements = 76 ;//Journals - Cash Disbursements
adttCashReciepts = 77 ;//Journals - Cash Reciepts
adttGeneral = 78 ;//Journals - General
adttPayroll = 79 ;//Journals - Payroll
adttPurchaseLedger = 80 ;//Journals - Purchase Ledger
adttStandardEntries = 81 ;//Journals - Standard
adttSales = 82 ;//Journals - Sales
adttCustom = 83 ;//Journals - Custom
adttLeadsheetGroupingAccountBalancesTotals = 84 ;//Leadsheet/Grouping - Account Balances with Totals
adttLeadsheetGroupingAccountBalancesGroupSubtotals = 85 ;//Leadsheet/Grouping - Account Balances with Group Subtotals
adttLeadsheetGroupingGroupSummary = 86 ;//Leadsheet/Grouping - Group Summary
adttLeadsheetGroupingDetailedTotals = 87 ;//Leadsheet/Grouping - Detailed with Totals
adttLeadsheetGroupingHierarchySubtotals = 88 ;//Leadsheet/Grouping - Group Hierarchy with Subtotals
adttLeadsheetGroupingHierarchySummary = 89 ;//Leadsheet/Grouping - Group Hierarchy Summary
adttMemorandum = 90 ;//Memorandum
adttProgramChecklistChecklist1 = 91 ;//Program/Checklist - Checklist - Format 1
adttProgramChecklistChecklist2 = 92 ;//Program/Checklist - Checklist - Format 2
adttProgramChecklistProgram1 = 93 ;//Program/Checklist - Program - Format 1
adttProgramChecklistProgram2 = 94 ;//Program/Checklist - Program - Format 1
adttProgramChecklistProgramAssertions1 = 95 ;//Program/Checklist - Program with Assertions - Format 1
adttProgramChecklistProgramAssertions2 = 96 ;//Program/Checklist - Program with Assertions - Format 2
adttSpreadsheetAnalysisAccruedExpenses = 97 ;//Spreadsheet Analysis - Accrued Expenses
adttSpreadsheetAnalysisCapitalAssets = 98 ;//Spreadsheet Analysis - Capital Assets
adttSpreadsheetAnalysisInvestments = 99 ;//Spreadsheet Analysis - Investments
adttSpreadsheetAnalysisNotesPayable = 100 ;//Spreadsheet Analysis - Notes Payable
adttSpreadsheetAnalysisPrepaidAnalysis = 101 ;//Spreadsheet Analysis - Prepaid Analysis
adttSpreadsheetAnalysisGenericSingle = 102 ;//Spreadsheet Analysis - Generic Single Column
adttSpreadsheetAnalysisGenericDouble = 103 ;//Spreadsheet Analysis - Generic Double Column
adttSpreadsheetAnalysisGenericTriple = 104 ;//Spreadsheet Analysis - Generic Triple Column
adttSpreadsheetAnalysisGenericTwelve = 105 ;//Spreadsheet Analysis - Generic Twelve Column
adttTaxReconciliationDetailTaxLineSubLine = 106 ;//Tax Reconciliation - Detail Tax Line and Sub-Line
adttTaxReconciliationTaxSubLineSummary = 107 ;//Tax Reconciliation - Tax Sub-Line Summary
adttTaxReconciliationTaxLineSummary = 108 ;//Tax Reconciliation - Tax Line Summary
adttTickmarks = 109 ;//Tickmarks
adttTrialBalanceAccountBalances = 110 ;//Trial Balance - Account Balances
adttTrialBalanceAccountFinalBalances = 111 ;//Trial Balance - Account Final Balance
adttTrialBalanceAccountBalancesGroupSubtotals = 112 ;//Trial Balance - Account Balances with Group Subtotals
adttTrialBalanceAccountBalancesSplitUp = 113 ;//Trial Balance - Account Balances with Split-up
adttTrialBalanceGroupHierarchySubtotals = 114 ;//Trial Balance - Group Hierarchy with Subtotals
adttTrialBalanceHierarchySubtotals = 115 ;//Trial Balance - Group Hierarchy Summary
adttTrialBalanceGroupSummary = 116 ;//Trial Balance - Group Summary
adttTrialBalanceDetailedEntriesTotals = 117 ;//Trial Balance - Detailed Entries with Totals
adttUnadjustedDifferences = 118 ;//Uncorrected Misstatements
adttIssuesCreatedBy = 119 ;//Issues - CreatedBy
adttIssuesAssignedTo = 120 ;//Issues - AssignedTo
edttLocal = 32767 ;//Local Document Link
adttAdjustingEntriesOtherBasis = 32768 ;//Journals - Adjusting Entries - Other Basis
adttHistorySignOutIn = adttIssuesAssignedTo+3 ;//History - Sign Out/In
adttHistoryAdjustingJournalEntry = adttIssuesAssignedTo+3 ;//History - Adjusting Journal Entry
adttHistoryFileAccess = adttIssuesAssignedTo+3 ;//History - File Access
adttHistoryDocumentCreation = adttIssuesAssignedTo+3 ;//History - Document Creation
adttHistoryDocumentModification = adttIssuesAssignedTo+3 ;//History - Document Modification
adttHistoryCheckOutIn = adttIssuesAssignedTo+3 ;//History - Check Out/In
adttHistoryRoleCompletion = adttIssuesAssignedTo+3 ;//History - Role Completion
adttHistoryDocumentDeletion = adttIssuesAssignedTo+3 ;//History - Document Deletion
adttHistoryIssueCreation = adttIssuesAssignedTo+3 ;//History - Issue Creation
adttHistoryDocumentAccess = adttIssuesAssignedTo+3 ;//History - Document Access
adttHistoryAllHistoryEvents = adttIssuesAssignedTo+3 ;//History - All History Events
adttHistoryPostLockdownEvents = adttIssuesAssignedTo+3 ;//History - Post Lockdown Events
adttIssuesCompletedBy = adttIssuesAssignedTo+3 ;//Issues - CompletedBy
adttIssuesClearedBy = adttIssuesAssignedTo+3 ;//Issues - ClearedBy
adttAdjustingEntriesProjected = adttIssuesAssignedTo+3 ;//Journals - Adjusting Entries - Unrecorded/Projected
adttAdjustingEntriesJudgmental = adttIssuesAssignedTo+3 ;//Journals - Adjusting Entries - Unrecorded/Judgmental
adttAdjustingEntriesOther2 = adttIssuesAssignedTo+3 ;//Journals - Adjusting Entries - Other Adjusting 2
adttAdjustingEntriesOther3 = adttIssuesAssignedTo+3 ;//Journals - Adjusting Entries - Other Adjusting 3
adttAllOtherEntries = adttIssuesAssignedTo+3 ;//Journals - All Other Entries
adttConsolidationNormal = adttIssuesAssignedTo+3 ;//Consolidation - Normal
adttConsolidationDetails = adttIssuesAssignedTo+3 ;//Consolidation - Details

/*CWAutoFillAction*/
autoFillDefault = 0 ;//Setting from ClientProfile.AutoFillAll
autoFillBlank = 1 ;//Autofill blank properties only
autoFillAll = 2 ;//Autofill all properties
autofillSelected = 3 ;//Autofill selected account properties

/*CWAutoMapField*/
amfId = 1 ;//Id
amfName = 2 ;//Name
amfExDescription = 3 ;//ExDescription
amfTitle = 4 ;//Title
amfAccountType = 5 ;//AccountType
amfSignType = 6 ;//SignType
amfClassType = 7 ;//ClassType
amfGrouping = 8 ;//Grouping
amfg1 = 0x00010008 ;//Grouping.1
amfg2 = 0x00020008 ;//Grouping.2
amfg3 = 0x00030008 ;//Grouping.3
amfg4 = 0x00040008 ;//Grouping.4
amfg5 = 0x00050008 ;//Grouping.5
amfg6 = 0x00060008 ;//Grouping.6
amfg7 = 0x00070008 ;//Grouping.7
amfg8 = 0x00080008 ;//Grouping.8
amfg9 = 0x00090008 ;//Grouping.9
amfg10 = 0x000A0008 ;//Grouping.10
amfAccumulatorType = 9 ;//AccumulatorType
amfTaxCode = 10 ;//TaxCode
amfGIFI = 11 ;//GIFI
amfUnits = 12 ;//Units
amfBehaviourType = 14 ;////BehaviourType
amfTaxExportCode = 15 ;//TaxExportCode
amfMappingFlip = 16 ;//MappingFlip
amfCalculation = 17 ;//Calculation
amfGroupingFlip = 18 ;//GroupingFlip
amfUnits2 = 19 ;//Units2
amfRate = 20 ;//Rate
amfProperty = 29 ;//Property
amfFlipType = 30 ;//FlipType
amfReportProperties = 31 ;//ReportProperties
amfrpAccumulatorType = 0x00010017;//	ReportProperties.AccumulatorType	32 ;//
amfrpPercentMapping = 0x00020017;//	ReportProperties.PercentMapping	33 ;//
amfrpCurrencySymbol = 0x00030017;//	ReportProperties.CurrencySymbol	34 ;//
amfrpUnderlineType = 0x00040017;//	ReportProperties.UnderlineType	35 ;//
amfrpNewPage = 0x00050017;//	ReportProperties.NewPage	36 ;//
amfrpColumnPosition = 0x00060017;//	ReportProperties.ColumnPosition	37 ;//
amfrpPrintFeature = 0x00070017;//	ReportProperties.PrintFeature	38 ;//
amfrpCashFlowCreditAccount = 0x00080017;//	ReportProperties.CashFlowCreditAccount	39 ;//
amfrpCashFlowDebitAccount = 0x00090017;//	ReportProperties.CashFlowDebitAccount	40 ;//
amfrpUpdateCashFlowAccounts = 0x000A0017;//	ReportProperties.UpdateCashFlowAccounts	41 ;//
amfrpCalculateBalanceUsing = 0x000B0017;//	ReportProperties.CalculateBalanceUsing	42 ;//
amfTaxEntityCode = 25;//	TaxEntityCode	43 ;//
amfTaxEntityCodeNone = 0x00000019;//	TaxEntityCode(None)	44 ;//
amfTaxEntityCodeCorporation = 0x00010019;//	TaxEntityCode(Corporation)	45 ;//
amfTaxEntityCodeSCorporation = 0x00020019;//	TaxEntityCode(S Corporation)	46 ;//
amfTaxEntityCodePartnership = 0x00030019;//	TaxEntityCode(Partnership)	47 ;//
amfTaxEntityCodeNonProfit = 0x00040019;//	TaxEntityCode(Non Profit)	48 ;//

/*CWBalanceSourceColumn*/
cwbColumn1 = 0 ;//Column 1
cwbColumn2 = 1 ;//Column 2
cwbColumn3 = 2 ;//Column 3

/*CWBalanceSourceType*/
bstUserDefinedBAL = 0 ;//User Defined Balance
bstForeignExchange = 1 ;//Foreign Exchange
bstPeriodBalances = 2 ;//Period Balances
bstOtherBasisAjustments = 3 ;//Include Other Basis Ajustments
bstDetailedColumnInformation = 4 ;//Include Detailed Column Information

/*CWBehaviourType*/
btNormal = 0 ;//Normal
btTitle = 1 ;//Title
btCalculated = 2 ;//Calculated
btUnassignable = 3 ;//Unassignable
btFiltered = 4 ;//Filtered

/*CWBooleanType*/
cwbNA = 0 ;//N/A
cwbYes = 1 ;//Yes
cwbNo = 2 ;//No

/*CWCalculateBalanceType*/
cbtNA = 0 ;//N/A
cbtActive = 1 ;//Active Period Balance
cbtActiveNegative = 2 ;//Active Period Balance Negative
cbtDifference = 3 ;//Difference between Active and Prior Period Balances
cbtDifferenceNegative = 4 ;//Difference between Active and Prior Period Balances Negative
cbtPrior = 5 ;//Prior Period Balance
cbtPriorNegative = 6 ;//Prior Period Balance Negative

/*CWCaseViewAccessLevelType*/
cvatNoAccess = 0 ;//No Access
cvatLevel1 = 1 ;//Level 1
cvatLevel2 = 2 ;//Level 2
cvatLevel3 = 3 ;//Level 3
cvatLevel4 = 4 ;//Level 4
cvatLevel5 = 5 ;//Level 5
cvatLevel6 = 6 ;//Level 6
cvatLevel7 = 7 ;//Level 7
cvatLevel8 = 8 ;//Level 8
cvatLevel9 = 9 ;//Level 9
cvatLevel10 = 10 ;//Level 10

/*CWClassType*/
ctNone = 0 ;//None
ctTR = 1 ;//Assets - Current - Trade Recievable
ctQA = 2 ;//Assets - Current - Other Quick
ctIN = 3 ;//Assets - Current - Inventory
ctCA = 4 ;//Assets - Current - Other
ctFD = 5 ;//Assets - Capital - Amortized Cost
ctFN = 6 ;//Assets - Capital - Non-Amortized Cost
ctFA = 7 ;//Assets - Capital - Accumulated Amortization
ctIA = 8 ;//Assets - Capital Asset
ctOA = 9 ;//Assets - Other Asset
ctTP = 10 ;//Liabilities - Current - Trade Payable
ctCL = 11 ;//Liabilities - Current - Other
ctLL = 12 ;//Liabilities - Long Term
ctDI = 13 ;//Equity - Dividend
ctRE = 14 ;//Equity - Retained Earnings
ctEQ = 15 ;//Equity - Other
ctSA = 16 ;//Revenue - Credit Sales
ctSO = 17 ;//Revenue - Other Sales
ctRO = 18 ;//Revenue - Other
ctCD = 19 ;//Expenses - Cost Of Sales - Amortization
ctCN = 20 ;//Expenses - Cost Of Sales - Non Cash
ctCI = 21 ;//Expenses - Cost Of Sales - Interest
ctPU = 22 ;//Expenses - Cost Of Sales - Purchase
ctCR = 23 ;//Expenses - Cost Of Sales - Repairs And Maintenance
ctCS = 24 ;//Expenses - Cost Of Sales - Other
ctOD = 25 ;//Expenses - Other - Amortization
ctON = 26 ;//Expenses - Other - Non Cash
ctIT = 27 ;//Expenses - Other - Income Tax Expense
ctOI = 28 ;//Expenses - Other - Interest
ctOR = 29 ;//Expenses - Other - Repairs And Maintenance
ctOE = 30 ;//Expenses - Other - Other

/*CWCloudIntegrationFeatureType*/
citftDisabled = 0 ;//Cloud Integration Feature Disabled
citfEnabledLicenceRequired = 1 ;//Cloud Integration Feature Enabled Licence Required
citfEnabledLicenceNotRequired = 2 ;//Cloud Integration Feature Enabled Licence Not Required

/*CWCloudIntegrationState*/
cistError = -1 ;//An error was generated while trying to determine the Cloud integration state
cistUnknown = 0 ;//The Cloud integration state is unknown
cistNotIntegrated = 1 ;//The client file is not integrated
cistIntegrated = 2 ;//The client file is integrated
cistUnintegrated = 3 ;//The client file was integrated but is now not integrated

/*CWCloudIntegrationType*/
citError = -1 ;//Could not determine the type
citNoIntegration = 0 ;//No integration
citLicencing = 1 ;//Licencing
citFilelessBundle = 2 ;//Fileless bundle
citSmartSync = 3 ;//SmartSync

/*CWColorSchemeType*/
cstDefault = 0 ;//Default
cstLedgerPad = 1 ;//Ledger Pad
cstPostItNote = 2 ;//Post-It Note
cstRainbow = 3 ;//Rainbow
cstSpectrum = 4 ;//Spectrum
cstSubliminal = 5 ;//Subliminal
cstVanilla = 6 ;//Vanilla
cstWave = 7 ;//Wave
cstSubtle = 8 ;//Subtle
cstTracker = 9 ;//Tracker
cstSocial = 10 ;//Social
cstVintage = 11 ;//Vintage
cstWebApp = 12 ;//WebApp

/*CWColumnPositionType*/
cptNA = 0 ;//N/A
cptColumn1 = 1 ;//Column 1
cptColumn2 = 2 ;//Column 2
cptColumn3 = 3 ;//Column 3

/*CWCompletionType*/
cwcNA	N/A	0 ;//
cwcNotStarted = 1 ;//Not started
cwcInProgress = 2 ;//In Progress
cwcNotCompleted = 3 ;//Not completed
cwcComplete = 4 ;//Complete

/*CWCompressFlags*/
cDefault = 0 ;//None
cOmitBackups = 1 ;//Omit Backups
cOmitIndices = 2 ;//Omit Database Indices
cOmitBackupAndIndicies = 3 ;//Omit Backups and Database Indices
cCompressAllSubFolders = 4 ;//Compress All SubFolders

/*CWConditionType*/
cdtEqual = 0 ;//Equal
cdtNotEqual = 1 ;//Not Equal
cdtGreaterThan = 2 ;//Greater Than
cdtLessThan = 3 ;//Less Than
cdtGreaterAndEqualThan = 4 ;//Greater And Equal Than
cdtLessAndEqualThan = 5 ;//Less And Equal Than

/*CWCopyFlagType*/
cftNone = 0 ;//None
cftCopy = 1 ;//Copy
cftCopyClearCells = 2 ;//Copy, Clear Cells
cftCopyClearParagraphs = 3 ;//Copy, Clear Paragraphs
cftCopyClearBoth = 4 ;//Copy, Clear Both

/*CWCopyGroupingsToType*/
cgttFirstAvailable = 0 ;//First Available
cgttGrouping1 = 1 ;//Grouping 1
cgttGrouping2 = 2 ;//Grouping 2
cgttGrouping3 = 3 ;//Grouping 3
cgttGrouping4 = 4 ;//Grouping 4
cgttGrouping5 = 5 ;//Grouping 5
cgttGrouping6 = 6 ;//Grouping 6
cgttGrouping7 = 7 ;//Grouping 7
cgttGrouping8 = 8 ;//Grouping 8
cgttGrouping9 = 9 ;//Grouping 9
cgttGrouping10 = 10 ;//Grouping 10

/*CWCopyStructuresToType*/
csttFirstAvailable = 0 ;//First Available
csttStructure1 = 1 ;//Structure 1
csttStructure2 = 2 ;//Structure 2
csttStructure3	3 ;//Structure 3	

/*CWCurrentPeriodCalculatedType*/
cpctIndependent = 0 ;//Independent
cpctThirteen = 1 ;//Thirteen
cpctMonthly = 2 ;//Monthly
cpctBiMonthly = 3 ;//Bi-Monthly
cpctQuarterly = 4 ;//Quarterly
cpctThirdly = 5 ;//Thirdly
cpctSemiAnnualy = 6 ;//Semi-Annual
cpctRandom = 7 ;//Random

/*CWCurrentPeriodType*/
cptPeriod1 = 1 ;//Period 1
cptPeriod2 = 2 ;//Period 2
cptPeriod3 = 3 ;//Period 3
cptPeriod4 = 4 ;//Period 4
cptPeriod5 = 5 ;//Period 5
cptPeriod6 = 6 ;//Period 6
cptPeriod7 = 7 ;//Period 7
cptPeriod8 = 8 ;//Period 8
cptPeriod9 = 9 ;//Period 9
cptPeriod10 = 10 ;//Period 10
cptPeriod11 = 11 ;//Period 11
cptPeriod12 = 12 ;//Period 12
cptPeriod13 = 13 ;//Period 13

/*CWCurrentPeriodTypeType*/
cpttThirteen = 0 ;//Thirteen
cpttMonthly	 ;//Monthly	1
cpttBiMonthly = 2 ;//Bi-Monthly
cpttQuarterly = 3 ;//Quarterly
cpttThirdly = 4 ;//Thirdly
cpttSemiAnnualy = 5 ;//Semi-Annual
cpttYearly = 6 ;//Yearly
cpttRandom = 7 ;//Random

/*CWCustomBalanceYearsType*/
typeCurrentPast = 0 ;//Current/Past
typeCurrentFuture = 1 ;//Current/Future

/*CWCustomYearType*/
yctCurrent = 0 ;//Current Year Custom
yctYear1 = 1 ;//Year 1 Custom
yctYear2 = 2 ;//Year 2 Custom
yctYear3 = 3 ;//Year 3 Custom
yctYear4 = 4 ;//Year 4 Custom

/*CWDifferenceType*/
difftypeTemporary = 0 ;//Temporary
difftypePermanent = 1 ;//Permanent

/*CWDisplayMappingsAsType*/
dmatShort = 0 ;//Short
dmatExtended = 1 ;//Extended
dmatExtendedIf = 2 ;//Extended If Available

/*CWDocBalanceType*/
dbtReport = 0 ;//Report
dbtUnadjusted = 1 ;//Unadjusted
dbtAdjusted = 2 ;//Adjusted
dbtConsolidated = 3 ;//Consolidated
dbtTaxFederal = 4 ;//Tax - Federal
dbtTaxState = 5 ;//Tax - State
dbtTaxCity = 6 ;//Tax - City

/*CWDocBalanceTypeZero*/
dbtzZero = 0 ;//Zero
dbtzReport = 1 ;//Report
dbtzUnadjusted = 2 ;//Unadjusted
dbtzAdjusted = 3 ;//Adjusted
dbtzConsolidated = 4 ;//Consolidated
dbtzTaxFederal = 5 ;//Tax - Federal
dbtzTaxState = 6 ;//Tax - State
dbtzTaxCity = 7 ;//Tax - City

/*CWDocBalancesType*/
dbstActivePrior = 0 ;//Active/Prior
dbstActivePreceding = 1 ;//Active/Preceding Period
dbstActiveBudget = 2 ;//Active/Budget
dbstActiveForecast = 3 ;//Active/Forecast
dbstFiveYear = 4 ;//Five Year
dbstFiveYearBudget = 5 ;//Five Year Budget
dbstFiveYearForecast = 6 ;//Five Year Forecast
dbstUserDefined = 7 ;//User Defined

/*CWDocGroupingType*/
dgtGroup1 = 0 ;//Group 1
dgtGroup2 = 1 ;//Group 2
dgtGroup3 = 2 ;//Group 3
dgtGroup4 = 3 ;//Group 4
dgtGroup5 = 4 ;//Group 5
dgtGroup6 = 5 ;//Group 6
dgtGroup7 = 6 ;//Group 7
dgtGroup8 = 7 ;//Group 8
dgtGroup9 = 8 ;//Group 9
dgtGroup10 = 9 ;//Group 10
dgtMapNumber = 10 ;//Map Number
dgtTaxCode = 11 ;//Tax Code
dgtGIFI = 12 ;//GIFI

/*CWDocOrderType*/
dotAccount = 0 ;//Account
dotMapNumber = 1 ;//Map Number
dotReport = 2 ;//Report
dotGroup1 = 3 ;//Group 1
dotGroup2 = 4 ;//Group 2
dotGroup3 = 5 ;//Group 3
dotGroup4 = 6 ;//Group 4
dotGroup5 = 7 ;//Group 5
dotGroup6 = 8 ;//Group 6
dotGroup7 = 9 ;//Group 7
dotGroup8 = 10 ;//Group 8
dotGroup9 = 11 ;//Group 9
dotGroup10 = 12 ;//Group 10
dotTaxCode = 13 ;//Tax Code
dotGIFI = 14 ;//GIFI

/*CWDocumentCustomBalanceType*/
dcbtNA = 0 ;//N/A
dcbtBudget = 1 ;//Budget
dcbtForecast = 2 ;//Forecast
dcbtCustom = 3 ;//Custom
dcbtAll = 4 ;//All

/*CWDocumentDeleteFlag*/
cwtDeleteNA = 0 ;//Fail if Multiply Linked on Document Manager
cwtDeleteReferenceOnly = 1;//Remove Document Manager Reference Only if Multiply Linked
cwtDeleteReferenceAndDocument = 2;//Remove Document Manager Reference and Document Always
cwtDeleteToRecyclebin = 3;//Send the deleted document to CaseWare Recycle Bin

/*CWDocumentLanguageType*/
langCurrent = 0;//Current Language
langSpecific = 1;//Specific Language

/*CWDocumentSyncFrequency*/
dsfDefault = 0;//Synchronize default
dsfAutomatic = 1;//Synchronize automatic
dsfOndemand = 2;//Synchronize ondemand
dsfNever = 3;//Synchronize never

/*CWDocumentSyncLock*/
dslOpenNormal = 0;//Lock acquired, caller can proceed to open file normally
dslExpiredLock = 1;//A lock was found, but it has expired
dslActiveLock = 2;//An active lock was found, computer and user name returned
dslPendingLock = 3;//File not locked, but latest version of document not available
dslOffline = 4;//SmartSync is currently operating offline
dslInstanceLost = 5;//The remote object instance on the server was lost
dslError = 6;//An error occurred trying to check or acquire the lock
dslNotExclusiveUse = 7;//The document is not an exclusive use document
dslAccessDenied = 8;//The server denied access to the lock

/*CWDocumentSyncLockCheck*/
dslcError = 0;//An error occurred
dslcOffline = 1;//The file is a sync file, but the server is not accessible
dslcNoLock = 2;//The file is a sync file, and is not locked
dslcLockedByOtherUser = 3;//The file is a sync file, and is locked by a different user
dslcLockedByThisUser = 4;//The file is a sync file, and is locked by the current user
dslcFormatError = 5;//The file is a sync file, a lock is present, but there is a problem with its format
dslcInstanceLost = 6;//The remote object instance in the server was lost
dslcAccessDenied = 7;//The server denied access to the lock

/*CWDocumentType*/
dtFolder = 0;//Folder
dtAutomatic	1;//	Automatic Document
dtCaseView = 2;//CaseView Document
dtManual = 3;//Manual Document Reference
dtExternalLink = 4;//External Document Link
dtLink = 5;//Document Link
dtURL = 6;//URL
dtIdea = 7;//IDEA
dtWord = 8;//Word (CW 2006.00.084)
dtExcel = 9;//Excel (CW 2006.00.084)

/*CWElementChangeType*/
ectNA = 0;//N/A
ectInsert = 1;//Insert
ectModify = 2;//Modify
/*ectVersionUpdate	Version Update	ectModify
ectDelete	Delete	ectModify
ectRemap	Remap	ectModify
ectRenumber	Renumber	ectRemap
ectInsertAs	Insert As	ectRemap
ectVersionUpdateCV	Content Version Update	ectRemap*/

/*CWElementUpdateStatus*/
eusRequired = 0;//Required
eusOptional = 1;//Optional
eusDenied = 2;//Denied

/*CWEmployeeType*/
etActive = 0;//Active
etInactive = 1;//Inactive
etActiveTimekeeper = 2;//Active Timekeeper
etInactiveTimekeeper = 3;//Inactive Timekeeper
etDefaultAdministrator = 4;//Default Administrator

/*CWEntitySelectType*/
cweNA	0;//N/A	
cweMultiple = 1;//Multiple Entities
cweDefaultContext = 2;//Default Context
cweSpecificEntities = 3;//Specific Entities

/*CWExportPackage*/
epNoPackageSelected = 0;//No package selected
epAccountView = 1;//Account view
epGIFI = 2;//GIFI

/*CWExportSetting*/
ExpEntitySpecificExport = 1;//Entity specific export
ExpIncludeSubEntities = 2;//Include Sub entities
ExpUseExtendedDescription = 4;//Use extended desciprtion
ExpDateFormatMMDDYY = 8;//MMDDYY format
ExpGIFIBalances = 16;//Export Balances
ExpGIFINotes = 32;//Export CaseView Notes
ExpGIFIFlipBalances = 64;//Flip GIFI balances
ExpGIFIExportBalanceSheet = 128;//Export Opening Balance Sheet
ExpGIFIExportMultipleIncomeStatements = 256;//Export Multiple income statements
ExpGIFIExportTopLevelEntity = 512;//Include top level entity
ExpGIFIExportForeignExchange = 1024;//Export Foreign Exchange


/*CWFlipType*/
ftIndividual = 0;//Individual
ftDebit = 1;//Debit
ftCredit = 2;//Credit
ftDebitSubs = 3;//Debit (use subs)
ftCreditSubs = 4;//Credit (use subs)
ftParent = 5;//Based on parent
ftDebitsCredits = 6;//Debits and credits

/*CWFontSettingType*/
fstDocumentNumber = 0;//Document Number
fstRole1 = 1;//Role 1
fstRole2 = 2;//Role 2
fstRole3 = 3;//Role 3
fstRole4 = 4;//Role 4
fstNumbers = 5;//Numbers
fstAnnotation = 6;//Annotation
fstTickmarksAndAnnotation = 7;//Tickmarks And Annotation
fstReference = 8;//Document Reference
fstTable = 9;//Table
fstHeading1 = 10;//Heading 1
fstHeading2 = 11;//Heading 2
fstHeading3 = 12;//Heading 3
fstHeading4 = 13;//Heading 4
fstFooter1 = 14;//Footer 1
fstTitle = 15;//Title
fstText = 16;//Text
fstNumbers1 = 17;//Numbers 1
fstNumbers2 = 18;//Numbers 2
fstNumbers3 = 19;//Numbers 3
fstCalculatedNumbers = 20;//Calculated Numbers
fstRole5 = 21;//Role 5
fstRole6 = 22;//Role 6
fstRole7 = 23;//Role 7
fstRole8 = 24;//Role 8

/*CWForecastYearType*/
yftCurrent = 0;//Current Year Forecast Year 1
yftFuture1 = 1;//Future Year 1 Forecast
yftFuture2 = 2;//Future Year 2 Forecast
yftFuture3 = 3;//Future Year 3 Forecast
yftFuture4 = 4;//Future Year 4 Forecast

/*CWGetAccountsType*/
gatAccount = 0;//Account
gatReport = 1;//Report
gatCashFlow = 2;//CashFlow
gatPerformance = 3;//Performance
gatDistributable = 4;//Distribution
gatFlipped = 5;//Flipped

/*CWGlobalGroupAutoSyncType*/
syncTypeNone = 0;//Don't automatically assign to all client files
syncTypeAutoAssign = 1;//Assign global group to all client files
syncTypeAutoAssignOverwriteLocal = 2;//Assign global group to all client files and overwrite local group with the same name

/*CWGroupingAccountType*/
gatNA = 0;//N/A
gatBalanceSheet = 1;//Balance Sheet
gatIncomeStatement = 2;//Income Statement
gatStatementOfCashFlow = 3;//Statement of cash flow
gatGroupPerformance = 4;//Performance

/*CWHistoryCategoryType*/
hcUser = 0;//User
hcSignout = 1;//Sign out
hcDocumentCreate = 2;//Document create
hcDocument = 3;//Document
hcCheckOut = 4;//Checkin/out
hcSignoff = 5;//Signoff
hcDocumentDelete = 6;//Document delete
hcMilestone = 7;//Milestone
hcTask = 8;//Task
hcAccess = 9;//Access
hcAdjustingJournalEntry = 10;//Adjusting Journal Entry
hcRecycleBin = 11;//Recycle Bin
hcFileCreated = 12;//File created
hcFileAccess = 13;//File access
hcYEC = 14;//Year End Close
hcTemplateUpdate = 15;//Template update

/*CWHistoryEventType*/
heUser = 0;//User
heDocumentDeleted = 1;//Document delete
heDocumentPropertyChanged = 2;//Prop changed
heDocumentCreated = 3;//Document created
heDocumentCopied = 4;//Document copied
heDocumentPasted = 5;//Document pasted
heDocumentDropped = 6;//Document dropped
heDocumentAccessed = 7;//Document accessed
heDocumentCheckedOut = 8;//Document checked out
heDocumentCheckedIn = 9;//Document checked in
heMilestoneSaved = 10;//Milestone saved
heDocumentSignedOff = 11;//Document signed off
heDocumentSignedOff1 = 12;//Document signed off (role 1)
heDocumentSignedOff2 = 13;//Document signed off (role 2)
heDocumentSignedOff3 = 14;//Document signed off (role 3)
heDocumentSignedOff4 = 15;//Document signed off (role 4)
heFileCheckedOut = 16;//File checked out
heFileCheckedIn = 17;//File checked in
heFileDatesChanged = 18;//File dates changed
heAJECreated = 19;//AJE created
heAJEDeleted = 20;//AJE deleted
heAJEImported = 21;//AJE imported
heTaskCrated = 22;//Task created
heFileSignedOut = 23;//File signed out
heFileSignedIn = 24;//File signed in
heDocumentSignedOff5 = 25;//Document signed off (role 5)
heDocumentSignedOff6 = 26;//Document signed off (role 6)
heDocumentSignedOff7 = 27;//Document signed off (role 7)
heDocumentSignedOff8 = 28;//Document signed off (role 8)
heAJERenumbered = 29;//AJE renumbered
heFileSignOutUndone = 30;//File sign out undone
heDocumentRecycled = 31;//Document recycled
heDocumentRestored = 32;//Document restored
heDocumentDeletedTemplate = 33;//Document deleted template
heDocumentPropertyTemplate = 34;//Document property template
heDocumentVersionTemplate = 35;//Document version template
heDocumentCreatedTemplate = 36;//Document created template
heDocumentMoved = 37;//Document moved
heDocumentConvertedToPDF = 38;//Document converted to PDF
heDocumentTaskCompleted = 39;//Document task completed
heFileAccessedByUser = 40;//File accessed by user
heDocumentLockedDown = 41;//Document locked down
heDocumentMilestonesRegenerated = 42;//Document milestones regenerated
heDocumentMilestonesConverted = 43;//Document milestones converted
heDocumentMilestonesSet = 44;//Document milestones set
heFileMilestonesSet = 45;//File milestones set
heFileLockedDown = 46;//File locked down
heFileUnlocked = 47;//File unlocked
heFileUndoLockDown = 48;//File undo lock down
heFileLockdownSet = 49;//File lockdown set
heFileLockdownRemovePending = 50;//File Remove pending lockdown
heFileLockdownMilestonesRegenerated = 51;//File lockdown milestones regenerated
heFileLockdownMilestonesRegeneratedAutoDocs = 52;//File lockdown auto doc milestones regenerated
heFileCreatedYEC = 53;//File created YEC
heFileCreatedFileNew = 54;//File created File New
heFileCreatedCopyTemplate = 55;//File created copy template
heYEC = 56;//YEC
heTemplateUpdate = 57;//Template update

/*CWImportComponents*/
icTrialBalance = 0;//Trial Balance
icGeneralLedger = 1;//GeneralLedger
icAdjustingJournalEntry = 2;//Adjusting Journal Entry

/*CWImportPackage*/
ipNoPackageSelected = 0;//No package selected
ipAccountView = 1;//Account view
ipASCII = 2;//Generic ASCII
ipProSysEngagement = 3;//ProSystem fx Engagement
ipXAF = 4;//Xml Audit File (XAF)

/*CWImportSetting*/
AVUsePartnerCode = 0;//Use partner code to build entity
AVPriorYearData = 1;//Prior Year Data
AVUseCostCenterCode = 2;//Use cost centre code
AVNonfinancialData = 3;//Non-financial data
AVOverwriteExistingJournals = 4;//Overwrite existing journals
AVOverwriteAccDescription = 5;//Overwrite existing account description
ProSysImportClientInfomation = 6;//Import Client Information
ProSysImportClientDocuments = 7;//Import Client Documents
ProSysImportAdjEntries = 8;//Import adj/reclass/tax entries
ProSysImportGL = 9;//Import General Ledger Detail
XafUpdateYearEndDate = 10;//Update year end date in client profile
XafUpdateClientInformation = 11;//Import client information
XafVersion3OrAbove = 12;//Import version 3.X auditfile
XafLeadCodeAsXBRL = 13;//Import the leadCode as XBRL tag
XafImportJournalOutsideFiscalYear = 14;//Import journal entries outside of the fiscal year

/*CWItemType*/
typeDocument = 0;//CWDocument
typeMapping = 1;//CWMapping
typeGrouping = 2;//CWGrouping
typeEntity = 3;//CWEntity
typeAccount = 4;//CWAccount
typeSplitUpAccount = 5;//CWSplitUpAccount
typeAdjustingEntry = 6;//CWAdjustingEntry
typeAdjustingEntryItem = 7;//CWAdjustingEntryItem
typeOtherEntry = 8;//CWOtherEntry
typeTaxExportCode = 9;//CWTaxExportCode
typeAnnotation = 10;//CWAnnotation
typeTickmark = 11;//CWTickmark
typeIssue = 12;//CWIssue
typeHistoryEvent = 13;//CWHistoryEvent
typeTaxonomy = 14;//CWTaxonomy
typeTaxonElement = 15;//CWTaxonElement
typeSecurityEmployee = 16;//CWSecurityEmployee
typeCustomBalanceIdentifier = 17;//CWCustomBalanceIdentifier
typeAccountDistribution = 18;//CWAccountDistribution
typeAccountDistributionItem = 19;//CWAccountDistributionItem
typeGlobalGroup = 20;//CCWGlobalGroup
typeGlobalUser = 21;//CCWGlobalUser
typePurge = 22;//CCWPurge
typeLanguageIdentifier = 23;//CWLanguageIdentifier
typeTemplateToolbarButton = 24;//CWTemplateToolbarButton
typeRoleSet = 25;//CWRoleSet

/*CWJournalAccountType*/
jatFinancial = 0;//Financial
jatCashFlow = 1;//Cash Flow
jatPerformance = 2;//Performance
jatGrouping1 = 3;//Grouping 1
jatGrouping2 = 4;//Grouping 2
jatGrouping3 = 5;//Grouping 3
jatGrouping4 = 6;//Grouping 4
jatGrouping5 = 7;//Grouping 5
jatGrouping6 = 8;//Grouping 6
jatGrouping7 = 9;//Grouping 7
jatGrouping8 = 10;//Grouping 8
jatGrouping9 = 11;//Grouping 9
jatGrouping10 = 12;//Grouping 10
jatMapping = 13;//Mapping
jatTaxCode = 14;//Tax Code

/*CWJournalApplyType*/
japtNA = 0;//N/A
japtAll = 1;//All
japtGrouping1 = 2;//Grouping 1
japtGrouping2 = 3;//Grouping 2
japtGrouping3 = 4;//Grouping 3
japtGrouping4 = 5;//Grouping 4
japtGrouping5 = 6;//Grouping 5
japtGrouping6 = 7;//Grouping 6
japtGrouping7 = 8;//Grouping 7
japtGrouping8 = 9;//Grouping 8
japtGrouping9 = 10;//Grouping 9
japtGrouping10 = 11;//Grouping 10
japtTaxExportCode = 12;//Tax Export Code


/*CWJournalBalanceType*/
jbtNA = 0;//N/A
jbtBudget = 1;//Budget
jbtForecast = 2;//Forecast
jbtCustom = 3;//Custom

/*CWLanguageComponentType*/
lctMapping = 0;//Mapping
lctGroup1 = 1;//Group 1
lctGroup2 = 2;//Group 2
lctGroup3 = 3;//Group 3
lctGroup4 = 4;//Group 4
lctGroup5 = 5;//Group 5
lctGroup6 = 6;//Group 6
lctGroup7 = 7;//Group 7
lctGroup8 = 8;//Group 8
lctGroup9 = 9;//Group 9
lctGroup10 = 10;//Group 10
lctDocuments = 11;//Documents
lctAccounts = 12;//Accounts

/*CWListFieldStyleType*/
lfstDescriptions = 0;//Descriptions
lfstAbbreviations = 1;//Abbreviations
lfstEntryCodes = 2;//Entry Codes

/*CWLockReturnType*/
lrtSuccess = 0;//Lock successfully obtained
lrtAlreadyLocked = 1;//File is already locked by another process
lrtBadClient = 2;//The client file specified does not exist or other file error
lrtAlreadyOpened = 3;//File is already opened by another process
lrtBadLock = 4;//Open or create lock file failed

/*CWMappingAccountType*/
matBalanceSheet = 0;//Balance Sheet
matIncomeStatement = 1;//Income Statement
matStatementOfCashFlow = 2;//Statement Of Cash Flow
matPerformance = 3;//Performance

/*CWMileStoneSaveType*/
msNothing = 0;//Nothing
msChecksum = 1;//Checksum
msPerEventType = 2;//Per Event Type
msPerUser = 3;//Per User
msPerEvent = 4;//Per Event

/*CWMisstatementType*/
mitNA = 0;//Not Applicable
mitFactual = 1;//Factual Misstatement
mitProjected = 2;//Projected Misstatement
mitJudgmental = 3;//Judgmental Misstatement
mitAll = 4;//All

/*CWMultipleContainsType*/
cwnNone = 0x0000;//None
cwmTickmarks = 0x0001;//Tickmarks
cwmnotes = 0x0002;//Notes
cwmCommentaryText = 0x0003;//Commentary Text
cwmDocumentReference = 0x0004;//Document Reference

/*CWMultipleDocType*/
cwmNone = 0x0000;//None
cwmAutomatic = 0x0001;//Automatic documents
cwmCaseView = 0x0002;//CaseView documents
cwmExternal = 0x0003;//External documents
cwmManual = 0x0004;//Manual documents
cwmPlacehoder = 0x0005;//Placehoder documents