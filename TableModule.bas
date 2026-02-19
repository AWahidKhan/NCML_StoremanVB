Attribute VB_Name = "TableModule"
Dim sql As String
Public UserType As String
Public gFName As String

Public AvgDslRs As Recordset        ' DelTmp.
Public TmpRs As Recordset           ' Attendance Driver
Public TmpRs1 As Recordset          ' Attendance Driver
Public TmpRs2 As Recordset
Public TmpRs3 As Recordset
Public TmpRs4 As Recordset
Public TmpRs5 As Recordset
Public TmpRs6 As Recordset
Public TmpRs7 As Recordset

Public TmpRs0 As Recordset

Public RsMst_Config As Recordset        ' Config

Public RsMst_CoExpense As Recordset

Public RsMst_Assayer As Recordset
Public RsMst_Category As Recordset
Public RsMst_Client As Recordset
Public RsMst_CMP As Recordset
Public RsMst_CommodityGroup As Recordset
Public RsMst_Commodity As Recordset
Public RsMst_CommodityDetail As Recordset
Public RsMst_License As Recordset

Public RsMst_Depository As Recordset
Public RsMst_ExchangeType As Recordset
Public RsMst_Godown As Recordset
Public RsMst_GSL As Recordset
Public RsMst_ISIN As Recordset
Public RsMst_Location As Recordset
Public RsMst_MenuDetail As Recordset
Public RsMst_MenuHasUserRole As Recordset
Public RsMst_RRN As Recordset
Public RsMst_State As Recordset
Public RsMst_StorageContract As Recordset
Public RsMst_StoreManLocations As Recordset
Public RsMst_UnitType As Recordset
Public RsMst_UserRole As Recordset
Public RsMst_Users As Recordset
Public RsTxn_CxTF_ClientDetail As Recordset
Public RsTxn_CxTF_Details As Recordset
Public RsTxn_CxTF_GSL As Recordset
Public RsTxn_CxTF_TruckDetail As Recordset
Public RsTxn_CDC As Recordset
Public RsTxn_CDC_GSLDetail As Recordset
Public RsTxn_CDC_AssayerDetail As Recordset
Public RsTxn_CDC_InsuranceDetail As Recordset
Public RsGSL_ISIN As Recordset

Public RsMst_PolicyDetail As Recordset
Public RsMst_InsuranceCo As Recordset
Public RsTxn_ISINRevalidation As Recordset
Public RsTxn_ISINGSLRevalidation As Recordset
Public RsTxn_WHRDetail As Recordset
Public RsTxn_WHRLicenseDetail As Recordset
Public RsMst_Bank     As Recordset
Public RsMst_Branch       As Recordset
Public RsTxn_Spillage As Recordset
Public RsTxn_PledgeCreation As Recordset
Public RsTxn_PledgeCreation_GSL As Recordset
Public RsTxn_PledgeRelease As Recordset
Public RsTxn_PledgeRelease_GSL As Recordset
Public RsTxn_CashReceipt As Recordset
Public RsTxn_CashReceiptInvoiceDetails As Recordset
Public RsMenuUserRole As Recordset
Public RsMenuDetail As Recordset
Public RsMst_Hologram As Recordset
Public RsLog_PrintHistory As Recordset
Public RsTxn_GSL_OpeningBalance As Recordset
Public RsMst_Agent As Recordset
Public RsTxn_GSL_Transfer As Recordset
Public RsMst_DR As Recordset
Public RsTxn_Benpos   As Recordset
Public RsTxn_BenposTemp   As Recordset
Public RsTxn_BenposDateWise As Recordset
Public RsTxn_InsertBenposDateWise As Recordset
Public RsTxn_AdjTransaction As Recordset
Public RsTxn_GSLAdjTransaction As Recordset
Public RsHelp_Detail As Recordset
Public RsMst_Designation As Recordset
Public RsMst_AssetType As Recordset
Public RsMst_AssetStatus As Recordset
Public RsMst_Employee As Recordset
Public RsMst_EmpCMPMap As Recordset
Public RsMst_EmpAssetMap As Recordset
Public RsMst_SecurityShifts As Recordset
Public RsMst_SecurityAgency As Recordset
Public RsMst_SecurityDetails As Recordset
Public RsMst_SecurityCMPMap As Recordset
Public RsMst_SecurityAssets As Recordset
Public RsMst_SecurityAssetMap As Recordset
Public RsMst_ExpenseType As Recordset
Public RsTxn_ClaimSubmission As Recordset
Public RsTxn_ClaimSubmissionDetails As Recordset


Public RsMst_AuditSegment As Recordset
Public RsMst_AuditType As Recordset
Public RsMst_AuditPoint As Recordset
Public RsTxn_Audit As Recordset
Public RsTxn_AuditDetails As Recordset
Public RsTxn_EmpCTC As Recordset

Public RsTxn_IncomeDetails As Recordset
Public RsTxn_DPIncomeDetails As Recordset
Public RsTxn_TelephoneDetail As Recordset
Public RsTxn_TelephoneTemp As Recordset
Public RsMst_ExpenseGroup As Recordset
Public RsMst_Vendor As Recordset
Public RsMst_ServiceTax As Recordset
Public RsMst_TDS As Recordset

Public RsMst_StockStatus As Recordset

Public RsTmp_DPIncomeDetails As Recordset
Public rstemp As Recordset

Public RsMst_TmpStockStatus As Recordset

Public RsMst_ClientLogin As Recordset
Public RsMst_ClientLoginDetail As Recordset
Public RsTxn_VendorClaim As Recordset
Public RsTxn_VendorClaimDetails As Recordset

Public RsLog_CMPReopenHistory  As Recordset
Public RsMst_CommodityPrice As Recordset
Public RsTxn_SRGodownMapping As Recordset
Public RsTxn_SRGodownMapDetails As Recordset
Public RsGSL_Contract As Recordset
Public RsMst_BagEquivalent As Recordset

Public RsMst_CA As Recordset
Public RsLog_KRDelete  As Recordset
Public RsLog_LoginDetail As Recordset

Public RsTxn_CashReceiptCDTFDetails As Recordset
Public RsTxn_CreateCDTFTemp As Recordset
Public RsTxn_CreatePITemp As Recordset
Public RsTxn_CDTFTemp As Recordset
Public RsTxn_PITemp As Recordset


Public RsMst_BillingCycle As Recordset
Public RsMst_StorageRateCard As Recordset
Public RsMst_WHRUOM As Recordset

Public RsLog_UnmarkPrint_History As Recordset

Public RsMst_EmpLedger As Recordset
Public RsMst_Lessor As Recordset
Public RsTxn_LessorAgreement As Recordset
Public RsTxn_LessorAgreementDetail As Recordset
Public RsTxn_GodownLessorAgreeDetail As Recordset
Public RsLog_UpdateCxTFDetails As Recordset


Public RsTxn_GodownRent As Recordset
Public RsTxn_GodownRentDetail As Recordset


Public RsMst_DPInvoiceUpload As Recordset
Public RsMst_GLAccType As Recordset


Public RsTxn_NewISINRevalidation As Recordset
Public RsTxn_NewISINGSLRevalidation As Recordset

Public RsMst_StorageCardRateDetail As Recordset

Public RsMst_SpecialRateCard As Recordset
Public RsMst_SpecialRateCardDetail As Recordset

Public RsTxn_InvMonGSLDetail As Recordset

Public RsTxn_CRDepositDetails  As Recordset
Public RsMst_FAGBank As Recordset

Public RsTxn_ProInvoice As Recordset
Public RsTxn_TempGSL As Recordset
Public RsTxn_CreateTempGSL As Recordset
Public RsTxn_TempProInvoice As Recordset
Public RsTxn_CreateTempProInvoice As Recordset

Public RsMst_ITVendor As Recordset
Public RsTxn_ITVendorClaim As Recordset
Public RsTxn_ITVendorClaimDetails As Recordset
Public RsMst_ClientGroup As Recordset
Public RsMst_FF_RateCard As Recordset
Public RsMst_FF_RateCard_Godown_Bag As Recordset
Public RsMst_FF_RateCard_Godown As Recordset


Public RsMst_NCDEXStorageRateCard As Recordset
Public RsMst_NCDEXStorageRateCardDetail As Recordset

Public RsTxn_DayWiseGSLDetail  As Recordset
Public RsTxn_CreateTempDayWise As Recordset
Public RsTxn_TempDailyGSL As Recordset

Public RsTxn_Fumigation As Recordset
Public RsTxn_FumigationDetail As Recordset
Public RsTxn_Spraying As Recordset

Public RsTxn_RAGMemberTransfer As Recordset
Public RsTxn_RAGMemberTransferDetail As Recordset


Public RsMst_QualityParameter As Recordset
Public RsMst_CommodityQualityParameter As Recordset
Public RsMst_TestingAgency As Recordset
Public RsTxn_QualityInformation As Recordset
Public RsTxn_QualityInformationDetail  As Recordset

Public RsTxn_CxTF_BOE  As Recordset

Public RsTxn_CreateTempGenInvoice As Recordset
Public RsTxn_TempGenInvoice As Recordset

Public RsTxn_Hologram As Recordset

Public RsTxn_CITFCR As Recordset
Public RsTxn_CITFCRDetail As Recordset

Public RsTxn_OtherIncome As Recordset

Public RsTxn_MonthEmp_History As Recordset
Public RsTxn_MonthEmp_HistoryDetail As Recordset
Public TmpRsHRPro As Recordset

Public RsMst_IncomeExpenseType As Recordset

Public RsMst_SecLocMapping              As Recordset
Public RsMst_SecLocMappingDetails    As Recordset
Public RsMst_SecurityDesignation    As Recordset



Public RsTxn_MonthwiseSecLoc       As Recordset
Public RsTxn_MonthwiseSecLocDetail As Recordset


Public RsTxn_SecurityBilling            As Recordset
Public RsTxn_SecurityBillingDetail      As Recordset

Public RsMst_StorageCardDepoRateDetail As Recordset
Public RsMst_SpecialCardDepoRateDetail As Recordset

Public RsMst_Region As Recordset

Public TmpRsCMPro As Recordset
Public RsMst_StateDRCMapping As Recordset

Public RsTxn_ReservationInvoice As Recordset

Public RsTxn_TempInvoice As Recordset
Public RsTxn_TIDetails As Recordset


Public RsTxn_NewCashReceipt As Recordset
Public RsTxn_NewCashReceiptClientDetail As Recordset
Public RsTxn_NewCashReceiptTIDetail As Recordset
Public RsTxn_NewCRDepositDetails As Recordset

Public RsMst_StateIndia As Recordset
Public RsTxn_ClientLedger As Recordset

Public RsTxn_CRTDSAdjustmentDetail As Recordset
Public RsTxn_CRTDSAdjustment As Recordset
Public RsTxn_CRAdvanceAllocation As Recordset

Public RsTxn_LockSealHologram As Recordset
Public RsTxn_LockSealHologramDetail As Recordset

Public RsLog_CxTF_Details As Recordset
Public RsLog_WHR_Detail As Recordset

Public RsLog_WHRPrint As Recordset

Public RsTxn_CDTFBOE As Recordset
Public RsTxn_CDTFBOE_GSL As Recordset

Public RsTxn_NAVInvoiceOutStanding As Recordset

Public RsLog_FormAccessDetail As Recordset

Public RsTxn_PledgeCashReceipt As Recordset
Public RsTxn_PledgeCashReceiptDetail As Recordset

Public RsTxn_FranchiseeRent As Recordset

Public RsMst_NCDEXAccreditedCapacity As Recordset

Public RsMst_EarthquakeZone As Recordset


Public RsLog_FranchiseeFeeRejectHistory As Recordset

Public RsMst_ChequeReturnReason As Recordset

Public RsLog_Client2ClientTransfer As Recordset
Public RsMst_EmployeeAllocationDetails As Recordset


Public RsTxn_FumigationApproval As Recordset
Public RsLog_ISINCommodity As Recordset

Public RsMst_SecurityRequest As Recordset
Public RsMst_SecurityRequestGodownMapping As Recordset
Public RsTxn_SecurityAttendance As Recordset


Public RsLog_TICollectionRemark As Recordset


Public RsTxn_LessorAgreeAttachment As Recordset
Public TmpRsSM_Attachment As Recordset

Public RsMst_FF_RateCard_Godown_Fix As Recordset

Public RsTxn_SContractAttachment As Recordset
Public RsMst_SContractBillingDetail As Recordset

Public RsTxn_RRNPWR As Recordset
Public RsTxn_RRNPWRGSL As Recordset

Public RsMst_LicenseAttachment As Recordset


Public RsTxn_ReservationDailyStock As Recordset
Public RsTxn_ReservationInvoiceDetail As Recordset

Public RsMst_InspectionExpenseType As Recordset
Public RsMst_InspectionExpenseDetail As Recordset
Public RsMst_InspectionCommodityDetail As Recordset
Public RsMst_InspectionGodownDetail As Recordset
Public RsMst_GodownInspection As Recordset

Public RsMst_TreatmentType As Recordset
Public RsMst_FumigationAgreement As Recordset
Public RsMst_FumigationTreatmentDetail As Recordset
Public RsMst_FumigationQuotationDetail As Recordset

Public RsTxn_TRF As Recordset
Public RsTxn_TRFDetail As Recordset

Public RsExpiredGSL_Contract As Recordset

Public RsTxn_MonthlyCommtrackFileDetail As Recordset
Public RsTxn_MonthlyCommtrackInvoice As Recordset

Public RsTxn_TempInvSRGodownDetails As Recordset

Public RsMst_RRNLotDetail As Recordset

Public RsLog_InvoiceDeletetion As Recordset

Public RsTxn_InsuranceExceptionMapping As Recordset

Public RsTxn_SecurityAgreement As Recordset
Public RsTxn_SecurityAgreementDetail As Recordset

Public RsTxn_TempInvGodownDetails As Recordset

Public RsTxn_MonthlyInsuranceCharges As Recordset
Public RsTxn_MonthlyInsuranceChargesDetail As Recordset

Public RsTxn_GodownRentST As Recordset
Public RsTxn_GodownRentSTClaimDetail As Recordset
Public RsTxn_GodownRentSTAgreeDetail As Recordset


Public RsTxn_SecurityAgreeAttachment As Recordset
Public RsMst_FSAgreeAttachment As Recordset

Public RsTxn_ScmQualityParameter As Recordset
Public RsTxn_ScmQualityParameterDetail As Recordset

Public RsTxn_QQSGodownMapping As Recordset
Public RsTxn_QQSGodownMappingDetail As Recordset

Public RsMst_DocumentType As Recordset
Public RsTxn_GodownAgreeDocCycle As Recordset
Public RsTxn_GADCGodownDetail As Recordset
Public RsTxn_GADCDocumentAttachDetail As Recordset
Public RsLog_GADCForwardToHistory As Recordset

Public RsMst_LicenseType As Recordset
Public RsTxn_LicenseAllocationDetail As Recordset
Public RsTxn_LicenseGodownDetail As Recordset

Public RsTxn_SCMDO As Recordset

Public RsMst_YearlyStorageRateCard As Recordset
Public RsMst_YearlyStorageRateCardDetail As Recordset

Public RsMst_VehicleType As Recordset
Public RsMst_InspectionAreaDetail As Recordset

Public RsTxn_ImportNcdexLotPleadge As Recordset

Public RsMst_MAW As Recordset
Public RsMst_MAWDetails As Recordset

Public RsMst_Stack As Recordset

Public RsLog_UpdateCMPBlockDays As Recordset

Public RsTxn_ClientFeedbackDoc As Recordset             ''-- Client Feedback Document
Public RsTxn_ClientFeedbackDocAttachment As Recordset           ''-- Client Feedback Attachment
Public RsTxn_MonthlyCommtrackFileDetailCMSP As Recordset

Public RsTxn_InvoicesForNAVISION As Recordset
Public RsTxn_CashRepeiptsForNAVISION  As Recordset
Public RsTxn_InvoiceAllocationForNAVISION As Recordset
Public RsTxn_SecurityBillForNavision As Recordset
Public RsTxn_IncomeExpenseForDailyPnL As Recordset

Public RsMst_ClientStateWiseGST              As Recordset
Public RsMst_DepositoryStateWiseGST          As Recordset
Public RsMst_LessorStateWiseGST              As Recordset
Public RsMst_SecurityAgencyStateWiseGST      As Recordset
Public RsMst_VendorStateWiseGST              As Recordset
Public RsMst_GLAccountIncomeTypeMapping      As Recordset

Public RsMst_InsuranceBy                    As Recordset
Public RsTxn_DepositContract                As Recordset
Public RsTxn_DepositContractGodownDetail    As Recordset
Public RsTxn_DepositContractAttachment      As Recordset

Public RsMst_InsuranceRiskCovered As Recordset
Public RsMst_InsuranceForUnitTypeDetail As Recordset
Public RsMst_InsuranceValidFor As Recordset
Public RsMst_InsuranceGodownMapping  As Recordset
Public RsMst_InsuranceCommodityDetail As Recordset
Public RsMst_InsuranceClientDetail As Recordset

Public RsTxn_InsurancePremium As Recordset              ''-- Ullhas Patil 13th June 2018
Public RsTxn_InsurancePremiumDetail As Recordset        ''-- Ullhas Patil 13th June 2018
Public RsLog_Txn_InsurancePremiumDetail As Recordset    ''-- Ullhas Patil 13th June 2018

Public RsTxn_GSL As Recordset
Public TmpRsAzureStoreMan As Recordset
Public TmpRs1AzureStoreMan As Recordset

Public RsMst_ClientAttachment As Recordset

Public RsTxn_WHRToNFin As Recordset

Public Function TmpTable()

Call DBC
sql = tmp
Set TmpRs = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs.CursorLocation = adUseClient
End If

TmpRs.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Public Function Tmp0Table()
Call DBC
sql = tmp
Set TmpRs0 = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs0.CursorLocation = adUseClient
End If

TmpRs0.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Public Function Tmp1Table()
Call DBC
sql = tmp
Set TmpRs1 = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs1.CursorLocation = adUseClient
End If

TmpRs1.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function
'FIXIT: Declare 'Tmp2Table' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
Public Function Tmp2Table()
Call DBC
sql = tmp
Set TmpRs2 = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs2.CursorLocation = adUseClient
End If

TmpRs2.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

'FIXIT: Declare 'Tmp2Table' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
Public Function Tmp3Table()
Call DBC
sql = tmp
Set TmpRs3 = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs3.CursorLocation = adUseClient
End If

TmpRs3.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Public Function Tmp4Table()
Call DBC
sql = tmp
Set TmpRs4 = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs4.CursorLocation = adUseClient
End If

TmpRs4.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function Tmp5Table()
Call DBC
sql = tmp
Set TmpRs5 = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs5.CursorLocation = adUseClient
End If

TmpRs5.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function Tmp6Table()
Call DBC
sql = tmp
Set TmpRs6 = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs6.CursorLocation = adUseClient
End If
TmpRs6.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function Tmp7Table()
Call DBC
sql = tmp
Set TmpRs7 = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs7.CursorLocation = adUseClient
End If
TmpRs7.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableMst_Config(Optional ByVal wc As String = "")
sql = "Select * from Mst_Config " & wc   '  & tmp & "'"
Set RsMst_Config = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Config.CursorLocation = adUseClient
End If

RsMst_Config.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

 ''===== Mst_AssayerMaster Table Declaration ====''
Function TableMst_Assayer(Optional ByVal wc As String = "")
sql = "Select * from Mst_Assayer " & wc
Set RsMst_Assayer = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Assayer.CursorLocation = adUseClient
End If

RsMst_Assayer.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_CategoryMaster Table Declaration ====''
Function TableMst_Category(Optional ByVal wc As String = "")
sql = "Select * from Mst_Category " & wc
Set RsMst_Category = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Category.CursorLocation = adUseClient
End If

RsMst_Category.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_ClientMaster Table Declaration ====''
Function TableMst_Client(Optional ByVal wc As String = "")
sql = "Select * from Mst_Client " & wc
Set RsMst_Client = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Client.CursorLocation = adUseClient
End If
RsMst_Client.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_CMPMaster Table Declaration ====''
Function TableMst_CMP(Optional ByVal wc As String = "")
sql = "Select * from Mst_CMP " & wc
Set RsMst_CMP = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_CMP.CursorLocation = adUseClient
End If
RsMst_CMP.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_CommodityMaster Table Declaration ====''
Function TableMst_Commodity(Optional ByVal wc As String = "")
sql = "Select * from Mst_Commodity " & wc
Set RsMst_Commodity = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Commodity.CursorLocation = adUseClient
End If
RsMst_Commodity.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''
 ''===== Mst_CommodityGroupMaster Table Declaration ====''
Function TableMst_CommodityGroup(Optional ByVal wc As String = "")
sql = "Select * from Mst_CommodityGroup " & wc
Set RsMst_CommodityGroup = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_CommodityGroup.CursorLocation = adUseClient
End If
RsMst_CommodityGroup.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_CommodityMaster Table Declaration ====''
Function TableMst_CommodityDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_CommodityDetail " & wc
Set RsMst_CommodityDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_CommodityDetail.CursorLocation = adUseClient
End If
RsMst_CommodityDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''


 ''===== Mst_DepositoryMaster Table Declaration ====''
Function TableMst_Depository(Optional ByVal wc As String = "")
sql = "Select * from Mst_Depository " & wc
Set RsMst_Depository = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Depository.CursorLocation = adUseClient
End If
RsMst_Depository.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_ExchangeTypeMaster Table Declaration ====''
Function TableMst_ExchangeType(Optional ByVal wc As String = "")
sql = "Select * from Mst_ExchangeType " & wc
Set RsMst_ExchangeType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_ExchangeType.CursorLocation = adUseClient
End If
RsMst_ExchangeType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_GodownMaster Table Declaration ====''
Function TableMst_Godown(Optional ByVal wc As String = "")
sql = "Select * from Mst_Godown " & wc
Set RsMst_Godown = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Godown.CursorLocation = adUseClient
End If
RsMst_Godown.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_GSLMaster Table Declaration ====''
Function TableMst_GSL(Optional ByVal wc As String = "", Optional ByVal Df As String = "")
If Df = "D" Then
   sql = "Select distinct StackNo from Mst_GSL " & wc
Else
   sql = "Select * from Mst_GSL " & wc
End If
Set RsMst_GSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_GSL.CursorLocation = adUseClient
End If
RsMst_GSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_ISINMaster Table Declaration ====''
Function TableMst_ISIN(Optional ByVal wc As String = "")
sql = "Select * from Mst_ISIN " & wc
Set RsMst_ISIN = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_ISIN.CursorLocation = adUseClient
End If
RsMst_ISIN.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_LocationMaster Table Declaration ====''
Function TableMst_Location(Optional ByVal wc As String = "")
sql = "Select * from Mst_Location " & wc
Set RsMst_Location = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Location.CursorLocation = adUseClient
End If
RsMst_Location.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_MenuDetailMaster Table Declaration ====''
Function TableMst_MenuDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_MenuDetail " & wc
Set RsMst_MenuDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_MenuDetail.CursorLocation = adUseClient
End If
RsMst_MenuDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_MenuHasUserRoleMaster Table Declaration ====''
Function TableMst_MenuHasUserRole(Optional ByVal wc As String = "")
sql = "Select * from Mst_MenuHasUserRole " & wc
Set RsMst_MenuHasUserRole = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_MenuHasUserRole.CursorLocation = adUseClient
End If
RsMst_MenuHasUserRole.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_RRNMaster Table Declaration ====''
Function TableMst_RRN(Optional ByVal wc As String = "")
sql = "Select * from Mst_RRN " & wc
Set RsMst_RRN = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_RRN.CursorLocation = adUseClient
End If
RsMst_RRN.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_StateMaster Table Declaration ====''
Function TableMst_State(Optional ByVal wc As String = "")
sql = "Select * from Mst_State " & wc
Set RsMst_State = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_State.CursorLocation = adUseClient
End If
RsMst_State.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_StorageContractMaster Table Declaration ====''
Function TableMst_StorageContract(Optional ByVal wc As String = "")
sql = "Select * from Mst_StorageContract " & wc
Set RsMst_StorageContract = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_StorageContract.CursorLocation = adUseClient
End If
RsMst_StorageContract.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_StoreManLocationsMaster Table Declaration ====''
Function TableMst_StoreManLocations(Optional ByVal wc As String = "")
sql = "Select * from Mst_StoreManLocations " & wc
Set RsMst_StoreManLocations = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_StoreManLocations.CursorLocation = adUseClient
End If
RsMst_StoreManLocations.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_UnitTypeMaster Table Declaration ====''
Function TableMst_UnitType(Optional ByVal wc As String = "")
sql = "Select * from Mst_UnitType " & wc
Set RsMst_UnitType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_UnitType.CursorLocation = adUseClient
End If
RsMst_UnitType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_UserRoleMaster Table Declaration ====''
Function TableMst_UserRole(Optional ByVal wc As String = "")
sql = "Select * from Mst_UserRole " & wc
Set RsMst_UserRole = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_UserRole.CursorLocation = adUseClient
End If
RsMst_UserRole.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Mst_UsersMaster Table Declaration ====''
Function TableMst_Users(Optional ByVal wc As String = "")
sql = "Select * from Mst_Users " & wc
Set RsMst_Users = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Users.CursorLocation = adUseServer
End If
RsMst_Users.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Txn_CxTF_ClientDetailMaster Table Declaration ====''
Function TableTxn_CxTF_ClientDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_CxTF_ClientDetail " & wc
Set RsTxn_CxTF_ClientDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CxTF_ClientDetail.CursorLocation = adUseClient
End If
RsTxn_CxTF_ClientDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Txn_CxTF_DetailsMaster Table Declaration ====''
Function TableTxn_CxTF_Details(Optional ByVal wc As String = "")
sql = "Select * from Txn_CxTF_Details " & wc
Set RsTxn_CxTF_Details = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CxTF_Details.CursorLocation = adUseClient
End If
RsTxn_CxTF_Details.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Txn_CxTF_GSLMaster Table Declaration ====''
Function TableTxn_CxTF_GSL(Optional ByVal wc As String = "")
sql = "Select * from Txn_CxTF_GSL " & wc
Set RsTxn_CxTF_GSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CxTF_GSL.CursorLocation = adUseClient
End If

RsTxn_CxTF_GSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''

 ''===== Txn_CxTF_TruckDetailMaster Table Declaration ====''
Function TableTxn_CxTF_TruckDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_CxTF_TruckDetail " & wc
Set RsTxn_CxTF_TruckDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CxTF_TruckDetail.CursorLocation = adUseClient
End If
RsTxn_CxTF_TruckDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''


 ''===== Txn_CxTF_TruckDetailMaster Table Declaration ====''
Function TableMst_License(Optional ByVal wc As String = "")
sql = "Select * from Mst_License " & wc
Set RsMst_License = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_License.CursorLocation = adUseClient
End If
RsMst_License.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''===== End Declaration ======''''


Function TableMst_InsuranceCo(Optional ByVal wc As String = "")
sql = "Select * from Mst_InsuranceCoMaster " & wc
Set RsMst_InsuranceCo = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InsuranceCo.CursorLocation = adUseClient
End If
RsMst_InsuranceCo.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_PolicyDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_InsuranceDetails " & wc
Set RsMst_PolicyDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_PolicyDetail.CursorLocation = adUseClient
End If
RsMst_PolicyDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableGSL_ISIN(Optional ByVal wc As String = "", Optional ByVal OrdBy As String = "")
'sql = "SELECT DISTINCT Mst_GSL.ClientIDRow, Mst_GSL.ISINID, Mst_GSL.Flag From Mst_GSL " & _
'      "WHERE (((Mst_GSL.ExchangeTypeID)=2 )) " & wc
'
sql = "SELECT Mst_GSL.ClientIDRow, Mst_GSL.ISINID, Mst_GSL.Flag, Max(Mst_GSL.CMPID) AS CMPID, " & _
"Max(Mst_GSL.CommodityID) As CommodityID From Mst_GSL " & _
"GROUP BY Mst_GSL.ClientIDRow, Mst_GSL.ISINID, Mst_GSL.Flag, Mst_GSL.ExchangeTypeID " & _
"HAVING ( " & wc & "  AND ((Mst_GSL.ExchangeTypeID)=2) And (Mst_GSL.Flag = 'J' OR Mst_GSL.Flag = 'O' )) " & OrdBy

'"ORDER BY Mst_GSL.ClientIDRow, Mst_GSL.ISINID;"

 Set RsGSL_ISIN = New ADODB.Recordset
 Call DBC
 If Gen_DBType = "SQL" Then
    RsGSL_ISIN.CursorLocation = adUseClient
 End If
 RsGSL_ISIN.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_CDC(Optional ByVal wc As String = "")
sql = "Select * from Txn_CDC " & wc
Set RsTxn_CDC = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CDC.CursorLocation = adUseClient
End If
RsTxn_CDC.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_CDC_GSLDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_CDC_GSL_Detail " & wc
Set RsTxn_CDC_GSLDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CDC_GSLDetail.CursorLocation = adUseClient
End If
RsTxn_CDC_GSLDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_CDC_AssayerDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_CDC_Assayer_Detail " & wc
Set RsTxn_CDC_AssayerDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CDC_AssayerDetail.CursorLocation = adUseClient
End If
RsTxn_CDC_AssayerDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_CDC_InsuranceDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_CDC_Insurance_Detail " & wc
Set RsTxn_CDC_InsuranceDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CDC_InsuranceDetail.CursorLocation = adUseClient
End If
RsTxn_CDC_InsuranceDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_ISINRevalidation(Optional ByVal wc As String = "")
sql = "Select * from Txn_ISINRevalidation " & wc
Set RsTxn_ISINRevalidation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ISINRevalidation.CursorLocation = adUseClient
End If
RsTxn_ISINRevalidation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_ISINGSLRevalidation(Optional ByVal wc As String = "")
sql = "Select * from Txn_ISINGSLRevalidation " & wc
Set RsTxn_ISINGSLRevalidation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ISINGSLRevalidation.CursorLocation = adUseClient
End If
RsTxn_ISINGSLRevalidation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_WHRDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_WHR_Detail " & wc
Set RsTxn_WHRDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_WHRDetail.CursorLocation = adUseClient
End If
RsTxn_WHRDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

'Function TableTxn_WHRLicenseDetail(Optional ByVal wc As String = "")
' sql = "Select * from Txn_WHRLicense_Detail " & wc
' Set RsTxn_WHRLicenseDetail = New ADODB.Recordset
' Call DBC
' RsTxn_WHRLicenseDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
'End Function


Function TableMst_Bank(Optional wc As String)
sql = "select * from Mst_Bank " & wc
Set RsMst_Bank = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Bank.CursorLocation = adUseClient
End If
RsMst_Bank.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_Branch(Optional wc As String)
sql = "select * from Mst_Branch " & wc
Set RsMst_Branch = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Branch.CursorLocation = adUseClient
End If
RsMst_Branch.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_Spillage(Optional wc As String)
sql = "select * from Txn_Spillage " & wc
Set RsTxn_Spillage = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_Spillage.CursorLocation = adUseClient
End If
RsTxn_Spillage.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_PledgeCreation(Optional wc As String)
sql = "select * from Txn_PledgeCreation " & wc
Set RsTxn_PledgeCreation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_PledgeCreation.CursorLocation = adUseClient
End If
RsTxn_PledgeCreation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_PledgeCreation_GSL(Optional wc As String)
sql = "select * from Txn_PledgeCreation_GSL " & wc
Set RsTxn_PledgeCreation_GSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_PledgeCreation_GSL.CursorLocation = adUseClient
End If
RsTxn_PledgeCreation_GSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_PledgeRelease(Optional wc As String)
sql = "select * from Txn_PledgeRelease " & wc
Set RsTxn_PledgeRelease = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_PledgeRelease.CursorLocation = adUseClient
End If
RsTxn_PledgeRelease.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_PledgeRelease_GSL(Optional wc As String)
sql = "select * from Txn_PledgeRelease_GSL " & wc
Set RsTxn_PledgeRelease_GSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_PledgeRelease_GSL.CursorLocation = adUseClient
End If
RsTxn_PledgeRelease_GSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CashReceipt(Optional wc As String)
sql = "select * from Txn_CashReceipt " & wc
Set RsTxn_CashReceipt = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CashReceipt.CursorLocation = adUseClient
End If
RsTxn_CashReceipt.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CashReceiptInvoiceDetails(Optional ByVal wc As String = "")
sql = "Select * from Txn_CashReceiptInvoiceDetails " & wc
Set RsTxn_CashReceiptInvoiceDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CashReceiptInvoiceDetails.CursorLocation = adUseClient
End If
RsTxn_CashReceiptInvoiceDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function Mst_MenuDetailTable(Optional wc As String)
sql = "Select * from Mst_MenuDetail " & wc ' & Wc & tmp & "'"
Set RsMenuDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMenuDetail.CursorLocation = adUseClient
End If
RsMenuDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function Mst_MenuHasUserRoleTable(Optional wc As String)
sql = "Select * from Mst_MenuHasUserRole " & wc ' & Wc & tmp & "'"
Set RsMenuUserRole = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMenuUserRole.CursorLocation = adUseClient
End If
RsMenuUserRole.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableMst_Hologram(Optional wc As String)
sql = "select * from Mst_Hologram " & wc
Set RsMst_Hologram = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Hologram.CursorLocation = adUseClient
End If
RsMst_Hologram.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_PrintHistory(Optional wc As String)
sql = "Select * from Log_Print_History " & wc
Set RsLog_PrintHistory = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_PrintHistory.CursorLocation = adUseClient
End If
RsLog_PrintHistory.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_GSL_OpeningBalance(Optional wc As String)
sql = "Select * from Txn_GSL_OpeningBalance " & wc
Set RsTxn_GSL_OpeningBalance = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GSL_OpeningBalance.CursorLocation = adUseClient
End If
RsTxn_GSL_OpeningBalance.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableMst_Agent(Optional ByVal wc As String = "")
sql = "Select * from Mst_Agent " & wc
Set RsMst_Agent = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Agent.CursorLocation = adUseClient
End If
RsMst_Agent.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_GSL_Transfer(Optional wc As String)
sql = "Select * from Txn_GSL_Transfer " & wc
Set RsTxn_GSL_Transfer = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GSL_Transfer.CursorLocation = adUseClient
End If
RsTxn_GSL_Transfer.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_DR(Optional ByVal wc As String = "")
sql = "Select * from Mst_DR " & wc
Set RsMst_DR = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_DR.CursorLocation = adUseClient
End If
RsMst_DR.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Public Function TableTxn_BenposTemp(Optional wc As String)
sql = "Select * from ##BenposTemp " & wc
Set RsTxn_Benpos = New ADODB.Recordset
Call DBC
RsTxn_Benpos.CursorLocation = adUseClient
RsTxn_Benpos.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TableTxn_CreateBenposTemp()
Dim sql As String
sql = "CREATE TABLE [dbo].[##BenposTemp] ([ISINID] [varchar] (36) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , " & _
"[ISINDescription] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , " & _
"[DPID] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , " & _
"[ClientID] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , " & _
"[ClientName] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , " & _
"[Depository] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , " & _
"[ElectronicPosition] [decimal](18, 4) NULL , " & _
"[Unit] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
"[ImportDate] Datetime ) ON [PRIMARY] "

Set RsTxn_BenposTemp = New ADODB.Recordset
Call DBC
RsTxn_BenposTemp.CursorLocation = adUseClient
RsTxn_BenposTemp.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Public Function TableTxn_BenposDatewise(Optional wc As String)
sql = "Select * from Txn_BenposDateWise " & wc
Set RsTxn_BenposDateWise = New ADODB.Recordset
Call DBC
RsTxn_BenposDateWise.CursorLocation = adUseClient
RsTxn_BenposDateWise.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TableTxn_InsertBenposDateWise(dtImport As Date)
Dim sql As String
'sql = "INSERT INTO BenposDateWise (ISIN,Qty,ImportDate) (SELECT ISIN,SUM(ElectronicPosition) 'Qty',ImportDate FROM ##BenposTemp GROUP BY ISIN,ImportDate )"
sql = "INSERT INTO Txn_BenposDateWise (ISINID,Qty,ImportDate) (SELECT ISINID,CASE WHEN Unit='KGS' THEN SUM(ElectronicPosition)/1000 ELSE SUM(ElectronicPosition) END 'Qty',ImportDate FROM ##BenposTemp GROUP BY ISINID,ImportDate,Unit )"
Set RsTxn_InsertBenposDateWise = New ADODB.Recordset
Call DBC
RsTxn_InsertBenposDateWise.CursorLocation = adUseClient
RsTxn_InsertBenposDateWise.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_AdjTransaction(Optional wc As String)
sql = "Select * from Txn_AdjTransaction " & wc
Set RsTxn_AdjTransaction = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_AdjTransaction.CursorLocation = adUseClient
End If
RsTxn_AdjTransaction.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_GSLAdjTransaction(Optional wc As String)
sql = "Select * from Txn_GSLAdjTransaction " & wc
Set RsTxn_GSLAdjTransaction = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GSLAdjTransaction.CursorLocation = adUseClient
End If
RsTxn_GSLAdjTransaction.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function Help_DetailTable(Optional wc As String)
sql = "select * from Help_Detail " & wc
Set RsHelp_Detail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsHelp_Detail.CursorLocation = adUseClient
End If
RsHelp_Detail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_Designation(Optional wc As String)
sql = "select * from Mst_Designation " & wc
Set RsMst_Designation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Designation.CursorLocation = adUseClient
End If
RsMst_Designation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_AssetType(Optional wc As String)
sql = "select * from Mst_AssetsType " & wc
Set RsMst_AssetType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_AssetType.CursorLocation = adUseClient
End If
RsMst_AssetType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_AssetStatus(Optional wc As String)
sql = "select * from Mst_AssetStatus " & wc
Set RsMst_AssetStatus = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_AssetStatus.CursorLocation = adUseClient
End If
RsMst_AssetStatus.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_Employee(Optional wc As String)
sql = "select * from Mst_Employee " & wc
Set RsMst_Employee = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Employee.CursorLocation = adUseClient
End If
RsMst_Employee.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_EmpCMPMap(Optional wc As String)
sql = "Select * from Mst_EmpCMPMapping " & wc
Set RsMst_EmpCMPMap = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_EmpCMPMap.CursorLocation = adUseClient
End If
RsMst_EmpCMPMap.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Function TableMst_EmpAssetMap(Optional wc As String)
sql = "Select * from Mst_EmpAssetMap " & wc
Set RsMst_EmpAssetMap = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_EmpAssetMap.CursorLocation = adUseClient
End If
RsMst_EmpAssetMap.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Function TableMst_SecurityShifts(Optional wc As String)
sql = "Select * from Mst_SecurityShifts " & wc
Set RsMst_SecurityShifts = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecurityShifts.CursorLocation = adUseClient
End If
RsMst_SecurityShifts.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_SecurityAgency(Optional wc As String)
sql = "Select * from Mst_SecurityAgency " & wc
Set RsMst_SecurityAgency = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecurityAgency.CursorLocation = adUseClient
End If

RsMst_SecurityAgency.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_SecurityDetail(Optional wc As String)
sql = "select * from Mst_SecurityDetails " & wc
Set RsMst_SecurityDetails = New ADODB.Recordset
Call DBC

If Gen_DBType = "SQL" Then
   RsMst_SecurityDetails.CursorLocation = adUseClient
End If

RsMst_SecurityDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_SecurityCMPMap(Optional wc As String)
sql = "select * from Mst_SecurityCMPMap " & wc
Set RsMst_SecurityCMPMap = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecurityCMPMap.CursorLocation = adUseClient
End If

RsMst_SecurityCMPMap.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_SecurityAssets(Optional wc As String)
sql = "select * from Mst_SecurityAssets " & wc
Set RsMst_SecurityAssets = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecurityAssets.CursorLocation = adUseClient
End If

RsMst_SecurityAssets.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_SecurityAssetMap(Optional wc As String)
sql = "select * from Mst_SecurityAssetMap " & wc
Set RsMst_SecurityAssetMap = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecurityAssetMap.CursorLocation = adUseClient
End If

RsMst_SecurityAssetMap.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function
Function TableMst_ExpenseType(Optional wc As String)
sql = "select * from Mst_ExpenseType " & wc
Set RsMst_ExpenseType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_ExpenseType.CursorLocation = adUseClient
End If
RsMst_ExpenseType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_ClaimSubmission(Optional wc As String)
sql = "Select * from Txn_ClaimSubmission " & wc
Set RsTxn_ClaimSubmission = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ClaimSubmission.CursorLocation = adUseClient
End If
RsTxn_ClaimSubmission.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_ClaimSubmissionDetails(Optional wc As String)
sql = "Select * from Txn_ClaimSubmissionDetail " & wc
Set RsTxn_ClaimSubmissionDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ClaimSubmissionDetails.CursorLocation = adUseClient
End If
RsTxn_ClaimSubmissionDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_AuditSegment(Optional wc As String)
sql = "Select * from Mst_AuditSegment " & wc
Set RsMst_AuditSegment = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_AuditSegment.CursorLocation = adUseClient
End If
RsMst_AuditSegment.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_AuditType(Optional wc As String)
sql = "Select * from Mst_AuditType " & wc
Set RsMst_AuditType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_AuditType.CursorLocation = adUseClient
End If
RsMst_AuditType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_AuditPoint(Optional wc As String)
sql = "Select * from Mst_AuditPoint " & wc
Set RsMst_AuditPoint = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_AuditPoint.CursorLocation = adUseClient
End If
RsMst_AuditPoint.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_Audit(Optional wc As String)
sql = "Select * from Txn_Audit " & wc
Set RsTxn_Audit = New ADODB.Recordset
RsTxn_Audit.CursorLocation = adUseClient
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_Audit.CursorLocation = adUseClient
End If
RsTxn_Audit.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_AuditDetails(Optional wc As String)
sql = "Select * from Txn_AuditDetails " & wc
Set RsTxn_AuditDetails = New ADODB.Recordset
RsTxn_AuditDetails.CursorLocation = adUseClient
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_AuditDetails.CursorLocation = adUseClient
End If
RsTxn_AuditDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_EmpCTC(Optional ByVal wc As String = "")
sql = "Select * from Txn_EmpCTC " & wc   '  & tmp & "'"
Set RsTxn_EmpCTC = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_EmpCTC.CursorLocation = adUseClient
End If

RsTxn_EmpCTC.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Function TableTxn_IncomeDetails(Optional ByVal wc As String = "")
sql = "Select * from Txn_IncomeDetails " & wc
Set RsTxn_IncomeDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_IncomeDetails.CursorLocation = adUseClient
End If

RsTxn_IncomeDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_DPIncomeDetails(Optional ByVal wc As String = "")
sql = "Select * from Txn_DPIncomeDetails " & wc   '  & tmp & "'"
Set RsTxn_DPIncomeDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_DPIncomeDetails.CursorLocation = adUseClient
End If

RsTxn_DPIncomeDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function



Function TableTxn_TelephoneDetails(Optional ByVal wc As String = "")
sql = "Select * from Txn_TelephoneDetails " & wc
Set RsTxn_TelephoneDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TelephoneDetail.CursorLocation = adUseClient
End If
RsTxn_TelephoneDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_TelephoneTemp(Optional ByVal wc As String = "")
 
sql = " Select * from Txn_TelephoneTemp " & wc
Set RsTxn_TelephoneTemp = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TelephoneTemp.CursorLocation = adUseClient
End If

RsTxn_TelephoneTemp.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function



Function TableMst_ExpenseGroup(Optional wc As String)
sql = "Select * from Mst_ExpenseGroup " & wc
Set RsMst_ExpenseGroup = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_ExpenseGroup.CursorLocation = adUseClient
End If
RsMst_ExpenseGroup.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_Vendor(Optional wc As String)
sql = "Select * from Mst_Vendor " & wc
Set RsMst_Vendor = New ADODB.Recordset
RsMst_Vendor.CursorLocation = adUseClient
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Vendor.CursorLocation = adUseClient
End If
RsMst_Vendor.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function





Function TableMst_ServiceTax(Optional wc As String)
sql = "Select * from Mst_ServiceTax " & wc
Set RsMst_ServiceTax = New ADODB.Recordset
RsMst_ServiceTax.CursorLocation = adUseClient
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_ServiceTax.CursorLocation = adUseClient
End If
RsMst_ServiceTax.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function



Function TableMst_TDS(Optional wc As String)
sql = "Select * from Mst_TDS " & wc
Set RsMst_TDS = New ADODB.Recordset
RsMst_TDS.CursorLocation = adUseClient
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_TDS.CursorLocation = adUseClient
End If
RsMst_TDS.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TempDataTableDelete(Optional wc As String)
sql = " Delete  From Temp " & wc
Call DBC
Set rstemp = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   rstemp.CursorLocation = adUseClient
End If

rstemp.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Public Function TempDataTable(Optional wc As String)
sql = " Select * From Temp " & wc
Call DBC
Set rstemp = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   rstemp.CursorLocation = adUseClient
End If

rstemp.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Public Function TableTmp_DPIncomeDetails(Optional wc As String)
sql = " Select * From Tmp_DPIncomeDetails " & wc
Call DBC
Set RsTmp_DPIncomeDetails = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsTmp_DPIncomeDetails.CursorLocation = adUseClient
End If

RsTmp_DPIncomeDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_VendorClaim(Optional wc As String)
sql = "Select * from Txn_VendorClaim " & wc
Set RsTxn_VendorClaim = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_VendorClaim.CursorLocation = adUseClient
End If
RsTxn_VendorClaim.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_VendorClaimDetails(Optional wc As String)
sql = "Select * from Txn_VendorClaimDetail " & wc
Set RsTxn_VendorClaimDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_VendorClaimDetails.CursorLocation = adUseClient
End If
RsTxn_VendorClaimDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TableMst_StockStatus(Optional wc As String)
sql = " Select * From Mst_StockStatus " & wc
Call DBC
Set RsMst_StockStatus = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsMst_StockStatus.CursorLocation = adUseClient
End If
RsMst_StockStatus.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Public Function TableMst_ClientLogin(Optional wc As String)
sql = " Select * From Mst_ClientLogin " & wc
Call DBC
Set RsMst_ClientLogin = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsMst_ClientLogin.CursorLocation = adUseClient
End If
RsMst_ClientLogin.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Public Function TableMst_ClientLoginDetail(Optional wc As String)
sql = " Select * From Mst_ClientLoginDetail " & wc
Call DBC
Set RsMst_ClientLoginDetail = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsMst_ClientLoginDetail.CursorLocation = adUseClient
End If
RsMst_ClientLoginDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TableMst_TmpStockStatus(Optional wc As String)
sql = " Select * From Mst_TmpStockStatus " & wc
Call DBC
Set RsMst_TmpStockStatus = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsMst_TmpStockStatus.CursorLocation = adUseClient
End If
RsMst_TmpStockStatus.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableLog_CMPReopenHistory(Optional wc As String)
sql = "Select * from Log_CMPReopenHistory " & wc
Set RsLog_CMPReopenHistory = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_CMPReopenHistory.CursorLocation = adUseClient
End If
RsLog_CMPReopenHistory.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableMst_CommodityPrice(Optional ByVal wc As String = "")
sql = "Select * from Mst_CommodityPrice " & wc
Set RsMst_CommodityPrice = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_CommodityPrice.CursorLocation = adUseClient
End If

RsMst_CommodityPrice.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_SRGodownMapping(Optional wc As String)
sql = "Select * from Txn_SRGodownMapping " & wc
Set RsTxn_SRGodownMapping = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_SRGodownMapping.CursorLocation = adUseClient
End If
RsTxn_SRGodownMapping.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_SRGodownMapDetails(Optional wc As String)
sql = "Select * from Txn_SRGodownMapDetails " & wc
Set RsTxn_SRGodownMapDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_SRGodownMapDetails.CursorLocation = adUseClient
End If
RsTxn_SRGodownMapDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableGSL_Contract(Optional ByVal wc As String = "", Optional ByVal OrdBy As String = "")

'sql = "SELECT Mst_GSL.GSLID,Mst_GSL.SM_Prefix, Mst_GSL.ContractID, Mst_GSL.Flag, Max(Mst_GSL.CMPID) as CMPID, " & _
'" Max(Mst_GSL.CommodityID) as CommodityID From Mst_GSL " & _
'" " & Wc & " And (Mst_GSL.Flag <> 'B' And Mst_GSL.Flag <> 'E' And Mst_GSL.Flag <> 'Z' ) " & _
'" GROUP BY Mst_GSL.GSLID,Mst_GSL.SM_Prefix, Mst_GSL.ContractID, Mst_GSL.Flag, Mst_GSL.ExchangeTypeID " & OrdBy

 
sql = "SELECT Distinct Mst_GSL.NContractID "
sql = sql & " From Mst_GSL "
sql = sql & " Inner Join Txn_SRGodownMapping On Mst_GSL.NContractID = Txn_SRGodownMapping.ContractID "
sql = sql & " Inner Join Mst_StorageContract On Txn_SRGodownMapping.ContractID = Mst_StorageContract.ContractID   "
sql = sql & "  " & wc & " And (Mst_GSL.Flag = 'O' ) "

If Gen_WcLocation <> "" Then
   sql = sql & " And Mst_StorageContract." & Gen_WcLocation
End If

sql = sql & "  " & OrdBy & " "

 
Set RsGSL_Contract = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsGSL_Contract.CursorLocation = adUseClient
End If
RsGSL_Contract.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Function TableMst_BagEquivalent(Optional wc As String)
sql = "Select * from Mst_BagEquivalent " & wc
Set RsMst_BagEquivalent = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_BagEquivalent.CursorLocation = adUseClient
End If
RsMst_BagEquivalent.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_CA(Optional wc As String)
sql = "Select * from Mst_CA " & wc
Set RsMst_CA = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_CA.CursorLocation = adUseClient
End If
RsMst_CA.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_KRDelete(Optional wc As String)
sql = "Select * from Log_KRDelete " & wc
Set RsLog_KRDelete = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_KRDelete.CursorLocation = adUseClient
End If
RsLog_KRDelete.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function Log_LoginDetailTable(Optional wc As String)
sql = "select * from Log_LoginDetail " & wc
Set RsLog_LoginDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_LoginDetail.CursorLocation = adUseClient
End If

RsLog_LoginDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_CashReceiptCDTFDetails(Optional ByVal wc As String = "")
sql = "Select * from Txn_CashReceiptCDTFDetails " & wc
Set RsTxn_CashReceiptCDTFDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CashReceiptCDTFDetails.CursorLocation = adUseClient
End If
RsTxn_CashReceiptCDTFDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CreateCDTFTemp()
Dim sql As String
sql = "Create Table ##CDTFDetails " & _
    "(CDTFNO Decimal,NetAmount Decimal," & _
    "ClientIDRow Decimal )"   ',DPID Varchar(50))"
    ''''"ClientID Varchar(50),DPID Varchar(50))"

Set RsTxn_CreateCDTFTemp = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CreateCDTFTemp.CursorLocation = adUseClient
End If
RsTxn_CreateCDTFTemp.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CreatePITemp()
Dim sql As String
sql = "Create Table ##PIDetails " & _
     "( PINo Decimal,NetAmount Decimal," & _
     "ClientIDRow Decimal ) " ',DPID Varchar(50))"
     '"ClientID Varchar(50),DPID Varchar(50))"

Set RsTxn_CreatePITemp = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CreatePITemp.CursorLocation = adUseClient
End If
RsTxn_CreatePITemp.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_CDTFDetailsTemp(Optional ByVal wc As String = "")
sql = "Select * from ##CDTFDetails " & wc
Set RsTxn_CDTFTemp = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CDTFTemp.CursorLocation = adUseClient
End If
RsTxn_CDTFTemp.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_PIDetailsTemp(Optional ByVal wc As String = "")
sql = "Select * from  ##PIDetails " & wc
Set RsTxn_PITemp = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_PITemp.CursorLocation = adUseClient
End If
RsTxn_PITemp.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_BillingCycle(Optional wc As String)
sql = "Select * from Mst_BillingCycle " & wc
Set RsMst_BillingCycle = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_BillingCycle.CursorLocation = adUseClient
End If
RsMst_BillingCycle.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_StorageCardRate(Optional wc As String)
sql = "Select * from Mst_StorageRateCard " & wc
Set RsMst_StorageRateCard = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_StorageRateCard.CursorLocation = adUseClient
End If
RsMst_StorageRateCard.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_WHRUOM(Optional ByVal wc As String = "")
sql = "Select * from  Mst_WHRUOM " & wc
Set RsMst_WHRUOM = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_WHRUOM.CursorLocation = adUseClient
End If
RsMst_WHRUOM.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_UnmarkPrint_History(Optional wc As String)
sql = "Select * from Log_UnmarkPrint_History " & wc
Set RsLog_UnmarkPrint_History = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_UnmarkPrint_History.CursorLocation = adUseClient
End If
RsLog_UnmarkPrint_History.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_EmpLedger(Optional ByVal wc As String = "")
sql = "Select * from Mst_EmpLedger " & wc
Set RsMst_EmpLedger = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_EmpLedger.CursorLocation = adUseClient
End If
RsMst_EmpLedger.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_Lessor(Optional ByVal wc As String = "")
sql = "Select * from Mst_Lessor " & wc
Set RsMst_Lessor = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Lessor.CursorLocation = adUseClient
End If
RsMst_Lessor.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_LessorAgreement(Optional ByVal wc As String = "")
sql = "Select * from Txn_LessorAgreement " & wc
Set RsTxn_LessorAgreement = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_LessorAgreement.CursorLocation = adUseClient
End If
RsTxn_LessorAgreement.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_LessorAgreementDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_LessorAgreementDetail " & wc
Set RsTxn_LessorAgreementDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_LessorAgreementDetail.CursorLocation = adUseClient
End If
RsTxn_LessorAgreementDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_GodownLessorAgreeDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_GodownLessorAgreeDetail " & wc
Set RsTxn_GodownLessorAgreeDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GodownLessorAgreeDetail.CursorLocation = adUseClient
End If
RsTxn_GodownLessorAgreeDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_UpdateCxTFDetails(Optional ByVal wc As String = "")
sql = "Select * from Log_UpdateCxTFDetails" & wc
Set RsLog_UpdateCxTFDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_UpdateCxTFDetails.CursorLocation = adUseClient
End If
RsLog_UpdateCxTFDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_GodownRent(Optional ByVal wc As String = "")
sql = "Select * from Txn_GodownRent " & wc
Set RsTxn_GodownRent = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GodownRent.CursorLocation = adUseClient
End If
RsTxn_GodownRent.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_GodownRentDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_GodownRentDetail " & wc
Set RsTxn_GodownRentDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GodownRentDetail.CursorLocation = adUseClient
End If
RsTxn_GodownRentDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_DPInvoiceUpload(Optional ByVal wc As String = "")
sql = "Select * from Mst_DPInvoiceUpload " & wc
Set RsMst_DPInvoiceUpload = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_DPInvoiceUpload.CursorLocation = adUseClient
End If
RsMst_DPInvoiceUpload.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_GLAccType(Optional wc As String)
sql = "select * from Mst_GLAccType " & wc
Set RsMst_GLAccType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_GLAccType.CursorLocation = adUseClient
End If
RsMst_GLAccType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_NewISINRevalidation(Optional ByVal wc As String = "")
sql = "Select * from Txn_NewISINRevalidation " & wc
Set RsTxn_NewISINRevalidation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_NewISINRevalidation.CursorLocation = adUseClient
End If
RsTxn_NewISINRevalidation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_NewISINGSLRevalidation(Optional ByVal wc As String = "")
sql = "Select * from Txn_NewISINGSLRevalidation " & wc
Set RsTxn_NewISINGSLRevalidation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_NewISINGSLRevalidation.CursorLocation = adUseClient
End If
RsTxn_NewISINGSLRevalidation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
 
Function TableMst_StorageCardRateDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_StorageRateCardDetail " & wc
Set RsMst_StorageCardRateDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_StorageCardRateDetail.CursorLocation = adUseClient
End If
RsMst_StorageCardRateDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableMst_SpecialRateCard(Optional wc As String)
sql = "Select * from Mst_SpecialRateCard " & wc
Set RsMst_SpecialRateCard = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SpecialRateCard.CursorLocation = adUseClient
End If
RsMst_SpecialRateCard.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_SpecialRateCardDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_SpecialRateCardDetail " & wc
Set RsMst_SpecialRateCardDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SpecialRateCardDetail.CursorLocation = adUseClient
End If
RsMst_SpecialRateCardDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_InvMonGSLDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_InvMonGSLDetail " & wc
Set RsTxn_InvMonGSLDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_InvMonGSLDetail.CursorLocation = adUseClient
End If
RsTxn_InvMonGSLDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_ProInvoice(Optional ByVal wc As String = "")
sql = "Select * from Txn_ProInvoice " & wc
Set RsTxn_ProInvoice = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ProInvoice.CursorLocation = adUseClient
End If

RsTxn_ProInvoice.Properties("Update Criteria") = adCriteriaKey

RsTxn_ProInvoice.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_TempGSL(Optional wc As String)
sql = "Select * from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " " & wc
Set RsTxn_TempGSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TempGSL.CursorLocation = adUseServer
End If
RsTxn_TempGSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_GSLBagSizeUpdate(Optional wc As String)
sql = "update mst_gsl set BagSize = (DepositeWeight/DepositeBags)*1000 where Status <> 'B' and SpillageFlag = 0 And DepositeBags <> 0 " & wc
Set TmpRs = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   TmpRs.CursorLocation = adUseClient
End If
TmpRs.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_CreateTempGSL(mFromDate_ As Variant, Optional mSM_Prefix_ As Variant, Optional mGenLocation_ As Variant)
Dim sql As String
'sql = "Select * into TempGSl from Mst_GSL where CDTFDate<'2010-04-01 00:00:00' and SpillageFlag=0 "
'sql = sql & " And flag in('O','R','J','E')"
Call TableTxn_GSLBagSizeUpdate(" And " & mGenLocation_)

'Commented on 12 Nov 2010 247 pm
'If mSM_Prefix_ <> "" Then
'   sql = "Select * into TempGSl from Mst_GSL Where SM_Prefix = '" & mSM_Prefix_ & "' And CDTFDate < '" & mFromDate_ & "' and SpillageFlag=0 "
'   sql = sql & " And flag in('O','R','J','E') "
'   If mGenLocation_ <> "" Then
'      sql = sql & " And " & mGenLocation_
'   End If
'Else
''   sql = "Select * into TempGSl from Mst_GSL Where CDTFDate < '" & mFromDate_ & "' and SpillageFlag=0 "
''   sql = sql & " And flag in('O','R','J','E')"
'   sql = "Select * into TempGSl From Mst_GSL Where CDTFDate < '" & mFromDate_ & "' and SpillageFlag=0"
'   sql = sql & " And flag in('O','R','J','E') And SM_Prefix Not In (Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate_ & "' And flag in('O','R','J','E') And TFlag = 'O')"
'   If mGenLocation_ <> "" Then
'      sql = sql & " And " & mGenLocation_
'   End If
'End If


'Commented on 21 July 2011
'sql = "Select * into TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " From Mst_GSL Where CDTFDate < '" & mFromDate_ & "' and SpillageFlag=0"
'sql = sql & " And flag in('O','R','J','E') And SM_Prefix Not In (Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate_ & "' And flag in('O','R','J','E') And TFlag = 'O')"
'sql = sql & " And " & mGenLocation_


'Commented on 08 May  2013 for made up bags exclusion
'sql = "Select * into TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " From Mst_GSL Where CDTFDate < '" & mFromDate_ & "' and SpillageFlag=0 "
'sql = sql & " And flag in('O','R','J','E') " 'And SM_Prefix Not In (Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate_ & "' And flag in('O','R','J','E') And TFlag = 'O')"
'sql = sql & " And " & mGenLocation_


sql = "Select * into TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " From Mst_GSL Where CDTFDate < '" & mFromDate_ & "' and SpillageFlag=0 and MadeupFlag=0 "
sql = sql & " And Status in('O','E') " 'And SM_Prefix Not In (Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate_ & "' And flag in('O','R','J','E') And TFlag = 'O')"
sql = sql & " And " & mGenLocation_


Set RsTxn_CreateTempGSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CreateTempGSL.CursorLocation = adUseClient
End If
RsTxn_CreateTempGSL.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function
Function TableTxn_TempProInvoice(Optional wc As String)
sql = "Select * from Txn_TempProInvoice" & Gen_UserLoginID & Gen_mTimeStamp & " " & wc
Set RsTxn_TempProInvoice = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TempProInvoice.CursorLocation = adUseClient
End If
RsTxn_TempProInvoice.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CreateTempProInvoice()
Dim sql As String

If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


sql = "CREATE TABLE Txn_TempProInvoice" & Gen_UserLoginID & Gen_mTimeStamp & "("
sql = sql & " LocationID decimal(18, 0) NOT NULL ,"
sql = sql & " Mst_GSL_ID int , "
sql = sql & " SM_Prefix " & TextType & " (5) NOT NULL ,"
sql = sql & " ExchangeTypeID decimal(18, 0) NOT NULL ,"
sql = sql & " ServiceTaxNAVCode " & TextType & " (30) NULL ,"
sql = sql & " ClientID int NOT NULL ,"
sql = sql & " CommodityID decimal(18, 0) NOT NULL ,"
sql = sql & " NoOfBags decimal (18,2)  ,"
sql = sql & " Weight decimal (18,3) ,"
sql = sql & " Rate decimal(18, 2) NOT NULL , "
sql = sql & " Amount decimal(18, 2) NOT NULL , "
sql = sql & " TFlag " & TextType & " (1)   NOT NULL , "
sql = sql & " BillingCycleID decimal(18, 0) NOT NULL ,"
sql = sql & " BillingUnit " & TextType & " (1) NULL ,"
sql = sql & " NoOfDays decimal(18, 0) NOT NULL ,"
sql = sql & " InvoiceDate DateTime ,"
sql = sql & " PINo  Decimal(18,0), "
sql = sql & " CxTFDate  DateTime, "
sql = sql & " MonG_UID " & TextType & "(36),"
sql = sql & " HSNNo " & TextType & "(20),"
sql = sql & " AccountType " & TextType & "(10)) "

Set RsTxn_CreateTempProInvoice = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CreateTempProInvoice.CursorLocation = adUseClient
End If
RsTxn_CreateTempProInvoice.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function



Function TableTxn_CRDepositDetails(Optional ByVal wc As String = "")
sql = "Select * from Txn_CashReceiptDepositDetails " & wc
Set RsTxn_CRDepositDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CRDepositDetails.CursorLocation = adUseClient
End If
RsTxn_CRDepositDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_FAGBank(Optional wc As String)
sql = "Select * from Mst_FAGBank " & wc
Set RsMst_FAGBank = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_FAGBank.CursorLocation = adUseClient
End If
RsMst_FAGBank.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_ITVendor(Optional wc As String)
sql = "Select * from Mst_ITVendor " & wc
Set RsMst_ITVendor = New ADODB.Recordset
RsMst_ITVendor.CursorLocation = adUseClient
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_ITVendor.CursorLocation = adUseClient
End If
RsMst_ITVendor.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function


Function TableTxn_ITVendorClaim(Optional wc As String)
sql = "Select * from Txn_ITVendorClaim " & wc
Set RsTxn_ITVendorClaim = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ITVendorClaim.CursorLocation = adUseClient
End If
RsTxn_ITVendorClaim.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_ITVendorClaimDetails(Optional wc As String)
sql = "Select * from Txn_ITVendorClaimDetail " & wc
Set RsTxn_ITVendorClaimDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ITVendorClaimDetails.CursorLocation = adUseClient
End If
RsTxn_ITVendorClaimDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_ClientGroup(Optional ByVal wc As String = "")
sql = "Select * from Mst_ClientGroup " & wc
Set RsMst_ClientGroup = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_ClientGroup.CursorLocation = adUseClient
End If
RsMst_ClientGroup.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableMst_FF_RateCard(Optional ByVal wc As String = "")
sql = "Select * from Mst_FF_RateCard " & wc
Set RsMst_FF_RateCard = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_FF_RateCard.CursorLocation = adUseClient
End If
RsMst_FF_RateCard.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_FF_RateCard_Godown(Optional ByVal wc As String = "")
sql = "Select * from Mst_FF_RateCard_Godown " & wc
Set RsMst_FF_RateCard_Godown = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_FF_RateCard_Godown.CursorLocation = adUseClient
End If
RsMst_FF_RateCard_Godown.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_FF_RateCard_Godown_Bag(Optional ByVal wc As String = "")
sql = "Select * from Mst_FF_RateCard_Godown_Bag" & wc
Set RsMst_FF_RateCard_Godown_Bag = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_FF_RateCard_Godown_Bag.CursorLocation = adUseClient
End If
RsMst_FF_RateCard_Godown_Bag.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_NCDEXStorageRateCard(Optional ByVal wc As String = "")
sql = "Select * from Mst_NCDEXStorageRateCard " & wc
Set RsMst_NCDEXStorageRateCard = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_NCDEXStorageRateCard.CursorLocation = adUseClient
End If
RsMst_NCDEXStorageRateCard.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_NCDEXStorageRateCardDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_NCDEXStorageRateCardDetail " & wc
Set RsMst_NCDEXStorageRateCardDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_NCDEXStorageRateCardDetail.CursorLocation = adUseClient
End If
RsMst_NCDEXStorageRateCardDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CreateTempGSLForDayWise(mFromDate_ As Variant, Optional mSM_Prefix_ As Variant)
Dim sql As String
If mSM_Prefix_ <> "" Then
   sql = "Select * into TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " from Mst_GSL Where SM_Prefix = '" & mSM_Prefix_ & "' And CDTFDate < '" & mFromDate_ & "' and SpillageFlag=0 "
   sql = sql & " And Status in('O','E')"
Else
   sql = "Select * into TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " From Mst_GSL Where CDTFDate < '" & mFromDate_ & "' and SpillageFlag=0"
   sql = sql & " And Status in('O','E') And SM_Prefix Not In (Select Distinct SM_Prefix From Txn_DayWiseGSLDetail" & Gen_UserName & Gen_mTimeStamp & " Where flag in('O','R','J','E') And TFlag = 'O')"
End If

Set RsTxn_CreateTempGSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CreateTempGSL.CursorLocation = adUseClient
End If
RsTxn_CreateTempGSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_DayWiseGSLDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " " & wc
Set RsTxn_DayWiseGSLDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_DayWiseGSLDetail.CursorLocation = adUseClient
End If
RsTxn_DayWiseGSLDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CreateTempDayWise()

Dim sql As String

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If
 

sql = " CREATE TABLE Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " ("
sql = sql & " Mst_GSL_ID int NOT NULL ,"
sql = sql & " SM_Prefix " & TextType & " (5)  NOT NULL ,"
sql = sql & " CMPID decimal(18, 0) NOT NULL ,"
sql = sql & " CommodityID decimal(18, 0) NULL ,"
sql = sql & " GodownID decimal(18, 0) NULL ,"
sql = sql & " StackNo " & TextType & " (10)  NOT NULL ,"
sql = sql & " LotNo " & TextType & " (25)  NOT NULL ,"
sql = sql & " BalanceBags decimal(18, 0) NULL ,"
sql = sql & " BalanceWeight decimal(18, 3) NULL ,"
sql = sql & " ExchangeTypeID decimal(18, 0) NULL ,"
sql = sql & " ClientID decimal(18, 0) NULL ,"
sql = sql & " Status " & TextType & " (1)  NULL ,"
sql = sql & " Flag1 " & TextType & " (1)  NULL ,"
sql = sql & " GrossWt decimal(18, 3) NULL ,"
sql = sql & " CxTFNO decimal(18, 0) NULL ,"
sql = sql & " TFlag " & TextType & " (1)  NULL ,"
sql = sql & " CDTFNo decimal(18, 0) NULL , "
sql = sql & " CxTF_DetailsID int NULL , "
sql = sql & " CDTFDate datetime NULL , "
sql = sql & " CxTFDate datetime NULL ) "

Set RsTxn_CreateTempDayWise = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CreateTempDayWise.CursorLocation = adUseClient
End If
RsTxn_CreateTempDayWise.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Function TableTxn_TempDailyGSL(Optional wc As String)
sql = "Select * from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " " & wc
Set RsTxn_TempDailyGSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TempDailyGSL.CursorLocation = adUseClient
End If
RsTxn_TempDailyGSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_Fumigation(Optional ByVal wc As String = "")
sql = "Select * from Txn_Fumigation " & wc
Set RsTxn_Fumigation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_Fumigation.CursorLocation = adUseClient
End If
RsTxn_Fumigation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_FumigationDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_FumigationDetail " & wc
Set RsTxn_FumigationDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_FumigationDetail.CursorLocation = adUseClient
End If
RsTxn_FumigationDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_Spraying(Optional ByVal wc As String = "")
sql = "Select * from Txn_Spraying " & wc
Set RsTxn_Spraying = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_Spraying.CursorLocation = adUseClient
End If
RsTxn_Spraying.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_RAGMemberTransfer(Optional ByVal wc As String = "")
sql = "Select * from Txn_RAGMemberTransfer " & wc
Set RsTxn_RAGMemberTransfer = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_RAGMemberTransfer.CursorLocation = adUseClient
End If
RsTxn_RAGMemberTransfer.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_RAGMemberTransferDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_RAGMemberTransferDetail " & wc
Set RsTxn_RAGMemberTransferDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_RAGMemberTransferDetail.CursorLocation = adUseClient
End If
RsTxn_RAGMemberTransferDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_QualityParameter(Optional wc As String)
sql = "Select * from Mst_QualityParameter " & wc   '  & tmp & "'"
Set RsMst_QualityParameter = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_QualityParameter.CursorLocation = adUseClient
End If
RsMst_QualityParameter.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_CommodityQualityParameter(Optional wc As String)
sql = "Select * from Mst_CommodityQualityParameter " & wc
Set RsMst_CommodityQualityParameter = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_CommodityQualityParameter.CursorLocation = adUseClient
End If
RsMst_CommodityQualityParameter.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_TestingAgency(Optional wc As String)
sql = "Select * from Mst_TestingAgency " & wc
Set RsMst_TestingAgency = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_TestingAgency.CursorLocation = adUseClient
End If
RsMst_TestingAgency.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_QualityInformation(Optional wc As String)
sql = "Select * from Txn_QualityInformation " & wc
Set RsTxn_QualityInformation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_QualityInformation.CursorLocation = adUseClient
End If
RsTxn_QualityInformation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_QualityInformationDetail(Optional wc As String)
sql = "Select * from Txn_QualityInformationDetail " & wc
Set RsTxn_QualityInformationDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_QualityInformationDetail.CursorLocation = adUseClient
End If
RsTxn_QualityInformationDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CxTF_BOE(Optional ByVal wc As String = "")
sql = "Select * from Txn_CxTF_BOE " & wc
Set RsTxn_CxTF_BOE = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CxTF_BOE.CursorLocation = adUseClient
End If
RsTxn_CxTF_BOE.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CreateTempGenInvoice()
Dim sql As String

If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


sql = "CREATE TABLE Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & "("
sql = sql & " Mst_GSL_ID int , "
sql = sql & " SM_Prefix " & TextType & " (5) ,"
sql = sql & " ExchangeTypeID decimal(18, 0)  ,"
sql = sql & " CMPID decimal(18, 0)  , "
sql = sql & " ClientID int ,"
sql = sql & " CommodityID decimal(18, 0) ,"
sql = sql & " BalanceBags decimal (18,2)  ,"
sql = sql & " BalanceWeight decimal (18,3) ,"
sql = sql & " Rate decimal(18, 4)  , "
sql = sql & " InvoiceAmount decimal(18, 2) , "
sql = sql & " TFlag " & TextType & " (1), "
sql = sql & " BillingCycleID decimal(18, 0) ,"
sql = sql & " BillingUnit " & TextType & " (1),"
sql = sql & " NoOfDays decimal(18, 0) ,"
sql = sql & " InvoiceDate DateTime ,"
sql = sql & " PINo Decimal(18,0), "
sql = sql & " Bagsize Decimal(18,0), "
sql = sql & " ContractID Decimal(18,0), "
sql = sql & " CxTFDate  DateTime ,"
sql = sql & " CxTF_DetailsID  int ,"
sql = sql & " CDTFDate DateTime, "
sql = sql & " YearMonth  DateTime, "
sql = sql & " GodownYearlyBillingFlg  " & TextType & "(1), "
sql = sql & " CommodityYearlyBillingFlg  " & TextType & "(1), " 'sql = sql & " CommodityYearlyBillingFlg  " & TextType & "(1)) "
'Added by Parag on 02 Sept 2016
sql = sql & " FranchClaimID  Decimal(18,0), "
sql = sql & " FranchBillingUnit  " & TextType & "(1), "
sql = sql & " MonG_UID  " & TextType & "(36), "
sql = sql & " GodownID  decimal(18, 0)) "

Set RsTxn_CreateTempGenInvoice = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CreateTempGenInvoice.CursorLocation = adUseClient
End If
RsTxn_CreateTempGenInvoice.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Function TableTxn_TempGenInvoice(Optional wc As String)

sql = "Select * from Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & " " & wc
Set RsTxn_TempGenInvoice = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TempGenInvoice.CursorLocation = adUseClient
End If
RsTxn_TempGenInvoice.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Function TableTxn_Hologram(Optional ByVal wc As String = "")
sql = "Select * from Txn_Hologram " & wc
Set RsTxn_Hologram = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_Hologram.CursorLocation = adUseClient
End If
RsTxn_Hologram.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CITFCR(Optional wc As String)
sql = "Select * from Txn_CITFCR " & wc
Set RsTxn_CITFCR = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CITFCR.CursorLocation = adUseClient
End If
RsTxn_CITFCR.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CITFCRDetail(Optional wc As String)
sql = "Select * from Txn_CITFCRDetail " & wc
Set RsTxn_CITFCRDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CITFCRDetail.CursorLocation = adUseClient
End If
RsTxn_CITFCRDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_OtherIncome(Optional ByVal wc As String = "")
sql = "Select * from Txn_OtherIncome " & wc
Set RsTxn_OtherIncome = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_OtherIncome.CursorLocation = adUseClient
End If
RsTxn_OtherIncome.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_MonthEmp_History(Optional wc As String)
sql = "select * from Txn_MonthEmp_History " & wc
Set RsTxn_MonthEmp_History = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_MonthEmp_History.CursorLocation = adUseClient
End If
RsTxn_MonthEmp_History.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_MonthEmp_HistoryDetail(Optional wc As String)
sql = "select * from Txn_MonthEmp_HistoryDetail " & wc
Set RsTxn_MonthEmp_HistoryDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_MonthEmp_HistoryDetail.CursorLocation = adUseClient
End If
RsTxn_MonthEmp_HistoryDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TmpTableHRPro()
Call DBCHRPro
sql = tmp
Set TmpRsHRPro = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRsHRPro.CursorLocation = adUseClient
End If

TmpRsHRPro.Open sql, cnHRPro, adOpenKeyset, adLockOptimistic

End Function

Function TableMst_IncomeExpenseType(Optional ByVal wc As String = "")
sql = "Select * from Mst_IncomeExpenseType " & wc
Set RsMst_IncomeExpenseType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_IncomeExpenseType.CursorLocation = adUseClient
End If
RsMst_IncomeExpenseType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_SecLocMapping(Optional ByVal wc As String = "")
sql = "Select * from Mst_SecLocMapping " & wc
Set RsMst_SecLocMapping = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecLocMapping.CursorLocation = adUseClient
End If
RsMst_SecLocMapping.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_SecLocMappingDetails(Optional wc As String)
sql = "Select * from Mst_SecLocMappingDetail " & wc
Set RsMst_SecLocMappingDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecLocMappingDetails.CursorLocation = adUseClient
End If
RsMst_SecLocMappingDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_SecurityDesignation(Optional wc As String)
sql = "Select * from Mst_SecurityDesignation " & wc
Set RsMst_SecurityDesignation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecurityDesignation.CursorLocation = adUseClient
End If
RsMst_SecurityDesignation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_MonthwiseSecLoc(Optional ByVal wc As String = "")
sql = "Select * from Txn_MonthwiseSecLoc " & wc
Set RsTxn_MonthwiseSecLoc = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_MonthwiseSecLoc.CursorLocation = adUseClient
End If
RsTxn_MonthwiseSecLoc.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_MonthwiseSecLocDetails(Optional wc As String)
sql = "Select * from Txn_MonthwiseSecLocDetail " & wc
Set RsTxn_MonthwiseSecLocDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_MonthwiseSecLocDetail.CursorLocation = adUseClient
End If
RsTxn_MonthwiseSecLocDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_SecurityBilling(Optional ByVal wc As String = "")
sql = "Select * from Txn_SecurityBilling " & wc
Set RsTxn_SecurityBilling = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_SecurityBilling.CursorLocation = adUseClient
End If
RsTxn_SecurityBilling.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_SecurityBillingDetail(Optional wc As String)
sql = "Select * from Txn_SecurityBillingDetail " & wc
Set RsTxn_SecurityBillingDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_SecurityBillingDetail.CursorLocation = adUseClient
End If
RsTxn_SecurityBillingDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_StorageCardDepoRateDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_StorageRateCardDepoDetail " & wc
Set RsMst_StorageCardDepoRateDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_StorageCardDepoRateDetail.CursorLocation = adUseClient
End If
RsMst_StorageCardDepoRateDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_SpecialCardDepoRateDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_SpecialRateCardDepoDetail " & wc
Set RsMst_SpecialCardDepoRateDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SpecialCardDepoRateDetail.CursorLocation = adUseClient
End If
RsMst_SpecialCardDepoRateDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_CoExpense(Optional wc As String)
sql = "Select * from Mst_CoExpense " & wc
Set RsMst_CoExpense = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_CoExpense.CursorLocation = adUseClient
End If
RsMst_CoExpense.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_Region(Optional wc As String)
sql = "select * from Mst_Region " & wc
Set RsMst_Region = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Region.CursorLocation = adUseClient
End If
RsMst_Region.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TmpTableCMPro()
Call DBCCMPro
sql = tmp
Set TmpRsCMPro = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRsCMPro.CursorLocation = adUseClient
End If

TmpRsCMPro.Open sql, cnCMPro, adOpenKeyset, adLockOptimistic

End Function

Function TableMst_StateDRCMapping(Optional wc As String)
sql = "Select * from Mst_StateDRCMapping " & wc
Set RsMst_StateDRCMapping = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_StateDRCMapping.CursorLocation = adUseClient
End If
RsMst_StateDRCMapping.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_ReservationInvoice(Optional wc As String)
sql = "Select * from Txn_ReservationInvoice " & wc
Set RsTxn_ReservationInvoice = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ReservationInvoice.CursorLocation = adUseServer
End If
RsTxn_ReservationInvoice.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_TempInvoice(Optional wc As String)
sql = "select * from Txn_TempInvoice " & wc
Set RsTxn_TempInvoice = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TempInvoice.CursorLocation = adUseClient
End If
RsTxn_TempInvoice.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_TIDetails(Optional wc As String)
sql = "select * from Txn_TIDetails " & wc
Set RsTxn_TIDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TIDetails.CursorLocation = adUseClient
End If
RsTxn_TIDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_NewCashReceipt(Optional wc As String)
sql = "select * from Txn_NewCashReceipt " & wc
Set RsTxn_NewCashReceipt = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_NewCashReceipt.CursorLocation = adUseClient
End If
RsTxn_NewCashReceipt.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_NewCashReceiptClientDetail(Optional wc As String)
sql = "select * from Txn_NewCashReceiptClientDetail " & wc
Set RsTxn_NewCashReceiptClientDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_NewCashReceiptClientDetail.CursorLocation = adUseClient
End If
RsTxn_NewCashReceiptClientDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_NewCashReceiptTIDetail(Optional wc As String)
sql = "select * from Txn_NewCashReceiptTIDetail " & wc
Set RsTxn_NewCashReceiptTIDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_NewCashReceiptTIDetail.CursorLocation = adUseClient
End If
RsTxn_NewCashReceiptTIDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_NewCRDepositDetails(Optional ByVal wc As String = "")
sql = "Select * from Txn_NewCashReceiptDepositDetails " & wc
Set RsTxn_NewCRDepositDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_NewCRDepositDetails.CursorLocation = adUseClient
End If
RsTxn_NewCRDepositDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_StateIndia(Optional ByVal wc As String = "")
sql = "Select * from Mst_StateIndia " & wc
Set RsMst_StateIndia = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_StateIndia.CursorLocation = adUseClient
End If
RsMst_StateIndia.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_ClientLedger(Optional ByVal wc As String = "")
sql = "Select * from Txn_ClientLedger " & wc
Set RsTxn_ClientLedger = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ClientLedger.CursorLocation = adUseClient
End If
RsTxn_ClientLedger.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CRTDSAdjustment(Optional wc As String)
sql = "select * from Txn_CRTDSAdjustment " & wc
Set RsTxn_CRTDSAdjustment = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CRTDSAdjustment.CursorLocation = adUseClient
End If
RsTxn_CRTDSAdjustment.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CRTDSAdjustmentDetail(Optional wc As String)
sql = "select * from Txn_CRTDSAdjustmentDetail " & wc
Set RsTxn_CRTDSAdjustmentDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CRTDSAdjustmentDetail.CursorLocation = adUseClient
End If
RsTxn_CRTDSAdjustmentDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CRAdvanceAllocation(Optional wc As String)
sql = "select * from Txn_CRAdvanceAllocation " & wc
Set RsTxn_CRAdvanceAllocation = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CRAdvanceAllocation.CursorLocation = adUseClient
End If
RsTxn_CRAdvanceAllocation.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function




Function TableTxn_LockSealHologram(Optional ByVal wc As String = "")
sql = "Select * from Txn_LockSealHologram " & wc
Set RsTxn_LockSealHologram = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_LockSealHologram.CursorLocation = adUseClient
End If
RsTxn_LockSealHologram.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function



Function TableTxn_LockSealHologramDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_LockSealHologramDetail " & wc
Set RsTxn_LockSealHologramDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_LockSealHologramDetail.CursorLocation = adUseClient
End If
RsTxn_LockSealHologramDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableLog_CxTF_Details(Optional ByVal wc As String = "")
sql = "Select * from Log_CxTF_Details " & wc
Set RsLog_CxTF_Details = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_CxTF_Details.CursorLocation = adUseClient
End If
RsLog_CxTF_Details.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_WHR_Detail(Optional ByVal wc As String = "")
sql = "Select * from Log_WHR_Detail " & wc
Set RsLog_WHR_Detail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_WHR_Detail.CursorLocation = adUseClient
End If
RsLog_WHR_Detail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_WHRPrint(Optional ByVal wc As String = "")
sql = "Select * from Log_WHRPrint " & wc
Set RsLog_WHRPrint = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_WHRPrint.CursorLocation = adUseClient
End If
RsLog_WHRPrint.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CDTFBOE(Optional ByVal wc As String = "")
sql = "Select * from Txn_CDTFBOE " & wc
Set RsTxn_CDTFBOE = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CDTFBOE.CursorLocation = adUseClient
End If
RsTxn_CDTFBOE.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_CDTFBOE_GSL(Optional ByVal wc As String = "")
sql = "Select * from Txn_CDTFBOE_GSL " & wc
Set RsTxn_CDTFBOE_GSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CDTFBOE_GSL.CursorLocation = adUseClient
End If
RsTxn_CDTFBOE_GSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_NAVInvoiceOutStanding(Optional ByVal wc As String = "")
sql = "Select * from Txn_NAVInvoiceOutStanding " & wc
Set RsTxn_NAVInvoiceOutStanding = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_NAVInvoiceOutStanding.CursorLocation = adUseClient
End If
RsTxn_NAVInvoiceOutStanding.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

'RsLog_FormAccessDetail
Function TableLog_FormAccessDetail(Optional ByVal wc As String = "")
sql = "Select * from Log_FormAccessDetail " & wc
Set RsLog_FormAccessDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_FormAccessDetail.CursorLocation = adUseClient
End If
RsLog_FormAccessDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_PledgeCashReceipt(Optional ByVal wc As String = "")
sql = "Select * from Txn_PledgeCashReceipt " & wc
Set RsTxn_PledgeCashReceipt = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_PledgeCashReceipt.CursorLocation = adUseClient
End If
RsTxn_PledgeCashReceipt.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_PledgeCashReceiptDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_PledgeCashReceiptDetail " & wc
Set RsTxn_PledgeCashReceiptDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_PledgeCashReceiptDetail.CursorLocation = adUseClient
End If
RsTxn_PledgeCashReceiptDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_FranchiseeRent(Optional ByVal wc As String = "")
sql = "Select * from Txn_FranchiseeRent " & wc
Set RsTxn_FranchiseeRent = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_FranchiseeRent.CursorLocation = adUseClient
End If
RsTxn_FranchiseeRent.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_NCDEXAccreditedCapacity(Optional wc As String)
sql = "Select * from Mst_NCDEXAccreditedCapacity " & wc   '  & tmp & "'"
Set RsMst_NCDEXAccreditedCapacity = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_NCDEXAccreditedCapacity.CursorLocation = adUseClient
End If
RsMst_NCDEXAccreditedCapacity.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_EarthquakeZone(Optional ByVal wc As String = "")
sql = "Select * from Mst_EarthquakeZone " & wc
Set RsMst_EarthquakeZone = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_EarthquakeZone.CursorLocation = adUseClient
End If
RsMst_EarthquakeZone.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_FranchiseeFeeRejectHistory(Optional ByVal wc As String = "")
sql = "Select * from Log_FranchiseeFeeRejectHistory " & wc
Set RsLog_FranchiseeFeeRejectHistory = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_FranchiseeFeeRejectHistory.CursorLocation = adUseClient
End If
RsLog_FranchiseeFeeRejectHistory.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_ChequeReturnReason(Optional wc As String)
sql = "Select * from Mst_ChequeReturnReason " & wc
Set RsMst_ChequeReturnReason = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_ChequeReturnReason.CursorLocation = adUseClient
End If
RsMst_ChequeReturnReason.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableLog_Client2ClientTransfer(Optional ByVal wc As String = "")
sql = "Select * from Log_Client2ClientTransfer " & wc
Set RsLog_Client2ClientTransfer = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_Client2ClientTransfer.CursorLocation = adUseClient
End If
RsLog_Client2ClientTransfer.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_EmployeeAllocationDetails(Optional wc As String)
sql = "Select * from Mst_EmployeeAllocationDetails " & wc
Set RsMst_EmployeeAllocationDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_EmployeeAllocationDetails.CursorLocation = adUseClient
End If
RsMst_EmployeeAllocationDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_FumigationApproval(Optional ByVal wc As String = "")
sql = "Select * from Txn_FumigationApproval " & wc
Set RsTxn_FumigationApproval = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_FumigationApproval.CursorLocation = adUseClient
End If
RsTxn_FumigationApproval.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableLog_ISINCommodity(Optional ByVal wc As String = "")
sql = "Select * from Log_ISINCommodity " & wc
Set RsLog_ISINCommodity = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_ISINCommodity.CursorLocation = adUseClient
End If
RsLog_ISINCommodity.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_SecurityRequest(Optional wc As String)
sql = "Select * from Mst_SecurityRequest " & wc
Set RsMst_SecurityRequest = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecurityRequest.CursorLocation = adUseClient
End If
RsMst_SecurityRequest.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_SecurityRequestGodownMapping(Optional wc As String)
sql = "Select * from Mst_SecurityRequestGodownMapping " & wc
Set RsMst_SecurityRequestGodownMapping = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecurityRequestGodownMapping.CursorLocation = adUseClient
End If
RsMst_SecurityRequestGodownMapping.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_SecurityAttendance(Optional wc As String)
sql = "Select * from Txn_SecurityAttendance " & wc
Set RsTxn_SecurityAttendance = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_SecurityAttendance.CursorLocation = adUseClient
End If
RsTxn_SecurityAttendance.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_TICollectionRemark(Optional wc As String)
sql = "Select * from Log_TICollectionRemark " & wc
Set RsLog_TICollectionRemark = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_TICollectionRemark.CursorLocation = adUseClient
End If
RsLog_TICollectionRemark.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TableTxn_LessorAgreeAttachment(Optional wc As String)
Call DBCSM_Attachment
sql = "Select * from Txn_LessorAgreeAttachment " & wc
Set RsTxn_LessorAgreeAttachment = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsTxn_LessorAgreeAttachment.CursorLocation = adUseClient
End If

RsTxn_LessorAgreeAttachment.Open sql, cnSM_Attachment, adOpenKeyset, adLockOptimistic

End Function

Public Function TmpTableSMAttachment()
Call DBCSM_Attachment
sql = tmp
Set TmpRsSM_Attachment = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRsSM_Attachment.CursorLocation = adUseClient
End If

TmpRsSM_Attachment.Open sql, cnSM_Attachment, adOpenKeyset, adLockOptimistic

End Function


Function TableMst_FF_RateCard_Godown_Fix(Optional wc As String)
sql = "Select * from Mst_FF_RateCard_Godown_Fix " & wc
Set RsMst_FF_RateCard_Godown_Fix = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_FF_RateCard_Godown_Fix.CursorLocation = adUseClient
End If
RsMst_FF_RateCard_Godown_Fix.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Public Function TableTxn_SContractAttachment(Optional wc As String)

sql = "Select * from Txn_StorageContractAttachment " & wc

Set RsTxn_SContractAttachment = New ADODB.Recordset
Call DBCSM_Attachment
If Gen_DBType = "SQL" Then
   RsTxn_SContractAttachment.CursorLocation = adUseClient
End If

RsTxn_SContractAttachment.Open sql, cnSM_Attachment, adOpenKeyset, adLockOptimistic

End Function

Public Function TableMst_SContractBillingDetail(Optional wc As String)
sql = "Select * from Mst_SContractBillingDetail " & wc
Set RsMst_SContractBillingDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SContractBillingDetail.CursorLocation = adUseClient
End If
RsMst_SContractBillingDetail.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Function TableTxn_RRNPWR(Optional wc As String)
sql = "Select * from Txn_RRNPWR " & wc
Set RsTxn_RRNPWR = New ADODB.Recordset
RsTxn_RRNPWR.CursorLocation = adUseClient
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_RRNPWR.CursorLocation = adUseClient
End If
RsTxn_RRNPWR.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_RRNPWRGSL(Optional wc As String)
sql = "Select * from Txn_RRNPWRGSL " & wc
Set RsTxn_RRNPWRGSL = New ADODB.Recordset
RsTxn_RRNPWRGSL.CursorLocation = adUseClient
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_RRNPWRGSL.CursorLocation = adUseClient
End If
RsTxn_RRNPWRGSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TableMst_LicenseAttachment(Optional wc As String)
Call DBCSM_Attachment
sql = "Select * from Mst_LicenseAttachment " & wc
Set RsMst_LicenseAttachment = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsMst_LicenseAttachment.CursorLocation = adUseClient
End If
RsMst_LicenseAttachment.Open sql, cnSM_Attachment, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_ReservationInvoiceDetail(Optional wc As String)
sql = "Select * from Txn_ReservationInvoiceDetail " & wc
Set RsTxn_ReservationInvoiceDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ReservationInvoiceDetail.CursorLocation = adUseClient
End If
RsTxn_ReservationInvoiceDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_ReservationDailyStock(Optional wc As String)
sql = "Select * from Txn_ReservationDailyStock " & wc
Set RsTxn_ReservationDailyStock = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ReservationDailyStock.CursorLocation = adUseClient
End If
RsTxn_ReservationDailyStock.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InspectionExpenseType(Optional ByVal wc As String = "")
sql = "Select * from Mst_InspectionExpenseType " & wc
Set RsMst_InspectionExpenseType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InspectionExpenseType.CursorLocation = adUseClient
End If
RsMst_InspectionExpenseType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_GodownInspection(Optional ByVal wc As String = "")
sql = "Select * from Mst_GodownInspection " & wc
Set RsMst_GodownInspection = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_GodownInspection.CursorLocation = adUseClient
End If
RsMst_GodownInspection.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InspectionGodownDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_InspectionGodownDetail " & wc
Set RsMst_InspectionGodownDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InspectionGodownDetail.CursorLocation = adUseClient
End If
RsMst_InspectionGodownDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InspectionCommodityDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_InspectionCommodityDetail " & wc
Set RsMst_InspectionCommodityDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InspectionCommodityDetail.CursorLocation = adUseClient
End If
RsMst_InspectionCommodityDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InspectionExpenseDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_InspectionExpenseDetail " & wc
Set RsMst_InspectionExpenseDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InspectionExpenseDetail.CursorLocation = adUseClient
End If
RsMst_InspectionExpenseDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableMst_TreatmentType(Optional ByVal wc As String = "")
sql = "Select * from Mst_TreatmentType" & wc
Set RsMst_TreatmentType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_TreatmentType.CursorLocation = adUseClient
End If
RsMst_TreatmentType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_FumigationAgreement(Optional ByVal wc As String = "")
sql = "Select * from Mst_FumigationAgreement" & wc
Set RsMst_FumigationAgreement = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_FumigationAgreement.CursorLocation = adUseClient
End If
RsMst_FumigationAgreement.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_FumigationTreatmentDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_FumigationTreatmentDetail " & wc
Set RsMst_FumigationTreatmentDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_FumigationTreatmentDetail.CursorLocation = adUseClient
End If
RsMst_FumigationTreatmentDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_FumigationQuotationDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_FumigationQuotationDetail " & wc
Set RsMst_FumigationQuotationDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_FumigationQuotationDetail.CursorLocation = adUseClient
End If
RsMst_FumigationQuotationDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_TRF(Optional wc As String)
sql = "Select * from Txn_TRF " & wc
Set RsTxn_TRF = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TRF.CursorLocation = adUseClient
End If
RsTxn_TRF.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_TRFDetail(Optional wc As String)
sql = "Select * from Txn_TRFDetail " & wc
Set RsTxn_TRFDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TRFDetail.CursorLocation = adUseClient
End If
RsTxn_TRFDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableExpiredGSL_Contract(Optional ByVal wc As String = "", Optional ByVal OrdBy As String = "")

sql = "SELECT Distinct Mst_GSL.NContractID "
sql = sql & " From Mst_GSL "
sql = sql & " Inner Join Txn_SRGodownMapping On Mst_GSL.NContractID = Txn_SRGodownMapping.ContractID "
sql = sql & " Inner Join Mst_StorageContract On Txn_SRGodownMapping.ContractID = Mst_StorageContract.ContractID   "
sql = sql & "  " & wc & " And (Mst_GSL.Flag = 'E' ) "

If Gen_WcLocation <> "" Then
   sql = sql & " And Mst_StorageContract." & Gen_WcLocation
End If
sql = sql & "  " & OrdBy & " "
 
Set RsExpiredGSL_Contract = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsExpiredGSL_Contract.CursorLocation = adUseClient
End If
RsExpiredGSL_Contract.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_MonthlyCommtrackFileDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_MonthlyCommtrackFileDetail " & wc
Set RsTxn_MonthlyCommtrackFileDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_MonthlyCommtrackFileDetail.CursorLocation = adUseClient
End If

RsTxn_MonthlyCommtrackFileDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_MonthlyCommtrackInvoice(Optional ByVal wc As String = "")
sql = "Select * from Txn_MonthlyCommtrackInvoice " & wc
Set RsTxn_MonthlyCommtrackInvoice = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_MonthlyCommtrackInvoice.CursorLocation = adUseClient
End If

RsTxn_MonthlyCommtrackInvoice.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_TempInvSRGodownDetails(Optional wc As String)
sql = "Select * from Txn_TempInvSRGodownDetails " & wc
Set RsTxn_TempInvSRGodownDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TempInvSRGodownDetails.CursorLocation = adUseClient
End If
RsTxn_TempInvSRGodownDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_RRN_LotDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_RRN_LotDetail " & wc
Set RsMst_RRNLotDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_RRNLotDetail.CursorLocation = adUseClient
End If
RsMst_RRNLotDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_InvoiceDeletetion(Optional ByVal wc As String = "")
sql = "Select * from Log_InvoiceDeletetion " & wc
Set RsLog_InvoiceDeletetion = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_InvoiceDeletetion.CursorLocation = adUseClient
End If
RsLog_InvoiceDeletetion.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_InsuranceExceptionMapping(Optional ByVal wc As String = "")
sql = "Select * from Txn_InsuranceExceptionMapping " & wc
Set RsTxn_InsuranceExceptionMapping = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_InsuranceExceptionMapping.CursorLocation = adUseClient
End If
RsTxn_InsuranceExceptionMapping.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_SecurityAgreement(Optional wc As String)
sql = "Select * from Txn_SecurityAgreement" & wc
Set RsTxn_SecurityAgreement = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_SecurityAgreement.CursorLocation = adUseClient
End If
RsTxn_SecurityAgreement.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_SecurityAgreementDetail(Optional wc As String)
sql = "Select * from Txn_SecurityAgreementDetail" & wc
Set RsTxn_SecurityAgreementDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_SecurityAgreementDetail.CursorLocation = adUseClient
End If
RsTxn_SecurityAgreementDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_TempInvGodownDetails(Optional wc As String)
sql = "Select * from Txn_TempInvGodownDetails " & wc
Set RsTxn_TempInvGodownDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_TempInvGodownDetails.CursorLocation = adUseClient
End If
RsTxn_TempInvGodownDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_MonthlyInsuranceCharges(Optional wc As String)
sql = "Select * from Txn_MonthlyInsuranceCharges " & wc
Set RsTxn_MonthlyInsuranceCharges = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_MonthlyInsuranceCharges.CursorLocation = adUseClient
End If
RsTxn_MonthlyInsuranceCharges.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_MonthlyInsuranceChargesDetail(Optional wc As String)
sql = "Select * from Txn_MonthlyInsuranceChargesDetail " & wc
Set RsTxn_MonthlyInsuranceChargesDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_MonthlyInsuranceChargesDetail.CursorLocation = adUseClient
End If
RsTxn_MonthlyInsuranceChargesDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_GodownRentST(Optional wc As String)
sql = "Select * from Txn_GodownRentST " & wc
Set RsTxn_GodownRentST = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GodownRentST.CursorLocation = adUseClient
End If
RsTxn_GodownRentST.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_GodownRentSTClaimDetail(Optional wc As String)
sql = "Select * from Txn_GodownRentSTClaimDetail " & wc
Set RsTxn_GodownRentSTClaimDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GodownRentSTClaimDetail.CursorLocation = adUseClient
End If
RsTxn_GodownRentSTClaimDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_GodownRentSTAgreeDetail(Optional wc As String)
sql = "Select * from Txn_GodownRentSTAgreeDetail " & wc
Set RsTxn_GodownRentSTAgreeDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GodownRentSTAgreeDetail.CursorLocation = adUseClient
End If
RsTxn_GodownRentSTAgreeDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TableTxn_SecurityAgreeAttachment(Optional wc As String)
Call DBCSM_Attachment
sql = "Select * from Txn_SecurityAgreeAttachment " & wc
Set RsTxn_SecurityAgreeAttachment = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsTxn_SecurityAgreeAttachment.CursorLocation = adUseClient
End If

RsTxn_SecurityAgreeAttachment.Open sql, cnSM_Attachment, adOpenKeyset, adLockOptimistic

End Function

Public Function TableMst_FSAgreeAttachment(Optional wc As String)
Call DBCSM_Attachment
sql = "Select * from Mst_FSAgreeAttachment " & wc
Set RsMst_FSAgreeAttachment = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsMst_FSAgreeAttachment.CursorLocation = adUseClient
End If

RsMst_FSAgreeAttachment.Open sql, cnSM_Attachment, adOpenKeyset, adLockOptimistic

End Function

Function TableTxn_ScmQualityParameter(Optional wc As String)
sql = "Select * from Txn_ScmQualityParameter " & wc   '  & tmp & "'"
Set RsTxn_ScmQualityParameter = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ScmQualityParameter.CursorLocation = adUseClient
End If
RsTxn_ScmQualityParameter.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_ScmQualityParameterDetail(Optional wc As String)
sql = "Select * from Txn_ScmQualityParameterDetail " & wc   '  & tmp & "'"
Set RsTxn_ScmQualityParameterDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ScmQualityParameterDetail.CursorLocation = adUseClient
End If
RsTxn_ScmQualityParameterDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_QQSGodownMapping(Optional wc As String)
sql = "Select * from Txn_QQSGodownMapping " & wc   '  & tmp & "'"
Set RsTxn_QQSGodownMapping = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_QQSGodownMapping.CursorLocation = adUseClient
End If
RsTxn_QQSGodownMapping.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_QQSGodownMappingDetail(Optional wc As String)
sql = "Select * from Txn_QQSGodownMappingDetail " & wc   '  & tmp & "'"
Set RsTxn_QQSGodownMappingDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_QQSGodownMappingDetail.CursorLocation = adUseClient
End If
RsTxn_QQSGodownMappingDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_DocumentType(Optional wc As String)
sql = "select * from Mst_DocumentType " & wc
Set RsMst_DocumentType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_DocumentType.CursorLocation = adUseClient
End If
RsMst_DocumentType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_GodownAgreeDocCycle(Optional wc As String)
sql = "select * from Txn_GodownAgreeDocCycle " & wc
Set RsTxn_GodownAgreeDocCycle = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GodownAgreeDocCycle.CursorLocation = adUseClient
End If
RsTxn_GodownAgreeDocCycle.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_GADCDocumentAttachDetail(Optional wc As String)
sql = "select * from Txn_GADCDocumentAttachDetail " & wc
Set RsTxn_GADCDocumentAttachDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GADCDocumentAttachDetail.CursorLocation = adUseClient
End If
RsTxn_GADCDocumentAttachDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_GADCGodownDetail(Optional wc As String)
sql = "select * from Txn_GADCGodownDetail " & wc
Set RsTxn_GADCGodownDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GADCGodownDetail.CursorLocation = adUseClient
End If
RsTxn_GADCGodownDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableLog_GADCForwardToHistory(Optional wc As String)
sql = "select * from Log_GADCForwardToHistory " & wc
Set RsLog_GADCForwardToHistory = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_GADCForwardToHistory.CursorLocation = adUseClient
End If
RsLog_GADCForwardToHistory.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_LicenseType(Optional ByVal wc As String = "")
sql = "Select * from Mst_LicenseType " & wc
Set RsMst_LicenseType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_LicenseType.CursorLocation = adUseClient
End If
RsMst_LicenseType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_LicenseAllocationDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_LicenseAllocationDetail " & wc
Set RsTxn_LicenseAllocationDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_LicenseAllocationDetail.CursorLocation = adUseClient
End If
RsTxn_LicenseAllocationDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_LicenseGodownDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_LicenseGodownDetail " & wc
Set RsTxn_LicenseGodownDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_LicenseGodownDetail.CursorLocation = adUseClient
End If
RsTxn_LicenseGodownDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_SCMDO(Optional wc As String)
sql = "Select * from Txn_SCMDO " & wc
Set RsTxn_SCMDO = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_SCMDO.CursorLocation = adUseClient
End If
RsTxn_SCMDO.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_YearlyStorageRateCard(Optional ByVal wc As String = "")
sql = "Select * from Mst_YearlyStorageRateCard " & wc
Set RsMst_YearlyStorageRateCard = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_YearlyStorageRateCard.CursorLocation = adUseClient
End If
RsMst_YearlyStorageRateCard.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_YearlyStorageRateCardDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_YearlyStorageRateCardDetail " & wc
Set RsMst_YearlyStorageRateCardDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_YearlyStorageRateCardDetail.CursorLocation = adUseClient
End If
RsMst_YearlyStorageRateCardDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_VehicleType(Optional ByVal wc As String = "")
sql = "Select * from Mst_VehicleType " & wc
Set RsMst_VehicleType = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_VehicleType.CursorLocation = adUseClient
End If
RsMst_VehicleType.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableMst_InspectionAreaDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_InspectionAreaDetail " & wc
Set RsMst_InspectionAreaDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InspectionAreaDetail.CursorLocation = adUseClient
End If
RsMst_InspectionAreaDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function



Public Function TableTxn_ImportNcdexLotPleadge(Optional wc As String)
sql = "Select * from Txn_ImportNcdexLotPleadge " & wc
Set RsTxn_ImportNcdexLotPleadge = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ImportNcdexLotPleadge.CursorLocation = adUseClient
End If
RsTxn_ImportNcdexLotPleadge.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_MAW(Optional ByVal wc As String = "")
sql = "Select * from Mst_MAW " & wc
Set RsMst_MAW = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_MAW.CursorLocation = adUseClient
End If
RsMst_MAW.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_MAWDetails(Optional ByVal wc As String = "")
sql = "Select * from Mst_MAWDetails " & wc
Set RsMst_MAWDetails = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_MAWDetails.CursorLocation = adUseClient
End If
RsMst_MAWDetails.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

''-- Stack Master

 ''===== Mst_Stack
Function TableMst_Stack(Optional ByVal wc As String = "")
sql = "Select * from Mst_Stack " & wc
Set RsMst_Stack = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_Stack.CursorLocation = adUseClient
End If
RsMst_Stack.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function



Function TableLog_UpdateCMPBlockDays(Optional ByVal wc As String = "")
sql = "Select * from Log_UpdateCMPBlockDays" & wc
Set RsLog_UpdateCMPBlockDays = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsLog_UpdateCMPBlockDays.CursorLocation = adUseClient
End If
RsLog_UpdateCMPBlockDays.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function



''===== Txn_ClientFeedbackDoc
Function TableTxn_ClientFeedbackDoc(Optional ByVal wc As String = "")
sql = "Select * from Txn_ClientFeedbackDoc " & wc
Set RsTxn_ClientFeedbackDoc = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_ClientFeedbackDoc.CursorLocation = adUseClient
End If
RsTxn_ClientFeedbackDoc.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

''===== Txn_ClientFeedbackDocAttachment
Public Function TableTxn_ClientFeedbackDocAttachment(Optional wc As String)
Call DBCSM_Attachment
sql = "Select * from Txn_ClientFeedbackDocAttachment " & wc
Set RsTxn_ClientFeedbackDocAttachment = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsTxn_ClientFeedbackDocAttachment.CursorLocation = adUseClient
End If
RsTxn_ClientFeedbackDocAttachment.Open sql, cnSM_Attachment, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_MonthlyCommtrackFileDetailCMSP(Optional ByVal wc As String = "")
sql = "Select * from Txn_MonthlyCommtrackFileDetailCMSP " & wc
Set RsTxn_MonthlyCommtrackFileDetailCMSP = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_MonthlyCommtrackFileDetailCMSP.CursorLocation = adUseClient
End If
RsTxn_MonthlyCommtrackFileDetailCMSP.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


''--- For Navision On 05 Aug 2016
Function TableTxn_InvoicesForNAVISION(Optional ByVal wc As String = "")
sql = "Select * from Txn_InvoicesForNAVISION " & wc
Set RsTxn_InvoicesForNAVISION = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_InvoicesForNAVISION.CursorLocation = adUseClient
End If
RsTxn_InvoicesForNAVISION.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_CashRepeiptsForNAVISION(Optional ByVal wc As String = "")
sql = "Select * from Txn_CashRepeiptsForNAVISION " & wc
Set RsTxn_CashRepeiptsForNAVISION = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_CashRepeiptsForNAVISION.CursorLocation = adUseClient
End If
RsTxn_CashRepeiptsForNAVISION.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_InvoiceAllocationForNAVISION(Optional ByVal wc As String = "")
sql = "Select * from Txn_InvoiceAllocationForNAVISION " & wc
Set RsTxn_InvoiceAllocationForNAVISION = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_InvoiceAllocationForNAVISION.CursorLocation = adUseClient
End If
RsTxn_InvoiceAllocationForNAVISION.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_SecurityBillForNavision(Optional ByVal wc As String = "")
sql = "Select * from Txn_SecurityBillForNavision " & wc
Set RsTxn_SecurityBillForNavision = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_SecurityBillForNavision.CursorLocation = adUseClient
End If
RsTxn_SecurityBillForNavision.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Function TableTxn_IncomeExpenseForDailyPnL(Optional ByVal wc As String = "")
sql = "Select * from Txn_IncomeExpenseForDailyPnL " & wc
Set RsTxn_IncomeExpenseForDailyPnL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_IncomeExpenseForDailyPnL.CursorLocation = adUseClient
End If
RsTxn_IncomeExpenseForDailyPnL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

''-- For Client State Wise GST On 17 July 2017

''---- For Client State Wise GST
Function TableMst_ClientStateWiseGST(Optional ByVal wc As String = "")
sql = "Select * from Mst_ClientStateWiseGST " & wc
Set RsMst_ClientStateWiseGST = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_ClientStateWiseGST.CursorLocation = adUseClient
End If
RsMst_ClientStateWiseGST.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''---- For Depository State Wise GST
Function TableMst_DepositoryStateWiseGST(Optional ByVal wc As String = "")
sql = "Select * from Mst_DepositoryStateWiseGST " & wc
Set RsMst_DepositoryStateWiseGST = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_DepositoryStateWiseGST.CursorLocation = adUseClient
End If
RsMst_DepositoryStateWiseGST.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''---- For Lessor State Wise GST
Function TableMst_LessorStateWiseGST(Optional ByVal wc As String = "")
sql = "Select * from Mst_LessorStateWiseGST " & wc
Set RsMst_LessorStateWiseGST = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_LessorStateWiseGST.CursorLocation = adUseClient
End If
RsMst_LessorStateWiseGST.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''---- For SecurityAgency State Wise GST
Function TableMst_SecurityAgencyStateWiseGST(Optional ByVal wc As String = "")
sql = "Select * from Mst_SecurityAgencyStateWiseGST " & wc
Set RsMst_SecurityAgencyStateWiseGST = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_SecurityAgencyStateWiseGST.CursorLocation = adUseClient
End If
RsMst_SecurityAgencyStateWiseGST.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
''---- For Vendor State Wise GST
Function TableMst_VendorStateWiseGST(Optional ByVal wc As String = "")
sql = "Select * from Mst_VendorStateWiseGST " & wc
Set RsMst_VendorStateWiseGST = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_VendorStateWiseGST.CursorLocation = adUseClient
End If
RsMst_VendorStateWiseGST.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Public Function TableMst_GLAccountIncomeTypeMapping(Optional ByVal wc As String = "")

sql = "Select * from Mst_GLAccountIncomeTypeMapping " & wc
Set RsMst_GLAccountIncomeTypeMapping = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_GLAccountIncomeTypeMapping.CursorLocation = adUseClient
End If
RsMst_GLAccountIncomeTypeMapping.Open sql, cn, adOpenKeyset, adLockOptimistic

End Function

Public Function TableTxn_DepositContract(Optional ByVal wc As String = "")
sql = "Select * from Txn_DepositContract " & wc
Set RsTxn_DepositContract = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_DepositContract.CursorLocation = adUseClient
End If
RsTxn_DepositContract.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TableTxn_DepositContractGodownDetail(Optional ByVal wc As String = "")
sql = "Select * from Txn_DepositContractGodownDetail " & wc
Set RsTxn_DepositContractGodownDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_DepositContractGodownDetail.CursorLocation = adUseClient
End If
RsTxn_DepositContractGodownDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TableMst_InsuranceBy(Optional ByVal wc As String = "")
sql = "Select * from Mst_InsuranceBy " & wc
Set RsMst_InsuranceBy = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InsuranceBy.CursorLocation = adUseClient
End If
RsMst_InsuranceBy.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


Public Function TableTxn_DepositContractAttachment(Optional wc As String)
Call DBCSM_Attachment
sql = "Select * from Txn_DepositContractAttachment " & wc
Set RsTxn_DepositContractAttachment = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsTxn_DepositContractAttachment.CursorLocation = adUseClient
End If
RsTxn_DepositContractAttachment.Open sql, cnSM_Attachment, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InsuranceRiskCovered(Optional ByVal wc As String = "")
sql = "Select * from Mst_InsuranceRiskCovered " & wc
Set RsMst_InsuranceRiskCovered = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InsuranceRiskCovered.CursorLocation = adUseClient
End If
RsMst_InsuranceRiskCovered.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InsuranceForUnitTypeDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_InsuranceForUnitTypeDetail " & wc
Set RsMst_InsuranceForUnitTypeDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InsuranceForUnitTypeDetail.CursorLocation = adUseClient
End If
RsMst_InsuranceForUnitTypeDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InsuranceValidFor(Optional ByVal wc As String = "")
sql = "Select * from Mst_InsuranceValidFor " & wc
Set RsMst_InsuranceValidFor = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InsuranceValidFor.CursorLocation = adUseClient
End If
RsMst_InsuranceValidFor.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InsuranceGodownMapping(Optional ByVal wc As String = "")
sql = "Select * from Mst_InsuranceGodownMapping " & wc
Set RsMst_InsuranceGodownMapping = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InsuranceGodownMapping.CursorLocation = adUseClient
End If
RsMst_InsuranceGodownMapping.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InsuranceCommodityDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_InsuranceCommodityDetail " & wc
Set RsMst_InsuranceCommodityDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InsuranceCommodityDetail.CursorLocation = adUseClient
End If
RsMst_InsuranceCommodityDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableMst_InsuranceClientDetail(Optional ByVal wc As String = "")
sql = "Select * from Mst_InsuranceClientDetail " & wc
Set RsMst_InsuranceClientDetail = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsMst_InsuranceClientDetail.CursorLocation = adUseClient
End If
RsMst_InsuranceClientDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function


''-- Ullhas Patil 13th June 2018
Function TableTxn_InsurancePremium(Optional ByVal wc As String = "")
    sql = "Select * from Txn_InsurancePremium " & wc
    Set RsTxn_InsurancePremium = New ADODB.Recordset
    Call DBC
    If Gen_DBType = "SQL" Then
       RsTxn_InsurancePremium.CursorLocation = adUseClient
    End If
    RsTxn_InsurancePremium.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableTxn_InsurancePremiumDetail(Optional ByVal wc As String = "")
    sql = "Select * from Txn_InsurancePremiumDetail " & wc
    Set RsTxn_InsurancePremiumDetail = New ADODB.Recordset
    Call DBC
    If Gen_DBType = "SQL" Then
       RsTxn_InsurancePremiumDetail.CursorLocation = adUseClient
    End If
    RsTxn_InsurancePremiumDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function
Function TableLog_Txn_InsurancePremiumDetail(Optional ByVal wc As String = "")
    sql = "Select * from Log_Txn_InsurancePremiumDetail " & wc
    Set RsLog_Txn_InsurancePremiumDetail = New ADODB.Recordset
    Call DBC
    If Gen_DBType = "SQL" Then
       RsLog_Txn_InsurancePremiumDetail.CursorLocation = adUseClient
    End If
    RsLog_Txn_InsurancePremiumDetail.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Function TableTxn_GSL(Optional ByVal wc As String = "")
sql = "Select * from Txn_GSL " & wc
Set RsTxn_GSL = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_GSL.CursorLocation = adUseClient
End If
RsTxn_GSL.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

Public Function TmpTableAzureStoreMan()

'Call DBCCMPro
'On Error GoTo Err
Call DBCAzureStoreMan
sql = tmp
Set TmpRsAzureStoreMan = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRsAzureStoreMan.CursorLocation = adUseClient
End If

TmpRsAzureStoreMan.Open sql, cnAzureStoreMan, adOpenKeyset, adLockOptimistic

Exit Function

Err: msg = Err.Description

End Function

Public Function TmpTableAzureStoreMan1()

'Call DBCCMPro
On Error GoTo Err
Call DBCAzureStoreMan
sql = tmp
Set TmpRs1AzureStoreMan = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   TmpRs1AzureStoreMan.CursorLocation = adUseClient
End If

TmpRs1AzureStoreMan.Open sql, cnAzureStoreMan, adOpenKeyset, adLockOptimistic

Exit Function

Err: msg = Err.Description

End Function
Public Function TableMst_ClientAttachment(Optional wc As String)
Call DBCSM_Attachment
sql = "Select * from Mst_ClientAttachment " & wc
Set RsMst_ClientAttachment = New ADODB.Recordset
If Gen_DBType = "SQL" Then
   RsMst_ClientAttachment.CursorLocation = adUseClient
End If

RsMst_ClientAttachment.Open sql, cnSM_Attachment, adOpenKeyset, adLockOptimistic

End Function
Function TableTxn_WHRToNFin(Optional ByVal wc As String = "")
sql = "Select * from Txn_WHRToNFin " & wc
Set RsTxn_WHRToNFin = New ADODB.Recordset
Call DBC
If Gen_DBType = "SQL" Then
   RsTxn_WHRToNFin.CursorLocation = adUseClient
End If
RsTxn_WHRToNFin.Open sql, cn, adOpenKeyset, adLockOptimistic
End Function

