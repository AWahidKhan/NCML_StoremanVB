VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportSecurityBilling 
   Caption         =   "Export Security Billing"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   135
      ScaleHeight     =   795
      ScaleWidth      =   14955
      TabIndex        =   9
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox CmbJournalCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9360
         TabIndex        =   4
         Top             =   240
         Width           =   2160
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13875
         TabIndex        =   7
         Top             =   195
         Width           =   975
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12780
         TabIndex        =   6
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Get Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11580
         TabIndex        =   5
         Top             =   195
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Old"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox CmbBatch 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CmbApprovedBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Journal Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   12
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Batch No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Approved By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8655
      Left            =   120
      TabIndex        =   8
      Top             =   1000
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15266
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Frm_ExportSecurityBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExpFlag, MaxID As Variant
Dim mForwardToID As Double
Dim mJournalBatchName, mJournalType As Variant
Dim mDocNo As Variant

Dim mSecurityExpenseID As Double
Dim mDocumentNo As Variant


Private Sub CmbApprovedBy_GotFocus()
tmp1 = CmbApprovedBy.Text
'Call TableMst_Employee("Where Flag = 'Y' And DesignationID=")
tmp = "Select ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName From Mst_Employee ME"
tmp = tmp & " Inner join Mst_Users MU On MU.EmployeeNo=ME.EmployeeNo Where MU.UserRoleID=7 And ME.Flag = 'Y'"
Call TmpTable
CmbApprovedBy.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Approved By Found!!!"
   CmbApprovedBy.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbApprovedBy.AddItem TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
    TmpRs.MoveNext
Next
CmbApprovedBy.Text = tmp1
End Sub

Private Sub CmbApprovedBy_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbApprovedBy.Text = "" Then
   mForwardToID = 0
   Exit Sub
End If

strSplitString = Split(CmbApprovedBy, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbApprovedBy.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


TmpRs.MoveFirst
TmpRs.Find "EmployeeNo = '" & EmpNo_ & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Selection "
   CmbApprovedBy.SetFocus
   Exit Sub
End If
mForwardToID = TmpRs!EmployeeID

End Sub

Private Sub CmbJournalCode_GotFocus()
Dim mR As Variant
tmp1 = CmbJournalCode
CmbJournalCode.Clear
tmp = " SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]" '''JT.JournalTypeName,JCM.JournalCode
tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where ME.EmployeeNo = '" & Gen_UserLoginID & "' Order by JCM.JournalTypeID"
Call TmpTableHRPro
If TmpRsHRPro.RecordCount > 0 Then
For mR = 1 To TmpRsHRPro.RecordCount
   CmbJournalCode.AddItem IIf(IsNull(TmpRsHRPro!JournalCode), "", TmpRsHRPro!JournalCode)
   TmpRsHRPro.MoveNext
Next
Else
   MsgBox "No Record found !!!! "
   Exit Sub
End If
CmbJournalCode = tmp1
End Sub

Private Sub CmbJournalCode_LostFocus()
Dim mSplit As Variant
If CmbJournalCode.Text = "" Then
   mJournalBatchName = ""
   mJournalType = ""
   Exit Sub
End If
mSplit = Split(CmbJournalCode.Text, "~")
If UBound(mSplit) <> 1 Then
   MsgBox "Invalid selection!!! "
   CmbJournalCode.SetFocus
   Exit Sub
End If
TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "JournalCode = '" & CmbJournalCode.Text & "'"
If TmpRsHRPro.EOF Then
   MsgBox "Invalid selection "
   CmbJournalCode.SetFocus
   Exit Sub
End If
mJournalType = UCase(mSplit(0))
mJournalBatchName = UCase(mSplit(1))
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportSecurityBilling.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportSecurityBilling.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExport_Click()
Dim Path, strFilName, strFilLocation As Variant
Dim mUserBatchID, mInvoiceType As Variant

Path = App.Path
''-- On 08 Aug 2016
'mJournalType = "GENERAL"
'mUserBatchID = GetGUID
'mJournalBatchName = GetJournalDetailsFromHRPro(Gen_UserLoginID, mJournalType)
'If Trim(mJournalBatchName) = "" Then
'   MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'   Exit Sub
'End If

''-- Ullhas Patil 04 Oct 2017
''''-- On 05 Aug 2016
If Option1.Value = True Then
    If CmbJournalCode.Text = "" Then
       MsgBox "Journal Code not selected"
       CmbJournalCode.SetFocus
       Exit Sub
    End If
    mUserBatchID = GetGUID
'    If Trim(mJournalType) = "" Or Trim(mJournalBatchName) = "" Then
'       MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'       Exit Sub
'    End If
End If
''--------------------------
mInvoiceType = "SECURITY"

If Option1.Value = True Then
   strFilName = "\Navision\SecurityNavUpload_SMClaim" & MaxID & ".csv"
Else
   strFilName = "\Navision\SecurityNavUpload_SMClaim" & CmbBatch.Text & ".csv"
End If

''-- On 27 July 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
    ''-- Ullhas Patil 04 Oct 2017
    Call TableTxn_SecurityBillForNavision(" Where G_UID is null ")
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
    ''--------------------------
End If

With MSHFlexGrid1
For I = 1 To MSHFlexGrid1.Rows - 1
    tmp = ""
    If Option1.Value = True Then
       If .TextMatrix(I, 0) <> "" Then
        ''-- Ullhas Patil 04 Oct 2017
        RsTxn_SecurityBillForNavision.AddNew
        RsTxn_SecurityBillForNavision!PostingDate = CDate(.TextMatrix(I, 0))
        RsTxn_SecurityBillForNavision!DocumentType = Trim(.TextMatrix(I, 1))
        RsTxn_SecurityBillForNavision!AccountType = Trim(.TextMatrix(I, 2))
        RsTxn_SecurityBillForNavision!AccountNo = Trim(.TextMatrix(I, 3))
        RsTxn_SecurityBillForNavision!Amount = Trim(.TextMatrix(I, 4))
        RsTxn_SecurityBillForNavision!TDSNatureofDeduction = Trim(.TextMatrix(I, 5))
        RsTxn_SecurityBillForNavision!ServiceTaxGroupCode = Trim(.TextMatrix(I, 6))
        RsTxn_SecurityBillForNavision!Narration = Trim(.TextMatrix(I, 7))
        RsTxn_SecurityBillForNavision!RegionDim = Trim(.TextMatrix(I, 8))
        RsTxn_SecurityBillForNavision!DeptDim = Trim(.TextMatrix(I, 9))
        RsTxn_SecurityBillForNavision!VendorDim_EmpDim = Trim(.TextMatrix(I, 10))
        RsTxn_SecurityBillForNavision!ExternalDocumentNo = Trim(.TextMatrix(I, 11))
        RsTxn_SecurityBillForNavision!InvoiceDate = IIf(Trim(.TextMatrix(I, 12)) = "", Null, Trim(.TextMatrix(I, 12)))
        RsTxn_SecurityBillForNavision!InvoiceType = mInvoiceType
        'RsTxn_SecurityBillForNavision!PINo = Left(Trim(.TextMatrix(I, 11)), InStr(Trim(.TextMatrix(I, 11)), "-") - 1)
        RsTxn_SecurityBillForNavision!CreatedBy = Gen_UserLoginID
        RsTxn_SecurityBillForNavision!G_UID = GetGUID
        RsTxn_SecurityBillForNavision!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
        RsTxn_SecurityBillForNavision!JournalBatchType = Trim(mJournalType)
        RsTxn_SecurityBillForNavision!JournalBatchName = Trim(mJournalBatchName)
        'RsTxn_SecurityBillForNavision!DocumentNo = Trim(.TextMatrix(i, 13))
        RsTxn_SecurityBillForNavision!UserBatchID = mUserBatchID
        
        RsTxn_SecurityBillForNavision!ProjectDim = Trim(.TextMatrix(I, 13))
        RsTxn_SecurityBillForNavision!GSTGroupCode = Trim(.TextMatrix(I, 14))
        RsTxn_SecurityBillForNavision!HSNNo = Trim(.TextMatrix(I, 15))
        RsTxn_SecurityBillForNavision!AccType = Trim(.TextMatrix(I, 16))
        RsTxn_SecurityBillForNavision!DocumentNo = Trim(.TextMatrix(I, 17))
        'RsTxn_SecurityBillForNavision!DataTransfer = "X"
        RsTxn_SecurityBillForNavision.Update

'        For c = 0 To MSHFlexGrid1.Cols - 2
'            If tmp = "" Then
'               tmp = tmp & MSHFlexGrid1.TextMatrix(i, c)
'            Else
'               tmp = tmp & "," & MSHFlexGrid1.TextMatrix(i, c)
'            End If
'        Next
'        Print #1, tmp

        ''--------------------------
       End If
    Else
        For C = 0 To MSHFlexGrid1.Cols - 1
            If tmp = "" Then
               tmp = tmp & MSHFlexGrid1.TextMatrix(I, C)
            Else
               tmp = tmp & "," & MSHFlexGrid1.TextMatrix(I, C)
            End If
        Next
        Print #1, tmp
    End If
Next
End With

If Option1.Value = True Then
'   tmp = "Update Txn_SecurityBilling Set BatchID = " & MaxID & " Where BatchID = 0 and  Flag = 'A'"
'   Call TmpTable
   If CmbApprovedBy.Text = "" Then
      tmp = "Update Txn_SecurityBilling set BatchID = " & MaxID & " Where Isnull(BatchID,0) = 0 and  Status = 'A'"
   Else
      tmp = "Update Txn_SecurityBilling set BatchID = " & MaxID & " Where Isnull(BatchID,0) = 0 and  Status = 'A' And ForwardTo = " & mForwardToID & ""
   End If
   Call TmpTable
End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
End Sub

Private Sub CmdGo_Click()

ExpFlag = ""
Call Grid_Head
If Option1.Value = True Then
   tmp = " Select TSB.ApprovedDate,'Invoice' As Inv,'G/L Account' As GLAccount,'7195900000000',"
   'tmp = tmp & " Sum (TSBD.BillDetailAmt + TSBD.BillDetailSTax) 'Amount', '' As TDS,ST.NAVCode 'ServiceTaxNAVCode',"
   tmp = tmp & " Sum (TSBD.BillDetailAmt) 'Amount', '' As TDS,ST.NAVCode 'ServiceTaxNAVCode',"
   tmp = tmp & " 'Being expenses incurred against Storeman Batch ID " & MaxID & "' As 'ApprovalRemarks',"
'Commented on 12 Oct 2017 after mail from Vishnu Sir on 10 Oct 2017
'   tmp = tmp & " l.NAVCode 'LocationNAVCode','CMP' As DeptDim ,"
'   tmp = tmp & " '' 'EmployeeNAVCode', TSB.BatchID,TSB.MonthYear,TSBD.BillTxnID ,L.LocationName,TSB.BillNo,ST.Description"
   
   tmp = tmp & " Max(l.NAVCode) 'LocationNAVCode','CMP' As DeptDim ,"
   tmp = tmp & " '' 'EmployeeNAVCode', TSB.BatchID,TSB.BillDate,TSBD.BillTxnID ,TSB.BillNo,ST.Description"
   
   tmp = tmp & " From Txn_SecurityBillingDetail TSBD"
   tmp = tmp & " Inner Join Txn_SecurityBilling TSB on TSBD.BillTxnID = TSB.BillTxnID"
   tmp = tmp & " Inner Join Mst_SecurityAgency E on TSB.SecurityAgencyID = E.SecurityAgencyID"
   tmp = tmp & " Inner Join Mst_Location L on TSBD.LocationID = L.LocationID"
   tmp = tmp & " Left  Join Mst_ServiceTax ST on TSB.ServiceTaxID = ST.ServiceTaxID"
   tmp = tmp & " Where Isnull(TSB.BatchID,0) =  0   and  TSB.Status in ('S', 'A')"
 
  If CmbApprovedBy.Text <> "" Then
     tmp = tmp & " And TSB.ForwardTo=" & mForwardToID & ""
  End If
  
  tmp = tmp & " Group By TSB.ApprovedDate,ST.NAVCode,TSB.ApprovedDate,"
  tmp = tmp & " TSB.BatchID, TSB.BillDate, TSBD.BillTxnID, "
  tmp = tmp & " TSB.BillNo,ST.Description"
  tmp = tmp & " Having Sum (TSBD.BillDetailAmt) > 0 "
  tmp = tmp & " Order By TSB.ApprovedDate"
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
   tmp = " Select TSB.ApprovedDate,'Invoice' As Inv,'G/L Account' As GLAccount,'7195900000000',"
   'tmp = tmp & " Sum (TSBD.BillDetailAmt + TSBD.BillDetailSTax) 'Amount', '' As TDS,ST.NAVCode 'ServiceTaxNAVCode',"
   tmp = tmp & " Sum (TSBD.BillDetailAmt) 'Amount', '' As TDS,ST.NAVCode 'ServiceTaxNAVCode',"
   tmp = tmp & " 'Being expenses incurred against Storeman Batch ID " & CmbBatch.Text & "' As 'ApprovalRemarks',"

'Commented on 12 Oct 2017 after mail from Vishnu Sir on 10 Oct 2017
'   tmp = tmp & " l.NAVCode 'LocationNAVCode','CMP' As DeptDim ,"
'   tmp = tmp & " '' 'EmployeeNAVCode', TSB.BatchID,TSB.MonthYear,TSBD.BillTxnID ,L.LocationName,TSB.BillNo,ST.Description"
   
   tmp = tmp & " Max(l.NAVCode) 'LocationNAVCode','CMP' As DeptDim ,"
   tmp = tmp & " '' 'EmployeeNAVCode', TSB.BatchID,TSB.BillDate,TSBD.BillTxnID ,TSB.BillNo,ST.Description"
   tmp = tmp & " From Txn_SecurityBillingDetail TSBD"
   tmp = tmp & " Inner Join Txn_SecurityBilling TSB on TSBD.BillTxnID = TSB.BillTxnID"
   tmp = tmp & " Inner Join Mst_SecurityAgency E on TSB.SecurityAgencyID = E.SecurityAgencyID"
   tmp = tmp & " Inner Join  Mst_Location L on TSBD.LocationID = L.LocationID"
   tmp = tmp & " Left  Join Mst_ServiceTax ST on TSB.ServiceTaxID = ST.ServiceTaxID"
   tmp = tmp & " Where TSB.BatchID =  " & Val(CmbBatch.Text) & " and  TSB.Status in ('S', 'A')"
   tmp = tmp & " Group By  TSB.ApprovedDate,E.NAVCode, TSB.TotalAmt,ST.NAVCode,"
   tmp = tmp & " TSBD.BillTxnID,"
   tmp = tmp & " TSB.FAGRemarks , E.NAVCode, TSB.BatchID, TSB.BillNo, TSB.BillTxnID, TSB.BillDate,TSB.BillNo,ST.Description"
   tmp = tmp & " Having Sum (TSBD.BillDetailAmt) > 0 "
   tmp = tmp & " Order By TSB.ApprovedDate"
 
End If
Call TmpTable

If Option1.Value = True Then
   tmp = " Select TSB.ApprovedDate,'Invoice' As Inv,'Vendor' As GLAccount,MSAGST.GSTNavCode 'ExpenseTypeNAVCode'," ''-- MSA.NAVCode
   'tmp = tmp & " (TSB.TotalAmt) 'Amount', T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',"
   tmp = tmp & " (TSB.BillAmt) 'Amount', T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',"
   tmp = tmp & " TSB.FAGRemarks + ' - Bill Txn ID:' + convert(varchar(10),TSB.BillTxnID) As 'ApprovalRemarks',"
   tmp = tmp & " '201001' 'LocationNAVCode','CMP' As DeptDim ,"
   tmp = tmp & " MSA.NAVCode 'AgencyNAVCode', TSB.BatchID ,'' As InvoiceDate,TSB.BillNo,TSB.BillTxnID, TSB.BillDate,TSB.BillingStateIndiaID"
   tmp = tmp & " From  Txn_SecurityBilling TSB"
   tmp = tmp & " Inner Join Mst_SecurityAgency MSA on TSB.SecurityAgencyID = MSA.SecurityAgencyID"
   tmp = tmp & " Left Join Mst_TDS T on MSA.TDSID = T.TDSID"
   
   ''-- Ullhas Patil 04 Oct 2017
   tmp = tmp & " Left Join Mst_SecurityAgencyStateWiseGST MSAGST on TSB.BillingStateIndiaID = MSAGST.StateIndiaID And TSB.SecurityAgencyID = MSAGST.SecurityAgencyID"
   ''----------------------------
   
   tmp = tmp & " Where Isnull(TSB.BatchID,0) =  0 and  TSB.Status in ('S', 'A') And TSB.BillAmt > 0 "
    If CmbApprovedBy.Text <> "" Then
       tmp = tmp & " And TSB.ForwardTo=" & mForwardToID & ""
    End If
    tmp = tmp & " Order By TSB.ApprovedDate"
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
   tmp = " Select TSB.ApprovedDate,'Invoice' As Inv,'Vendor' As GLAccount,MSAGST.GSTNavCode 'ExpenseTypeNAVCode'," ''-- MSA.NAVCode
   'tmp = tmp & " (TSB.TotalAmt) 'Amount', T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',"
   tmp = tmp & " (TSB.BillAmt) 'Amount', T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',"
   tmp = tmp & " TSB.FAGRemarks + ' - Bill Txn ID:' + convert(varchar(10),TSB.BillTxnID) As 'ApprovalRemarks',"
   tmp = tmp & " '201001' 'LocationNAVCode','CMP' As DeptDim ,"
   tmp = tmp & " MSA.NAVCode 'AgencyNAVCode', TSB.BatchID ,'' As InvoiceDate,TSB.BillNo,TSB.BillTxnID, TSB.BillDate,TSB.BillingStateIndiaID"
   tmp = tmp & " From  Txn_SecurityBilling TSB"
   tmp = tmp & " Inner Join Mst_SecurityAgency MSA on TSB.SecurityAgencyID = MSA.SecurityAgencyID"
   
   ''-- Ullhas Patil 04 Oct 2017
   tmp = tmp & " Left Join Mst_SecurityAgencyStateWiseGST MSAGST on TSB.BillingStateIndiaID = MSAGST.StateIndiaID And TSB.SecurityAgencyID = MSAGST.SecurityAgencyID"
   ''----------------------------
   
   tmp = tmp & " Left Join Mst_TDS T on MSA.TDSID = T.TDSID"
   tmp = tmp & " Where TSB.BatchID =  " & Val(CmbBatch.Text) & " and  TSB.Status in ('S', 'A') And TSB.BillAmt > 0 "
   tmp = tmp & " Order By TSB.ApprovedDate"
End If
Call Tmp1Table
Dim mDat_ As Variant
If TmpRs.RecordCount > 0 And TmpRs1.RecordCount > 0 Then
    Dim x As Variant
    x = 0
    ExpFlag = "Y"
    mDat_ = IIf(IsNull(TmpRs.Fields(0)), "", TmpRs.Fields(0))
    For x = 1 To TmpRs.RecordCount
        If TmpRs!LocationNAVCode = "" Or IsNull(TmpRs!LocationNAVCode) Or TmpRs1!AgencyNAVCode = "" Or IsNull(TmpRs1!AgencyNAVCode) Then
           ExpFlag = "N"
        End If
        I = MSHFlexGrid1.Rows - 1
        
        'For c = 0 To MSHFlexGrid1.Cols - 1
        '    MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
        'Next
        
        
        ''-- Ullhas Patil 03 Oct 2017
        MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!ApprovedDate), "", TmpRs!ApprovedDate)
        MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!Inv), "", TmpRs!Inv)
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!GLAccount), "", TmpRs!GLAccount)
        
        MSHFlexGrid1.TextMatrix(I, 3) = ""
        If LCase(TmpRs!Description) = LCase("GST with Tax") Then
            MSHFlexGrid1.TextMatrix(I, 3) = GetGLAccNAVCodeFromExpenseType(mSecurityExpenseID, "NAGRI", "G")
        ElseIf LCase(TmpRs!Description) = LCase("GST without Tax") Then
            MSHFlexGrid1.TextMatrix(I, 3) = GetGLAccNAVCodeFromExpenseType(mSecurityExpenseID, "AGRI", "G")
        End If
        
        MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
        MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!TDS), "", TmpRs!TDS)
        
        MSHFlexGrid1.TextMatrix(I, 6) = ""  '' ServiceTaxNAVCode
        
        MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!ApprovalRemarks), "", TmpRs!ApprovalRemarks)
        MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!LocationNAVCode), "", TmpRs!LocationNAVCode)
        MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!DeptDim), "", TmpRs!DeptDim)
        MSHFlexGrid1.TextMatrix(I, 10) = "EMP-1569"
        MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!BillDate), "", TmpRs!BillDate)
        '''---------------
        
        'MSHFlexGrid1.TextMatrix(i, 0) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 0))
        'MSHFlexGrid1.TextMatrix(i, 12) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 12))
''        If CmbBatch.Text = "" Then
''           MSHFlexGrid1.TextMatrix(I, 11) = "SMS" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
''        Else
''           MSHFlexGrid1.TextMatrix(I, 11) = "SMS" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
''        End If
        
        If Option1.Value = True Then  ' If CmbBatch.Text = "" Then
           MSHFlexGrid1.TextMatrix(I, 11) = "SMS" & Trim(MaxID) & "-" & TmpRs!BillTxnID & "-" & IIf(IsNull(TmpRs!BillNo), "", "Inv^" & TmpRs!BillNo)
        Else
           If CmbBatch.Text > 274 Then
              MSHFlexGrid1.TextMatrix(I, 11) = "SMS" & Trim(CmbBatch.Text) & "-" & TmpRs!BillTxnID & "-" & IIf(IsNull(TmpRs!BillNo), "", "Inv^" & TmpRs!BillNo)
           Else
              MSHFlexGrid1.TextMatrix(I, 11) = "SMS" & Trim(CmbBatch.Text) & "-" & TmpRs!BillTxnID & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
           End If
        End If
        mDocNo = "DocNo" & TmpRs!BillTxnID
        
        ''--- Ullhas Patil 03 Oct 2017
        ''MSHFlexGrid1.TextMatrix(i, 13) = mDocNo
        MSHFlexGrid1.TextMatrix(I, 13) = ""
        MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!ServiceTaxNAVCode), "", TmpRs!ServiceTaxNAVCode)
        If LCase(TmpRs!Description) = LCase("GST with Tax") Then
            MSHFlexGrid1.TextMatrix(I, 15) = GetGLAccNAVCodeFromExpenseType(mSecurityExpenseID, "NAGRI", "H")
        ElseIf LCase(TmpRs!Description) = LCase("GST without Tax") Then
            MSHFlexGrid1.TextMatrix(I, 15) = GetGLAccNAVCodeFromExpenseType(mSecurityExpenseID, "AGRI", "H")
        End If
        MSHFlexGrid1.TextMatrix(I, 16) = ""
        
        If MSHFlexGrid1.TextMatrix(I, 14) = "" Or MSHFlexGrid1.TextMatrix(I, 15) = "" Then
            ExpFlag = "N"
        End If
        ''-----------------------------------------
        mDocumentNo = "DocNo-" & TmpRs!BillTxnID
        MSHFlexGrid1.TextMatrix(I, 17) = mDocumentNo
        
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        mDat_ = TmpRs!BillTxnID
        
        TmpRs.MoveNext
        If TmpRs.EOF Then
           CrEntry (mDat_)
           Exit For
        End If
        If mDat_ <> IIf(IsNull(TmpRs!BillTxnID), "", TmpRs!BillTxnID) Then
           Call CrEntry(mDat_)
           mDat_ = IIf(IsNull(TmpRs!BillTxnID), "", TmpRs!BillTxnID)
        End If
        
    Next

Else
   ExpFlag = "N"
End If

If ExpFlag = "N" Then
   CmdExport.Enabled = False
Else
   CmdExport.Enabled = True
End If

MSHFlexGrid1.Col = 11
MSHFlexGrid1.Sort = flexSortNumericDescending
End Sub
Sub CrEntry(mPDat_ As Variant)

TmpRs1.MoveFirst
TmpRs1.Find "BillTxnID = " & mPDat_ 'TmpRs!BillTxnID
'For x = 1 To TmpRs1.RecordCount
If TmpRs1.EOF = False Then
    If TmpRs1!ExpenseTypeNAVCode = "" Or IsNull(TmpRs1!ExpenseTypeNAVCode) Or TmpRs1!TDS = "" Or IsNull(TmpRs1!TDS) Or TmpRs1!LocationNAVCode = "" Or IsNull(TmpRs1!LocationNAVCode) Then
       ExpFlag = "N"
    End If
    I = MSHFlexGrid1.Rows - 1
    ''For c = 0 To MSHFlexGrid1.Cols - 1
    ''    MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs1.Fields(c)), "", TmpRs1.Fields(c))
    ''Next
    
    ''--- Ullhas Patil 03 Oct 2017
        MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!ApprovedDate), "", TmpRs1!ApprovedDate)
        MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!Inv), "", TmpRs1!Inv)
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!GLAccount), "", TmpRs1!GLAccount)
        MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!ExpenseTypeNAVCode), "", TmpRs1!ExpenseTypeNAVCode)
        MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!Amount), "", TmpRs1!Amount)
        MSHFlexGrid1.TextMatrix(I, 4) = Val(MSHFlexGrid1.TextMatrix(I, 4)) * -1
        MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!TDS), "", TmpRs1!TDS)
        MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!ServiceTaxNAVCode), "", TmpRs1!ServiceTaxNAVCode)
        MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!ApprovalRemarks), "", TmpRs1!ApprovalRemarks)
        MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!LocationNAVCode), "", TmpRs1!LocationNAVCode)
        MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!DeptDim), "", TmpRs1!DeptDim)
        MSHFlexGrid1.TextMatrix(I, 10) = "EMP-1569"
        MSHFlexGrid1.TextMatrix(I, 11) = ""
        MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!BillDate), "", TmpRs1!BillDate) 'ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 12))
        
    ''-----------------------------
    
    
'    MSHFlexGrid1.TextMatrix(i, 0) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 0))
'    MSHFlexGrid1.TextMatrix(i, 4) = Val(MSHFlexGrid1.TextMatrix(i, 4)) * -1
'    MSHFlexGrid1.TextMatrix(i, 10) = "EMP-1569"
'    MSHFlexGrid1.TextMatrix(i, 12) = IIf(IsNull(TmpRs1!MonthYear), "", TmpRs1!MonthYear) 'ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 12))
'    MSHFlexGrid1.TextMatrix(i, 12) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 12))
'    If CmbBatch.Text = "" Then
'       MSHFlexGrid1.TextMatrix(I, 11) = "SMV" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
'    Else
'       MSHFlexGrid1.TextMatrix(I, 11) = "SMV" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
'    End If

'        If CmbBatch.Text = "" Then
           ''MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!BillNo), "", "Invoice " & TmpRs1!BillNo) '"SMV" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
'        Else
'           MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!InvoiceNo), "", TmpRs!InvoiceNo) 'TmpRs!InvoiceNo  '"SMV" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
'        End If
   If Option1.Value = True Then  ' If CmbBatch.Text = "" Then
      MSHFlexGrid1.TextMatrix(I, 11) = "SMS" & Trim(MaxID) & "-" & TmpRs1!BillTxnID & "-" & IIf(IsNull(TmpRs1!BillNo), "", "Inv^" & TmpRs1!BillNo)
   Else
      If CmbBatch.Text > 274 Then
         MSHFlexGrid1.TextMatrix(I, 11) = "SMS" & Trim(CmbBatch.Text) & "-" & TmpRs1!BillTxnID & "-" & IIf(IsNull(TmpRs1!BillNo), "", "Inv^" & TmpRs1!BillNo)
      Else
         MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!BillNo), "", "Invoice " & TmpRs1!BillNo & "-" & TmpRs1!BillTxnID)  '"SMF" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
      End If
   End If
    
    ''--- Ullhas Patil 03 Oct 2017
    ''MSHFlexGrid1.TextMatrix(i, 13) = mDocNo
    MSHFlexGrid1.TextMatrix(I, 13) = ""
    MSHFlexGrid1.TextMatrix(I, 14) = ""
    MSHFlexGrid1.TextMatrix(I, 15) = ""
    MSHFlexGrid1.TextMatrix(I, 16) = ""
    MSHFlexGrid1.TextMatrix(I, 17) = mDocumentNo
    ''-----------------------------------------
    
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'    TmpRs1.MoveNext
'    If TmpRs1.EOF Then Exit For
'    If mPDat_ <> IIf(IsNull(TmpRs1.Fields(0)), "", TmpRs1.Fields(0)) Then Exit For
'Next
End If

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Picture1.Enabled = False
   MSHFlexGrid1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

''-- Ullhas Patil 03 Oct 2017
mSecurityExpenseID = 57
''--------------

Picture1.Enabled = True
MSHFlexGrid1.Enabled = True
Call Grid_Head
tmp = "Select Distinct BatchID from Txn_SecurityBilling Where BatchID > 0  order by BatchID"
Call Tmp4Table
CmbBatch.Clear
If TmpRs4.RecordCount > 0 Then
   For I = 1 To TmpRs4.RecordCount
       CmbBatch.AddItem TmpRs4.Fields(0)
       MaxID = TmpRs4.Fields(0) + 1
       TmpRs4.MoveNext
   Next
Else
   MaxID = 1
End If
End Sub
Private Sub Option2_Click()
CmdExport.Enabled = False
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbApprovedBy.Text = ""
   CmbApprovedBy.Enabled = False
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Enabled = False
   CmbApprovedBy.Text = ""
   CmbApprovedBy.Enabled = True
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Public Sub Grid_Head()
CmdExport.Enabled = False
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 18 '17 ''14

    .TextMatrix(0, 0) = "Posting Date":             .ColWidth(0) = 1200
    .TextMatrix(0, 1) = "Document Type":            .ColWidth(1) = 1000
    .TextMatrix(0, 2) = "Account Type":             .ColWidth(2) = 1200
    .TextMatrix(0, 3) = "Account No.":              .ColWidth(3) = 1500
    .TextMatrix(0, 4) = "Amount":                   .ColWidth(4) = 1200
    .TextMatrix(0, 5) = "TDS Nature of Deduction":  .ColWidth(5) = 800
    .TextMatrix(0, 6) = "Service Tax Group Code":   .ColWidth(6) = 1000
    .TextMatrix(0, 7) = "Narration":                .ColWidth(7) = 1500
    .TextMatrix(0, 8) = "RegionDim":                .ColWidth(8) = 1500
    .TextMatrix(0, 9) = "DeptDim":                  .ColWidth(9) = 1000
    .TextMatrix(0, 10) = "VendorDim":               .ColWidth(10) = 1000
    .TextMatrix(0, 11) = "External Document No.":   .ColWidth(11) = 1200
    .TextMatrix(0, 12) = "Invoice Date":            .ColWidth(12) = 1000
    
    ''--- Ullhas Patil 03 Oct 2017
    ''.TextMatrix(0, 13) = "Document No.":          .ColWidth(13) = 1000
    .TextMatrix(0, 13) = "Project Dim":             .ColWidth(13) = 1200
    .TextMatrix(0, 14) = "GST Group Code":          .ColWidth(14) = 1200
    .TextMatrix(0, 15) = "HSN Code":                .ColWidth(15) = 1200
    .TextMatrix(0, 16) = "Invoice Type":            .ColWidth(16) = 1200
    .TextMatrix(0, 17) = "Document No":              .ColWidth(17) = 1200
    ''--------------------------------------


End With
End Sub





