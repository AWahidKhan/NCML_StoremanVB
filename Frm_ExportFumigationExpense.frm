VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportFumigationExpense 
   Caption         =   "Export Fumigation Expense"
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
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   14955
      TabIndex        =   8
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
         Left            =   11040
         TabIndex        =   4
         Top             =   200
         Width           =   1920
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
         Height          =   360
         Left            =   13980
         TabIndex        =   6
         Top             =   200
         Width           =   855
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
         Height          =   360
         Left            =   13020
         TabIndex        =   5
         Top             =   200
         Width           =   975
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
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   75
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
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   350
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
         Left            =   2280
         TabIndex        =   2
         Top             =   200
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
         Left            =   5040
         TabIndex        =   3
         Top             =   200
         Width           =   4575
      End
      Begin VB.Label Label4 
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
         Left            =   9720
         TabIndex        =   11
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   1320
         TabIndex        =   10
         Top             =   255
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3720
         TabIndex        =   9
         Top             =   255
         Width           =   1185
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8775
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15478
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Frm_ExportFumigationExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExpFlag, MaxID As Variant
Dim mForwardToID As Double
Dim mJournalBatchName, mJournalType As Variant
Dim mDocNo As Variant

Dim mFumigationExpenseID As Double
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
''-- On 05 Aug 2016
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
mInvoiceType = "FUMIGATION"

If Option1.Value = True Then
   strFilName = "\Navision\VendorNavUpload_SMFumigation" & MaxID & ".csv"
Else
   strFilName = "\Navision\VendorNavUpload_SMFumigation" & CmbBatch.Text & ".csv"
End If

''-- On 08 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
    ''----- Ullhas Patil 03 Oct 2017
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
    Call TableTxn_SecurityBillForNavision(" Where G_UID is null ")
    ''--------------------------
End If

With MSHFlexGrid1
For I = 1 To MSHFlexGrid1.Rows - 1
    tmp = ""
    If Option1.Value = True Then
       If .TextMatrix(I, 0) <> "" Then
        ''----- Ullhas Patil 03 Oct 2017
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
'        RsTxn_SecurityBillForNavision!DocumentNo = Trim(.TextMatrix(i, 13))
        
        
        RsTxn_SecurityBillForNavision!ProjectDim = Trim(.TextMatrix(I, 13))
        RsTxn_SecurityBillForNavision!GSTGroupCode = Trim(.TextMatrix(I, 14))
        RsTxn_SecurityBillForNavision!HSNNo = Trim(.TextMatrix(I, 15))
        RsTxn_SecurityBillForNavision!AccType = Trim(.TextMatrix(I, 16))
        RsTxn_SecurityBillForNavision!DocumentNo = Trim(.TextMatrix(I, 17))
        
        RsTxn_SecurityBillForNavision!UserBatchID = mUserBatchID
        
        'RsTxn_SecurityBillForNavision!DataTransfer = "X"
        
        RsTxn_SecurityBillForNavision.Update
        
'        For c = 0 To MSHFlexGrid1.Cols - 1
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
'   tmp = "Update Txn_VendorClaim Set BatchID = " & MaxID & " Where BatchID = 0 and  Flag = 'A'"
'   Call TmpTable
   If CmbApprovedBy.Text = "" Then
      tmp = "Update Txn_FumigationApproval set BatchID = " & MaxID & " Where (BatchID = 0 Or BatchID is null) and  Flag = 'A'"
   Else
      tmp = "Update Txn_FumigationApproval set BatchID = " & MaxID & " Where (BatchID = 0 Or BatchID is null) and  Flag= 'A' And ForwardTo = " & mForwardToID & ""
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

''    tmp = " Select TFG.ApprovalDate , 'Invoice' As Inv, 'G/L Account' As GLAccount, "
''    tmp = tmp & " 7060201000000 'ExpenseTypeNAVCode', TFG.Amount, '' As TDS , '' 'ServiceTaxNAVCode', 'Being expenses incurred against Storeman Batch ID " & MaxID & "' As 'ApprovalRemark' ,"
''    tmp = tmp & " ML.NAVCode 'LocationNAVCode', 'CMP' As DeptDim ,  '' 'EmployeeNAVCode', TFG.BatchID, TFG.FumigationDate, TFG.FumigationID"
''    tmp = tmp & " From Txn_Fumigation TFG"
''    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TFG.GodownID"
''    tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
''    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
''    tmp = tmp & " Where (TFG.BatchID = 0 or TFG.BatchID is null) and  TFG.Status = 'A'"
''    tmp = tmp & " Order By TFG.ApprovalDate"

    tmp = "Select TFG.ApprovedDate , 'Invoice' As Inv, 'G/L Account' As GLAccount,"
    tmp = tmp & " 7060201000000 'ExpenseTypeNAVCode', TFG.ActualAmount, '' As TDS , MS.NAVCode as 'ServiceTaxNAVCode', 'Being expenses incurred against Storeman Batch ID " & MaxID & "' As 'ApprovalRemark' ,"
    tmp = tmp & " ML.NAVCode 'LocationNAVCode', 'CMP' As DeptDim ,  '' 'EmployeeNAVCode', TFG.BatchID, TFG.BillDate 'MappingDate', TFG.MappingID,TFG.BillNo,MS.Description"
    tmp = tmp & " From Txn_FumigationApproval TFG"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TFG.LocationID"
    tmp = tmp & " Inner Join Mst_ServiceTax MS on MS.ServiceTaxID = TFG.ServiceTaxID"
    tmp = tmp & " Where (TFG.BatchID = 0 or TFG.BatchID is null) and  TFG.Flag in ('A')"
    If CmbApprovedBy.Text <> "" Then
       tmp = tmp & " And TFG.ForwardTo=" & mForwardToID & ""
    End If
    tmp = tmp & " Order By TFG.ApprovedDate"

Else
    If CmbBatch.Text = "" Then
        MsgBox "Select Batch ID !!!! "
        Exit Sub
    End If
    
''    tmp = " Select TFG.ApprovalDate , 'Invoice' As Inv, 'G/L Account' As GLAccount, "
''    tmp = tmp & " 7060201000000 'ExpenseTypeNAVCode', TFG.Amount, '' As TDS , '' 'ServiceTaxNAVCode', 'Being expenses incurred against Storeman Batch ID " & CmbBatch.Text & "' as 'ApprovalRemark' ,"
''    tmp = tmp & " ML.NAVCode 'LocationNAVCode', 'CMP' As DeptDim ,  '' 'EmployeeNAVCode', TFG.BatchID, TFG.FumigationDate, TFG.FumigationID"
''    tmp = tmp & " From Txn_Fumigation TFG"
''    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TFG.GodownID"
''    tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
''    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
''    tmp = tmp & " Where TFG.BatchID =  " & Val(CmbBatch.Text) & " and  TFG.Status = 'A'"
''    tmp = tmp & " Order By TFG.ApprovalDate "

    tmp = "Select TFG.ApprovedDate , 'Invoice' As Inv, 'G/L Account' As GLAccount,"
    tmp = tmp & " 7060201000000 'ExpenseTypeNAVCode', TFG.ActualAmount, '' As TDS , MS.NAVCode as  'ServiceTaxNAVCode', 'Being expenses incurred against Storeman Batch ID " & MaxID & "' As 'ApprovalRemark' ,"
    tmp = tmp & " ML.NAVCode 'LocationNAVCode', 'CMP' As DeptDim ,  '' 'EmployeeNAVCode', TFG.BatchID, TFG.MappingDate, TFG.MappingID,TFG.BillNo,MS.Description"
    tmp = tmp & " From Txn_FumigationApproval TFG"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TFG.LocationID"
    tmp = tmp & " Inner Join Mst_ServiceTax MS on MS.ServiceTaxID = TFG.ServiceTaxID"
    tmp = tmp & " Where TFG.BatchID =  " & Val(CmbBatch.Text) & " And TFG.Flag in ('S', 'A')"
    tmp = tmp & " Order By TFG.ApprovedDate"

End If
Call TmpTable

If Option1.Value = True Then
''
''    tmp = " Select TFG.ApprovalDate, 'Invoice' As Inv, 'Vendor' As GLAccount, 7060201000000 'ExpenseTypeNAVCode', TFG.Amount,"
''    tmp = tmp & " T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',  TFG.ApprovalRemark + ' - Fumigation ID:' + convert(varchar(10),TFG.FumigationID) As 'ApprovalRemarks',"
''    tmp = tmp & " '201001' 'LocationNAVCode', 'CMP' As DeptDim ,  E.NAVCode 'VendorNAVCode',"
''    tmp = tmp & " TFG.BatchID ,'' As InvoiceDate, '' InvoiceNo , TFG.FumigationID"
''    tmp = tmp & " From  Txn_Fumigation TFG"
''    tmp = tmp & " Inner Join Mst_Vendor E on TFG.VendorID = E.VendorID"
''    tmp = tmp & " Inner Join Mst_TDS T on E.TDSID = T.TDSID"
''    tmp = tmp & " Where (TFG.BatchID = 0 Or TFG.BatchID is null) and  TFG.Status = 'A' "
''    'If CmbApprovedBy.Text <> "" Then
''    '   tmp = tmp & " And VC.ForwardToID=" & mForwardToID & ""
''    'End If
''    tmp = tmp & " Order By TFG.ApprovalDate"

    tmp = "Select TFG.ApprovedDate, 'Invoice' As Inv, 'Vendor' As GLAccount,MVGST.GSTNavCode 'ExpenseTypeNAVCode', TFG.ActualAmount," ''-- 7060201000000 'ExpenseTypeNAVCode'
    tmp = tmp & " T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',  TFG.ApprovalRemark + ' - Fumigation Map ID:' + convert(varchar(10),TFG.MappingID) As 'ApprovalRemarks',"
    tmp = tmp & " '201001' 'LocationNAVCode', 'CMP' As DeptDim ,  E.NAVCode 'VendorNAVCode',"
    tmp = tmp & " TFG.BatchID ,'' As InvoiceDate, TFG.BillNo as InvoiceNo, TFG.MappingID,TFG.BillDate 'MappingDate',TFG.BillingStateIndiaID"
    tmp = tmp & " From  Txn_FumigationApproval TFG"
    tmp = tmp & " Inner Join Mst_Vendor E on TFG.VendorID = E.VendorID"
    tmp = tmp & " Inner Join Mst_TDS T on E.TDSID = T.TDSID"
    
    ''-- Ullhas Patil 03 Oct 2017
    tmp = tmp & " Inner Join Mst_Location ML on TFG.LocationID = ML.LocationID"
    'tmp = tmp & " LEft Join Mst_State MS On BillingStateIndiaID.StateID = MS.StateID"
    tmp = tmp & " Left Join Mst_VendorStateWiseGST MVGST On TFG.BillingStateIndiaID = MVGST.StateIndiaID And TFG.VendorID = MVGST.VendorID"
    ''---------
    
    tmp = tmp & " Where (TFG.BatchID = 0 Or TFG.BatchID is null) and TFG.Flag in ('A')"
    If CmbApprovedBy.Text <> "" Then
       tmp = tmp & " And TFG.ForwardTo=" & mForwardToID & ""
    End If
    tmp = tmp & " Order By TFG.ApprovedDate"

Else
    If CmbBatch.Text = "" Then
        MsgBox "Select Batch ID !!!"
        Exit Sub
    End If
    
''    tmp = " Select TFG.ApprovalDate, 'Invoice' As Inv, 'Vendor' As GLAccount, 7060201000000 'ExpenseTypeNAVCode', TFG.Amount,"
''    tmp = tmp & " T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',  TFG.ApprovalRemark + ' - Fumigation ID:' + convert(varchar(10),TFG.FumigationID) As 'ApprovalRemarks',"
''    tmp = tmp & " '201001' 'LocationNAVCode', 'CMP' As DeptDim ,  E.NAVCode 'VendorNAVCode',"
''    tmp = tmp & " TFG.BatchID ,'' As InvoiceDate, '' InvoiceNo , TFG.FumigationID"
''    tmp = tmp & " From  Txn_Fumigation TFG"
''    tmp = tmp & " Inner Join Mst_Vendor E on TFG.VendorID = E.VendorID"
''    tmp = tmp & " Inner Join Mst_TDS T on E.TDSID = T.TDSID"
''    tmp = tmp & " Where TFG.BatchID =  " & Val(CmbBatch.Text) & " and  TFG.Status = 'A'"
''    tmp = tmp & " Order By TFG.ApprovalDate "

    tmp = "Select TFG.ApprovedDate, 'Invoice' As Inv, 'Vendor' As GLAccount, MVGST.GSTNavCode 'ExpenseTypeNAVCode', TFG.ActualAmount," ''-- 7060201000000 'ExpenseTypeNAVCode'
    tmp = tmp & " T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',  TFG.ApprovalRemark + ' - Fumigation Map ID:' + convert(varchar(10),TFG.MappingID) As 'ApprovalRemarks',"
    tmp = tmp & " '201001' 'LocationNAVCode', 'CMP' As DeptDim ,  E.NAVCode 'VendorNAVCode',"
    tmp = tmp & " TFG.BatchID ,'' As InvoiceDate, TFG.BillNo as InvoiceNo , TFG.MappingID,TFG.MappingDate,TFG.BillingStateIndiaID"
    tmp = tmp & " From  Txn_FumigationApproval TFG"
    tmp = tmp & " Inner Join Mst_Vendor E on TFG.VendorID = E.VendorID"
    tmp = tmp & " Inner Join Mst_TDS T on E.TDSID = T.TDSID"
    
    ''-- Ullhas Patil 03 Oct 2017
'    tmp = tmp & " Inner Join Mst_Location ML on TFG.LocationID = ML.LocationID"
'    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Left Join Mst_VendorStateWiseGST MVGST On TFG.BillingStateIndiaID = MVGST.StateIndiaID And TFG.VendorID = MVGST.VendorID"
    ''---------
    
    tmp = tmp & " Where TFG.BatchID =  " & Val(CmbBatch.Text) & " and  TFG.Flag in ('S', 'A')"
    tmp = tmp & " Order By TFG.ApprovedDate"

End If
Call Tmp1Table
Dim mDat_ As Variant
If TmpRs.RecordCount > 0 And TmpRs1.RecordCount > 0 Then
    Dim x As Variant
    x = 0
    ExpFlag = "Y"
    mDat_ = IIf(IsNull(TmpRs!ApprovedDate), "", TmpRs!ApprovedDate)
    For x = 1 To TmpRs.RecordCount
        If TmpRs!ExpenseTypeNAVCode = "" Or IsNull(TmpRs!ExpenseTypeNAVCode) Or TmpRs!LocationNAVCode = "" Or IsNull(TmpRs!LocationNAVCode) Then
           ExpFlag = "N"
        End If
        I = MSHFlexGrid1.Rows - 1
        'For c = 0 To MSHFlexGrid1.Cols - 1
        '    MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
        'Next
        MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!ApprovedDate), "", TmpRs!ApprovedDate)
        MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!Inv), "", TmpRs!Inv)
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!GLAccount), "", TmpRs!GLAccount)
        
        ''-- Ullhas Patil 29 Sep 2017
        ''MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(TmpRs!ExpenseTypeNAVCode), "", TmpRs!ExpenseTypeNAVCode)
        If LCase(TmpRs!Description) = LCase("GST with Tax") Then
            MSHFlexGrid1.TextMatrix(I, 3) = GetGLAccNAVCodeFromExpenseType(mFumigationExpenseID, "NAGRI", "G")
        ElseIf LCase(TmpRs!Description) = LCase("GST without Tax") Then
            MSHFlexGrid1.TextMatrix(I, 3) = GetGLAccNAVCodeFromExpenseType(mFumigationExpenseID, "AGRI", "G")
        End If
        '''---------------
        
        MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!ActualAmount), 0, TmpRs!ActualAmount)
        MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!TDS), "", TmpRs!TDS)
        
        ''-- Ullhas Patil 29 Sep 2017
        ''MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(TmpRs!ServiceTaxNAVCode), "", TmpRs!ServiceTaxNAVCode)
        MSHFlexGrid1.TextMatrix(I, 6) = ""
        ''----------------
        
        MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!ApprovalRemark), "", TmpRs!ApprovalRemark)
        MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!LocationNAVCode), "", TmpRs!LocationNAVCode)
        MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!DeptDim), "", TmpRs!DeptDim)
        'MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!EmployeeNAVCode), "", TmpRs!EmployeeNAVCode)
        MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!MappingDate), "", TmpRs!MappingDate)
        
'        MSHFlexGrid1.TextMatrix(i, 0) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 0))
        MSHFlexGrid1.TextMatrix(I, 10) = "EMP-1569"
        'MSHFlexGrid1.TextMatrix(i, 12) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 12))
        
        If Option1.Value = True Then  ' If CmbBatch.Text = "" Then
           MSHFlexGrid1.TextMatrix(I, 11) = "SMF" & Trim(MaxID) & "-" & TmpRs!MappingID & "-" & IIf(IsNull(TmpRs!BillNo), "", "Inv^" & TmpRs!BillNo)
        Else
           If CmbBatch.Text > 164 Then
              MSHFlexGrid1.TextMatrix(I, 11) = "SMF" & Trim(CmbBatch.Text) & "-" & TmpRs!MappingID & "-" & IIf(IsNull(TmpRs!BillNo), "", "Inv^" & TmpRs!BillNo)
           Else
              MSHFlexGrid1.TextMatrix(I, 11) = "SMF" & Trim(CmbBatch.Text) & "-" & TmpRs!MappingID & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
           End If
        End If
        mDocNo = "DocNo" & TmpRs!MappingID
        ''-- Ullhas Patil 29 Sep 2017
        ''MSHFlexGrid1.TextMatrix(i, 13) = mDocNo
        MSHFlexGrid1.TextMatrix(I, 13) = ""
        MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!ServiceTaxNAVCode), "", TmpRs!ServiceTaxNAVCode)
        
        If LCase(TmpRs!Description) = LCase("GST with Tax") Then
            MSHFlexGrid1.TextMatrix(I, 15) = GetGLAccNAVCodeFromExpenseType(mFumigationExpenseID, "NAGRI", "H")
        ElseIf LCase(TmpRs!Description) = LCase("GST without Tax") Then
            MSHFlexGrid1.TextMatrix(I, 15) = GetGLAccNAVCodeFromExpenseType(mFumigationExpenseID, "AGRI", "H")
        End If

        MSHFlexGrid1.TextMatrix(I, 16) = ""
        ''----------------------
        mDocumentNo = "DocNo-" & TmpRs!MappingID
        MSHFlexGrid1.TextMatrix(I, 17) = mDocumentNo

        If MSHFlexGrid1.TextMatrix(I, 14) = "" Or MSHFlexGrid1.TextMatrix(I, 15) = "" Then
            ExpFlag = "N"
        End If
        
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        mDat_ = TmpRs!MappingID
        
        TmpRs.MoveNext
        If TmpRs.EOF Then
           CrEntry (mDat_)
           Exit For
        End If
        If mDat_ <> IIf(IsNull(TmpRs!MappingID), "", TmpRs!MappingID) Then
           Call CrEntry(mDat_)
           mDat_ = IIf(IsNull(TmpRs!MappingID), "", TmpRs!MappingID)
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
TmpRs1.Find "MappingID = " & mPDat_ 'TmpRs!ClaimID
If TmpRs1.EOF = False Then
   If TmpRs1!ExpenseTypeNAVCode = "" Or IsNull(TmpRs1!ExpenseTypeNAVCode) Or TmpRs1!LocationNAVCode = "" Or IsNull(TmpRs1!LocationNAVCode) Or TmpRs1!VendorNAVCode = "" Or IsNull(TmpRs1!VendorNAVCode) Then
      ExpFlag = "N"
   End If
   I = MSHFlexGrid1.Rows - 1
    'For c = 0 To MSHFlexGrid1.Cols - 1
    '    MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs1.Fields(c)), "", TmpRs1.Fields(c))
    'Next
   MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!ApprovedDate), "", TmpRs1!ApprovedDate)
   MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!Inv), "", TmpRs1!Inv)
   MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!GLAccount), "", TmpRs1!GLAccount)
   MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!ExpenseTypeNAVCode), "", TmpRs1!ExpenseTypeNAVCode) '' IIf(IsNull(TmpRs1!VendorNAVCode), "", TmpRs1!VendorNAVCode)
   MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!ActualAmount), 0, TmpRs1!ActualAmount)
   MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!TDS), "", TmpRs1!TDS)
   MSHFlexGrid1.TextMatrix(I, 6) = "" ''IIf(IsNull(TmpRs1!ServiceTaxNAVCode), "", TmpRs1!ServiceTaxNAVCode)
   MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!ApprovalRemarks), "", TmpRs1!ApprovalRemarks)
   MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!LocationNAVCode), "", TmpRs1!LocationNAVCode)
   MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!DeptDim), "", TmpRs1!DeptDim)
   'MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!FumigationDate), "", TmpRs1!FumigationDate)
        
'   MSHFlexGrid1.TextMatrix(i, 0) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 0))
   MSHFlexGrid1.TextMatrix(I, 4) = Val(MSHFlexGrid1.TextMatrix(I, 4)) * -1
   MSHFlexGrid1.TextMatrix(I, 10) = "EMP-1569"
  ' MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!InvoiceNo), "", "Invoice " & TmpRs1!InvoiceNo) '"SMV" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
   If Option1.Value = True Then  ' If CmbBatch.Text = "" Then
      MSHFlexGrid1.TextMatrix(I, 11) = "SMF" & Trim(MaxID) & "-" & TmpRs1!MappingID & "-" & IIf(IsNull(TmpRs1!InvoiceNo), "", "Inv^" & TmpRs1!InvoiceNo)
   Else
      If CmbBatch.Text > 164 Then
         MSHFlexGrid1.TextMatrix(I, 11) = "SMF" & Trim(CmbBatch.Text) & "-" & TmpRs1!MappingID & "-" & IIf(IsNull(TmpRs1!InvoiceNo), "", "Inv^" & TmpRs1!InvoiceNo)
      Else
         MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!InvoiceNo), "", "Invoice " & TmpRs1!InvoiceNo & "-" & TmpRs1!MappingID)  '"SMF" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
      End If
   End If
        
   MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!MappingDate), "", TmpRs1!MappingDate)
   'MSHFlexGrid1.TextMatrix(i, 12) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 12))
   ''MSHFlexGrid1.TextMatrix(i, 13) = mDocNo
   
   MSHFlexGrid1.TextMatrix(I, 13) = ""
   MSHFlexGrid1.TextMatrix(I, 14) = ""
   MSHFlexGrid1.TextMatrix(I, 15) = ""
   MSHFlexGrid1.TextMatrix(I, 16) = ""
   MSHFlexGrid1.TextMatrix(I, 17) = mDocumentNo
   
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1

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
Picture1.Enabled = True
MSHFlexGrid1.Enabled = True

''-- Ullhas Patil 29 Sep 2017
mFumigationExpenseID = 18
''--------------

Call Grid_Head
tmp = "Select Distinct BatchID from Txn_FumigationApproval Where BatchID > 0  order by BatchID"
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
With MSHFlexGrid1
    CmdExport.Enabled = False
    .Clear
    .Rows = 2
    .Cols = 18 '17 ''14
    .TextMatrix(0, 0) = "Posting Date":                 .ColWidth(0) = 1200
    .TextMatrix(0, 1) = "Document Type":                .ColWidth(1) = 1000
    .TextMatrix(0, 2) = "Account Type":                 .ColWidth(2) = 1200
    .TextMatrix(0, 3) = "Account No.":                  .ColWidth(3) = 1500
    .TextMatrix(0, 4) = "Amount":                       .ColWidth(4) = 1200
    .TextMatrix(0, 5) = "TDS Nature of Deduction":      .ColWidth(5) = 800
    .TextMatrix(0, 6) = "Service Tax Group Code":       .ColWidth(6) = 800
    .TextMatrix(0, 7) = "Narration":                    .ColWidth(7) = 1500
    .TextMatrix(0, 8) = "RegionDim":                    .ColWidth(8) = 1000
    .TextMatrix(0, 9) = "DeptDim":                      .ColWidth(9) = 1000
    .TextMatrix(0, 10) = "VendorDim":                   .ColWidth(10) = 1000
    .TextMatrix(0, 11) = "External Document No.":       .ColWidth(11) = 1200
    .TextMatrix(0, 12) = "Invoice Date":                .ColWidth(12) = 1000
    ''---- Ullhas Patil 29 Sep 2017
    ''.TextMatrix(0, 13) = "Document No.":                .ColWidth(13) = 1000
    .TextMatrix(0, 13) = "Project Dim":                 .ColWidth(13) = 1000
    .TextMatrix(0, 14) = "GSTGroup Code":               .ColWidth(14) = 1000
    .TextMatrix(0, 15) = "HSN Code":                    .ColWidth(15) = 1000
    .TextMatrix(0, 16) = "Invoice Type":                .ColWidth(16) = 1000
    .TextMatrix(0, 17) = "Document No":                 .ColWidth(17) = 1800
End With
End Sub
