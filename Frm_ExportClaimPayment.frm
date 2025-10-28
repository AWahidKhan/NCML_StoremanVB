VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportClaimPayment 
   Caption         =   "Export Claim Payment"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   30
      Width           =   14985
      Begin VB.OptionButton OptFranchiseeRent 
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
         Left            =   6600
         TabIndex        =   26
         Top             =   240
         Width           =   210
      End
      Begin VB.OptionButton optFumigation 
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
         Left            =   4200
         TabIndex        =   2
         Top             =   233
         Width           =   210
      End
      Begin VB.OptionButton OptGodownRent 
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
         Top             =   240
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.OptionButton OptSecurity 
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
         Left            =   2355
         TabIndex        =   1
         Top             =   233
         Width           =   210
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Franchisee Rent"
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
         Left            =   6885
         TabIndex        =   27
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fumigation Expense"
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
         Left            =   4485
         TabIndex        =   22
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Godown Rent"
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
         Left            =   420
         TabIndex        =   15
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Security"
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
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   720
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7545
      Left            =   120
      TabIndex        =   11
      Top             =   2205
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   13309
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   3825
      Left            =   120
      TabIndex        =   20
      Top             =   9840
      Visible         =   0   'False
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6747
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   14985
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
         Left            =   2820
         TabIndex        =   28
         Top             =   1080
         Width           =   9855
      End
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
         Left            =   9960
         TabIndex        =   8
         Top             =   555
         Width           =   2760
      End
      Begin VB.ComboBox CmbModeOfPay 
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
         ItemData        =   "Frm_ExportClaimPayment.frx":0000
         Left            =   8100
         List            =   "Frm_ExportClaimPayment.frx":0010
         TabIndex        =   7
         Top             =   555
         Width           =   1845
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
         Height          =   400
         Left            =   12750
         TabIndex        =   21
         Top             =   600
         Width           =   1200
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
         Height          =   400
         Left            =   13950
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   200
         Width           =   960
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
         Height          =   400
         Left            =   12750
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   200
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
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
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.OptionButton Option2 
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
         TabIndex        =   4
         Top             =   600
         Width           =   210
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
         Left            =   1395
         TabIndex        =   5
         Top             =   555
         Width           =   1455
      End
      Begin VB.ComboBox CmbFAGBank 
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
         Left            =   2895
         TabIndex        =   6
         Top             =   555
         Width           =   5175
      End
      Begin VB.CommandButton CmdLetter 
         Caption         =   "Letter"
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
         Height          =   400
         Left            =   13950
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Payment Updated By"
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
         TabIndex        =   29
         Top             =   1140
         Width           =   1920
      End
      Begin VB.Label Label9 
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
         Left            =   10853
         TabIndex        =   25
         Top             =   278
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Payment Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8325
         TabIndex        =   24
         Top             =   285
         Width           =   1515
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
         Left            =   1755
         TabIndex        =   19
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Left            =   435
         TabIndex        =   18
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   435
         TabIndex        =   17
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "FAG Bank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5010
         TabIndex        =   16
         Top             =   270
         Width           =   1050
      End
   End
End
Attribute VB_Name = "Frm_ExportClaimPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExpFlag, MaxID, mFAGBankID As Variant
Dim mForwardToID As Double
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mPaymentLimitation, mModeOfPay, mFileNo, mVendorID, letternm As Variant

Dim mJournalBatchName, mJournalType As Variant
Dim mDocNo, mInvoiceType As Variant

Dim mpType, mpBankID As Variant
Dim mNewLetterFormatFAGBankID As Variant
Dim mAllowNewLetterFormatTobePrint As Boolean
Dim mUpdatedBy As Variant

Private Sub CmbBatch_GotFocus()
tmp1 = CmbBatch.Text

If OptGodownRent.Value = True Then
    tmp = "select distinct PBatchID from Txn_GodownRent Where PBatchID > 0  order by PBatchID"
ElseIf OptSecurity.Value = True Then
    tmp = "select distinct PBatchID from Txn_SecurityBilling  Where PBatchID > 0  order by PBatchID"
ElseIf optFumigation.Value = True Then
    tmp = "select distinct PBatchID from Txn_FumigationApproval Where PBatchID > 0  order by PBatchID"
ElseIf OptFranchiseeRent.Value = True Then
    tmp = "select distinct PBatchID from Txn_FranchiseeRent Where PBatchID > 0  order by PBatchID"
End If
Call Tmp1Table

CmbBatch.Clear
If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
       CmbBatch.AddItem TmpRs1.Fields(0)
       MaxID = TmpRs1.Fields(0) + 1
       TmpRs1.MoveNext
   Next
Else
   MaxID = 1
End If
CmbBatch.Text = tmp1
End Sub

Private Sub CmbBatch_LostFocus()
If CmbBatch.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "PBatchID = " & Val(CmbBatch.Text) & ""
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbBatch.SetFocus
   Exit Sub
End If

'tmp = "Select MF.BankName,TA.ModeOfPay,MF.PaymentLimitation " 'MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName 'VendorBankName',MV.IFSCCode,MV.BankBranch "
'tmp = tmp & " From Txn_VendorClaim TA "
'tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID "
'tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text)

tmp = "Select Distinct MF.BankName,TA.ModeOfPay,MF.PaymentLimitation "
tmp = tmp & " From Txn_GodownRent TA"
tmp = tmp & " Inner Join Mst_Lessor MV on TA.LessorID=MV.LessorID"
tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text)
Call Tmp4Table
If TmpRs4.RecordCount > 0 Then
    mModeOfPay = IIf(IsNull(TmpRs4!ModeOfPay), "", TmpRs4!ModeOfPay)
    mPaymentLimitation = IIf(IsNull(TmpRs4!PaymentLimitation), 1000, IIf(TmpRs4!PaymentLimitation = 0, 1000, TmpRs4!PaymentLimitation))
Else
    mModeOfPay = ""
    mPaymentLimitation = ""
End If

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

Private Sub CmbModeOfPay_GotFocus()
tmp = CmbModeOfPay.Text
End Sub
Private Sub CmbModeOfPay_LostFocus()
If CmbModeOfPay.Text = "" Then Exit Sub
If OptGodownRent.Value = True Then
   If LCase(CmbModeOfPay.Text) <> "bank transfer" And LCase(CmbModeOfPay.Text) <> "cheque" And LCase(CmbModeOfPay.Text) <> "rtgs" And LCase(CmbModeOfPay.Text) <> "neft" Then
      MsgBox "invalid Selection."
      CmbModeOfPay.SetFocus
      Exit Sub
   End If
End If
If tmp <> CmbModeOfPay.Text Then
   Call Grid_Head
End If
End Sub

Private Sub CmdExport_Click()
Dim I, C As Variant
Dim Path, strFilName, strFilLocation As Variant
Dim mUserBatchID, mInvoiceType As Variant

Path = App.Path
Call CmbBatch_GotFocus
If OptGodownRent.Value = True Then mInvoiceType = "GODOWNRENT"
If OptSecurity.Value = True Then mInvoiceType = "SECURITY"
If optFumigation.Value = True Then mInvoiceType = "FUMIGATION"
If OptFranchiseeRent.Value = True Then mInvoiceType = "FRANCHISEE"
mUserBatchID = GetGUID

If Option1.Value = True Then
    If CmbJournalCode.Text = "" Then
       MsgBox "Journal Code not selected"
       CmbJournalCode.SetFocus
       Exit Sub
    End If
    mUserBatchID = GetGUID
    If Trim(mJournalType) = "" Or Trim(mJournalBatchName) = "" Then
       MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
       Exit Sub
    End If
End If

If OptGodownRent.Value = True Then
   mInvoiceType = "GODOWNRENT"
   If Option1.Value = True Then
      strFilName = "\Navision\NavUpload_GodownRentPayment" & MaxID & ".csv"
   Else
      strFilName = "\Navision\NavUpload_GodownRentPayment" & CmbBatch.Text & ".csv"
   End If
ElseIf OptSecurity.Value = True Then
   mInvoiceType = "SECURITY"
   If Option1.Value = True Then
      strFilName = "\Navision\NavUpload_SecurityBilling" & MaxID & ".csv"
   Else
      strFilName = "\Navision\NavUpload_SecurityBilling" & CmbBatch.Text & ".csv"
   End If
ElseIf optFumigation.Value = True Then
   mInvoiceType = "FUMIGATION"
   If Option1.Value = True Then
      strFilName = "\Navision\NavUpload_FumigationExpense" & MaxID & ".csv"
   Else
      strFilName = "\Navision\NavUpload_FumigationExpense" & CmbBatch.Text & ".csv"
   End If
ElseIf OptFranchiseeRent.Value = True Then
   mInvoiceType = "FRANCHISEE"
   If Option1.Value = True Then
      strFilName = "\Navision\NavUpload_FranchiseeRent" & MaxID & ".csv"
   Else
      strFilName = "\Navision\NavUpload_FranchiseeRent" & CmbBatch.Text & ".csv"
   End If
End If

If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
   Call TableTxn_SecurityBillForNavision(" Where G_UID is null ")
End If

With MSHFlexGrid1
For I = 2 To MSHFlexGrid1.Rows - 1
    tmp = ""
    If Option1.Value = True Then
       If .TextMatrix(I, 0) <> "" Then
        RsTxn_SecurityBillForNavision.AddNew
        RsTxn_SecurityBillForNavision!PostingDate = CDate(.TextMatrix(I, 0))
        RsTxn_SecurityBillForNavision!DocumentType = Trim(.TextMatrix(I, 1))
        RsTxn_SecurityBillForNavision!AccountType = Trim(.TextMatrix(I, 2))
        RsTxn_SecurityBillForNavision!AccountNo = Trim(.TextMatrix(I, 3))
        RsTxn_SecurityBillForNavision!Amount = Trim(.TextMatrix(I, 4))
        RsTxn_SecurityBillForNavision!InvoiceNo = Trim(.TextMatrix(I, 5))
        RsTxn_SecurityBillForNavision!TDSNatureofDeduction = ""
        RsTxn_SecurityBillForNavision!ServiceTaxGroupCode = ""
        RsTxn_SecurityBillForNavision!Narration = Trim(.TextMatrix(I, 6))
        RsTxn_SecurityBillForNavision!RegionDim = Trim(.TextMatrix(I, 7))
        RsTxn_SecurityBillForNavision!DeptDim = Trim(.TextMatrix(I, 8))
        RsTxn_SecurityBillForNavision!VendorDim_EmpDim = Trim(.TextMatrix(I, 9))
        RsTxn_SecurityBillForNavision!ChequeNo = Trim(.TextMatrix(I, 10))
        RsTxn_SecurityBillForNavision!InvoiceDate = IIf(Trim(.TextMatrix(I, 11)) = "", Null, Trim(.TextMatrix(I, 11)))
        RsTxn_SecurityBillForNavision!PaidDate = IIf(Trim(.TextMatrix(I, 12)) = "", Null, Trim(.TextMatrix(I, 12)))
        RsTxn_SecurityBillForNavision!PINo = Trim(.TextMatrix(I, 13))
        RsTxn_SecurityBillForNavision!InvoiceType = mInvoiceType
        RsTxn_SecurityBillForNavision!CreatedBy = Gen_UserLoginID
        RsTxn_SecurityBillForNavision!G_UID = GetGUID
        RsTxn_SecurityBillForNavision!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
        RsTxn_SecurityBillForNavision!JournalBatchType = Trim(mJournalType)
        RsTxn_SecurityBillForNavision!JournalBatchName = Trim(mJournalBatchName)
        RsTxn_SecurityBillForNavision!DocumentNo = Trim(.TextMatrix(I, 16))
        RsTxn_SecurityBillForNavision!UserBatchID = mUserBatchID
        RsTxn_SecurityBillForNavision.Update
        
       End If
    Else
        For C = 0 To MSHFlexGrid1.Cols - 2
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
   If CmbFAGBank.Text = "" Then
      If OptGodownRent.Value = True Then
         tmp = "Update Txn_GodownRent set PBatchID = " & MaxID & ",FAGPayExportedDate =Getdate() Where isnull(PBatchID,0) = 0 and Flag = 'S' And DisbuAmount>0 And Isnull(BatchID,0) <> 0 "
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And UpdatedBy = '" & mUpdatedBy & "'"
         End If
      ElseIf OptSecurity.Value = True Then
         tmp = "Update Txn_SecurityBilling set PBatchID = " & MaxID & ",FAGPayExportedDate =Getdate() Where isnull(PBatchID,0) = 0 and Status = 'S' And DisbuAmount>0 And Isnull(BatchID,0) <> 0 "
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And UpdatedBy = '" & mUpdatedBy & "'"
         End If
      ElseIf optFumigation.Value = True Then
         tmp = "Update Txn_FumigationApproval set PBatchID = " & MaxID & ",FAGPayExportedDate =Getdate() Where isnull(PBatchID,0) = 0 and Flag = 'S' And DisbuAmount>0 And Isnull(BatchID,0) <> 0 "
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And UpdatedBy = '" & mUpdatedBy & "'"
         End If
      ElseIf OptFranchiseeRent.Value = True Then
         tmp = "Update Txn_FranchiseeRent set PBatchID = " & MaxID & ",FAGPayExportedDate =Getdate() Where isnull(PBatchID,0) = 0 and Flag = 'S' And DisbuAmount>0 And Isnull(BatchID,0) <> 0 "
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And UpdatedBy = '" & mUpdatedBy & "'"
         End If
      
      End If
   Else
      If OptGodownRent.Value = True Then
         tmp = "Update Txn_GodownRent set PBatchID = " & MaxID & ",FAGPayExportedDate =Getdate() Where isnull(PBatchID,0) = 0 and  Flag = 'S' And DisbuAmount>0 And FAGBankID = " & mFAGBankID & "  And Isnull(BatchID,0) <> 0 "
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And UpdatedBy = '" & mUpdatedBy & "'"
         End If
      ElseIf OptSecurity.Value = True Then
         tmp = "Update Txn_SecurityBilling set PBatchID = " & MaxID & ",FAGPayExportedDate =Getdate() Where isnull(PBatchID,0) = 0 and  Status = 'S' And DisbuAmount>0 And FAGBankID = " & mFAGBankID & "  And Isnull(BatchID,0) <> 0 "
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And UpdatedBy = '" & mUpdatedBy & "'"
         End If
      ElseIf optFumigation.Value = True Then
         tmp = "Update Txn_FumigationApproval set PBatchID = " & MaxID & ",FAGPayExportedDate =Getdate() Where isnull(PBatchID,0) = 0 and  Flag = 'S' And DisbuAmount>0 And FAGBankID = " & mFAGBankID & "  And Isnull(BatchID,0) <> 0 "
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And UpdatedBy = '" & mUpdatedBy & "'"
         End If
      ElseIf OptFranchiseeRent.Value = True Then
         tmp = "Update Txn_FranchiseeRent set PBatchID = " & MaxID & ",FAGPayExportedDate =Getdate() Where isnull(PBatchID,0) = 0 and  Flag = 'S' And DisbuAmount>0 And FAGBankID = " & mFAGBankID & "  And Isnull(BatchID,0) <> 0 "
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And UpdatedBy = '" & mUpdatedBy & "'"
         End If
      End If
      If CmbModeOfPay.Text <> "" Then
         tmp = tmp & " And ModeOfPay= '" & CmbModeOfPay.Text & "'"
      End If
   End If
   Call TmpTable
End If

MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Close #1
Call Grid_Head
'MsgBox "Done !!!"
'Close #1
'Call Grid_Head
End Sub


Private Sub CmdGo_Click()
ExpFlag = ""
Call Grid_Head

If Option1.Value = True Then
   If CmbFAGBank.Text = "" Then
      MsgBox "Select FAG Bank Name!!"
      CmbFAGBank.SetFocus
      Exit Sub
   End If
   
   If CmbModeOfPay.Text = "" Then
      MsgBox "Select Payment Type!!"
      CmbModeOfPay.SetFocus
      Exit Sub
   End If
   
   If CmbApprovedBy.Text = "" Then
      ans = MsgBox("Are you Sure you want to get data of all claims which may be setteled by others. If 'No' please select payment updated by from list. ", vbYesNo)
      If ans = vbNo Then
         Exit Sub
      End If
   End If
   
   Call CmbBatch_GotFocus
   
   ''--- All Payment Entries are Credit / Debit wise as Informed By Shailesh in UAT and HR Pro having same so changed here too
   If LCase(CmbModeOfPay.Text) = "bank transfer" Or LCase(CmbModeOfPay.Text) = "neft" Or LCase(CmbModeOfPay.Text) = "rtgs" Or LCase(CmbModeOfPay.Text) = "cheque" Then
      
''      If OptGodownRent.Value = True Then
''      ''--- Godown Rent (Bank Trnf/NEFT/RTGS) New Batch
''         tmp = "Select PostingDate,DocumentType,AccountType,AccountNo,SUM(Amount)Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,LessorID,advanceTypeGroupID,DisbuRemark"
''         tmp = tmp & " From ("
''         tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.LessorID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-Claim ID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
''         tmp = tmp & " From Txn_GodownRent TAR"
''         tmp = tmp & " inner Join Mst_Lessor ME on TAR.LessorID=ME.LessorID"
''         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''         tmp = tmp & " Inner Join Txn_GodownRentDetail GD on TAR.ClaimID = GD.ClaimID"
''         tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
''         tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
''         tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
''         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''
''         If CmbFAGBank.Text <> "" Then
''            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
''         End If
''         If CmbModeOfPay.Text <> "" Then
''            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
''         End If
''         tmp = tmp & " Group By DisbuDate,ME.NAVCode,TAR.DisbuAmount,L.NAVCode,TAR.ClaimID,TAR.LessorID,TAR.DisbuRemark "
''         tmp = tmp & " Union All"
''         tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLMB' as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as LessorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
''         tmp = tmp & " From Txn_GodownRent TAR"
''         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''         tmp = tmp & " inner Join Mst_Lessor ME on TAR.LessorID=ME.LessorID"
''         tmp = tmp & " Inner Join Txn_GodownRentDetail GD on TAR.ClaimID = GD.ClaimID"
''         tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
''         tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
''         tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
''         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''
''         If CmbFAGBank.Text <> "" Then
''            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
''         End If
''         If CmbModeOfPay.Text <> "" Then
''            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
''         End If
''         tmp = tmp & "Group By TAR.DisbuRemark,TAR.ClaimID "
''         tmp = tmp & " ) a"
''         tmp = tmp & " Group By PostingDate,DocumentType,AccountType,AccountNo,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,LessorID,advanceTypeGroupID,DisbuRemark Order by PINo,Seq"
''
''      ElseIf OptSecurity.Value = True Then
''      ''--- security (Bank Trnf/NEFT/RTGS) New Batch
''         tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,SUM(Amount)Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,SecurityAgencyID,advanceTypeGroupID,DisbuRemark"
''         tmp = tmp & " From ("
''         tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,Max(L.NAVCode) as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.BillTxnID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.SecurityAgencyID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-BillTxnID:' +  convert(varchar,TAR.BillTxnID) as DisbuRemark"
''         tmp = tmp & " From Txn_SecurityBilling TAR"
''         tmp = tmp & " inner Join Mst_SecurityAgency ME on TAR.SecurityAgencyID=ME.AgencyID"
''         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''         tmp = tmp & " Inner Join Txn_SecurityBillingDetail GD on TAR.BillTxnID = GD.BillTxnID"
''         tmp = tmp & " Inner Join Mst_Location L on GD.LocationID = L.LocationID"
''         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Status='S' And isnull(TAR.DisbuAmount,0)>0"
''
''         If CmbFAGBank.Text <> "" Then
''            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
''         End If
''         If CmbModeOfPay.Text <> "" Then
''            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
''         End If
''         tmp = tmp & " Group By DisbuDate,ME.NAVCode,TAR.DisbuAmount,TAR.SecurityAgencyID,TAR.DisbuRemark,TAR.BillTxnID "
''         tmp = tmp & " Union All"
''         tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,Max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLMB' as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as SecurityAgencyID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark " ''TAR.DisbuRemark + '-BillingTxnID:' +  convert(varchar,TAR.BillTxnID)
''         tmp = tmp & " From Txn_SecurityBilling TAR"
''         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''         tmp = tmp & " inner Join Mst_SecurityAgency ME on TAR.SecurityAgencyID=ME.AgencyID"
''         tmp = tmp & " Inner Join Txn_SecurityBillingDetail GD on TAR.BillTxnID = GD.BillTxnID"
''         tmp = tmp & " Inner Join Mst_Location L on GD.LocationID = L.LocationID"
''         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Status='S' And isnull(TAR.DisbuAmount,0)>0"
''
''         If CmbFAGBank.Text <> "" Then
''            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
''         End If
''         If CmbModeOfPay.Text <> "" Then
''            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
''         End If
''
''         tmp = tmp & " Group By TAR.DisbuRemark,TAR.BillTxnID ) a"
''         tmp = tmp & " Group By PostingDate,DocumentType,AccountType,AccountNo,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,SecurityAgencyID,advanceTypeGroupID,DisbuRemark Order by PINo,Seq"
''
''      ElseIf optFumigation.Value = True Then
''      ''--- Fumigation (Bank Trnf/NEFT/RTGS) New Batch
''         tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,Sum(Amount)Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark"
''         tmp = tmp & " From ("
''         tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.MappingID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.VendorID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-MappingID:' +  convert(varchar,TAR.MappingID) as DisbuRemark"
''         tmp = tmp & " From Txn_FumigationApproval TAR"
''         tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
''         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''         tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
''         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''
''         If CmbFAGBank.Text <> "" Then
''            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
''         End If
''         If CmbModeOfPay.Text <> "" Then
''            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
''         End If
''
''         tmp = tmp & " Union All"
''         tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,"
''         tmp = tmp & " max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,"
''         tmp = tmp & " 'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,"
''         tmp = tmp & " '201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLMB' as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as VendorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
''         tmp = tmp & " From Txn_FumigationApproval TAR"
''         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''         tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
''         tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
''         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''
''         If CmbFAGBank.Text <> "" Then
''             tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
''         End If
''         If CmbModeOfPay.Text <> "" Then
''            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
''         End If
''         tmp = tmp & " Group By TAR.DisbuRemark,TAR.MappingID "
''         tmp = tmp & " ) a"
''         tmp = tmp & " Group By PostingDate,DocumentType,AccountType,AccountNo ,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark Order by PINo,Seq"
''      ElseIf OptFranchiseeRent.Value = True Then
''      ''--- FranchiseeRent (Bank Trnf/NEFT/RTGS) New Batch
''         tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,Sum(Amount)Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark"
''         tmp = tmp & " From ("
''         tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.VendorID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-ClaimID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
''         tmp = tmp & " From Txn_FranchiseeRent TAR"
''         tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
''         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''         tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
''         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''
''         If CmbFAGBank.Text <> "" Then
''            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
''         End If
''         If CmbModeOfPay.Text <> "" Then
''            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
''         End If
''
''         tmp = tmp & " Union All"
''         tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,"
''         tmp = tmp & " max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,"
''         tmp = tmp & " 'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,"
''         tmp = tmp & " '201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLMB' as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as VendorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
''         tmp = tmp & " From Txn_FranchiseeRent TAR"
''         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''         tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
''         tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
''         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''
''         If CmbFAGBank.Text <> "" Then
''             tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
''         End If
''         If CmbModeOfPay.Text <> "" Then
''            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
''         End If
''         tmp = tmp & " Group By TAR.DisbuRemark,TAR.ClaimID "
''         tmp = tmp & " ) a"
''         tmp = tmp & " Group By PostingDate,DocumentType,AccountType,AccountNo ,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark Order by PINo,Seq"
''      End If
''   Else

      If OptGodownRent.Value = True Then
      ''--- GodownRent (Bank Trnf/NEFT/RTGS/Cheque) New Batch
         tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,LessorID, advanceTypeGroupID,DisbuRemark"
         tmp = tmp & " From ("
         tmp = tmp & " Select Distinct DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,ML.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.LessorID, 0 as AdvanceTypeGroupID,TAR.DisbuRemark + '-Claim ID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
         tmp = tmp & " From Txn_GodownRent TAR"
         tmp = tmp & " inner Join Mst_Lessor ME on TAR.LessorID=ME.LessorID"
         tmp = tmp & " Inner Join Txn_GodownRentDetail GD on TAR.ClaimID = GD.ClaimID"
         tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
         tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
         tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID = ML.LocationID"
         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
         
         If CmbFAGBank.Text <> "" Then
            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
         End If
         If CmbModeOfPay.Text <> "" Then
            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
         End If
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And TAR.UpdatedBy = '" & mUpdatedBy & "'"
         End If
         
         tmp = tmp & " Union All"
         tmp = tmp & " Select Distinct DisbuDate as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,MFG.NAVCode As AccountNo ,-isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached employees.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,TAR.LessorID,0 as AdvanceTypeGroupID,TAR.DisbuRemark + '-Claim ID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
         tmp = tmp & " From Txn_GodownRent TAR"
         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
         tmp = tmp & " inner Join Mst_Lessor ME on TAR.LessorID=ME.LessorID"
         tmp = tmp & " Inner Join Txn_GodownRentDetail GD on TAR.ClaimID = GD.ClaimID"
         tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
         tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
         tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"

         If CmbFAGBank.Text <> "" Then
            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
         End If
         If CmbModeOfPay.Text <> "" Then
            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
         End If
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And TAR.UpdatedBy = '" & mUpdatedBy & "'"
         End If
         tmp = tmp & " ) a"
         tmp = tmp & " Order by PINo,Seq"
      
      ElseIf OptSecurity.Value = True Then
      ''--- Security (Bank Trnf/NEFT/RTGS/Cheque) New Batch
         tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,SecurityAgencyID,advanceTypeGroupID,DisbuRemark,BillTxnID"
         tmp = tmp & " From ("
         tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,Max(L.NAVCode) as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.BillTxnID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.SecurityAgencyID, 0 as AdvanceTypeGroupID,TAR.DisbuRemark + '-BillTxnID :' +  convert(varchar,TAR.BillTxnID) as DisbuRemark,TAR.BillTxnID"
         tmp = tmp & " From Txn_SecurityBilling TAR"
         tmp = tmp & " inner Join Mst_SecurityAgency ME on TAR.SecurityAgencyID=ME.SecurityAgencyID"
         tmp = tmp & " Inner Join Txn_SecurityBillingDetail GD on TAR.BillTxnID = GD.BillTxnID"
         tmp = tmp & " Inner Join Mst_Location L on GD.LocationID = L.LocationID"
         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0 And TAR.Status='S' And isnull(TAR.DisbuAmount,0)>0 "
         If CmbFAGBank.Text <> "" Then
            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
         End If
         If CmbModeOfPay.Text <> "" Then
            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
         End If
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And TAR.UpdatedBy = '" & mUpdatedBy & "'"
         End If
         tmp = tmp & " Group By TAR.DisbuDate,ME.NAVCode,TAR.BillTxnID,TAR.SecurityAgencyID,TAR.DisbuRemark,TAR.DisbuAmount"
         tmp = tmp & " Union All"
         tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,max(MFG.NAVCode) As AccountNo ,-(isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached employees.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.BillTxnID) as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as VendorID,0 as AdvanceTypeGroupID,TAR.DisbuRemark + '-Bill TxnID:' +  convert(varchar,TAR.BillTxnID) as DisbuRemark,TAR.BillTxnID"
         tmp = tmp & " From Txn_SecurityBilling TAR"
         tmp = tmp & " Inner Join Mst_SecurityAgency ME on TAR.SecurityAgencyID=ME.SecurityAgencyID"
         tmp = tmp & " Inner Join Txn_SecurityBillingDetail GD on TAR.BillTxnID = GD.BillTxnID"
         tmp = tmp & " Inner Join Mst_Location L on GD.LocationID = L.LocationID"
         tmp = tmp & " Inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Status='S' And isnull(TAR.DisbuAmount,0)>0 "
         If CmbFAGBank.Text <> "" Then
            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
         End If
         If CmbModeOfPay.Text <> "" Then
            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
         End If
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And TAR.UpdatedBy = '" & mUpdatedBy & "'"
         End If
         tmp = tmp & " Group By TAR.DisbuRemark,TAR.BillTxnID,TAR.DisbuAmount) a"
         tmp = tmp & " Order by PINo,Seq"
      ElseIf optFumigation.Value = True Then
      ''--- Fumigation (Bank Trnf/NEFT/RTGS/Cheque) New Batch
         tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark"
         tmp = tmp & " From ("
         tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.MappingID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.VendorID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-MappingID:' +  convert(varchar,TAR.MappingID) as DisbuRemark"
         tmp = tmp & " From Txn_FumigationApproval TAR"
         tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
         tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
         
         If CmbFAGBank.Text <> "" Then
            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
         End If
           
         If CmbModeOfPay.Text <> "" Then
            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
         End If
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And TAR.UpdatedBy = '" & mUpdatedBy & "'"
         End If
            
         tmp = tmp & " Union All"
         tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,"
         tmp = tmp & " max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,"
         tmp = tmp & " 'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,"
         tmp = tmp & " '201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.MappingID) as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as VendorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
         tmp = tmp & " From Txn_FumigationApproval TAR"
         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
         tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
         tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
            
         If CmbFAGBank.Text <> "" Then
            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
         End If
         If CmbModeOfPay.Text <> "" Then
            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
         End If
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And TAR.UpdatedBy = '" & mUpdatedBy & "'"
         End If
         
         tmp = tmp & " Group By TAR.DisbuRemark,TAR.MappingID "
         tmp = tmp & " ) a"
         tmp = tmp & " Group By PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark Order by PINo,Seq"
      
      ElseIf OptFranchiseeRent.Value = True Then
      ''--- FranchiseeRent (Bank Trnf/NEFT/RTGS/Cheque) New Batch
         tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark"
         tmp = tmp & " From ("
         tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.VendorID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-ClaimID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
         tmp = tmp & " From Txn_FranchiseeRent TAR"
         tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
         tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
         
         If CmbFAGBank.Text <> "" Then
            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
         End If
           
         If CmbModeOfPay.Text <> "" Then
            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
         End If
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And TAR.UpdatedBy = '" & mUpdatedBy & "'"
         End If
            
         tmp = tmp & " Union All"
         tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,"
         tmp = tmp & " max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,"
         tmp = tmp & " 'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,"
         tmp = tmp & " '201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as VendorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
         tmp = tmp & " From Txn_FranchiseeRent TAR"
         tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
         tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
         tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
         tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = 0  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
            
         If CmbFAGBank.Text <> "" Then
            tmp = tmp & " And TAR.FAGBankID = " & mFAGBankID & ""
         End If
         If CmbModeOfPay.Text <> "" Then
            tmp = tmp & " And TAR.ModeOfPay= '" & CmbModeOfPay.Text & "'"
         End If
         If CmbApprovedBy.Text <> "" Then
            tmp = tmp & " And TAR.UpdatedBy = '" & mUpdatedBy & "'"
         End If
         
         tmp = tmp & " Group By TAR.DisbuRemark,TAR.ClaimID "
         tmp = tmp & " ) a"
         tmp = tmp & " Group By PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark Order by PINo,Seq"
      
      End If
   End If
Else
    If CmbBatch.Text = "" Then
       MsgBox "Select Batch ID !!!! "
       Exit Sub
    End If
   If OptGodownRent.Value = True Then
        mpType = GetPaymentType("V", "Txn_GodownRent")
        mpBankID = GetPaymentFAGBankID("V", "Txn_GodownRent")
        ''--- Godown Rent (Bank Trnf/NEFT/RTGS) Old Batch
        If LCase(mpType) = "bank transfer" Or LCase(mpType) = "neft" Or LCase(mpType) = "rtgs" Or LCase(mpType) = "cheque" Then
''            tmp = "Select PostingDate,DocumentType,AccountType,AccountNo,SUM(Amount)Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,LessorID,advanceTypeGroupID,DisbuRemark"
''            tmp = tmp & " From ("
''            tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.LessorID, 0 as AdvanceTypeGroupID,TAR.DisbuRemark + '-Claim ID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
''            tmp = tmp & " From Txn_GodownRent TAR"
''            tmp = tmp & " inner Join Mst_Lessor ME on TAR.LessorID=ME.LessorID"
''            tmp = tmp & " Inner Join Txn_GodownRentDetail GD on TAR.ClaimID = GD.ClaimID"
''            tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
''            tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
''            tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
''            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S'"
''            tmp = tmp & " Group By DisbuDate,ME.NAVCode,TAR.DisbuAmount,L.NAVCode,TAR.ClaimID,TAR.LessorID,TAR.DisbuRemark"
''            tmp = tmp & " Union All"
''            tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached employees.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLMB' as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as LessorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
''            tmp = tmp & " From Txn_GodownRent TAR"
''            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''            tmp = tmp & " inner Join Mst_Lessor ME on TAR.LessorID=ME.LessorID"
''            tmp = tmp & " Inner Join Txn_GodownRentDetail GD on TAR.ClaimID = GD.ClaimID"
''            tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
''            tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
''            tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
''            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0 Group By TAR.DisbuRemark,TAR.ClaimID ) a"
''            tmp = tmp & " Group By PostingDate,DocumentType,AccountType,AccountNo ,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,LessorID,advanceTypeGroupID,DisbuRemark "
''            tmp = tmp & " Order by PINo,Seq"
''        Else
        ''--- Godown Rent (Cheque) Old Batch
            tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,LessorID,advanceTypeGroupID,DisbuRemark"
            tmp = tmp & " From ("
            tmp = tmp & " Select Distinct DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.LessorID, 0 as advanceTypeGroupID,TAR.DisbuRemark + '-Claim ID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
            tmp = tmp & " From Txn_GodownRent TAR"
            tmp = tmp & " inner Join Mst_Lessor ME on TAR.LessorID=ME.LessorID"
            tmp = tmp & " Inner Join Txn_GodownRentDetail GD on TAR.ClaimID = GD.ClaimID"
            tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
            tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
            tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "   And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
            tmp = tmp & " Union All"
            tmp = tmp & " Select Distinct DisbuDate as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,MFG.NAVCode As AccountNo ,-isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached employees.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,TAR.LessorID, 0 as AdvanceTypeGroupID,TAR.DisbuRemark + '-Claim ID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
            tmp = tmp & " From Txn_GodownRent TAR"
            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
            tmp = tmp & " inner Join Mst_Lessor ME on TAR.LessorID=ME.LessorID"
            tmp = tmp & " Inner Join Txn_GodownRentDetail GD on TAR.ClaimID = GD.ClaimID"
            tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
            tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
            tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "   And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
            tmp = tmp & " ) a"
            tmp = tmp & " Order by PINo,Seq"
        End If
   ElseIf OptSecurity.Value = True Then
        mpType = GetPaymentType("V", "Txn_SecurityBilling")
        mpBankID = GetPaymentFAGBankID("V", "Txn_SecurityBilling")
        ''--- Security (Bank Trnf/NEFT/RTGS) Old Batch
        If LCase(mpType) = "bank transfer" Or LCase(mpType) = "neft" Or LCase(mpType) = "rtgs" Or LCase(mpType) = "cheque" Then
''          tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,SUM(Amount)Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,SecurityAgencyID,advanceTypeGroupID,DisbuRemark"
''          tmp = tmp & " From ("
''          tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,Max(L.NAVCode) as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.BillTxnID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.SecurityAgencyID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-BillTxnID ID:' +  convert(varchar,TAR.BillTxnID) as DisbuRemark"
''          tmp = tmp & " From Txn_SecurityBilling TAR"
''          tmp = tmp & " inner Join Mst_SecurityAgency ME on TAR.SecurityAgencyID=ME.AgencyID"
''          tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''          tmp = tmp & " Inner Join Txn_SecurityBillingDetail GD on TAR.BillTxnID = GD.BillTxnID"
''          tmp = tmp & " Inner Join Mst_Location L on GD.LocationID = L.LocationID"
''          tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Status='S' And isnull(TAR.DisbuAmount,0)>0"
''          tmp = tmp & " Group By DisbuDate,ME.NAVCode,TAR.DisbuAmount,TAR.SecurityAgencyID,TAR.DisbuRemark,TAR.BillTxnID   "
''          tmp = tmp & " Union All"
''          tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,Max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLMB' as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as SecurityAgencyID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark " ''TAR.DisbuRemark + '-BillingTxnID:' +  convert(varchar,TAR.BillTxnID)
''          tmp = tmp & " From Txn_SecurityBilling TAR"
''          tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''          tmp = tmp & " inner Join Mst_SecurityAgency ME on TAR.SecurityAgencyID=ME.AgencyID"
''          tmp = tmp & " Inner Join Txn_SecurityBillingDetail GD on TAR.BillTxnID = GD.BillTxnID"
''          tmp = tmp & " Inner Join Mst_Location L on GD.LocationID = L.LocationID"
''          tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & " And TAR.Status='S' And isnull(TAR.DisbuAmount,0)>0"
''          tmp = tmp & " Group By TAR.DisbuRemark,TAR.BillTxnID ) a"
''          tmp = tmp & " group by PostingDate,DocumentType,AccountType,AccountNo,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,SecurityAgencyID,advanceTypeGroupID,DisbuRemark "
''          tmp = tmp & " Order by PINo,Seq"
''        Else
        ''--- Security (Cheque) Old Batch
          tmp = "Select PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,SecurityAgencyID,advanceTypeGroupID,DisbuRemark"
          tmp = tmp & " From ("
          tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,Max(L.NAVCode) as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.BillTxnID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.SecurityAgencyID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-BillTxnID ID:' +  convert(varchar,TAR.BillTxnID) as DisbuRemark"
          tmp = tmp & " From Txn_SecurityBilling TAR"
          tmp = tmp & " inner Join Mst_SecurityAgency ME on TAR.SecurityAgencyID=ME.SecurityAgencyID"
          tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
          tmp = tmp & " Inner Join Txn_SecurityBillingDetail GD on TAR.BillTxnID = GD.BillTxnID"
          tmp = tmp & " Inner Join Mst_Location L on GD.LocationID = L.LocationID"
          tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Status='S' And isnull(TAR.DisbuAmount,0)>0"
          tmp = tmp & " Group By TAR.DisbuDate,ME.NAVCode,TAR.BillTxnID,TAR.SecurityAgencyID,TAR.DisbuRemark,TAR.DisbuAmount "
          tmp = tmp & " Union All"
          tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,"
          tmp = tmp & " max(MFG.NAVCode) As AccountNo ,-isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,"
          tmp = tmp & " 'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,"
          tmp = tmp & " '201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM'+ convert(varchar,TAR.BillTxnID) as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as LessorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
          tmp = tmp & " From Txn_SecurityBilling TAR"
          tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
          tmp = tmp & " inner Join Mst_SecurityAgency ME on TAR.SecurityAgencyID=ME.SecurityAgencyID"
          tmp = tmp & " Inner Join Txn_SecurityBillingDetail GD on TAR.BillTxnID = GD.BillTxnID"
          tmp = tmp & " Inner Join Mst_Location L on GD.LocationID = L.LocationID"
          tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Status='S' And isnull(TAR.DisbuAmount,0)>0"
          tmp = tmp & " Group By TAR.DisbuRemark,TAR.BillTxnID,TAR.DisbuAmount"
          tmp = tmp & " ) a"
          tmp = tmp & " Order by PINo,Seq"
          
        End If
   ElseIf optFumigation.Value = True Then
        ''--- Fumigation (Bank Trnf/NEFT/RTGS) Old Batch
        mpType = GetPaymentType("V", "Txn_FumigationApproval")
        mpBankID = GetPaymentFAGBankID("V", "Txn_FumigationApproval")
        If LCase(mpType) = "bank transfer" Or LCase(mpType) = "neft" Or LCase(mpType) = "rtgs" Or LCase(mpType) = "cheque" Then
''            tmp = "Select PostingDate,DocumentType,AccountType,AccountNo,Sum(Amount)Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark"
''            tmp = tmp & " From ("
''            tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.MappingID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.VendorID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-MappingID:' +  convert(varchar,TAR.MappingID) as DisbuRemark"
''            tmp = tmp & " From Txn_FumigationApproval TAR"
''            tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
''            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''            tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
''            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''            tmp = tmp & " Union All"
''            tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLMB' as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as VendorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
''            tmp = tmp & " From Txn_FumigationApproval TAR"
''            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''            tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
''            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''            tmp = tmp & " Group By TAR.DisbuRemark,TAR.MappingID "
''            tmp = tmp & " ) a"
''            tmp = tmp & " group by PostingDate,DocumentType,AccountType,AccountNo,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark "
''            tmp = tmp & " Order by PINo,Seq"
''        Else
        ''--- Fumigation (Cheque) Old Batch
            tmp = "Select PostingDate,DocumentType,AccountType,AccountNo,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark"
            tmp = tmp & " From ("
            tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.MappingID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.VendorID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-MappingID:' +  convert(varchar,TAR.MappingID) as DisbuRemark"
            tmp = tmp & " From Txn_FumigationApproval TAR"
            tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
            tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
            tmp = tmp & " Union All"
            tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.MappingID) as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as VendorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
            tmp = tmp & " From Txn_FumigationApproval TAR"
            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
            tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
            tmp = tmp & " Group By TAR.DisbuRemark,TAR.MappingID "
            tmp = tmp & " ) a"
            tmp = tmp & " group by PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark "
            tmp = tmp & " Order by PINo,Seq"
        End If
   ElseIf OptFranchiseeRent.Value = True Then
        ''--- FranchiseeRent (Bank Trnf/NEFT/RTGS) Old Batch
        mpType = GetPaymentType("V", "Txn_FranchiseeRent")
        mpBankID = GetPaymentFAGBankID("V", "Txn_FranchiseeRent")
        If LCase(mpType) = "bank transfer" Or LCase(mpType) = "neft" Or LCase(mpType) = "rtgs" Or LCase(mpType) = "cheque" Then
''            tmp = "Select PostingDate,DocumentType,AccountType,AccountNo,Sum(Amount)Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark"
''            tmp = tmp & " From ("
''            tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.VendorID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-ClaimID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
''            tmp = tmp & " From Txn_FranchiseeRent TAR"
''            tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
''            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''            tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
''            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''            tmp = tmp & " Union All"
''            tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLMB' as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as VendorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
''            tmp = tmp & " From Txn_FranchiseeRent TAR"
''            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
''            tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
''            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
''            tmp = tmp & " Group By TAR.DisbuRemark,TAR.ClaimID "
''            tmp = tmp & " ) a"
''            tmp = tmp & " group by PostingDate,DocumentType,AccountType,AccountNo,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark "
''            tmp = tmp & " Order by PINo,Seq"
''        Else
        ''--- FranchiseeRent (Cheque) Old Batch
            tmp = "Select PostingDate,DocumentType,AccountType,AccountNo,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark"
            tmp = tmp & " From ("
            tmp = tmp & " Select DisbuDate as PostingDate,'Payment'as DocumentType,'Vendor' as AccountType,ME.NAVCode As AccountNo ,isnull(TAR.DisbuAmount,0) as Amount,'' as InvoiceNo,'Being claim paid' as Narration ,L.NAVCode as RegionDim,'' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,1 as Seq,TAR.VendorID,0 as advanceTypeGroupID,TAR.DisbuRemark + '-ClaimID:' +  convert(varchar,TAR.ClaimID) as DisbuRemark"
            tmp = tmp & " From Txn_FranchiseeRent TAR"
            tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
            tmp = tmp & " Inner Join Mst_Location L on TAR.LocationID = L.LocationID"
            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
            tmp = tmp & " Union All"
            tmp = tmp & " Select max(DisbuDate) as PostingDate,'Payment'as DocumentType,'Bank Account' as AccountType,max(MFG.NAVCode) As AccountNo ,Max(-isnull(TAR.DisbuAmount,0)) as Amount,'' as InvoiceNo,'Being amt paid towards expenses claim settlement to the attached vendors.' as Narration ,'201001' as RegionDim,'COO' as DeptDim,''InstrumentType,''InstrumentNo,''InstrumentDate,''DepositDate,'CLM' + convert(varchar,TAR.ClaimID) as PINo,'' PostInvoiceNo,Null as EntryNo,2 as Seq,0 as VendorID,0 as AdvanceTypeGroupID,'Being amt paid towards expenses claim settlement to the attached vendors.' as DisbuRemark"
            tmp = tmp & " From Txn_FranchiseeRent TAR"
            tmp = tmp & " inner Join Mst_FAGBank  MFG on TAR.FAGBankID=MFG.FAGBankID"
            tmp = tmp & " inner Join Mst_Vendor ME on TAR.VendorID=ME.VendorID"
            tmp = tmp & " Where isnull(TAR.BatchID,0) <> 0 And isnull(TAR.PBatchID,0) = " & Val(CmbBatch.Text) & "  And TAR.Flag='S' And isnull(TAR.DisbuAmount,0)>0"
            tmp = tmp & " Group By TAR.DisbuRemark,TAR.ClaimID "
            tmp = tmp & " ) a"
            tmp = tmp & " group by PostingDate,DocumentType,AccountType,AccountNo ,Amount,InvoiceNo,Narration ,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo,Seq,VendorID,advanceTypeGroupID,DisbuRemark "
            tmp = tmp & " Order by PINo,Seq"
        End If
   End If
End If

Call TmpTable
If Option1.Value = False Then CmbModeOfPay.Text = mpType

If TmpRs.RecordCount > 0 Then
With MSHFlexGrid1
   For r = 1 To TmpRs.RecordCount
       I = .Rows - 1
       .TextMatrix(I, 0) = Format(TmpRs!PostingDate, "mm/dd/yyyy")
       .TextMatrix(I, 1) = TmpRs!DocumentType
       .TextMatrix(I, 2) = TmpRs!AccountType
       .TextMatrix(I, 6) = IIf(IsNull(TmpRs!DisbuRemark), "", TmpRs!DisbuRemark)
       .TextMatrix(I, 3) = IIf(IsNull(TmpRs!AccountNo), "", TmpRs!AccountNo)
       
       If TmpRs!AccountNo = "" Or IsNull(TmpRs!AccountNo) Then ExpFlag = "N"

       .TextMatrix(I, 4) = IIf(IsNull(TmpRs!Amount), 0, TmpRs!Amount)
       .TextMatrix(I, 5) = IIf(IsNull(TmpRs!InvoiceNo), 0, TmpRs!InvoiceNo)
       .TextMatrix(I, 7) = IIf(IsNull(TmpRs!RegionDim), 0, TmpRs!RegionDim)
       
       If TmpRs!AccountNo = "" Or IsNull(TmpRs!AccountNo) Then ExpFlag = "N"
       If .TextMatrix(I, 7) = "" Then ExpFlag = "N"
       
       If .TextMatrix(I, 2) = "Bank Account" Then
          .TextMatrix(I, 8) = TmpRs!DeptDim
          If LCase(CmbModeOfPay.Text) = "bank transfer" Or LCase(CmbModeOfPay.Text) = "rtgs" Or LCase(CmbModeOfPay.Text) = "neft" Or LCase(CmbModeOfPay.Text) = "cheque" Then
'             If Option1.Value = True Then
'                .TextMatrix(I, 13) = TmpRs!PINo & MaxID
'             Else
'                .TextMatrix(I, 13) = TmpRs!PINo & CmbBatch.Text
'             End If
'          Else
             .TextMatrix(I, 13) = TmpRs!PINo
          End If
       Else
          If OptGodownRent.Value = True Then
             .TextMatrix(I, 8) = "CMP" 'GetTopSegment(TmpRs!EmployeeID)
          ElseIf OptSecurity.Value = True Then
             .TextMatrix(I, 8) = "CMP" 'GetTopVendorSegment(Replace(TmpRs!PINo, "CLM", ""))
          ElseIf optFumigation.Value = True Then
             .TextMatrix(I, 8) = "CMP" 'GetTopVendorSegment(Replace(TmpRs!PINo, "CLM", ""))
          ElseIf OptFranchiseeRent.Value = True Then
             .TextMatrix(I, 8) = "CMP"
          End If
          .TextMatrix(I, 13) = TmpRs!PINo
       End If
       
       If .TextMatrix(I, 8) = "" Then ExpFlag = "N"
       .TextMatrix(I, 9) = "EMP-1000"
       .TextMatrix(I, 14) = IIf(IsNull(TmpRs!PostInvoiceNo), "", TmpRs!PostInvoiceNo)
       .TextMatrix(I, 15) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
       
       If LCase(CmbModeOfPay.Text) = "bank transfer" Or LCase(CmbModeOfPay.Text) = "rtgs" Or LCase(CmbModeOfPay.Text) = "neft" Or LCase(CmbModeOfPay.Text) = "cheque" Then
''           If OptGodownRent.Value = True Then mDocNo = "G" & IIf(Option1.Value = True, "CLM" & MaxID, "CLM" & CmbBatch.Text)
''           If OptSecurity.Value = True Then mDocNo = "S" & IIf(Option1.Value = True, "CLM" & MaxID, "CLM" & CmbBatch.Text)
''           If optFumigation.Value = True Then mDocNo = "F" & IIf(Option1.Value = True, "CLM" & MaxID, "CLM" & CmbBatch.Text)
''           If OptFranchiseeRent.Value = True Then mDocNo = "FR" & IIf(Option1.Value = True, "CLM" & MaxID, "CLM" & CmbBatch.Text)
''       Else    ''-- Cheque
           If OptGodownRent.Value = True Then mDocNo = IIf(LCase(TmpRs!PINo) = "clm", mDocNo, "G" & TmpRs!PINo)
           If OptSecurity.Value = True Then mDocNo = IIf(LCase(TmpRs!PINo) = "clm", mDocNo, "S" & TmpRs!PINo)
           If optFumigation.Value = True Then mDocNo = IIf(LCase(TmpRs!PINo) = "clm", mDocNo, "FU" & TmpRs!PINo)
           If OptFranchiseeRent.Value = True Then mDocNo = IIf(LCase(TmpRs!PINo) = "clm", mDocNo, "FR" & TmpRs!PINo)
       End If
       .TextMatrix(I, 16) = mDocNo
       .Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
End With
Else
   ExpFlag = "N"
End If

If ExpFlag = "N" Then
   CmdExport.Enabled = False
Else
   CmdExport.Enabled = True
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MSHFlexGrid1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
'OptGodownRent

''-- Bank ID for New Letter format
mNewLetterFormatFAGBankID = Gen_ReadLetterFormatFAGBankID

Frame1.Enabled = True
Frame2.Enabled = True
MSHFlexGrid1.Enabled = True
Call Grid_Head
''tmp = "select distinct BatchID from Txn_AdvanceRequest Where BatchID > 0  order by BatchID"
''Call Tmp4Table
''CmbBatch.Clear
''If TmpRs4.RecordCount > 0 Then
''   For i = 1 To TmpRs4.RecordCount
''       CmbBatch.AddItem TmpRs4.Fields(0)
''       MaxID = TmpRs4.Fields(0) + 1
''       TmpRs4.MoveNext
''   Next
''Else
''   MaxID = 1
''End If
End Sub

Private Sub MSHFlexGrid1_Click()
'Dim mAmount As Variant
'mAmount = 0
'
'With MSHFlexGrid1
'
'     If .Col = 16 Then
'        If .TextMatrix(.Row, .Col) = "" Then Exit Sub
'        .Col = 16
'        If .TextMatrix(.Row, .Col) = strUnChecked Then
'           If .TextMatrix(.Row, 3) <> "" Then
'              If Val(.TextMatrix(.Row, 4)) < 0 Then
'                 mPINo = .TextMatrix(.Row, 13)
'                 .TextMatrix(.Row, .Col) = strChecked
'                ' mAmount = Val(.TextMatrix(.Row, 4))
'                 mAmount = 0
'                 For x = 2 To MSHFlexGrid1.Rows - 2
'                     If mPINo = .TextMatrix(x, 13) Then
'                        .TextMatrix(x, .Col) = strChecked
'                        mAmount = Val(mAmount) + Val(.TextMatrix(x, 4))
'                     End If
'                 Next
'                 .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(mAmount)
'                 mAmount = 0
'              End If
'           End If
'        Else
'           mAmount = 0
'           .TextMatrix(.Row, .Col) = strUnChecked
'       '    mAmount = Val(.TextMatrix(.Row, 4))
'           mAmount = 0
'           mPINo = .TextMatrix(.Row, 13)
'           For x = 2 To MSHFlexGrid1.Rows - 1
'               If mPINo = .TextMatrix(x, 13) Then
'                  .TextMatrix(x, .Col) = strUnChecked
'                   mAmount = Val(mAmount) + Val(.TextMatrix(x, 4))
'               End If
'           Next
'           .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) - Val(mAmount)
'        End If
'     End If
'End With

End Sub

Private Sub OptGodownRent_Click()
Call Grid_Head
End Sub

Private Sub Option2_Click()
CmdExport.Enabled = False
Call Grid_Head
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbFAGBank.Text = ""
   CmbFAGBank.Enabled = False
   CmbModeOfPay.Enabled = False
   CmbModeOfPay.Text = ""
   CmdLetter.Enabled = True
   CmbApprovedBy.Text = ""
   CmbApprovedBy.Enabled = False
Else
   CmbBatch.Enabled = False
   CmbBatch.Text = ""
   CmbFAGBank.Text = ""
   CmbFAGBank.Enabled = True
   CmbModeOfPay.Enabled = True
   CmbModeOfPay.Text = ""
   CmdLetter.Enabled = False
   CmbApprovedBy.Text = ""
   CmbApprovedBy.Enabled = True
End If
End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 17
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(1, 0) = "Posting Date":         .ColWidth(0) = 1200
    .TextMatrix(1, 1) = "Document Type":        .ColWidth(1) = 1200
    .TextMatrix(1, 2) = "Account Type":         .ColWidth(2) = 1400
    .TextMatrix(1, 3) = "Account No.":          .ColWidth(3) = 1400
    .TextMatrix(1, 4) = "Amount":               .ColWidth(4) = 1600
    .TextMatrix(1, 5) = "Invoice No":           .ColWidth(5) = 1600
    .TextMatrix(1, 6) = "Narration":            .ColWidth(6) = 1600
    .TextMatrix(1, 7) = "RegionDim":            .ColWidth(7) = 2000
    .TextMatrix(1, 8) = "DeptDim":              .ColWidth(8) = 1600
    .TextMatrix(1, 9) = "EmpDim":               .ColWidth(9) = 1400
    .TextMatrix(1, 10) = "Cheque No.":          .ColWidth(10) = 1400
    .TextMatrix(1, 11) = "Cheque Date":         .ColWidth(11) = 1000
    .TextMatrix(1, 12) = "Deposit Date":        .ColWidth(12) = 1000
    .TextMatrix(1, 13) = "PINo":                .ColWidth(13) = 1200
    .TextMatrix(1, 14) = "PostInvoiceNo":       .ColWidth(14) = 1200
    .TextMatrix(1, 15) = "EntryNo":             .ColWidth(15) = 1200
    .TextMatrix(1, 16) = "Document No.":        .ColWidth(16) = 1200

End With
End Sub

Public Sub Grid_HeadVendor()
With MSHFlexGrid2
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 11
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(0, 0) = "Vendor Name"
    .TextMatrix(0, 1) = "Disbusment Amount"
    .TextMatrix(0, 2) = "FAG Bank Name"
    .TextMatrix(0, 3) = "FAG Account No"
    .TextMatrix(0, 4) = "Contact No"
    .TextMatrix(0, 5) = "Address"
    .TextMatrix(0, 6) = "Vendor Bank Acc No"
    .TextMatrix(0, 7) = "Vendor Bank Name"
    .TextMatrix(0, 8) = "IFSC Code"
    .TextMatrix(0, 9) = "Vendor Branch"
    .TextMatrix(0, 10) = "Vendor ID"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1600
    .ColWidth(5) = 1600
    .ColWidth(6) = 1600
    .ColWidth(7) = 2000
    .ColWidth(8) = 1600
    .ColWidth(9) = 1400
    
End With

End Sub

Private Sub CmbApprovedBy_GotFocus()

tmp1 = CmbApprovedBy.Text
'Call TableMst_Employee("Where Flag = 'Y' And DesignationID=")
tmp = "Select ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName From Mst_Employee ME"
tmp = tmp & " Inner join Mst_Users MU On MU.EmployeeNo=ME.EmployeeNo Where MU.UserRoleID in(7,14) And ME.Flag = 'Y'"

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
mUpdatedBy = TmpRs!EmployeeNo

End Sub
Private Sub CmbFAGBank_GotFocus()
tmp = CmbFAGBank.Text
Call TableMst_FAGBank("")
CmbFAGBank.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
    CmbFAGBank.AddItem RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
    RsMst_FAGBank.MoveNext
Next
CmbFAGBank.Text = tmp
End Sub

Private Sub CmbFAGBank_LostFocus()
Dim strSplitString() As String
Dim AcNo_, BankName_ As String

If CmbFAGBank.Text = "" Then
   Exit Sub
End If
strSplitString = Split(CmbFAGBank, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   CmbFAGBank.SetFocus
   Exit Sub
End If
BankName_ = Trim(strSplitString(0))
AcNo_ = Trim(strSplitString(1))
RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo='" & AcNo_ & "'"
If RsMst_FAGBank.EOF Then
   MsgBox "Invalid selection "
   CmbFAGBank.SetFocus
   Exit Sub
End If
If tmp <> CmbFAGBank.Text Then
   Call Grid_Head
End If
mFAGBankID = RsMst_FAGBank!FAGBankID
End Sub
Private Function GetPaymentType(Flg_ As Variant, tablenm As Variant) As Variant
Dim Retval As Variant
Retval = ""
If LCase(Flg_) = "e" Then
   tmp = "Select Top(1) ModeOfPay from " & tablenm & " Where PBatchID=" & Val(CmbBatch.Text)
ElseIf LCase(Flg_) = "v" Then
   tmp = "Select Top(1) ModeOfPay from " & tablenm & " Where PBatchID=" & Val(CmbBatch.Text) & "  And ModeofPay is not null "
End If
Call Tmp3Table
If TmpRs3.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs3!ModeOfPay), "", TmpRs3!ModeOfPay)
End If
GetPaymentType = Retval
End Function

Private Sub CmdLetter_Click()
Dim mSplit As Variant
Dim mR As Double
If CmbBatch.Text = "" Then
   MsgBox " Select Batch ID"
   Exit Sub
End If

If mNewLetterFormatFAGBankID <> "" Then
    mSplit = Split(mNewLetterFormatFAGBankID, ",")
    If UBound(mSplit) >= 0 Then
       For mR = 0 To UBound(mSplit)
           If Val(mSplit(mR)) = Val(mpBankID) Then
                mAllowNewLetterFormatTobePrint = True
           End If
       Next mR
    End If
Else
    mAllowNewLetterFormatTobePrint = False
End If

Gen_mTimeStamp = GetTimeStamp
If OptGodownRent.Value = True Then

'   Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_BankLetterGodownRent.xls")
'   Call Xls_Detail_RptVendor
   letternm = "_BankLetterVendor_"
   tmp = "Select MV.LessorName ,Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo 'FAGAccount',MF.ContactPerson,MF.Address,MV.AccountNo,MB.BankName 'VendorBankName',MV.RTGSNo,MBr.BranchName,MV.LessorID,isnull(MF.PaymentLimitation,1) 'PaymentLimitation' "
   tmp = tmp & " From Txn_GodownRent TA"
   tmp = tmp & " Inner Join Mst_Lessor MV on TA.LessorID=MV.LessorID"
   tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
   tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
   tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
   tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text)
   tmp = tmp & " Group By MV.LessorName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.AccountNo,MB.BankName,MV.RTGSNo,MBr.BranchName,MV.LessorID,MF.PaymentLimitation"
   Call Tmp4Table
    
   If TmpRs4.RecordCount = 0 Then Exit Sub
   
   tmp = "Select * from Txn_GodownRent Where LetterPrintFlag = 'Y' And PBatchID = " & Val(CmbBatch.Text) & " "
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      ans = MsgBox("Bank Letter already generated for this batch. Still you want to generate letter?", vbYesNo)
      If ans = vbNo Then
         Exit Sub
      End If
   End If
   
   MsgBox ("Wait till Next Message")
   
   Call Grid_HeadVendor
   mFileNo = 1
   MSHFlexGrid2.Rows = TmpRs4.RecordCount + 1
   
   For I = 1 To TmpRs4.RecordCount
       mPaymentLimitation = TmpRs4!PaymentLimitation
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs4.Fields(C)), "", TmpRs4.Fields(C))
       Next
       TmpRs4.MoveNext
   Next
   
   a = Round(Val(TmpRs4.RecordCount) / Val(mPaymentLimitation))
   y = Val(TmpRs4.RecordCount) / Val(mPaymentLimitation)
   If a < y Then
      a = a + 1
   End If
   mtot = Val(TmpRs4.RecordCount)
   
   I = 0
   y = 1
   mVendorID = ""
   For x = 0 To MSHFlexGrid2.Rows - 1
        If y < a Then
           If I = mPaymentLimitation Then
              If mVendorID = "" Then
                 mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10) & ")"
              Else
                 mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10) & " )"
              End If
              Call Create_Xls(Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
              
              ''-- Ullhas Patil on 17 Feb 2017 For Yes Bank Format
              If mAllowNewLetterFormatTobePrint = True Then   ''-- If Yes Bank
                Call Xls_Detail_RptVendorYesBank
              Else  ''-- If Other Banks
                Call Xls_Detail_RptVendor
              End If
              ''----------
              
              mFileNo = mFileNo + 1
              mVendorID = ""
              y = y + 1
              I = 0
              mtot = mtot - mPaymentLimitation
           Else
              If x > 0 Then
                 If mVendorID = "" Then
                    mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10)
                 Else
                    mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10)
                 End If
              End If
           End If
           
        ElseIf a = y Then
           If I = mtot Then
              If mVendorID = "" Then
                 mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10) & ")"
              Else
                 mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10) & " )"
              End If
              Call Create_Xls(Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
              ''Call Xls_Detail_RptVendor
              ''-- Ullhas Patil on 17 Feb 2017 For Yes Bank Format
              If mAllowNewLetterFormatTobePrint = True Then   ''-- If Yes Bank
                Call Xls_Detail_RptVendorYesBank
              Else  ''-- If Other Banks
                Call Xls_Detail_RptVendor
              End If
              ''----------
              mFileNo = mFileNo + 1
              mVendorID = ""
              y = y + 1
              I = 0
           Else
              If x > 0 Then
                 If mVendorID = "" Then
                    mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10)
                 Else
                    mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10)
                 End If
              
              End If
           End If
        End If
        I = I + 1
    Next
    MsgBox ("Excel File Created !!! ")

'   Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_BankLetterVendor.xls")
'   Call Xls_Detail_RptVendor

   tmp = "Update Txn_GodownRent Set LetterPrintFlag = 'Y' Where PBatchID = " & Val(CmbBatch.Text) & " "
   Call TmpTable
   
ElseIf OptSecurity.Value = True Then
   
   letternm = "_BankLetterSecurity_"
   
   tmp = "Select MV.AgencyName ,Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo 'FAGAccount',MF.ContactPerson,MF.Address,MV.BankAcno,MB.BankName 'VendorBankName',MV.IFSCCODE,MBr.BranchName,MV.SecurityAgencyID,isnull(MF.PaymentLimitation,1) 'PaymentLimitation' "
   tmp = tmp & " From Txn_SecurityBilling TA"
   tmp = tmp & " Inner Join Mst_SecurityAgency MV on TA.SecurityAgencyID=MV.SecurityAgencyID"
   tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
   tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
   tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
   tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text)
   tmp = tmp & " Group By MV.AgencyName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcno,MB.BankName,MV.IFSCCODE,MBr.BranchName,MV.SecurityAgencyID,MF.PaymentLimitation"
   
   Call Tmp4Table
    
   If TmpRs4.RecordCount = 0 Then Exit Sub
   
   
   tmp = "Select * from Txn_SecurityBilling Where LetterPrintFlag = 'Y' And PBatchID = " & Val(CmbBatch.Text) & " "
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      
      ans = MsgBox("Bank Letter already generated for this batch. Still you want to generate letter?", vbYesNo)
      If ans = vbNo Then
         Exit Sub
      End If
   End If
   
   MsgBox ("Wait till Next Message")
   
   Call Grid_HeadVendor
   mFileNo = 1
   MSHFlexGrid2.Rows = TmpRs4.RecordCount + 1
   
   For I = 1 To TmpRs4.RecordCount
       mPaymentLimitation = TmpRs4!PaymentLimitation
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs4.Fields(C)), "", TmpRs4.Fields(C))
       Next
       TmpRs4.MoveNext
   Next
   
   a = Round(Val(TmpRs4.RecordCount) / Val(mPaymentLimitation))
   y = Val(TmpRs4.RecordCount) / Val(mPaymentLimitation)
   If a < y Then
      a = a + 1
   End If
   mtot = Val(TmpRs4.RecordCount)
   
   I = 0
   y = 1
   mVendorID = ""
   For x = 0 To MSHFlexGrid2.Rows - 1
        If y < a Then
           If I = mPaymentLimitation Then
              If mVendorID = "" Then
                 mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10) & ")"
              Else
                 mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10) & " )"
              End If
              Call Create_Xls(Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
              ''Call Xls_Detail_RptVendor
              ''-- Ullhas Patil on 17 Feb 2017 For Yes Bank Format
              If mAllowNewLetterFormatTobePrint = True Then   ''-- If Yes Bank
                Call Xls_Detail_RptVendorYesBank
              Else  ''-- If Other Banks
                Call Xls_Detail_RptVendor
              End If
              ''----------
              mFileNo = mFileNo + 1
              mVendorID = ""
              y = y + 1
              I = 0
              mtot = mtot - mPaymentLimitation
           Else
              If x > 0 Then
                 If mVendorID = "" Then
                    mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10)
                 Else
                    mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10)
                 End If
              End If
           End If
           
        ElseIf a = y Then
           If I = mtot Then
              If mVendorID = "" Then
                 mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10) & ")"
              Else
                 mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10) & " )"
              End If
              Call Create_Xls(Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
              'Call Xls_Detail_RptVendor
              ''-- Ullhas Patil on 17 Feb 2017 For Yes Bank Format
              If mAllowNewLetterFormatTobePrint = True Then   ''-- If Yes Bank
                Call Xls_Detail_RptVendorYesBank
              Else  ''-- If Other Banks
                Call Xls_Detail_RptVendor
              End If
              ''----------
              
              mFileNo = mFileNo + 1
              mVendorID = ""
              y = y + 1
              I = 0
           Else
              If x > 0 Then
                 If mVendorID = "" Then
                    mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10)
                 Else
                    mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10)
                 End If
              
              End If
           End If
        End If
        I = I + 1
    Next
    MsgBox ("Excel File Created !!! ")

'   Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_BankLetterVendor.xls")
'   Call Xls_Detail_RptVendor

   tmp = "Update Txn_SecurityBilling Set LetterPrintFlag = 'Y' Where PBatchID = " & Val(CmbBatch.Text) & " "
   Call TmpTable
ElseIf optFumigation.Value = True Then
   
   letternm = "_BankLetterFumigation_"
   tmp = "Select MV.VendorName,Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo 'FAGAccount',MF.ContactPerson,MF.Address,MV.BankAcno,MB.BankName 'VendorBankName',MV.IFSCCODE,MBr.BranchName,MV.VendorID,isnull(MF.PaymentLimitation,1) 'PaymentLimitation' "
   tmp = tmp & " From Txn_FumigationApproval TA"
   tmp = tmp & " Inner Join Mst_Vendor MV on TA.VendorID=MV.VendorID"
   tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
   tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
   tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
   tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text)
   tmp = tmp & " Group By MV.VendorName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName,MV.IFSCCODE,MBr.BranchName,MV.VendorID,MF.PaymentLimitation"
   
   Call Tmp4Table
    
   If TmpRs4.RecordCount = 0 Then Exit Sub
   
   
   tmp = "Select * from Txn_FumigationApproval Where LetterPrintFlag = 'Y' And PBatchID = " & Val(CmbBatch.Text) & " "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      ans = MsgBox("Bank Letter already generated for this batch. Still you want to generate letter?", vbYesNo)
      If ans = vbNo Then
         Exit Sub
      End If
   End If
   
   MsgBox ("Wait till Next Message")
   
   Call Grid_HeadVendor
   mFileNo = 1
   MSHFlexGrid2.Rows = TmpRs4.RecordCount + 1
   
   For I = 1 To TmpRs4.RecordCount
       mPaymentLimitation = TmpRs4!PaymentLimitation
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs4.Fields(C)), "", TmpRs4.Fields(C))
       Next
       TmpRs4.MoveNext
   Next
   
   a = Round(Val(TmpRs4.RecordCount) / Val(mPaymentLimitation))
   y = Val(TmpRs4.RecordCount) / Val(mPaymentLimitation)
   If a < y Then
      a = a + 1
   End If
   mtot = Val(TmpRs4.RecordCount)
   
   I = 0
   y = 1
   mVendorID = ""
   For x = 0 To MSHFlexGrid2.Rows - 1
        If y < a Then
           If I = mPaymentLimitation Then
              If mVendorID = "" Then
                 mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10) & ")"
              Else
                 mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10) & " )"
              End If
              Call Create_Xls(Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
              'Call Xls_Detail_RptVendor
              ''-- Ullhas Patil on 17 Feb 2017 For Yes Bank Format
              If mAllowNewLetterFormatTobePrint = True Then   ''-- If Yes Bank
                Call Xls_Detail_RptVendorYesBank
              Else  ''-- If Other Banks
                Call Xls_Detail_RptVendor
              End If
              ''----------
              mFileNo = mFileNo + 1
              mVendorID = ""
              y = y + 1
              I = 0
              mtot = mtot - mPaymentLimitation
           Else
              If x > 0 Then
                 If mVendorID = "" Then
                    mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10)
                 Else
                    mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10)
                 End If
              End If
           End If
           
        ElseIf a = y Then
           If I = mtot Then
              If mVendorID = "" Then
                 mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10) & ")"
              Else
                 mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10) & " )"
              End If
              Call Create_Xls(Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
              'Call Xls_Detail_RptVendor
              ''-- Ullhas Patil on 17 Feb 2017 For Yes Bank Format
              If mAllowNewLetterFormatTobePrint = True Then   ''-- If Yes Bank
                Call Xls_Detail_RptVendorYesBank
              Else  ''-- If Other Banks
                Call Xls_Detail_RptVendor
              End If
              ''----------
              
              mFileNo = mFileNo + 1
              mVendorID = ""
              y = y + 1
              I = 0
           Else
              If x > 0 Then
                 If mVendorID = "" Then
                    mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10)
                 Else
                    mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10)
                 End If
              
              End If
           End If
        End If
        I = I + 1
    Next
    MsgBox ("Excel File Created !!! ")

'   Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_BankLetterVendor.xls")
'   Call Xls_Detail_RptVendor

   tmp = "Update Txn_FumigationApproval Set LetterPrintFlag = 'Y' Where PBatchID = " & Val(CmbBatch.Text) & " "
   Call TmpTable
ElseIf OptFranchiseeRent.Value = True Then


   letternm = "_BankLetterFranchiseeRent_"
   tmp = "Select MV.VendorName,Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo 'FAGAccount',MF.ContactPerson,MF.Address,MV.BankAcno,MB.BankName 'VendorBankName',MV.IFSCCODE,MBr.BranchName,MV.VendorID,isnull(MF.PaymentLimitation,1) 'PaymentLimitation' "
   tmp = tmp & " From Txn_FranchiseeRent TA"
   tmp = tmp & " Inner Join Mst_Vendor MV on TA.VendorID=MV.VendorID"
   tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
   tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
   tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
   tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text)
   tmp = tmp & " Group By MV.VendorName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName,MV.IFSCCODE,MBr.BranchName,MV.VendorID,MF.PaymentLimitation"
   
   Call Tmp4Table
    
   If TmpRs4.RecordCount = 0 Then Exit Sub
   
   
   tmp = "Select * from Txn_FranchiseeRent Where LetterPrintFlag = 'Y' And PBatchID = " & Val(CmbBatch.Text) & " "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      ans = MsgBox("Bank Letter already generated for this batch. Still you want to generate letter?", vbYesNo)
      If ans = vbNo Then
         Exit Sub
      End If
   End If
   
   MsgBox ("Wait till Next Message")
   
   Call Grid_HeadVendor
   mFileNo = 1
   MSHFlexGrid2.Rows = TmpRs4.RecordCount + 1
   
   For I = 1 To TmpRs4.RecordCount
       mPaymentLimitation = TmpRs4!PaymentLimitation
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs4.Fields(C)), "", TmpRs4.Fields(C))
       Next
       TmpRs4.MoveNext
   Next
   
   a = Round(Val(TmpRs4.RecordCount) / Val(mPaymentLimitation))
   y = Val(TmpRs4.RecordCount) / Val(mPaymentLimitation)
   If a < y Then
      a = a + 1
   End If
   mtot = Val(TmpRs4.RecordCount)
   
   I = 0
   y = 1
   mVendorID = ""
   For x = 0 To MSHFlexGrid2.Rows - 1
        If y < a Then
           If I = mPaymentLimitation Then
              If mVendorID = "" Then
                 mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10) & ")"
              Else
                 mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10) & " )"
              End If
              Call Create_Xls(Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
              'Call Xls_Detail_RptVendor
              ''-- Ullhas Patil on 17 Feb 2017 For Yes Bank Format
              If mAllowNewLetterFormatTobePrint = True Then   ''-- If Yes Bank
                Call Xls_Detail_RptVendorYesBank
              Else  ''-- If Other Banks
                Call Xls_Detail_RptVendor
              End If
              ''----------
              mFileNo = mFileNo + 1
              mVendorID = ""
              y = y + 1
              I = 0
              mtot = mtot - mPaymentLimitation
           Else
              If x > 0 Then
                 If mVendorID = "" Then
                    mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10)
                 Else
                    mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10)
                 End If
              End If
           End If
           
        ElseIf a = y Then
           If I = mtot Then
              If mVendorID = "" Then
                 mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10) & ")"
              Else
                 mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10) & " )"
              End If
              Call Create_Xls(Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
              'Call Xls_Detail_RptVendor
              ''-- Ullhas Patil on 17 Feb 2017 For Yes Bank Format
              If mAllowNewLetterFormatTobePrint = True Then   ''-- If Yes Bank
                Call Xls_Detail_RptVendorYesBank
              Else  ''-- If Other Banks
                Call Xls_Detail_RptVendor
              End If
              ''----------
              mFileNo = mFileNo + 1
              mVendorID = ""
              y = y + 1
              I = 0
           Else
              If x > 0 Then
                 If mVendorID = "" Then
                    mVendorID = "(" & MSHFlexGrid2.TextMatrix(x, 10)
                 Else
                    mVendorID = mVendorID & "," & MSHFlexGrid2.TextMatrix(x, 10)
                 End If
              
              End If
           End If
        End If
        I = I + 1
    Next
    MsgBox ("Excel File Created !!! ")

   tmp = "Update Txn_FranchiseeRent Set LetterPrintFlag = 'Y' Where PBatchID = " & Val(CmbBatch.Text) & " "
   Call TmpTable

End If
End Sub
Public Sub Create_Xls(FileName As String)
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Gen_FileName = FileName
oBook.SaveAs (Pat & "\Excel\" & FileName)

Set oSheet = oBook.Worksheets(1)
With oExcel
'.Application.Sheets.Delete
.SheetsInNewWorkbook = 15 'MSHFlexGrid2.Rows - 2
.Workbooks.Add
End With
'oBook.SaveAs (Pat & "\" & FileName)
'oExcel.Show
oExcel.Quit
'oExcel.Close
End Sub
Public Sub Xls_Detail_Rpt()
Dim ans As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path

tmp = "Select TA.ClaimID,TA.DisbuDate,ME.FullName ,TA.DisbuAmount,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,ME.AdvACNo"
tmp = tmp & " From Txn_ClaimSubmission TA"
tmp = tmp & " Inner Join Mst_Employee ME on TA.EmployeeID=ME.EmployeeID"
tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text)

Call Tmp4Table

If TmpRs4.RecordCount = 0 Then Exit Sub

Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_BankLetter.xls")
Set oWS = oWB.Worksheets("Sheet1")
    oWS.ClearArrows
    mRow = 3
    
    oWS.Range("a3:d3").Merge
    
    oWS.Range("a:a").ColumnWidth = 9.157 '10.86
    oWS.Range("b:b").ColumnWidth = 40.71 '17.71
    oWS.Range("c:c").ColumnWidth = 16.71 '17.14
    oWS.Range("d:d").ColumnWidth = 12 '12.86
    
    'oWS.Range("A1:E1").Borders.Color = RGB(0, 0, 0)
'    oWS.Cells(mRow, c + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
 '   oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
 '   oWS.Cells(mRow, c + 1).Font.Bold = True
 '   oWS.Cells(mRow, c + 1).Font.Size = 14
    
    'oWS.Cells(1, 1).Font.ColorIndex = 4
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1).RowHeight = 8
    
    mRow = mRow + 1
    'oWS.Range("a3:e3").Merge
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    tmp = "c" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    tmp = Format(Date, "dd-mmm-yyyy")
    oWS.Cells(mRow, C + 3) = "Date : " & tmp
    oWS.Cells(mRow, C + 3).HorizontalAlignment = -4108
    'oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, C + 3).Font.Size = 9
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1).RowHeight = 8
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "To"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "The Branch Manager,"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Bank Name : " & TmpRs4!BankName
    oWS.Cells(mRow, C + 1).Font.Size = 10
       
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = TmpRs4!Address
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    
    mRow = mRow + 2
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Kind Attn: " & TmpRs4!ContactPerson
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    
''    mRow = mRow + 2
''    tmp = "a" & mRow & ":" & "j" & mRow
''    oWS.Range(tmp).Merge
''    oWS.Cells(mRow, c + 1) = "Dear Sir,"
''    oWS.Cells(mRow, c + 1).Font.Size = 10
''
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "SUB : Transfer of Funds"
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Underline = True
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Dear Sir,"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 2
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Please debit our current account no. viz. " & TmpRs4!AccountNo & " with you and credit following accounts"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    mRow1 = mRow
    'tmp = "a" & mRow & ":" & "d" & mRow
    'oWS.Range(tmp).Merge
    'oWS.Cells(mRow, c + 1) = "with amount indicated; totaling to " & NumStrConv(
    'oWS.Cells(mRow, c + 1).Font.Size = 10
    
    mRow = mRow + 2
    oWS.Cells(mRow, C + 1).Font.Bold = True
    'oWS.Cells(mRow, c + 1).RowHeight = 40
    oWS.Cells(mRow, C + 1).WrapText = True
    oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 1) = "Sr. No."
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    oWS.Cells(mRow, C + 2).Font.Bold = True
    'oWS.Cells(mRow, c + 2).RowHeight = 40
    oWS.Cells(mRow, C + 2).WrapText = True
    oWS.Cells(mRow, C + 2).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 2) = "Name of Account Holder"
    oWS.Cells(mRow, C + 2).Font.Size = 10
    
    oWS.Cells(mRow, C + 3).Font.Bold = True
    'oWS.Cells(mRow, c + 3).RowHeight = 40
    oWS.Cells(mRow, C + 3).WrapText = True
    oWS.Cells(mRow, C + 3).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 3) = "Account No."
    oWS.Cells(mRow, C + 3).Font.Size = 10
    
    oWS.Cells(mRow, C + 4).Font.Bold = True
    'oWS.Cells(mRow, c + 4).RowHeight = 40
    oWS.Cells(mRow, C + 4).WrapText = True
    oWS.Cells(mRow, C + 4).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 4) = "Amount (Rs.)"
    oWS.Cells(mRow, C + 4).Font.Size = 10
    
    mTotAmt = 0
    For I = 1 To TmpRs4.RecordCount
        mRow = mRow + 1 ' + 6
        
        For C = 0 To 3
            oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
        Next
        C = 0
        oWS.Cells(mRow, C + 1) = I
        oWS.Cells(mRow, C + 1).Font.Size = 10
        oWS.Cells(mRow, C + 2) = TmpRs4!FullName
        oWS.Cells(mRow, C + 2).Font.Size = 10
        oWS.Cells(mRow, C + 3) = IIf(IsNull(TmpRs4!AdvACNo), "", TmpRs4!AdvACNo)
        oWS.Cells(mRow, C + 3).Font.Size = 10
        oWS.Cells(mRow, C + 4) = IIf(IsNull(TmpRs4!DisbuAmount), "", TmpRs4!DisbuAmount)
        oWS.Cells(mRow, C + 4).Font.Size = 10
        mTotAmt = mTotAmt + IIf(IsNull(TmpRs4!DisbuAmount), 0, TmpRs4!DisbuAmount)
        TmpRs4.MoveNext
    Next
    
    mRow = mRow + 1
    oWS.Cells(mRow, 3) = "Total"
    oWS.Cells(mRow, 3).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 3).Font.Size = 10
    oWS.Cells(mRow, 4) = mTotAmt
    oWS.Cells(mRow, 4).Font.Size = 10
    oWS.Cells(mRow, 4).Borders.Color = RGB(0, 0, 0)
    
    tmp = "a" & mRow1 & ":" & "d" & mRow1
    oWS.Range(tmp).Merge
    oWS.Cells(mRow1, C + 1) = "with amount indicated; totaling to Rs." & mTotAmt & "/- (" & NumStrConv(Val(mTotAmt)) & ")"
    oWS.Cells(mRow1, C + 1).Font.Size = 10
    
    
    mRow = mRow + 3
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, 1) = "Thanking You,"
    oWS.Cells(mRow, 1).Font.Size = 10
    
    mRow = mRow + 3
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, 1) = "For National Collateral Management Services Ltd."
    oWS.Cells(mRow, 1).Font.Size = 8
    oWS.Range(tmp).Font.Name = "Zurich BT"
    
    mRow = mRow + 3
    
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    
    oWS.Cells(mRow, 1) = "Authorised Signatories"
    oWS.Cells(mRow, 1).Font.Size = 8
    oWS.Range(tmp).Font.Name = "Zurich BT"

oWB.Save
'oWB.Close
'oXL.Quit
MsgBox ("Excel File Created !!! ")
oXL.Visible = True


End Sub
Private Sub OptVendorClaim_Click()
Call Grid_Head
End Sub




Public Sub Xls_Detail_RptVendor()
Dim ans As Variant
'MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path

If OptGodownRent.Value = True Then
    letternm = "_BankLetterVendor_"
    tmp = "Select MV.LessorName ,Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo 'FAGAccount',MF.ContactPerson,MF.Address,MV.AccountNo,MB.BankName 'VendorBankName',MV.RTGSNo,MBr.BranchName "
    tmp = tmp & " From Txn_GodownRent TA"
    tmp = tmp & " Inner Join Mst_Lessor MV on TA.LessorID=MV.LessorID"
    tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
    tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
    tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
    tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text) & " And MV.LessorID in " & mVendorID & ""
    tmp = tmp & " Group By MV.LessorName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.AccountNo,MB.BankName,MV.RTGSNo,MBr.BranchName"
ElseIf OptSecurity.Value = True Then
    letternm = "_BankLetterSecurity_"
    tmp = "Select MV.AgencyName ,Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo 'FAGAccount',MF.ContactPerson,MF.Address,MV.BankAcno,MB.BankName 'VendorBankName',MV.IFSCCODE,MBr.BranchName "
    tmp = tmp & " From Txn_SecurityBilling TA"
    tmp = tmp & " Inner Join Mst_SecurityAgency MV on TA.SecurityAgencyID=MV.SecurityAgencyID"
    tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
    tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
    tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
    tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text) & " And MV.SecurityAgencyID in " & mVendorID & ""
    tmp = tmp & " Group By MV.AgencyName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcno,MB.BankName,MV.IFSCCODE,MBr.BranchName"
ElseIf optFumigation.Value = True Then
    letternm = "_BankLetterFumigation_"
    tmp = "Select MV.VendorName,Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo 'FAGAccount',MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName 'VendorBankName',MV.IFSCCODE,MBr.BranchName "
    tmp = tmp & " From Txn_FumigationApproval TA"
    tmp = tmp & " Inner Join Mst_Vendor MV on TA.VendorID=MV.VendorID"
    tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
    tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
    tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
    tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text) & " And MV.VendorID in " & mVendorID & ""
    tmp = tmp & " Group By MV.VendorName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName,MV.IFSCCODE,MBr.BranchName"
ElseIf OptFranchiseeRent.Value = True Then
    letternm = "_BankLetterFranchiseeRent_"
    tmp = "Select MV.VendorName,Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo 'FAGAccount',MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName 'VendorBankName',MV.IFSCCODE,MBr.BranchName "
    tmp = tmp & " From Txn_FranchiseeRent TA"
    tmp = tmp & " Inner Join Mst_Vendor MV on TA.VendorID=MV.VendorID"
    tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
    tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
    tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
    tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text) & " And MV.VendorID in " & mVendorID & ""
    tmp = tmp & " Group By MV.VendorName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName,MV.IFSCCODE,MBr.BranchName"
End If

Call Tmp5Table

'If TmpRs5.RecordCount = 0 Then Exit Sub

'Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_BankLetterGodownRent.xls")
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
    oWS.ClearArrows
    mRow = 3
    
    oWS.Range("a3:g3").Merge
    
    oWS.Range("a:a").ColumnWidth = 9.157 '10.86
    oWS.Range("b:b").ColumnWidth = 21.71 '17.71
    oWS.Range("c:c").ColumnWidth = 13.71 '17.14
    oWS.Range("d:d").ColumnWidth = 10 '12.86
    oWS.Range("e:e").ColumnWidth = 10 '12.86
    oWS.Range("f:f").ColumnWidth = 10 '12.86
    oWS.Range("g:g").ColumnWidth = 10 '12.86
 
   'oWS.Range("A1:E1").Borders.Color = RGB(0, 0, 0)
'    oWS.Cells(mRow, c + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
 '   oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
 '   oWS.Cells(mRow, c + 1).Font.Bold = True
 '   oWS.Cells(mRow, c + 1).Font.Size = 14
    
    'oWS.Cells(1, 1).Font.ColorIndex = 4
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1).RowHeight = 8
    
    mRow = mRow + 1
    'oWS.Range("a3:e3").Merge
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    tmp = "c" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    tmp = Format(Date, "dd-mmm-yyyy")
    oWS.Cells(mRow, C + 3) = "Date : " & tmp
    oWS.Cells(mRow, C + 3).HorizontalAlignment = -4108
    'oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, C + 3).Font.Size = 9
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1).RowHeight = 8
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "To"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "The Branch Manager,"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Bank Name : " & TmpRs5!BankName
    oWS.Cells(mRow, C + 1).Font.Size = 10
       
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = TmpRs5!Address
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    
    mRow = mRow + 2
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Kind Attn: " & TmpRs5!ContactPerson
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    
''    mRow = mRow + 2
''    tmp = "a" & mRow & ":" & "j" & mRow
''    oWS.Range(tmp).Merge
''    oWS.Cells(mRow, c + 1) = "Dear Sir,"
''    oWS.Cells(mRow, c + 1).Font.Size = 10
''
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "SUB : Transfer of Funds"
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Underline = True
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Dear Sir,"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 2
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Please debit our current account no. viz. " & TmpRs5!FAGAccount & " with you and credit following accounts"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    mRow1 = mRow
    'tmp = "a" & mRow & ":" & "g" & mRow
    'oWS.Range(tmp).Merge
    'oWS.Cells(mRow, c + 1) = "with amount indicated; totaling to " & NumStrConv(
    'oWS.Cells(mRow, c + 1).Font.Size = 10
    
    mRow = mRow + 2
    oWS.Cells(mRow, C + 1).Font.Bold = True
    'oWS.Cells(mRow, c + 1).RowHeight = 40
    oWS.Cells(mRow, C + 1).WrapText = True
    oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 1) = "Sr. No."
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    oWS.Cells(mRow, C + 2).Font.Bold = True
    'oWS.Cells(mRow, c + 2).RowHeight = 40
    oWS.Cells(mRow, C + 2).WrapText = True
    oWS.Cells(mRow, C + 2).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 2) = "Name of Account Holder"
    oWS.Cells(mRow, C + 2).Font.Size = 10
    
    oWS.Cells(mRow, C + 3).Font.Bold = True
    'oWS.Cells(mRow, c + 3).RowHeight = 40
    oWS.Cells(mRow, C + 3).WrapText = True
    oWS.Cells(mRow, C + 3).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 3) = "Bank Name"
    oWS.Cells(mRow, C + 3).Font.Size = 10
    
    oWS.Cells(mRow, C + 4).Font.Bold = True
    'oWS.Cells(mRow, c + 4).RowHeight = 40
    oWS.Cells(mRow, C + 4).WrapText = True
    oWS.Cells(mRow, C + 4).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 4) = "Branch Details"
    oWS.Cells(mRow, C + 4).Font.Size = 10
    
    
    oWS.Cells(mRow, C + 5).Font.Bold = True
    'oWS.Cells(mRow, c + 4).RowHeight = 40
    oWS.Cells(mRow, C + 5).WrapText = True
    oWS.Cells(mRow, C + 5).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 5) = "RTGS Code"
    oWS.Cells(mRow, C + 5).Font.Size = 10
    
    
    
    oWS.Cells(mRow, C + 6).Font.Bold = True
    'oWS.Cells(mRow, c + 4).RowHeight = 40
    oWS.Cells(mRow, C + 6).WrapText = True
    oWS.Cells(mRow, C + 6).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 6) = "Account Number"
    oWS.Cells(mRow, C + 6).Font.Size = 10
    
    
    oWS.Cells(mRow, C + 7).Font.Bold = True
    'oWS.Cells(mRow, c + 4).RowHeight = 40
    oWS.Cells(mRow, C + 7).WrapText = True
    oWS.Cells(mRow, C + 7).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 7) = "Amt (Rs.)"
    oWS.Cells(mRow, C + 7).Font.Size = 10
    
    
    mTotAmt = 0
    For I = 1 To TmpRs5.RecordCount
        mRow = mRow + 1 ' + 6
        
        For C = 0 To 6
            oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
        Next
        C = 0
        
        oWS.Cells(mRow, C + 1) = I
        oWS.Cells(mRow, C + 1).Font.Size = 10
        
        If OptGodownRent.Value = True Then
            oWS.Cells(mRow, C + 2) = TmpRs5!LessorName
        ElseIf OptSecurity.Value = True Then
            oWS.Cells(mRow, C + 2) = IIf(IsNull(TmpRs5!AgencyName), "", TmpRs5!AgencyName)
        Else
            oWS.Cells(mRow, C + 2) = IIf(IsNull(TmpRs5!VendorName), "", TmpRs5!VendorName)
        End If
        oWS.Cells(mRow, C + 2).Font.Size = 10
        oWS.Cells(mRow, C + 3) = IIf(IsNull(TmpRs5!VendorBankName), "", TmpRs5!VendorBankName)
        oWS.Cells(mRow, C + 3).Font.Size = 10
        oWS.Cells(mRow, C + 4) = IIf(IsNull(TmpRs5!BranchName), "", TmpRs5!BranchName)
        oWS.Cells(mRow, C + 4).Font.Size = 10
        If OptGodownRent.Value = True Then
            oWS.Cells(mRow, C + 5) = IIf(IsNull(TmpRs5!RTGSNo), "", TmpRs5!RTGSNo)
        ElseIf OptSecurity.Value = True Then
            oWS.Cells(mRow, C + 5) = IIf(IsNull(TmpRs5!IFSCCode), "", TmpRs5!IFSCCode)
        Else
            oWS.Cells(mRow, C + 5) = IIf(IsNull(TmpRs5!IFSCCode), "", TmpRs5!IFSCCode)
        End If
        oWS.Cells(mRow, C + 5).Font.Size = 10
        If OptGodownRent.Value = True Then
            oWS.Cells(mRow, C + 6) = IIf(IsNull(TmpRs5!AccountNo), "", TmpRs5!AccountNo)
        ElseIf OptSecurity.Value = True Then
            oWS.Cells(mRow, C + 6) = IIf(IsNull(TmpRs5!BankAcNo), "", TmpRs5!BankAcNo)
        Else
            oWS.Cells(mRow, C + 6) = IIf(IsNull(TmpRs5!BankAcNo), "", TmpRs5!BankAcNo)
        End If
        
        oWS.Cells(mRow, C + 6).Font.Size = 10
        oWS.Cells(mRow, C + 7) = IIf(IsNull(TmpRs5!DisbuAmount), "", TmpRs5!DisbuAmount)
        oWS.Cells(mRow, C + 7).Font.Size = 10
        mTotAmt = mTotAmt + IIf(IsNull(TmpRs5!DisbuAmount), 0, TmpRs5!DisbuAmount)
        
        oWS.Cells(mRow, C + 2).WrapText = True
        oWS.Cells(mRow, C + 3).WrapText = True
        oWS.Cells(mRow, C + 4).WrapText = True
        TmpRs5.MoveNext
    Next
    
    mRow = mRow + 1
    oWS.Cells(mRow, 6) = "Total"
    oWS.Cells(mRow, 6).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 6).Font.Size = 10
    oWS.Cells(mRow, 7) = mTotAmt
    oWS.Cells(mRow, 7).Font.Size = 10
    oWS.Cells(mRow, 7).Borders.Color = RGB(0, 0, 0)
    
    tmp = "a" & mRow1 & ":" & "g" & mRow1
    oWS.Range(tmp).Merge
    oWS.Cells(mRow1, C + 1) = "with amount indicated totaling to Rs." & mTotAmt & "/- (" & NumStrConv(Val(mTotAmt)) & ")"
    oWS.Cells(mRow1, C + 1).Font.Size = 10
    
    
    mRow = mRow + 3
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, 1) = "Thanking You,"
    oWS.Cells(mRow, 1).Font.Size = 10
    
    mRow = mRow + 3
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, 1) = "For National Collateral Management Services Ltd."
    oWS.Cells(mRow, 1).Font.Size = 8
    oWS.Range(tmp).Font.Name = "Zurich BT"
    
    mRow = mRow + 3
    
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    
    oWS.Cells(mRow, 1) = "Authorised Signatories"
    oWS.Cells(mRow, 1).Font.Size = 8
    oWS.Range(tmp).Font.Name = "Zurich BT"

oWB.Save
'oWB.Close
'oXL.Quit
'MsgBox ("Excel File Created !!! ")
oXL.Visible = True


End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExpensePaymentJVExcel.xls")
Call Xls_Detail_RptExcel
End Sub
Public Sub Xls_Detail_RptExcel()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExpensePaymentJVExcel.xls")
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

''-- Ullhas Patil On 17 Feb 2017
Private Function GetPaymentFAGBankID(Flg_ As Variant, tablenm As Variant) As Variant
Dim Retval As Variant
Retval = ""
If LCase(Flg_) = "e" Then
   tmp = "Select Top(1) FAGBankID from " & tablenm & " Where PBatchID=" & Val(CmbBatch.Text)
ElseIf LCase(Flg_) = "v" Then
   tmp = "Select Top(1) FAGBankID from " & tablenm & " Where PBatchID=" & Val(CmbBatch.Text) & "  And ModeofPay is not null "
End If
Call Tmp3Table
If TmpRs3.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs3!FAGBankID), "", TmpRs3!FAGBankID)
End If
GetPaymentFAGBankID = Retval
End Function

''-- Ullhas Patil On 17 Feb 2017
Public Sub Xls_Detail_RptVendorYesBank()
Dim ans As Variant
'MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path


''tmp = "Select MV.VendorName , Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName 'VendorBankName',MV.IFSCCode,MV.BankBranch "
''tmp = tmp & " From Txn_VendorClaim TA"
''tmp = tmp & " Inner Join Mst_Vendor MV on TA.VendorID=MV.VendorID"
''tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
''tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
''tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text) & " and TA.VendorID in " & mVendorID & ""
''tmp = tmp & " Group By MV.VendorName,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName,MV.IFSCCode,MV.BankBranch "
''Call Tmp5Table

If OptGodownRent.Value = True Then
    letternm = "_BankLetterVendor_"
    tmp = "Select MV.LessorName 'VendorName',Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.AccountNo 'BankAcNo',MB.BankName 'VendorBankName',MV.RTGSNo 'IFSCCode',MBr.BranchName 'BankBranch' "
    tmp = tmp & " From Txn_GodownRent TA"
    tmp = tmp & " Inner Join Mst_Lessor MV on TA.LessorID=MV.LessorID"
    tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
    tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
    tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
    tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text) & " And MV.LessorID in " & mVendorID & ""
    tmp = tmp & " Group By MV.LessorName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.AccountNo,MB.BankName,MV.RTGSNo,MBr.BranchName"
ElseIf OptSecurity.Value = True Then
    letternm = "_BankLetterSecurity_"
    tmp = "Select MV.AgencyName 'VendorName',Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcno 'BankAcNo',MB.BankName 'VendorBankName',MV.IFSCCODE 'IFSCCODE',MBr.BranchName 'BankBranch' "
    tmp = tmp & " From Txn_SecurityBilling TA"
    tmp = tmp & " Inner Join Mst_SecurityAgency MV on TA.SecurityAgencyID=MV.SecurityAgencyID"
    tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
    tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
    tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
    tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text) & " And MV.SecurityAgencyID in " & mVendorID & ""
    tmp = tmp & " Group By MV.AgencyName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcno,MB.BankName,MV.IFSCCODE,MBr.BranchName"
ElseIf optFumigation.Value = True Then
    letternm = "_BankLetterFumigation_"
    tmp = "Select MV.VendorName 'VendorName',Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo 'BankAcNo',MB.BankName 'VendorBankName',MV.IFSCCODE 'IFSCCODE',MBr.BranchName 'BankBranch'"
    tmp = tmp & " From Txn_FumigationApproval TA"
    tmp = tmp & " Inner Join Mst_Vendor MV on TA.VendorID=MV.VendorID"
    tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
    tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
    tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
    tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text) & " And MV.VendorID in " & mVendorID & ""
    tmp = tmp & " Group By MV.VendorName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName,MV.IFSCCODE,MBr.BranchName"
ElseIf OptFranchiseeRent.Value = True Then
    letternm = "_BankLetterFranchiseeRent_"
    tmp = "Select MV.VendorName 'VendorName',Sum(TA.DisbuAmount) 'DisbuAmount',MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo 'BankAcNo',MB.BankName 'VendorBankName',MV.IFSCCODE 'IFSCCODE',MBr.BranchName 'BankBranch'"
    tmp = tmp & " From Txn_FranchiseeRent TA"
    tmp = tmp & " Inner Join Mst_Vendor MV on TA.VendorID=MV.VendorID"
    tmp = tmp & " Inner Join Mst_FAGBank  MF on MF.FAGBankID=TA.FAGBankID"
    tmp = tmp & " Inner Join Mst_Bank  MB on MV.BankID=MB.BankID"
    tmp = tmp & " Inner Join Mst_Branch  MBr on MV.BranchID=MBr.BranchID"
    tmp = tmp & " Where TA.PBatchID=" & Val(CmbBatch.Text) & " And MV.VendorID in " & mVendorID & ""
    tmp = tmp & " Group By MV.VendorName ,MF.BankName,MF.AccountNo,MF.ContactPerson,MF.Address,MV.BankAcNo,MB.BankName,MV.IFSCCODE,MBr.BranchName"
End If

Call Tmp5Table
If TmpRs5.RecordCount = 0 Then
Exit Sub
End If
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & letternm & CmbBatch.Text & "_" & mFileNo & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
    oWS.ClearArrows
    mRow = 3
    
    oWS.Range("a3:g3").Merge
    
    oWS.Range("a:a").ColumnWidth = 9.157 '10.86
    oWS.Range("b:b").ColumnWidth = 21.71 '17.71
    oWS.Range("c:c").ColumnWidth = 13.71 '17.14
    oWS.Range("d:d").ColumnWidth = 10 '12.86
    oWS.Range("e:e").ColumnWidth = 10 '12.86
    oWS.Range("f:f").ColumnWidth = 10 '12.86
    oWS.Range("g:g").ColumnWidth = 10 '12.86
 
   'oWS.Range("A1:E1").Borders.Color = RGB(0, 0, 0)
'    oWS.Cells(mRow, c + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
 '   oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
 '   oWS.Cells(mRow, c + 1).Font.Bold = True
 '   oWS.Cells(mRow, c + 1).Font.Size = 14
    
    'oWS.Cells(1, 1).Font.ColorIndex = 4
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1).RowHeight = 8
    
    mRow = mRow + 1
    'oWS.Range("a3:e3").Merge
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    tmp = "c" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    'tmp = Format(Date, "dd-mmm-yyyy")
    'oWS.Cells(mRow, c + 3) = "Date : " & tmp
    'oWS.Cells(mRow, c + 3).HorizontalAlignment = -4108
    'oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, C + 3).Font.Size = 9
    'oWS.Cells(mRow, c + 3).HorizontalAlignment =
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1).RowHeight = 8
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "" ''-- Blank
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "" ''-- Blank
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    'oWS.Cells(mRow, c + 1) = "Bank Name : " & TmpRs5!BankName
    oWS.Cells(mRow, C + 1) = TmpRs5!BankName
    oWS.Cells(mRow, C + 1).Font.Size = 10
       
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = TmpRs5!Address
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    
    mRow = mRow + 2
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Kind Attn: " & TmpRs5!ContactPerson
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).Font.Bold = True
    
''    mRow = mRow + 2
''    tmp = "a" & mRow & ":" & "j" & mRow
''    oWS.Range(tmp).Merge
''    oWS.Cells(mRow, c + 1) = "Dear Sir,"
''    oWS.Cells(mRow, c + 1).Font.Size = 10
''
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    'oWS.Cells(mRow, c + 1) = "SUB : Transfer of Funds"
    'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, C + 1) = "Sub : Funds transfer through NEFT/RTGS -"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).Font.Bold = True
    'oWS.Cells(mRow, c + 1).Font.Underline = True
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    'oWS.Cells(mRow, c + 1) = "Dear Sir,"
    oWS.Cells(mRow, C + 1) = ""
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 2
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    'oWS.Cells(mRow, c + 1) = "Please debit our current account no. viz. " & TmpRs5!AccountNo & " with you and credit following accounts"
    oWS.Cells(mRow, C + 1) = "Please debit our current account no. " & TmpRs5!AccountNo & " with you and credit following accounts"
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    mRow1 = mRow
    'tmp = "a" & mRow & ":" & "g" & mRow
    'oWS.Range(tmp).Merge
    'oWS.Cells(mRow, c + 1) = "with amount indicated; totaling to " & NumStrConv(
    'oWS.Cells(mRow, c + 1).Font.Size = 10
    
    mRow = mRow + 2
    
    oWS.Cells(mRow, C + 1).Font.Bold = True
    'oWS.Cells(mRow, c + 1).RowHeight = 40
    oWS.Cells(mRow, C + 1).WrapText = True
    oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 1) = "Sr. No."
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    oWS.Cells(mRow, C + 2).Font.Bold = True
    oWS.Cells(mRow, C + 2).WrapText = True
    oWS.Cells(mRow, C + 2).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 2) = "Name of Account Holder"
    oWS.Cells(mRow, C + 2).Font.Size = 10
    
    oWS.Cells(mRow, C + 3).Font.Bold = True
    oWS.Cells(mRow, C + 3).WrapText = True
    oWS.Cells(mRow, C + 3).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 3) = "Account Number"
    oWS.Cells(mRow, C + 3).Font.Size = 10
    
    oWS.Cells(mRow, C + 4).Font.Bold = True
    oWS.Cells(mRow, C + 4).WrapText = True
    oWS.Cells(mRow, C + 4).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 4) = "RTGS Code"
    oWS.Cells(mRow, C + 4).Font.Size = 10
    
    oWS.Cells(mRow, C + 5).Font.Bold = True
    oWS.Cells(mRow, C + 5).WrapText = True
    oWS.Cells(mRow, C + 5).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, C + 5) = "Amt (Rs.)"
    oWS.Cells(mRow, C + 5).Font.Size = 10
    
    
''    oWS.Cells(mRow, c + 1).Font.Bold = True
''    'oWS.Cells(mRow, c + 1).RowHeight = 40
''    oWS.Cells(mRow, c + 1).WrapText = True
''    oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
''    oWS.Cells(mRow, c + 1) = "Sr. No."
''    oWS.Cells(mRow, c + 1).Font.Size = 10
''
''    oWS.Cells(mRow, c + 2).Font.Bold = True
''    'oWS.Cells(mRow, c + 2).RowHeight = 40
''    oWS.Cells(mRow, c + 2).WrapText = True
''    oWS.Cells(mRow, c + 2).Borders.Color = RGB(0, 0, 0)
''    oWS.Cells(mRow, c + 2) = "Name of Account Holder"
''    oWS.Cells(mRow, c + 2).Font.Size = 10
''
''    oWS.Cells(mRow, c + 3).Font.Bold = True
''    'oWS.Cells(mRow, c + 3).RowHeight = 40
''    oWS.Cells(mRow, c + 3).WrapText = True
''    oWS.Cells(mRow, c + 3).Borders.Color = RGB(0, 0, 0)
''    oWS.Cells(mRow, c + 3) = "Bank Name"
''    oWS.Cells(mRow, c + 3).Font.Size = 10
''
''    oWS.Cells(mRow, c + 4).Font.Bold = True
''    'oWS.Cells(mRow, c + 4).RowHeight = 40
''    oWS.Cells(mRow, c + 4).WrapText = True
''    oWS.Cells(mRow, c + 4).Borders.Color = RGB(0, 0, 0)
''    oWS.Cells(mRow, c + 4) = "Branch Details"
''    oWS.Cells(mRow, c + 4).Font.Size = 10
''
''    oWS.Cells(mRow, c + 5).Font.Bold = True
''    'oWS.Cells(mRow, c + 4).RowHeight = 40
''    oWS.Cells(mRow, c + 5).WrapText = True
''    oWS.Cells(mRow, c + 5).Borders.Color = RGB(0, 0, 0)
''    oWS.Cells(mRow, c + 5) = "RTGS Code"
''    oWS.Cells(mRow, c + 5).Font.Size = 10
''
''    oWS.Cells(mRow, c + 6).Font.Bold = True
''    'oWS.Cells(mRow, c + 4).RowHeight = 40
''    oWS.Cells(mRow, c + 6).WrapText = True
''    oWS.Cells(mRow, c + 6).Borders.Color = RGB(0, 0, 0)
''    oWS.Cells(mRow, c + 6) = "Account Number"
''    oWS.Cells(mRow, c + 6).Font.Size = 10
''
''    oWS.Cells(mRow, c + 7).Font.Bold = True
''    'oWS.Cells(mRow, c + 4).RowHeight = 40
''    oWS.Cells(mRow, c + 7).WrapText = True
''    oWS.Cells(mRow, c + 7).Borders.Color = RGB(0, 0, 0)
''    oWS.Cells(mRow, c + 7) = "Amt (Rs.)"
''    oWS.Cells(mRow, c + 7).Font.Size = 10
    
    mTotAmt = 0
    For iq = 1 To TmpRs5.RecordCount
        mRow = mRow + 1 ' + 6
        
        For C = 0 To 4
            oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
        Next
        C = 0
        
        oWS.Cells(mRow, C + 1) = iq
        oWS.Cells(mRow, C + 1).Font.Size = 10
        oWS.Cells(mRow, C + 2) = TmpRs5!VendorName
        oWS.Cells(mRow, C + 2).Font.Size = 10
        'oWS.Cells(mRow, c + 3) = IIf(IsNull(TmpRs5!VendorBankName), "", TmpRs5!VendorBankName)
        'oWS.Cells(mRow, c + 3).Font.Size = 10
        'oWS.Cells(mRow, c + 4) = IIf(IsNull(TmpRs5!BankBranch), "", TmpRs5!BankBranch)
        'oWS.Cells(mRow, c + 4).Font.Size = 10
        'oWS.Cells(mRow, c + 5) = IIf(IsNull(TmpRs5!IFSCCODE), "", TmpRs5!IFSCCODE)
        'oWS.Cells(mRow, c + 5).Font.Size = 10
        'oWS.Cells(mRow, c + 6) = IIf(IsNull(TmpRs5!BankAcNo), "", TmpRs5!BankAcNo)
        'oWS.Cells(mRow, c + 6).Font.Size = 10
        'oWS.Cells(mRow, c + 7) = IIf(IsNull(TmpRs5!DisbuAmount), "", TmpRs5!DisbuAmount)
        'oWS.Cells(mRow, c + 7).Font.Size = 10
        'mTotAmt = mTotAmt + IIf(IsNull(TmpRs5!DisbuAmount), 0, TmpRs5!DisbuAmount)
        
        oWS.Cells(mRow, C + 3) = IIf(IsNull(TmpRs5!BankAcNo), "", TmpRs5!BankAcNo)
        oWS.Cells(mRow, C + 3).Font.Size = 10
        oWS.Cells(mRow, C + 4) = IIf(IsNull(TmpRs5!IFSCCode), "", TmpRs5!IFSCCode)
        oWS.Cells(mRow, C + 4).Font.Size = 10
        oWS.Cells(mRow, C + 5) = IIf(IsNull(TmpRs5!DisbuAmount), "", TmpRs5!DisbuAmount)
        oWS.Cells(mRow, C + 5).Font.Size = 10
        
        mTotAmt = mTotAmt + IIf(IsNull(TmpRs5!DisbuAmount), 0, TmpRs5!DisbuAmount)
        
        oWS.Cells(mRow, C + 2).WrapText = True
        oWS.Cells(mRow, C + 3).WrapText = True
        oWS.Cells(mRow, C + 4).WrapText = True
        TmpRs5.MoveNext
    Next
    
    mRow = mRow + 1
'    oWS.Cells(mRow, 6) = "Total"
'    oWS.Cells(mRow, 6).Borders.Color = RGB(0, 0, 0)
'    oWS.Cells(mRow, 6).Font.Size = 10
'    oWS.Cells(mRow, 7) = mTotAmt
'    oWS.Cells(mRow, 7).Font.Size = 10
'    oWS.Cells(mRow, 7).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 1).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 2).Borders.Color = RGB(0, 0, 0)

    oWS.Cells(mRow, 3) = "Total"
    oWS.Cells(mRow, 3).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 3).Font.Size = 10
    oWS.Cells(mRow, 4).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 5) = mTotAmt
    oWS.Cells(mRow, 5).Font.Size = 10
    oWS.Cells(mRow, 5).Borders.Color = RGB(0, 0, 0)

    tmp = "a" & mRow1 & ":" & "g" & mRow1
    oWS.Range(tmp).Merge
    oWS.Cells(mRow1, C + 1) = "with amount indicated totaling to Rs." & mTotAmt & "/- (" & NumStrConv(Val(mTotAmt)) & ")"
    oWS.Cells(mRow1, C + 1).Font.Size = 10
    
    mRow = mRow + 3
    tmp = "a" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, 1) = "Thanking You,"
    oWS.Cells(mRow, 1).Font.Size = 10
    
    mRow = mRow + 3
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, 1) = "For National Collateral Management Services Ltd."
    oWS.Cells(mRow, 1).Font.Size = 8
    oWS.Range(tmp).Font.Name = "Zurich BT"
    
    mRow = mRow + 3
    
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    
    oWS.Cells(mRow, 1) = "Authorised Signatories"
    oWS.Cells(mRow, 1).Font.Size = 8
    oWS.Range(tmp).Font.Name = "Zurich BT"

oWB.Save
'oWB.Close
'oXL.Quit
'MsgBox ("Wait till Next Message")
'MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub



