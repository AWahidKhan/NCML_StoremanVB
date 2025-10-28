VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CashReceiptForNavision 
   Caption         =   "Cash Receipt For Navision Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7455
      Left            =   120
      TabIndex        =   30
      Top             =   2220
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   31
         Top             =   7245
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6975
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12303
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   15015
      Begin VB.CommandButton CmdRemoveErrorFlag 
         Caption         =   "Remove Error Flag"
         Height          =   375
         Left            =   12880
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1620
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CheckBox ChkSelectAll 
         Height          =   240
         Left            =   120
         TabIndex        =   32
         Top             =   1755
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.ComboBox CmbErrorActionFlag 
         Height          =   360
         ItemData        =   "FrmRpt_CashReceiptForNavision.frx":0000
         Left            =   4470
         List            =   "FrmRpt_CashReceiptForNavision.frx":000A
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.ComboBox CmbDataTransfer 
         Height          =   360
         ItemData        =   "FrmRpt_CashReceiptForNavision.frx":003E
         Left            =   2400
         List            =   "FrmRpt_CashReceiptForNavision.frx":0048
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   2010
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   12150
         TabIndex        =   16
         Top             =   1620
         Width           =   735
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   11415
         TabIndex        =   15
         Top             =   1620
         Width           =   735
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         Height          =   375
         Left            =   10200
         TabIndex        =   14
         Top             =   1620
         Width           =   1215
      End
      Begin VB.OptionButton OptPostingDate 
         Caption         =   "Posting Date"
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton OptCRExportDate 
         Caption         =   "CR Export Date"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   645
         Width           =   2175
      End
      Begin VB.CheckBox ChkReverseCR 
         Caption         =   "Reverse Cash Receipt"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1005
         Width           =   2295
      End
      Begin VB.ComboBox CmbCashReceiptType 
         Height          =   360
         ItemData        =   "FrmRpt_CashReceiptForNavision.frx":0062
         Left            =   5450
         List            =   "FrmRpt_CashReceiptForNavision.frx":006F
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   3120
      End
      Begin VB.ComboBox CmbExportedBy 
         Height          =   360
         Left            =   8600
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   3720
      End
      Begin VB.ComboBox CmbJournalCode 
         Height          =   360
         Left            =   12360
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   2520
      End
      Begin VB.TextBox TxtAccountNo 
         Height          =   360
         Left            =   5475
         TabIndex        =   9
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox TxtPINO 
         Height          =   360
         Left            =   2400
         TabIndex        =   8
         Top             =   1080
         Width           =   2925
      End
      Begin VB.TextBox TxtBatchNo 
         Height          =   360
         Left            =   8640
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TxtUserBatchID 
         Height          =   360
         Left            =   10560
         TabIndex        =   11
         Top             =   1080
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   345
         Left            =   2400
         TabIndex        =   2
         Top             =   475
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111738881
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   345
         Left            =   3960
         TabIndex        =   3
         Top             =   475
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111738881
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Select All"
         Height          =   240
         Left            =   360
         TabIndex        =   35
         Top             =   1755
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Action On Error Flag"
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
         Left            =   4545
         TabIndex        =   34
         Top             =   1560
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Navision Push Flag"
         Height          =   240
         Left            =   2475
         TabIndex        =   33
         Top             =   1560
         Width           =   1740
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   240
         Left            =   2655
         TabIndex        =   29
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   240
         Left            =   4320
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cash Receipt Type"
         Height          =   240
         Left            =   6140
         TabIndex        =   27
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Exported By"
         Height          =   240
         Left            =   9913
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Journal Code"
         Height          =   240
         Left            =   13027
         TabIndex        =   25
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Account No"
         Height          =   240
         Left            =   6495
         TabIndex        =   24
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "PI No"
         Height          =   240
         Left            =   3600
         TabIndex        =   23
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Batch No"
         Height          =   240
         Left            =   9180
         TabIndex        =   22
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "User Batch ID"
         Height          =   240
         Left            =   12112
         TabIndex        =   21
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label lblRec 
         Caption         =   "Total Record Found :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   20
         Top             =   1680
         Width           =   2325
      End
      Begin VB.Label lblRecordCount 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Left            =   9240
         TabIndex        =   19
         Top             =   1680
         Width           =   210
      End
   End
End
Attribute VB_Name = "FrmRpt_CashReceiptForNavision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCRType, mExportedBy, mExportedByID As Variant
Dim mJournalType, mJournalCode As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mUpdateFlag As Boolean
Dim mNavisionEditRole As Variant
Dim mEditRights As Boolean

Private Sub CmbExportedBy_GotFocus()
Dim mR As Variant
tmp1 = CmbExportedBy.Text
CmbExportedBy.Clear
tmp = " SELECT Distinct JCM.EMployeeID,ME.EmployeeNo,ME.FullName FROM Mst_EmpJournalCodeMapping JCM"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Order by ME.FullName"
Call TmpTableHRPro
If TmpRsHRPro.RecordCount > 0 Then
For mR = 1 To TmpRsHRPro.RecordCount
   CmbExportedBy.AddItem IIf(IsNull(TmpRsHRPro!FullName), "", TmpRsHRPro!FullName) & "~" & IIf(IsNull(TmpRsHRPro!EmployeeNo), "", TmpRsHRPro!EmployeeNo)
   TmpRsHRPro.MoveNext
Next
Else
   MsgBox "No Record found !!!! "
   Exit Sub
End If
CmbExportedBy = tmp1
End Sub

Private Sub CmbExportedBy_LostFocus()
Dim mSplit As Variant
If CmbExportedBy.Text = "" Then
   mExportedBy = ""
   mExportedByID = ""
   Exit Sub
End If
mSplit = Split(CmbExportedBy.Text, "~")
If UBound(mSplit) <> 1 Then
   MsgBox "Invalid selection!!! "
   CmbExportedBy.SetFocus
   Exit Sub
End If
TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "EmployeeNo = '" & mSplit(1) & "'"
If TmpRsHRPro.EOF Then
   MsgBox "Invalid selection "
   CmbJournalCode.SetFocus
   Exit Sub
End If
mExportedBy = TmpRsHRPro!EmployeeNo
mExportedByID = TmpRsHRPro!EmployeeID
End Sub

Private Sub CmbCashReceiptType_LostFocus()
    mCRType = ""
    If CmbCashReceiptType.Text = "" Then Exit Sub
    
    If LCase(CmbCashReceiptType.Text) = LCase("Cash Receipt") Then
        mCRType = "CR"
    ElseIf LCase(CmbCashReceiptType.Text) = LCase("Pledge Cash Receipt") Then
        mCRType = "PCR"
    ElseIf LCase(CmbCashReceiptType.Text) = LCase("Cash Receipt Adjustment") Then
        mCRType = "CRADJ"
    Else
       mCRType = ""
       MsgBox "Invalid Selection"
       CmbCashReceiptType.SetFocus
       Exit Sub
    End If
End Sub

Private Sub CmbJournalCode_GotFocus()
Dim mR As Variant
If Trim(CmbExportedBy.Text) = "" Then
    MsgBox "Select Exported By."
    CmbExportedBy.SetFocus
    Exit Sub
End If

tmp1 = CmbJournalCode
CmbJournalCode.Clear
tmp = " SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]" '''JT.JournalTypeName,JCM.JournalCode
tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where ME.EmployeeNo = '" & mExportedBy & "' Order by JCM.JournalTypeID"
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
mJournalCode = UCase(mSplit(1))
End Sub

Private Sub CmbServiceTax_LostFocus()
    If CmbServiceTax.Text = "" Then Exit Sub
    If LCase(CmbServiceTax.Text) <> LCase("Service Tax") And LCase(CmbServiceTax.Text) <> LCase("No Service Tax") Then
       MsgBox "Invalid Selection"
       CmbServiceTax.SetFocus
       Exit Sub
    End If
End Sub

Private Sub CmdClear_Click()
    CmdGetReport.Enabled = True
    CmdExcel.Enabled = True
    TxtFromDate.Value = CDate(Date) - 30
    TxtToDate.Value = CDate(Date)
    CmbCashReceiptType.Text = ""
    CmbExportedBy.Text = ""
    CmbJournalCode.Text = ""
    ChkReverseCR.Value = vbUnchecked
    TxtPINO = ""
    TxtAccountNo = ""
    TxtBatchNo = ""
    TxtUserBatchID = ""
    lblRecordCount.Caption = "-"
    ProgressBar1.Value = False
    
    CmbDataTransfer.Text = ""
    CmdRemoveErrorFlag.Enabled = False
    CmbErrorActionFlag.Text = ""
    If mEditRights = True Then
        CmdRemoveErrorFlag.Visible = True
        ChkSelectAll.Visible = True
        ChkSelectAll.Value = vbUnchecked
        Label14.Visible = True
        CmbErrorActionFlag.Visible = True
        Label20.Visible = True
    Else
        CmdRemoveErrorFlag.Visible = False
        ChkSelectAll.Visible = False
        ChkSelectAll.Value = vbUnchecked
        Label14.Visible = False
        CmbErrorActionFlag.Visible = False
        Label20.Visible = False
    End If
    Call Grid_Head
End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant
Dim mR As Double
If TxtFromDate.Value > TxtToDate.Value Then
    MsgBox "Invalid Date."
    TxtFromDate.SetFocus
    Exit Sub
End If
Call Grid_Head
lblRecordCount.Caption = "-"
CmdExcel.Enabled = False
''-- Cash Receipt
wc = ""
tmp = " SELECT TCRFN.PostingDate,TCRFN.DocumentType,TCRFN.AccountType,TCRFN.Amount,TCRFN.AccountNo,TCRFN.InvoiceNo,TCRFN.Narration,TCRFN.RegionDim,"
tmp = tmp & " TCRFN.DeptDim,TCRFN.EmpDim,TCRFN.ChequeNo,TCRFN.ChequeDate,TCRFN.PINo,TCRFN.CRADJEntryNo,TCRFN.PostInvoiceNo,TCRFN.EntryNo,TCRFN.BatchID,"
tmp = tmp & " TCRFN.JournalBatchType,TCRFN.JournalBatchName,TCRFN.InvoiceType,TCRFN.ReverseInvoice,TCRFN.DataTransfer , TCRFN.EntryNoTransfer, TCRFN.Remark,"
tmp = tmp & " TCRFN.Concurrency, TCRFN.CreatedBy, TCRFN.[LineNo], TCRFN.UserBatchID FROM Txn_CashRepeiptsForNAVISION TCRFN "

''-- Date
If OptPostingDate.Value = True Then
    wc = " Where TCRFN.PostingDate BETWEEN '" & Format(TxtFromDate, "YYYY-MM-DD") & "' AND '" & Format(TxtToDate, "YYYY-MM-DD") & "'"
Else
    wc = " Where CONVERT(VARCHAR,TCRFN.Concurrency,112) BETWEEN '" & Format(TxtFromDate, "YYYYMMDD") & "' AND '" & Format(TxtToDate, "YYYYMMDD") & "'"
End If
''-- Cash Receipt Type
If CmbCashReceiptType.Text <> "" Then
    wc = wc & " And TCRFN.InvoiceType = '" & mCRType & "' "
End If
''-- Exported By
If CmbExportedBy.Text <> "" Then
    wc = wc & " And TCRFN.CreatedBy = '" & mExportedBy & "' "
End If
''-- Journal Code
If CmbJournalCode.Text <> "" Then
    wc = wc & " And (TCRFN.JournalBatchType = '" & mJournalType & "' And TCRFN.JournalBatchName = '" & mJournalCode & "') "
End If
''-- Reverse CR
If ChkReverseCR.Value = vbChecked Then
    wc = wc & " And Isnull(TCRFN.ReverseInvoice,'N') = 'Y' "
End If
''-- Navision Push Flag
If CmbDataTransfer.Text <> "" Then
    If LCase(CmbDataTransfer.Text) = LCase("Data Transfer") Then
        wc = wc & " And Isnull(TCRFN.DataTransfer,'') = 'T' "
    Else
        wc = wc & " And Isnull(TCRFN.DataTransfer,'') = 'E' "
    End If
End If
''-- PI No
If Trim(TxtPINO) <> "" Then
    wc = wc & " And TCRFN.PINo like '%" & Trim(TxtPINO) & "%' "
End If
''-- Account No
If Trim(TxtAccountNo) <> "" Then
    wc = wc & " And TCRFN.AccountNo like '%" & Trim(TxtAccountNo) & "%' "
End If
''-- Batch No
If Trim(TxtBatchNo) <> "" Then
    wc = wc & " And TCRFN.BatchID = " & Val(TxtBatchNo) & " "
End If
''-- USER Batch ID
If Trim(TxtUserBatchID) <> "" Then
    wc = wc & " And TCRFN.UserBatchID like '%" & Trim(TxtUserBatchID) & "%' "
End If

tmp = tmp & wc
tmp = tmp & " Order By TCRFN.[LineNo] "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   With MSHFlexGrid1
     ProgressBar1.Visible = True
     ProgressBar1.Max = TmpRs.RecordCount
     For mR = 1 To TmpRs.RecordCount
        
        If mEditRights = True Then
            .row = mR
            .Col = 0
            .CellFontName = "Wingdings"
            .CellFontSize = 12
            .CellAlignment = flexAlignCenterCenter
            .Text = strUnChecked
            .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!PostingDate), "", Format(TmpRs!PostingDate, "DD-MMM-YYYY"))
            .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!DocumentType), "", TmpRs!DocumentType)
            .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!AccountType), "", TmpRs!AccountType)
            .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!AccountNo), "", TmpRs!AccountNo)
            .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
            .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!InvoiceNo), "", TmpRs!InvoiceNo)
            .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!Narration), "", TmpRs!Narration)
            .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!RegionDim), "", TmpRs!RegionDim)
            .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!DeptDim), "", TmpRs!DeptDim)
            .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!EmpDim), "", TmpRs!EmpDim)
            .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!ChequeNo), "", TmpRs!ChequeNo)
            .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!ChequeDate), "", TmpRs!ChequeDate)
            .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!PINo), "", TmpRs!PINo)
            .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
            .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!PostInvoiceNo), "", TmpRs!PostInvoiceNo)
            .TextMatrix(mR, 16) = IIf(IsNull(TmpRs!CRADJEntryNo), "", TmpRs!CRADJEntryNo)
            .TextMatrix(mR, 17) = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID)
            .TextMatrix(mR, 18) = IIf(IsNull(TmpRs!JournalBatchType), "", TmpRs!JournalBatchType)
            .TextMatrix(mR, 19) = IIf(IsNull(TmpRs!JournalBatchName), "", TmpRs!JournalBatchName)
            .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!InvoiceType), "", TmpRs!InvoiceType)
            .TextMatrix(mR, 21) = IIf(IsNull(TmpRs!ReverseInvoice), "", TmpRs!ReverseInvoice)
            .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!DataTransfer), "", TmpRs!DataTransfer)
            .TextMatrix(mR, 23) = IIf(IsNull(TmpRs!EntryNoTransfer), "", TmpRs!EntryNoTransfer)
            .TextMatrix(mR, 24) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
            .TextMatrix(mR, 25) = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & ":-" & IIf(IsNull(TmpRs!Concurrency), "", TmpRs!Concurrency)
            .TextMatrix(mR, 26) = IIf(IsNull(TmpRs!LineNo), "", TmpRs!LineNo)
            .TextMatrix(mR, 27) = IIf(IsNull(TmpRs!UserBatchID), "", TmpRs!UserBatchID)
        Else
            .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!PostingDate), "", Format(TmpRs!PostingDate, "DD-MMM-YYYY"))
            .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!DocumentType), "", TmpRs!DocumentType)
            .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!AccountType), "", TmpRs!AccountType)
            .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!AccountNo), "", TmpRs!AccountNo)
            .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
            .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!InvoiceNo), "", TmpRs!InvoiceNo)
            .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!Narration), "", TmpRs!Narration)
            .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!RegionDim), "", TmpRs!RegionDim)
            .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!DeptDim), "", TmpRs!DeptDim)
            .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!EmpDim), "", TmpRs!EmpDim)
            .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!ChequeNo), "", TmpRs!ChequeNo)
            .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!ChequeDate), "", TmpRs!ChequeDate)
            .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!PINo), "", TmpRs!PINo)
            .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
            .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!PostInvoiceNo), "", TmpRs!PostInvoiceNo)
            .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!CRADJEntryNo), "", TmpRs!CRADJEntryNo)
            .TextMatrix(mR, 16) = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID)
            .TextMatrix(mR, 17) = IIf(IsNull(TmpRs!JournalBatchType), "", TmpRs!JournalBatchType)
            .TextMatrix(mR, 18) = IIf(IsNull(TmpRs!JournalBatchName), "", TmpRs!JournalBatchName)
            .TextMatrix(mR, 19) = IIf(IsNull(TmpRs!InvoiceType), "", TmpRs!InvoiceType)
            .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!ReverseInvoice), "", TmpRs!ReverseInvoice)
            .TextMatrix(mR, 21) = IIf(IsNull(TmpRs!DataTransfer), "", TmpRs!DataTransfer)
            .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!EntryNoTransfer), "", TmpRs!EntryNoTransfer)
            .TextMatrix(mR, 23) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
            .TextMatrix(mR, 24) = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & ":-" & IIf(IsNull(TmpRs!Concurrency), "", TmpRs!Concurrency)
            .TextMatrix(mR, 25) = IIf(IsNull(TmpRs!LineNo), "", TmpRs!LineNo)
            .TextMatrix(mR, 26) = IIf(IsNull(TmpRs!UserBatchID), "", TmpRs!UserBatchID)
        End If
     .Rows = .Rows + 1
     ProgressBar1.Value = mR
     TmpRs.MoveNext
     Next mR
   End With
End If
lblRecordCount.Caption = TmpRs.RecordCount
CmdExcel.Enabled = True
MsgBox "Done."
End Sub

Private Sub CmdRemoveErrorFlag_Click()
Dim mR, mRow As Variant
Dim mTmpDocNo, mTmpUserBatchID, mTmpLineNo As Variant
Dim mTmpInvoiceType, mTmpReverse, mTmpAccountType, mTmpProjectType, mTmpBatchNo, mTmpPINo, mTmpClaimId As Variant
Dim mActionFlag As Variant
Dim strSplitString, PCRID_, Prefix_ As Variant

mActionFlag = ""
If CmbErrorActionFlag.Text = "" Then
   MsgBox "Select Action On Error Flag."
   CmbErrorActionFlag.SetFocus
   Exit Sub
Else
    If LCase(CmbErrorActionFlag.Text) = LCase("Remove Error Flag") Then
        mActionFlag = "Error"
    ElseIf LCase(CmbErrorActionFlag.Text) = LCase("Pending Entry for Re-Export") Then
        mActionFlag = "Pending"
    Else
        mActionFlag = ""
    End If
End If

Select Case MsgBox("Do you want to '" & CmbErrorActionFlag.Text & "'", vbYesNo)
    Case vbYes
        CmdRemoveErrorFlag.Enabled = False
        If mEditRights = True Then
            With MSHFlexGrid1
               If LCase(mActionFlag) = LCase("Error") Then
                    ''--- Remove Error Flag From Exported Table Only
                    For mR = 1 To .Rows - 2
                        mTmpPINo = ""
                        mTmpUserBatchID = ""
                        mTmpLineNo = ""
                        If .TextMatrix(mR, 0) = strChecked Then
                            mTmpPINo = .TextMatrix(mR, 13)
                            mTmpUserBatchID = .TextMatrix(mR, 27)
                            For mRow = 1 To .Rows - 2
                                If .TextMatrix(mRow, 0) = strChecked And mTmpPINo = .TextMatrix(mRow, 13) And mTmpUserBatchID = .TextMatrix(mRow, 27) And LCase(.TextMatrix(mRow, 22)) = LCase("E") Then
                                    mTmpLineNo = mTmpLineNo & "," & Val(.TextMatrix(mRow, 26))
                                End If
                            Next mRow
                            If mTmpLineNo <> "" Then
                                mTmpLineNo = Right(mTmpLineNo, Len(mTmpLineNo) - 1)
                                ''tmp = " SELECT * FROM Txn_CashRepeiptsForNAVISION Where [LineNo] In (" & mTmpLineNo & ") And DataTransfer = 'E' And ISNULL(PostInvoiceNo,'') = '' And UserBatchID = '" & mTmpUserBatchID & "' And PINo = '" & mTmpPINo & "'"
                                tmp = " UPDATE Txn_CashRepeiptsForNAVISION SET DataTransfer = Null,Remark = '' Where [LineNo] In (" & mTmpLineNo & ") And DataTransfer = 'E' And ISNULL(PostInvoiceNo,'') = '' And UserBatchID = '" & mTmpUserBatchID & "' And PINo = '" & mTmpPINo & "'"
                                Call TmpTable
                            End If
                        End If
                    Next mR
               ElseIf LCase(mActionFlag) = LCase("Pending") Then
                    ''--- Mark pending Source Entry so entry can export again
                    For mR = 1 To .Rows - 2
                        mTmpUserBatchID = ""
                        mTmpLineNo = ""
                        mTmpInvoiceType = ""
                        mTmpBatchNo = ""
                        mTmpClaimId = ""
                        mTmpPINo = ""
                        mTmpReverse = ""
                        mTmpAccountType = ""
                        If .TextMatrix(mR, 0) = strChecked Then
                            mTmpPINo = .TextMatrix(mR, 13)
                            mTmpUserBatchID = .TextMatrix(mR, 27)
                            mTmpDocumentType = .TextMatrix(mR, 2)
                            mTmpAccountType = .TextMatrix(mR, 3)
                            mTmpBatchNo = .TextMatrix(mR, 17)
                            mTmpReverse = .TextMatrix(mR, 21)
                            mTmpInvoiceType = .TextMatrix(mR, 20)
                            
                            If LCase(mTmpReverse) = "" And LCase(mTmpInvoiceType) = LCase("CR") Then
                            ''--- Normal CR Entry
                                ''tmp = " SELECT * FROM Txn_NewCashReceipt Where ISNULL(BatchID,0) <> 0 And ISNULL(RBatchID,0) = 0 And BatchID = " & Val(mTmpBatchNo) & " And CRID = " & Val(Replace(mTmpPINo, "CR", "")) & " "
                                tmp = " Update Txn_NewCashReceipt set BatchID = 0 Where ISNULL(BatchID,0) <> 0 And ISNULL(RBatchID,0) = 0 And BatchID = " & Val(mTmpBatchNo) & " And CRID = " & Val(Replace(mTmpPINo, "CR", "")) & " "
                                Call TmpTable
                            ElseIf LCase(mTmpReverse) = "" And LCase(mTmpInvoiceType) = LCase("PCR") Then
                            ''--- Pledge CR Entry
                                strSplitString = Split(Replace(mTmpPINo, "PCR", ""), "-")
                                PCRID_ = strSplitString(0)
                                Prefix_ = strSplitString(1)
                                ''tmp = " SELECT * FROM Txn_PledgeCashReceipt WHERE Flag = 'P' And ISNULL(BatchID,0) <> 0 And BatchID = " & Val(mTmpBatchNo) & " And ISNULL(RBatchID,0) = 0 And ReturnFlag = 'N' And PCRID = " & Val(PCRID_) & " And SM_Prefix = '" & Prefix_ & "' "
                                tmp = " Update Txn_PledgeCashReceipt set BatchID = 0 WHERE Flag = 'P' And ISNULL(BatchID,0) <> 0 And BatchID = " & Val(mTmpBatchNo) & " And ISNULL(RBatchID,0) = 0 And ReturnFlag = 'N' And PCRID = " & Val(PCRID_) & " And SM_Prefix = '" & Prefix_ & "' "
                                Call TmpTable
                            ElseIf LCase(mTmpReverse) = LCase("") And LCase(mTmpInvoiceType) = LCase("CRADJ") Then
                            ''--- CR ADJ Entry
                                ''tmp = " SELECT * FROM Txn_CRTDSAdjustment Where ISNULL(BatchID,0) <> 0 And BatchID = " & Val(mTmpBatchNo) & " And TDSAID = " & Val(Replace(mTmpPINo, "AD", "")) & " "
                                tmp = " Update Txn_CRTDSAdjustment set BatchID = 0 Where ISNULL(BatchID,0) <> 0 And BatchID = " & Val(mTmpBatchNo) & " And TDSAID = " & Val(Replace(mTmpPINo, "TDS", "")) & " "
                                Call TmpTable
                            End If
                        End If
                    Next mR
               End If
            End With
        End If
        MsgBox "Done."
        Call Grid_Head
End Select
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
''--- Edit Role for Specific Users Only by EmailConfig INI File
mNavisionEditRole = Split(Replace(LCase(Gen_ReadNavisionEditUserRole), "navisioneditrole-", ""), ",")
mEditRights = False
If UBound(mNavisionEditRole) >= 0 Then
For mR = 0 To UBound(mNavisionEditRole)
    If LCase(Gen_UserLoginID) = LCase(mNavisionEditRole(mR)) Then
        mEditRights = True
    End If
Next mR
End If
''---
Call CmdClear_Click
CmdExcel.Enabled = False
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0
    .WordWrap = True
    If mEditRights = True Then
        .Cols = 28
        .TextMatrix(0, 0) = "Select":                   .ColWidth(0) = 800
        .TextMatrix(0, 1) = "Posting Date":             .ColWidth(1) = 1200
        .TextMatrix(0, 2) = "Document Type":            .ColWidth(2) = 1500
        .TextMatrix(0, 3) = "Account Type":             .ColWidth(3) = 1500
        .TextMatrix(0, 4) = "Account No":               .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "Amount":                   .ColWidth(5) = 1500
        .TextMatrix(0, 6) = "Invoice No":               .ColWidth(6) = 1500
        .TextMatrix(0, 7) = "Narration":                .ColWidth(7) = 2500
        .TextMatrix(0, 8) = "Region Dim":               .ColWidth(8) = 1500
        .TextMatrix(0, 9) = "Dept Dim":                 .ColWidth(9) = 1500
        .TextMatrix(0, 10) = "Emp Dim":                 .ColWidth(10) = 1500
        .TextMatrix(0, 11) = "Cheque No":               .ColWidth(11) = 1500
        .TextMatrix(0, 12) = "Cheque Date":             .ColWidth(12) = 1500
        .TextMatrix(0, 13) = "PI No":                   .ColWidth(13) = 1500
        .TextMatrix(0, 14) = "CR Adj Entry No":         .ColWidth(14) = 1500    ''-- Reversal Entry No
        .TextMatrix(0, 15) = "Post Invoice No":         .ColWidth(15) = 2000
        .TextMatrix(0, 16) = "Entry No":                .ColWidth(16) = 1500
        .TextMatrix(0, 17) = "Batch ID":                .ColWidth(17) = 1500
        .TextMatrix(0, 18) = "JourNal Batch Type":      .ColWidth(18) = 1800
        .TextMatrix(0, 19) = "JourNal Batch Name":      .ColWidth(19) = 1800
        .TextMatrix(0, 20) = "Cash Receipt Type":       .ColWidth(20) = 1800
        .TextMatrix(0, 21) = "Reverse CR (Y/N)":        .ColWidth(21) = 1000
        .TextMatrix(0, 22) = "Navision Push Flag":      .ColWidth(22) = 1000
        .TextMatrix(0, 23) = "Navision Pull Flag":      .ColWidth(23) = 1000
        .TextMatrix(0, 24) = "Remark":                  .ColWidth(24) = 2500
        .TextMatrix(0, 25) = "Created By":              .ColWidth(25) = 2500
        .TextMatrix(0, 26) = "Line No":                 .ColWidth(26) = 1500
        .TextMatrix(0, 27) = "User Batch ID":           .ColWidth(27) = 2500
    Else
        .Cols = 27
        .TextMatrix(0, 0) = "Posting Date":             .ColWidth(0) = 1200
        .TextMatrix(0, 1) = "Document Type":            .ColWidth(1) = 1500
        .TextMatrix(0, 2) = "Account Type":             .ColWidth(2) = 1500
        .TextMatrix(0, 3) = "Account No":               .ColWidth(3) = 1500
        .TextMatrix(0, 4) = "Amount":                   .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "Invoice No":               .ColWidth(5) = 1500
        .TextMatrix(0, 6) = "Narration":                .ColWidth(6) = 2500
        .TextMatrix(0, 7) = "Region Dim":               .ColWidth(7) = 1500
        .TextMatrix(0, 8) = "Dept Dim":                 .ColWidth(8) = 1500
        .TextMatrix(0, 9) = "Emp Dim":                  .ColWidth(9) = 1500
        .TextMatrix(0, 10) = "Cheque No":               .ColWidth(10) = 1500
        .TextMatrix(0, 11) = "Cheque Date":             .ColWidth(11) = 1500
        .TextMatrix(0, 12) = "PI No":                   .ColWidth(12) = 1500
        .TextMatrix(0, 13) = "CR Adj Entry No":         .ColWidth(13) = 1500    ''-- Reversal Entry No
        .TextMatrix(0, 14) = "Post Invoice No":         .ColWidth(14) = 2000
        .TextMatrix(0, 15) = "Entry No":                .ColWidth(15) = 1500
        .TextMatrix(0, 16) = "Batch ID":                .ColWidth(16) = 1500
        .TextMatrix(0, 17) = "JourNal Batch Type":      .ColWidth(17) = 1800
        .TextMatrix(0, 18) = "JourNal Batch Name":      .ColWidth(18) = 1800
        .TextMatrix(0, 19) = "Cash Receipt Type":       .ColWidth(19) = 1800
        .TextMatrix(0, 20) = "Reverse CR (Y/N)":        .ColWidth(20) = 1000
        .TextMatrix(0, 21) = "Navision Push Flag":      .ColWidth(21) = 1000
        .TextMatrix(0, 22) = "Navision Pull Flag":      .ColWidth(22) = 1000
        .TextMatrix(0, 23) = "Remark":                  .ColWidth(23) = 2500
        .TextMatrix(0, 24) = "Created By":              .ColWidth(24) = 2500
        .TextMatrix(0, 25) = "Line No":                 .ColWidth(25) = 1500
        .TextMatrix(0, 26) = "User Batch ID":           .ColWidth(26) = 2500
    End If
    .RowHeight(0) = 800
End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CashReceiptForNavision.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CashReceiptForNavision.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

With MSHFlexGrid1
ProgressBar1.Max = MSHFlexGrid1.Rows - 1
    For I = 0 To .Rows - 1
        mRow = I + 1 ' + 6
        For C = 0 To .Cols - 1
            If InStr(.TextMatrix(I, C), "/") > 0 Then
               If IsDate(.TextMatrix(I, C)) Then
                  oWS.Cells(mRow, C + 1) = Format(.TextMatrix(I, C), "MM/dd/yyyy")
               Else
                  oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
               End If
            Else
               oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
            End If
        Next
        ProgressBar1.Value = I
    Next
End With
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub CmbDataTransfer_LostFocus()
    If CmbDataTransfer.Text = "" Then Exit Sub
    If LCase(CmbDataTransfer.Text) <> LCase("Data Transfer") And LCase(CmbDataTransfer.Text) <> LCase("Error") Then
       MsgBox "Invalid Selection"
       CmbDataTransfer.SetFocus
       Exit Sub
    End If
End Sub
Private Sub CmbErrorActionFlag_LostFocus()
    If CmbErrorActionFlag.Text = "" Then Exit Sub
    If LCase(CmbErrorActionFlag.Text) <> LCase("Remove Error Flag") And LCase(CmbErrorActionFlag.Text) <> LCase("Pending Entry for Re-Export") Then
       MsgBox "Invalid Selection"
       CmbErrorActionFlag.SetFocus
       Exit Sub
    End If
End Sub

Private Sub ChkSelectAll_Click()
Dim mR, mTempPINo As Variant
If MSHFlexGrid1.Rows = 2 Then
   ChkSelectAll.Value = vbUnchecked
   Exit Sub
End If
With MSHFlexGrid1
If ChkSelectAll.Value = vbChecked Then
   For I = 1 To .Rows - 2
       .row = I
       mTempPINo = Trim(.TextMatrix(I, 13))
       If Trim(.TextMatrix(I, 4)) <> "" And Trim(.TextMatrix(I, 15)) = "" And LCase(.TextMatrix(I, 22)) = LCase("E") And Trim(.TextMatrix(I, 13)) <> "" Then
          .row = I
          .TextMatrix(I, 0) = strChecked
          For mR = I To .Rows - 2
            If mTempPINo = Trim(.TextMatrix(mR, 13)) And LCase(.TextMatrix(mR, 22)) = LCase("E") Then
                .TextMatrix(mR, .Col) = strChecked
            End If
          Next mR
      End If
   Next
Else
   For I = 1 To .Rows - 2
       .row = I
       .TextMatrix(I, 0) = strUnChecked
   Next
End If

mUpdateFlag = False
For mR = 1 To .Rows - 2
    If .TextMatrix(mR, 0) = strChecked And LCase(.TextMatrix(mR, 22)) = LCase("E") Then
        mUpdateFlag = True
    End If
Next mR
If mUpdateFlag = True Then
    CmdRemoveErrorFlag.Enabled = True
Else
    CmdRemoveErrorFlag.Enabled = False
End If
End With
End Sub


Private Sub MSHFlexGrid1_Click()
Dim mR, mTempPINo As Variant
If mEditRights = True Then
    With MSHFlexGrid1
         If .Col = 0 Then
            If .TextMatrix(.row, .Col) = "" Then Exit Sub
            .Col = 0
            mTempPINo = Trim(.TextMatrix(.row, 13))
            If .TextMatrix(.row, .Col) = strUnChecked Then
               If Trim(.TextMatrix(.row, 4)) <> "" And Trim(.TextMatrix(.row, 15)) = "" And LCase(.TextMatrix(.row, 22)) = LCase("E") And Trim(.TextMatrix(.row, 13)) <> "" Then
                  .TextMatrix(.row, .Col) = strChecked
                  For mR = 1 To .Rows - 2
                    If mTempPINo = Trim(.TextMatrix(mR, 13)) And LCase(.TextMatrix(mR, 22)) = LCase("E") Then
                        .TextMatrix(mR, .Col) = strChecked
                    End If
                  Next mR
               End If
            Else
                .TextMatrix(.row, .Col) = strUnChecked
                For mR = 1 To .Rows - 2
                  If mTempPINo = Trim(.TextMatrix(mR, 13)) Then
                      .TextMatrix(mR, .Col) = strUnChecked
                  End If
                Next mR
            End If
         End If
         
    mUpdateFlag = False
    For mR = 1 To .Rows - 2
        If .TextMatrix(mR, 0) = strChecked And LCase(.TextMatrix(mR, 22)) = LCase("E") Then
            mUpdateFlag = True
        End If
    Next mR
    If mUpdateFlag = True Then
        CmdRemoveErrorFlag.Enabled = True
    Else
        CmdRemoveErrorFlag.Enabled = False
    End If
    End With
End If
End Sub

