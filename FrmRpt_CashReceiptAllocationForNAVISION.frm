VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CashReceiptAllocationForNAVISION 
   Caption         =   "Cash Receipt Allocation For Navision Report"
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
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7455
      Left            =   120
      TabIndex        =   25
      Top             =   2340
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   26
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
         TabIndex        =   12
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
      Height          =   2355
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   15015
      Begin VB.CommandButton CmdRemoveErrorFlag 
         Caption         =   "Remove Error Flag"
         Height          =   375
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1740
         Visible         =   0   'False
         Width           =   2080
      End
      Begin VB.ComboBox CmbDataTransfer 
         Height          =   360
         ItemData        =   "FrmRpt_CashReceiptAllocationForNAVISION.frx":0000
         Left            =   1320
         List            =   "FrmRpt_CashReceiptAllocationForNAVISION.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   1845
         Width           =   2370
      End
      Begin VB.ComboBox CmbErrorActionFlag 
         Height          =   360
         ItemData        =   "FrmRpt_CashReceiptAllocationForNAVISION.frx":0024
         Left            =   3750
         List            =   "FrmRpt_CashReceiptAllocationForNAVISION.frx":002E
         TabIndex        =   28
         Top             =   1845
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.CheckBox ChkSelectAll 
         Height          =   240
         Left            =   120
         TabIndex        =   27
         Top             =   1650
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox TxtUserBatchID 
         Height          =   360
         Left            =   10800
         TabIndex        =   8
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox TxtBatchNo 
         Height          =   360
         Left            =   6315
         TabIndex        =   7
         Top             =   1080
         Width           =   4440
      End
      Begin VB.TextBox TxtPaymentEntryNo 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2925
      End
      Begin VB.TextBox TxtInvoiceEntryNo 
         Height          =   360
         Left            =   3285
         TabIndex        =   6
         Top             =   1080
         Width           =   2925
      End
      Begin VB.ComboBox CmbJournalCode 
         Height          =   360
         Left            =   10800
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   4080
      End
      Begin VB.ComboBox CmbExportedBy 
         Height          =   360
         Left            =   6315
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   4440
      End
      Begin VB.ComboBox CmbCashReceiptType 
         Height          =   360
         ItemData        =   "FrmRpt_CashReceiptAllocationForNAVISION.frx":0062
         Left            =   3285
         List            =   "FrmRpt_CashReceiptAllocationForNAVISION.frx":006F
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2925
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         Height          =   375
         Left            =   9600
         TabIndex        =   9
         Top             =   1740
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   10815
         TabIndex        =   10
         Top             =   1740
         Width           =   975
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   11790
         TabIndex        =   11
         Top             =   1740
         Width           =   975
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   345
         Left            =   120
         TabIndex        =   0
         Top             =   480
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
         Format          =   59113473
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   345
         Left            =   1680
         TabIndex        =   1
         Top             =   480
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
         Format          =   59113473
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Navision Push Flag"
         Height          =   240
         Left            =   1635
         TabIndex        =   32
         Top             =   1560
         Width           =   1740
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
         Left            =   3945
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Select All"
         Height          =   240
         Left            =   360
         TabIndex        =   30
         Top             =   1650
         Visible         =   0   'False
         Width           =   1080
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
         Left            =   8760
         TabIndex        =   24
         Top             =   1800
         Width           =   210
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
         Left            =   6360
         TabIndex        =   23
         Top             =   1800
         Width           =   2325
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "User Batch ID"
         Height          =   240
         Left            =   12345
         TabIndex        =   22
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Batch No"
         Height          =   240
         Left            =   8220
         TabIndex        =   21
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Payment Entry No"
         Height          =   240
         Left            =   840
         TabIndex        =   20
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Invoice Entry No"
         Height          =   240
         Left            =   4215
         TabIndex        =   19
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Journal Code"
         Height          =   240
         Left            =   12240
         TabIndex        =   18
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Exported By"
         Height          =   240
         Left            =   8048
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cash Receipt Type"
         Height          =   240
         Left            =   3975
         TabIndex        =   16
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   240
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   240
         Left            =   375
         TabIndex        =   14
         Top             =   240
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmRpt_CashReceiptAllocationForNAVISION"
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

Private Sub ChkSelectAll_Click()
Dim mR, mTempPaymentEntryNo, mTempInvoiceEntryNo As Variant
If MSHFlexGrid1.Rows = 2 Then
   ChkSelectAll.Value = vbUnchecked
   Exit Sub
End If
With MSHFlexGrid1
If ChkSelectAll.Value = vbChecked Then
   For I = 1 To .Rows - 2
       .row = I
       mTempPaymentEntryNo = Trim(.TextMatrix(I, 1))
       mTempInvoiceEntryNo = Trim(.TextMatrix(I, 2))
       If Trim(.TextMatrix(I, 1)) <> "" And Trim(.TextMatrix(I, 2)) <> "" And Trim(.TextMatrix(I, 3)) <> "" And LCase(.TextMatrix(I, 8)) = LCase("E") Then
          .row = I
          .TextMatrix(I, 0) = strChecked
          For mR = I To .Rows - 2
            If mTempPaymentEntryNo = Trim(.TextMatrix(mR, 1)) And mTempInvoiceEntryNo = Trim(.TextMatrix(mR, 2)) And LCase(.TextMatrix(mR, 8)) = LCase("E") Then
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
    If .TextMatrix(mR, 0) = strChecked And LCase(.TextMatrix(mR, 8)) = LCase("E") Then
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
    ElseIf LCase(CmbCashReceiptType.Text) = LCase("Cash Receipt Advance") Then
        mCRType = "CRA"
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
    TxtPaymentEntryNo = ""
    TxtInvoiceEntryNo = ""
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

''--Invoice Allocation
wc = ""
tmp = " SELECT TIAFN.PaymentEntryNo,TIAFN.InvoiceEntryNo,TIAFN.Amount,TIAFN.BatchID,TIAFN.JournalBatchType,TIAFN.JournalBatchName,"
tmp = tmp & " TIAFN.InvoiceType , TIAFN.DataTransfer, TIAFN.EntryNoTransfer, TIAFN.Remark, TIAFN.Concurrency, TIAFN.CreatedBy, TIAFN.[LineNo], TIAFN.UserBatchID"
tmp = tmp & " FROM Txn_InvoiceAllocationForNAVISION TIAFN "

''-- Export Date
'wc = " Where TCRFN.PostingDate BETWEEN '" & Format(TxtFromDate, "YYYY-MM-DD") & "' AND '" & Format(TxtToDate, "YYYY-MM-DD") & "'"
wc = " Where  CONVERT(VARCHAR,TIAFN.Concurrency,112) BETWEEN '" & Format(TxtFromDate, "YYYYMMDD") & "' AND '" & Format(TxtToDate, "YYYYMMDD") & "'"
''-- Cash Receipt Type
If CmbCashReceiptType.Text <> "" Then
    wc = wc & " And TIAFN.InvoiceType = '" & mCRType & "' "
End If
''-- Exported By
If CmbExportedBy.Text <> "" Then
    wc = wc & " And TIAFN.CreatedBy = '" & mExportedBy & "' "
End If
''-- Journal Code
If CmbJournalCode.Text <> "" Then
    wc = wc & " And (TIAFN.JournalBatchType = '" & mJournalType & "' And TIAFN.JournalBatchName = '" & mJournalCode & "') "
End If
''--Payment Entry No
If Trim(TxtPaymentEntryNo) <> "" Then
    wc = wc & " And TIAFN.PaymentEntryNo like '%" & Trim(TxtPaymentEntryNo) & "%' "
End If
''--Invoice Entry No
If Trim(TxtInvoiceEntryNo) <> "" Then
    wc = wc & " And TIAFN.InvoiceEntryNo like '%" & Trim(TxtInvoiceEntryNo) & "%' "
End If
''-- Batch No
If Trim(TxtBatchNo) <> "" Then
    wc = wc & " And TIAFN.BatchID = " & Val(TxtBatchNo) & " "
End If
''-- USER Batch ID
If Trim(TxtUserBatchID) <> "" Then
    wc = wc & " And TIAFN.UserBatchID like '%" & Trim(TxtUserBatchID) & "%' "
End If
''-- Navision Push Flag
If CmbDataTransfer.Text <> "" Then
    If LCase(CmbDataTransfer.Text) = LCase("Data Transfer") Then
        wc = wc & " And Isnull(TIAFN.DataTransfer,'') = 'T' "
    Else
        wc = wc & " And Isnull(TIAFN.DataTransfer,'') = 'E' "
    End If
End If


tmp = tmp & wc
tmp = tmp & " Order By TIAFN.[LineNo] "
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
            .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!PaymentEntryNo), "", TmpRs!PaymentEntryNo)
            .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!InvoiceEntryNo), "", TmpRs!InvoiceEntryNo)
            .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
            .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID)
            .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!JournalBatchType), "", TmpRs!JournalBatchType)
            .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!JournalBatchName), "", TmpRs!JournalBatchName)
            .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!InvoiceType), "", TmpRs!InvoiceType)
            .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!DataTransfer), "", TmpRs!DataTransfer)
            .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!EntryNoTransfer), "", TmpRs!EntryNoTransfer)
            .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
            .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & ":-" & IIf(IsNull(TmpRs!Concurrency), "", TmpRs!Concurrency)
            .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!LineNo), "", TmpRs!LineNo)
            .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!UserBatchID), "", TmpRs!UserBatchID)
        Else
            .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!PaymentEntryNo), "", TmpRs!PaymentEntryNo)
            .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!InvoiceEntryNo), "", TmpRs!InvoiceEntryNo)
            .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
            .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID)
            .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!JournalBatchType), "", TmpRs!JournalBatchType)
            .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!JournalBatchName), "", TmpRs!JournalBatchName)
            .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!InvoiceType), "", TmpRs!InvoiceType)
            .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!DataTransfer), "", TmpRs!DataTransfer)
            .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!EntryNoTransfer), "", TmpRs!EntryNoTransfer)
            .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
            .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & ":-" & IIf(IsNull(TmpRs!Concurrency), "", TmpRs!Concurrency)
            .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!LineNo), "", TmpRs!LineNo)
            .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!UserBatchID), "", TmpRs!UserBatchID)
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
Dim mTmpPaymentEntryNo, mTmpInvoiceEntryNo, mTmpUserBatchID, mTmpLineNo, mTmpBatchNo, mTmpInvoiceType As Variant
Dim mActionFlag As Variant

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
                        mTmpPaymentEntryNo = ""
                        mTmpInvoiceEntryNo = ""
                        mTmpUserBatchID = ""
                        mTmpLineNo = ""
                        If .TextMatrix(mR, 0) = strChecked Then
                            mTmpPaymentEntryNo = .TextMatrix(mR, 1)
                            mTmpInvoiceEntryNo = .TextMatrix(mR, 2)
                            mTmpUserBatchID = .TextMatrix(mR, 13)
                            For mRow = 1 To .Rows - 2
                                If .TextMatrix(mRow, 0) = strChecked And mTmpPaymentEntryNo = .TextMatrix(mRow, 1) And mTmpInvoiceEntryNo = .TextMatrix(mRow, 2) And mTmpUserBatchID = .TextMatrix(mRow, 13) And LCase(.TextMatrix(mRow, 8)) = LCase("E") Then
                                    mTmpLineNo = mTmpLineNo & "," & Val(.TextMatrix(mRow, 12))
                                End If
                            Next mRow
                            If mTmpLineNo <> "" Then
                                mTmpLineNo = Right(mTmpLineNo, Len(mTmpLineNo) - 1)
                                ''tmp = " SELECT * FROM Txn_InvoiceAllocationForNAVISION Where PaymentEntryNo = '" & mTmpPaymentEntryNo & "' And InvoiceEntryNo = '" & mTmpInvoiceEntryNo & "' And UserBatchID = '" & mTmpUserBatchID & "' And DataTransfer = 'E' And [LineNo] In (" & mTmpLineNo & ") "
                                tmp = " UPDATE Txn_InvoiceAllocationForNAVISION SET DataTransfer = Null,Remark = '' Where PaymentEntryNo = '" & mTmpPaymentEntryNo & "' And InvoiceEntryNo = '" & mTmpInvoiceEntryNo & "' And UserBatchID = '" & mTmpUserBatchID & "' And DataTransfer = 'E' And [LineNo] In (" & mTmpLineNo & ") "
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
                        If .TextMatrix(mR, 0) = strChecked Then
                            mTmpPaymentEntryNo = .TextMatrix(mR, 1)
                            mTmpInvoiceEntryNo = .TextMatrix(mR, 2)
                            mTmpUserBatchID = .TextMatrix(mR, 13)
                            mTmpBatchNo = .TextMatrix(mR, 4)
                            mTmpInvoiceType = .TextMatrix(mR, 7)
                            
                            If LCase(mTmpInvoiceType) = LCase("CR") Then
                            ''--- CR Invoice Allocation
                                ''tmp = " SELECT * FROM Txn_NewCashReceipt Where ISNULL(InvAllocFlag,0) = 1 ANd BatchID = " & Val(mTmpBatchNo) & ""
                                tmp = " Update Txn_NewCashReceipt set InvAllocFlag = Null Where ISNULL(InvAllocFlag,0) = 1 ANd BatchID = " & Val(mTmpBatchNo) & ""
                                Call TmpTable
                            ElseIf LCase(mTmpInvoiceType) = LCase("CRA") Then
                            ''--- CR Adv Inv Allocation
                                ''tmp = " SELECT * FROM Txn_CRAdvanceAllocation Where ISNULL(BatchID,0) <> 0 And BatchID = " & Val(mTmpBatchNo) & " And CRAAID =1 "
                                ''tmp = " Update Txn_NewCashReceipt set BatchID = 0 Where ISNULL(BatchID,0) <> 0 And ISNULL(RBatchID,0) <> 0 And ReturnFlag='Y' And RBatchID = " & Val(mTmpBatchNo) & " And CRID = " & Val(Replace(mTmpPINo, "RCR", "")) & " "
                                ''Call TmpTable
                            ElseIf LCase(mTmpInvoiceType) = LCase("CRADJ") Then
                            ''--- Export TDS Adj Invoice
                                ''Call TmpTable
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
        .Cols = 14
        .TextMatrix(0, 0) = "Select":                   .ColWidth(0) = 800
        .TextMatrix(0, 1) = "Payment Entry No":         .ColWidth(1) = 1200
        .TextMatrix(0, 2) = "Invoice Entry No":         .ColWidth(2) = 1500
        .TextMatrix(0, 3) = "Amount":                   .ColWidth(3) = 1500
        .TextMatrix(0, 4) = "Batch ID":                 .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "JourNal Batch Type":       .ColWidth(5) = 1800
        .TextMatrix(0, 6) = "JourNal Batch Name":       .ColWidth(6) = 1800
        .TextMatrix(0, 7) = "Cash Receipt Type":        .ColWidth(7) = 1800
        .TextMatrix(0, 8) = "Navision Push Flag":       .ColWidth(8) = 1000
        .TextMatrix(0, 9) = "Navision Pull Flag":       .ColWidth(9) = 1000
        .TextMatrix(0, 10) = "Remark":                  .ColWidth(10) = 2500
        .TextMatrix(0, 11) = "Created By":              .ColWidth(11) = 2500
        .TextMatrix(0, 12) = "Line No":                 .ColWidth(12) = 1500
        .TextMatrix(0, 13) = "User Batch ID":           .ColWidth(13) = 2500
    Else
        .Cols = 13
        .TextMatrix(0, 0) = "Payment Entry No":         .ColWidth(0) = 1200
        .TextMatrix(0, 1) = "Invoice Entry No":         .ColWidth(1) = 1500
        .TextMatrix(0, 2) = "Amount":                   .ColWidth(2) = 1500
        .TextMatrix(0, 3) = "Batch ID":                 .ColWidth(3) = 1500
        .TextMatrix(0, 4) = "JourNal Batch Type":       .ColWidth(4) = 1800
        .TextMatrix(0, 5) = "JourNal Batch Name":       .ColWidth(5) = 1800
        .TextMatrix(0, 6) = "Cash Receipt Type":        .ColWidth(6) = 1800
        .TextMatrix(0, 7) = "Navision Push Flag":       .ColWidth(7) = 1000
        .TextMatrix(0, 8) = "Navision Pull Flag":       .ColWidth(8) = 1000
        .TextMatrix(0, 9) = "Remark":                   .ColWidth(9) = 2500
        .TextMatrix(0, 10) = "Created By":              .ColWidth(10) = 2500
        .TextMatrix(0, 11) = "Line No":                 .ColWidth(11) = 1500
        .TextMatrix(0, 12) = "User Batch ID":           .ColWidth(12) = 2500
    End If
    .RowHeight(0) = 800
End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CashReceiptAllocationForNavision.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CashReceiptAllocationForNavision.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

With MSHFlexGrid1
ProgressBar1.Max = MSHFlexGrid1.Rows - 1
    For I = 0 To .Rows - 1
        mRow = I + 1 ' + 6
        For c = 0 To .Cols - 1
            If InStr(.TextMatrix(I, c), "/") > 0 Then
               If IsDate(.TextMatrix(I, c)) Then
                  oWS.Cells(mRow, c + 1) = Format(.TextMatrix(I, c), "MM/dd/yyyy")
               Else
                  oWS.Cells(mRow, c + 1) = .TextMatrix(I, c)
               End If
            Else
               oWS.Cells(mRow, c + 1) = .TextMatrix(I, c)
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

Private Sub MSHFlexGrid1_Click()
Dim mR, mTempPaymentEntryNo, mTempInvoiceEntryNo As Variant
If mEditRights = True Then
    With MSHFlexGrid1
         If .Col = 0 Then
            If .TextMatrix(.row, .Col) = "" Then Exit Sub
            .Col = 0
            mTempPaymentEntryNo = Trim(.TextMatrix(.row, 1))
            mTempInvoiceEntryNo = Trim(.TextMatrix(.row, 2))
            If .TextMatrix(.row, .Col) = strUnChecked Then
               If Trim(.TextMatrix(.row, 1)) <> "" And Trim(.TextMatrix(.row, 2)) <> "" And Trim(.TextMatrix(.row, 3)) <> "" And LCase(.TextMatrix(.row, 8)) = LCase("E") Then
                  .TextMatrix(.row, .Col) = strChecked
                  For mR = 1 To .Rows - 2
                    If mTempPaymentEntryNo = Trim(.TextMatrix(mR, 1)) And mTempInvoiceEntryNo = Trim(.TextMatrix(mR, 2)) And LCase(.TextMatrix(mR, 8)) = LCase("E") Then
                        .TextMatrix(mR, .Col) = strChecked
                    End If
                  Next mR
               End If
            Else
                .TextMatrix(.row, .Col) = strUnChecked
                For mR = 1 To .Rows - 2
                  If mTempPaymentEntryNo = Trim(.TextMatrix(mR, 1)) And mTempInvoiceEntryNo = Trim(.TextMatrix(mR, 2)) Then
                      .TextMatrix(mR, .Col) = strUnChecked
                  End If
                Next mR
            End If
         End If
         
    mUpdateFlag = False
    For mR = 1 To .Rows - 2
        If .TextMatrix(mR, 0) = strChecked And LCase(.TextMatrix(mR, 8)) = LCase("E") Then
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
