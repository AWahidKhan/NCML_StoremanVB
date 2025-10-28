VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportCRTDSAdj 
   Caption         =   "Export CR TDS Adjustment"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8385
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   8040
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14182
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbType 
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
         ItemData        =   "Frm_ExportCRTDSAdj.frx":0000
         Left            =   5400
         List            =   "Frm_ExportCRTDSAdj.frx":000A
         TabIndex        =   15
         Top             =   240
         Width           =   3000
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
         Left            =   5400
         TabIndex        =   3
         Top             =   720
         Width           =   3600
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9120
         TabIndex        =   5
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   12540
         TabIndex        =   10
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   11400
         TabIndex        =   9
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10260
         TabIndex        =   8
         Top             =   240
         Width           =   1140
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
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
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
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
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
         Left            =   2355
         TabIndex        =   2
         Top             =   240
         Width           =   1560
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
         Height          =   350
         Left            =   13680
         TabIndex        =   7
         Top             =   240
         Width           =   1140
      End
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Type"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   270
         Width           =   1095
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
         Left            =   4080
         TabIndex        =   14
         Top             =   780
         Width           =   1575
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
         Left            =   1350
         TabIndex        =   11
         Top             =   270
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_ExportCRTDSAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCodeTDS, mGLAccNAVCodeTDSDP, mGLAccNAVCodeOther As Variant
Dim ExpFlag, MaxID As Variant
Dim mSDepoBankID As Double
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mBatchID As Double
Dim mJournalBatchName, mJournalType As Variant

Private Sub ChkCheckAll_Click()

If MSHFlexGrid3.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid3.Rows - 2
       MSHFlexGrid3.row = I
       
       If MSHFlexGrid3.TextMatrix(I, 3) <> "" Then 'And MSHFlexGrid3.TextMatrix(I, 6) <> "" And MSHFlexGrid3.TextMatrix(I, 13) <> "" Then
          If mPINo <> MSHFlexGrid3.TextMatrix(I, 13) Then
             MSHFlexGrid3.row = I
             If LCase(CmbType.Text) <> LCase("PartialReversal") Then
                MSHFlexGrid3.TextMatrix(I, 19) = strChecked
             Else
                MSHFlexGrid3.TextMatrix(I, 22) = strChecked
             End If
             MSHFlexGrid3.TextMatrix(0, 4) = Format(Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I, 4)), "#####0.00")
          End If
       Else
          mPINo = MSHFlexGrid3.TextMatrix(I, 13)
       End If
   Next
Else
   MSHFlexGrid3.TextMatrix(0, 4) = ""
   For I = 2 To MSHFlexGrid3.Rows - 2
       MSHFlexGrid3.row = I
       If LCase(CmbType.Text) <> LCase("PartialReversal") Then
          MSHFlexGrid3.TextMatrix(I, 19) = strUnChecked
       Else
          MSHFlexGrid3.TextMatrix(I, 22) = strUnChecked
       End If
   Next
End If

End Sub

Private Sub CmbBatch_GotFocus()
tmp = "select distinct BatchID As BatchID from Txn_CRTDSAdjustment Where BatchID > 0  order by BatchID"
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
End Sub

Private Sub CmbBatch_LostFocus()
If CmbBatch.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "BatchID = " & Val(CmbBatch.Text) & ""
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbBatch.SetFocus
   Exit Sub
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

Private Sub CmbType_LostFocus()
If CmbType.Text = "" Then Exit Sub
   
If LCase(CmbType.Text) <> LCase("TDS/Otherdeduction") And LCase(CmbType.Text) <> LCase("PartialReversal") Then
   MsgBox "Invalid Selection!!!"
   CmbType.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportCashReceipt.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportCashReceipt.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    If Option1.Value = True Then
       For C = 0 To MSHFlexGrid3.Cols - 2
           If InStr(MSHFlexGrid3.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid3.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid3.TextMatrix(I, C), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
              End If
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
           End If
       Next
    Else
       For C = 0 To MSHFlexGrid3.Cols - 1
           If InStr(MSHFlexGrid3.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid3.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid3.TextMatrix(I, C), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
              End If
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
           End If
       Next
    End If
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdExport_Click()
Dim Path, strFilName, strFilLocation As Variant
Dim strSplitString() As String
Dim CRID_, Prefix_ As String

''-- On 05 Aug 2016
Dim mUserBatchID, mInvoiceType As Variant
'mJournalType = "REPEIPTS"
'mUserBatchID = GetGUID
'mJournalBatchName = GetJournalDetailsFromHRPro(Gen_UserLoginID, mJournalType)
'If Trim(mJournalBatchName) = "" Then
'   MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'   Exit Sub
'End If
''-- On 05 Aug 2016

If CmbType.Text = "" Then
   MsgBox "Select Type!!!"
   CmbType.SetFocus
   Exit Sub
End If
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
mInvoiceType = "CRADJ"

If LCase(CmbType.Text) = LCase("TDS/Otherdeduction") Then
   If Val(MSHFlexGrid3.TextMatrix(0, 4)) > 0 Then
      MsgBox "Total of Col 'Amount' must be zero!!!"
      Exit Sub
   End If
End If

Path = App.Path
MaxID = GetMaxBatchID

If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_CRTDSAdjustment" & "_" & MaxID & ".csv"
Else
   strFilName = "\Navision\NavUpload_CRTDSAdjustment" & "_" & CmbBatch.Text & ".csv"
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
    wc = ""
Else
   Call TableTxn_CashRepeiptsForNAVISION(" Where G_UID Is Null ")
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
    wc = ""
End If

'If Option1.Value = True Then
'   If CmbSDepositedBank.Text <> "" Then
''      tmp = "Select * Txn_NewCashReceipt Where BatchID=0 And CRDepositID In ( select CRDepositID From Txn_NewCashReceiptDepositDetails Where FAGBankID = " & mSDepoBankID & ")"
'      Wc = " And CRDepositID In ( select CRDepositID From Txn_NewCashReceiptDepositDetails Where FAGBankID = " & mSDepoBankID & ") "
'   Else
''      tmp = "Select * Txn_NewCashReceipt Where BatchID=0 And CRDepositID > 0"
'      Wc = " And CRDepositID > 0 "
'   End If
'
'
''   Call Tmp1Table
'End If

With MSHFlexGrid3
For I = 2 To .Rows - 1
    tmp = ""
    If Option1.Value = True Then
       If LCase(CmbType.Text) = LCase("TDS/Otherdeduction") Then
           If .TextMatrix(I, 19) = strChecked Then
    '          For c = 0 To .Cols - 2
    '              If tmp = "" Then
    '                 tmp = tmp & .TextMatrix(I, c)
    '              Else
    '                 tmp = tmp & "," & .TextMatrix(I, c)
    '              End If
    '          Next
    '          Print #1, tmp
    
                RsTxn_CashRepeiptsForNAVISION.AddNew
                RsTxn_CashRepeiptsForNAVISION!PostingDate = Trim(.TextMatrix(I, 0))
                RsTxn_CashRepeiptsForNAVISION!DocumentType = Trim(.TextMatrix(I, 1))
                RsTxn_CashRepeiptsForNAVISION!AccountType = Trim(.TextMatrix(I, 2))
                RsTxn_CashRepeiptsForNAVISION!AccountNo = Trim(.TextMatrix(I, 3))
                RsTxn_CashRepeiptsForNAVISION!Amount = Trim(.TextMatrix(I, 4))
                RsTxn_CashRepeiptsForNAVISION!InvoiceNo = Trim(.TextMatrix(I, 5))
                RsTxn_CashRepeiptsForNAVISION!Narration = Trim(.TextMatrix(I, 6))
                RsTxn_CashRepeiptsForNAVISION!RegionDim = Trim(.TextMatrix(I, 7))
                RsTxn_CashRepeiptsForNAVISION!DeptDim = Trim(.TextMatrix(I, 8))
                RsTxn_CashRepeiptsForNAVISION!EmpDim = Trim(.TextMatrix(I, 9))
                RsTxn_CashRepeiptsForNAVISION!ChequeNo = Trim(.TextMatrix(I, 10))
                RsTxn_CashRepeiptsForNAVISION!ChequeDate = IIf(Trim(.TextMatrix(I, 11)) = "", Null, Trim(.TextMatrix(I, 11)))
                RsTxn_CashRepeiptsForNAVISION!DepositDate = IIf(Trim(.TextMatrix(I, 12)) = "", Null, Trim(.TextMatrix(I, 12)))
                RsTxn_CashRepeiptsForNAVISION!PINo = Trim(.TextMatrix(I, 13))
                RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo = IIf(Trim(.TextMatrix(I, 14)) = "", Null, Trim(.TextMatrix(I, 14)))
                RsTxn_CashRepeiptsForNAVISION!EntryNo = IIf(Trim(.TextMatrix(I, 15)) = "", Null, Trim(.TextMatrix(I, 15)))
                'RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo = Null
                RsTxn_CashRepeiptsForNAVISION!CreatedBy = Gen_UserLoginID
                RsTxn_CashRepeiptsForNAVISION!G_UID = GetGUID
                RsTxn_CashRepeiptsForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
                'RsTxn_CashRepeiptsForNAVISION!DataTransfer = Null
                RsTxn_CashRepeiptsForNAVISION!InvoiceType = mInvoiceType
                RsTxn_CashRepeiptsForNAVISION!JournalBatchType = mJournalType
                RsTxn_CashRepeiptsForNAVISION!JournalBatchName = mJournalBatchName
                'RsTxn_CashRepeiptsForNAVISION!ReverseInvoice = ""
                RsTxn_CashRepeiptsForNAVISION!STCode = IIf(Trim(.TextMatrix(I, 16)) = "", Null, Trim(.TextMatrix(I, 16)))
                RsTxn_CashRepeiptsForNAVISION!BalAccType = IIf(Trim(.TextMatrix(I, 17)) = "", Null, Trim(.TextMatrix(I, 17)))
                RsTxn_CashRepeiptsForNAVISION!BalAccountNo = IIf(Trim(.TextMatrix(I, 18)) = "", Null, Trim(.TextMatrix(I, 18)))
                'RsTxn_CashRepeiptsForNAVISION!EntryNoTransfer = Null
                'RsTxn_CashRepeiptsForNAVISION!EntryNoTransferRemark = Null
                RsTxn_CashRepeiptsForNAVISION!UserBatchID = mUserBatchID
                RsTxn_CashRepeiptsForNAVISION.Update
    
                If Option1.Value = True Then
                  If Val(.TextMatrix(I, 4)) < 0 Then
                     strSplitString = Split(.TextMatrix(I, 13), "-")
                     CRID_ = strSplitString(0)
                     Prefix_ = strSplitString(1)
                     tmp = " Update  Txn_CRTDSAdjustment Set BatchID = " & MaxID & " Where CRTDSAdjustmentID = " & Replace(CRID_, "TDS", "") & " And SM_Prefix = '" & Prefix_ & "' And BatchID = 0  " '& Wc 'And CRDepositID In ( select CRDepositID From Txn_NewCashReceiptDepositDetails Where FAGBankID = " & mSDepoBankID & ")"
                     Call Tmp1Table
                  End If
                End If
              
           End If
       Else
           If .TextMatrix(I, 22) = strChecked Then
'               For C = 0 To .Cols - 2
'                  If tmp = "" Then
'                     tmp = tmp & .TextMatrix(I, C)
'                  Else
'                     tmp = tmp & "," & .TextMatrix(I, C)
'                  End If
'               Next
'               Print #1, tmp
           
                RsTxn_CashRepeiptsForNAVISION.AddNew
                RsTxn_CashRepeiptsForNAVISION!PostingDate = Trim(.TextMatrix(I, 0))
                RsTxn_CashRepeiptsForNAVISION!DocumentType = Trim(.TextMatrix(I, 1))
                RsTxn_CashRepeiptsForNAVISION!AccountType = Trim(.TextMatrix(I, 2))
                RsTxn_CashRepeiptsForNAVISION!AccountNo = Trim(.TextMatrix(I, 3))
                RsTxn_CashRepeiptsForNAVISION!Amount = Trim(.TextMatrix(I, 4))
                
                RsTxn_CashRepeiptsForNAVISION!BalAccType = IIf(Trim(.TextMatrix(I, 17)) = "", Null, Trim(.TextMatrix(I, 17)))
                RsTxn_CashRepeiptsForNAVISION!BalAccountNo = IIf(Trim(.TextMatrix(I, 18)) = "", Null, Trim(.TextMatrix(I, 18)))
                
                RsTxn_CashRepeiptsForNAVISION!Narration = Trim(.TextMatrix(I, 6))
                RsTxn_CashRepeiptsForNAVISION!RegionDim = Trim(.TextMatrix(I, 7))
                RsTxn_CashRepeiptsForNAVISION!DeptDim = Trim(.TextMatrix(I, 8))
                
                RsTxn_CashRepeiptsForNAVISION!PINo = Trim(.TextMatrix(I, 13))
                RsTxn_CashRepeiptsForNAVISION!EntryNo = Trim(.TextMatrix(I, 16))
                RsTxn_CashRepeiptsForNAVISION!STCode = Trim(.TextMatrix(I, 15))
                 ''TmpRs4!AgreeNonAgree = Trim(MSHFlexGrid3.TextMatrix(I, 20))
                
                RsTxn_CashRepeiptsForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text) ''MaxID  '' CmbBatch.Text
                RsTxn_CashRepeiptsForNAVISION!InvoiceType = Trim(CmbType.Text)
                RsTxn_CashRepeiptsForNAVISION!JournalBatchType = Trim(mJournalType)
                RsTxn_CashRepeiptsForNAVISION!JournalBatchName = Trim(mJournalBatchName)
                RsTxn_CashRepeiptsForNAVISION!CreatedBy = Gen_UserLoginID
                RsTxn_CashRepeiptsForNAVISION!G_UID = GetGUID
                RsTxn_CashRepeiptsForNAVISION!ReverseInvoice = "Y"   ''-- If Invoice Is ReverseType
                RsTxn_CashRepeiptsForNAVISION!UserBatchID = mUserBatchID
                  
                RsTxn_CashRepeiptsForNAVISION!GSTGroupCode = Trim(.TextMatrix(I, 19))
                RsTxn_CashRepeiptsForNAVISION!HSNNo = Trim(.TextMatrix(I, 20))
                RsTxn_CashRepeiptsForNAVISION!AccType = Trim(.TextMatrix(I, 21))
                 
                RsTxn_CashRepeiptsForNAVISION.Update
                RsTxn_CashRepeiptsForNAVISION.Requery
                
                If Option1.Value = True Then
                  If Val(.TextMatrix(I, 4)) > 0 Then
                     strSplitString = Split(.TextMatrix(I, 13), "-")
                     CRID_ = strSplitString(0)
                     Prefix_ = strSplitString(1)
                     tmp = " Update  Txn_CRTDSAdjustment Set BatchID = " & MaxID & " Where CRTDSAdjustmentID = " & Replace(CRID_, "TDS", "") & " And SM_Prefix = '" & Prefix_ & "' And BatchID = 0  " '& Wc 'And CRDepositID In ( select CRDepositID From Txn_NewCashReceiptDepositDetails Where FAGBankID = " & mSDepoBankID & ")"
                     Call Tmp1Table
                  End If
                End If
           End If
       End If
    Else
       For C = 0 To .Cols - 1
           If tmp = "" Then
              tmp = tmp & .TextMatrix(I, C)
           Else
              tmp = tmp & "," & .TextMatrix(I, C)
           End If
       Next
       Print #1, tmp
    End If
'    For c = 0 To MSHFlexGrid3.Cols - 1
'        If tmp = "" Then
'           tmp = tmp & MSHFlexGrid3.TextMatrix(I, c)
'        Else
'           tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, c)
'        End If
'    Next
'    Print #1, tmp
Next
End With
'If Option1.Value = True Then
'   If CmbSDepositedBank.Text <> "" Then
'      tmp = "Update Txn_NewCashReceipt set BatchID = " & MaxID & " Where BatchID=0 And CRDepositID In ( select CRDepositID From Txn_NewCashReceiptDepositDetails Where FAGBankID = " & mSDepoBankID & ")"
'   Else
'      tmp = "Update Txn_NewCashReceipt set BatchID = " & MaxID & " Where BatchID=0 And CRDepositID > 0"
'   End If
'   Call TmpTable
'End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Close #1
Call CmdReset_Click
End Sub

Private Function GetMaxBatchID()
tmp = "select Max(BatchID) from Txn_CRTDSAdjustment"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   GetMaxBatchID = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0)) + 1
Else
   GetMaxBatchID = 1
End If
End Function

Private Sub CmdGetReport_Click()
Dim mSum1, mSum2 As Variant
Dim wc As Variant

Dim mNAVCode As String
ExpFlag = ""
If Option1.Value = True Then
   mBatchID = 0
   If CmbType.Text = "" Then
      MsgBox "Select Type!!!"
      CmbType.SetFocus
      Exit Sub
   End If
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
   mBatchID = Val(CmbBatch.Text)
End If

Call Grid_Head

If LCase(CmbType.Text) = LCase("TDS/Otherdeduction") Or Option2.Value = True Then

    tmp = " Select PostingDate,DocumentType,AccountType,AccountNo,Amount,InvoiceNo,Narration,RegionDim,DeptDim,ClientNAVCode,DPNAVCode,PINo,PostInvoiceNo,EntryNo,RNAVCode,ServiceTaxID,Sequense"
    tmp = tmp & " From"
    tmp = tmp & " (Select FAGPostingDate as PostingDate,'Payment'as DocumentType,'Customer' as AccountType,'' As AccountNo ,-sum(TCRT.RecTDS) as Amount,'' as InvoiceNo,'Being receipt for services' as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,convert(varchar,TCR.CRTDSAdjustmentID) + '-'+ TCR.SM_Prefix as PINo,'' PostInvoiceNo,Null as EntryNo,'' As 'RNAVCode',0 as ServiceTaxID,1 as Sequense"
    tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
    tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
    tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
    tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
    tmp = tmp & " Where TCRT.RecTDS > 0 And TCR.BatchID = " & mBatchID & "   And TCR.Flag='A'"
    tmp = tmp & " Group By FAGPostingDate,ML.NAVCode,TI.NAVCode,TCR.CRTDSAdjustmentID,TCR.SM_Prefix"
    tmp = tmp & " Union"
    tmp = tmp & " Select FAGPostingDate as PostingDate,'Payment'as DocumentType,'G/L Account' as AccountType,'' As AccountNo ,sum(TCRT.RecTDS) as Amount,'' as InvoiceNo,'Being TDS deducted  by Customer ' as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,convert(varchar,TCR.CRTDSAdjustmentID) + '-'+ TCR.SM_Prefix as PINo,'' as PostInvoiceNo,Null as EntryNo,'' as RNAVCode,0 as ServiceTaxID,2 as Sequense"
    tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
    tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
    tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
    tmp = tmp & " Where TCRT.RecTDS > 0 And TCR.BatchID = " & mBatchID & "  And TCR.Flag='A'"
    tmp = tmp & " Group By FAGPostingDate,ML.NAVCode,TI.NAVCode,TCR.CRTDSAdjustmentID,TCR.SM_Prefix"
    
    tmp = tmp & " Union"
    
    tmp = tmp & " Select FAGPostingDate as PostingDate,'Credit Memo'as DocumentType,'Customer' as AccountType,ME.NAVCode As AccountNo ,-sum(TCRT.RecOtherDed-TCRT.STAmount) as Amount,'' as InvoiceNo,'Being Other Deduction' as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,convert(varchar,TCR.CRTDSAdjustmentID) + '-'+ TCR.SM_Prefix as PINo,'' PostInvoiceNo,Null as EntryNo,'' As 'RNAVCode',isnull(ME.ServiceTaxID,0) as ServiceTaxID,3 as Sequense"
    tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
    tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
    tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
    tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
    tmp = tmp & " inner join Mst_ExpenseType ME on TCRT.ExpenseAccID=ME.ExpenseTypeID"
    tmp = tmp & " Where ME.ReverseFlag='N' And TCR.BatchID = " & mBatchID & "   And TCR.Flag='A' and TCRT.STAmount>0"
    tmp = tmp & " Group By FAGPostingDate,ML.NAVCode,TI.NAVCode,TCR.CRTDSAdjustmentID,TCR.SM_Prefix,ME.ServiceTaxID,ME.NAVCode"
    
    tmp = tmp & " Union"
    
    tmp = tmp & " Select FAGPostingDate as PostingDate,'Credit Memo'as DocumentType,'G/L Account' as AccountType,ME.NAVCode As AccountNo ,sum(TCRT.RecOtherDed-TCRT.STAmount) as Amount,'' as InvoiceNo,'Being Other Deduction ' as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,convert(varchar,TCR.CRTDSAdjustmentID) + '-'+ TCR.SM_Prefix as PINo,'' as PostInvoiceNo,Null as EntryNo,'' as RNAVCode,isnull(ME.ServiceTaxID,0) ServiceTaxID,4 as Sequense"
    tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
    tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
    tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
    tmp = tmp & " inner join Mst_ExpenseType ME on TCRT.ExpenseAccID=ME.ExpenseTypeID"
    tmp = tmp & " Where ME.ReverseFlag='N' And TCR.BatchID = " & mBatchID & "  And TCR.Flag='A'  and TCRT.STAmount>0"
    tmp = tmp & " Group By ME.NAVCode,FAGPostingDate,ML.NAVCode,TI.NAVCode,TCR.CRTDSAdjustmentID,TCR.SM_Prefix,ME.ServiceTaxID"
    
    tmp = tmp & " Union"
    
    tmp = tmp & " Select FAGPostingDate as PostingDate,'Credit Memo'as DocumentType,'Customer' as AccountType,ME.NAVCode As AccountNo ,-sum(TCRT.RecOtherDed-TCRT.STAmount) as Amount,'' as InvoiceNo,'Being Other Deduction' as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,convert(varchar,TCR.CRTDSAdjustmentID) + '-'+ TCR.SM_Prefix as PINo,'' PostInvoiceNo,Null as EntryNo,'' As 'RNAVCode',0 as ServiceTaxID,5 as Sequense"
    tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
    tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
    tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
    tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
    tmp = tmp & " inner join Mst_ExpenseType ME on TCRT.ExpenseAccID=ME.ExpenseTypeID"
    tmp = tmp & " Where ME.ReverseFlag='N' And TCR.BatchID = " & mBatchID & "   And TCR.Flag='A' and TCRT.STAmount=0"
    tmp = tmp & " Group By FAGPostingDate,ML.NAVCode,TI.NAVCode,TCR.CRTDSAdjustmentID,TCR.SM_Prefix,ME.ServiceTaxID,ME.NAVCode"
    
    tmp = tmp & " Union"
    
    tmp = tmp & " Select FAGPostingDate as PostingDate,'Credit Memo'as DocumentType,'G/L Account' as AccountType,ME.NAVCode As AccountNo ,sum(TCRT.RecOtherDed-TCRT.STAmount) as Amount,'' as InvoiceNo,'Being Other Deduction ' as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,convert(varchar,TCR.CRTDSAdjustmentID) + '-'+ TCR.SM_Prefix as PINo,'' as PostInvoiceNo,Null as EntryNo,'' as RNAVCode,0 ServiceTaxID,6 as Sequense"
    tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
    tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
    tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
    tmp = tmp & " inner join Mst_ExpenseType ME on TCRT.ExpenseAccID=ME.ExpenseTypeID"
    tmp = tmp & " Where ME.ReverseFlag='N' And TCR.BatchID = " & mBatchID & "  And TCR.Flag='A' and TCRT.STAmount=0"
    tmp = tmp & " Group By ME.NAVCode,FAGPostingDate,ML.NAVCode,TI.NAVCode,TCR.CRTDSAdjustmentID,TCR.SM_Prefix,ME.ServiceTaxID"
    
    
    tmp = tmp & " ) A Order By PINO,AccountNo,sequense"
    
    Call TmpTable

    mSum1 = 0
    mSum2 = 0
    
    mGLAccNAVCodeTDS = GetGLAccNAVCode("TDSNavCode")
    mGLAccNAVCodeTDSDP = GetGLAccNAVCode("TDSNavCodeDP")
    
    
    If TmpRs.RecordCount > 0 Then
       For r = 1 To TmpRs.RecordCount
           I = MSHFlexGrid3.Rows - 1
           MSHFlexGrid3.TextMatrix(I, 0) = Format(TmpRs!PostingDate, "mm/dd/yyyy") 'Format(TmpRs!PostingDate, "dd/mm/yyyy")
           MSHFlexGrid3.TextMatrix(I, 1) = TmpRs!DocumentType
           MSHFlexGrid3.TextMatrix(I, 2) = TmpRs!AccountType
           MSHFlexGrid3.TextMatrix(I, 6) = TmpRs!Narration
           If TmpRs!AccountType = "Customer" Then
              If IsNull(TmpRs!ClientNAVCode) = False Then
                 MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
                 mNAVCode = MSHFlexGrid3.TextMatrix(I, 3)
              End If
              If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
           ElseIf TmpRs!AccountType = "G/L Account" Then
              If IsNull(TmpRs!ClientNAVCode) = False Then
                 mNAVCode = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
              ElseIf IsNull(TmpRs!RNAVCode) = False Then
                 mNAVCode = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)
              Else
                 mNAVCode = IIf(IsNull(TmpRs!DPNAVCode), "", TmpRs!DPNAVCode)
              End If
              If InStr(1, TmpRs!Narration, "TDS") > 0 Then
                 MSHFlexGrid3.TextMatrix(I, 6) = TmpRs!Narration & " " & mNAVCode
                 If InStr(1, mNAVCode, "TDS") And TmpRs!AccountNo = "" Then
                    MSHFlexGrid3.TextMatrix(I, 3) = mGLAccNAVCodeTDSDP
                 ElseIf TmpRs!AccountNo = "" Then
                    MSHFlexGrid3.TextMatrix(I, 3) = mGLAccNAVCodeTDS
                 End If
              Else
                 MSHFlexGrid3.TextMatrix(I, 3) = TmpRs!AccountNo
              End If
              If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
           End If
           If TmpRs!InvoiceNo <> "Credit Memo" Then
              MSHFlexGrid3.TextMatrix(I, 4) = TmpRs!Amount
           End If
           MSHFlexGrid3.TextMatrix(I, 7) = TmpRs!RegionDim
           MSHFlexGrid3.TextMatrix(I, 8) = TmpRs!DeptDim
           MSHFlexGrid3.TextMatrix(I, 9) = "EMP-1000"
           
           MSHFlexGrid3.TextMatrix(I, 13) = "TDS" & TmpRs!PINo
           MSHFlexGrid3.TextMatrix(I, 14) = IIf(IsNull(TmpRs!PostInvoiceNo), "", TmpRs!PostInvoiceNo)
           MSHFlexGrid3.TextMatrix(I, 15) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
           
           If TmpRs!ServiceTaxID > 0 Then
              MSHFlexGrid3.TextMatrix(I, 16) = GetServiceTaxNAVCode(TmpRs!ServiceTaxID)
           End If
    
           If Option1.Value = True Then
              
              MSHFlexGrid3.row = I
              MSHFlexGrid3.Col = 19
              MSHFlexGrid3.CellFontName = "Wingdings"
              MSHFlexGrid3.CellFontSize = 12
              MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
              MSHFlexGrid3.Text = strUnChecked
           End If
           MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
           TmpRs.MoveNext
       Next
    End If
End If
If LCase(CmbType.Text) = LCase("PartialReversal") Or Option2.Value = True Then
    
    tmp = "Select FAGPostingDate as PostingDate,'Credit Memo' as DocumentType,'Customer' as AccountType,TI.NAVCode As AccountNo, (TCRT.RecOtherDed-TCRT.STAmount) as Amount ,'' as InvoiceNo,TCR.FAGRemarks as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,convert(varchar,TCR.CRTDSAdjustmentID) + '-'+ TCR.SM_Prefix + '-'+ convert(varchar,TCRT.TIID) as PINo,TI.EntryNo as EntryNo,isnull(ME.ServiceTaxID,0) ServiceTaxID,'G/L Account' as BalAccType,ME.NAVCode as BalAcc,MST.NAVCOde 'STNAVCode',MST.InvoiceType,ME.HSNNo "
    tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
    tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
    tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
    tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
    tmp = tmp & " Left join Mst_ExpenseType ME on TCRT.ExpenseAccID=ME.ExpenseTypeID"
    tmp = tmp & " Left join Mst_ServiceTax MST on ME.ServiceTaxID=MST.ServiceTaxID "
    tmp = tmp & " Where TCRT.RecOtherDed > 0  And TCR.BatchID =  " & mBatchID & "   And TCR.Flag='A' And ME.ReverseFlag='Y'"
    
    Call Tmp2Table

    If TmpRs2.RecordCount > 0 Then
       For r = 1 To TmpRs2.RecordCount
           I = MSHFlexGrid3.Rows - 1
           MSHFlexGrid3.TextMatrix(I, 0) = Format(TmpRs2!PostingDate, "mm/dd/yyyy") 'Format(TmpRs2!PostingDate, "dd/mm/yyyy")
           MSHFlexGrid3.TextMatrix(I, 1) = TmpRs2!DocumentType
           MSHFlexGrid3.TextMatrix(I, 2) = TmpRs2!AccountType
           MSHFlexGrid3.TextMatrix(I, 6) = TmpRs2!Narration
           MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs2!AccountNo), "", TmpRs2!AccountNo)
           mNAVCode = MSHFlexGrid3.TextMatrix(I, 3)
           If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
           MSHFlexGrid3.TextMatrix(I, 4) = TmpRs2!Amount
           MSHFlexGrid3.TextMatrix(I, 7) = TmpRs2!RegionDim
           MSHFlexGrid3.TextMatrix(I, 8) = TmpRs2!DeptDim
           MSHFlexGrid3.TextMatrix(I, 13) = "TDS" & TmpRs2!PINo
           MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs2!EntryNo), "", TmpRs2!EntryNo)
           MSHFlexGrid3.TextMatrix(I, 17) = TmpRs2!BalAccType
           MSHFlexGrid3.TextMatrix(I, 18) = TmpRs2!BalAcc
           MSHFlexGrid3.TextMatrix(I, 19) = IIf(IsNull(TmpRs2!STNAVCode), "", TmpRs2!STNAVCode)
           MSHFlexGrid3.TextMatrix(I, 20) = IIf(IsNull(TmpRs2!HSNNo), "", TmpRs2!HSNNo)
           MSHFlexGrid3.TextMatrix(I, 21) = IIf(IsNull(TmpRs2!InvoiceType), "", TmpRs2!InvoiceType)
           If MSHFlexGrid3.TextMatrix(I, 20) = "" Then ExpFlag = "N"
           If MSHFlexGrid3.TextMatrix(I, 19) = "" Then ExpFlag = "N"
           If Option1.Value = True Then
              MSHFlexGrid3.row = I
              MSHFlexGrid3.Col = 22
              MSHFlexGrid3.CellFontName = "Wingdings"
              MSHFlexGrid3.CellFontSize = 12
              MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
              MSHFlexGrid3.Text = strUnChecked
           End If
           MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
           TmpRs2.MoveNext
       Next
    End If
End If

End Sub
Private Sub AddToCustAcc(mNAVCode_ As Variant, mPINo_ As Variant, mAmt_ As Variant)
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 3) = mNAVCode_ And MSHFlexGrid3.TextMatrix(I, 13) = mPINo_ And LCase(MSHFlexGrid3.TextMatrix(I, 2)) = "customer" Then
       MSHFlexGrid3.TextMatrix(I, 4) = Val(MSHFlexGrid3.TextMatrix(I, 4)) - Val(mAmt_)
       Exit Sub
    End If
Next
End Sub
Private Sub CmdReset_Click()
Call Grid_Head
CmbBatch.Text = ""
CmdExport.Enabled = True
Option1.Value = True
ChkCheckAll.Value = 0
CmbType.Text = ""
Call Option1_Click
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame5.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame5.Enabled = True
Call Grid_Head
End Sub
Public Sub Grid_Head()
If LCase(CmbType.Text) <> LCase("PartialReversal") Then

    With MSHFlexGrid3
        .Clear
        .Rows = 3
        .FixedRows = 2
        If Option1.Value = True Then
           .Cols = 20
         Else
           .Cols = 19
        End If
        .AllowUserResizing = flexResizeColumns
        .TextMatrix(1, 0) = "Posting Date"
        .TextMatrix(1, 1) = "Document Type"
        .TextMatrix(1, 2) = "Account Type"
        .TextMatrix(1, 3) = "Account No."
        .TextMatrix(1, 4) = "Amount"
        .TextMatrix(1, 5) = "Invoice No"
        .TextMatrix(1, 6) = "Narration"
        .TextMatrix(1, 7) = "RegionDim"
        .TextMatrix(1, 8) = "DeptDim"
        .TextMatrix(1, 9) = "EmpDim"
        .TextMatrix(1, 10) = "Cheque No."
        .TextMatrix(1, 11) = "Cheque Date"
        .TextMatrix(1, 12) = "Deposit Date"
        .TextMatrix(1, 13) = "PINo"
        .TextMatrix(1, 14) = "PostInvoiceNo"
        .TextMatrix(1, 15) = "EntryNo"
        .TextMatrix(1, 16) = "STCode"
        .TextMatrix(1, 17) = "Bal Acc Type"
        .TextMatrix(1, 18) = "Bal Acc No"
        
        If Option1.Value = True Then
           .TextMatrix(1, 19) = "Selection"
        End If
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
        .ColWidth(10) = 1400
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        
    End With
Else
    With MSHFlexGrid3
        .Clear
        .Rows = 3
        .FixedRows = 2
        If Option1.Value = True Then
           .Cols = 23
         Else
           .Cols = 22
        End If
        .AllowUserResizing = flexResizeColumns
        .TextMatrix(1, 0) = "Posting Date"
        .TextMatrix(1, 1) = "Document Type"
        .TextMatrix(1, 2) = "Account Type"
        .TextMatrix(1, 3) = "Account No."
        .TextMatrix(1, 4) = "Amount"
        .TextMatrix(1, 5) = "Invoice No"
        .TextMatrix(1, 6) = "Narration"
        .TextMatrix(1, 7) = "RegionDim"
        .TextMatrix(1, 8) = "DeptDim"
        .TextMatrix(1, 9) = "EmpDim"
        .TextMatrix(1, 10) = "Cheque No."
        .TextMatrix(1, 11) = "Cheque Date"
        .TextMatrix(1, 12) = "Deposit Date"
        .TextMatrix(1, 13) = "PINo"
        .TextMatrix(1, 14) = "PostInvoiceNo"
        
    '    .TextMatrix(1, 15) = "EntryNo"
    '    .TextMatrix(1, 16) = "STCode"
    '    .TextMatrix(1, 17) = "Bal Acc Type"
    '    .TextMatrix(1, 18) = "Bal Acc No"
     
        'Commented on 01 Nov 2017 for GST Modification
        '.TextMatrix(1, 15) = "EntryNo"
        '.TextMatrix(1, 16) = "STCode"
    
        .TextMatrix(1, 15) = "STCode"
        .TextMatrix(1, 16) = "EntryNo"
        .TextMatrix(1, 17) = "Bal Acc Type"
        .TextMatrix(1, 18) = "Bal Acc No"
        
        'Commented on 01 Nov 2017 for GST Modification New Column for
        .TextMatrix(1, 19) = "GST Group Code"
        .TextMatrix(1, 20) = "HSN No"
        .TextMatrix(1, 21) = "Invoice Type"
        
        If Option1.Value = True Then
           .TextMatrix(1, 22) = "Selection"
        End If
        
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
        .ColWidth(10) = 1400
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        
    End With
End If


End Sub

Private Sub MSHFlexGrid3_Click()
Dim mAmount As Variant
mAmount = 0

With MSHFlexGrid3
     If LCase(CmbType.Text) = LCase("TDS/Otherdeduction") Then
        If .Col = 19 Then
           If .TextMatrix(.row, .Col) = "" Then Exit Sub
           .Col = 19
           If .TextMatrix(.row, .Col) = strUnChecked Then
              If .TextMatrix(.row, 3) <> "" Then
                 If Val(.TextMatrix(.row, 4)) < 0 Then
                    mPINo = .TextMatrix(.row, 13)
                    .TextMatrix(.row, .Col) = strChecked
                   ' mAmount = Val(.TextMatrix(.Row, 4))
                    mAmount = 0
                    For x = 2 To MSHFlexGrid3.Rows - 2
                        If mPINo = .TextMatrix(x, 13) Then
                           .TextMatrix(x, .Col) = strChecked
                           mAmount = Val(mAmount) + Val(.TextMatrix(x, 4))
                        End If
                    Next
                    .TextMatrix(0, 4) = Format(Val(.TextMatrix(0, 4)) + Val(mAmount), "#####0.00")
                    mAmount = 0
                 End If
              End If
           Else
              mAmount = 0
              .TextMatrix(.row, .Col) = strUnChecked
          '    mAmount = Val(.TextMatrix(.Row, 4))
              mAmount = 0
              mPINo = .TextMatrix(.row, 13)
              For x = 2 To MSHFlexGrid3.Rows - 1
                  If mPINo = .TextMatrix(x, 13) Then
                     .TextMatrix(x, .Col) = strUnChecked
                      mAmount = Val(mAmount) + Val(.TextMatrix(x, 4))
                  End If
              Next
              .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) - Val(mAmount)
           End If
        End If
     Else
        If .Col = 22 Then
           If .TextMatrix(.row, .Col) = "" Then Exit Sub
           .Col = 22
           If .TextMatrix(.row, .Col) = strUnChecked Then
              If .TextMatrix(.row, 3) <> "" Then
                 If Val(.TextMatrix(.row, 4)) > 0 Then
                    mPINo = .TextMatrix(.row, 13)
                    .TextMatrix(.row, .Col) = strChecked
                   ' mAmount = Val(.TextMatrix(.Row, 4))
                    mAmount = 0
                    For x = 2 To MSHFlexGrid3.Rows - 2
                        If mPINo = .TextMatrix(x, 13) Then
                           .TextMatrix(x, .Col) = strChecked
                           mAmount = Val(mAmount) + Val(.TextMatrix(x, 4))
                        End If
                    Next
                    .TextMatrix(0, 4) = Format(Val(.TextMatrix(0, 4)) + Val(mAmount), "#####0.00")
                    mAmount = 0
                 End If
              End If
           Else
              mAmount = 0
              .TextMatrix(.row, .Col) = strUnChecked
          '    mAmount = Val(.TextMatrix(.Row, 4))
              mAmount = 0
              mPINo = .TextMatrix(.row, 13)
              For x = 2 To MSHFlexGrid3.Rows - 1
                  If mPINo = .TextMatrix(x, 13) Then
                     .TextMatrix(x, .Col) = strUnChecked
                      mAmount = Val(mAmount) + Val(.TextMatrix(x, 4))
                  End If
              Next
              .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) - Val(mAmount)
           End If
        End If
     
     End If

End With

End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Private Sub Option2_Click()
'CmdExport.Enabled = False
If Option2.Value = True Then
   CmbBatch.Enabled = True
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = False
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   ChkCheckAll.Enabled = True
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub

Private Sub PutTotInvAmt(mPINo_ As Variant, mNAVCode_ As Variant, mExpenseNAVCode_ As Variant, mROw_ As Variant)
Dim mAmtExclST, mRatio, mRatioAmt As Double

tmp = "Select PostingDate,DocumentType,AccountType,AccountNo,Amount,InvoiceNo,Narration,RegionDim,DeptDim,ClientNAVCode,DPNAVCode,PINo,PostInvoiceNo,EntryNo,RNAVCode,TIAmount,TISTAmount,ServiceTaxID"
tmp = tmp & " From ("
tmp = tmp & " Select FAGPostingDate as PostingDate,'Credit Memo'as DocumentType,'G/L Account' as AccountType,ME.NAVCode As AccountNo ,sum(TCRT.RecOtherDed) as Amount,'' as InvoiceNo,'Being Other deduction  by Customer ' as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,convert(varchar,TCR.CRTDSAdjustmentID) + '-'+ TCR.SM_Prefix as PINo,'' as PostInvoiceNo,Null as EntryNo,'' as RNAVCode,TI.TIID,TI.TIAmount,TI.TISTAmount,isnull(ME.ServiceTaxID,0)ServiceTaxID"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
tmp = tmp & " Left join Mst_ExpenseType ME on TCRT.ExpenseAccID=ME.ExpenseTypeID"
tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
tmp = tmp & " Where TCRT.RecOtherDed > 0 And TCR.BatchID =  " & mBatchID & " And TCR.Flag='A' And ME.ReverseFlag='Y' And TI.NAVCode='" & mNAVCode_ & "' and ME.NAVCode='" & mExpenseNAVCode_ & "'"
tmp = tmp & " Group By FAGPostingDate,ML.NAVCode,TI.NAVCode,TCR.CRTDSAdjustmentID,TCR.SM_Prefix,ME.NAVCode,TI.TIID,TI.TIAmount,TI.TISTAmount,ME.ServiceTaxID) A"
tmp = tmp & " Where PINO='" & mPINo_ & "' "

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   For r = 1 To TmpRs2.RecordCount
       mAmtExclST = 0
       mRatio = 0
       mRatioAmt = 0
       If TmpRs2!TISTAmount > 0 Then
          mAmtExclST = TmpRs2!TIAmount - TmpRs2!TISTAmount
          mRatio = TmpRs2!TIAmount / TmpRs2!TISTAmount
          mRatio = Format(TmpRs2!Amount / mRatio, "#####0.00")
          mRatioAmt = Format(TmpRs2!Amount - mRatio, "#####0.00")
       Else
          mRatioAmt = TmpRs2!Amount
       End If
       If MSHFlexGrid3.TextMatrix(mROw_ - 1, 2) = "Customer" And MSHFlexGrid3.TextMatrix(mROw_ - 1, 3) = mNAVCode_ Then
          MSHFlexGrid3.TextMatrix(mROw_ - 1, 4) = Val(MSHFlexGrid3.TextMatrix(mROw_ - 1, 4)) - mRatioAmt
''          If TmpRs2!ServiceTaxID > 0 Then
''             MSHFlexGrid3.TextMatrix(mRow_ - 1, 16) = GetServiceTaxNAVCode(TmpRs2!ServiceTaxID)
''          End If
          If TmpRs2!TISTAmount > 0 Then MSHFlexGrid3.TextMatrix(mROw_ - 1, 16) = "STWS"
          MSHFlexGrid3.TextMatrix(mROw_, 4) = Val(MSHFlexGrid3.TextMatrix(mROw_, 4)) + mRatioAmt
       End If
       
       TmpRs2.MoveNext
   Next
End If

End Sub
