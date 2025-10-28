VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportPledgeCashReceipt 
   Caption         =   "Export Pledge Cash Receipt"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   10500
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
      Height          =   7545
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7200
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12700
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   0
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
         Left            =   5880
         TabIndex        =   3
         Top             =   240
         Width           =   3120
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         Top             =   255
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
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   705
         Width           =   855
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
         Left            =   2835
         TabIndex        =   2
         Top             =   240
         Width           =   1440
      End
      Begin VB.ComboBox CmbSDepositedBank 
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
         ItemData        =   "Frm_ExportPledgeCashReceipt.frx":0000
         Left            =   2835
         List            =   "Frm_ExportPledgeCashReceipt.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   675
         Width           =   10065
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
         TabIndex        =   9
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
         Left            =   13560
         TabIndex        =   11
         Top             =   720
         Width           =   1215
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
         Left            =   4560
         TabIndex        =   16
         Top             =   300
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
         Left            =   1320
         TabIndex        =   13
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label42 
         Caption         =   "Deposited Bank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   12
         Top             =   743
         Width           =   1680
      End
   End
End
Attribute VB_Name = "Frm_ExportPledgeCashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCodeTDS, mGLAccNAVCodeTDSDP, mGLAccNAVCodeOther As Variant
Dim ExpFlag, MaxID As Variant
Dim mSDepoBankID As Double
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mJournalBatchName, mJournalType As Variant

Private Sub ChkCheckAll_Click()
If MSHFlexGrid3.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If
If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid3.Rows - 2
       MSHFlexGrid3.row = I
       If MSHFlexGrid3.TextMatrix(I, 3) <> "" And MSHFlexGrid3.TextMatrix(I, 7) <> "" Then 'And MSHFlexGrid3.TextMatrix(I, 6) <> "" And MSHFlexGrid3.TextMatrix(I, 13) <> "" Then
          If mPINo <> MSHFlexGrid3.TextMatrix(I, 13) Then
             MSHFlexGrid3.row = I
             MSHFlexGrid3.TextMatrix(I, 16) = strChecked
             MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I, 4))
          End If
       Else
          mPINo = MSHFlexGrid3.TextMatrix(I, 13)
       End If
   Next
Else
   MSHFlexGrid3.TextMatrix(0, 4) = ""
   For I = 2 To MSHFlexGrid3.Rows - 2
       MSHFlexGrid3.row = I
       MSHFlexGrid3.TextMatrix(I, 16) = strUnChecked
   Next
End If
End Sub

Private Sub CmbBatch_GotFocus()
tmp = "select distinct BatchID As BatchID from Txn_PledgeCashReceipt Where BatchID > 0  order by BatchID"
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

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportPledgeCashReceipt.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportPledgeCashReceipt.xls")
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
mInvoiceType = "PCR"

''If Val(MSHFlexGrid3.TextMatrix(0, 4)) > 0 Then
''   MsgBox "Total of Col 'Amount' must be zero!!!"
''   Exit Sub
''End If

Path = App.Path
'MaxID = GetMaxBatchID
If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_PledgeCashReceipt" & "_" & MaxID & ".csv"
Else
   strFilName = "\Navision\NavUpload_PledgeCashReceipt" & "_" & CmbBatch.Text & ".csv"
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
    wc = ""
Else
   Call TableTxn_CashRepeiptsForNAVISION(" Where G_UID Is Null ")
End If

'If Option1.Value = True Then
'   If CmbSDepositedBank.Text <> "" Then
''      tmp = "Select * Txn_PledgeCashReceipt Where BatchID=0 And CRDepositID In ( select CRDepositID From Txn_PledgeCashReceiptDepositDetails Where FAGBankID = " & mSDepoBankID & ")"
'      Wc = " And CRDepositID In ( select CRDepositID From Txn_PledgeCashReceiptDepositDetails Where FAGBankID = " & mSDepoBankID & ") "
'   Else
''      tmp = "Select * Txn_PledgeCashReceipt Where BatchID=0 And CRDepositID > 0"
'      Wc = " And CRDepositID > 0 "
'   End If
'
'
''   Call Tmp1Table
'End If

With MSHFlexGrid3
For I = 2 To MSHFlexGrid3.Rows - 2
    tmp = ""
    If Option1.Value = True Then
       If MSHFlexGrid3.TextMatrix(I, 16) = strChecked Then
'          For c = 0 To MSHFlexGrid3.Cols - 2
'              If tmp = "" Then
'                 tmp = tmp & MSHFlexGrid3.TextMatrix(I, c)
'              Else
'                 tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, c)
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
            'RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo = Null
            'RsTxn_CashRepeiptsForNAVISION!EntryNo = Null
            'RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo = Null
            RsTxn_CashRepeiptsForNAVISION!CreatedBy = Gen_UserLoginID
            RsTxn_CashRepeiptsForNAVISION!G_UID = GetGUID
            RsTxn_CashRepeiptsForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
            'RsTxn_CashRepeiptsForNAVISION!DataTransfer = Null
            RsTxn_CashRepeiptsForNAVISION!InvoiceType = mInvoiceType
            RsTxn_CashRepeiptsForNAVISION!JournalBatchType = mJournalType
            RsTxn_CashRepeiptsForNAVISION!JournalBatchName = mJournalBatchName
            'RsTxn_CashRepeiptsForNAVISION!ReverseInvoice = ""
            'RsTxn_CashRepeiptsForNAVISION!STCODE = Null
            'RsTxn_CashRepeiptsForNAVISION!BalAccType = Null
            'RsTxn_CashRepeiptsForNAVISION!BalAccountNo = Null
            'RsTxn_CashRepeiptsForNAVISION!EntryNoTransfer = Null
            'RsTxn_CashRepeiptsForNAVISION!EntryNoTransferRemark = Null
            RsTxn_CashRepeiptsForNAVISION!UserBatchID = mUserBatchID
            RsTxn_CashRepeiptsForNAVISION.Update
                  
          If Option1.Value = True Then
            If MSHFlexGrid3.TextMatrix(I, 2) = "G/L Account" < 0 Then
               strSplitString = Split(Replace(MSHFlexGrid3.TextMatrix(I, 13), "PCR", ""), "-")
               CRID_ = strSplitString(0)
               Prefix_ = strSplitString(1)
            
               tmp = " Update Txn_PledgeCashReceipt Set BatchID = " & MaxID & " Where PCRID = " & CRID_ & " And SM_Prefix = '" & Prefix_ & "' And BatchID = 0  "
               Call Tmp1Table
            End If
          End If
       End If
    Else
       For C = 0 To MSHFlexGrid3.Cols - 1
           If tmp = "" Then
              tmp = tmp & MSHFlexGrid3.TextMatrix(I, C)
           Else
              tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, C)
           End If
       Next
       Print #1, tmp
    End If
Next
End With

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Close #1
Call CmdReset_Click
End Sub
Private Function GetMaxBatchID()

tmp = "select Max(BatchID) from Txn_PledgeCashReceipt"
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
Dim mBatchID As Double

ExpFlag = ""

If Option1.Value = True Then
   mBatchID = GetMaxBatchID 'MaxID = GetMaxBatchID
   MaxID = mBatchID
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
   mBatchID = Val(CmbBatch.Text)
End If

Call Grid_Head

mGLAccNAVCodePledge = GetGLAccNAVCode("PledgeNavCode")

tmp = "Select PostingDate,DocumentType,AccountType,AccountNo,Amount,Narration,RegionDim,DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo From"
tmp = tmp & " (Select FAGPostingDate as PostingDate,'Payment'as DocumentType,'G/L Account' as AccountType,'' As AccountNo ,sum(isnull(TCRT.RecAmount,0)) + TCR.ExcessAmount as Amount,'Being pledge received for ' + ML.LocationName + ' BatchID - " & mBatchID & "' as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'PCR' + convert(varchar,TCR.PCRID) + '-'+ TCR.SM_Prefix as PINo"
tmp = tmp & " From Txn_PledgeCashReceiptDetail TCRT"
tmp = tmp & " Inner Join Txn_PledgeCashReceipt TCR  on TCRT.PCRID=TCR.PCRID And TCRT.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
If Option2.Value = True Then
   tmp = tmp & " Where TCR.BatchID = " & mBatchID & " "
Else
   tmp = tmp & " Where TCR.BatchID = 0 "
End If
tmp = tmp & " And TCR.FAGBankID > 0 And TCR.FAGPostingDate Is Not Null"
If CmbSDepositedBank.Text <> "" Then
   tmp = tmp & " And TCR.FAGBankID = " & mSDepoBankID
End If
tmp = tmp & " Group By FAGPostingDate,ML.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,TCR.PCRID,TCR.SM_Prefix,ML.LocationName,TCR.ExcessAmount"
tmp = tmp & " Union All"
tmp = tmp & " Select FAGPostingDate as PostingDate,'Payment' as DocumentType,'Bank Account' as AccountType,MFG.NAVCode As AccountNo,InstrumentAmount as Amount,'Being pledge received for '+ ML.LocationName + ' BatchID - " & mBatchID & "' as Narration,'201001' as RegionDim,'COO' as DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'PCR' + convert(varchar,TCR.PCRID) + '-'+ TCR.SM_Prefix as PINo"
tmp = tmp & " From Txn_PledgeCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
tmp = tmp & " left Join Mst_FAGBank  MFG on TCR.FAGBankID=MFG.FAGBankID"
If Option2.Value = True Then
   tmp = tmp & " Where TCR.BatchID = " & mBatchID & "  "
Else
   tmp = tmp & " Where TCR.BatchID = 0  "
End If
tmp = tmp & " And TCR.FAGBankID > 0 And TCR.FAGPostingDate is Not Null"
If CmbSDepositedBank.Text <> "" Then
   tmp = tmp & " And TCR.FAGBankID = " & mSDepoBankID
End If
tmp = tmp & " )A Order By PINo Asc,AccountType Desc"

Call TmpTable

mSum1 = 0

If TmpRs.RecordCount > 0 Then
   For r = 1 To TmpRs.RecordCount
       I = MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.TextMatrix(I, 0) = Format(TmpRs!PostingDate, "mm/dd/yyyy") 'Format(TmpRs!PostingDate, "dd/mm/yyyy") 'Format(TmpRs!PostingDate, "mm/dd/yyyy")
       MSHFlexGrid3.TextMatrix(I, 1) = TmpRs!DocumentType
       MSHFlexGrid3.TextMatrix(I, 2) = TmpRs!AccountType
       MSHFlexGrid3.TextMatrix(I, 6) = TmpRs!Narration
       
       If TmpRs!AccountType = "G/L Account" Then
          MSHFlexGrid3.TextMatrix(I, 3) = mGLAccNAVCodePledge
          If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
          MSHFlexGrid3.TextMatrix(I, 4) = "-" & Val(TmpRs!Amount)
       Else
          MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!AccountNo), "", TmpRs!AccountNo)
          If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
          MSHFlexGrid3.TextMatrix(I, 4) = TmpRs!Amount
       End If
       'MSHFlexGrid3.TextMatrix(I, 4) = TmpRs!Amount
       
       MSHFlexGrid3.TextMatrix(I, 5) = "" 'TmpRs!InvoiceNo
       MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(TmpRs!RegionDim), "", TmpRs!RegionDim)
       MSHFlexGrid3.TextMatrix(I, 8) = TmpRs!DeptDim
       MSHFlexGrid3.TextMatrix(I, 9) = "EMP-1000"
       
       If LCase(TmpRs!InstrumentType) <> "cash" And (TmpRs!AccountType = "Bank Account") Then
          MSHFlexGrid3.TextMatrix(I, 10) = TmpRs!InstrumentNo
          MSHFlexGrid3.TextMatrix(I, 11) = Format(TmpRs!InstrumentDate, "mm/dd/yyyy")  'Format(TmpRs!InstrumentDate, "dd/mm/yyyy") 'Format(TmpRs!InstrumentDate, "mm/dd/yyyy")
          MSHFlexGrid3.TextMatrix(I, 12) = Format(TmpRs!DepositDate, "mm/dd/yyyy") 'Format(TmpRs!DepositDate, "dd/mm/yyyy") 'Format(TmpRs!DepositDate, "mm/dd/yyyy")
       End If
       
       MSHFlexGrid3.TextMatrix(I, 13) = TmpRs!PINo

       If Option1.Value = True Then
          MSHFlexGrid3.row = I
          MSHFlexGrid3.Col = 16
          MSHFlexGrid3.CellFontName = "Wingdings"
          MSHFlexGrid3.CellFontSize = 12
          MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid3.Text = strUnChecked
       End If
       
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       TmpRs.MoveNext
   Next
End If

End Sub

Private Sub CmdReset_Click()

Call Grid_Head

CmbBatch.Text = ""
CmdExport.Enabled = True
Option1.Value = True
ChkCheckAll.Value = 0

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
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    If Option1.Value = True Then
       .Cols = 17
    Else
       .Cols = 16
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
    If Option1.Value = True Then
       .TextMatrix(1, 16) = "Selection"
    End If
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
    .ColWidth(6) = 1600
    .ColWidth(7) = 1400
    .ColWidth(8) = 1400
    .ColWidth(9) = 1400
    .ColWidth(10) = 1400
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    
End With

End Sub

Private Sub MSHFlexGrid3_Click()

Dim mAmount As Variant

mAmount = 0
With MSHFlexGrid3
     If .Col = 16 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        .Col = 16
        If .TextMatrix(.row, .Col) = strUnChecked Then
           If .TextMatrix(.row, 3) <> "" And .TextMatrix(.row, 7) <> "" Then
                 mPINo = .TextMatrix(.row, 13)
                ' mAmount = Val(.TextMatrix(.Row, 4))
                 mAmount = 0
                 For x = 2 To MSHFlexGrid3.Rows - 2
                     If mPINo = .TextMatrix(x, 13) Then
                        If .TextMatrix(x, 3) <> "" And .TextMatrix(x, 7) <> "" Then
                           .TextMatrix(x, .Col) = strChecked
                           mAmount = Val(mAmount) + Val(.TextMatrix(x, 4))
                        Else
                           .TextMatrix(x, .Col) = strUnChecked
                           GoTo z:
                        End If
                     End If
                 Next
                 .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(mAmount)
                 mAmount = 0
                 .TextMatrix(.row, .Col) = strChecked
z:           End If
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
End With

End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Private Sub Option2_Click()
'CmdExport.Enabled = False
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbSDepositedBank.Text = ""
   CmbSDepositedBank.Enabled = False
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = False
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   CmbSDepositedBank.Enabled = True
   ChkCheckAll.Enabled = True
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub

Private Sub CmbSDepositedBank_GotFocus()
tmp = CmbSDepositedBank.Text
Call TableMst_FAGBank("")
CmbSDepositedBank.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbSDepositedBank.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
    CmbSDepositedBank.AddItem RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
    RsMst_FAGBank.MoveNext
Next
CmbSDepositedBank.Text = tmp
End Sub
Private Sub CmbSDepositedBank_LostFocus()
Dim strSplitString() As String
Dim AccountNo_, BankName_ As String

If CmbSDepositedBank.Text = "" Then
   Exit Sub
End If

strSplitString = Split(CmbSDepositedBank, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbSDepositedBank.SetFocus
   Exit Sub
End If

BankName_ = Trim(strSplitString(0))
AccountNo_ = Trim(strSplitString(1))

RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo = '" & AccountNo_ & "'"

If RsMst_FAGBank.EOF Then
   MsgBox "Invalid Deposited Bank selection "
   CmbSDepositedBank.SetFocus
   Exit Sub
End If
mSDepoBankID = RsMst_FAGBank!FAGBankID
End Sub


