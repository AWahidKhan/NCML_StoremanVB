VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportReturnPledgeCashReceipt 
   Caption         =   "Export Return Pledge Cash Receipt"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
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
         Left            =   6600
         TabIndex        =   11
         Top             =   235
         Width           =   3120
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
         Left            =   3120
         TabIndex        =   9
         Top             =   235
         Width           =   2040
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
         Left            =   1320
         TabIndex        =   8
         Top             =   255
         Width           =   855
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
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
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
         Left            =   11100
         TabIndex        =   6
         Top             =   240
         Width           =   1260
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
         Left            =   12360
         TabIndex        =   5
         Top             =   240
         Width           =   1260
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
         Left            =   13620
         TabIndex        =   4
         Top             =   240
         Width           =   1260
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
         Left            =   9840
         TabIndex        =   3
         Top             =   240
         Width           =   1260
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
         Left            =   5280
         TabIndex        =   12
         Top             =   288
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
         Left            =   2160
         TabIndex        =   10
         Top             =   288
         Width           =   1095
      End
   End
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
      Height          =   8025
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7680
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13547
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "Frm_ExportReturnPledgeCashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCodeTDS, mGLAccNAVCodeTDSDP, mGLAccNAVCodeOther As Variant
Dim ExpFlag, MaxID As Variant
Dim mJournalBatchName, mJournalType As Variant

Private Sub CmbBatch_GotFocus()
tmp = "select distinct RBatchID As BatchID from Txn_PledgeCashReceipt Where RBatchID > 0  order by RBatchID"
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


Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdExport_Click()
Dim Path, strFilName, strFilLocation As Variant
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

Path = App.Path
MaxID = GetMaxBatchID
If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_ReturnPledgeCashReceipt" & "_" & MaxID & ".csv"
Else
   strFilName = "\Navision\NavUpload_ReturnPledgeCashReceipt" & "_" & CmbBatch.Text & ".csv"
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
   Call TableTxn_CashRepeiptsForNAVISION(" Where G_UID Is Null ")
End If

With MSHFlexGrid3
For I = 2 To MSHFlexGrid3.Rows - 1
    tmp = ""
    If Option1.Value = True Then
    
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
    RsTxn_CashRepeiptsForNAVISION!EntryNo = IIf(Trim(.TextMatrix(I, 15)) = "", Null, Trim(.TextMatrix(I, 15)))
    'RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo = Null
    RsTxn_CashRepeiptsForNAVISION!CreatedBy = Gen_UserLoginID
    RsTxn_CashRepeiptsForNAVISION!G_UID = GetGUID
    RsTxn_CashRepeiptsForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
    'RsTxn_CashRepeiptsForNAVISION!DataTransfer = Null
    RsTxn_CashRepeiptsForNAVISION!InvoiceType = mInvoiceType
    RsTxn_CashRepeiptsForNAVISION!JournalBatchType = mJournalType
    RsTxn_CashRepeiptsForNAVISION!JournalBatchName = mJournalBatchName
    RsTxn_CashRepeiptsForNAVISION!ReverseInvoice = "Y"
    'RsTxn_CashRepeiptsForNAVISION!STCODE = Null
    'RsTxn_CashRepeiptsForNAVISION!BalAccType = Null
    'RsTxn_CashRepeiptsForNAVISION!BalAccountNo = Null
    'RsTxn_CashRepeiptsForNAVISION!EntryNoTransfer = Null
    'RsTxn_CashRepeiptsForNAVISION!EntryNoTransferRemark = Null
    RsTxn_CashRepeiptsForNAVISION!UserBatchID = mUserBatchID
    RsTxn_CashRepeiptsForNAVISION.Update
    
    Else
        For c = 0 To MSHFlexGrid3.Cols - 1
            If tmp = "" Then
               tmp = tmp & MSHFlexGrid3.TextMatrix(I, c)
            Else
               tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, c)
            End If
        Next
        Print #1, tmp
    End If
Next
End With

If Option1.Value = True Then
   tmp = "Update Txn_PledgeCashReceipt set RBatchID = " & MaxID & " Where ReturnFlag='Y' And RBatchID=0 "
   Call TmpTable
End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
End Sub
Private Function GetMaxBatchID()
tmp = "select Max(RBatchID) from Txn_PledgeCashReceipt"
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
tmp = tmp & " (Select FAGPostingDate as PostingDate,'Payment'as DocumentType,'G/L Account' as AccountType,'' As AccountNo ,sum(isnull(TCR.InstrumentAmount,0)) as Amount,'Being pledge Reverse for ' + ML.LocationName + ' RBatchID - " & mBatchID & "' as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'PCR' + convert(varchar,TCR.PCRID) + '-'+ TCR.SM_Prefix as PINo"
tmp = tmp & " From Txn_PledgeCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
If Option2.Value = True Then
   tmp = tmp & " Where isnull(TCR.RBatchID,0) = " & mBatchID & " And ReturnFlag='Y'"
Else
   tmp = tmp & " Where isnull(TCR.RBatchID,0) = 0 And ReturnFlag='Y'"
End If
tmp = tmp & " And TCR.FAGBankID > 0 And TCR.FAGPostingDate Is Not Null"
tmp = tmp & " Group By FAGPostingDate,ML.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,TCR.PCRID,TCR.SM_Prefix,ML.LocationName"
tmp = tmp & " Union All"
tmp = tmp & " Select FAGPostingDate as PostingDate,'Payment' as DocumentType,'Bank Account' as AccountType,MFG.NAVCode As AccountNo,InstrumentAmount as Amount,'Being pledge Reverse for '+ ML.LocationName + ' RBatchID - " & mBatchID & "' as Narration,'201001' as RegionDim,'COO' as DeptDim,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'PCR' + convert(varchar,TCR.PCRID) + '-'+ TCR.SM_Prefix as PINo"
tmp = tmp & " From Txn_PledgeCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
tmp = tmp & " left Join Mst_FAGBank  MFG on TCR.FAGBankID=MFG.FAGBankID"
If Option2.Value = True Then
   tmp = tmp & " Where isnull(TCR.RBatchID,0) = " & mBatchID & " And ReturnFlag='Y' "
Else
   tmp = tmp & " Where isnull(TCR.RBatchID,0) = 0  And ReturnFlag='Y'"
End If
tmp = tmp & " And TCR.FAGBankID > 0 And TCR.FAGPostingDate is Not Null"
tmp = tmp & " )A Order By PINo Asc,AccountType Desc"

Call TmpTable

mSum1 = 0

If TmpRs.RecordCount > 0 Then
   For r = 1 To TmpRs.RecordCount
       I = MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.TextMatrix(I, 0) = Format(TmpRs!PostingDate, "mm/dd/yyyy")
       MSHFlexGrid3.TextMatrix(I, 1) = TmpRs!DocumentType
       MSHFlexGrid3.TextMatrix(I, 2) = TmpRs!AccountType
       MSHFlexGrid3.TextMatrix(I, 6) = TmpRs!Narration
       
       If TmpRs!AccountType = "G/L Account" Then
          MSHFlexGrid3.TextMatrix(I, 3) = mGLAccNAVCodePledge
          If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
          MSHFlexGrid3.TextMatrix(I, 4) = TmpRs!Amount
       Else
          MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!AccountNo), "", TmpRs!AccountNo)
          If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
          MSHFlexGrid3.TextMatrix(I, 4) = "-" & Val(TmpRs!Amount)
       End If
       'MSHFlexGrid3.TextMatrix(I, 4) = TmpRs!Amount
       MSHFlexGrid3.TextMatrix(I, 5) = "" 'TmpRs!InvoiceNo
       MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(TmpRs!RegionDim), "", TmpRs!RegionDim)
       MSHFlexGrid3.TextMatrix(I, 8) = TmpRs!DeptDim
       MSHFlexGrid3.TextMatrix(I, 9) = "EMP-1000"
       
       If LCase(TmpRs!InstrumentType) <> "cash" And (TmpRs!AccountType = "Bank Account") Then
          MSHFlexGrid3.TextMatrix(I, 10) = TmpRs!InstrumentNo
          MSHFlexGrid3.TextMatrix(I, 11) = Format(TmpRs!InstrumentDate, "mm/dd/yyyy")
          MSHFlexGrid3.TextMatrix(I, 12) = Format(TmpRs!DepositDate, "mm/dd/yyyy")
       End If
       
       MSHFlexGrid3.TextMatrix(I, 13) = TmpRs!PINo

       
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       TmpRs.MoveNext
   Next
   If ExpFlag <> "N" Then
      CmdExport.Enabled = True
   End If

End If

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmbBatch.Text = ""
CmdExport.Enabled = False
Option1.Value = True
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
    .Cols = 16
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

End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Private Sub Option2_Click()
CmdExport.Enabled = False
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub
