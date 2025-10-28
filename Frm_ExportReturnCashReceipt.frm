VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportReturnCashReceipt 
   Caption         =   "Export Returned Cash Receipt "
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
      Height          =   8625
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   8280
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14605
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
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
         TabIndex        =   4
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
         Left            =   9840
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   240
         Width           =   1260
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
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   240
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
         Height          =   360
         Left            =   1320
         TabIndex        =   1
         Top             =   240
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
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   2040
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
         Height          =   360
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_ExportReturnCashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCodeTDS, mGLAccNAVCodeTDSDP, mGLAccNAVCodeOther As Variant
Dim ExpFlag, MaxID As Variant
Dim mJournalBatchName, mJournalType As Variant

Private Sub CmbBatch_GotFocus()
tmp = "select distinct RBatchID As BatchID from Txn_NewCashReceipt Where RBatchID > 0  order by RBatchID"
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
mInvoiceType = "CR"

Path = App.Path
MaxID = GetMaxBatchID
If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_ReturnCashReceipt" & "_" & MaxID & ".csv"
Else
   strFilName = "\Navision\NavUpload_ReturnCashReceipt" & "_" & CmbBatch.Text & ".csv"
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
For I = 2 To MSHFlexGrid3.Rows - 2
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

If Option1.Value = True Then
   tmp = "Update Txn_NewCashReceipt set RBatchID = " & MaxID & " Where ReturnFlag='Y' And RBatchID=0 "
   Call TmpTable
End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
End Sub

Private Function GetMaxBatchID()
tmp = "select Max(RBatchID) from Txn_NewCashReceipt"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   GetMaxBatchID = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0)) + 1
Else
   GetMaxBatchID = 1
End If
End Function
Private Sub CmdGetReport_Click()
Dim wc As Variant
Dim mRBatchID As Double
ExpFlag = ""
If Option1.Value = True Then
   mRBatchID = 0
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
   mRBatchID = Val(CmbBatch.Text)
End If

Call Grid_Head

mGLAccNAVCodeTDS = GetGLAccNAVCode("TDSNavCode")
mGLAccNAVCodeTDSDP = GetGLAccNAVCode("TDSNavCodeDP")
mGLAccNAVCodeOther = GetGLAccNAVCode("OtherDedNavCode")

''tmp = " Select PostingDate,DocumentType,AccountType,AccountNo,Amount,InvoiceNo,Narration,RegionDim,DeptDim,ClientNAVCode,DPNAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,PINo,PostInvoiceNo,EntryNo"
''tmp = tmp & " From"
''tmp = tmp & " ( Select ReturnDate as PostingDate,''as DocumentType,'Customer' as AccountType,'' As AccountNo ,(sum(isnull(TCRT.RecAmount,0))) as Amount,'' as InvoiceNo,TCR.ReturnRemark as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'R'+convert(varchar,TCR.CRID) + '-'+ TCR.SM_Prefix as PINo,'' PostInvoiceNo,Max(TCRT.EntryNo) as EntryNo"
''tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR  on TCRT.CRID=TCR.CRID And TCRT.SM_Prefix=TCR.SM_Prefix"
''tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
''tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
''tmp = tmp & " Where  TCR.ReturnFlag='Y' and TCR.RBatchID = " & mRBatchID
''tmp = tmp & " Group By DepositDate,ReturnDate,ML.NAVCode,TI.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,TCR.CRID,TCR.SM_Prefix,TCR.ReturnRemark "
''tmp = tmp & " Union"
''tmp = tmp & " Select ReturnDate as PostingDate,''as DocumentType,'Customer' as AccountType,'' As AccountNo ,(sum(TCRT.Amount)) as Amount,'' as InvoiceNo,TCR.ReturnRemark as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,MC.NAVCode as ClientNAVCode,MD.NAVCode as DPNAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'R'+convert(varchar,TCR.CRID) + '-'+ TCR.SM_Prefix as PINo,'' as PostInvoiceNo,Max(TCRT.EntryNo) as EntryNo"
''tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR  on TCRT.CRID=TCR.CRID And TCRT.SM_Prefix=TCR.SM_Prefix"
''tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
''tmp = tmp & " Left Join Mst_Client MC on TCRT.ClientIDRow=MC.ClientIDRow"
''tmp = tmp & " Left Join Mst_Depository MD on TCRT.DPID=MD.DepositoryID"
''tmp = tmp & " Where  TCR.ReturnFlag='Y' and TCRT.Amount > 0 And TCR.RBatchID = " & mRBatchID
''tmp = tmp & " Group By ReturnDate,DepositDate,ML.NAVCode,MC.NAVCode,MD.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,TCR.CRID,TCR.SM_Prefix,TCR.ReturnRemark "
''
''tmp = tmp & " ) A"
''tmp = tmp & " Order By PINO"
''
''Call TmpTable


'COmmented on 09 July 2012 for display of Return date in export issue raied by Basant ji
'tmp = " Select PostingDate,DocumentType,AccountType,AccountNo,sum(Amount)Amount,InvoiceNo,Narration,RegionDim,DeptDim,ClientNAVCode,"
'tmp = tmp & " DPNAVCode , InstrumentType, InstrumentNo, InstrumentDate, DepositDate, PINO, PostInvoiceNo, EntryNo, RNAVCode"
'tmp = tmp & " From"
'tmp = tmp & " (Select FAGPostingDate as PostingDate,''as DocumentType,'Customer' as AccountType,'' As AccountNo ,(sum(isnull(TCRT.RecAmount,0))+sum(isnull(TCRT.RecTDS,0))) as Amount,'' as InvoiceNo,TCR.ReturnRemark as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'R'+convert(varchar,TCR.CRID) + '-'+ TCR.SM_Prefix as PINo,'' PostInvoiceNo,Max(TCRT.EntryNo)EntryNo,'' As 'RNAVCode'"
'tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
'tmp = tmp & " Inner Join Txn_NewCashReceipt TCR  on TCRT.CRID=TCR.CRID And TCRT.SM_Prefix=TCR.SM_Prefix"
'tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
'tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
'tmp = tmp & " Where TCR.ReturnFlag='Y' And TCR.FAGBankID > 0 And TCR.FAGPostingDate Is Not Null And IsNull(TCRT.CRAAID, 0) = 0  And TCR.RBatchID = " & mRBatchID
'tmp = tmp & " Group By FAGPostingDate,TCR.ReturnRemark,ML.NAVCode,TI.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,TCR.CRID,TCR.SM_Prefix"
'tmp = tmp & " Union"
'tmp = tmp & " Select FAGPostingDate as PostingDate,'Payment'as DocumentType,'Customer' as AccountType,'' As AccountNo ,(sum(TCRT.Amount)) as Amount,'' as InvoiceNo,TCR.ReturnRemark as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,isnull(TCRT.AdvNAVCode,MC.NAVCode) as ClientNAVCode,isnull(MD.NAVCode,'') as DPNAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'R'+convert(varchar,TCR.CRID) + '-'+ TCR.SM_Prefix as PINo,'' as PostInvoiceNo,Max(TCRT.EntryNo)EntryNo,isnull(MC.RNAVCode ,'')"
'tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRT"
'tmp = tmp & " Inner Join Txn_NewCashReceipt TCR  on TCRT.CRID=TCR.CRID And TCRT.SM_Prefix=TCR.SM_Prefix"
'tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
'tmp = tmp & " Left Join Mst_Client MC on TCRT.ClientIDRow=MC.ClientIDRow"
'tmp = tmp & " Left Join Mst_Depository MD on TCRT.DPID=MD.DepositoryID"
'tmp = tmp & " Where TCR.ReturnFlag='Y' And TCRT.Amount > 0 And TCR.FAGBankID > 0 And TCR.FAGPostingDate Is Not Null  And TCR.RBatchID = " & mRBatchID
'tmp = tmp & " Group By FAGPostingDate,TCR.ReturnRemark,ML.NAVCode,TCRT.AdvNAVCode,MC.NAVCode,MD.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,TCR.CRID,TCR.SM_Prefix,MC.RNAVCode"
'tmp = tmp & " Union"
'tmp = tmp & " Select FAGPostingDate as PostingDate,''as DocumentType,'G/L Account' as AccountType,'TDS'As AccountNo,-(sum(TCRT.RecTDS)) as Amount,'' as InvoiceNo,TCR.ReturnRemark as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,TI.NAVCode as DPNAVCode,'' as InstrumentType,'' as InstrumentNo,'' as InstrumentDate,'' as DepositDate,'R'+convert(varchar,TCR.CRID) + '-'+ TCR.SM_Prefix as PINo,'' as PostInvoiceNo,Null EntryNo,'' As 'RNAVCode'"
'tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
'tmp = tmp & " Inner Join Txn_NewCashReceipt TCR  on TCRT.CRID=TCR.CRID And TCRT.SM_Prefix=TCR.SM_Prefix"
'tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
'tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
'tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
'tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
'tmp = tmp & " Where TCR.ReturnFlag='Y' And TCRT.RecTDS > 0 And TCR.RBatchID = " & mRBatchID & " And TCR.FAGBankID > 0 And TCR.FAGPostingDate Is Not Null"
'tmp = tmp & " Group By FAGPostingDate,TCR.ReturnRemark,ML.NAVCode,TI.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,TCR.CRID,TCR.SM_Prefix"
'tmp = tmp & " ) A"
'tmp = tmp & " group by PostingDate,DocumentType,AccountType,AccountNo,InvoiceNo,Narration,RegionDim,DeptDim,ClientNAVCode,"
'tmp = tmp & " DPNAVCode , InstrumentType, InstrumentNo, InstrumentDate, DepositDate, PINO, PostInvoiceNo, EntryNo, RNAVCode"
'tmp = tmp & " Order By PINo"


tmp = " Select PostingDate,DocumentType,AccountType,AccountNo,sum(Amount)Amount,InvoiceNo,Narration,RegionDim,DeptDim,ClientNAVCode,"
tmp = tmp & " DPNAVCode , InstrumentType, InstrumentNo, InstrumentDate, DepositDate, PINO, PostInvoiceNo, EntryNo, RNAVCode"
tmp = tmp & " From"
tmp = tmp & " (Select ReturnDate as PostingDate,''as DocumentType,'Customer' as AccountType,'' As AccountNo ,(sum(isnull(TCRT.RecAmount,0))+sum(isnull(TCRT.RecTDS,0))) as Amount,'' as InvoiceNo,TCR.ReturnRemark as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,'' as DPNAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'R'+convert(varchar,TCR.NewCRID) + '-'+ TCR.SM_Prefix as PINo,'' PostInvoiceNo,Max(TCRT.EntryNo)EntryNo,'' As 'RNAVCode'"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR  on TCRT.NewCRID=TCR.NewCRID And TCRT.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
tmp = tmp & " Where TCR.ReturnFlag='Y' And TCR.FAGBankID > 0 And TCR.FAGPostingDate Is Not Null And IsNull(TCRT.CRAAID, 0) = 0  And TCR.RBatchID = " & mRBatchID
tmp = tmp & " Group By ReturnDate,TCR.ReturnRemark,ML.NAVCode,TI.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,TCR.NewCRID,TCR.SM_Prefix"
tmp = tmp & " Union"
tmp = tmp & " Select ReturnDate as PostingDate,'Payment'as DocumentType,'Customer' as AccountType,'' As AccountNo ,(sum(TCRT.Amount)) as Amount,'' as InvoiceNo,TCR.ReturnRemark as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,isnull(TCRT.AdvNAVCode,MC.NAVCode) as ClientNAVCode,isnull(MD.NAVCode,'') as DPNAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'R'+convert(varchar,TCR.NewCRID) + '-'+ TCR.SM_Prefix as PINo,'' as PostInvoiceNo,Max(TCRT.EntryNo)EntryNo,isnull(MC.RNAVCode ,'')"
tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR  on TCRT.NewCRID=TCR.NewCRID And TCRT.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_Client MC on TCRT.ClientID=MC.ClientID"
tmp = tmp & " Left Join Mst_Depository MD on TCRT.DPID=MD.DepositoryID"
tmp = tmp & " Where TCR.ReturnFlag='Y' And TCRT.Amount > 0 And TCR.FAGBankID > 0 And TCR.FAGPostingDate Is Not Null  And TCR.RBatchID = " & mRBatchID
tmp = tmp & " Group By ReturnDate,TCR.ReturnRemark,ML.NAVCode,TCRT.AdvNAVCode,MC.NAVCode,MD.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,TCR.NewCRID,TCR.SM_Prefix,MC.RNAVCode"
tmp = tmp & " Union"
tmp = tmp & " Select ReturnDate as PostingDate,''as DocumentType,'G/L Account' as AccountType,'TDS'As AccountNo,-(sum(TCRT.RecTDS)) as Amount,'' as InvoiceNo,TCR.ReturnRemark as Narration ,ML.NAVCode as RegionDim,'CMP' as DeptDim,TI.NAVCode as ClientNAVCode,TI.NAVCode as DPNAVCode,'' as InstrumentType,'' as InstrumentNo,'' as InstrumentDate,'' as DepositDate,'R'+convert(varchar,TCR.NewCRID) + '-'+ TCR.SM_Prefix as PINo,'' as PostInvoiceNo,Null EntryNo,'' As 'RNAVCode'"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR  on TCRT.NewCRID=TCR.NewCRID And TCRT.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_Client MC on TI.ClientID=MC.ClientID"
tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
tmp = tmp & " Where TCR.ReturnFlag='Y' And TCRT.RecTDS > 0 And TCR.RBatchID = " & mRBatchID & " And TCR.FAGBankID > 0 And TCR.FAGPostingDate Is Not Null"
tmp = tmp & " Group By ReturnDate,TCR.ReturnRemark,ML.NAVCode,TI.NAVCode,InstrumentType,InstrumentNo,InstrumentDate,TCR.NewCRID,TCR.SM_Prefix"
tmp = tmp & " ) A"
tmp = tmp & " group by PostingDate,DocumentType,AccountType,AccountNo,InvoiceNo,Narration,RegionDim,DeptDim,ClientNAVCode,"
tmp = tmp & " DPNAVCode , InstrumentType, InstrumentNo, InstrumentDate, DepositDate, PINO, PostInvoiceNo, EntryNo, RNAVCode"
tmp = tmp & " Order By PINo"

Call TmpTable

tmp = " Select ReturnDate as PostingDate,'' as DocumentType,'Bank Account' as AccountType,MFG.NAVCode As AccountNo,-(InstrumentAmount) as Amount,'' as InvoiceNo,TCR.ReturnRemark as Narration,'201001' as RegionDim,'COO' as DeptDim,Null as DepositoryID,Null as ClientID,'' as ClientNAVCode,'' as DPNAVCode,InstrumentType,InstrumentNo,InstrumentDate,DepositDate,'R'+ convert(varchar,TCR.NewCRID) + '-'+ TCR.SM_Prefix as PINo"
tmp = tmp & " From Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML on TCR.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_FAGBank  MFG on TCR.FAGBankID=MFG.FAGBankID"
tmp = tmp & " Where TCR.ReturnFlag='Y' and TCR.RBatchID = " & mRBatchID
tmp = tmp & " Order By PINO"

Call Tmp1Table

If TmpRs.RecordCount > 0 Then
   For r = 1 To TmpRs.RecordCount
       I = MSHFlexGrid3.Rows - 1
       If I > 2 Then
          If MSHFlexGrid3.TextMatrix(I - 1, 13) <> TmpRs!PINo Then
             Call FillBankEntry(MSHFlexGrid3.TextMatrix(I - 1, 13))
          End If
       End If
       I = MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.TextMatrix(I, 0) = Format(TmpRs!PostingDate, "mm/dd/yyyy") 'Format(TmpRs!PostingDate, "dd/mm/yyyy")
       MSHFlexGrid3.TextMatrix(I, 1) = TmpRs!DocumentType
       MSHFlexGrid3.TextMatrix(I, 2) = TmpRs!AccountType
       MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs!Narration), "", TmpRs!Narration)
       If TmpRs!AccountType = "Customer" Then
          If IsNull(TmpRs!ClientNAVCode) = False Then
             MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
          Else
             MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!DPNAVCode), "", TmpRs!DPNAVCode)
          End If
          If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
       ElseIf TmpRs!AccountType = "G/L Account" Then
''          If IsNull(TmpRs!ClientNAVCode) = False Then
''             MSHFlexGrid3.TextMatrix(I, 6) = TmpRs!Narration & " " & IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
''          Else
''             MSHFlexGrid3.TextMatrix(I, 6) = TmpRs!Narration & " " & IIf(IsNull(TmpRs!DPNAVCode), "", TmpRs!DPNAVCode)
''          End If
          If TmpRs!AccountNo = "Other" Then
             MSHFlexGrid3.TextMatrix(I, 3) = mGLAccNAVCodeOther
          Else
             If IsNull(TmpRs!ClientNAVCode) = False Then
                MSHFlexGrid3.TextMatrix(I, 3) = mGLAccNAVCodeTDS
             Else
                MSHFlexGrid3.TextMatrix(I, 3) = mGLAccNAVCodeTDSDP
             End If
          End If
          If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
       End If
       MSHFlexGrid3.TextMatrix(I, 4) = TmpRs!Amount
       'MSHFlexGrid3.TextMatrix(I, 5) = TmpRs!InvoiceNo
       MSHFlexGrid3.TextMatrix(I, 7) = TmpRs!RegionDim
       MSHFlexGrid3.TextMatrix(I, 8) = TmpRs!DeptDim
       MSHFlexGrid3.TextMatrix(I, 9) = "EMP-1000"
       
       If LCase(TmpRs!InstrumentType) <> "cash" And (TmpRs!AccountType = "Bank Account" Or TmpRs!AccountType = "Customer") Then
          MSHFlexGrid3.TextMatrix(I, 10) = TmpRs!InstrumentNo
          MSHFlexGrid3.TextMatrix(I, 11) = Format(TmpRs!InstrumentDate, "mm/dd/yyyy") 'Format(TmpRs!InstrumentDate, "dd/mm/yyyy")
          MSHFlexGrid3.TextMatrix(I, 12) = Format(TmpRs!DepositDate, "mm/dd/yyyy") 'Format(TmpRs!DepositDate, "dd/mm/yyyy")
       End If
       MSHFlexGrid3.TextMatrix(I, 13) = TmpRs!PINo
       'MSHFlexGrid3.TextMatrix(I, 14) = IIf(IsNull(TmpRs!PostInvoiceNo), "", TmpRs!PostInvoiceNo)
       MSHFlexGrid3.TextMatrix(I, 15) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       TmpRs.MoveNext
   Next
   Call FillBankEntry(MSHFlexGrid3.TextMatrix(I, 13))
   If ExpFlag = "N" Then
      CmdExport.Enabled = False
   Else
      CmdExport.Enabled = True
   End If
Else
   CmdExport.Enabled = False
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

Private Sub FillBankEntry(CRID_ As Variant)
If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
TmpRs1.Find "PINO= '" & CRID_ & "'"

If TmpRs1.EOF = False Then
   I = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.TextMatrix(I, 0) = Format(TmpRs1!PostingDate, "mm/dd/yyyy") 'Format(TmpRs1!PostingDate, "dd/mm/yyyy")
   MSHFlexGrid3.TextMatrix(I, 1) = TmpRs1!DocumentType
   MSHFlexGrid3.TextMatrix(I, 2) = TmpRs1!AccountType
   MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!AccountNo), "", TmpRs1!AccountNo)
   If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
   MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!Narration), "", TmpRs1!Narration)
   MSHFlexGrid3.TextMatrix(I, 4) = TmpRs1!Amount
   MSHFlexGrid3.TextMatrix(I, 5) = TmpRs1!InvoiceNo
   MSHFlexGrid3.TextMatrix(I, 7) = TmpRs1!RegionDim
   MSHFlexGrid3.TextMatrix(I, 8) = TmpRs1!DeptDim
   MSHFlexGrid3.TextMatrix(I, 9) = "EMP-1000"
   If LCase(TmpRs1!InstrumentType) <> "cash" Then
      MSHFlexGrid3.TextMatrix(I, 10) = TmpRs1!InstrumentNo
      MSHFlexGrid3.TextMatrix(I, 11) = Format(TmpRs1!InstrumentDate, "mm/dd/yyyy") 'Format(TmpRs1!InstrumentDate, "dd/mm/yyyy")
      MSHFlexGrid3.TextMatrix(I, 12) = Format(TmpRs1!DepositDate, "mm/dd/yyyy") 'Format(TmpRs1!DepositDate, "dd/mm/yyyy")
   End If
   MSHFlexGrid3.TextMatrix(I, 13) = TmpRs1!PINo
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If
End Sub

