VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportReverseNNNRInvoice 
   Caption         =   "Export Reverse NNNR Invoice"
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
      Height          =   7665
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7320
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12912
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1080
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
         Left            =   1560
         TabIndex        =   4
         Top             =   620
         Width           =   4440
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
         Left            =   13800
         TabIndex        =   9
         Top             =   225
         Width           =   1020
      End
      Begin VB.ComboBox CmbInvoiceType 
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
         ItemData        =   "Frm_ExportReverseNNNRInvoice.frx":0000
         Left            =   1560
         List            =   "Frm_ExportReverseNNNRInvoice.frx":0016
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2280
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
         Left            =   8880
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
         Left            =   12660
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
         Left            =   10140
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
         Height          =   330
         Left            =   3960
         TabIndex        =   1
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
         Left            =   5160
         TabIndex        =   2
         Top             =   270
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
         Left            =   7080
         TabIndex        =   3
         Top             =   240
         Width           =   1680
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
         Left            =   120
         TabIndex        =   15
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Invoice Type"
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
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1170
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
         Left            =   6120
         TabIndex        =   11
         Top             =   300
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_ExportReverseNNNRInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCodeNNNR As Variant
Dim ExpFlag, MaxID As Variant
Dim mJournalBatchName, mJournalType, mIncomeTypeID, mGLAccNAVCodeNAGRI, mGLAccNAVCodeAGRI, mHSNNAGRI, mHSNAGRI, mInvType As Variant

Private Sub CmbBatch_GotFocus()
If CmbInvoiceType.Text = "" Then
   MsgBox "Please select Invoice Type"
   CmbInvoiceType.SetFocus
   Exit Sub
End If
If LCase(CmbInvoiceType.Text) = "nnnr" Then
   tmp = "select distinct RBatchID As BatchID from Txn_ProInvoice Where RBatchID > 0  order by RBatchID"
ElseIf LCase(CmbInvoiceType.Text) = "7 days" Then
   tmp = "select distinct RBatchID As BatchID from Txn_CDTFBOE Where RBatchID > 0 And TxnType='D' order by RBatchID"
ElseIf LCase(CmbInvoiceType.Text) = "rrn" Then
   tmp = "select distinct RBatchID As BatchID from Txn_CDTFBOE Where RBatchID > 0 And TxnType='R' order by RBatchID"
ElseIf LCase(CmbInvoiceType.Text) = "nnnry" Then
   tmp = "select distinct RBatchID As BatchID from Txn_CDTFBOE Where RBatchID > 0 And TxnType='Y' order by RBatchID"
Else
   tmp = "select distinct RBatchID As BatchID from Txn_CDTFBOE Where RBatchID > 0 And TxnType='W' order by RBatchID"
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


Private Sub CmbInvoiceType_LostFocus()
If CmbInvoiceType.Text = "" Then Exit Sub

If LCase(CmbInvoiceType.Text) <> "nnnr" And LCase(CmbInvoiceType.Text) <> "7 days" And LCase(CmbInvoiceType.Text) <> "rejected" And LCase(CmbInvoiceType.Text) <> "rrn" And LCase(CmbInvoiceType.Text) <> "nnnry" And LCase(CmbInvoiceType.Text) <> "pwr" Then
   MsgBox "Invalid Selection!!!"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

Call Grid_Head

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
'mJournalType = "SALES"
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

Path = App.Path
If CmbInvoiceType.Text = "" Then
   MsgBox "Please select Invoice Type"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

MaxID = GetMaxBatchID
If Option1.Value = True Then
   If LCase(CmbInvoiceType.Text) = "nnnr" Then
      strFilName = "\Navision\NavUpload_ReverseNNNRInvoice" & "_" & MaxID & ".csv"
      mInvoiceType = "NNNR"
   ElseIf LCase(CmbInvoiceType.Text) = "7 days" Then
      strFilName = "\Navision\NavUpload_Reverse7DaysInvoice" & "_" & MaxID & ".csv"
      mInvoiceType = "7Days"
   ElseIf LCase(CmbInvoiceType.Text) = "rrn" Then
      strFilName = "\Navision\NavUpload_ReverseRRNInvoice" & "_" & MaxID & ".csv"
      mInvoiceType = "RRN"
   ElseIf LCase(CmbInvoiceType.Text) = "nnnry" Then
      strFilName = "\Navision\NavUpload_ReverseNNNRYearlyInvoice" & "_" & MaxID & ".csv"
      mInvoiceType = "NNNRY"
   ElseIf LCase(CmbInvoiceType.Text) = "pwr" Then
      strFilName = "\Navision\NavUpload_ReversePWRInvoice" & "_" & MaxID & ".csv"
      mInvoiceType = "PWR"
   Else
      strFilName = "\Navision\NavUpload_ReverseRejectedInvoice" & "_" & MaxID & ".csv"
      mInvoiceType = "REJECTED"
   End If
Else
   If LCase(CmbInvoiceType.Text) = "nnnr" Then
      strFilName = "\Navision\NavUpload_ReverseNNNRInvoice" & "_" & CmbBatch.Text & ".csv"
      mInvoiceType = "NNNR"
   ElseIf LCase(CmbInvoiceType.Text) = "7 days" Then
      strFilName = "\Navision\NavUpload_Reverse7DaysInvoice" & "_" & CmbBatch.Text & ".csv"
      mInvoiceType = "7Days"
   ElseIf LCase(CmbInvoiceType.Text) = "rrn" Then
      strFilName = "\Navision\NavUpload_ReverseRRNInvoice" & "_" & CmbBatch.Text & ".csv"
      mInvoiceType = "RRN"
   ElseIf LCase(CmbInvoiceType.Text) = "nnnry" Then
      strFilName = "\Navision\NavUpload_ReverseNNNRYearlyInvoice" & "_" & CmbBatch.Text & ".csv"
      mInvoiceType = "NNNRY"
   ElseIf LCase(CmbInvoiceType.Text) = "pwr" Then
      strFilName = "\Navision\NavUpload_ReversePWRInvoice" & "_" & CmbBatch.Text & ".csv"
      mInvoiceType = "PWR"
   Else
      strFilName = "\Navision\NavUpload_ReverseRejectedInvoice" & "_" & CmbBatch.Text & ".csv"
      mInvoiceType = "REJECTED"
   End If
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
    Call TableTxn_InvoicesForNAVISION(" Where G_UID Is Null ")
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
End If

For I = 2 To MSHFlexGrid3.Rows - 2
    tmp = ""
    If Option1.Value = True Then
        
        RsTxn_InvoicesForNAVISION.AddNew
        RsTxn_InvoicesForNAVISION!Date = Trim(MSHFlexGrid3.TextMatrix(I, 0))
        RsTxn_InvoicesForNAVISION!Invoice = Trim(MSHFlexGrid3.TextMatrix(I, 1))
        RsTxn_InvoicesForNAVISION!Customer = Trim(MSHFlexGrid3.TextMatrix(I, 2))
        RsTxn_InvoicesForNAVISION!NavCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
        RsTxn_InvoicesForNAVISION!Amount = Trim(MSHFlexGrid3.TextMatrix(I, 4))
        RsTxn_InvoicesForNAVISION!Narration = Trim(MSHFlexGrid3.TextMatrix(I, 6))
        RsTxn_InvoicesForNAVISION!LocationNAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 7))
        RsTxn_InvoicesForNAVISION!DeptDim = Trim(MSHFlexGrid3.TextMatrix(I, 8))
        RsTxn_InvoicesForNAVISION!EmpDim = Trim(MSHFlexGrid3.TextMatrix(I, 9))
        RsTxn_InvoicesForNAVISION!PINo = Trim(MSHFlexGrid3.TextMatrix(I, 13))
        'RsTxn_InvoicesForNAVISION!PostInvoiceNo = Null
        RsTxn_InvoicesForNAVISION!EntryNo = Trim(MSHFlexGrid3.TextMatrix(I, 16))
        RsTxn_InvoicesForNAVISION!STCode = Trim(MSHFlexGrid3.TextMatrix(I, 15))
        RsTxn_InvoicesForNAVISION!GLAccount = Trim(MSHFlexGrid3.TextMatrix(I, 17))
        RsTxn_InvoicesForNAVISION!GLACNAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 18))
        
        RsTxn_InvoicesForNAVISION!GSTGroupCode = Trim(MSHFlexGrid3.TextMatrix(I, 19))
        RsTxn_InvoicesForNAVISION!HSNNo = Trim(MSHFlexGrid3.TextMatrix(I, 20))
        RsTxn_InvoicesForNAVISION!AccountType = Trim(MSHFlexGrid3.TextMatrix(I, 21))
        

        RsTxn_InvoicesForNAVISION!CreatedBy = Gen_UserLoginID
        RsTxn_InvoicesForNAVISION!G_UID = GetGUID
        RsTxn_InvoicesForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
        'RsTxn_InvoicesForNAVISION!DataTransfer = Null
        RsTxn_InvoicesForNAVISION!InvoiceType = mInvoiceType
        RsTxn_InvoicesForNAVISION!JournalBatchType = Trim(mJournalType)
        RsTxn_InvoicesForNAVISION!JournalBatchName = Trim(mJournalBatchName)
        RsTxn_InvoicesForNAVISION!ReverseInvoice = "Y"
        'RsTxn_InvoicesForNAVISION!EntryNoTransfer = Null
        'RsTxn_InvoicesForNAVISION!EntryNoTransferRemark = Null
        RsTxn_InvoicesForNAVISION!UserBatchID = mUserBatchID
        'RsTxn_InvoicesForNAVISION!DataTransfer = "X"
        RsTxn_InvoicesForNAVISION.Update
        
'        For c = 0 To MSHFlexGrid3.Cols - 1
'            If tmp = "" Then
'               tmp = tmp & MSHFlexGrid3.TextMatrix(i, c)
'            Else
'               tmp = tmp & "," & MSHFlexGrid3.TextMatrix(i, c)
'            End If
'        Next
'        Print #1, tmp
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

If Option1.Value = True Then
   If LCase(CmbInvoiceType.Text) = "nnnr" Then
      tmp = "Update Txn_ProInvoice set RBatchID = " & MaxID & " Where ReverseFlag='Y' And RBatchID=0 "
   ElseIf LCase(CmbInvoiceType.Text) = "7 days" Then
      tmp = "Update Txn_CDTFBOE set RBatchID = " & MaxID & " Where ReverseFlag='Y' And RBatchID=0 And TxnType='D'"
   ElseIf LCase(CmbInvoiceType.Text) = "rrn" Then
      tmp = "Update Txn_CDTFBOE set RBatchID = " & MaxID & " Where ReverseFlag='Y' And RBatchID=0 And TxnType='R'"
   ElseIf LCase(CmbInvoiceType.Text) = "nnnry" Then
      tmp = "Update Txn_CDTFBOE set RBatchID = " & MaxID & " Where ReverseFlag='Y' And RBatchID=0 And TxnType='Y'"
   ElseIf LCase(CmbInvoiceType.Text) = "pwr" Then
      tmp = "Update Txn_CDTFBOE set RBatchID = " & MaxID & " Where ReverseFlag='Y' And RBatchID=0 And TxnType='P'"
   Else
      tmp = "Update Txn_CDTFBOE set RBatchID = " & MaxID & " Where ReverseFlag='Y' And RBatchID=0 And TxnType='W'"
   End If
   Call TmpTable
End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Close #1
Call Grid_Head
End Sub
Private Function GetMaxBatchID()
If LCase(CmbInvoiceType.Text) = "nnnr" Then
   tmp = "select Max(RBatchID) from Txn_ProInvoice"
Else
   tmp = "select Max(RBatchID) from Txn_CDTFBOE"
End If
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

If CmbInvoiceType.Text = "" Then
   MsgBox "Please select Invoice Type"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

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

'''tmp = " Select TPI.PINo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MC.NAVCode 'ClientNAVCode',MC.RNAVCode as 'RClientNAVCode',-(TPI.Amount) Amount,TPI.BatchID,MS.NAVCode, TPI.ReverseDate,TPI.ReverseRemark,TPI.InvoiceDate,TPI.InvoiceTypeID"
'''tmp = tmp & " From Txn_TempInvoice TPI"
'''tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
'''tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
'''tmp = tmp & " Inner Join Mst_ServiceTax MS on MS.ServiceTaxID=TPI.ServiceTaxID"
'''tmp = tmp & " Where ReverseFlag='Y' And RBatchID=" & mRBatchID
'''tmp = tmp & " Order By TPI.PIDate"

mNumberSeriesFlag = ""
mGLAccNAVCodeAGRI = ""
mGLAccNAVCodeNAGRI = ""
mHSNAGRI = ""
mHSNNAGRI = ""

If LCase(CmbInvoiceType.Text) = "nnnr" Then

    'mGLAccNAVCodeNNNR = GetGLAccNAVCode("NNNC Taxable")
    
    
    mIncomeTypeID = 7
    
    tmp = "Select TPI.InvoiceNo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,TPI.ServiceTaxNAVCode,TPI.ReverseDate,TPI.ReverseRemark"
    tmp = tmp & " ,TPI.InvoiceDate , TPI.Discount, TPI.Premium, TPI.FAGPostingDate,MST.InvoiceType,TPI.HSNNo "
    tmp = tmp & " From Txn_ProInvoice TPI"
    tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
    tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
    tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"
    tmp = tmp & " Inner Join Mst_ServiceTax MST on TPI.ServiceTaxNAVCode = MST.NavCode"
    tmp = tmp & " Where ReverseFlag='Y' And RBatchID=" & mRBatchID
    tmp = tmp & " Order By TPI.InvoiceDate"
    
    mGLAccNAVCodeAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "G")
    mGLAccNAVCodeNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "G")
    mHSNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "H")
    mHSNNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "H")
   
    
ElseIf LCase(CmbInvoiceType.Text) = "7 days" Then

    'mGLAccNAVCodeNNNR = GetGLAccNAVCode("7 Days Taxable")

    tmp = "Select TPI.BOEID as InvoiceNo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,MST.NAVCode ServiceTaxNAVCode,TPI.ReverseDate,TPI.ReverseRemark"
    tmp = tmp & " ,TPI.InvoiceDate,isnull(TPI.Discount,0)Discount, isnull(TPI.Premium,0)Premium, TPI.FAGPostingDate,MST.InvoiceType,MComm.HSNNo "
    tmp = tmp & " From Txn_CDTFBOE TPI"
    tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
    tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType='D' "
    tmp = tmp & " inner join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
    tmp = tmp & " inner join Mst_ServiceTax MST On MComm.ServiceTaxID = MST.ServiceTaxID "
    tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
    tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"
    tmp = tmp & " Where ReverseFlag='Y' And RBatchID=" & mRBatchID & " and TPI.TxnType='D'"
    tmp = tmp & " Order By TPI.InvoiceDate"
    
    mIncomeTypeID = 1

    mGLAccNAVCodeAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "G")
    mGLAccNAVCodeNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "G")
    mHSNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "H")
    mHSNNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "H")

''tmp = "Select Distinct TPI.BOEID,TPI.SM_Prefix,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MC.NAVCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,MS.NAVCode,"
''tmp = tmp & " TPI.InvoiceDate , isnull(TPI.Discount,0)Discount, isnull(TPI.Premium,0)Premium, TPI.FAGPostingDate,MComm.Commodity "
''tmp = tmp & " From Txn_CDTFBOE TPI"
''tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
''tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
''tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CDTFNo = TC.CxTFNo And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType='D' "
''tmp = tmp & " inner join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
''tmp = tmp & " inner join Mst_ServiceTax MS On MComm.ServiceTaxID = MS.ServiceTaxID "

ElseIf LCase(CmbInvoiceType.Text) = "rrn" Then

    'mGLAccNAVCodeNNNR = GetGLAccNAVCode("RRN Withdrawal Taxable")

    tmp = "Select TPI.BOEID as InvoiceNo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,MST.NAVCode ServiceTaxNAVCode,TPI.ReverseDate,TPI.ReverseRemark"
    tmp = tmp & " ,TPI.InvoiceDate,isnull(TPI.Discount,0)Discount, isnull(TPI.Premium,0)Premium, TPI.FAGPostingDate,MST.InvoiceType,MComm.HSNNo,Sum(TCG.NCDEXSpaceCharges) 'NCDEXSpaceCharges' "
    tmp = tmp & " From Txn_CDTFBOE TPI"
    tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
    tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType = 'W' and TPI.TxnType='R' "
    tmp = tmp & " inner join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
    tmp = tmp & " inner join Mst_ServiceTax MST On MComm.ServiceTaxID = MST.ServiceTaxID "
    tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
    tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"
    tmp = tmp & " Inner join Txn_CDTFBOE_GSL TCG on TPI.BOEID = TCG.BOEID "
    tmp = tmp & " Where ReverseFlag='Y' And RBatchID=" & mRBatchID & " and TPI.TxnType='R'"
    tmp = tmp & " Group By TPI.BOEID,ML.LocationName,ML.NAVCode,MC.ClientName,MCSG.GSTNavCode,TPI.Amount,TPI.BatchID,MST.NAVCode,TPI.ReverseDate,TPI.ReverseRemark"
    tmp = tmp & " ,TPI.InvoiceDate,TPI.Discount, TPI.Premium, TPI.FAGPostingDate,MST.InvoiceType,MComm.HSNNo"
    tmp = tmp & " Order By TPI.InvoiceDate"

    mIncomeTypeID = 13

    mGLAccNAVCodeAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "G")
    mGLAccNAVCodeNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "G")
    mHSNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "H")
    mHSNNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "H")


ElseIf LCase(CmbInvoiceType.Text) = "nnnry" Then

    'mGLAccNAVCodeNNNR = GetGLAccNAVCode("NNNC Taxable")

    tmp = "Select TPI.BOEID as InvoiceNo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode  'ClientNAVCode',TPI.Amount,TPI.BatchID,MST.NAVCode ServiceTaxNAVCode,TPI.ReverseDate,TPI.ReverseRemark"
    tmp = tmp & " ,TPI.InvoiceDate,isnull(TPI.Discount,0)Discount, isnull(TPI.Premium,0)Premium, TPI.FAGPostingDate,MST.InvoiceType,MComm.HSNNo"
    tmp = tmp & " From Txn_CDTFBOE TPI"
    tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
    tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType = 'W' and TPI.TxnType='Y' "
    tmp = tmp & " inner join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
    tmp = tmp & " inner join Mst_ServiceTax MST On MComm.ServiceTaxID = MST.ServiceTaxID "
    tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
    tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"
    tmp = tmp & " Where ReverseFlag='Y' And RBatchID=" & mRBatchID & " and TPI.TxnType='Y'"
    tmp = tmp & " Order By TPI.InvoiceDate"

    mIncomeTypeID = 1

    mGLAccNAVCodeAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "G")
    mGLAccNAVCodeNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "G")
    mHSNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "H")
    mHSNNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "H")

ElseIf LCase(CmbInvoiceType.Text) = "pwr" Then

    'mGLAccNAVCodeNNNR = GetGLAccNAVCode("RRN Withdrawal Taxable")

    tmp = "Select TPI.BOEID as InvoiceNo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode  'ClientNAVCode',TPI.Amount,TPI.BatchID,MST.NAVCode ServiceTaxNAVCode,TPI.ReverseDate,TPI.ReverseRemark"
    tmp = tmp & " ,TPI.InvoiceDate,isnull(TPI.Discount,0)Discount, isnull(TPI.Premium,0)Premium, TPI.FAGPostingDate,MST.InvoiceType,MComm.HSNNo"
    tmp = tmp & " From Txn_CDTFBOE TPI"
    tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
    tmp = tmp & " Inner Join Txn_RRNPWR TC On TPI.CDTFNo = TC.ConvertedID And TPI.TxnType = 'P' "
    tmp = tmp & " inner join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
    tmp = tmp & " inner join Mst_ServiceTax MST On MComm.ServiceTaxID = MST.ServiceTaxID "
    tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
    tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"
    tmp = tmp & " Where ReverseFlag='Y' And RBatchID=" & mRBatchID & " and TPI.TxnType='P'"
    tmp = tmp & " Order By TPI.InvoiceDate"

    mIncomeTypeID = 13

    mGLAccNAVCodeAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "G")
    mGLAccNAVCodeNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "G")
    mHSNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "H")
    mHSNNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "H")

Else

    'mGLAccNAVCodeNNNR = GetGLAccNAVCode("Rejected Taxable")

    tmp = "Select TPI.BOEID as InvoiceNo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,MST.NAVCode ServiceTaxNAVCode,TPI.ReverseDate,TPI.ReverseRemark"
    tmp = tmp & " ,TPI.InvoiceDate,isnull(TPI.Discount,0)Discount, isnull(TPI.Premium,0)Premium, TPI.FAGPostingDate,MST.InvoiceType,MComm.HSNNo"
    tmp = tmp & " From Txn_CDTFBOE TPI"
    tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
    tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType='W' "
    tmp = tmp & " inner join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
    tmp = tmp & " inner join Mst_ServiceTax MST On MComm.ServiceTaxID = MST.ServiceTaxID "
    tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
    tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"
    tmp = tmp & " Where ReverseFlag='Y' And RBatchID=" & mRBatchID & " and TPI.TxnType='W'"
    tmp = tmp & " Order By TPI.InvoiceDate"

    mIncomeTypeID = 2

    mGLAccNAVCodeAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "G")
    mGLAccNAVCodeNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "G")
    mHSNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "H")
    mHSNNAGRI = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "H")

End If

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For r = 1 To TmpRs.RecordCount
       I = MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.TextMatrix(I, 0) = Format(TmpRs!ReverseDate, "mm/dd/yyyy") ', "dd/mm/yyyy")
       MSHFlexGrid3.TextMatrix(I, 1) = "Credit Memo"
       MSHFlexGrid3.TextMatrix(I, 2) = "Customer"
       MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs!ReverseRemark), "", TmpRs!ReverseRemark)
       MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
       MSHFlexGrid3.TextMatrix(I, 17) = "G/L Account"
       MSHFlexGrid3.TextMatrix(I, 19) = TmpRs!ServiceTaxNAVCode
       MSHFlexGrid3.TextMatrix(I, 21) = TmpRs!InvoiceType
       MSHFlexGrid3.TextMatrix(I, 20) = IIf(IsNull(TmpRs!HSNNo), "", TmpRs!HSNNo)
       If MSHFlexGrid3.TextMatrix(I, 19) = "GSTS18" Then
          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeNAGRI
          If MSHFlexGrid3.TextMatrix(I, 20) = "" Then
             MSHFlexGrid3.TextMatrix(I, 20) = mHSNNAGRI
          End If
       Else
          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeAGRI
          If MSHFlexGrid3.TextMatrix(I, 20) = "" Then
             MSHFlexGrid3.TextMatrix(I, 20) = mHSNAGRI
          End If
       End If
       
       MSHFlexGrid3.TextMatrix(I, 4) = (TmpRs!Amount + TmpRs!Premium - TmpRs!Discount) '-(TmpRs!Amount + TmpRs!Premium - TmpRs!Discount)
       'MSHFlexGrid3.TextMatrix(I, 5) = TmpRs!InvoiceNo
       MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(TmpRs!LocatioNAVCode), "", TmpRs!LocatioNAVCode)
       MSHFlexGrid3.TextMatrix(I, 8) = "CMP"
       'MSHFlexGrid3.TextMatrix(I, 9) = "EMP-1000"
       'MSHFlexGrid3.TextMatrix(I, 15) = ""
       If LCase(CmbInvoiceType.Text) = "nnnr" Then
            MSHFlexGrid3.TextMatrix(I, 13) = "RNN" & TmpRs!InvoiceNo
            tmp = "Select EntryNo,NAVCode from Txn_TIDetails Where  PIFlag='NNNR' And PIID=" & TmpRs!InvoiceNo & ""
            Call Tmp3Table
            If TmpRs3.RecordCount > 0 Then
               MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs3!EntryNo), "", TmpRs3!EntryNo)
               MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
            End If
       ElseIf LCase(CmbInvoiceType.Text) = "7 days" Then
            MSHFlexGrid3.TextMatrix(I, 13) = "RSD" & TmpRs!InvoiceNo
            tmp = "Select EntryNo,NAVCode from Txn_TIDetails Where  PIFlag='SD' And PIID=" & TmpRs!InvoiceNo & ""
            Call Tmp3Table
            If TmpRs3.RecordCount > 0 Then
               MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs3!EntryNo), "", TmpRs3!EntryNo)
               MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
            End If
       ElseIf LCase(CmbInvoiceType.Text) = "rrn" Then
            If Val(MSHFlexGrid3.TextMatrix(I, 4)) > Val(TmpRs!NCDEXSpaceCharges) Then
               MSHFlexGrid3.TextMatrix(I, 4) = Val(MSHFlexGrid3.TextMatrix(I, 4)) - Val(TmpRs!NCDEXSpaceCharges)
            End If
            MSHFlexGrid3.TextMatrix(I, 13) = "RRRN" & TmpRs!InvoiceNo
            tmp = "Select EntryNo,NAVCode from Txn_TIDetails Where  PIFlag='RRN' And PIID=" & TmpRs!InvoiceNo & ""
            Call Tmp3Table
            If TmpRs3.RecordCount > 0 Then
               MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs3!EntryNo), "", TmpRs3!EntryNo)
               MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
            End If
       ElseIf LCase(CmbInvoiceType.Text) = "nnnry" Then
            MSHFlexGrid3.TextMatrix(I, 13) = "RNNNRY" & TmpRs!InvoiceNo
            tmp = "Select EntryNo,NAVCode from Txn_TIDetails Where  PIFlag='NNNRY' And PIID=" & TmpRs!InvoiceNo & ""
            Call Tmp3Table
            If TmpRs3.RecordCount > 0 Then
               MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs3!EntryNo), "", TmpRs3!EntryNo)
               MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
            End If
       ElseIf LCase(CmbInvoiceType.Text) = "pwr" Then
            MSHFlexGrid3.TextMatrix(I, 13) = "RPWR" & TmpRs!InvoiceNo
            tmp = "Select EntryNo,NAVCode from Txn_TIDetails Where  PIFlag='PWR' And PIID=" & TmpRs!InvoiceNo & ""
            Call Tmp3Table
            If TmpRs3.RecordCount > 0 Then
               MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs3!EntryNo), "", TmpRs3!EntryNo)
               MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
            End If
       Else
            MSHFlexGrid3.TextMatrix(I, 13) = "RRJ" & TmpRs!InvoiceNo
            tmp = "Select EntryNo,NAVCode from Txn_TIDetails Where  PIFlag='RJ' And PIID=" & TmpRs!InvoiceNo & ""
            Call Tmp3Table
            If TmpRs3.RecordCount > 0 Then
               MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs3!EntryNo), "", TmpRs3!EntryNo)
               MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
            End If
       End If
       
       If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
       If MSHFlexGrid3.TextMatrix(I, 16) = "" Then ExpFlag = "N"
       If MSHFlexGrid3.TextMatrix(I, 18) = "" Then ExpFlag = "N"
       If MSHFlexGrid3.TextMatrix(I, 20) = "" Then ExpFlag = "N"
              
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       TmpRs.MoveNext
   Next
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
CmbInvoiceType.Text = ""
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
    .Cols = 22 '19
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
Call Grid_Head
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "_ExportReverseNNNRInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_ExportReverseNNNRInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
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
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
