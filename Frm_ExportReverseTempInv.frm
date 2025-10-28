VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportReverseTempInv 
   Caption         =   "Export Reverse Ad-Hoc Invoice"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   10875
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
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
         Left            =   6840
         TabIndex        =   3
         Top             =   240
         Width           =   2880
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
         Left            =   3360
         TabIndex        =   2
         Top             =   240
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
         Left            =   1440
         TabIndex        =   1
         Top             =   270
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
         Left            =   240
         TabIndex        =   0
         Top             =   255
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   4
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label2 
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
         Left            =   5520
         TabIndex        =   12
         Top             =   293
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
         Left            =   2400
         TabIndex        =   11
         Top             =   293
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
      TabIndex        =   8
      Top             =   840
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7680
         Left            =   120
         TabIndex        =   9
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
Attribute VB_Name = "Frm_ExportReverseTempInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCodeTDS, mGLAccNAVCodeTDSDP, mGLAccNAVCodeOther, mGLAccNAVCodeLoadUnload, mGLAccNAVCodeCRC, mGLAccNAVCodeSCMIncome, mGLAccNAVCodeCommtrack, mGLAccNAVCodeRRN, mGLAccNAVCodeWBIncome As Variant
Dim ExpFlag, MaxID As Variant
Dim mJournalBatchName, mJournalType As Variant
Dim tmp1, mIncomeTypeID, mGLAccNAVCodeNAGRI, mGLAccNAVCodeAGRI, mHSNNAGRI, mHSNAGRI, mInvType As Variant

Private Sub CmbBatch_GotFocus()
tmp = "select distinct RBatchID As BatchID from Txn_TempInvoice Where RBatchID > 0  order by RBatchID"
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
'If Option1.Value = True Then
'If CmbJournalCode.Text = "" Then
'   MsgBox "Journal Code not selected"
'   CmbJournalCode.SetFocus
'   Exit Sub
'End If
''mJournalType = "SALES"
'mUserBatchID = GetGUID
''mJournalBatchName = GetJournalDetailsFromHRPro(Gen_UserLoginID, mJournalType)
'If Trim(mJournalType) = "" Or Trim(mJournalBatchName) = "" Then
'   MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'   Exit Sub
'End If
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
MaxID = GetMaxBatchID
If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_ReverseTempInvoice" & "_" & MaxID & ".csv"
Else
   strFilName = "\Navision\NavUpload_ReverseTempInvoice" & "_" & CmbBatch.Text & ".csv"
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
        RsTxn_InvoicesForNAVISION!InvoiceType = "ADHOC"
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
   tmp = "Update Txn_TempInvoice set RBatchID = " & MaxID & " Where ReverseFlag='Y' And RBatchID=0 "
   Call TmpTable
End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
End Sub
Private Function GetMaxBatchID()
tmp = "select Max(RBatchID) from Txn_TempInvoice"
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

tmp = " Select TPI.PINo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode 'ClientNAVCode',MCSG.RGSTNavCode as 'RClientNAVCode',(TPI.Amount) Amount,TPI.BatchID,MST.NAVCode, TPI.ReverseDate,TPI.ReverseRemark,TPI.InvoiceDate,TPI.IETypeID,MDSG.GSTNavCode as 'DPNAVCode',MST.InvoiceType"
tmp = tmp & " From Txn_TempInvoice TPI"
tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_Client MC On TPI.ClientID = MC.ClientID"
tmp = tmp & " Left Join Mst_Depository MD On TPI.DPID = MD.DepositoryID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Left Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"
tmp = tmp & " Left Join Mst_DepositoryStateWiseGST MDSG on MD.DepositoryID = MDSG.DepositoryID And MSI.StateIndiaID = MDSG.StateIndiaID"
tmp = tmp & " Inner Join Mst_ServiceTax MST on MST.ServiceTaxID=TPI.ServiceTaxID"

tmp = tmp & " Where ReverseFlag='Y' And RBatchID=" & mRBatchID
tmp = tmp & " Order By TPI.PIDate"

Call TmpTable
'
'mGLAccNAVCodeResr = GetGLAccNAVCode("Reservation Taxable")
'mGLAccNAVCodeSevDays = GetGLAccNAVCode("7 Days Taxable")
'mGLAccNAVCodeNNNR = GetGLAccNAVCode("NNNC Taxable")
'mGLAccNAVCodeRej = GetGLAccNAVCode("Rejected Taxable")
'mGLAccNAVCodeSpot = GetGLAccNAVCode("SPOT Taxable")
'mGLAccNAVCodeSpil = GetGLAccNAVCode("Spillage Taxable")
'mGLAccNAVCodeDPClient = GetGLAccNAVCode("DP")
'mGLAccNAVCodeCRC = GetGLAccNAVCode("Cheque Return Charges Taxable")
'mGLAccNAVCodeLoadUnload = GetGLAccNAVCode("Loading unloading charges reimbursement from client Taxable")
'mGLAccNAVCodeSCMIncome = GetGLAccNAVCode("SCM Income Taxable")
'mGLAccNAVCodeOther = GetGLAccNAVCode("Other Income Taxable")
'mGLAccNAVCodeCommtrack = GetGLAccNAVCode("Commtrack Taxable")
'mGLAccNAVCodeRRN = GetGLAccNAVCode("RRN Withdrawal Taxable")
'mGLAccNAVCodeWBIncome = GetGLAccNAVCode("Weighbridge Income Taxable")
'mGLAccNAVCodeSOC = GetGLAccNAVCode("Space Occupancy charges Taxable")

If TmpRs.RecordCount > 0 Then
   For r = 1 To TmpRs.RecordCount
       I = MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.TextMatrix(I, 0) = Format(TmpRs!ReverseDate, "mm/dd/yyyy") ', "dd/mm/yyyy")
       MSHFlexGrid3.TextMatrix(I, 1) = "Credit Memo"
       MSHFlexGrid3.TextMatrix(I, 2) = "Customer"
       MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs!ReverseRemark), "", TmpRs!ReverseRemark)
       If TmpRs!IETypeID <> 8 And TmpRs!IETypeID <> 12 Then
          MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
       Else
          MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!DPNAVCode), "", TmpRs!DPNAVCode)
       End If
       
       MSHFlexGrid3.TextMatrix(I, 4) = TmpRs!Amount
       MSHFlexGrid3.TextMatrix(I, 5) = "G/L Account"
       MSHFlexGrid3.TextMatrix(I, 19) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
       MSHFlexGrid3.TextMatrix(I, 20) = "" 'IIf(IsNull(TmpRs!HSNNo), "", TmpRs!HSNNo)
       MSHFlexGrid3.TextMatrix(I, 21) = IIf(IsNull(TmpRs!InvoiceType), "", TmpRs!InvoiceType)
       If TmpRs!NavCode = "GSTS00" Then
          MSHFlexGrid3.TextMatrix(I, 18) = GetGLAccNAVCodeFromInvoiceType(TmpRs!IETypeID, "AGRI", "G")
          If MSHFlexGrid3.TextMatrix(I, 20) = "" Then
             MSHFlexGrid3.TextMatrix(I, 20) = GetGLAccNAVCodeFromInvoiceType(TmpRs!IETypeID, "AGRI", "H")
          End If
       Else
          MSHFlexGrid3.TextMatrix(I, 18) = GetGLAccNAVCodeFromInvoiceType(TmpRs!IETypeID, "NAGRI", "G")
          If MSHFlexGrid3.TextMatrix(I, 20) = "" Then
             MSHFlexGrid3.TextMatrix(I, 20) = GetGLAccNAVCodeFromInvoiceType(TmpRs!IETypeID, "NAGRI", "H")
          End If
       End If
       
       MSHFlexGrid3.TextMatrix(I, 17) = "G/L Account"
'       If TmpRs!IETypeID = 1 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeSevDays
'       ElseIf TmpRs!IETypeID = 2 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeRej
'       ElseIf TmpRs!IETypeID = 3 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeSpot
'       ElseIf TmpRs!IETypeID = 4 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeResr
'          MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!RClientNAVCode), "", TmpRs!RClientNAVCode)
'       ElseIf TmpRs!IETypeID = 5 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeOther
'       ElseIf TmpRs!IETypeID = 6 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeSpil
'       ElseIf TmpRs!IETypeID = 7 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeNNNR
'       ElseIf TmpRs!IETypeID = 8 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeDPClient
'       ElseIf TmpRs!IETypeID = 9 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeCRC
'       ElseIf TmpRs!IETypeID = 10 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeLoadUnload
'       ElseIf TmpRs!IETypeID = 11 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeSCMIncome
'       ElseIf TmpRs!IETypeID = 12 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeCommtrack
'       ElseIf TmpRs!IETypeID = 13 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeRRN
'       ElseIf TmpRs!IETypeID = 14 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeWBIncome
'       ElseIf TmpRs!IETypeID = 15 Then
'          MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCodeSOC
'       End If
       

       
       MSHFlexGrid3.TextMatrix(I, 4) = TmpRs!Amount
       'MSHFlexGrid3.TextMatrix(I, 5) = TmpRs!InvoiceNo
       MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(TmpRs!LocatioNAVCode), "", TmpRs!LocatioNAVCode)
       MSHFlexGrid3.TextMatrix(I, 8) = "CMP"
      
       MSHFlexGrid3.TextMatrix(I, 13) = "RAH" & TmpRs!PINo
       tmp = "Select EntryNo,NAVCode from Txn_TIDetails Where  PIFlag='ADHOC' And PIID=" & TmpRs!PINo & ""
       Call Tmp3Table
       If TmpRs3.RecordCount > 0 Then
          MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs3!EntryNo), "", TmpRs3!EntryNo)
          MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
       End If
       If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
       If MSHFlexGrid3.TextMatrix(I, 18) = "" Then ExpFlag = "N"
       If MSHFlexGrid3.TextMatrix(I, 7) = "" Then ExpFlag = "N"
       
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
CmbJournalCode.Text = ""
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
    .TextMatrix(1, 21) = "Number Series Flag"
     
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

