VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportCRAdvInvAlloc 
   Caption         =   "Export CR Advance Invoice Allocation"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   12
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
         Left            =   3720
         TabIndex        =   4
         Top             =   690
         Width           =   3600
      End
      Begin VB.CheckBox Check1 
         Caption         =   "All"
         Height          =   255
         Left            =   1470
         TabIndex        =   2
         Top             =   278
         Width           =   735
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
         Top             =   720
         Width           =   1215
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
         Left            =   3720
         TabIndex        =   3
         Top             =   225
         Width           =   3600
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
         Left            =   9900
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
         Left            =   11160
         TabIndex        =   7
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
         Left            =   12420
         TabIndex        =   8
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
         Left            =   8640
         TabIndex        =   5
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
         Left            =   2400
         TabIndex        =   14
         Top             =   743
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
         Left            =   2400
         TabIndex        =   13
         Top             =   278
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
      TabIndex        =   10
      Top             =   1200
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7680
         Left            =   120
         TabIndex        =   11
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
Attribute VB_Name = "Frm_ExportCRAdvInvAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCodeTDS, mGLAccNAVCodeTDSDP, mGLAccNAVCodeOther As Variant
Dim ExpFlag, MaxID As Variant
Dim mCRAAIDs As String
Dim mJournalBatchName, mJournalType As Variant

Private Sub CmbBatch_GotFocus()
tmp = "select distinct BatchID As BatchID from Txn_CRAdvanceAllocation Where BatchID > 0  order by BatchID"
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
mInvoiceType = "CRA"
 
Path = App.Path
MaxID = GetMaxBatchID
If CmbBatch.Text <> "" Then
   strFilName = "\Navision\NavUpload_CRAdvInvoices" & "_" & CmbBatch.Text & ".csv"
Else
   strFilName = "\Navision\NavUpload_CRAdvInvoices" & "_" & MaxID & ".csv"
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
   Call TableTxn_InvoiceAllocationForNAVISION(" Where G_UID Is Null ")
End If

With MSHFlexGrid3
For I = 2 To MSHFlexGrid3.Rows - 2
    tmp = ""
    If Option1.Value = True Then
    
    RsTxn_InvoiceAllocationForNAVISION.AddNew
    RsTxn_InvoiceAllocationForNAVISION!PaymentEntryNo = Trim(.TextMatrix(I, 0))
    RsTxn_InvoiceAllocationForNAVISION!InvoiceEntryNo = Trim(.TextMatrix(I, 1))
    RsTxn_InvoiceAllocationForNAVISION!Amount = Trim(.TextMatrix(I, 2))
    RsTxn_InvoiceAllocationForNAVISION!CreatedBy = Gen_UserLoginID
    RsTxn_InvoiceAllocationForNAVISION!G_UID = GetGUID
    RsTxn_InvoiceAllocationForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
    'RsTxn_InvoiceAllocationForNAVISION!DataTransfer = Null
    RsTxn_InvoiceAllocationForNAVISION!InvoiceType = mInvoiceType
    RsTxn_InvoiceAllocationForNAVISION!JournalBatchType = mJournalType
    RsTxn_InvoiceAllocationForNAVISION!JournalBatchName = mJournalBatchName
    'RsTxn_InvoiceAllocationForNAVISION!EntryNoTransfer = Null
    'RsTxn_InvoiceAllocationForNAVISION!EntryNoTransferRemark = Null
    RsTxn_InvoiceAllocationForNAVISION!UserBatchID = mUserBatchID
    RsTxn_InvoiceAllocationForNAVISION.Update
    
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
   If mCRAAIDs <> "" Then
       tmp = "Update Txn_CRAdvanceAllocation set BatchID = " & MaxID & " Where CRAAID in (" & mCRAAIDs & ")" 'BatchID=0 "
       Call TmpTable
   End If
End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant
Dim mBatchID As Double
ExpFlag = ""
If Option1.Value = True Then
   mBatchID = 0
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
   mBatchID = Val(CmbBatch.Text)
End If

Call Grid_Head

''mGLAccNAVCodeTDS = GetGLAccNAVCode("TDSNavCode")
''mGLAccNAVCodeTDSDP = GetGLAccNAVCode("TDSNavCodeDP")
''mGLAccNAVCodeOther = GetGLAccNAVCode("OtherDedNavCode")

tmp = "Select TCRT.EntryNo,TI.EntryNo as InvEntryNo,TCRT.RecAmount as Amount,TI.NAVCode,TCR.NewCRID,TCR.SM_Prefix,CRA.CRAAID"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR  on TCRT.NewCRID=TCR.NewCRID And TCRT.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
tmp = tmp & " Inner Join Txn_CRAdvanceAllocation CRA  on TCRT.CRAAID=CRA.CRAAID"
tmp = tmp & " Where TCR.AdvanceFlag=1 and CRA.BatchID = " & mBatchID
tmp = tmp & " Order By TCR.CRID,TCR.SM_Prefix,TI.EntryNo"

Call TmpTable

mCRAAIDs = ""
If TmpRs.RecordCount > 0 Then
   For r = 1 To TmpRs.RecordCount
       tmp = "Select Amount,EntryNo,AdvNAVCode"
       tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRD"
       tmp = tmp & " left join Mst_Client MC on MC.ClientID=TCRD.ClientID"
       tmp = tmp & " Left join Mst_Depository MD on MD.DepositoryID=TCRD.DPID And MD.ExchangeTypeID = TCRD.ExchangeTypeID"
       tmp = tmp & " Where  TCRD.NewCRID = " & TmpRs!NewCRID & " and TCRD.SM_Prefix= '" & TmpRs!SM_Prefix & "' And (TCRD.AdvNAVCode='" & TmpRs!NavCode & "' Or MC.NAVCode='" & TmpRs!NavCode & "' or MC.RNAVCode='" & TmpRs!NavCode & "' or MD.NAVCode='" & TmpRs!NavCode & "')"
        
       Call Tmp6Table

       If TmpRs6.RecordCount > 0 Then
          If (Check1.Visible = True And Check1.Value = 0) Or Option2.Value = True Then
             If IIf(IsNull(TmpRs6!EntryNo), "", TmpRs6!EntryNo) <> "" Then
                I = MSHFlexGrid3.Rows - 1
                MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs!InvEntryNo), "", TmpRs!InvEntryNo)
                MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
                MSHFlexGrid3.TextMatrix(I, 0) = IIf(IsNull(TmpRs6!EntryNo), "", TmpRs6!EntryNo)
                MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
                If mCRAAIDs = "" Then
                   mCRAAIDs = TmpRs!CRAAID
                Else
                   mCRAAIDs = mCRAAIDs & "," & TmpRs!CRAAID
                End If
                If MSHFlexGrid3.TextMatrix(I, 0) = "" Then ExpFlag = "N"
                If MSHFlexGrid3.TextMatrix(I, 1) = "" Then ExpFlag = "N"
             End If
          Else
                I = MSHFlexGrid3.Rows - 1
       
                MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs!InvEntryNo), "", TmpRs!InvEntryNo)
                MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
                MSHFlexGrid3.TextMatrix(I, 0) = IIf(IsNull(TmpRs6!EntryNo), "", TmpRs6!EntryNo)
                MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
                If MSHFlexGrid3.TextMatrix(I, 0) = "" Then ExpFlag = "N"
                If MSHFlexGrid3.TextMatrix(I, 1) = "" Then ExpFlag = "N"
          End If
       Else
          If Check1.Visible = True And Check1.Value = 1 Then
             I = MSHFlexGrid3.Rows - 1
        
             MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs!InvEntryNo), "", TmpRs!InvEntryNo)
             MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
             MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
          End If
       End If
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
    .Cols = 3
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(1, 0) = "Payment Entry No"
    .TextMatrix(1, 1) = "Invoice Entry No"
    .TextMatrix(1, 2) = "Amount"
    .ColWidth(0) = 1600
    .ColWidth(1) = 1600
    .ColWidth(2) = 1500
End With
End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   CmbBatch.Enabled = True
   Check1.Visible = False
   Check1.Value = 0
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   Check1.Visible = True
   Check1.Value = 0
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub
Private Function GetMaxBatchID()
tmp = "select Max(BatchID) from Txn_CRAdvanceAllocation"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   GetMaxBatchID = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0)) + 1
Else
   GetMaxBatchID = 1
End If
End Function

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportCashReceiptAdvInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportCashReceiptAdvInvoice.xls")
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

