VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportTDSAdjInvoice 
   Caption         =   "TDS Adj Invoice"
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
      TabIndex        =   9
      Top             =   120
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
         Left            =   5040
         TabIndex        =   1
         Top             =   240
         Width           =   2640
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
         Left            =   11040
         TabIndex        =   4
         Top             =   240
         Width           =   1260
      End
      Begin VB.ComboBox CmbBatch 
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
         TabIndex        =   0
         Top             =   235
         Width           =   2040
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
         Left            =   9780
         TabIndex        =   3
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
         Left            =   12315
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
         Left            =   13575
         TabIndex        =   6
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
         Left            =   8520
         TabIndex        =   2
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label3 
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
         Left            =   3720
         TabIndex        =   11
         Top             =   300
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
         Left            =   240
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
      TabIndex        =   8
      Top             =   960
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7560
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13335
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "Frm_ExportTDSAdjInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCodeTDS, mGLAccNAVCodeTDSDP, mGLAccNAVCodeOther As Variant
Dim ExpFlag, MaxID As Variant
Dim mJournalBatchName, mJournalType As Variant
Dim mDateLock, mBatchDate As Variant

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
'If CmbBatch.Text = "" Then
'   Exit Sub
'End If
'TmpRs1.MoveFirst
'TmpRs1.Find "BatchID = " & Val(CmbBatch.Text) & ""
'If TmpRs1.EOF Then
'   MsgBox "Invalid selection "
'   CmbBatch.SetFocus
'   Exit Sub
'End If

mBatchDate = ""
If CmbBatch.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "BatchID = " & Val(CmbBatch.Text) & ""
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbBatch.SetFocus
   Exit Sub
Else
    tmp = " SELECT Distinct MAX(AdjDate)[AdjDate] FROM Txn_CRTDSAdjustment Where BatchID =  " & Val(CmbBatch.Text) & " "
    Call Tmp4Table
    If TmpRs4.RecordCount > 0 Then
        mBatchDate = IIf(IsNull(TmpRs4!AdjDate), "", Format(TmpRs4!AdjDate, "YYYY-MM-DD"))
    End If
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
'If Option1.Value = True Then
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
'End If
mInvoiceType = "CRADJ"

Path = App.Path
MaxID = GetMaxBatchID
strFilName = "\Navision\NavUpload_TDSAdjInvoices" & "_" & CmbBatch.Text & ".csv"
''-- On 05 Aug 2016
'If Option2.Value = True Then
'    strFilLocation = Path + strFilName
'    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
'    Open strFilLocation For Output As #1
'Else
   Call TableTxn_InvoiceAllocationForNAVISION(" Where G_UID Is Null ")
'End If

With MSHFlexGrid3
For I = 2 To MSHFlexGrid3.Rows - 1
    tmp = ""
    If Trim(.TextMatrix(I, 0)) = "" Then Exit For
    RsTxn_InvoiceAllocationForNAVISION.AddNew
    RsTxn_InvoiceAllocationForNAVISION!PaymentEntryNo = Trim(.TextMatrix(I, 0))
    RsTxn_InvoiceAllocationForNAVISION!InvoiceEntryNo = Trim(.TextMatrix(I, 1))
    RsTxn_InvoiceAllocationForNAVISION!Amount = Trim(.TextMatrix(I, 2))
    RsTxn_InvoiceAllocationForNAVISION!CreatedBy = Gen_UserLoginID
    RsTxn_InvoiceAllocationForNAVISION!G_UID = GetGUID
    RsTxn_InvoiceAllocationForNAVISION!BatchID = CmbBatch.Text
    'RsTxn_InvoiceAllocationForNAVISION!DataTransfer = Null
    RsTxn_InvoiceAllocationForNAVISION!InvoiceType = mInvoiceType
    RsTxn_InvoiceAllocationForNAVISION!JournalBatchType = mJournalType
    RsTxn_InvoiceAllocationForNAVISION!JournalBatchName = mJournalBatchName
    'RsTxn_InvoiceAllocationForNAVISION!EntryNoTransfer = Null
    'RsTxn_InvoiceAllocationForNAVISION!EntryNoTransferRemark = Null
    RsTxn_InvoiceAllocationForNAVISION!UserBatchID = mUserBatchID
    RsTxn_InvoiceAllocationForNAVISION.Update
    
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

'MsgBox "Done !!!"
MsgBox " Batch No - " & CmbBatch.Text & " is Exported."
Close #1

End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant
Dim mBatchID As Double
ExpFlag = ""
If CmbBatch.Text = "" Then
   MsgBox "Select Batch ID !!!! "
   Exit Sub
End If
mBatchID = Val(CmbBatch.Text)

Call Grid_Head

''mGLAccNAVCodeTDS = GetGLAccNAVCode("TDSNavCode")
''mGLAccNAVCodeTDSDP = GetGLAccNAVCode("TDSNavCodeDP")
''mGLAccNAVCodeOther = GetGLAccNAVCode("OtherDedNavCode")

tmp = "Select TCRT.TDSEntryNo EntryNo,TI.EntryNo as InvEntryNo,TCRT.RecTDS as Amount"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
tmp = tmp & " Where TCR.BatchID = " & mBatchID & " And TCRT.RecTDS > 0"
tmp = tmp & " Union All"
tmp = tmp & " Select TCRT.OtherDedEntryNo EntryNo,TI.EntryNo as InvEntryNo,TCRT.RecOtherDed as Amount"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR  on TCRT.CRTDSAdjustmentID=TCR.CRTDSAdjustmentID And TCRT.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Inner Join Mst_ExpenseType ME  on ME.ExpenseTypeID=TCRT.ExpenseAccID"
tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
tmp = tmp & " Where TCR.BatchID = " & mBatchID & " And TCRT.RecOtherDed > 0 And OtherDedEntryNo > 0 And ME.ReverseFlag='N'"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For r = 1 To TmpRs.RecordCount
       I = MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.TextMatrix(I, 0) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
       MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs!InvEntryNo), "", TmpRs!InvEntryNo)
       MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
       If MSHFlexGrid3.TextMatrix(I, 0) = "" Then ExpFlag = "N"
       If MSHFlexGrid3.TextMatrix(I, 1) = "" Then ExpFlag = "N"
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

''-- Chk that Batch ID already Exported / Not
If MSHFlexGrid3.Rows > 3 Then
    If mBatchDate < mDateLock Then
        CmdExport.Enabled = False
        MsgBox "Given Batch No is already Exported." & vbNewLine & "You can Export in Excel."
        Exit Sub
    End If
    tmp = " SELECT Distinct BatchID FROM Txn_InvoiceAllocationForNAVISION Where InvoiceTYpe = 'CRADJ' And BatchID = " & Val(CmbBatch.Text) & " "
    Call Tmp4Table
    If TmpRs4.RecordCount > 0 Then
        CmdExport.Enabled = False
        MsgBox "Given Batch No is already Exported." & vbNewLine & "You can Export in Excel."
        Exit Sub
    End If
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
mDateLock = Format("2016-04-01", "YYYY-MM-DD")
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

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportTDSAdjInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportTDSAdjInvoice.xls")
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


