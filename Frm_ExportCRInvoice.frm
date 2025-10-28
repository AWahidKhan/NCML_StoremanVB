VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportCRInvoice 
   Caption         =   "Export Cash Receipt Invoices"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   9690
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
      TabIndex        =   7
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7320
         Left            =   120
         TabIndex        =   8
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
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   12360
         TabIndex        =   13
         Top             =   600
         Width           =   1260
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
         Left            =   6120
         TabIndex        =   11
         Top             =   240
         Width           =   3600
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
         TabIndex        =   10
         Top             =   120
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
         TabIndex        =   9
         Top             =   480
         Width           =   855
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         Left            =   2640
         TabIndex        =   1
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
         Left            =   4800
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
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   300
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_ExportCRInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCodeTDS, mGLAccNAVCodeTDSDP, mGLAccNAVCodeOther As Variant
Dim ExpFlag, MaxID As Variant
Dim mBatchID As Double
Dim mJournalBatchName, mJournalType As Variant

Private Sub CmbBatch_GotFocus()
If Option1.Value = True Then
   tmp = "select distinct BatchID As BatchID from Txn_NewCashReceipt Where BatchID > 0 And isnull(InvAllocFlag,0)=0 order by BatchID"
Else
   tmp = "select distinct BatchID As BatchID from Txn_NewCashReceipt Where BatchID > 0 and InvAllocFlag=1 order by BatchID"
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportCRInvoiceAllocation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportCRInvoiceAllocation.xls")
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
strFilName = "\Navision\NavUpload_CashReceiptInvoices" & "_" & CmbBatch.Text & ".csv"

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
    RsTxn_InvoiceAllocationForNAVISION!BatchID = CmbBatch.Text  ''IIf(Option1.Value = True, MaxID, CmbBatch.Text)
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
   tmp = "Update Txn_NewCashReceipt Set InvAllocFlag=1 Where BatchID=" & Val(CmbBatch.Text) & ""
   Call TmpTable
End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & CmbBatch.Text & " is Exported."
Call Grid_Head
Close #1
End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant

ExpFlag = ""
If CmbBatch.Text = "" Then
   MsgBox "Select Batch No !!!! "
   Exit Sub
End If
mBatchID = Val(CmbBatch.Text)

Call Grid_Head

''mGLAccNAVCodeTDS = GetGLAccNAVCode("TDSNavCode")
''mGLAccNAVCodeTDSDP = GetGLAccNAVCode("TDSNavCodeDP")
''mGLAccNAVCodeOther = GetGLAccNAVCode("OtherDedNavCode")

tmp = "Select TCRT.EntryNo,TI.EntryNo as InvEntryNo,TCRT.RecAmount+TCRT.RecTDS as Amount"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR  on TCRT.NewCRID=TCR.NewCRID And TCRT.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Inner Join Txn_TIDetails TI  on TCRT.TIID=TI.TIID"
tmp = tmp & " Where TCR.BatchID = " & mBatchID & " And isnull(TCRT.CRAAID,0)=0"
tmp = tmp & " Order By TCRT.EntryNo,TI.EntryNo"

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

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmbBatch.Text = ""
CmdExport.Enabled = False
CmbJournalCode.Text = ""
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

