VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_NNNRMismatchInvoice 
   Caption         =   "NNNR Mismatch Invoice"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   10005
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   14970
      Begin VB.ComboBox CmbState 
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
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   2760
      End
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmRpt_NNNRMismatchInvoice.frx":0000
         Left            =   8925
         List            =   "FrmRpt_NNNRMismatchInvoice.frx":000D
         TabIndex        =   5
         Top             =   540
         Width           =   1965
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
         Height          =   360
         Left            =   12600
         TabIndex        =   9
         Top             =   1035
         Width           =   1125
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
         Height          =   360
         Left            =   11475
         TabIndex        =   8
         Top             =   1035
         Width           =   1125
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
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
         Left            =   13725
         TabIndex        =   10
         Top             =   1035
         Width           =   1125
      End
      Begin VB.ComboBox CmbLocation 
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
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   2880
      End
      Begin VB.ComboBox CmbNonNcdexClient 
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
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   9345
      End
      Begin VB.CheckBox ChkFAG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   13080
         TabIndex        =   7
         Top             =   210
         Width           =   225
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7935
         Left            =   90
         TabIndex        =   11
         Top             =   1545
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13996
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Format          =   42991617
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   1575
         TabIndex        =   2
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Format          =   42991617
         CurrentDate     =   39884
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   90
         TabIndex        =   20
         Top             =   9600
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4148
         TabIndex        =   21
         Top             =   210
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9570
         TabIndex        =   19
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1875
         TabIndex        =   18
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6968
         TabIndex        =   17
         Top             =   210
         Width           =   765
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Client Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   15
         Top             =   210
         Width           =   945
      End
      Begin VB.Label LblTotalRecord 
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
         Left            =   13590
         TabIndex        =   14
         Top             =   615
         Width           =   90
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   11190
         TabIndex        =   13
         Top             =   615
         Width           =   2280
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Approved By FAG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11160
         TabIndex        =   12
         Top             =   210
         Width           =   1635
      End
   End
End
Attribute VB_Name = "FrmRpt_NNNRMismatchInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLOcationID, mClientID As Variant

Private Sub Grid_Head()
CmdExcel.Enabled = False
LblTotalRecord.Caption = "--"
ProgressBar1.Value = 0
With MSHFlexGrid1
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 25
   
   .TextMatrix(1, 0) = "Invoice No"
   .TextMatrix(1, 1) = "Invoice Date"
   .TextMatrix(1, 2) = "Billing Month"
   .TextMatrix(1, 3) = "State Name"
   .TextMatrix(1, 4) = "Location Name"
   .TextMatrix(1, 5) = "Exchange Type"
   .TextMatrix(1, 6) = "Client Name"
   .TextMatrix(1, 7) = "GSL Invoice Amount"
   .TextMatrix(1, 8) = "Amount"
   .TextMatrix(1, 9) = "Discount"
   .TextMatrix(1, 10) = "Premium"
   .TextMatrix(1, 11) = "BatchID"
   .TextMatrix(1, 12) = "GST Group Code" '"Service Tax NAVCode"
   .TextMatrix(1, 13) = "FAG Posting Date"
   .TextMatrix(1, 14) = "Flag"
   .TextMatrix(1, 15) = "TIID"
   .TextMatrix(1, 16) = "TINo"
   .TextMatrix(1, 17) = "TI Date"
   .TextMatrix(1, 18) = "TI Amount"
   .TextMatrix(1, 19) = "TIST Amount"
   .TextMatrix(1, 20) = "Entry No"
   .TextMatrix(1, 21) = "Reverse Flag"
   .TextMatrix(1, 22) = "Reverse Remarks"
   .TextMatrix(1, 23) = "Reverse Date"
   .TextMatrix(1, 24) = "RBatchID"
   
   .RowHeight(1) = 700
   
   .WordWrap = True
   
End With
End Sub

Private Sub CmbLocation_GotFocus()
If Trim(CmbState.Text) = "" Then
   MsgBox "Select State first!!!"
   CmbState.SetFocus
   Exit Sub
End If
tmp = CmbLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID = " & mStateID)
Else
   Call TableMst_Location("where StateID = " & mStateID & " " & Gen_WcLocation)
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
mLOcationID = RsMst_Location!LocationID
End Sub

Private Sub CmbState_GotFocus()
tmp = CmbState.Text

If Gen_WcState = "" Then
   Call TableMst_State("")
Else
   Call TableMst_State("where " & Gen_WcLocation)
End If
CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
   CmbState.AddItem RsMst_State!StateName
   RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub

Private Sub CmbState_LostFocus()
If Trim(CmbState.Text) = "" Then
   CmbLocation.Text = ""
   Exit Sub
End If
If CmbState.Text <> tmp Then
   CmbLocation.Text = ""
End If

RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection!!!"
   CmbState.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmbStatus_LostFocus()

If Trim(CmbStatus.Text) = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "confirmed" And LCase(CmbStatus.Text) <> "reject" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date
TxtToDate.Value = Date
Call TxtFromDate_LostFocus
Call TxtToDate_LostFocus
CmbState.Text = ""
CmbLocation.Text = ""
CmbStatus.Text = ""
CmbNonNcdexClient.Text = ""
ChkFAG.Value = vbUnchecked
Call Grid_Head
End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant

If Format(TxtFromDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future month not allowed!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If
If Format(TxtToDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future month not allowed!!!"
   TxtToDate.SetFocus
   Exit Sub
End If
If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date Should Be Less Than To Date!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If

wc = " Where TPI.Billingmonth between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"

If Trim(CmbState.Text) <> "" Then
   wc = wc & " And ML.StateID = " & mStateID
End If

If Trim(CmbLocation.Text) <> "" Then
   wc = wc & " And ML.LocationID = " & mLOcationID
End If

If Trim(CmbNonNcdexClient.Text) <> "" Then
   wc = wc & " And MC.ClientIDRow = " & mClientID
End If

If Trim(CmbStatus.Text) <> "" Then
   wc = wc & " And TPI.Flag = '" & Left$(CmbStatus.Text, 1) & "'"
End If

If ChkFAG.Value = vbUnchecked Then
   wc = wc & " And  TPI.BatchID = 0 "
End If

Call Grid_Head

MsgBox "Wait till next message!!!"

tmp = " Select"
tmp = tmp & " MS.StateName, TPI.InvoiceNo,TPI.InvoiceDate,TPI.BillingMonth,ML.LocationName,ME.ExchangeType,MC.ClientName,"
tmp = tmp & " TPI.Discount, TPI.Premium, TPI.BatchID, TPI.ServiceTaxNAVCode, TPI.FAGPostingDate, TPI.Flag, TPI.Amount, "
tmp = tmp & " sum(InvoiceAmount) GSLInvoiceAmount"
tmp = tmp & " ,TTD.TIID,TTD.TINo,TTD.TIDate,TTD.TIAmount,TTD.TISTAmount,TTD.EntryNo,TPI.ReverseFlag,ReverseRemark,ReverseDate,RBatchID"
tmp = tmp & " From Txn_ProInvoice TPI"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientIDRow = MC.ClientIDRow"
tmp = tmp & " Left Join Txn_TIDetails TTD On TPI.InvoiceNo =  TTD.PIID And TTD.PIFlag = 'NNNR' And TTD.TIAmount <> 0"
tmp = tmp & " Left jOin Txn_InvMonGSLDetail TIM on TIM.PINo = TPI.InvoiceNo and TIM.ExchangeTypeID = 1"
tmp = tmp & wc
tmp = tmp & " group by TPI.InvoiceNo,TPI.InvoiceDate,TPI.BillingMonth,ML.LocationName,ME.ExchangeType,MC.ClientName,"
tmp = tmp & " TPI.Amount , TPI.Discount, TPI.Premium, TPI.BatchID, TPI.ServiceTaxNAVCode, TPI.FAGPostingDate, TPI.Flag, MS.StateName"
tmp = tmp & " ,TTD.TIID,TTD.TINo,TTD.TIDate,TTD.TIAmount,TTD.TISTAmount,TTD.EntryNo,TPI.ReverseFlag,ReverseRemark,ReverseDate,RBatchID"
tmp = tmp & " Having abs(Sum(isnull(invoiceamount,0)) - tpi.amount) > 1 " ''Sum(isnull(invoiceamount,0)) <> tpi.amount"
tmp = tmp & " Order by TPI.InvoiceNO"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   CmdExcel.Enabled = True
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   ProgressBar1.Max = MSHFlexGrid1.Rows
   LblTotalRecord.Caption = TmpRs.RecordCount
   With MSHFlexGrid1
      For I = 2 To TmpRs.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!InvoiceNo), "", TmpRs!InvoiceNo) '
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!InvoiceDate), "", Format(TmpRs!InvoiceDate, "dd-MMM-yyyy")) '"Invoice Date"
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!BillingMonth), "", Format(TmpRs!BillingMonth, "dd-MMM-yyyy")) '"Billing Month"
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName) '"State Name"
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) '"Location Name"
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType) '"Exchange Type"
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName) '"Client Name"
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs!GSLInvoiceAmount), "", TmpRs!GSLInvoiceAmount) '"GSL Invoice Amount"
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount) '"Amount"
         .TextMatrix(I, 9) = IIf(IsNull(TmpRs!Discount), "", TmpRs!Discount) '"Discount"
         .TextMatrix(I, 10) = IIf(IsNull(TmpRs!Premium), "", TmpRs!Premium) '"Premium"
         .TextMatrix(I, 11) = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID) '"BatchID"
         .TextMatrix(I, 12) = IIf(IsNull(TmpRs!ServiceTaxNAVCode), "", TmpRs!ServiceTaxNAVCode) '"Service Tax NAVCode"
         .TextMatrix(I, 13) = IIf(IsNull(TmpRs!FAGPostingDate), "", Format(TmpRs!FAGPostingDate, "dd-MMM-yyyy")) '"FAG Posting Date"
         .TextMatrix(I, 14) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag) '"Flag"
         If LCase(.TextMatrix(I, 14)) = "c" Then
            .TextMatrix(I, 14) = "Confirmed"
         ElseIf LCase(.TextMatrix(I, 14)) = "p" Then
            .TextMatrix(I, 14) = "Pending"
         ElseIf LCase(.TextMatrix(I, 14)) = "r" Then
            .TextMatrix(I, 14) = "Reject"
         End If
         .TextMatrix(I, 15) = IIf(IsNull(TmpRs!TIID), "", TmpRs!TIID) '"TIID"
         .TextMatrix(I, 16) = IIf(IsNull(TmpRs!TINo), "", TmpRs!TINo) '"TINo"
         .TextMatrix(I, 17) = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-MMM-yyyy")) '"TI Date"
         .TextMatrix(I, 18) = IIf(IsNull(TmpRs!TIAmount), "", TmpRs!TIAmount) '"TI Amount"
         .TextMatrix(I, 19) = IIf(IsNull(TmpRs!TISTAmount), "", TmpRs!TISTAmount) '"TIST Amount"
         .TextMatrix(I, 20) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo) '"Entry No"
         .TextMatrix(I, 21) = IIf(IsNull(TmpRs!ReverseFlag), "", TmpRs!ReverseFlag) '"Reverse Flag"
         .TextMatrix(I, 22) = IIf(IsNull(TmpRs!ReverseRemark), "", TmpRs!ReverseRemark) '"Reverse Remark"
         .TextMatrix(I, 23) = IIf(IsNull(TmpRs!ReverseDate), "", Format(TmpRs!ReverseDate, "dd-MMM-yyyy")) '"Reverse Date"
         .TextMatrix(I, 24) = IIf(IsNull(TmpRs!RBatchID), "", TmpRs!RBatchID) '"RBatchID"
         
         .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) + Val(.TextMatrix(I, 7))
         .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(I, 8))
         .TextMatrix(0, 9) = Val(.TextMatrix(0, 9)) + Val(.TextMatrix(I, 9))
         .TextMatrix(0, 10) = Val(.TextMatrix(0, 10)) + Val(.TextMatrix(I, 10))
         .TextMatrix(0, 18) = Val(.TextMatrix(0, 18)) + Val(.TextMatrix(I, 18))
         .TextMatrix(0, 19) = Val(.TextMatrix(0, 19)) + Val(.TextMatrix(I, 19))
         
         TmpRs.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox "Done!!!"
Else
   MsgBox "No record found!!!"
End If
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame0.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame0.Enabled = True
Call CmdClear_Click
End Sub

Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = 1 ")
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub
Private Sub CmbNonNcdexClient_LostFocus()
If CmbNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mClientID = RsMst_Client!ClientIDRow
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NNNRMismatchInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NNNRMismatchInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
   Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub TxtFromDate_LostFocus()
TxtFromDate.Value = DateSerial(TxtFromDate.Year, TxtFromDate.Month, 1)
End Sub

Private Sub TxtToDate_LostFocus()
TxtToDate.Value = DateSerial(TxtToDate.Year, TxtToDate.Month + 1, 0)
End Sub
