VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CRTDSDetails 
   Caption         =   "CR TDS Details"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   9480
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbExpenseTypeID 
         Appearance      =   0  'Flat
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
         ItemData        =   "FrmRpt_CRTDSDetails.frx":0000
         Left            =   1395
         List            =   "FrmRpt_CRTDSDetails.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   1350
         Width           =   6375
      End
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmRpt_CRTDSDetails.frx":0004
         Left            =   11160
         List            =   "FrmRpt_CRTDSDetails.frx":0011
         TabIndex        =   22
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox CmbStateID 
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
         Left            =   4260
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   3540
      End
      Begin VB.ComboBox CmbLocationID 
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
         Left            =   7845
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   3240
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
         Height          =   375
         Left            =   12510
         TabIndex        =   8
         Top             =   945
         Width           =   1215
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
         Height          =   375
         Left            =   11295
         TabIndex        =   7
         Top             =   945
         Width           =   1215
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
         Height          =   375
         Left            =   13725
         TabIndex        =   9
         Top             =   945
         Width           =   1215
      End
      Begin VB.TextBox TxtSearchBy 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   6
         Top             =   900
         Width           =   2850
      End
      Begin VB.ComboBox CmbSearchBy 
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
         ItemData        =   "FrmRpt_CRTDSDetails.frx":0032
         Left            =   1395
         List            =   "FrmRpt_CRTDSDetails.frx":003F
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   915
         Width           =   2760
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CR Date"
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
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "TI Date"
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
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   1410
         TabIndex        =   1
         Top             =   480
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   108068865
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   2835
         TabIndex        =   2
         Top             =   480
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   108068865
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label5 
         Caption         =   "Expense Ac."
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
         TabIndex        =   25
         Top             =   1403
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   11790
         TabIndex        =   23
         Top             =   165
         Width           =   720
      End
      Begin VB.Label LblStateID 
         Caption         =   "State Name"
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
         Left            =   5483
         TabIndex        =   19
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "To  Date"
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
         Left            =   3090
         TabIndex        =   18
         Top             =   165
         Width           =   870
      End
      Begin VB.Label Label6 
         Caption         =   "From  Date"
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
         Left            =   1568
         TabIndex        =   17
         Top             =   165
         Width           =   1065
      End
      Begin VB.Label LblLocationID 
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
         Left            =   9083
         TabIndex        =   16
         Top             =   165
         Width           =   765
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
         Left            =   11775
         TabIndex        =   15
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   9495
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblSearchBy 
         Caption         =   "TI No"
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
         Left            =   4320
         TabIndex        =   13
         Top             =   960
         Width           =   2460
      End
      Begin VB.Label Label2 
         Caption         =   "Search By"
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
         Left            =   150
         TabIndex        =   12
         Top             =   975
         Width           =   1005
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7335
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12938
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_CRTDSDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCommodityID, mExchangeTypeID, mRow, mExpenseTypeID As Variant

Private Sub CmbExpenseTypeID_GotFocus()
Dim wc As String
tmp = CmbExpenseTypeID.Text
Call TableMst_ExpenseType(" Where (AllocateTo = 'B' Or AllocateTo = 'E') And OtherDeductionFlag = 'Y' ")
CmbExpenseTypeID.Clear
If RsMst_ExpenseType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExpenseType.RecordCount
    CmbExpenseTypeID.AddItem RsMst_ExpenseType!ExpenseType
    RsMst_ExpenseType.MoveNext
Next
CmbExpenseTypeID.Text = tmp
End Sub
'----
Private Sub CmbExpenseTypeID_LostFocus()
If CmbExpenseTypeID.Text = "" Then
   mExpenseTypeID = 0
   Exit Sub
End If
RsMst_ExpenseType.MoveFirst
RsMst_ExpenseType.Find "ExpenseType = '" & CmbExpenseTypeID.Text & "'"
If RsMst_ExpenseType.EOF Then
   MsgBox "Invalid selection "
   CmbExpenseTypeID.SetFocus
   Exit Sub
End If
mExpenseTypeID = RsMst_ExpenseType!ExpenseTypeID
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then
   Exit Sub
End If

If LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdClear_Click()
Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
CmbLocationID.Text = ""
CmbSearchBy.Text = ""
CmbStateID.Text = ""
TxtSearchBy = ""
CmbSStatus.Text = ""
LblTotalRecord.Caption = ""
End Sub

Private Sub CmdGetReport_Click()
Dim Wc1 As Variant
Wc1 = ""

ProgressBar1.Value = 0
Call Grid_Head

MsgBox "Wait till next message!!!"

If TxtToDate.Value < TxtFromDate.Value Then
   MsgBox "To Date should be greater than From Date !!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If Option1.Value = True Then
   Wc1 = " Where TCT.AdjDate >='" & TxtFromDate.Value & "' And TCT.AdjDate<='" & TxtToDate.Value & "'"
Else
   Wc1 = " Where TI.TIDate >='" & TxtFromDate.Value & "' And TI.TIDate<='" & TxtToDate.Value & "'"
End If

If CmbStateID.Text <> "" Then
   Wc1 = Wc1 & " And MS.StateName= '" & CmbStateID.Text & "' "
End If

If CmbLocationID.Text <> "" Then
   Wc1 = Wc1 & " And ML.LocationName= '" & CmbLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      Wc1 = Wc1 & " And ML." & Gen_WcLocation
   End If
End If

If CmbSearchBy.Text <> "" Then
   If Left(LCase(CmbSearchBy.Text), 1) = "t" Then
      Wc1 = Wc1 & " And TI.TINo like '%" & TxtSearchBy & "%'"
   ElseIf Left(LCase(CmbSearchBy.Text), 1) = "p" Then
      Wc1 = Wc1 & " And TI.PostInvoiceNo like '%" & TxtSearchBy & "%'"
   ElseIf Left(LCase(CmbSearchBy.Text), 1) = "e" Then
      Wc1 = Wc1 & " And TI.EntryNo like '%" & TxtSearchBy & "%'"
   End If
End If

If CmbSStatus.Text <> "" Then
    Wc1 = Wc1 & " And  TCT.Flag = '" & Left(CmbSStatus.Text, 1) & "'"
End If

If CmbExpenseTypeID.Text <> "" Then
   Wc1 = Wc1 & " And  MEx.ExpenseType = '" & CmbExpenseTypeID.Text & "'"
End If


tmp = " Select MS.StateName,ML.LocationName,TCTD.CRTDSAdjustmentID,TCT.AdjDate,ME.ExchangeType,MCl.ClientName+'~'+MCl.ClientIDRow 'ClientName',"
tmp = tmp & " MD.DepositoryName+'~'+MD.DepositoryID 'DP', TCT.Remark, TCT.Flag ,TCT.BatchID,MEx.ExpenseType,"
tmp = tmp & " TI.TINo , TI.EntryNo, TI.TIDate, TI.TIAmount, TCTD.RecTDS, TCTD.RecOtherDed,TCT.FAGPostingDate,TI.NAVCode,TI.PostInvoiceNo "
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCTD"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCT On TCTD.CRTDSAdjustmentID = TCT.CRTDSAdjustmentID And TCTD.SM_Prefix = TCT.SM_Prefix"
tmp = tmp & " Inner Join Mst_Location ML On TCT.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TCT.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_State MS on ML.STateID = MS.StateID"
tmp = tmp & " Left Join Mst_Client MCl on TCT.ClientID = MCl.ClientID"
tmp = tmp & " Left Join Mst_Depository MD On TCT.DepositoryID = MD.DepositoryID"
tmp = tmp & " Inner Join Txn_TIDetails TI On TCTD.TIID = TI.TIID"
tmp = tmp & " Left Join Mst_ExpenseType MEx On TCTD.ExpenseAccID = MEx.ExpenseTypeID"
tmp = tmp & Wc1
tmp = tmp & " Order By ML.LocationName,TCTD.CRTDSAdjustmentID"

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 3

For I = 2 To TmpRs.RecordCount + 1
    MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!StateName
    MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!Locationname
    MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!CRTDSAdjustmentID
    MSHFlexGrid1.TextMatrix(I, 3) = Format(TmpRs!AdjDate, "dd-MMM-yyyy")
    MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!ExchangeType
    MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
    MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!DP), "", TmpRs!DP)
    MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
    MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag)
    
    If MSHFlexGrid1.TextMatrix(I, 8) = "P" Then
       MSHFlexGrid1.TextMatrix(I, 8) = "Pending"
    ElseIf MSHFlexGrid1.TextMatrix(I, 8) = "A" Then
       MSHFlexGrid1.TextMatrix(I, 8) = "Approved"
    ElseIf MSHFlexGrid1.TextMatrix(I, 8) = "R" Then
       MSHFlexGrid1.TextMatrix(I, 8) = "Rejected"
    End If
    
    MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID)
    MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!ExpenseType), "TDS", TmpRs!ExpenseType)
    MSHFlexGrid1.TextMatrix(I, 11) = TmpRs!TINo
    MSHFlexGrid1.TextMatrix(I, 12) = TmpRs!EntryNo
    MSHFlexGrid1.TextMatrix(I, 13) = Format(TmpRs!TIDate, "dd-MMM-yyyy") 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "D", "B")
    MSHFlexGrid1.TextMatrix(I, 14) = Format(TmpRs!TIAmount, "#####0.00") 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "W", "B")
    MSHFlexGrid1.TextMatrix(I, 15) = Format(TmpRs!RecTDS, "#####0.00") 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "D", "Q") '"Deposited Qty(in MT)"
    MSHFlexGrid1.TextMatrix(I, 16) = Format(TmpRs!RecOtherDed, "#####0.00") 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "W", "Q") '"Withdrawn Qty(in MT)"
    If IsNull(TmpRs!FAGPostingDate) = False Then
       MSHFlexGrid1.TextMatrix(I, 17) = Format(TmpRs!FAGPostingDate, "dd-MMM-yyyy")
    Else
       MSHFlexGrid1.TextMatrix(I, 17) = ""
    End If
    MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
    MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs!PostInvoiceNo), "", TmpRs!PostInvoiceNo)
    MSHFlexGrid1.TextMatrix(0, 15) = Val(MSHFlexGrid1.TextMatrix(0, 15)) + Val(MSHFlexGrid1.TextMatrix(I, 15))
    MSHFlexGrid1.TextMatrix(0, 16) = Val(MSHFlexGrid1.TextMatrix(0, 16)) + Val(MSHFlexGrid1.TextMatrix(I, 16))
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
LblTotalRecord.Caption = TmpRs.RecordCount
MsgBox "Done!!!"
End Sub
Private Sub Form_Load()
Call Grid_Head
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
TxtFromDate.Value = Date
TxtToDate.Value = Date
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2

MSHFlexGrid1.Cols = 20
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.AllowUserResizing = flexResizeColumns
MSHFlexGrid1.TextMatrix(1, 0) = "State Name"
MSHFlexGrid1.TextMatrix(1, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(1, 2) = "Adjustment ID"
MSHFlexGrid1.TextMatrix(1, 3) = "Adjustment Date"
MSHFlexGrid1.TextMatrix(1, 4) = "Exchange Type"
MSHFlexGrid1.TextMatrix(1, 5) = "Client Name and ID"
MSHFlexGrid1.TextMatrix(1, 6) = "DP Name and ID"
MSHFlexGrid1.TextMatrix(1, 7) = "Remarks"
MSHFlexGrid1.TextMatrix(1, 8) = "Flag"
MSHFlexGrid1.TextMatrix(1, 9) = "Batch ID"
MSHFlexGrid1.TextMatrix(1, 10) = "Expense Type"
MSHFlexGrid1.TextMatrix(1, 11) = "TI No"
MSHFlexGrid1.TextMatrix(1, 12) = "Entry No"
MSHFlexGrid1.TextMatrix(1, 13) = "TI Date"
MSHFlexGrid1.TextMatrix(1, 14) = "TI Amount"
MSHFlexGrid1.TextMatrix(1, 15) = "Received TDS against this CR"
MSHFlexGrid1.TextMatrix(1, 16) = "Other Deduduction"
MSHFlexGrid1.TextMatrix(1, 17) = "FAG Posting Date"
MSHFlexGrid1.TextMatrix(1, 18) = "NAV Code"
MSHFlexGrid1.TextMatrix(1, 19) = "Post Invoice No"

MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 800
MSHFlexGrid1.ColWidth(10) = 1500
MSHFlexGrid1.ColWidth(11) = 1500
MSHFlexGrid1.ColWidth(12) = 1500
MSHFlexGrid1.ColWidth(13) = 1200
MSHFlexGrid1.ColWidth(14) = 1500
MSHFlexGrid1.ColWidth(15) = 1500
MSHFlexGrid1.ColWidth(16) = 1500
MSHFlexGrid1.ColWidth(17) = 1500
MSHFlexGrid1.ColWidth(18) = 1500
MSHFlexGrid1.ColWidth(19) = 2000

MSHFlexGrid1.RowHeight(1) = 800
MSHFlexGrid1.Font.Size = 10


End Sub
Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState = "" Then
   Call TableMst_State
Else
   Call TableMst_State(" Where " & Gen_WcState)
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub
'----
Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

If mStateID <> RsMst_State!StateID Then
   CmbLocationID.Clear
   mLocationID = Null
End If

mStateID = RsMst_State!StateID

End Sub

Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CRTDSDetails.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CRTDSDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

oWS.Range("a:a").ColumnWidth = 7
oWS.Range("b:b").ColumnWidth = 13.57
oWS.Range("c:c").ColumnWidth = 9.43
oWS.Range("d:d").ColumnWidth = 10.57
oWS.Range("e:e").ColumnWidth = 19.29
oWS.Range("f:f").ColumnWidth = 9.29
oWS.Range("g:g").ColumnWidth = 10
oWS.Range("h:h").ColumnWidth = 11
oWS.Range("i:i").ColumnWidth = 11.43
oWS.Range("j:j").ColumnWidth = 10.71
oWS.Range("k:k").ColumnWidth = 11.43
oWS.Range("l:l").ColumnWidth = 11.43
oWS.Range("m:m").ColumnWidth = 11.43
oWS.Range("n:n").ColumnWidth = 11.43
oWS.Range("o:o").ColumnWidth = 11.43
oWS.Range("p:p").ColumnWidth = 11.43
oWS.Range("q:q").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "q" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "CR TDS Adjustment Details"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 3

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = mRow + 1  ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
        If I = 0 Then
           oWS.Cells(mRow, C + 1).Interior.Color = &HD9D9FF
           oWS.Cells(mRow, C + 1).Font.Bold = True
        ElseIf I = 1 Then
           oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, C + 1).Font.Bold = True
           oWS.Cells(mRow, C + 1).RowHeight = 45
           oWS.Cells(mRow, C + 1).WrapText = True
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
tmp = "a" & 1 & ":" & "q" & mRow
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub

Private Sub CmbSearchBy_LostFocus()

If CmbSearchBy.Text = "" Then
   lblSearchBy.Caption = "TI No"
   TxtSearchBy = ""
   TxtSearchBy.Locked = True
   Exit Sub
End If

If LCase(CmbSearchBy.Text) <> "ti no" And LCase(CmbSearchBy.Text) <> "post invoice no" And LCase(CmbSearchBy.Text) <> "entry no" Then
   MsgBox "Invalid Selection!!!"
   CmbSearchBy.SetFocus
   Exit Sub
End If

lblSearchBy.Caption = CmbSearchBy.Text
TxtSearchBy.Locked = False

End Sub


