VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CRInvoices 
   Caption         =   "Cash Receipt wise Invoices"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1710
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.TextBox TxtCRNo 
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
         Left            =   11160
         TabIndex        =   3
         Top             =   480
         Width           =   2850
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
         TabIndex        =   21
         Top             =   600
         Width           =   1095
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
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
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
         ItemData        =   "FrmRpt_CRInvoices.frx":0000
         Left            =   1395
         List            =   "FrmRpt_CRInvoices.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   915
         Width           =   2760
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
         TabIndex        =   16
         Top             =   900
         Width           =   2850
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
         TabIndex        =   8
         Top             =   1185
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
         TabIndex        =   5
         Top             =   1185
         Width           =   1215
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
         TabIndex        =   4
         Top             =   1185
         Width           =   1215
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
         TabIndex        =   2
         Top             =   480
         Width           =   3240
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
         TabIndex        =   1
         Top             =   480
         Width           =   3540
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   1410
         TabIndex        =   6
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
         Format          =   111738881
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   2835
         TabIndex        =   7
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
         Format          =   111738881
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CR No"
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
         Left            =   12210
         TabIndex        =   22
         Top             =   165
         Width           =   600
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
         TabIndex        =   19
         Top             =   975
         Width           =   1005
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
         TabIndex        =   18
         Top             =   960
         Width           =   2460
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
         Left            =   7095
         TabIndex        =   14
         Top             =   1320
         Width           =   2175
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
         Left            =   9375
         TabIndex        =   13
         Top             =   1320
         Width           =   90
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
         TabIndex        =   12
         Top             =   165
         Width           =   765
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
         TabIndex        =   11
         Top             =   165
         Width           =   1065
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
         TabIndex        =   10
         Top             =   165
         Width           =   870
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
         TabIndex        =   9
         Top             =   165
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7815
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13785
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_CRInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCommodityID, mExchangeTypeID, mRow As Variant

Private Sub CmdClear_Click()
Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
CmbLocationID.Text = ""
CmbSearchBy.Text = ""
CmbStateID.Text = ""
TxtSearchBy = ""
LblTotalRecord.Caption = ""
End Sub

Private Sub CmdGetReport_Click()
Dim Wc1 As Variant
Wc1 = ""

Call Grid_Head

MsgBox "Wait till next message!!!"
If TxtToDate.Value < TxtFromDate.Value Then
   MsgBox "To Date should be greater than From Date !!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If Option1.Value = True Then
   Wc1 = " Where TCR.CRDate >='" & TxtFromDate.Value & "' And TCR.CRDate<='" & TxtToDate.Value & "'"
Else
   Wc1 = " Where TI.TIDate >='" & TxtFromDate.Value & "' And TI.TIDate<='" & TxtToDate.Value & "'"
End If

If CmbStateID.Text <> "" Then
   Wc1 = Wc1 & " And ML.StateID=" & mStateID & ""
End If

If CmbLocationID.Text <> "" Then
   Wc1 = Wc1 & " And ML.LocationID=" & mLocationID & ""
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

If Trim(TxtCRNo) <> "" Then
   Wc1 = Wc1 & " And TCR.NewCRID = " & Val(TxtCRNo) & ""
End If

Wc1 = Wc1 & " And TCR.ReturnFlag <> 'Y'"

tmp = "Select TCR.NewCRID,ML.LocationName,TCR.CRDate,TCR.InstrumentAmount,TI.TINo,TI.PostInvoiceNo,TI.EntryNo,TI.TIDate,TI.TIAmount, "
tmp = tmp & " TCRTID.RecAmount , TCRTID.RecTDS, TCRTID.RecOtherDed,TCR.ReceivedFrom,TCR.FAGPostingDate,TCR.InstrumentType,TCR.InstrumentNo,TCR.InstrumentDate,TCR.Remark,TCR.FAGRemarks,TCR.BatchID,TI.NAVCode,MC.ClientName,MD.DepositoryName,ME.ExchangeType 'ClientExchange',ME1.ExchangeType 'DepoExchange',MLI.LocationName 'InvoiceLocationName' "
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRTID"
'Parag
'tmp = tmp & " Inner Join Txn_NewCashReceipt TCR On TCRTID.CRID = TCR.CRID And TCRTID.SM_prefix = TCR.SM_Prefix"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR On TCRTID.NewCRID = TCR.NewCRID And TCRTID.SM_prefix = TCR.SM_Prefix"
tmp = tmp & " Inner Join Txn_TIDetails TI On TCRTID.TIID = TI.TIID"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocatioNID = ML.LocationID"
'Parag
'tmp = tmp & " Left Join Mst_Client MC On TI.ClientIDRow = MC.ClientIDRow"
tmp = tmp & " Left Join Mst_Client MC On TI.ClientID = MC.ClientID"
tmp = tmp & " Left Join Mst_Depository MD On TI.DepositoryID = MD.DepositoryID"
tmp = tmp & " Left Join Mst_ExchangeType ME On MC.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Left Join Mst_ExchangeType ME1 On MD.ExchangeTypeID = ME1.ExchangeTypeID"
tmp = tmp & " Left Join Mst_Location MLI On TI.LocatioNID = MLI.LocationID"
tmp = tmp & Wc1
tmp = tmp & " Order By ML.LocationName,TCR.NewCRID"


Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 3

For I = 2 To TmpRs.RecordCount + 1


    MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!NewCRID
    MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!LocationName
    MSHFlexGrid1.TextMatrix(I, 2) = Format(TmpRs!CRDate, "dd-MMM-yyyy")
    MSHFlexGrid1.TextMatrix(I, 3) = Format(TmpRs!InstrumentAmount, "#####0.00")
    MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!ReceivedFrom), "", TmpRs!ReceivedFrom)
    MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!FAGPostingDate), "", Format(TmpRs!FAGPostingDate, "dd-mmm-yyyy"))
    MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!InstrumentType), "", TmpRs!InstrumentType)
    MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!InstrumentNo), "", TmpRs!InstrumentNo)
    MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!InstrumentDate), "", Format(TmpRs!InstrumentDate, "dd-mmm-yyyy"))
    MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
    MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!FAGRemarks), "", TmpRs!FAGRemarks)
    MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID)
    MSHFlexGrid1.TextMatrix(I, 12) = TmpRs!TINo
    MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs!PostInvoiceNo), "", TmpRs!PostInvoiceNo)
    MSHFlexGrid1.TextMatrix(I, 14) = TmpRs!EntryNo
    MSHFlexGrid1.TextMatrix(I, 15) = Format(TmpRs!TIDate, "dd-MMM-yyyy") 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "D", "B")
    MSHFlexGrid1.TextMatrix(I, 16) = Format(TmpRs!TIAmount, "#####0.00") 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "W", "B")
    MSHFlexGrid1.TextMatrix(I, 17) = Format(TmpRs!RecAmount, "#####0.00")
    MSHFlexGrid1.TextMatrix(I, 18) = Format(TmpRs!RecTDS, "#####0.00") 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "D", "Q") '"Deposited Qty(in MT)"
    MSHFlexGrid1.TextMatrix(I, 19) = Format(TmpRs!RecOtherDed, "#####0.00") 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "W", "Q") '"Withdrawn Qty(in MT)"
    MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
    
    If IsNull(TmpRs!ClientName) = False Then
       MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
       MSHFlexGrid1.TextMatrix(I, 22) = IIf(IsNull(TmpRs!ClientExchange), "", TmpRs!ClientExchange)
    Else
       If IsNull(TmpRs!DepositoryName) = False Then
          MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
          MSHFlexGrid1.TextMatrix(I, 22) = IIf(IsNull(TmpRs!DepoExchange), "", TmpRs!DepoExchange)
       End If
    End If
    
    MSHFlexGrid1.TextMatrix(I, 23) = IIf(IsNull(TmpRs!InvoiceLocationName), "", TmpRs!InvoiceLocationName)
    
    MSHFlexGrid1.TextMatrix(0, 17) = Val(MSHFlexGrid1.TextMatrix(0, 17)) + Val(MSHFlexGrid1.TextMatrix(I, 17))
    MSHFlexGrid1.TextMatrix(0, 18) = Val(MSHFlexGrid1.TextMatrix(0, 18)) + Val(MSHFlexGrid1.TextMatrix(I, 18))
    MSHFlexGrid1.TextMatrix(0, 19) = Val(MSHFlexGrid1.TextMatrix(0, 19)) + Val(MSHFlexGrid1.TextMatrix(I, 19))
    TmpRs.MoveNext
Next
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
'
'ReceivedFrom
'FAGPostingDate
'InstrumentType
'InstrumentNo
'InstrumentDate
'Remark
'FAGRemarks
'BatchID
'NAV Code of the client

MSHFlexGrid1.Cols = 24 '23 '21 '12
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.AllowUserResizing = flexResizeColumns
MSHFlexGrid1.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid1.TextMatrix(1, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(1, 2) = "CR Date"
MSHFlexGrid1.TextMatrix(1, 3) = "CR Amount"

MSHFlexGrid1.TextMatrix(1, 4) = "Received From"
MSHFlexGrid1.TextMatrix(1, 5) = "FAG Posting Date"
MSHFlexGrid1.TextMatrix(1, 6) = "Instrument Type"
MSHFlexGrid1.TextMatrix(1, 7) = "Instrument No"
MSHFlexGrid1.TextMatrix(1, 8) = "Instrument Date"
MSHFlexGrid1.TextMatrix(1, 9) = "Remark"
MSHFlexGrid1.TextMatrix(1, 10) = "FAG Remarks"
MSHFlexGrid1.TextMatrix(1, 11) = "Batch ID"

MSHFlexGrid1.TextMatrix(1, 12) = "TI No"
MSHFlexGrid1.TextMatrix(1, 13) = "Post Invoice No"
MSHFlexGrid1.TextMatrix(1, 14) = "Entry No"
MSHFlexGrid1.TextMatrix(1, 15) = "TI Date"
MSHFlexGrid1.TextMatrix(1, 16) = "TI Amount"
MSHFlexGrid1.TextMatrix(1, 17) = "Received Amt against this CR"
MSHFlexGrid1.TextMatrix(1, 18) = "Received TDS against this CR"
MSHFlexGrid1.TextMatrix(1, 19) = "Other Deduduction"
MSHFlexGrid1.TextMatrix(1, 20) = "NAV Code of the client"
MSHFlexGrid1.TextMatrix(1, 21) = "Name of the client"
MSHFlexGrid1.TextMatrix(1, 22) = "Type Of Invoice"
MSHFlexGrid1.TextMatrix(1, 23) = "Invoice Location"


MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1500
MSHFlexGrid1.ColWidth(9) = 1500
MSHFlexGrid1.ColWidth(10) = 1500
MSHFlexGrid1.ColWidth(11) = 1500
MSHFlexGrid1.ColWidth(12) = 1500
MSHFlexGrid1.ColWidth(13) = 1500
MSHFlexGrid1.ColWidth(14) = 1500
MSHFlexGrid1.ColWidth(15) = 1500
MSHFlexGrid1.ColWidth(16) = 1500
MSHFlexGrid1.ColWidth(17) = 1500
MSHFlexGrid1.ColWidth(18) = 1500
MSHFlexGrid1.ColWidth(19) = 1500
MSHFlexGrid1.ColWidth(20) = 1500

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
    CmbLocationID.AddItem RsMst_Location!LocationName
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


   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CRInvoices.xls")
   Call Xls_Detail_Rpt


'''Dim Path, strFilName, strFilLocation As Variant
'''Gen_mTimeStamp = GetTimeStamp
'''tmp = ""
'''ans = MsgBox("Do you want to Genearate CSV File", vbYesNo)
'''If ans = vbNo Then
'''   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CRInvoices.xls")
'''   Call Xls_Detail_Rpt
'''Else
'''   Path = App.Path
'''   strFilName = "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CRInvoices.csv"
'''   strFilLocation = Path + strFilName
'''   If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
'''   Open strFilLocation For Output As #1
'''   For I = 0 To MSHFlexGrid1.Rows - 2
'''       tmp = ""
'''       For c = 0 To MSHFlexGrid1.Cols - 1
'''          If c = 0 Then
'''             tmp = tmp & Replace(Replace(MSHFlexGrid1.TextMatrix(I, c), vbNewLine, ""), ",", "~")
'''          ElseIf c = 12 Or c = 15 Then
'''             tmp = tmp & "," & Replace(Replace(MSHFlexGrid1.TextMatrix(I, c), vbNewLine, ""), ",", "~")
'''          Else
'''             tmp = tmp & "," & Replace(Replace(MSHFlexGrid1.TextMatrix(I, c), vbNewLine, ""), ",", "~")
'''          End If
'''       Next
'''       Print #1, tmp
'''   Next
'''   Close #1
'''   MsgBox "Done!!!"
'''End If


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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CRInvoices.xls")
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

tmp = "a" & mRow & ":" & "l" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Cash Receipt Invoices"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F

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
Next
tmp = "a" & 1 & ":" & "l" & mRow
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

