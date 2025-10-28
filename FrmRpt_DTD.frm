VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_DTD 
   Caption         =   "DTD Report"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12726
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15000
      Begin VB.ComboBox CmbTransType 
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
         ItemData        =   "FrmRpt_DTD.frx":0000
         Left            =   120
         List            =   "FrmRpt_DTD.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1155
         Width           =   2880
      End
      Begin VB.ComboBox CmbCommodityID 
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
         Left            =   6630
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1170
         Width           =   3480
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         ItemData        =   "FrmRpt_DTD.frx":0023
         Left            =   3120
         List            =   "FrmRpt_DTD.frx":0025
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1170
         Width           =   3480
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
         Left            =   6630
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   3480
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
         Left            =   3105
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   3465
      End
      Begin VB.ComboBox CmbCMPName 
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
         Left            =   10155
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   420
         Width           =   3375
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
         Height          =   420
         Left            =   11715
         TabIndex        =   10
         Top             =   1125
         Width           =   1455
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
         Height          =   420
         Left            =   10260
         TabIndex        =   9
         Top             =   1125
         Width           =   1455
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
         Height          =   420
         Left            =   13170
         TabIndex        =   11
         Top             =   1125
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
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
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss"
         Format          =   119799811
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   1620
         TabIndex        =   2
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
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
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss"
         Format          =   119799811
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label3 
         Caption         =   "TransactionType"
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
         Left            =   773
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LblCommodityID 
         Caption         =   "Commodity"
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
         Left            =   7830
         TabIndex        =   18
         Top             =   885
         Width           =   1080
      End
      Begin VB.Label Label14 
         Caption         =   "Exchange Type"
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
         Left            =   4133
         TabIndex        =   17
         Top             =   855
         Width           =   1455
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
         Left            =   7920
         TabIndex        =   16
         Top             =   180
         Width           =   765
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4305
         TabIndex        =   15
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label LblCMPName 
         Caption         =   "CMP"
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
         Left            =   11535
         TabIndex        =   14
         Top             =   180
         Width           =   600
      End
      Begin VB.Label Label1 
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
         Left            =   300
         TabIndex        =   19
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   165
         Width           =   780
      End
   End
End
Attribute VB_Name = "FrmRpt_DTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MaxRow, mCMPID, mLocationID, mStateID, mExchangeTypeID, mCommodityID, mSM_Prefix As Variant
Dim dtCurrentDate As Date

Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection!!! "
   CmbCommodityID.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID

End Sub

Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
   
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub
Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Qry = " where " & Gen_WcState & ""
Else
   Qry = ""
End If
Call TableMst_State(Qry)
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
mStateID = RsMst_State!StateID
End Sub
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation <> "" Then
   Qry = " Where  " & Gen_WcLocation & " And StateID= " & mStateID & ""
Else
   Qry = " Where LocationID <> 0 And StateID= " & mStateID & ""
End If
If CmbStateID.Text = "" Then
   MsgBox "Please Select State !!"
   CmbStateID.SetFocus
   Exit Sub
End If

'qry = " Where LocationID <> 0 And StateID= " & mStateID & ""
Call TableMst_Location(Qry)
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "First Select State  !!!! "
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
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbTransType_LostFocus()

If CmbTransType.Text = "" Then Exit Sub

If LCase(CmbTransType.Text) <> "deposit" And LCase(CmbTransType.Text) <> "withdrawal" Then
   MsgBox "Invalid Selection"
   CmbTransType.SetFocus
   Exit Sub
End If

End Sub


Private Sub CmdClear_Click()
Call Grid_Head
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbExchangeTypeID.Text = ""
CmbCommodityID.Text = ""
CmbTransType.Text = ""

dtCurrentDate = Date
TxtFromDate.Value = Date
TxtToDate.Value = Date
CmdExcel.Enabled = False
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
'   CmdGetReport.Enabled = False
'   CmdExcel.Enabled = False
'   CmdClear.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
CmdExcel.Enabled = True 'False
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 24
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Deposit(D) Or Withdrawal(W)"
MSHFlexGrid1.TextMatrix(0, 1) = "Ncdex or Non-NCDEX Transaction"
MSHFlexGrid1.TextMatrix(0, 2) = "Date (dd/mm/yyyy format)"
MSHFlexGrid1.TextMatrix(0, 3) = "CMP"
MSHFlexGrid1.TextMatrix(0, 4) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 5) = "CDTF /CITF No"
MSHFlexGrid1.TextMatrix(0, 6) = "Name of Depositor / Withdrawee"

'--
'MSHFlexGrid1.TextMatrix(0, 7) = "Address of Depositor / Withdrawee"
'--
'MSHFlexGrid1.TextMatrix(0, 8) = "TTD No"
'--
'9
MSHFlexGrid1.TextMatrix(0, 7) = "Truck No"

MSHFlexGrid1.TextMatrix(0, 8) = "WB Slip No" '9'
'10
MSHFlexGrid1.TextMatrix(0, 9) = "Total Weight in MT" '10'
'11
MSHFlexGrid1.TextMatrix(0, 10) = "Tare Weight in MT" '11'

'12
MSHFlexGrid1.TextMatrix(0, 11) = "Gross Weight (including Gunny Weight) in MT"
'13
'--
'MSHFlexGrid1.TextMatrix(0, 13) = "GCTD No"
'--

'14
MSHFlexGrid1.TextMatrix(0, 12) = "Godown No (Enter only ONE Godown Number per row)"
'15
MSHFlexGrid1.TextMatrix(0, 13) = "Stack No (Enter only ONE Stack Number per row)"
'16
MSHFlexGrid1.TextMatrix(0, 14) = "Lot No ( Enter only ONE Lot number per row)"
'17
MSHFlexGrid1.TextMatrix(0, 15) = "Lotwise No of Bags"
'18
MSHFlexGrid1.TextMatrix(0, 16) = "Lotwise Qty in MT (Gross Weight)"
'19

MSHFlexGrid1.TextMatrix(0, 17) = "Truck No mapped to Lot"
MSHFlexGrid1.TextMatrix(0, 18) = "Weigh Bridge Slip No mapped to Lot"


MSHFlexGrid1.TextMatrix(0, 19) = "Commodity"

'20
MSHFlexGrid1.TextMatrix(0, 20) = "Client ID (for NCDEX Deposits)"
'21
MSHFlexGrid1.TextMatrix(0, 21) = "DP ID (for NCDEX Deposits)"
'22
'--
'MSHFlexGrid1.TextMatrix(0, 22) = "Depositor 's Name in WHR (for Non-NCDEX Deposits)"
'--
'23
MSHFlexGrid1.TextMatrix(0, 22) = "WHR No (for Non-NCDEX Transactions)"
'24
MSHFlexGrid1.TextMatrix(0, 23) = "RRN No / Rejected (for NCDEX Withdrawal)"


MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 2000
'MSHFlexGrid1.ColWidth(7) = 2000
'MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1200
'MSHFlexGrid1.ColWidth(13) = 1200
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1200
MSHFlexGrid1.ColWidth(14) = 1000
MSHFlexGrid1.ColWidth(15) = 1200
MSHFlexGrid1.ColWidth(16) = 1200
MSHFlexGrid1.ColWidth(17) = 1000
MSHFlexGrid1.ColWidth(18) = 1200
'MSHFlexGrid1.ColWidth(22) = 1000
MSHFlexGrid1.ColWidth(19) = 1200
MSHFlexGrid1.ColWidth(20) = 1200
MSHFlexGrid1.ColWidth(21) = 1200
MSHFlexGrid1.ColWidth(22) = 1200
MSHFlexGrid1.ColWidth(23) = 1200

MSHFlexGrid1.RowHeight(0) = 1600

End Sub

Private Sub CmdGetReport_Click()
Dim SR, SSR As Variant
Dim tmp1 As String
Dim wc As String
wc = ""

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date Must be less than To Date !!!"
   TxtFromDate.SetFocus
   Exit Sub
End If


If TxtToDate.Value - TxtFromDate.Value > 15 Then
   If CmbStateID.Text = "" Then
      MsgBox "State Selection is must for more than 15 days !!!"
      CmbStateID.SetFocus
      Exit Sub
   End If
End If
Wc1 = ""
If CmbStateID.Text <> "" Then
   tmp = "Select Distinct SM_Prefix From Mst_Location WHere StateID = " & mStateID & " "
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      Wc1 = "SM_Prefix in (''"
      For I = 1 To TmpRs.RecordCount
          Wc1 = Wc1 & ",'" & TmpRs!SM_Prefix & "'"
          TmpRs.MoveNext
      Next
      Wc1 = Wc1 & ")"
   End If
   
Else
   If Gen_WcState <> "" Then
      tmp = "Select Distinct SM_Prefix From Mst_Location WHere " & Gen_WcState & " " 'StateID in " & mStateID & " "
      Call TmpTable
   
      If TmpRs.RecordCount > 0 Then
         Wc1 = "SM_Prefix in (''"
         For I = 1 To TmpRs.RecordCount
             Wc1 = Wc1 & ",'" & TmpRs!SM_Prefix & "'"
             TmpRs.MoveNext
         Next
         Wc1 = Wc1 & ")"
      End If
   End If
End If


'9821694217
tmp = "Select * from Txn_CxTF_Details TCD"
tmp = tmp & " Inner Join Mst_Commodity MC on MC.CommodityID=TCD.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType ME on ME.ExchangeTypeID=TCD.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID=TCD.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.SM_PreFix=TCD.SM_PreFix"
tmp = tmp & " Inner Join Mst_State MS on Ms.StateID=ML.StateID "
tmp = tmp & " Left Join  Mst_Agent MA on TCD.AgentID=MA.AgentID "
If Gen_DBType = "MDB" Then
   wc = " Where (TCD.CxTFDate between  #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# and #" & Format(TxtToDate.Value, Gen_DatFormat) & "# )"
Else
   wc = " Where (TCD.CxTFDate between  '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' and '" & Format(TxtToDate.Value, Gen_DatFormat) & "' )"
End If

If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "MS.StateID", "N", "N")
End If

If CmbLocationID.Text <> "" Then
   wc = GenWc(wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbCMPName.Text <> "" Then
   wc = GenWc(wc, mCMPID, "MCMP.CMPID", "N", "N")
End If

If CmbExchangeTypeID.Text <> "" Then
   wc = GenWc(wc, mExchangeTypeID, "ME.ExchangeTypeID", "N", "N")
End If

If CmbCommodityID.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "MC.CommodityID", "N", "N")
End If

'If Gen_WcLocation <> "" Then
'   Wc = " And  ML." & Gen_WcLocation
'End If

If CmbTransType.Text <> "" Then
   wc = wc & " And TCD.TxnType='" & UCase(Left(CmbTransType.Text, 1)) & "'"
End If

tmp = tmp & wc
tmp = tmp & " Order By TCD.TxnType"

Call TmpTable

WC2 = ""
If TmpRs.RecordCount > 0 Then
   WC2 = "Where CxTF_DetailsID in (Select CxTF_DetailsID From Txn_CxTF_Details Where CxTFDate between  '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' and '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
End If

Call Grid_Head
'Call TableTxn_CxTF_Details("Where CmpID='" & mCMPName & "' And CxTFDate = '" & DTPicker1.Value & "' order by TxnType")

If CmbLocationID.Text <> "" Then
   If WC2 <> "" Then
      WC2 = WC2 & " And SM_Prefix = '" & mSM_Prefix & "'"
   Else
      WC2 = " Where SM_Prefix = '" & mSM_Prefix & "'"
   End If
End If

If CmbTransType.Text <> "" Then
   If WC2 <> "" Then
      WC2 = WC2 & " And TxnType = '" & Left(CmbTransType.Text, 1) & "'"
   Else
      WC2 = " Where TxnType = '" & Left(CmbTransType.Text, 1) & "'"
   End If
End If

If WC2 = "" Then
   If Wc1 <> "" Then
      WC2 = "Where " & Wc1
   End If
End If

Call TableTxn_CxTF_ClientDetail(WC2)
Call TableTxn_CxTF_TruckDetail(WC2) '(" where  SM_Prefix = '" & RsTxn_CxTF_Details!SM_Prefix & "' And CxTFNo = " & RsTxn_CxTF_Details!CxTFNo & "  And TxnType = '" & RsTxn_CxTF_Details!TxnType & "' ")
Call TableTxn_CxTF_GSL(WC2) '(" where  SM_Prefix = '" & RsTxn_CxTF_Details!SM_Prefix & "' And CxTFNo = " & RsTxn_CxTF_Details!CxTFNo & "  And TxnType = '" & RsTxn_CxTF_Details!TxnType & "' ")

If TmpRs.RecordCount > 0 Then
   MsgBox ("Wait till next message")
   SR = 1
   For I = 1 To TmpRs.RecordCount
        
       MSHFlexGrid1.TextMatrix(SR, 0) = TmpRs!TxnType
       If LCase(TmpRs!ExchangeType) = "spot" Then 'GetExchangeName(TmpRs!ExchangeTypeID) = "SPOT" Then
          MSHFlexGrid1.TextMatrix(SR, 1) = "Non NCDEX" 'GetExchangeName(TmpRs!ExchangeTypeID)
       Else
          MSHFlexGrid1.TextMatrix(SR, 1) = TmpRs!ExchangeType 'GetExchangeName(TmpRs!ExchangeTypeID)
       End If
       MSHFlexGrid1.TextMatrix(SR, 2) = Format(TmpRs!CxTFDate, "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(SR, 3) = TmpRs!EHOPCMPID 'GetEHOPCMPID(TmpRs!CMPID)
       MSHFlexGrid1.TextMatrix(SR, 4) = TmpRs!CMPName 'GetCMPName(TmpRs!CMPID)
       MSHFlexGrid1.TextMatrix(SR, 5) = TmpRs!SM_Prefix & "-" & TmpRs!CxTF_DetailsID
       If IsNull(TmpRs!AgentID) = False Then
          MSHFlexGrid1.TextMatrix(SR, 6) = TmpRs!AgentName 'GetAgentName(TmpRs!AgentID) 'RsTxn_CxTF_GSL!SM_Prefix & "-" & RsTxn_CxTF_TruckDetail!CxTFNo
       Else
          MSHFlexGrid1.TextMatrix(SR, 6) = ""
       End If
       'MSHFlexGrid1.TextMatrix(SR, 7) = "" 'RsTxn_CxTF_GSL!WeighBridgeSlipno
       MSHFlexGrid1.TextMatrix(SR, 19) = TmpRs!Commodity 'GetCommodityName(TmpRs!CommodityID)
       
       RsTxn_CxTF_TruckDetail.Filter = " SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID = " & TmpRs!CxTF_DetailsID & "  And TxnType = '" & TmpRs!TxnType & "' "
       RsTxn_CxTF_GSL.Filter = "SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID = " & TmpRs!CxTF_DetailsID & "  And TxnType = '" & TmpRs!TxnType & "' "
            
       If RsTxn_CxTF_TruckDetail.RecordCount > RsTxn_CxTF_GSL.RecordCount Then
          MaxRow = RsTxn_CxTF_TruckDetail.RecordCount
       Else
          MaxRow = RsTxn_CxTF_GSL.RecordCount
       End If
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + MaxRow
       SSR = SR
       Call FillTruckDetails(SR)
       SR = SSR
       Call FillGSLDetails(SR)
       TmpRs.MoveNext
       
       SR = SSR + MaxRow
       RsTxn_CxTF_TruckDetail.Filter = "CxTF_DetailsID <> 0"
       RsTxn_CxTF_GSL.Filter = "CxTF_DetailsID <> 0"
       RsTxn_CxTF_ClientDetail.Filter = "CxTF_DetailsID <> 0"
       RsTxn_CxTF_TruckDetail.MoveFirst
       RsTxn_CxTF_GSL.MoveFirst
       RsTxn_CxTF_ClientDetail.MoveFirst
   Next
   
   'Label2.Caption = MaxRow
   CmdExcel.Enabled = True
   MsgBox ("Done!!!")
Else
   MsgBox "No record found for given Date!!!"
   Call Grid_Head
   Exit Sub
End If

End Sub

Private Sub FillTruckDetails(i1 As Variant)
For x = 1 To RsTxn_CxTF_TruckDetail.RecordCount
    MSHFlexGrid1.TextMatrix(i1, 0) = TmpRs!TxnType
    If LCase(TmpRs!ExchangeType) = "spot" Then 'GetExchangeName(TmpRs!ExchangeTypeID) = "SPOT" Then
       MSHFlexGrid1.TextMatrix(i1, 1) = "Non NCDEX" 'GetExchangeName(TmpRs!ExchangeTypeID)
    Else
       MSHFlexGrid1.TextMatrix(i1, 1) = TmpRs!ExchangeType 'GetExchangeName(TmpRs!ExchangeTypeID)
    End If
    MSHFlexGrid1.TextMatrix(i1, 2) = Format(TmpRs!CxTFDate, "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(i1, 3) = TmpRs!EHOPCMPID 'GetEHOPCMPID(TmpRs!CMPID)
    MSHFlexGrid1.TextMatrix(i1, 4) = TmpRs!CMPName 'GetCMPName(TmpRs!CMPID)
    MSHFlexGrid1.TextMatrix(i1, 5) = TmpRs!SM_Prefix & "-" & TmpRs!CxTF_DetailsID
    If IsNull(TmpRs!AgentID) = False Then
       MSHFlexGrid1.TextMatrix(i1, 6) = TmpRs!AgentName 'GetAgentName(TmpRs!AgentID) 'RsTxn_CxTF_GSL!SM_Prefix & "-" & RsTxn_CxTF_TruckDetail!CxTFNo
    Else
       MSHFlexGrid1.TextMatrix(i1, 6) = ""
    End If
    MSHFlexGrid1.TextMatrix(i1, 19) = TmpRs!Commodity 'GetCommodityName(TmpRs!CommodityID)
       
    MSHFlexGrid1.TextMatrix(i1, 7) = RsTxn_CxTF_TruckDetail!TruckNo
    MSHFlexGrid1.TextMatrix(i1, 8) = RsTxn_CxTF_TruckDetail!WeighBridgeSlipno
    MSHFlexGrid1.TextMatrix(i1, 9) = RsTxn_CxTF_TruckDetail!TotalWeight
    MSHFlexGrid1.TextMatrix(i1, 10) = RsTxn_CxTF_TruckDetail!TareWeight
    MSHFlexGrid1.TextMatrix(i1, 11) = RsTxn_CxTF_TruckDetail!GrossWeight
    RsTxn_CxTF_TruckDetail.MoveNext
    i1 = i1 + 1
Next

End Sub

Private Sub FillGSLDetails(SR2 As Variant)
Dim ans, a As Variant
For x = 1 To RsTxn_CxTF_GSL.RecordCount
    MSHFlexGrid1.TextMatrix(SR2, 0) = TmpRs!TxnType
    If LCase(TmpRs!ExchangeType) = "spot" Then 'GetExchangeName(TmpRs!ExchangeTypeID) = "SPOT" Then
       MSHFlexGrid1.TextMatrix(SR2, 1) = "Non NCDEX" 'GetExchangeName(TmpRs!ExchangeTypeID)
    Else
       MSHFlexGrid1.TextMatrix(SR2, 1) = TmpRs!ExchangeType 'GetExchangeName(TmpRs!ExchangeTypeID)
    End If
    MSHFlexGrid1.TextMatrix(SR2, 2) = Format(TmpRs!CxTFDate, "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(SR2, 3) = TmpRs!EHOPCMPID 'GetEHOPCMPID(TmpRs!CMPID)
    MSHFlexGrid1.TextMatrix(SR2, 4) = TmpRs!CMPName 'GetCMPName(TmpRs!CMPID)
    MSHFlexGrid1.TextMatrix(SR2, 5) = TmpRs!SM_Prefix & "-" & TmpRs!CxTF_DetailsID
    If IsNull(TmpRs!AgentID) = False Then
       MSHFlexGrid1.TextMatrix(SR2, 6) = TmpRs!AgentName ' GetAgentName(TmpRs!AgentID) 'RsTxn_CxTF_GSL!SM_Prefix & "-" & RsTxn_CxTF_TruckDetail!CxTFNo
    Else
       MSHFlexGrid1.TextMatrix(SR2, 6) = ""
    End If
    MSHFlexGrid1.TextMatrix(SR2, 17) = IIf(IsNull(RsTxn_CxTF_GSL!TruckNo), "", RsTxn_CxTF_GSL!TruckNo) '
    MSHFlexGrid1.TextMatrix(SR2, 18) = IIf(IsNull(RsTxn_CxTF_GSL!WeighBridgeSlipno), "", RsTxn_CxTF_GSL!WeighBridgeSlipno) '
    
    MSHFlexGrid1.TextMatrix(SR2, 19) = TmpRs!Commodity 'GetCommodityName(TmpRs!CommodityID)
    
    
    'MSHFlexGrid1.TextMatrix(SR2, 13) = RsTxn_CxTF_GSL!SM_Prefix & "-" & RsTxn_CxTF_GSL!CxTFNo
    ans = GetGSLByGLSID(RsTxn_CxTF_GSL!Mst_GSL_ID, RsTxn_CxTF_GSL!SM_Prefix)
    MSHFlexGrid1.TextMatrix(SR2, 12) = ans(3)
    MSHFlexGrid1.TextMatrix(SR2, 13) = ans(4) 'RsTxn_CxTF_GSL!GrossWeight
    MSHFlexGrid1.TextMatrix(SR2, 14) = ans(5) 'RsTxn_CxTF_GSL!GrossWeight
    MSHFlexGrid1.TextMatrix(SR2, 15) = RsTxn_CxTF_GSL!NoofBags
    MSHFlexGrid1.TextMatrix(SR2, 16) = RsTxn_CxTF_GSL!GrossWeight 'RsTxn_CxTF_GSL!GrossWeight
    
    If RsTxn_CxTF_GSL!TxnType = "D" Then
       'ans = GetClientNameByRow(RsTxn_CxTF_GSL!ClientIDRow)
       If TmpRs!ExchangeTypeID = 2 Then
          ans = GetClientNameByRow(RsTxn_CxTF_GSL!ClientID)
          MSHFlexGrid1.TextMatrix(SR2, 20) = ans(0) 'Client ID
          MSHFlexGrid1.TextMatrix(SR2, 21) = ans(2) 'DPID
          'MSHFlexGrid1.TextMatrix(SR2, 22) = "" 'DPName
       Else
          MSHFlexGrid1.TextMatrix(SR2, 20) = "" 'Client ID
          MSHFlexGrid1.TextMatrix(SR2, 21) = "" 'DPID
          'MSHFlexGrid1.TextMatrix(SR2, 22) = ans(1) 'DPName
       End If
    Else
       MSHFlexGrid1.TextMatrix(SR2, 20) = "" 'ans(0)
       MSHFlexGrid1.TextMatrix(SR2, 21) = "" 'ans(2) 'RsTxn_CxTF_GSL!GrossWeight
       'MSHFlexGrid1.TextMatrix(SR2, 22) = "" 'ans(3) 'DPName
    End If
    MSHFlexGrid1.TextMatrix(SR2, 22) = "" 'RsTxn_CxTF_GSL!WHRNo
    If MSHFlexGrid1.TextMatrix(SR2, 0) = "D" Then
       MSHFlexGrid1.TextMatrix(SR2, 23) = "" 'RsTxn_CxTF_GSL!RRNNo
    Else
       If TmpRs!ExchangeTypeID = 2 Then
          RsTxn_CxTF_ClientDetail.Filter = " SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID = " & TmpRs!CxTF_DetailsID & "  And TxnType = '" & TmpRs!TxnType & "' "
          For a = 0 To RsTxn_CxTF_ClientDetail.RecordCount - 1
              If MSHFlexGrid1.TextMatrix(SR2, 23) = "" Then
                 'MSHFlexGrid1.TextMatrix(SR2, 24) = RsTxn_CxTF_ClientDetail!RRNID
                 MSHFlexGrid1.TextMatrix(SR2, 23) = IIf(IsNull(RsTxn_CxTF_ClientDetail!RRNID), "", RsTxn_CxTF_ClientDetail!RRNID)
                 If MSHFlexGrid1.TextMatrix(SR2, 23) <> "" Then
                    MSHFlexGrid1.TextMatrix(SR2, 23) = GetRRNNo(MSHFlexGrid1.TextMatrix(SR2, 23))
                 End If
              Else
                 MSHFlexGrid1.TextMatrix(SR2, 23) = MSHFlexGrid1.TextMatrix(SR2, 23) & "," & IIf(IsNull(RsTxn_CxTF_ClientDetail!RRNID), "", GetRRNNo(RsTxn_CxTF_ClientDetail!RRNID))
              End If
              RsTxn_CxTF_ClientDetail.MoveNext
          Next
          
       Else
          MSHFlexGrid1.TextMatrix(SR2, 23) = ""
       End If
    End If
    RsTxn_CxTF_GSL.MoveNext
    SR2 = SR2 + 1
 
Next
'MaxRow = SR2
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp

'Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_DTD.xls")
'Call Xls_Detail_Rpt


tmp = ""
ans = MsgBox("Do you want to Genearate CSV File", vbYesNo)
If ans = vbNo Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_DTD.xls")
   Call Xls_Detail_Rpt
Else
   Path = App.Path
   strFilName = "\Excel\" & Gen_mTimeStamp & "Rpt_DTD" & Gen_UserName & ".csv"
   strFilLocation = Path + strFilName
   If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
   Open strFilLocation For Output As #1
   For I = 0 To MSHFlexGrid1.Rows - 1
       tmp = ""
       For C = 0 To MSHFlexGrid1.Cols - 1
          If C = 0 Then
             tmp = tmp & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
'          ElseIf c = 6 Then
'             tmp = tmp & "," & "'" & Replace(Replace(MSHFlexGrid1.TextMatrix(I, c), vbNewLine, ""), ",", "~")
          Else
             tmp = tmp & "," & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          End If
       Next
       Print #1, tmp
   Next
   Close #1
   MsgBox "CSV Generated!!!"
End If



End Sub

Public Sub Xls_Detail_Rpt()
Dim mSplitStr() As String
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_DTD.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              mSplitStr = Split(MSHFlexGrid1.TextMatrix(I, C), "/")
              If UBound(mSplitStr) = 2 Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "mm/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
              End If
           Else
              If C = 8 Or C = 12 Or C = 13 Or C = 14 Or C = 17 Or C = 18 Then
                 oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
           End If
        Else
           If C = 8 Or C = 12 Or C = 13 Or C = 14 Or C = 17 Or C = 18 Then
              oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbCMPName_GotFocus()


If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!"
   CmbLocationID.SetFocus
   Exit Sub
End If
tmp = CmbCMPName.Text

If Gen_WcCMP <> "" Then
   Qry = " Where " & Gen_WcCMP & " and LocationID =" & mLocationID & " "
Else
   Qry = " Where LocationID =" & mLocationID & ""
End If
'qry = " Where LocationID =" & mLocationID & " "
Call TableMst_CMP(Qry)
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub
'----
Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
    mCMPID = Null
   Exit Sub
End If
RsMst_CMP.MoveFirst

RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
'dtCurrentDate = RsMst_CMP!CurrentDate 'CDate(Format(RsMst_CMP!CurrentDate, "mm/dd/yyyy"))
'RptDate.Value = RsMst_CMP!CurrentDate
End Sub

