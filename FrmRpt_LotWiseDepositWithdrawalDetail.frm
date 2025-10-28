VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_LotWiseDepositWithdrawalDetail 
   Caption         =   "Lot Wise Deposit/Withdrawal detail Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   20340
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1830
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   15000
      Begin VB.ComboBox CmbCommodity 
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
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   1245
         Width           =   3510
      End
      Begin VB.TextBox TxtCxTFNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5985
         TabIndex        =   16
         Top             =   1230
         Width           =   1860
      End
      Begin VB.ComboBox CmbType 
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
         ItemData        =   "FrmRpt_LotWiseDepositWithdrawalDetail.frx":0000
         Left            =   3675
         List            =   "FrmRpt_LotWiseDepositWithdrawalDetail.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   1245
         Width           =   2280
      End
      Begin VB.ComboBox CmbExchangeType 
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
         ItemData        =   "FrmRpt_LotWiseDepositWithdrawalDetail.frx":0023
         Left            =   10935
         List            =   "FrmRpt_LotWiseDepositWithdrawalDetail.frx":0025
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   3960
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
         Left            =   7035
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   3900
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
         Height          =   495
         Left            =   13695
         TabIndex        =   7
         Top             =   1110
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
         Height          =   495
         Left            =   11265
         TabIndex        =   6
         Top             =   1110
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
         Height          =   495
         Left            =   12480
         TabIndex        =   5
         Top             =   1110
         Width           =   1215
      End
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
         Left            =   3105
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   3915
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   105
         TabIndex        =   17
         Top             =   480
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
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
         Format          =   64552961
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1650
         TabIndex        =   18
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Format          =   64552961
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label10 
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
         Left            =   7935
         TabIndex        =   24
         Top             =   1290
         Width           =   2280
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "--"
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
         Left            =   10335
         TabIndex        =   23
         Top             =   1290
         Width           =   165
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   1328
         TabIndex        =   22
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   375
         TabIndex        =   20
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   1935
         TabIndex        =   19
         Top             =   210
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "CxTF No"
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
         Left            =   6435
         TabIndex        =   15
         Top             =   975
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type"
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
         Left            =   4590
         TabIndex        =   14
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   12375
         TabIndex        =   12
         Top             =   210
         Width           =   1425
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   8670
         TabIndex        =   11
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   4635
         TabIndex        =   10
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7875
      Left            =   120
      TabIndex        =   0
      Top             =   1905
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   7485
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12726
         _Version        =   393216
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_LotWiseDepositWithdrawalDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCommodityID, mExchangeTypeID As Variant

Private Sub CmbCommodity_GotFocus()
tmp = CmbCommodity.Text
Call TableMst_Commodity
CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
 CmbCommodity.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp
End Sub


Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodity.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID
End Sub


Private Sub CmbExchangeType_GotFocus()
tmp = CmbExchangeType.Text
Call TableMst_ExchangeType
CmbExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeType.Text = tmp
End Sub
'----
Private Sub CmbExchangeType_LostFocus()
If CmbExchangeType.Text = "" Then
   mExchangeTypeID = 0
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeType.SetFocus
   Exit Sub
End If

mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

End Sub

Private Sub CmbLocation_GotFocus()
If mStateID = "" Then Exit Sub

tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("where " & Gen_WcLocation & " And StateID=" & mStateID & "")
End If
CmbLocation.Clear

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location.Fields("LocationID")
mSMPrefix = RsMst_Location.Fields("SM_Prefix")
End Sub

Private Sub CmbState_GotFocus()
tmp = CmbState.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub


Private Sub CmbState_LostFocus()
If CmbState.Text = "" Then
   CmbLocation.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbState.SetFocus
   Exit Sub
End If
If RsMst_State!StateName <> CmbState.Text Then
   CmbLocation.Text = ""
End If

mStateID = RsMst_State!StateID
End Sub
Private Sub CmbType_LostFocus()

If CmbType.Text = "" Then
   Exit Sub
End If

If LCase(CmbType.Text) <> "deposit" And LCase(CmbType.Text) <> "withdrawal" Then
   MsgBox "Invalid Selection"
   CmbType.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdClear_Click()
Call Grid_Head1
CmbState.Text = ""
CmbLocation.Text = ""
CmbExchangeType.Text = ""
TxtCxTFNo.Text = ""
CmdExcel.Enabled = False
TxtFromDate = Date - 30
TxtToDate = Date
CmbType = ""
ProgressBar1.Value = 0
CmbCommodity.Text = ""
Label8.Caption = ""
End Sub
Public Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 9
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 1) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 2) = "GSL ID"
MSHFlexGrid1.TextMatrix(0, 3) = "CxTF No"
MSHFlexGrid1.TextMatrix(0, 4) = "CxTF Date"
MSHFlexGrid1.TextMatrix(0, 5) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 6) = "Demat Weight"
MSHFlexGrid1.TextMatrix(0, 7) = "Type"
MSHFlexGrid1.TextMatrix(0, 8) = "Exchange Type"

MSHFlexGrid1.ColWidth(0) = 2000
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 2000

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "LotWiseDepositWithdrawaldetailReport.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "LotWiseDepositWithdrawaldetailReport.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
wc = ""

Call Grid_Head1
Label8.Caption = ""

If CmbType.Text = "" Then
   MsgBox "Blank Type Not allowed!!!"
   CmbType.SetFocus
   Exit Sub
End If

MsgBox ("Wait till next message")
ProgressBar1.Value = 0

If Gen_DBType = "MDB" Then
   wc = wc & " Where a.CxTFDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
Else
   wc = wc & " Where a.CxTFDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "' "
End If

If CmbState.Text <> "" Then
   wc = wc & " And s.StateName='" & CmbState.Text & "'"
End If

If CmbLocation.Text <> "" Then
   wc = wc & " And l.LocationName='" & CmbLocation.Text & "'"
End If

If CmbExchangeType.Text <> "" Then
   wc = wc & " And e.ExchangeType='" & CmbExchangeType.Text & "'"
End If

If CmbType.Text <> "" Then
   wc = wc & " And a.TxnType='" & UCase(Left(CmbType.Text, 1)) & "'"
End If

If TxtCxTFNo.Text <> "" Then
   wc = wc & " And a.CxTFNo='" & TxtCxTFNo.Text & "'"
End If

If CmbCommodity.Text <> "" Then
   wc = wc & " And com.Commodity='" & CmbCommodity.Text & "'"
End If

tmp = " Select  l.locationname , com.Commodity, b.gslid , a.CxTFNo, a.CxTFDate  , b.NoOfBags , b.DematWeight ,a.TxnType , e.ExchangeType"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_location l on a.sm_prefix = l.sm_prefix"
tmp = tmp & " inner join mst_commodity com on a.commodityid = com.commodityid"
tmp = tmp & " inner join mst_State s on l.StateID = s.StateID"
tmp = tmp & " inner join mst_ExchangeType e on a.ExchangeTypeID = e.ExchangeTypeID"
tmp = tmp & wc
tmp = tmp & " order by l.locationname , com.Commodity, a.CxTFDate"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount + 2
   For i = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For c = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid1.TextMatrix(i, 4) = Format(MSHFlexGrid1.TextMatrix(i, 4), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(i, 7) = IIf(MSHFlexGrid1.TextMatrix(i, 7) = "D", "Deposit", "Withdrawal")
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   Label8.Caption = TmpRs.RecordCount
   ProgressBar1.Value = ProgressBar1.Max
   CmdExcel.Enabled = True
   MsgBox ("Done!!!")
Else
   MsgBox "No record found !!!"
   Call Grid_Head1
   Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
CmdClear_Click
End Sub



