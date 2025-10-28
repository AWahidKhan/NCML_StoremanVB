VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_UnMappedCITF 
   Caption         =   "Goods withdrawn without collection of SC"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbPrintFlag 
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
         ItemData        =   "FrmRpt_UnMappedCITF.frx":0000
         Left            =   7020
         List            =   "FrmRpt_UnMappedCITF.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1125
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.ComboBox CmbCMP 
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
         Left            =   10980
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   405
         Width           =   3900
      End
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
         Left            =   1755
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1125
         Width           =   5160
      End
      Begin VB.ComboBox CmbExchange 
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
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1125
         Visible         =   0   'False
         Width           =   2760
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
         Left            =   4635
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   405
         Width           =   2820
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
         Left            =   7515
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   405
         Width           =   3360
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
         Left            =   13830
         TabIndex        =   15
         Top             =   1110
         Width           =   1095
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Go"
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
         Left            =   11880
         TabIndex        =   10
         Top             =   1110
         Width           =   975
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
         Left            =   12855
         TabIndex        =   14
         Top             =   1110
         Width           =   975
      End
      Begin VB.OptionButton OptTxDate 
         Caption         =   "Transaction Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton OptCrDate 
         Caption         =   "Creation Date"
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
         Left            =   90
         TabIndex        =   13
         Top             =   780
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   420
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
         Format          =   111542273
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   3225
         TabIndex        =   3
         Top             =   420
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
         Format          =   111542273
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Print Flag"
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
         Left            =   7380
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   12713
         TabIndex        =   24
         Top             =   120
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3825
         TabIndex        =   23
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Exchange"
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
         Left            =   1800
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   900
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
         Left            =   5513
         TabIndex        =   21
         Top             =   120
         Width           =   1065
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
         Left            =   1965
         TabIndex        =   20
         Top             =   120
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
         Left            =   3465
         TabIndex        =   19
         Top             =   120
         Width           =   870
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
         Left            =   8813
         TabIndex        =   18
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lblRecordCount 
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
         Left            =   10080
         TabIndex        =   17
         Top             =   1680
         Width           =   210
      End
      Begin VB.Label lblRec 
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
         Left            =   7680
         TabIndex        =   16
         Top             =   1680
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   2085
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   7200
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12091
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
Attribute VB_Name = "FrmRpt_UnMappedCITF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mFlag, mLocationID, mCommodityID, mStateID, mExchangeID, mCMPID As Variant

Private Sub CmbCMP_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
tmp = CmbCMP.Text
wc = ""
If Gen_WcCMP <> "" Then
   wc = "And " & Gen_WcCMP
End If

Call TableMst_CMP(" Where CloseDate is Null And LocationID = " & mLocationID & " " & wc)
CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
   CmbCMP.AddItem RsMst_CMP!CMPName
   RsMst_CMP.MoveNext
Next
CmbCMP.Text = tmp

End Sub

Private Sub CmbCMP_LostFocus()
If CmbCMP.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbCMP.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbCommodity_GotFocus()
tmp = CmbCommodity.Text
Call TableMst_Commodity
CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
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
   MsgBox "Invalid selection!!!"
   CmbCommodity.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID
End Sub

Private Sub CmbExchange_GotFocus()
tmp = CmbExchange.Text
Call TableMst_ExchangeType("")

CmbExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Exchange found"
   CmbExchange.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchange.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchange.Text = tmp

End Sub

Private Sub CmbExchange_LostFocus()
If CmbExchange.Text = "" Then
   mExchangeID = 0
   Exit Sub
End If
' Code for clearing Godown combox in case of change

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid Exchange selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mExchangeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
If CmbStateID.Text = "" Then
   MsgBox "Select State!!!"
   CmbStateID.SetFocus
   Exit Sub
End If
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where StateID = " & mStateID & "  And " & Gen_WcLocation)
Else
   Call TableMst_Location(" Where StateID = " & mStateID & " ")
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
' Code for clearing Godown combox in case of change

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbPrintFlag_LostFocus()
If CmbPrintFlag.Text = "" Then Exit Sub

If LCase(CmbPrintFlag.Text) = "pending" Then
   mFlag = "Pr"
ElseIf LCase(CmbPrintFlag.Text) = "printed" Then
   mFlag = "P"
Else
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" Where " & Gen_WcState)
Else
   Call TableMst_State("")
End If

CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No State found"
   CmbStateID.SetFocus
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
   CmbLocationID.Clear
   CmbCMP.Clear
   mStateID = 0
   Exit Sub
End If
' Code for clearing Godown combox in case of change

RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid State Selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
TxtFromDate = Date - 30
TxtToDate = Date
CmbLocationID.Text = ""
CmbPrintFlag.Text = ""
CmbCMP.Text = ""
CmbCommodity.Text = ""
CmbExchange.Text = ""
CmbStateID.Text = ""
lblRecordCount.Caption = "-"
ProgressBar1.Value = 0
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 21
    .Font.Size = 10
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "State"
    .TextMatrix(0, 1) = "Location"
    .TextMatrix(0, 2) = "CMP"
    .TextMatrix(0, 3) = "Exchange"
    .TextMatrix(0, 4) = "Commodity"
    ''============Remove Agent Column
        '.TextMatrix(0, 5) = "Agent"
    ''============Remove Agent Column
    .TextMatrix(0, 5) = "Name of the Clients in this CITF"
    .TextMatrix(0, 6) = "CITF No" 'NCDEX / Non - NCDEX
    .TextMatrix(0, 7) = "CITF Date"
    
    ''============Remove Following Column
    '.TextMatrix(0, 8) = "Created Date"
    '.TextMatrix(0, 9) = "Created By"
    '.TextMatrix(0, 10) = "Print Flag"
    ''============Remove Following Column
    .TextMatrix(0, 8) = "Reservation Qty Withdrawn" ' "Total CITF Qty Withdrawn"
    .TextMatrix(0, 9) = "Unreserved Qty Withdrawn" '"Reservation Qty Withdrawn"
    .TextMatrix(0, 10) = "Total CITF Qty Withdrawn" '"Unreserved Qty Withdrawn"
    .TextMatrix(0, 11) = "Rejected"
    .TextMatrix(0, 12) = "Area Manager"
    .TextMatrix(0, 13) = "Operational Controller"
    .TextMatrix(0, 14) = "CMPM"
    .TextMatrix(0, 15) = "Storage Charges to be Collected" '"BOE Amount"
    .TextMatrix(0, 16) = "Latest Rate of Commodity"
    .TextMatrix(0, 17) = "Value of Commodity"
    .TextMatrix(0, 18) = "Balance Stock held by the clients of this CITF(in MT)"
    .TextMatrix(0, 19) = "Balance Value"
    .TextMatrix(0, 20) = "SCM Withdrawal"
'    .TextMatrix(0, 16) = "Total CITF Qty"
'    .TextMatrix(0, 17) = "Reservation Qty"
'    .TextMatrix(0, 18) = "Unreserved Qty"
    ''============Remove Following Column
    '.TextMatrix(0, 12) = "Hologram No"
    ''============Remove Following Column
    
    .ColWidth(0) = 1400
    .ColWidth(1) = 1700
    .ColWidth(2) = 2200
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 2200
    .ColWidth(6) = 1000
    .ColWidth(7) = 1200
    .ColWidth(8) = 1500
    .ColWidth(9) = 1500
    .ColWidth(10) = 1500
    .ColWidth(11) = 1500
    .ColWidth(12) = 2200
    .ColWidth(13) = 2200
    .ColWidth(14) = 2200
    .ColWidth(15) = 1700
    .ColWidth(16) = 1400
    .ColWidth(17) = 1200
    .ColWidth(18) = 2500
    .ColWidth(19) = 2500
    .ColWidth(20) = 2000
    
    .RowHeight(0) = 600
    .WordWrap = True
End With
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_UnmappedCITF.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_UnmappedCITF.xls")
Set oWS = oWB.Worksheets("Sheet1")

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1

oWS.ClearArrows
mRow = 1
oWS.Range("a:a").ColumnWidth = 13
oWS.Range("b:b").ColumnWidth = 13
oWS.Range("c:c").ColumnWidth = 15
oWS.Range("d:d").ColumnWidth = 10.43
oWS.Range("e:e").ColumnWidth = 15
oWS.Range("f:f").ColumnWidth = 36.71
oWS.Range("g:g").ColumnWidth = 7
oWS.Range("h:h").ColumnWidth = 9.29
oWS.Range("i:i").ColumnWidth = 11
oWS.Range("j:j").ColumnWidth = 11
oWS.Range("k:k").ColumnWidth = 11
oWS.Range("l:l").ColumnWidth = 11.86
oWS.Range("m:m").ColumnWidth = 30
oWS.Range("n:n").ColumnWidth = 30
oWS.Range("o:o").ColumnWidth = 30
oWS.Range("p:p").ColumnWidth = 9
oWS.Range("q:q").ColumnWidth = 12.43
oWS.Range("r:r").ColumnWidth = 10.57
oWS.Range("s:s").ColumnWidth = 8.43
oWS.Range("t:t").ColumnWidth = 12.43

tmp = "a" & mRow & ":" & "t" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, C + 1) = "Unmapped CITF"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, C + 1).WrapText = True

For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)

        If I = 0 Then
           oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, C + 1).Font.Bold = True
           oWS.Cells(mRow, C + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, C + 1).WrapText = True
        oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True


'oWS.ClearArrows
'mRow = 1
'ProgressBar1.Value = 0
'ProgressBar1.Max = MSHFlexGrid1.Rows + 1
'For I = 0 To MSHFlexGrid1.Rows - 1
'    mRow = I + 1 ' + 6
'    For c = 0 To MSHFlexGrid1.Cols - 1
'        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
'           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
'              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "mm/dd/yyyy")
'           Else
'              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
'           End If
'        Else
'           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
'        End If
'    Next
'    ProgressBar1.Value = ProgressBar1.Value + 1
'Next
'ProgressBar1.Value = ProgressBar1.Max
'oWB.Save
'Set oWB = Nothing
'MsgBox ("Excel File Created !!! ")
'oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head
lblRecordCount.Caption = ""
ProgressBar1.Value = 0

MsgBox " Wait till next message !!!"

tmp = "Select MS.StateName,ML.LocationName,MC.CMPName,ME.ExchangeType,MCom.Commodity, TCD.CxTFNo, TCD.SM_Prefix, TCD.TxnType,"
tmp = tmp & " TCD.CxTFDate,TCD.CreatedDate,MU.EmployeeName+ '~'+  TCD.CreatedBy 'CreatedBy',TCD.BookedSpaceFlag,"
tmp = tmp & " TCD.Flag,TCD.HologramNo, MS.StateID, MCom.CommodityID,TCD.SCMFlag,"
tmp = tmp & " sum(TCB.BOEAmount) BOEAmount "
''===========Remove Column Agent
'MA.AgentName,
''===========Remove Column Agent
tmp = tmp & " From Txn_CxTF_Details TCD"
tmp = tmp & " Left Join Txn_CxTF_BOE TCB on TCB.SM_Prefix = TCD.SM_Prefix And TCB.CxTFNo = TCD.CxTFNo"
tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_Commodity MCom on TCD.CommodityID = MCom.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TCD.ExchangeTypeID = ME.ExchangeTypeID"
''===========Remove Column Agent
'tmp = tmp & " Left Join  Mst_Agent MA On TCD.AgentID = MA.AgentID"
''===========Remove Column Agent
tmp = tmp & " Inner Join Mst_Users MU On TCD.CreatedBy = MU.EmployeeNo"
tmp = tmp & " Where TCD.TxnType = 'W' And ME.ExchangeTypeID = 1 And TCD.CxTFNo not in (Select CITFNo From  Txn_CITFCRDetail TCCD Where TCCD.SM_Prefix = ML.SM_Prefix)"
tmp = tmp & " And TCD.Flag='P'"

If OptTxDate.Value = True Then
   If Gen_DBType = "MDB" Then
      tmp = tmp & " And (TCD.CxTFDate between #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#)"
   Else
      tmp = tmp & " And TCD.CxTFDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' and '" & Format(TxtToDate.Value, Gen_DatFormat) & "'" '''   tmp = tmp & " And (CDTFDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'  And '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
   End If
ElseIf OptCrDate.Value = True Then
   If Gen_DBType = "MDB" Then
      tmp = tmp & " And (TCD.CreatedDate between #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#)"
   Else
      tmp = tmp & " And TCD.CreatedDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' and '" & Format(TxtToDate.Value, Gen_DatFormat) & "'" '''   tmp = tmp & " And (CDTFDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'  And '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
   End If
End If

If CmbStateID.Text <> "" Then
   tmp = tmp & " And MS.StateID=" & mStateID & ""
Else
   If Gen_WcState <> "" Then
      tmp = tmp & " And ML." & Gen_WcState
   End If
End If

If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID & ""
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

If CmbCMP.Text <> "" Then
   tmp = tmp & " And MC.CMPID=" & mCMPID & ""
Else
   If Gen_WcCMP <> "" Then
      tmp = tmp & "MC." & Gen_WcCMP
   End If
End If

If CmbCommodity.Text <> "" Then
   tmp = tmp & " And MCom.CommodityID=" & mCommodityID & ""
End If

'If CmbExchange.Text <> "" Then
'   tmp = tmp & " And ME.ExchangeTypeID=" & mExchangeID & ""
'End If

'If CmbPrintFlag.Text <> "" Then
'   If mFlag = "P" Then
'      tmp = tmp & " And TCD.Flag='P'"
'   ElseIf mFlag = "Pr" Then
'      tmp = tmp & " And (TCD.Flag is null Or TCD.Flag = '') "
'   End If
'End If
tmp = tmp & " group by MS.StateName,ML.LocationName,MC.CMPName,ME.ExchangeType,MCom.Commodity, TCD.CxTFNo, TCD.SM_Prefix,"
tmp = tmp & " TCD.TxnType, TCD.CxTFDate,TCD.CreatedDate,MU.EmployeeName, TCD.CreatedBy , TCD.BookedSpaceFlag,"
tmp = tmp & " TCD.Flag , TCD.HologramNo, ms.StateID, MCom.CommodityID,TCD.SCMFlag"
tmp = tmp & " having sum(TCB.BOEAmount) > 0"
tmp = tmp & " Order by MS.StateName,ML.LocationName,MC.CMPName "

Call TmpTable

MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 2
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) 'Format(TmpRs!BillTxnDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName) 'Format(TmpRs!MonthYear, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       'MSHFlexGrid1.TextMatrix(I, 5) = "name of the clients"   '"IIf(IsNull(TmpRs!AgentName), "", TmpRs!AgentName)
       Call GetClientName1(I)
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!CxTFNO), "", TmpRs!CxTFNO) 'Format(TmpRs!BillDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!CxTFDate), "", Format(TmpRs!CxTFDate, "dd-MMM-yyyy"))
       'MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!CreatedDate), "", Format(TmpRs!CreatedDate, "dd-MMM-yyyy"))
       'MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy)
       'MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag)
       MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!BookedSpaceFlag), "", TmpRs!BookedSpaceFlag)
       'MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!HologramNo), "", TmpRs!HologramNo)
       'If MSHFlexGrid1.TextMatrix(I, 10) = "P" Then
       '    MSHFlexGrid1.TextMatrix(I, 10) = "Printed"
       'ElseIf MSHFlexGrid1.TextMatrix(I, 10) = "" Then
       '    MSHFlexGrid1.TextMatrix(I, 10) = "Pending"
       'End If
       If MSHFlexGrid1.TextMatrix(I, 11) = "0" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Not Rejected"
       ElseIf MSHFlexGrid1.TextMatrix(I, 11) = "1" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Rejected"
       End If
       Call GetAMOCName(I)
       'Call GetBOEAmount(I)
       MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs!BOEAmount), 0, TmpRs!BOEAmount)
       Call GetTotalCITFQuanity(I)
       Call GetCommodityLatestRate(I)
       Call GetReservationQty(I)
       Call GetUnReservationQty(I)
       Call GetTotalCITFQuanity(I)
       Call GetBalanceStockbyClient(I)
       MSHFlexGrid1.TextMatrix(I, 17) = Val(MSHFlexGrid1.TextMatrix(I, 10)) * Val(MSHFlexGrid1.TextMatrix(I, 16))
       MSHFlexGrid1.TextMatrix(I, 19) = Val(MSHFlexGrid1.TextMatrix(I, 18)) * Val(MSHFlexGrid1.TextMatrix(I, 16))
       
       MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs!SCMFlag), 0, TmpRs!SCMFlag)
       If MSHFlexGrid1.TextMatrix(I, 20) = "0" Then
          MSHFlexGrid1.TextMatrix(I, 20) = "No"
       ElseIf MSHFlexGrid1.TextMatrix(I, 20) = "1" Then
          MSHFlexGrid1.TextMatrix(I, 20) = "Yes"
       End If
       
       
       
       'MSHFlexGrid1.TextMatrix(I, 14) = "Value of Commodity"
       
       'MSHFlexGrid1.TextMatrix(I, 15) = "Balance Stock held by the clients of this CITF"
       
       'MSHFlexGrid1.TextMatrix(I, 16) = "Total CITF Qty"
       'MSHFlexGrid1.TextMatrix(I, 17) = "Reservation Qty"
       'MSHFlexGrid1.TextMatrix(I, 18) = "Unreserved Qty"
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox ("Done!!!")
   lblRecordCount.Caption = TmpRs.RecordCount
Else
   MsgBox "No Record Found !!!"
End If
End Sub

Private Sub GetBalanceStockbyClient(x_ As Variant)
tmp = " Select sum(MGSL.BalanceWeight) BalanceWeight from Mst_GSL MGSL"
tmp = tmp & " Where MGSL.ClientIDRow in (Select"
tmp = tmp & " Distinct TWD.ClientIDRow"
tmp = tmp & " From TxN_WHR_Detail TWD"
tmp = tmp & " Inner Join Txn_Cxtf_ClientDetail TCCD on TCCD.RRNID = TWD.WHRNo"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD on TCD.SM_Prefix = TCCD.SM_Prefix And TCD.CxTFNo = TCCD.CxTFNo And TCD.TxnType = TCCD.TxnType"
tmp = tmp & " Inner Join Mst_CLient MCL on MCL.ClientIDRow = TWD.ClientIDRow"
tmp = tmp & " Where TCCD.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCCD.CxTFNo = '" & TmpRs!CxTFNO & "' And TCCD.TxnType = 'W'"
tmp = tmp & " )"
tmp = tmp & " And MGSL.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
tmp = tmp & " And CommodityID = " & TmpRs!CommodityID & ""
Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
    MSHFlexGrid1.TextMatrix(x_, 18) = IIf(IsNull(TmpRs2!BalanceWeight), 0, TmpRs2!BalanceWeight)
    If Val(MSHFlexGrid1.TextMatrix(x_, 18)) < 0 Then
        MSHFlexGrid1.TextMatrix(x_, 18) = 0
    End If
End If

End Sub


Private Sub GetUnReservationQty(x_ As Variant)
tmp = " Select sum(TGSL.DematWeight) DematWeight From Txn_CXTF_GSL TGSL"
tmp = tmp & " Inner Join Mst_GSL MGSL on TGSL.GSLID = MGSL.GSLID And TGSL.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Where TGSL.CxTFNo = " & TmpRs!CxTFNO & " And TGSL.SM_Prefix = '" & TmpRs!SM_Prefix & "' and TGSL.TxnType = 'W'"
tmp = tmp & " And (MGSL.NContractID is null or  MGSL.NContractID = 0)"
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
    MSHFlexGrid1.TextMatrix(x_, 9) = IIf(IsNull(TmpRs2!DematWeight), 0, TmpRs2!DematWeight)
End If
End Sub

Private Sub GetReservationQty(x_ As Variant)
tmp = " Select sum(TGSL.DematWeight) DematWeight From Txn_CXTF_GSL TGSL"
tmp = tmp & " Inner Join Mst_GSL MGSL on TGSL.GSLID = MGSL.GSLID And TGSL.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Where TGSL.CxTFNo = " & TmpRs!CxTFNO & " And TGSL.SM_Prefix = '" & TmpRs!SM_Prefix & "' and TGSL.TxnType = 'W'"
tmp = tmp & " And (MGSL.NContractID is not null And  MGSL.NContractID <> 0)"
Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
    MSHFlexGrid1.TextMatrix(x_, 8) = IIf(IsNull(TmpRs2!DematWeight), 0, TmpRs2!DematWeight)
End If
End Sub

Private Sub GetTotalCITFQuanity(x_ As Variant)
tmp = " Select sum(DematWeight) DematWeight  From Txn_CXTF_GSL"
tmp = tmp & " Where CxTFNo = " & TmpRs!CxTFNO & " And SM_Prefix = '" & TmpRs!SM_Prefix & "' and TxnType = 'W'"
Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
    MSHFlexGrid1.TextMatrix(x_, 10) = IIf(IsNull(TmpRs2!DematWeight), 0, TmpRs2!DematWeight)
End If
End Sub

Private Sub GetCommodityLatestRate(x_ As Variant)
'tmp = " Select MCP.YearMonth, MCP.Day28,MCP.Day29,MCP.Day30,MCP.Day31 from Mst_CommodityPrice MCP" 'MS.StateName,MC.Commodity,MCP.MonthYear,MCP.YearMonth,
''tmp = tmp & " Inner Join Mst_Commodity MC on MCP.CommodityID = MC.CommodityID"
''tmp = tmp & " Inner Join Mst_State MS On MCP.StateID = MS.StateID"
''tmp = tmp & " Where YearMonth =  '" & Left(MonthName(Month((TxtToDate.Value))), 3) & "-" & Year(TxtToDate.Value) & "'"
''tmp = tmp & " And MS.StateName = '" & MSHFlexGrid1.TextMatrix(x_, 0) & "' And MC.Commodity = '" & MSHFlexGrid1.TextMatrix(x_, 3) & "' "
'tmp = tmp & " Where MCP.StateID = " & TmpRs!StateID & " And MCP.CommodityID = " & TmpRs!CommodityID & ""
'tmp = tmp & " Order By MCP.YearMonth "


tmp = " Select MCP.StateID,MCP.CommodityID,MCP.YearMonth,"
tmp = tmp & " Case When Convert(date,MCP.YearMonth+'-01') = Convert(date,Convert(Varchar,DATEPART(Year,GETDATE()))+'-'+Convert(Varchar,DATEPART(Month,GETDATE()))+'-01')"
tmp = tmp & " Then"
tmp = tmp & "   COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & " Else"
tmp = tmp & "   COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & " End [Rate]"
tmp = tmp & " from Mst_CommodityPrice MCP"
tmp = tmp & " Where StateID = " & TmpRs!StateID & " And CommodityID = " & mCommodityID_ & "  And Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,YearMonth) In  (Select Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,MAX(YearMonth)) From Mst_CommodityPrice Group By StateID,CommodityID )"

Call Tmp2Table

MSHFlexGrid1.TextMatrix(x_, 16) = 0

'If TmpRs2.RecordCount > 0 Then
'   TmpRs2.MoveLast
'   'If Month(TxtToDate.Value) = 1 Or Month(TxtToDate.Value) = 3 Or Month(TxtToDate.Value) = 5 Or Month(TxtToDate.Value) = 7 Or Month(TxtToDate.Value) = 8 Or Month(TxtToDate.Value) = 10 Or Month(TxtToDate.Value) = 12 Then
'   If Right(TmpRs2!YearMonth, 2) = "01" Or Right(TmpRs2!YearMonth, 2) = "03" Or Right(TmpRs2!YearMonth, 2) = "05" Or Right(TmpRs2!YearMonth, 2) = "07" Or Right(TmpRs2!YearMonth, 2) = "08" Or Right(TmpRs2!YearMonth, 2) = "10" Or Right(TmpRs2!YearMonth, 2) = "12" Then
'      MSHFlexGrid1.TextMatrix(x_, 16) = TmpRs2!Day31
'   ElseIf Right(TmpRs2!YearMonth, 2) = "04" Or Right(TmpRs2!YearMonth, 2) = "06" Or Right(TmpRs2!YearMonth, 2) = "09" Or Right(TmpRs2!YearMonth, 2) = "11" Then
'      MSHFlexGrid1.TextMatrix(x_, 16) = TmpRs2!Day30
'   Else
'      If IsNull(TmpRs2!Day29) = False Then
'         MSHFlexGrid1.TextMatrix(x_, 16) = TmpRs2!Day29
'      Else
'         MSHFlexGrid1.TextMatrix(x_, 16) = TmpRs2!Day28
'      End If
'   End If
'Else
'   tmp = " Select MinRate from Mst_Commodity Where CommodityID = " & TmpRs!CommodityID & ""
'
'   Call Tmp2Table
'
'   If TmpRs2.RecordCount > 0 Then
'      MSHFlexGrid1.TextMatrix(x_, 16) = TmpRs2!MinRate
'   End If
'End If


If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x_, 16) = IIf(IsNull(TmpRs2!Rate), 0, TmpRs2!Rate)
Else
   tmp = " Select MinRate from Mst_Commodity Where CommodityID = " & TmpRs!CommodityID & ""
   Call Tmp2Table

   If TmpRs2.RecordCount > 0 Then
      MSHFlexGrid1.TextMatrix(x_, 16) = TmpRs2!MinRate
   End If
End If

'tmp = " Select * From Mst_CommodityPrice"
'tmp = tmp & " Where StateID = " & TmpRs!StateID & " And CommodityID = " & TmpRs!CommodityID & ""
'tmp = tmp & " And YearMonth = (Select Max(YearMonth) from Mst_CommodityPrice Where StateID = " & TmpRs!StateID & " And CommodityID = " & TmpRs!CommodityID & ")"
'Call Tmp2Table
'
'If TmpRs2.RecordCount > 0 Then
'    For k = 31 To 1 Step -1
'        If IsNull(TmpRs2!Day & k) = False And Val(TmpRs2!Day & k) <> 0 Then
'            MSHFlexGrid1.TextMatrix(I, 13) = TmpRs2!Day & k
'            Exit For
'        End If
'    Next
'End If
End Sub

Private Sub GetBOEAmount(x_ As Variant)
tmp = " Select sum(TCB.BOEAmount) BOEAmount"
tmp = tmp & " From Txn_CxTF_BOE TCB"
tmp = tmp & " Where TCB.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCB.CxTFNo = '" & TmpRs!CxTFNO & "'"
Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
    MSHFlexGrid1.TextMatrix(x_, 15) = IIf(IsNull(TmpRs2!BOEAmount), 0, TmpRs2!BOEAmount)
End If

End Sub

Private Sub GetClientName1(x_ As Variant)
tmp = " Select Distinct MCL.ClientName, TWD.ClientIDRow From TxN_WHR_Detail TWD"
tmp = tmp & " Inner Join Txn_Cxtf_ClientDetail TCCD on TCCD.RRNID = TWD.WHRNo"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD on TCD.SM_Prefix = TCCD.SM_Prefix And TCD.CxTFNo = TCCD.CxTFNo And TCD.TxnType = TCCD.TxnType"
tmp = tmp & " Inner Join Mst_CLient MCL on MCL.ClientIDRow = TWD.ClientIDRow"
tmp = tmp & " Where TCCD.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCCD.CxTFNo = '" & TmpRs!CxTFNO & "' And TCCD.TxnType = 'W'"
Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
    For k = 1 To TmpRs2.RecordCount
        If MSHFlexGrid1.TextMatrix(x_, 5) = "" Then
            MSHFlexGrid1.TextMatrix(x_, 5) = TmpRs2!ClientName
        Else
            MSHFlexGrid1.TextMatrix(x_, 5) = TmpRs2!ClientName & "~" & MSHFlexGrid1.TextMatrix(x_, 5)
        End If
        TmpRs2.MoveNext
    Next
End If
End Sub

Private Sub GetAMOCName(x_ As Variant)

'tmp = " Select Distinct  ME1.EmployeeName + '~'+ ME1.EmployeeNo 'AMName',ME2.EmployeeName + '~'+ ME2.EmployeeNo 'OCName'"
'tmp = tmp & " From Mst_EmpCMPMapping MECM"
'tmp = tmp & " Inner Join Mst_CMP MC On MECM.CMPID = MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML On MC.LocatioNID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_Employee ME On MECM.EmployeeID = ME.EmployeeID"
'tmp = tmp & " Inner Join Mst_Employee ME1 On ME.DirectReportingTo = ME1.EmployeeID"
'tmp = tmp & " Inner Join Mst_Employee ME2 On ME1.EmployeeID = ME2.DirectReportingTo And ME2.DesignationID = 10"
'tmp = tmp & " Where ME.Flag = 'Y' And ML.LocationName = '" & MSHFlexGrid1.TextMatrix(x_, 1) & "'"
'
'Call Tmp1Table

'If TmpRs1.RecordCount > 0 Then
'   TmpRs1.MoveFirst
'   TmpRs1.Find "LocationName = '" & MSHFlexGrid1.TextMatrix(x_, 1) & "' "
'   If TmpRs1.EOF = False Then
'      MSHFlexGrid1.TextMatrix(x_, 9) = IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname) ' & "~" & IIf(IsNull(TmpRs!Employee_Name), "", TmpRs!Employee_Name)
'      MSHFlexGrid1.TextMatrix(x_, 10) = IIf(IsNull(TmpRs1!OCName), "", TmpRs1!OCName) ' & "~" & IIf(IsNull(TmpRs!OCName), "", TmpRs!OCName)
'   End If
'End If


tmp = " Select Distinct  ML.LocationName, ME.EmployeeName + '~'+ ME.EmployeeNo 'CMPMName', ME1.EmployeeName + '~'+ ME1.EmployeeNo 'AMName',ME2.EmployeeName + '~'+ ME2.EmployeeNo 'OCName'"
tmp = tmp & " From Mst_EmpCMPMapping MECM"
tmp = tmp & " Inner Join Mst_CMP MC On MECM.CMPID = MC.CMPID"
tmp = tmp & " inner join Mst_Location ML On MC.LocatioNID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Employee ME On MECM.EmployeeID = ME.EmployeeID"
tmp = tmp & " Inner Join Mst_Employee ME1 On ME.DirectReportingTo = ME1.EmployeeID"
tmp = tmp & " Inner Join Mst_Employee ME2 On ME1.EmployeeID = ME2.DirectReportingTo And ME2.DesignationID = 10"
tmp = tmp & " Where ME.Flag = 'Y' And ML.LocationName = '" & MSHFlexGrid1.TextMatrix(x_, 1) & "'"

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
    For j = 1 To TmpRs2.RecordCount
        If MSHFlexGrid1.TextMatrix(x_, 12) = "" Then
            MSHFlexGrid1.TextMatrix(x_, 12) = IIf(IsNull(TmpRs2!AMname), "", TmpRs2!AMname)
        Else
            If InStr(1, MSHFlexGrid1.TextMatrix(x_, 12), IIf(IsNull(TmpRs2!AMname), "", TmpRs2!AMname)) <= 0 Then
                MSHFlexGrid1.TextMatrix(x_, 12) = IIf(IsNull(TmpRs2!AMname), "", TmpRs2!AMname) & " , " & MSHFlexGrid1.TextMatrix(x_, 12)
            End If
        End If
        If MSHFlexGrid1.TextMatrix(x_, 13) = "" Then
            MSHFlexGrid1.TextMatrix(x_, 13) = IIf(IsNull(TmpRs2!OCName), "", TmpRs2!OCName)
        Else
            If InStr(1, MSHFlexGrid1.TextMatrix(x_, 13), IIf(IsNull(TmpRs2!OCName), "", TmpRs2!OCName)) <= 0 Then
                MSHFlexGrid1.TextMatrix(x_, 13) = IIf(IsNull(TmpRs2!OCName), "", TmpRs2!OCName) & " , " & MSHFlexGrid1.TextMatrix(x_, 13)
            End If
        End If
        'MSHFlexGrid1.TextMatrix(x_, 12) = IIf(IsNull(TmpRs2!AMname), "", TmpRs2!AMname) ' & "~" & IIf(IsNull(TmpRs!Employee_Name), "", TmpRs!Employee_Name)
        'MSHFlexGrid1.TextMatrix(x_, 13) = IIf(IsNull(TmpRs2!OCName), "", TmpRs2!OCName) ' & "~" & IIf(IsNull(TmpRs!OCName), "", TmpRs!OCName)
        If MSHFlexGrid1.TextMatrix(x_, 14) = "" Then
            MSHFlexGrid1.TextMatrix(x_, 14) = IIf(IsNull(TmpRs2!CMPMName), "", TmpRs2!CMPMName)
        Else
            'If Not InStr(1, .TextMatrix(m, 10), IIf(IsNull(TmpRsStoreMan!CMPM), "", TmpRsStoreMan!CMPM)) Then
            '        .TextMatrix(m, 9) = IIf(IsNull(TmpRsStoreMan!CMPM), "", TmpRsStoreMan!CMPM) & " , " & .TextMatrix(m, 9)  '"CMPM" '"CMPM"
            'End If
            If InStr(1, MSHFlexGrid1.TextMatrix(x_, 14), IIf(IsNull(TmpRs2!CMPMName), "", TmpRs2!CMPMName)) <= 0 Then
                MSHFlexGrid1.TextMatrix(x_, 14) = IIf(IsNull(TmpRs2!CMPMName), "", TmpRs2!CMPMName) & " , " & MSHFlexGrid1.TextMatrix(x_, 14)  '"CMPM" '"CMPM"
            End If
        End If
        TmpRs2.MoveNext
    Next
End If
End Sub

Private Sub Form_Load()
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True

Call CmdClear_Click
CmdExcel.Enabled = True

End Sub

