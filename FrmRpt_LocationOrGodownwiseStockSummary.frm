VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_LocationOrGodownwiseStockSummary 
   Caption         =   "Location Or Godown Wise Stock Summary"
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
      Height          =   1725
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   15000
      Begin VB.OptionButton Opt_LocationWise 
         Caption         =   "Location Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Opt_GodownWise 
         Caption         =   "Godown Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
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
         ItemData        =   "FrmRpt_LocationOrGodownwiseStockSummary.frx":0000
         Left            =   11265
         List            =   "FrmRpt_LocationOrGodownwiseStockSummary.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   570
         Width           =   2520
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
         Left            =   7080
         TabIndex        =   3
         Top             =   570
         Width           =   3975
      End
      Begin VB.ComboBox CmbGodown 
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
         Left            =   7080
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1290
         Visible         =   0   'False
         Width           =   4005
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1290
         Visible         =   0   'False
         Width           =   6855
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
         Left            =   4440
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   570
         Width           =   2535
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
         Left            =   13695
         TabIndex        =   9
         Top             =   1275
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
         Left            =   11265
         TabIndex        =   7
         Top             =   1275
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
         Left            =   12480
         TabIndex        =   8
         Top             =   1275
         Width           =   1215
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
         Left            =   1740
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   570
         Width           =   2655
      End
      Begin VB.Label Label7 
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
         Left            =   11798
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCMP 
         AutoSize        =   -1  'True
         Caption         =   "CMP Name"
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
         Left            =   3030
         TabIndex        =   18
         Top             =   1005
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label LblGodown 
         AutoSize        =   -1  'True
         Caption         =   "Godown"
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
         Left            =   8707
         TabIndex        =   17
         Top             =   1005
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label1 
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
         Left            =   8467
         TabIndex        =   16
         Top             =   210
         Width           =   1200
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
         Left            =   5280
         TabIndex        =   15
         Top             =   233
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   11280
         TabIndex        =   14
         Top             =   960
         Width           =   2325
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Left            =   13680
         TabIndex        =   13
         Top             =   960
         Width           =   75
      End
      Begin VB.Label Label6 
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
         Left            =   2655
         TabIndex        =   12
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7680
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   7215
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   10
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
Attribute VB_Name = "FrmRpt_LocationOrGodownwiseStockSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID, mExchangeTypeID, mGodownID, mCommodityID As Variant 'mGodownID,

Public Sub Grid_Head()
ProgressBar1.Value = 0
Label5.Caption = ""
With MSHFlexGrid1
    .Clear
    .Cols = 6
    .Rows = 3
    .FixedRows = 2
    .FixedCols = 0
    
    .TextMatrix(1, 0) = "State Name"
    .TextMatrix(1, 1) = "Location Name"
    .TextMatrix(1, 2) = "Commodity"
    If Opt_LocationWise.Value = True Then
        .TextMatrix(1, 3) = "Exchange Type"
        .TextMatrix(1, 4) = "Balance Bags"
        .TextMatrix(1, 5) = "Balance Weight"
    ElseIf Opt_GodownWise.Value = True Then
        .Cols = 8
        .TextMatrix(1, 3) = "CMP Name"
        .TextMatrix(1, 4) = "Godown No"
        .TextMatrix(1, 5) = "Exchange Type"
        .TextMatrix(1, 6) = "Balance Bags"
        .TextMatrix(1, 7) = "Balance Weight"
    End If
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 2500
    .ColWidth(2) = 1800
    .ColWidth(3) = 2200
    .ColWidth(4) = 1800
    .ColWidth(5) = 1500
    
    .RowHeight(1) = 700
    .WordWrap = True
End With
End Sub


Private Sub CmbCMPName_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" where LocationID=" & mLocationID & " And CloseDate is null ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "First Select Location!!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub

Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
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
End Sub

Private Sub CmbCommodity_GotFocus()
tmp = CmbCommodity.Text
Call TableMst_Commodity("Order by Commodity")
CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   MsgBox "Invalid selection "
   CmbCommodity.SetFocus
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

Private Sub CmbGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPID & " And CloseDate is null ")
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()
If CmbGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!godownid
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" where " & Gen_WcState & "")
Else
   Call TableMst_State
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

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbLocation.Clear
   CmbCMPName.Clear
   CmbGodown.Clear
   Exit Sub
End If
If CmbStateID.Text <> tmp Then
   CmbLocation.Clear
   CmbCMPName.Clear
   CmbGodown.Clear
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

Private Sub CmbLocation_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State !!"
   CmbStateID.SetFocus
   Exit Sub
End If
tmp = CmbLocation.Text

Call TableMst_Location(" Where LocationID <> 0 And StateID= " & mStateID & "")

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
    CmbCMPName.Clear
    CmbGodown.Clear
    Exit Sub
End If
' Code for clearing Godown combox in case of change
If CmbLocation.Text <> tmp Then
    CmbCMPName.Clear
    CmbGodown.Clear
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdClear_Click()

Label5.Caption = ""
CmdExcel.Enabled = False
CmbLocation.Text = ""
CmbStateID.Text = ""
CmbCommodity.Text = ""
CmbCMPName.Text = ""
CmbGodown.Text = ""
CmbExchangeTypeID.Text = ""
ProgressBar1.Value = 0
Call Grid_Head

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_LocationOrGodownWiseStockSummary.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_LocationOrGodownWiseStockSummary.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
oWS.Range("a:a").ColumnWidth = 15
oWS.Range("b:b").ColumnWidth = 15
oWS.Range("c:c").ColumnWidth = 28
If Opt_GodownWise.Value = True Then
    oWS.Range("d:d").ColumnWidth = 23.14
    oWS.Range("e:e").ColumnWidth = 10.71
    oWS.Range("f:f").ColumnWidth = 13.43
    oWS.Range("g:g").ColumnWidth = 11.57
    oWS.Range("h:h").ColumnWidth = 14.04
    tmp = "a" & mRow & ":" & "h" & mRow
ElseIf Opt_LocationWise.Value = True Then
    oWS.Range("d:d").ColumnWidth = 13.43
    oWS.Range("e:e").ColumnWidth = 11.57
    oWS.Range("f:f").ColumnWidth = 14.04
    tmp = "a" & mRow & ":" & "f" & mRow
End If


oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 1) = "Location Or Godown Wise Stock Summary"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, c + 1).WrapText = True


For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        If I = 1 Then
           oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, c + 1).Font.Bold = True
           oWS.Cells(mRow, c + 1).RowHeight = 30
           oWS.Cells(mRow, c + 1).WrapText = True
        End If
        oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
wc = ""

Call Grid_Head

MsgBox "Please wait till next Message."
ProgressBar1.Value = 0

Label5.Caption = ""

If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "S.StateID", "N", "N")
Else
   If Gen_WcState <> "" Then
      tmp = Replace(Gen_WcState, "StateID", "S.StateID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If

If CmbLocation.Text <> "" Then
   wc = GenWc(wc, mLocationID, "L.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp = Replace(Gen_WcLocation, "LocationID", "L.LocationID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If

If CmbCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "C.CommodityID", "N", "N")
End If

If CmbCMPName.Text <> "" Then
   wc = GenWc(wc, mCMPID, "cmp.cmpID", "N", "N")
End If

If CmbGodown.Text <> "" Then
   wc = GenWc(wc, mGodownID, "g.godownID", "N", "N")
End If

If CmbExchangeTypeID.Text <> "" Then
   wc = GenWc(wc, mExchangeTypeID, "MET.ExchangeTypeID", "N", "N")
End If

tmp = " Select "
tmp = tmp & " S.StateName, l.LocationName , MET.ExchangeType, c.Commodity, sum(gsl.BalanceBags) 'BalanceBags', sum(gsl.BalanceWeight) 'BalanceWeight'"
If Opt_GodownWise.Value = True Then
    tmp = tmp & " ,cmp.cmpname,g.godownno "
End If
'If CmbCMPName.Text <> "" Then
''   If CmbGodown.Text = "" Then
''      tmp = tmp & " ,cmp.cmpname "
''   Else
'      tmp = tmp & " ,cmp.cmpname,g.godownno "
''   End If
'End If

tmp = tmp & " from mst_gsl gsl "
tmp = tmp & " inner join mst_location l on l.sm_prefix = gsl.sm_prefix "
tmp = tmp & " inner join mst_State s on s.StateID = l.stateID "
If Opt_GodownWise.Value = True Then
    tmp = tmp & " left join mst_cmp cmp on cmp.cmpid = gsl.cmpid "
    tmp = tmp & " left join mst_godown g on g.godownid = gsl.godownid "
End If
'If CmbCMPName.Text <> "" Then
'   tmp = tmp & " left join mst_cmp cmp on cmp.cmpid = gsl.cmpid "
'   tmp = tmp & " left join mst_godown g on g.godownid = gsl.godownid "
'End If
'If CmbGodown.Text <> "" Then

'End If
tmp = tmp & " inner join mst_Commodity c on gsl.CommodityID = c.CommodityID "
tmp = tmp & " inner join mst_ExchangeType MET on gsl.ExchangeTypeID = MET.ExchangeTypeID "
If wc = "" Then
    tmp = tmp & " where gsl.flag in ('O','J','R')"
Else
    tmp = tmp & " " & wc & " And gsl.flag in ('O','J','R')"
End If

tmp = tmp & " group by s.StateName, l.LocationName, MET.ExchangeType, c.Commodity"
If Opt_GodownWise.Value = True Then
    tmp = tmp & " , cmp.cmpname, g.godownno "
End If

'If CmbCMPName.Text <> "" Then
''   If CmbGodown.Text <> "" Then
'      tmp = tmp & " ,cmp.cmpname,g.godownno "
''   Else
'      tmp = tmp & " ,cmp.cmpname "
''   End If
'End If

tmp = tmp & " order by s.StateName ,l.LocationName , c.Commodity "
If Opt_GodownWise.Value = True Then
    tmp = tmp & " ,cmp.cmpname,g.godownno "
End If

'If CmbCMPName.Text <> "" Then
''   If CmbGodown.Text <> "" Then
'      tmp = tmp & " ,cmp.cmpname,g.godownno "
''   Else
''      tmp = tmp & " ,cmp.cmpname "
''   End If
'End If

Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount + 2

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       If Opt_LocationWise.Value = True Then
          MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
          MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
          MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
          MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
          MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
       ElseIf Opt_GodownWise.Value = True Then
          MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
          MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
          MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
          MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
          MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
          MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
          MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
       End If

'       If CmbCMPName.Text <> "" Then
' '         If CmbGodown.Text <> "" Then
'             MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
'             MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
'             MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
'             MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
'             MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
'             MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
'             MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
''          Else
''             MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
''             MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
''             MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
''             MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
''            MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
''             MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
' '         End If
'       Else
'          MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
'          MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
'          MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
'          MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
'          MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
'       End If
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
End If

ProgressBar1.Value = ProgressBar1.Max
Label5.Caption = TmpRs.RecordCount
Label5.Refresh
CmdExcel.Enabled = True
MsgBox "Done!!"

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Grid_Head
CmdExcel.Enabled = False
End Sub

Private Sub Opt_GodownWise_Click()
lblCMP.Visible = True
LblGodown.Visible = True
CmbGodown.Visible = True
CmbCMPName.Visible = True
CmbGodown.Text = ""
CmbCMPName.Text = ""
Call Grid_Head
End Sub

Private Sub Opt_LocationWise_Click()
lblCMP.Visible = False
LblGodown.Visible = False
CmbGodown.Visible = False
CmbCMPName.Visible = False
Call Grid_Head
CmbGodown.Text = ""
CmbCMPName.Text = ""
End Sub
