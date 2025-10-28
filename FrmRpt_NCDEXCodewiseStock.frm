VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_NCDEXCodewiseStock 
   Caption         =   "NCDEX Code wise Stock"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   15000
      Begin VB.TextBox TxtSClientName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5760
         TabIndex        =   7
         Top             =   1290
         Width           =   9045
      End
      Begin VB.ComboBox CmbSExchangeTypeID 
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
         ItemData        =   "FrmRpt_NCDEXCodewiseStock.frx":0000
         Left            =   2880
         List            =   "FrmRpt_NCDEXCodewiseStock.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1290
         Width           =   2805
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
         Left            =   5760
         TabIndex        =   3
         Top             =   570
         Width           =   3615
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
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1290
         Width           =   2640
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
         Left            =   9420
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   570
         Width           =   5415
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
         Left            =   2880
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   570
         Width           =   2805
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
         Left            =   13635
         TabIndex        =   12
         Top             =   1755
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
         Left            =   11205
         TabIndex        =   8
         Top             =   1755
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
         Left            =   12420
         TabIndex        =   11
         Top             =   1755
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
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   570
         Width           =   2655
      End
      Begin VB.Label Label8 
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
         Height          =   225
         Left            =   9622
         TabIndex        =   22
         Top             =   1005
         Width           =   1320
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
         Left            =   3517
         TabIndex        =   21
         Top             =   1005
         Width           =   1530
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
         Left            =   11610
         TabIndex        =   19
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label Label2 
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
         Left            =   1125
         TabIndex        =   18
         Top             =   1005
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
         Left            =   6967
         TabIndex        =   17
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
         Left            =   3855
         TabIndex        =   16
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
         Left            =   6960
         TabIndex        =   15
         Top             =   1815
         Width           =   2325
      End
      Begin VB.Label Label5 
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
         Left            =   9360
         TabIndex        =   14
         Top             =   1815
         Width           =   90
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
         Left            =   975
         TabIndex        =   13
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7275
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   9
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   9480
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmRpt_NCDEXCodewiseStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID, mExchangeTypeID, mGodownID, mCommodityID, mSExchangeTypeID As Variant 'mGodownID,

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 14
    .Font.Size = 10
    .AllowUserResizing = flexResizeBoth
    
    .TextMatrix(1, 0) = "NCDEX CODE"
    .TextMatrix(1, 1) = "Commtrack NCDEX Code"
    .TextMatrix(1, 2) = "NCom NCDEX Code"
    .TextMatrix(1, 3) = "State"
    .TextMatrix(1, 4) = "Location"
    .TextMatrix(1, 5) = "CMP Name"
    .TextMatrix(1, 6) = "Godown No"
    .TextMatrix(1, 7) = "Agreement Expiry Date"
    .TextMatrix(1, 8) = "Godown Capacity"
    .TextMatrix(1, 9) = "Commodity"
    .TextMatrix(1, 10) = "Exchange Type"
    .TextMatrix(1, 11) = "Client Name"
    .TextMatrix(1, 12) = "Balance Bags"
    .TextMatrix(1, 13) = "Balance Qty"
    
    .ColWidth(0) = 1600
    .ColWidth(1) = 1600
    .ColWidth(2) = 1600
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1600
    .ColWidth(10) = 1400
    .ColWidth(11) = 1800
    .ColWidth(12) = 1500
    .ColWidth(13) = 1500
    
    .RowHeight(1) = 700
    .WordWrap = True
End With
End Sub


Private Sub CmbCMPName_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Select Location first !!! "
   CmbLocation.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" Where LocationID=" & mLocationID & " ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "First Select Location!!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!Cmpname
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
Call TableMst_Commodity(" Order by Commodity")
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


Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
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
Call TableMst_Godown("Where CMPID = " & mCMPID)
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
    CmbGodown.AddItem RsMst_Godown!godownno
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
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbSCMFlag_LostFocus()
If Trim(CmbSCMFlag.Text) = "" Then Exit Sub
If LCase(CmbSCMFlag.Text) <> "yes" And LCase(CmbSCMFlag.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSCMFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbsExchangeTypeID_GotFocus()
tmp = CmbSExchangeTypeID.Text
Call TableMst_ExchangeType
CmbSExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbsExchangeTypeID_LostFocus()
If CmbSExchangeTypeID.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSExchangeTypeID.SetFocus
   Exit Sub
End If

mSExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
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
For i = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!stateName
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
For i = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!locationname
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
Call Grid_Head
Label5.Caption = ""
CmdExcel.Enabled = False
CmbLocation.Text = ""
CmbStateID.Text = ""
CmbCommodity.Text = ""
CmbCMPName.Text = ""
CmbGodown.Text = ""
ProgressBar1.Value = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXCodewiseStock.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXCodewiseStock.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
oWS.Range("a:a").ColumnWidth = 15
oWS.Range("b:b").ColumnWidth = 15
oWS.Range("c:c").ColumnWidth = 19
oWS.Range("d:d").ColumnWidth = 8 '40
oWS.Range("e:e").ColumnWidth = 40 '8
oWS.Range("f:f").ColumnWidth = 8 '25
oWS.Range("g:g").ColumnWidth = 25 '10
oWS.Range("h:h").ColumnWidth = 10 '13
oWS.Range("i:i").ColumnWidth = 13 '13
oWS.Range("j:j").ColumnWidth = 13 '19
oWS.Range("k:k").ColumnWidth = 13
oWS.Range("l:l").ColumnWidth = 12
oWS.Range("m:m").ColumnWidth = 12
oWS.Range("n:n").ColumnWidth = 12

tmp = "a" & mRow & ":" & "n" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 1) = "NCDEX Code Wise Stock"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, c + 1).WrapText = True

ProgressBar1.Value = 0

For i = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        If i = 1 Then
           oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, c + 1).Font.Bold = True
           oWS.Cells(mRow, c + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, c + 1).WrapText = True
        oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
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
MsgBox "Please wait till next Message."

Call Grid_Head
ProgressBar1.Value = 0
Label5.Caption = ""

If CmbLocation.Text <> "" Then
   wc = GenWc(wc, mLocationID, "l.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp = Replace(Gen_WcLocation, "LocationID", "L.LocationID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If

If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "S.StateID", "N", "N")
Else
   If Gen_WcState <> "" Then
      tmp = Replace(Gen_WcState, "StateID", "S.StateID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If
If CmbCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "Com.CommodityID", "N", "N")
End If
If CmbCMPName.Text <> "" Then
    wc = GenWc(wc, mCMPID, "CMP.CMPID", "N", "N")
End If
If CmbGodown.Text <> "" Then
    wc = GenWc(wc, mGodownID, "g.GodownID", "N", "N")
End If

If CmbSExchangeTypeID.Text <> "" Then
   wc = GenWc(wc, mSExchangeTypeID, "et.ExchangeTypeID", "N", "N")
End If

If TxtSClientName <> "" Then
   If wc = "" Then
      wc = " Where c.ClientName like '%" & TxtSClientName & "%'"
   Else
      wc = wc & " And c.ClientName like '%" & TxtSClientName & "%'"
   End If
End If

tmp = "Select g.NCDEXCODE,g.CommtrackNCDEXCode, "
tmp = tmp & " g.NComNCDEXCode,s.stateName,l.locationname,cmp.Cmpname,g.godownno,g.GodownID, "
tmp = tmp & " g.GodownCapacity , com.Commodity, et.ExchangeType, c.ClientName, Sum(gsl.BalanceBags) 'Bags', Sum(gsl.BalanceWeight) 'Wt' "
tmp = tmp & " From Mst_GSL gsl "
tmp = tmp & " inner join Mst_Commodity com on gsl.commodityid  = com.commodityid "
tmp = tmp & " inner join mst_Exchangetype et on gsl.ExchangeTypeID  = et.ExchangeTypeID "
tmp = tmp & " inner join Mst_Godown g on gsl.GodownID  = g.GodownID "
tmp = tmp & " inner join Mst_CMP cmp on gsl.CMPID  = cmp.CMPID "
tmp = tmp & " inner join Mst_Location l on cmp.LocationID  = l.LocationID "
tmp = tmp & " inner join Mst_State s on l.Stateid  = s.Stateid "
tmp = tmp & " inner join Mst_Client c on gsl.ClientIDRow  = c.ClientIDRow "
If wc <> "" Then
   tmp = tmp & wc & " And gsl.flag not in ('Z','E','B') And g.closedate is null "
Else
   tmp = tmp & " where gsl.flag not in ('Z','E','B') And g.closedate is null "
End If

tmp = tmp & " Group By g.NCDEXCODE,g.CommtrackNCDEXCode, g.NComNCDEXCode,s.stateName,l.locationname,cmp.Cmpname,"
tmp = tmp & " g.godownno,g.godownid,g.GodownCapacity , com.Commodity, et.ExchangeType, "
tmp = tmp & " c.ClientName "

Call TmpTable


If TmpRs.RecordCount > 0 Then
       
   tmp = "Select b.godownid,isnull(Max(a.ExpiryDate),'') From Txn_GodownLessorAgreeDetail b "
   tmp = tmp & " Inner Join Txn_LessorAgreement a On b.LAID = a.LAID "
   tmp = tmp & " Inner Join Mst_Godown g on b.godownid = g.godownid "
   tmp = tmp & " Where a.Flag = 'A' and g.closedate is null "
   tmp = tmp & " Group by b.godownid "
   
   Call Tmp1Table
    
   ProgressBar1.Max = TmpRs.RecordCount + 2
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   For i = 2 To TmpRs.RecordCount + 1
       For c = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       If Val(MSHFlexGrid1.TextMatrix(i, 13)) < 0 Then
          MSHFlexGrid1.TextMatrix(i, 13) = 0
       End If
       MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(i, 12))
       MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(i, 13))
       Call GetAgreementExpiry(i)
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
ProgressBar1.Value = ProgressBar1.Max
End If

Label5.Caption = TmpRs.RecordCount
Label5.Refresh
CmdExcel.Enabled = True
MsgBox "Done!!"


End Sub

Private Sub GetAgreementExpiry(x_ As Variant)

If TmpRs1.RecordCount > 1 Then
   TmpRs1.MoveFirst
   TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(x_, 7))
End If

If TmpRs1.EOF = False Then
   MSHFlexGrid1.TextMatrix(x_, 7) = IIf(IsNull(TmpRs1.Fields(1)), "", Format(TmpRs1.Fields(1), "dd-mmm-yyyy"))
Else
   MSHFlexGrid1.TextMatrix(x_, 7) = ""
End If


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


