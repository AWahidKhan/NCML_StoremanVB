VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GodownWiseBalanceStockandCommodityValue 
   Caption         =   "Godown Wise Balance Stock and Commodity Value"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7755
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7335
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12938
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
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbSCMFlag 
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
         ItemData        =   "FrmRpt_GodownWiseBalanceStockandCommodityValue.frx":0000
         Left            =   4440
         List            =   "FrmRpt_GodownWiseBalanceStockandCommodityValue.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1290
         Width           =   1245
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
         TabIndex        =   0
         Top             =   570
         Width           =   2655
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
         TabIndex        =   7
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
         Left            =   11205
         TabIndex        =   6
         Top             =   1275
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
         Left            =   13635
         TabIndex        =   8
         Top             =   1275
         Width           =   1215
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
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   570
         Width           =   2655
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
         TabIndex        =   3
         Top             =   570
         Width           =   5415
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
         TabIndex        =   4
         Top             =   1290
         Width           =   4005
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
         TabIndex        =   2
         Top             =   570
         Width           =   3615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "SCM Flag"
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
         Left            =   4620
         TabIndex        =   20
         Top             =   1005
         Width           =   885
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
         TabIndex        =   18
         Top             =   240
         Width           =   1065
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
         Left            =   9360
         TabIndex        =   15
         Top             =   1335
         Width           =   75
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
         TabIndex        =   14
         Top             =   1335
         Width           =   2325
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
         Left            =   3900
         TabIndex        =   13
         Top             =   233
         Width           =   855
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
         TabIndex        =   12
         Top             =   210
         Width           =   1200
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
         Left            =   1807
         TabIndex        =   11
         Top             =   1005
         Width           =   750
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
         TabIndex        =   10
         Top             =   285
         Width           =   1035
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   9600
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmRpt_GodownWiseBalanceStockandCommodityValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID, mExchangeTypeID, mGodownID, mCommodityID As Variant 'mGodownID,

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 13
    .Font.Size = 10
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(1, 0) = "State"
    .TextMatrix(1, 1) = "Location"
    .TextMatrix(1, 2) = "CMP Name"
    .TextMatrix(1, 3) = "Godown No"
    .TextMatrix(1, 4) = "Godown Address"
    .TextMatrix(1, 5) = "Godown Capacity"
    .TextMatrix(1, 6) = "Commodity"
    .TextMatrix(1, 7) = "Balance Bags"
    .TextMatrix(1, 8) = "Balance Qty"
    .TextMatrix(1, 9) = "Latest Rate of Commodity"
    .TextMatrix(1, 10) = "Value of Commodity"
    .TextMatrix(1, 11) = "License No"
    .TextMatrix(1, 12) = "License Expiry Date"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 2000
    .ColWidth(2) = 2000
    .ColWidth(3) = 1000
    .ColWidth(4) = 3500
    .ColWidth(5) = 1200
    .ColWidth(6) = 1800
    .ColWidth(7) = 1200
    .ColWidth(8) = 1500
    .ColWidth(9) = 1800
    .ColWidth(10) = 1400
    .ColWidth(11) = 1500
    .ColWidth(12) = 1500
    
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
Call TableMst_CMP(" where LocationID=" & mLocationID & " ")
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
Call TableMst_Commodity(" Order by Commodity")
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
Call TableMst_Godown("Where CMPID = " & mCMPID)
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWiseBalanceStockandCommodityValue.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWiseBalanceStockandCommodityValue.xls")
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

tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, C + 1) = "Godown Wise Balance Stock and Commodity Value"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, C + 1).WrapText = True

ProgressBar1.Value = 0

For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        If I = 1 Then
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
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
wc = ""
MsgBox "Please wait till next Message."

Call Grid_Head
ProgressBar1.Value = 0
Label5.Caption = ""

If CmbLocation.Text <> "" Then
   wc = GenWc(wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If

If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "MS.StateID", "N", "N")
Else
   If Gen_WcState <> "" Then
      tmp = Replace(Gen_WcState, "StateID", "MS.StateID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If
If CmbCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "MCom.CommodityID", "N", "N")
End If
If CmbCMPName.Text <> "" Then
    wc = GenWc(wc, mCMPID, "MCMP.CMPID", "N", "N")
End If
If CmbGodown.Text <> "" Then
    wc = GenWc(wc, mGodownID, "MG.GodownID", "N", "N")
End If
If LCase(CmbSCMFlag.Text) = "yes" Then
    wc = GenWc(wc, 1, "MGSL.SCMFlag", "N", "N")
ElseIf LCase(CmbSCMFlag.Text) = "no" Then
    wc = GenWc(wc, 0, "MGSL.SCMFlag", "N", "N")
End If


tmp = " Select MS.StateName,ML.LocationName,MCMP.CMPName, MG.GodownNo,MG.Address,MG.GodownCapacity,MCom.Commodity,Sum(MGSL.BalanceBags) 'Bags',Sum(MGSL.BalanceWeight) 'Qty', MS.StateID, MCom.CommodityID,MLC.LicenseNo,MLC.LicenseExpiryDate"
tmp = tmp & " From Mst_GSL MGSL"
tmp = tmp & " Inner Join Mst_Godown MG On MGSL.GodownID = MG.GodownID"
tmp = tmp & " Left Join Mst_License MLC On MG.LicenseID = MLC.LicenseID"
tmp = tmp & " Inner Join Mst_CMP MCMP On MG.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_Commodity MCom on MGSL.CommodityID = MCom.CommodityID"
If wc <> "" Then
    tmp = tmp & " " & wc & " And MG.CloseDate is null And MGSL.FLag not in ('Z','E','B') "
Else
    tmp = tmp & " Where MG.CloseDate is null And MGSL.FLag not in ('Z','E','B') "
End If
tmp = tmp & " Group by MS.StateName,ML.LocationName,MCMP.CMPName, MG.GodownNo,MG.Address,MG.GodownCapacity,MCom.Commodity, MS.StateID, MCom.CommodityID,MLC.LicenseNo,MLC.LicenseExpiryDate"
Call TmpTable

If TmpRs.RecordCount > 0 Then
    ProgressBar1.Max = TmpRs.RecordCount + 2
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
    For I = 2 To TmpRs.RecordCount + 1
        MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
        MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) 'Format(TmpRs!BillTxnDate, "dd-MMM-yyyy"))
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
        MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo) 'Format(TmpRs!MonthYear, "dd-MMM-yyyy"))
        MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
        MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!GodownCapacity), "", TmpRs!GodownCapacity)
        MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
        MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!Bags), "", TmpRs!Bags)
        MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!Qty), "", TmpRs!Qty)
        If Val(MSHFlexGrid1.TextMatrix(I, 8)) < 0 Then
            MSHFlexGrid1.TextMatrix(I, 8) = 0
        End If
        MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
        MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 8))
        Call GetCommodityLatestRate(I)
        MSHFlexGrid1.TextMatrix(I, 10) = Val(MSHFlexGrid1.TextMatrix(I, 8)) * Val(MSHFlexGrid1.TextMatrix(I, 9))
        MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(I, 10))
        MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!LicenseNo), "", TmpRs!LicenseNo)
        MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!LicenseExpiryDate), "", Format(TmpRs!LicenseExpiryDate, "dd-MMM-yyyy"))
        
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

Private Sub GetCommodityLatestRate(x_ As Variant)

'tmp = " Select MCP.YearMonth, MCP.Day28,MCP.Day29,MCP.Day30,MCP.Day31 from Mst_CommodityPrice MCP" 'MS.StateName,MC.Commodity,MCP.MonthYear,MCP.YearMonth,
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
tmp = tmp & " Where StateID = " & TmpRs!StateID & " And CommodityID = " & TmpRs!CommodityID & "  And Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,YearMonth) In  (Select Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,MAX(YearMonth)) From Mst_CommodityPrice Group By StateID,CommodityID )"


Call Tmp2Table

MSHFlexGrid1.TextMatrix(x_, 9) = 0
'If TmpRs2.RecordCount > 0 Then
'    TmpRs2.MoveLast
'    If Right(TmpRs2!YearMonth, 2) = "01" Or Right(TmpRs2!YearMonth, 2) = "03" Or Right(TmpRs2!YearMonth, 2) = "05" Or Right(TmpRs2!YearMonth, 2) = "07" Or Right(TmpRs2!YearMonth, 2) = "08" Or Right(TmpRs2!YearMonth, 2) = "10" Or Right(TmpRs2!YearMonth, 2) = "12" Then
'        MSHFlexGrid1.TextMatrix(x_, 9) = TmpRs2!Day31
'    ElseIf Right(TmpRs2!YearMonth, 2) = "04" Or Right(TmpRs2!YearMonth, 2) = "06" Or Right(TmpRs2!YearMonth, 2) = "09" Or Right(TmpRs2!YearMonth, 2) = "11" Then
'        MSHFlexGrid1.TextMatrix(x_, 9) = TmpRs2!Day30
'    Else
'        If IsNull(TmpRs2!Day29) = False Then
'            MSHFlexGrid1.TextMatrix(x_, 9) = TmpRs2!Day29
'        Else
'            MSHFlexGrid1.TextMatrix(x_, 9) = TmpRs2!Day28
'        End If
'    End If
'Else
'    tmp = " Select MinRate from Mst_Commodity Where CommodityID = " & TmpRs!CommodityID & ""
'
'    Call Tmp2Table
'
'    If TmpRs2.RecordCount > 0 Then
'        MSHFlexGrid1.TextMatrix(x_, 9) = IIf(IsNull(TmpRs2!MinRate), 0, TmpRs2!MinRate)
'    End If
'End If


If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x_, 9) = IIf(IsNull(TmpRs2!Rate), 0, TmpRs2!Rate)
Else
   tmp = " Select MinRate from Mst_Commodity Where CommodityID = " & TmpRs!CommodityID & ""
   Call Tmp2Table
   If TmpRs2.RecordCount > 0 Then
      MSHFlexGrid1.TextMatrix(x_, 9) = IIf(IsNull(TmpRs2!MinRate), 0, TmpRs2!MinRate)
   End If
End If




End Sub


Private Sub Combo1_Change()

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

