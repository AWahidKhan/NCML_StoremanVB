VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_DepositContract 
   Caption         =   "Deposit Contract Report"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   12375
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbSContractType 
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
         ItemData        =   "FrmRpt_DepositContract.frx":0000
         Left            =   9400
         List            =   "FrmRpt_DepositContract.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1100
         Width           =   1700
      End
      Begin VB.ComboBox CmbSType 
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
         ItemData        =   "FrmRpt_DepositContract.frx":0020
         Left            =   13080
         List            =   "FrmRpt_DepositContract.frx":002A
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1100
         Width           =   1785
      End
      Begin VB.ComboBox CmbSVerified 
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
         ItemData        =   "FrmRpt_DepositContract.frx":003E
         Left            =   11160
         List            =   "FrmRpt_DepositContract.frx":0048
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1860
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
         Left            =   13650
         TabIndex        =   12
         Top             =   1550
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
         Height          =   420
         Left            =   11220
         TabIndex        =   10
         Top             =   1550
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
         Height          =   420
         Left            =   12435
         TabIndex        =   11
         Top             =   1550
         Width           =   1215
      End
      Begin VB.ComboBox CmbClientID 
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
         TabIndex        =   5
         Top             =   1100
         Width           =   6900
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
         TabIndex        =   2
         Top             =   420
         Width           =   3960
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
         Left            =   7110
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   4000
      End
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmRpt_DepositContract.frx":0055
         Left            =   7080
         List            =   "FrmRpt_DepositContract.frx":0057
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1100
         Width           =   2280
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
         Left            =   11160
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   3720
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   120
         TabIndex        =   0
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
         Format          =   106823683
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   1620
         TabIndex        =   1
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
         Format          =   106823683
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label14 
         Caption         =   "Contract Type"
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
         Index           =   6
         Left            =   9572
         TabIndex        =   28
         Top             =   820
         Width           =   1320
      End
      Begin VB.Label Label14 
         Caption         =   "Quantity / Bags"
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
         Index           =   5
         Left            =   13252
         TabIndex        =   27
         Top             =   825
         Width           =   1440
      End
      Begin VB.Label Label14 
         Caption         =   "Verified"
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
         Index           =   4
         Left            =   11730
         TabIndex        =   26
         Top             =   825
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "Total Record Found :"
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
         Left            =   7200
         TabIndex        =   25
         Top             =   1633
         Width           =   2205
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
         Left            =   9480
         TabIndex        =   24
         Top             =   1640
         Width           =   90
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
         TabIndex        =   22
         Top             =   165
         Width           =   780
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
         TabIndex        =   21
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label LblCMPName 
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
         Height          =   300
         Left            =   2670
         TabIndex        =   20
         Top             =   820
         Width           =   1200
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
         Left            =   4553
         TabIndex        =   19
         Top             =   150
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
         Left            =   8708
         TabIndex        =   18
         Top             =   180
         Width           =   765
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
         Left            =   12480
         TabIndex        =   17
         Top             =   150
         Width           =   1080
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   7920
         TabIndex        =   16
         Top             =   820
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7455
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7100
         Left            =   120
         TabIndex        =   13
         Top             =   150
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12515
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   120
         Left            =   120
         TabIndex        =   23
         Top             =   7250
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   212
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_DepositContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID As Variant, mLocationID As Variant, mCommodityID As Variant, mClientRowID As Variant, mStatus As Variant
Dim wc As Variant, I As Double

Dim mSContractTypeFlag As Variant, mSType As Variant, mSVerified As Variant
Dim tmp1 As Variant

Private Sub CmbLocationID_GotFocus()
tmp1 = CmbLocationID.Text
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
If LCase(tmp1) <> LCase(CmbLocationID.Text) Then
    CmbCommodityID.Text = ""
End If
End Sub

Private Sub CmbCommodityID_GotFocus()

If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbCommodityID.Text
''Call TableMst_Commodity


Call TableMst_Commodity(" Where CommodityID IN (SELECT Distinct CommodityID FROM Mst_LocationCommodityMapping Where LocationID = " & mLocationID & " ) ")
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

Private Sub CmbClientId_GotFocus()
tmp = CmbClientID.Text
Call TableMst_Client("Where ExchangeTypeID = " & GetExchangeID("Non NCDEX") & " Order By ClientName ")
CmbClientID.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClientID.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClientID.Text = tmp
End Sub
Private Sub CmbClientId_LostFocus()
If CmbClientID.Text = "" Then Exit Sub
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClientID.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbClientID.SetFocus
   Exit Sub
End If
mClientRowID = RsMst_Client!ClientID
End Sub

Private Sub CmbStateID_GotFocus()
tmp1 = CmbStateID.Text
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
If LCase(tmp1) <> LCase(CmbStateID.Text) Then
    CmbLocationID.Text = ""
    CmbCommodityID.Text = ""
End If
End Sub

Private Sub CmbStatus_GotFocus()
CmbStatus.Clear
CmbStatus.AddItem "System Approved"
CmbStatus.AddItem "Approved"
CmbStatus.AddItem "Forwarded"
CmbStatus.AddItem "Rejected"
CmbStatus.AddItem "Closed"
End Sub
Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "system approved" And LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "rejected" And LCase(CmbStatus.Text) <> "forwarded" And LCase(CmbStatus.Text) <> "closed" Then
   MsgBox "Invalid Selection !!!"
   CmbStatus.SetFocus
   Exit Sub
End If
If LCase(CmbStatus.Text) = "system approved" Then mStatus = "'S'"
If LCase(CmbStatus.Text) = "approved" Then mStatus = "'A'"
If LCase(CmbStatus.Text) = "rejected" Then mStatus = "'R'"
If LCase(CmbStatus.Text) = "forwarded" Then mStatus = "'F'"
If LCase(CmbStatus.Text) = "closed" Then mStatus = "'C'"
End Sub

''-- 19 Feb 2018
Private Sub CmbSContractType_LostFocus()
If CmbSContractType.Text = "" Then Exit Sub
If LCase(CmbSContractType.Text) <> "normal" And LCase(CmbSContractType.Text) <> "revision" Then
   MsgBox "Invalid Selection !!!"
   CmbSContractType.SetFocus
   Exit Sub
End If
If LCase(CmbSContractType.Text) = "normal" Then mSContractTypeFlag = "N"
If LCase(CmbSContractType.Text) = "revision" Then mSContractTypeFlag = "R"
End Sub
Private Sub CmbSType_LostFocus()
If CmbSType.Text = "" Then Exit Sub
If LCase(CmbSType.Text) <> "quantity" And LCase(CmbSType.Text) <> "bags" Then
   MsgBox "Invalid Selection !!!"
   CmbSType.SetFocus
   Exit Sub
End If
If LCase(CmbSType.Text) = "quantity" Then mSType = "N"
If LCase(CmbSType.Text) = "bags" Then mSType = "Y"
End Sub
Private Sub CmbSVerified_LostFocus()
If CmbSVerified.Text = "" Then Exit Sub
If LCase(CmbSVerified.Text) <> "yes" And LCase(CmbSVerified.Text) <> "no" Then
   MsgBox "Invalid Selection !!!"
   CmbSVerified.SetFocus
   Exit Sub
End If
mSVerified = UCase(Left(CmbSVerified.Text, 1))
End Sub
''------------------

Private Sub CmdClear_Click()
Call Grid_Head
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbCommodityID.Text = ""
CmbClientID.Text = ""
CmbStatus.Text = ""
'ChkBagSize.Value = vbUnchecked
TxtFromDate.Value = Date
TxtToDate.Value = Date

''-- 19 Feb 2018
CmbSContractType.Text = ""
CmbSType.Text = ""
CmbSVerified.Text = ""
Label5.Caption = "-"
''-----

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_DepositContract.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_DepositContract.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
With MSHFlexGrid1
mRow = 1
For I = 0 To .Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To .Cols - 1
        If InStr(.TextMatrix(I, C), "/") > 0 Then
           If IsDate(.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
           End If
        Else
           If C = 8 Then
            oWS.Cells(mRow, C + 1) = "'" & .TextMatrix(I, C)
           Else
            oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
           End If
           If C = 18 Or C = 20 Then
              .row = I
              .Col = C
              If .CellBackColor = vbRed Then oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
              If .CellBackColor = vbGreen Then oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbGreen
              If .CellBackColor = vbYellow Then oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbYellow
           End If
           ''oWS.Cells(mRow, c + 1) = .TextMatrix(i, c)
        End If
    Next
Next
End With
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGetReport_Click()
Dim wc As String
Call Grid_Head
Label5.Caption = "-"
MSHFlexGrid1.Refresh
MsgBox ("Wait till next message")

wc = ""
If TxtToDate.Value < TxtFromDate.Value Then
   MsgBox "To date must be always greater than from date!!!"
   TxtToDate.SetFocus
   Exit Sub
End If
wc = " Where TDC.Date Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "' "
If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "MS.StateID", "N", "N")
Else
   If Gen_WcState <> "" Then
      tmp = Replace(Gen_WcState, "StateID", "MS.StateID")
      wc = wc & " And " & tmp
   End If
End If
If CmbLocationID.Text <> "" Then
   wc = GenWc(wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = wc & " And " & tmp
   End If
End If
If CmbCommodityID.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "Mco.CommodityID", "N", "N")
End If
If CmbClientID.Text <> "" Then
   wc = GenWc(wc, mClientRowID, "MC.ClientID", "N", "N")
End If
If CmbStatus.Text <> "" Then
   If wc = "" Then
      wc = " TDC.ApprovalFlag in (" & mStatus & ")"
   Else
      wc = wc & " And TDC.ApprovalFlag in (" & mStatus & ")"
   End If
End If

''--- Ullhas Patil 19 Feb 2018
''If ChkBagSize.Value = vbChecked Then
''    wc = wc & " And TDC.BagSizeFlag = 'Y' "
''End If
If CmbSContractType.Text <> "" Then
    wc = wc & " And TDC.ContractType = '" & mSContractTypeFlag & "' "
End If
If CmbSVerified.Text <> "" Then
    wc = wc & " And Isnull(TDC.Verified,'N') = '" & mSVerified & "' "
End If
If CmbSType.Text <> "" Then
    wc = wc & " And TDC.BagSizeFlag = '" & mSType & "' "
End If
''-----------

tmp = " Select TDC.DepConID,TDC.Date,MS.StateName,ML.LocationName,Mco.Commodity,MC.ClientName,MC.ClientID,MC.ClientIDRow, TDC.QuantumMT,MIB.InsuranceByName,TDC.StartDate,TDC.EndDate, MBC.BillingCycle As BillingCycle, MBCDep.BillingCycle As BillingCycleDep,TDC.BagSizeFlag,TDC.RatePer,TDC.TentativeDur,TDC.LockPeriod,ME.EmployeeName+'~'+Me.EmployeeNo As ForwardedTo,TDC.ApprovalFlag,TDC.ApproveOn,MEE.EmployeeName+'~'+MEE.EmployeeNo As ApproveBy,TDC.ClosedOn,TDC.ClosedBy,MEEE.EmployeeName+'~'+MEEE.EmployeeNo As CloseBy,CMP.CMPName,TDCD.GodownId,MG.GodownNo,MG.Address,MUT.UnitType,TDCD.FinalPrice,TDCD.DepositFinalPrice,TDCD.WithdrawalFinalPrice,TDCD.BaseRate,TDCD.DepositBaseRate,TDCD.WithdrawalBaseRate,TDCD.OldBillingRate,TDCD.OldDepositBillingRate,TDCD.OldWithdrawalBillingRate,TDCD.MinBagSize,TDCD.MaxBagSize,TDCD.BaseRateColor,TDCD.DepositBaseRateColor,TDCD.WithdrawalBaseRateColor,TDCD.RatePerMonthPerMT,TDCD.Rent,"
tmp = tmp & " TDCD.DepConDID,isnull(TDC.Verified,'N')Verified,MEEEE.EmployeeName+'~'+MEEEE.EmployeeNo As VerifiedBy,TDC.VerifiedDate,TDC.ContractType"
tmp = tmp & " From Txn_DepositContract TDC"
tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On TDC.DepConID = TDCD.DepConID"
tmp = tmp & " Inner Join Mst_Godown MG on TDCD.GodownId = MG.GodownId"
tmp = tmp & " Inner Join MST_CMP CMP on MG.CMPID = CMP.CMPID"
tmp = tmp & " Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID"
tmp = tmp & " Inner Join MST_Location ML on TDC.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Inner Join MST_Commodity Mco on TDC.CommodityID=Mco.CommodityID"
tmp = tmp & " Inner Join MST_Client MC on TDC.ClientID=MC.ClientID"
tmp = tmp & " Inner Join MST_BillingCycle MBC on TDC.BillingCycleID=MBC.BillingCycleID"
tmp = tmp & " Inner Join MST_BillingCycle MBCDep on TDC.DepBillingCycleID=MBCDep.BillingCycleID"
tmp = tmp & " Inner Join Mst_InsuranceBy MIB on TDC.InsuranceByID=MIB.InsuranceByID"
tmp = tmp & " Left Join MST_Employee ME on TDC.EmployeeID= ME.EmployeeID"
tmp = tmp & " Left Join MST_Employee MEE on TDC.ApproveBy=MEE.EmployeeNo"
tmp = tmp & " Left Join MST_Employee MEEE on TDC.ClosedBy=MEEE.EmployeeNo"
tmp = tmp & " Left Join MST_Employee MEEEE on TDC.VerifiedBy=MEEEE.EmployeeNo"

If wc <> "" Then
   tmp = tmp & wc
End If
tmp = tmp & " ORDER BY TDC.DepConID,TDC.Date "
Call TmpTable
If TmpRs.RecordCount > 0 Then
ProgressBar1.Max = TmpRs.RecordCount
For I = 1 To TmpRs.RecordCount
    With MSHFlexGrid1
        .TextMatrix(I, 0) = IIf(IsNull(TmpRs!DepConID), "", TmpRs!DepConID)                     ''-- Deposit Contract ID
        .TextMatrix(I, 1) = IIf(IsNull(TmpRs!Date), "", Format(TmpRs!Date, "DD-MMM-YYYY"))      ''-- Date
        .TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)                   ''-- State
        .TextMatrix(I, 3) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)             ''-- Location
        .TextMatrix(I, 4) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)                   ''-- Commodity
        .TextMatrix(I, 5) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName & "#" & TmpRs!ClientIDRow & "#" & TmpRs!ClientID)               ''-- Client
        .TextMatrix(I, 6) = IIf(IsNull(TmpRs!InsuranceByName), "", TmpRs!InsuranceByName)       ''-- Insurance By
        .TextMatrix(I, 7) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
        .TextMatrix(I, 8) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
        .TextMatrix(I, 9) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
        .TextMatrix(I, 10) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "DD-MMM-YYYY"))  ''-- Deposit Contract Start Date
        .TextMatrix(I, 11) = IIf(IsNull(TmpRs!EndDate), "", Format(TmpRs!EndDate, "DD-MMM-YYYY"))      ''-- Deposit Contract End Date
        .TextMatrix(I, 12) = IIf(IsNull(TmpRs!BagSizeFlag), "", IIf(LCase(TmpRs!BagSizeFlag) = "n", "Quantity Wise", IIf(LCase(TmpRs!BagSizeFlag) = "y", "Bag Wise", "")))           ''-- Quantity / Bag
        .TextMatrix(I, 13) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle)                   ''-- Billing Cycle
        .TextMatrix(I, 14) = IIf(IsNull(TmpRs!BillingCycleDep), "", TmpRs!BillingCycleDep)            ''-- Deposit/ Withdrawal Billing Cycle
        .TextMatrix(I, 15) = IIf(IsNull(TmpRs!MinBagSize), "", TmpRs!MinBagSize)
        .TextMatrix(I, 16) = IIf(IsNull(TmpRs!MaxBagSize), "", TmpRs!MaxBagSize)
        .TextMatrix(I, 17) = IIf(IsNull(TmpRs!OldBillingRate), "", TmpRs!OldBillingRate)
        
        If LCase(TmpRs!BaseRateColor) = "r" Then
            .row = I: .Col = 18: .CellBackColor = vbRed
        ElseIf LCase(TmpRs!BaseRateColor) = "y" Then
            .row = I: .Col = 18: .CellBackColor = vbYellow
        ElseIf LCase(TmpRs!BaseRateColor) = "g" Then
            .row = I: .Col = 18: .CellBackColor = vbGreen
        End If
        
        .TextMatrix(I, 18) = Format(IIf(IsNull(TmpRs!FinalPrice), "", TmpRs!FinalPrice), "#0.0000")
        .TextMatrix(I, 19) = IIf(IsNull(TmpRs!OldDepositBillingRate), "", TmpRs!OldDepositBillingRate)
        
        If LCase(TmpRs!DepositBaseRateColor) = "r" Then
            .row = I: .Col = 20: .CellBackColor = vbRed
        ElseIf LCase(TmpRs!DepositBaseRateColor) = "y" Then
            .row = I: .Col = 20: .CellBackColor = vbYellow
        ElseIf LCase(TmpRs!DepositBaseRateColor) = "g" Then
            .row = I: .Col = 20: .CellBackColor = vbGreen
        End If
        
        .TextMatrix(I, 20) = Format(IIf(IsNull(TmpRs!DepositFinalPrice), "", TmpRs!DepositFinalPrice), "#0.0000")
        .TextMatrix(I, 21) = IIf(IsNull(TmpRs!RatePerMonthPerMT), "", TmpRs!RatePerMonthPerMT)
        .TextMatrix(I, 22) = IIf(IsNull(TmpRs!RatePer), "", IIf(UCase(TmpRs!RatePer) = "I", "Included", "Excluded"))    ''-- Service Tax (GST)
        .TextMatrix(I, 23) = IIf(IsNull(TmpRs!ForwardedTo), "", TmpRs!ForwardedTo)                                      ''-- Forwarded To
        .TextMatrix(I, 24) = IIf(IsNull(TmpRs!ApprovalFlag), "", UCase(TmpRs!ApprovalFlag))                             ''-- Status
            If .TextMatrix(I, 24) = "A" Then .TextMatrix(I, 24) = "Approved"
            If .TextMatrix(I, 24) = "S" Then .TextMatrix(I, 24) = "System Approved"
            If .TextMatrix(I, 24) = "F" Then .TextMatrix(I, 24) = "Forwarded"
            If .TextMatrix(I, 24) = "R" Then .TextMatrix(I, 24) = "Rejected"
            If .TextMatrix(I, 24) = "C" Then .TextMatrix(I, 24) = "Closed"
        .TextMatrix(I, 25) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
        .TextMatrix(I, 26) = IIf(IsNull(TmpRs!TentativeDur), "", TmpRs!TentativeDur)    ''-- Tentative Duration of Storage (in months)
        .TextMatrix(I, 27) = IIf(IsNull(TmpRs!LockPeriod), "", TmpRs!LockPeriod)        ''-- Lock-in Period (in months)
        .TextMatrix(I, 28) = IIf(IsNull(TmpRs!QuantumMT), "", TmpRs!QuantumMT)          ''-- Indicative Quantum
        If LCase(TmpRs!ApprovalFlag) = "a" Or LCase(TmpRs!ApprovalFlag) = "c" Then
            .TextMatrix(I, 29) = IIf(IsNull(TmpRs!ApproveOn), "", Format(TmpRs!ApproveOn, "DD-MMM-YYYY hh:mm:sss"))               ''-- Approveon
            .TextMatrix(I, 30) = IIf(IsNull(TmpRs!ApproveBy), "", TmpRs!ApproveBy)                                                ''-- ApproveBy
        End If
        If LCase(TmpRs!ApprovalFlag) = "c" Then
            .TextMatrix(I, 31) = IIf(IsNull(TmpRs!ClosedOn), "", Format(TmpRs!ClosedOn, "DD-MMM-YYYY hh:mm:sss"))           ''-- ClosedNo
            .TextMatrix(I, 32) = IIf(IsNull(TmpRs!CloseBy), "", TmpRs!CloseBy)                                              ''-- ClosedBy
        End If
        
        .TextMatrix(I, 33) = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)
        
        ''-- Ullhas Patil 19 Feb 2018
        .TextMatrix(I, 34) = IIf(IsNull(TmpRs!ContractType), "", IIf(LCase(TmpRs!ContractType) = "n", "Normal", IIf(LCase(TmpRs!ContractType) = "r", "Revision", "")))           ''-- Normal / Revision
        .TextMatrix(I, 35) = IIf(IsNull(TmpRs!Verified), "No", IIf(LCase(TmpRs!Verified) = "n", "No", IIf(LCase(TmpRs!Verified) = "y", "Yes", "")))                              ''-- Yes / No
        If LCase(IIf(IsNull(TmpRs!Verified), "n", TmpRs!Verified)) = "y" Then
            .TextMatrix(I, 36) = IIf(IsNull(TmpRs!VerifiedDate), "", Format(TmpRs!VerifiedDate, "DD-MMM-YYYY hh:mm:sss"))           ''-- VerifiedDate
            .TextMatrix(I, 37) = IIf(IsNull(TmpRs!VerifiedBy), "", TmpRs!VerifiedBy)                                      ''-- VerifiedBy
        End If
        ''------------------------------
        
        .Rows = .Rows + 1
    End With
    TmpRs.MoveNext
    ProgressBar1.Value = I
Next
End If
Label5.Caption = TmpRs.RecordCount
MsgBox ("Done.")
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
Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
.Clear
.Rows = 2
.Cols = 38
.WordWrap = True
.RowHeight(0) = 800
.Font.Size = 8
.TextMatrix(0, 0) = "Deposit Contract ID":                                  .ColWidth(0) = 1000
.TextMatrix(0, 1) = "Date":                                                 .ColWidth(1) = 1200
.TextMatrix(0, 2) = "State":                                                .ColWidth(2) = 1500
.TextMatrix(0, 3) = "Location":                                             .ColWidth(3) = 1500
.TextMatrix(0, 4) = "Commodity":                                            .ColWidth(4) = 1500
.TextMatrix(0, 5) = "Client":                                               .ColWidth(5) = 4500
.TextMatrix(0, 6) = "Insurance By":                                         .ColWidth(6) = 1000
.TextMatrix(0, 7) = "CMP Name":                                             .ColWidth(7) = 1300
.TextMatrix(0, 8) = "Godown No":                                            .ColWidth(8) = 800
.TextMatrix(0, 9) = "Godown Type":                                          .ColWidth(9) = 1000
.TextMatrix(0, 10) = "Deposit Contract Start Date":                         .ColWidth(10) = 1200
.TextMatrix(0, 11) = "Deposit Contract End Date":                           .ColWidth(11) = 1200
.TextMatrix(0, 12) = "Qunatity / Bag":                                      .ColWidth(12) = 1200
.TextMatrix(0, 13) = "Billing Cycle":                                       .ColWidth(13) = 1200
.TextMatrix(0, 14) = "Deposit/ Withdrawal Billing Cycle":                   .ColWidth(14) = 1200
.TextMatrix(0, 15) = "Bag Size (Minimum)":                                  .ColWidth(15) = 900
.TextMatrix(0, 16) = "Bag Size (Maximum)":                                  .ColWidth(16) = 900
.TextMatrix(0, 17) = "Billing Rate/MT (As per System)":                     .ColWidth(17) = 1000
.TextMatrix(0, 18) = "Billing Rate/MT (Proposed/ Approved)":                .ColWidth(18) = 1200
.TextMatrix(0, 19) = "Deposit/ Withdrawal Rate/MT (As per System)":         .ColWidth(19) = 1200
.TextMatrix(0, 20) = "Deposit/ Withdrawal Rate/MT (Proposed/ Approved)":    .ColWidth(20) = 1200
.TextMatrix(0, 21) = "Billing Rate/MT (Monthly Cycle)":                     .ColWidth(21) = 1200
.TextMatrix(0, 22) = "GST":                                                 .ColWidth(22) = 1200
.TextMatrix(0, 23) = "Forwarded To":                                        .ColWidth(23) = 1800
.TextMatrix(0, 24) = "Status":                                              .ColWidth(24) = 1200
.TextMatrix(0, 25) = "Address":                                             .ColWidth(25) = 2000
.TextMatrix(0, 26) = "Tentative Duration of Storage (in months)":           .ColWidth(26) = 1000
.TextMatrix(0, 27) = "Lock-in Period (in months)":                          .ColWidth(27) = 1000
.TextMatrix(0, 28) = "Indicative Quantum":                                  .ColWidth(28) = 1000
.TextMatrix(0, 29) = "Approved On":                                         .ColWidth(29) = 2000
.TextMatrix(0, 30) = "Approved By":                                         .ColWidth(30) = 2500
.TextMatrix(0, 31) = "Closed On":                                           .ColWidth(31) = 2000
.TextMatrix(0, 32) = "Closed By":                                           .ColWidth(32) = 2500
.TextMatrix(0, 33) = "Detail ID":                                           .ColWidth(33) = 800

''-- Ullhas Patil 19 Feb 2018
.TextMatrix(0, 34) = "Contract Type":                                       .ColWidth(34) = 1000
.TextMatrix(0, 35) = "Verified":                                            .ColWidth(35) = 800
.TextMatrix(0, 36) = "Verified On":                                         .ColWidth(36) = 2000
.TextMatrix(0, 37) = "Verified By":                                         .ColWidth(37) = 2500
''-------

End With
End Sub

