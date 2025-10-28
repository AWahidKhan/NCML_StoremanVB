VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_DailyPnL 
   Caption         =   "Daily P & L Report"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Daily P&&L"
      TabPicture(0)   =   "FrmRpt_DailyPnL.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details for Avg Rate Calculation"
      TabPicture(1)   =   "FrmRpt_DailyPnL.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid4"
      Tab(1).Control(1)=   "MSHFlexGrid3"
      Tab(1).ControlCount=   2
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   19
         Top             =   420
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   12726
         _Version        =   393216
         WordWrap        =   -1  'True
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   6800
         _Version        =   393216
         WordWrap        =   -1  'True
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   21
         Top             =   4320
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   6588
         _Version        =   393216
         WordWrap        =   -1  'True
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
   Begin VB.Frame Frame11 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Apply Rate to Transactions only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   840
         Width           =   3255
      End
      Begin VB.CommandButton CmdUpdateBillingRate 
         Caption         =   "Update Billing Rate on GSL"
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
         Left            =   8400
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox ChkRateDetails 
         Appearance      =   0  'Flat
         Caption         =   "Get AvgRate Calculation Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   12480
         TabIndex        =   22
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton CmdProcess 
         Caption         =   "Process"
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
         Left            =   11040
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Opt_Stock 
         Caption         =   "Stock Wise"
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
         Left            =   1920
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Opt_Capacity 
         Caption         =   "Capacity Wise"
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
         TabIndex        =   6
         Top             =   795
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   12270
         TabIndex        =   5
         Top             =   840
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
         Left            =   11055
         TabIndex        =   4
         Top             =   840
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
         Left            =   13485
         TabIndex        =   3
         Top             =   840
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
         Left            =   4950
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   4680
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3285
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110034945
         CurrentDate     =   39850
      End
      Begin VB.Label Label44 
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
         Left            =   360
         TabIndex        =   12
         Top             =   150
         Width           =   945
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
         Left            =   4560
         TabIndex        =   11
         Top             =   1275
         Width           =   2325
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   6960
         TabIndex        =   10
         Top             =   1275
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
         Left            =   6908
         TabIndex        =   9
         Top             =   150
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
         Left            =   2670
         TabIndex        =   8
         Top             =   150
         Width           =   1065
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   9390
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3836
      _Version        =   393216
      WordWrap        =   -1  'True
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid13 
      Height          =   2175
      Left            =   360
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3836
      _Version        =   393216
      WordWrap        =   -1  'True
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
Attribute VB_Name = "FrmRpt_DailyPnL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLOcationID, mStateID As Variant 'mTotalDays, mTotalMonth
Dim mStartDate, mEndDate, mMonth, mYear, mTotalDays, wc, mCurrStartDate, mWiseStartDate, mWideEndDate, mWiseTotalDays, mTotalMonths As Variant
Dim mGodownIDs As Variant

Enum ETempDailyPnL
RptDate = 0
StateName = 1
StateID = 2
LocationID = 3
GodownID = 4
LocationName = 5
AMname = 6
RMName = 7
CMPName = 8
godownno = 9
GodownCapacity = 10
UnitType = 11
Stock = 12
RStock = 13 'AvgRate = 13
IncomewithoutReservation = 14
ReservationIncome = 15
CMinCMP = 16
TotalIncome = 17
LeaseRent = 18
EmployeeClaim = 19
VendorClaim = 20
FumigationClaim = 21
FranchiseeFees = 22
SalaryCost = 23
SecurityCost = 24
InterestCharges = 25
CoExpense = 26
NotionalRent = 27
TotalExpense = 28
NetIncome = 29
GStartDate = 30

End Enum

Dim mTotalCapacity, mContractID As Variant

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
Call TableMst_State
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub
'----
Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbLocationID.Text = ""
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
   CmbLocationID.Text = ""
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

Call TableMst_Location("Where StateID=" & mStateID & "")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbStateID.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLOcationID = RsMst_Location!LocationID
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Call Grid_HeadIncomeExpense
TxtFromDate.Value = Date
CmbStateID.Text = ""
CmbLocationID.Text = ""
LblTotalRecords.Caption = "-"
Call Grid_HeadGodownStock
Call Grid_HeadGodownIncome
ChkRateDetails.Value = 0
Check1.Value = 1
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_DailyPnL.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Variant
Dim mCol As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_DailyPnL.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Godown Wise Profitability"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        Else
           If C = 9 Then
              oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(i, C)
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
       End If
       oWS.Cells(mRow, C + 1).WrapText = True
       oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
       If i = 1 Then
          oWS.Cells(mRow, C + 1).Font.Bold = True
       ElseIf i > 1 Then
          If C < 19 Then
             MSHFlexGrid1.row = i
             MSHFlexGrid1.Col = C
             If MSHFlexGrid1.CellBackColor = vbRed Then
                oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
             End If
          End If
       End If
       
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
 
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
  
mRow = 2
For C = 0 To MSHFlexGrid1.Cols - 1
    If C = 14 Or C = 16 Or C = 23 Or C = 24 Then
       oWS.Cells(mRow, C + 1).Interior.Color = vbGreen
    ElseIf C = 15 Or C = 18 Or C = 25 Or C = 26 Or C = 27 Then
       oWS.Cells(mRow, C + 1).Interior.Color = &HC0FF&
    ElseIf C = 19 Or C = 20 Or C = 21 Or C = 22 Then
       oWS.Cells(mRow, C + 1).Interior.Color = vbRed
    ElseIf C = 17 Or C = 28 Or C = 29 Then
       oWS.Cells(mRow, C + 1).Interior.Color = &HF3EEDB
    End If

Next


Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Godownwise Wt Stock"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid3.Rows + 2
mRow = 0
For i = 0 To MSHFlexGrid3.Rows - 2
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid3.Cols - 1
        If C = 3 Then
           oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid3.TextMatrix(i, C)
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(i, C)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max

Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "Godown Income"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid4.Rows + 2
mRow = 0
For i = 0 To MSHFlexGrid4.Rows - 2
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid4.Cols - 1
       If C = 3 Then
          oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid4.TextMatrix(i, C)
       Else
          oWS.Cells(mRow, C + 1) = MSHFlexGrid4.TextMatrix(i, C)
       End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub

Private Sub Grid_Head()
CmdExcel.Enabled = False
LblTotalRecords.Caption = ""
ProgressBar1.Value = 0

tmp = "Select MAX(MonthYear) 'MonthYear' from Txn_MonGodownWiseProfitability"
Call TmpTable
mCurrStartDate = Year(TmpRs!MonthYear) & "-" & Month(TmpRs!MonthYear) & "-1"

With MSHFlexGrid1
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 31 '60 '59
   
   .TextMatrix(1, ETempDailyPnL.RptDate) = "Date"
   .TextMatrix(1, ETempDailyPnL.StateName) = "State Name"
   .TextMatrix(1, ETempDailyPnL.LocationName) = "Location Name"
   .TextMatrix(1, ETempDailyPnL.AMname) = "AM Name"
   .TextMatrix(1, ETempDailyPnL.RMName) = "RM Name"
   .TextMatrix(1, ETempDailyPnL.CMPName) = "CMP Name"
   .TextMatrix(1, ETempDailyPnL.godownno) = "Godown No"
   .TextMatrix(1, ETempDailyPnL.GodownCapacity) = "Godown Capacity "
   .TextMatrix(1, ETempDailyPnL.UnitType) = "Unit Type "
   .TextMatrix(1, ETempDailyPnL.Stock) = "Wt Stock (Non Reservation)"
   .TextMatrix(1, ETempDailyPnL.RStock) = "Wt Stock (Reservation)" '.TextMatrix(1, ETempDailyPnL.AvgRate) = "Avg Rate From PnL of - " & Format(mCurrStartDate, "MMM-yyyy")
   .TextMatrix(1, ETempDailyPnL.IncomewithoutReservation) = "Income excluding Reservation From PnL of - " & Format(mCurrStartDate, "MMM-yyyy")
   .TextMatrix(1, ETempDailyPnL.ReservationIncome) = "Reservation Income"
   .TextMatrix(1, ETempDailyPnL.CMinCMP) = "CM in CMP Income From PnL of - " & Format(mCurrStartDate, "MMM-yyyy")
   .TextMatrix(1, ETempDailyPnL.TotalIncome) = "Total Income "
   .TextMatrix(1, ETempDailyPnL.LeaseRent) = "Lease Rent From Current Rent Agreement"
   .TextMatrix(1, ETempDailyPnL.EmployeeClaim) = "Employee Claim From PnL Starting FY upto - " & Format(mCurrStartDate, "MMM-yyyy")
   .TextMatrix(1, ETempDailyPnL.VendorClaim) = "Vendor Claim From PnL Starting FY upto - " & Format(mCurrStartDate, "MMM-yyyy")
   .TextMatrix(1, ETempDailyPnL.FumigationClaim) = "Fumigation Claim From PnL Starting FY upto - " & Format(mCurrStartDate, "MMM-yyyy")
   .TextMatrix(1, ETempDailyPnL.FranchiseeFees) = "Franchisee Fees From PnL Starting FY upto - " & Format(mCurrStartDate, "MMM-yyyy")
   .TextMatrix(1, ETempDailyPnL.SalaryCost) = "Salary Cost From PnL of - " & Format(mCurrStartDate, "MMM-yyyy")
   .TextMatrix(1, ETempDailyPnL.SecurityCost) = "Security Cost From PnL of - " & Format(mCurrStartDate, "MMM-yyyy")
   .TextMatrix(1, ETempDailyPnL.InterestCharges) = "Interest Charges From Current Rent Agreement "
   .TextMatrix(1, ETempDailyPnL.CoExpense) = "CO Expense (current year figure)"
   .TextMatrix(1, ETempDailyPnL.NotionalRent) = "Notional Rent (Figures from CMG)"
   .TextMatrix(1, ETempDailyPnL.TotalExpense) = "Total Expense"
   .TextMatrix(1, ETempDailyPnL.NetIncome) = "Net Income"
   .TextMatrix(1, ETempDailyPnL.GStartDate) = "Godown Start Date"
   .TextMatrix(1, ETempDailyPnL.StateID) = "State ID"
   .TextMatrix(1, ETempDailyPnL.LocationID) = "LocationID"
   .TextMatrix(1, ETempDailyPnL.GodownID) = "Godown ID"
   
   .ColWidth(0) = 1500
   .ColWidth(1) = 1800
   .ColWidth(2) = 0
   .ColWidth(3) = 0
   .ColWidth(4) = 0
   .ColWidth(5) = 1800
   .ColWidth(6) = 1000
   .ColWidth(28) = 1800
   .ColWidth(29) = 1800
   
   .WordWrap = True
   .RowHeight(1) = 700
End With

End Sub

Private Sub CmdGetReport_Click()


Call Grid_Head
Call Grid_HeadGodownStock
Call Grid_HeadGodownIncome

SSTab1.Tab = 0

LblTotalRecords.Caption = "-"
MsgBox "Wait till next message!!!"

mYear = Year(TxtFromDate.Value)
'If MonthName(Month(TxtFromDate.Value)) = MonthName(1) Or CmbMonth = MonthName(2) Or CmbMonth = MonthName(3) Then
'   mYear = mYear - 1
'   mWideStartDate = mYear & "-" & 4 & "-1"
'ElseIf MonthName(Month(TxtFromDate.Value)) = MonthName(4) Or CmbMonth = MonthName(5) Or CmbMonth = MonthName(6) Or CmbMonth = MonthName(7) Or CmbMonth = MonthName(8) Or CmbMonth = MonthName(9) Or CmbMonth = MonthName(10) Or CmbMonth = MonthName(11) Or CmbMonth = MonthName(12) Then
'   mWideStartDate = mYear & "-" & 4 & "-1"
'End If

mWiseTotalDays = Day(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value), 0))
mWideEndDate = Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-" & mWiseTotalDays
mCurrStartDate = Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-1"
mStartDate = Format(CDate(mWideStartDate), "YYYY-MM-DD") 'mYear & "-" & mMonth & "-1"
mTotalDays = Day(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value) + 1, 0))
mEndDate = Year(mCurrStartDate) & "-" & Month(mCurrStartDate) & "-" & mTotalDays
'mTotalMonths = DateDiff("M", CDate(mWideStartDate), CDate(mWideEndDate)) + 1

tmp = "Select Distinct MG.GodownID"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & TxtFromDate.Value & "') And MG.StartDate <= '" & TxtFromDate.Value & "'"

If CmbStateID.Text <> "" Then
   tmp = tmp & " And MS.STateName = '" & CmbStateID.Text & "'"
Else
   If Gen_WcState <> "" Then
      tmp = tmp & " And MS." & Gen_WcState
   End If
End If

If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationName = '" & CmbLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

Call TmpTable

mGodownIDs = ""
For i = 1 To TmpRs.RecordCount
    If mGodownIDs = "" Then
       mGodownIDs = "(" & TmpRs!GodownID
    Else
       mGodownIDs = mGodownIDs & "," & TmpRs!GodownID
    End If
    TmpRs.MoveNext
Next

If mGodownIDs <> "" Then
   mGodownIDs = mGodownIDs & ")"
End If

Call TableTxn_CreateTempGSLForDailyPnL

tmp = "Select MS.StateName, MS.StateID, ML.LocationID, MG.GodownID, ML.LocationName, MCMP.CMPName, MG.GodownNo, MG.Address, MG.GodownCapacity, MUT.UnitType, MG.StartDate, MG.CloseDate"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & TxtFromDate.Value & "') And MG.StartDate <= '" & TxtFromDate.Value & "' And MG.GodownID in " & mGodownIDs & ""
tmp = tmp & " Order by MS.StateName,ML.LocationName "

Call TmpTable

Call TableTxn_IncomeExpenseForDailyPnL(" Where GodownID in " & mGodownIDs & " ")

'tmp = "Select GodownID,SUM(WtStock) 'WtStk',Sum(IncomewithoutReservation) 'IncomewithoutReservation' from  TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & "  Where GodownID in " & mGodownIDs & " Group by GodownID Order By GodownID "

tmp = " Select GodownID,Sum(WtStk) 'WtStk',Sum(RWtStk) 'RWtStk' ,Sum(IncomewithoutReservation) 'IncomewithoutReservation',Sum(IncomewithReservation) 'IncomewithReservation'  from ("
tmp = tmp & " Select GodownID,SUM(WtStock) 'WtStk',0 'RWtStk',Sum(IncomewithoutReservation) 'IncomewithoutReservation',0 'IncomewithReservation' "
tmp = tmp & " From TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & " "
tmp = tmp & " Where (ContractID is null or ContractID = 0) And GodownID in " & mGodownIDs & ""
tmp = tmp & " Group by GodownID"
tmp = tmp & " Union All"
tmp = tmp & " Select GodownID,0 'WtStk',SUM(WtStock) 'RWtStk',0 'IncomewithoutReservation' ,Sum(IncomewithoutReservation) 'IncomewithReservation' "
tmp = tmp & " From TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & " "
tmp = tmp & " Where ContractID is not null And ContractID <> 0 And GodownID in " & mGodownIDs & ""
tmp = tmp & " Group by GodownID) A Group By GodownID Order By GodownID "


Call Tmp1Table

tmp = " Select Distinct GodownID,ResQty,Rate From ("
tmp = tmp & " Select GodownID,Sum(ReservationQty) 'ResQty',Max(d.BillingRate) 'Rate' from Txn_SRGodownMapDetails a"
tmp = tmp & " Inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_StorageContract c on b.ContractID = c.ContractID"
tmp = tmp & " Inner Join Mst_SContractBillingDetail d On C.contractID = d.ContractID"
tmp = tmp & " Where d.BillingCycleID <> 5  And c.BookedAsPer <> 'Fixed' And c.Flag = 'A' And c.ExtendedPeriod >= '" & TxtFromDate.Value & "' And '" & TxtFromDate & "' Between d.FromDate and d.ToDate and GodownID in " & mGodownIDs & " "
tmp = tmp & " Group By GodownID"
tmp = tmp & " Having Sum(ReservationQty) > 0 "
tmp = tmp & " Union All "
tmp = tmp & " Select GodownID,Sum(ReservationQty) 'ResQty',Max(d.BillingRate) 'Rate' from Txn_SRGodownMapDetails a"
tmp = tmp & " Inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_StorageContract c on b.ContractID = c.ContractID"
tmp = tmp & " Inner Join Mst_SContractBillingDetail d On C.contractID = d.ContractID"
tmp = tmp & " Where d.BillingCycleID = 5 And c.BookedAsPer <> 'Fixed' And c.Flag = 'A' And c.ExtendedPeriod >= '" & TxtFromDate.Value & "' And '" & TxtFromDate & "' Between d.FromDate and d.ToDate and GodownID in " & mGodownIDs & " "
tmp = tmp & " Group By GodownID "
tmp = tmp & " Having Sum(ReservationQty) > 0) A "

Call Tmp2Table


tmp = " Select GodownID,Sum(ReservationQty) 'ResQty',Max(d.BillingRate) 'Rate' from Txn_SRGodownMapDetails a"
tmp = tmp & " Inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_StorageContract c on b.ContractID = c.ContractID"
tmp = tmp & " Inner Join Mst_SContractBillingDetail d On C.contractID = d.ContractID"
tmp = tmp & " Where c.BookedAsPer <> 'Fixed' And c.Flag = 'A' And c.ExtendedPeriod >= '" & TxtFromDate.Value & "' And '" & TxtFromDate & "' Between d.FromDate and d.ToDate and GodownID in " & mGodownIDs & " "
tmp = tmp & " Group By GodownID "
tmp = tmp & " Having Sum(ReservationQty) = 0 "

Call Tmp4Table


With MSHFlexGrid1

For i = 2 To TmpRs.RecordCount + 1
   .Rows = .Rows + 1
   .TextMatrix(i, ETempDailyPnL.RptDate) = Format(TxtFromDate.Value, "dd-mmm-yyyy")
   .TextMatrix(i, ETempDailyPnL.StateName) = TmpRs!StateName
   .TextMatrix(i, ETempDailyPnL.LocationName) = TmpRs!LocationName
   .TextMatrix(i, ETempDailyPnL.StateID) = TmpRs!StateID
   .TextMatrix(i, ETempDailyPnL.LocationID) = TmpRs!LocationID
   .TextMatrix(i, ETempDailyPnL.GodownID) = TmpRs!GodownID
   .TextMatrix(i, ETempDailyPnL.AMname) = ""
   .TextMatrix(i, ETempDailyPnL.RMName) = ""
   .TextMatrix(i, ETempDailyPnL.CMPName) = TmpRs!CMPName
   .TextMatrix(i, ETempDailyPnL.godownno) = TmpRs!godownno
   .TextMatrix(i, ETempDailyPnL.GodownCapacity) = TmpRs!GodownCapacity
   .TextMatrix(i, ETempDailyPnL.UnitType) = TmpRs!UnitType
   .TextMatrix(i, ETempDailyPnL.GStartDate) = Format(TmpRs!StartDate, "dd-mmm-yyyy")
   
   If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
   TmpRs1.Find "GodownID = " & Val(.TextMatrix(i, ETempDailyPnL.GodownID))
   If TmpRs1.EOF = False Then
      .TextMatrix(i, ETempDailyPnL.Stock) = IIf(IsNull(TmpRs1!WtStk), 0, TmpRs1!WtStk)
      .TextMatrix(i, ETempDailyPnL.RStock) = IIf(IsNull(TmpRs1!RWtStk), 0, TmpRs1!RWtStk)
      .TextMatrix(i, ETempDailyPnL.IncomewithoutReservation) = Format(TmpRs1!IncomewithoutReservation, "#######0")
      .TextMatrix(i, ETempDailyPnL.ReservationIncome) = Format(TmpRs1!IncomewithReservation, "#######0")
   Else
      .TextMatrix(i, ETempDailyPnL.Stock) = 0
      .TextMatrix(i, ETempDailyPnL.RStock) = 0
   End If
   'RStock
   If RsTxn_IncomeExpenseForDailyPnL.RecordCount > 0 Then RsTxn_IncomeExpenseForDailyPnL.MoveFirst
   RsTxn_IncomeExpenseForDailyPnL.Find "GodownID = " & Val(.TextMatrix(i, ETempDailyPnL.GodownID))
   If RsTxn_IncomeExpenseForDailyPnL.EOF = False Then
      If Val(.TextMatrix(i, ETempDailyPnL.IncomewithoutReservation)) = 0 Then
         mAvgRate = IIf(IsNull(RsTxn_IncomeExpenseForDailyPnL!AvgRate), 0, Format(RsTxn_IncomeExpenseForDailyPnL!AvgRate, "#######0"))
      End If
      .TextMatrix(i, ETempDailyPnL.CMinCMP) = IIf(IsNull(RsTxn_IncomeExpenseForDailyPnL!CMinCMP), 0, Format(RsTxn_IncomeExpenseForDailyPnL!CMinCMP, "#######0"))
      .TextMatrix(i, ETempDailyPnL.EmployeeClaim) = IIf(IsNull(RsTxn_IncomeExpenseForDailyPnL!EmployeeClaim), 0, Format(RsTxn_IncomeExpenseForDailyPnL!EmployeeClaim, "#######0"))
      .TextMatrix(i, ETempDailyPnL.VendorClaim) = IIf(IsNull(RsTxn_IncomeExpenseForDailyPnL!VendorClaim), 0, Format(RsTxn_IncomeExpenseForDailyPnL!VendorClaim, "#######0"))
      .TextMatrix(i, ETempDailyPnL.FumigationClaim) = IIf(IsNull(RsTxn_IncomeExpenseForDailyPnL!FumigationClaim), 0, Format(RsTxn_IncomeExpenseForDailyPnL!FumigationClaim, "#######0"))
      .TextMatrix(i, ETempDailyPnL.FranchiseeFees) = IIf(IsNull(RsTxn_IncomeExpenseForDailyPnL!FranchiseeFees), 0, Format(RsTxn_IncomeExpenseForDailyPnL!FranchiseeFees, "#######0"))
      .TextMatrix(i, ETempDailyPnL.SalaryCost) = IIf(IsNull(RsTxn_IncomeExpenseForDailyPnL!SalaryCost), 0, Format(RsTxn_IncomeExpenseForDailyPnL!SalaryCost, "#######0"))
      .TextMatrix(i, ETempDailyPnL.SecurityCost) = IIf(IsNull(RsTxn_IncomeExpenseForDailyPnL!SecurityCost), 0, Format(RsTxn_IncomeExpenseForDailyPnL!SecurityCost, "#######0"))
   Else
      '.TextMatrix(I, ETempDailyPnL.AvgRate) = 0
       tmp = "  Select l.LocationID,Avg(AvgRate) 'AvgRate' from Txn_IncomeExpenseForDailyPnL a"
       tmp = tmp & " inner join Mst_Godown g On a.GodownID = g.GodownID"
       tmp = tmp & " Inner join Mst_CMP b on g.CMPID = b.CMPID"
       tmp = tmp & " Inner Join Mst_Location l on b.LocationID = l.LocationID"
       tmp = tmp & " Where l.LocationID = " & .TextMatrix(i, ETempDailyPnL.LocationID) & " "
       tmp = tmp & " Group by l.LocationID"
       Call Tmp6Table
       
       If TmpRs6.RecordCount > 0 Then
          If Val(.TextMatrix(i, ETempDailyPnL.IncomewithoutReservation)) = 0 Then
             mAvgRate = IIf(IsNull(TmpRs6!AvgRate), 0, Format(TmpRs6!AvgRate, "#######0"))
          End If
       Else
          mAvgRate = 0
       End If
      
      .TextMatrix(i, ETempDailyPnL.CMinCMP) = 0
      .TextMatrix(i, ETempDailyPnL.EmployeeClaim) = 0
      .TextMatrix(i, ETempDailyPnL.VendorClaim) = 0
      .TextMatrix(i, ETempDailyPnL.FumigationClaim) = 0
      .TextMatrix(i, ETempDailyPnL.FranchiseeFees) = 0
      .TextMatrix(i, ETempDailyPnL.SalaryCost) = 0
      .TextMatrix(i, ETempDailyPnL.SecurityCost) = 0
   End If
   
   If IsNull(.TextMatrix(i, ETempDailyPnL.IncomewithoutReservation)) = True Or Val(.TextMatrix(i, ETempDailyPnL.IncomewithoutReservation)) = 0 Then
      .TextMatrix(i, ETempDailyPnL.IncomewithoutReservation) = Val(mAvgRate) * Val(.TextMatrix(i, ETempDailyPnL.Stock))
      .row = i
      If .row <> 0 Then
         For o = 0 To .Cols - 1
             .Col = o
             .CellBackColor = vbRed
         Next
      End If
   End If
   
   If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
   If TmpRs4.RecordCount > 0 Then TmpRs4.MoveFirst
   TmpRs2.Find "GodownID = " & Val(.TextMatrix(i, ETempDailyPnL.GodownID))
   If TmpRs2.EOF = False Then
      If IsNull(TmpRs2!ResQty) = True Or TmpRs2!ResQty = 0 Then
      Else
        .TextMatrix(i, ETempDailyPnL.ReservationIncome) = Val(.TextMatrix(i, ETempDailyPnL.ReservationIncome)) + IIf(IsNull(TmpRs2!ResQty), 0, TmpRs2!ResQty) * IIf(IsNull(TmpRs2!Rate), 0, TmpRs2!Rate)
      End If
   Else
'Commented by Parag
'      .TextMatrix(I, ETempDailyPnL.ReservationIncome) = 0
'      TmpRs4.Find "GodownID = " & Val(.TextMatrix(I, ETempDailyPnL.GodownID))
'      If TmpRs4.EOF = False Then
'         '.TextMatrix(I, ETempDailyPnL.ReservationIncome) = Val(.TextMatrix(I, ETempDailyPnL.ReservationIncome)) + IIf(IsNull(TmpRs4!Rate), 0, TmpRs4!Rate) * Val(.TextMatrix(I, ETempDailyPnL.Stock))
'      End If
   End If
   .TextMatrix(i, ETempDailyPnL.TotalIncome) = Val(.TextMatrix(i, ETempDailyPnL.IncomewithoutReservation)) + Val(.TextMatrix(i, ETempDailyPnL.ReservationIncome)) + Val(.TextMatrix(i, ETempDailyPnL.CMinCMP))
   
   TmpRs.MoveNext
   
Next


End With


Call GetAMRMName
Call GetFixedReservationIncome
Call GetNotionalExpense
Call GetCoExpense
Call GetRentExpense
Call GetOutstandingInterest


With MSHFlexGrid1
     For i = 2 To .Rows - 2
         .TextMatrix(i, ETempDailyPnL.TotalExpense) = Format(Val(.TextMatrix(i, ETempDailyPnL.EmployeeClaim)) + Val(.TextMatrix(i, ETempDailyPnL.VendorClaim)) + Val(.TextMatrix(i, ETempDailyPnL.FumigationClaim)) + Val(.TextMatrix(i, ETempDailyPnL.FranchiseeFees)) + Val(.TextMatrix(i, ETempDailyPnL.SalaryCost)) + Val(.TextMatrix(i, ETempDailyPnL.SecurityCost)) + Val(.TextMatrix(i, ETempDailyPnL.CoExpense)) + Val(.TextMatrix(i, ETempDailyPnL.NotionalRent)) + Val(.TextMatrix(i, ETempDailyPnL.LeaseRent)) + Val(.TextMatrix(i, ETempDailyPnL.InterestCharges)), "#######0")
         .TextMatrix(i, ETempDailyPnL.NetIncome) = Format(Val(.TextMatrix(i, ETempDailyPnL.TotalIncome)) - Val(.TextMatrix(i, ETempDailyPnL.TotalExpense)), "#######0")
     Next
End With


With MSHFlexGrid1
    .TextMatrix(0, ETempDailyPnL.CMinCMP) = 0
    .TextMatrix(0, ETempDailyPnL.EmployeeClaim) = 0
    .TextMatrix(0, ETempDailyPnL.VendorClaim) = 0
    .TextMatrix(0, ETempDailyPnL.FumigationClaim) = 0
    .TextMatrix(0, ETempDailyPnL.FranchiseeFees) = 0
    .TextMatrix(0, ETempDailyPnL.SalaryCost) = 0
    .TextMatrix(0, ETempDailyPnL.SecurityCost) = 0
    .TextMatrix(0, ETempDailyPnL.ReservationIncome) = 0
    .TextMatrix(0, ETempDailyPnL.CoExpense) = 0
    .TextMatrix(0, ETempDailyPnL.NotionalRent) = 0
    .TextMatrix(0, ETempDailyPnL.LeaseRent) = 0
    .TextMatrix(0, ETempDailyPnL.InterestCharges) = 0
    .TextMatrix(0, ETempDailyPnL.IncomewithoutReservation) = 0
    .TextMatrix(0, ETempDailyPnL.TotalIncome) = 0
    .TextMatrix(0, ETempDailyPnL.TotalExpense) = 0
    .TextMatrix(0, ETempDailyPnL.NetIncome) = 0
End With

With MSHFlexGrid1
     For i = 2 To .Rows - 2
         .TextMatrix(0, ETempDailyPnL.CMinCMP) = Val(.TextMatrix(0, ETempDailyPnL.CMinCMP)) + Val(.TextMatrix(i, ETempDailyPnL.CMinCMP))
         .TextMatrix(0, ETempDailyPnL.EmployeeClaim) = Val(.TextMatrix(0, ETempDailyPnL.EmployeeClaim)) + Val(.TextMatrix(i, ETempDailyPnL.EmployeeClaim))
         .TextMatrix(0, ETempDailyPnL.VendorClaim) = Val(.TextMatrix(0, ETempDailyPnL.VendorClaim)) + Val(.TextMatrix(i, ETempDailyPnL.VendorClaim))
         .TextMatrix(0, ETempDailyPnL.FumigationClaim) = Val(.TextMatrix(0, ETempDailyPnL.FumigationClaim)) + Val(.TextMatrix(i, ETempDailyPnL.FumigationClaim))
         .TextMatrix(0, ETempDailyPnL.FranchiseeFees) = Val(.TextMatrix(0, ETempDailyPnL.FranchiseeFees)) + Val(.TextMatrix(i, ETempDailyPnL.FranchiseeFees))
         .TextMatrix(0, ETempDailyPnL.SalaryCost) = Val(.TextMatrix(0, ETempDailyPnL.SalaryCost)) + Val(.TextMatrix(i, ETempDailyPnL.SalaryCost))
         .TextMatrix(0, ETempDailyPnL.SecurityCost) = Val(.TextMatrix(0, ETempDailyPnL.SecurityCost)) + Val(.TextMatrix(i, ETempDailyPnL.SecurityCost))
         .TextMatrix(0, ETempDailyPnL.ReservationIncome) = Val(.TextMatrix(0, ETempDailyPnL.ReservationIncome)) + Val(.TextMatrix(i, ETempDailyPnL.ReservationIncome))
         .TextMatrix(0, ETempDailyPnL.CoExpense) = Val(.TextMatrix(0, ETempDailyPnL.CoExpense)) + Val(.TextMatrix(i, ETempDailyPnL.CoExpense))
         .TextMatrix(0, ETempDailyPnL.NotionalRent) = Val(.TextMatrix(0, ETempDailyPnL.NotionalRent)) + Val(.TextMatrix(i, ETempDailyPnL.NotionalRent))
         .TextMatrix(0, ETempDailyPnL.LeaseRent) = Val(.TextMatrix(0, ETempDailyPnL.LeaseRent)) + Val(.TextMatrix(i, ETempDailyPnL.LeaseRent))
         .TextMatrix(0, ETempDailyPnL.InterestCharges) = Val(.TextMatrix(0, ETempDailyPnL.InterestCharges)) + Val(.TextMatrix(i, ETempDailyPnL.InterestCharges))
         .TextMatrix(0, ETempDailyPnL.IncomewithoutReservation) = Val(.TextMatrix(0, ETempDailyPnL.IncomewithoutReservation)) + Val(.TextMatrix(i, ETempDailyPnL.IncomewithoutReservation))
         .TextMatrix(0, ETempDailyPnL.TotalIncome) = Val(.TextMatrix(0, ETempDailyPnL.TotalIncome)) + Val(.TextMatrix(i, ETempDailyPnL.TotalIncome))
         .TextMatrix(0, ETempDailyPnL.TotalExpense) = Val(.TextMatrix(0, ETempDailyPnL.TotalExpense)) + Val(.TextMatrix(i, ETempDailyPnL.TotalExpense))
         .TextMatrix(0, ETempDailyPnL.NetIncome) = Val(.TextMatrix(0, ETempDailyPnL.NetIncome)) + Val(.TextMatrix(i, ETempDailyPnL.NetIncome))
    Next
End With

For mRow = 2 To MSHFlexGrid1.Rows - 2
    MSHFlexGrid1.row = mRow
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.Col = C
        If C = 14 Or C = 16 Or C = 23 Or C = 24 Then
           MSHFlexGrid1.CellBackColor = vbGreen
        ElseIf C = 15 Or C = 18 Or C = 25 Or C = 26 Or C = 27 Then
           MSHFlexGrid1.CellBackColor = &HC0FF&
        ElseIf C = 19 Or C = 20 Or C = 21 Or C = 22 Then
           MSHFlexGrid1.CellBackColor = vbRed
        ElseIf C = 17 Or C = 28 Or C = 29 Then
           MSHFlexGrid1.CellBackColor = &HF3EEDB
        End If
    Next
Next
             
CmdExcel.Enabled = True

tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & "') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
tmp = tmp & " drop table dbo.TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ""
Call TmpTable

LblTotalRecords.Caption = MSHFlexGrid1.Rows - 3


If ChkRateDetails.Value = vbChecked Then

   tmp = "Select MAX(MonthYear) 'MonthYear' from Txn_MonGodownWiseProfitability"
   Call TmpTable
    
    If TmpRs.RecordCount > 0 Then
       mCurrStartDate = Year(TmpRs!MonthYear) & "-" & Month(TmpRs!MonthYear) & "-1"
       

       tmp = " Select s.StateName,l.LocationName,c.CMPName,g.GodownNo,g.Address,a.GodownID,SUM(WtQty) 'TotalStock' from Txn_InvMonGSLDetail a"
       tmp = tmp & " Inner Join Mst_Godown g On a.GodownID = g.GodownID"
       tmp = tmp & " Inner Join Mst_CMP c On g.CMPID = c.CMPID"
       tmp = tmp & " Inner Join Mst_Location l On c.LocationID = l.LocationID"
       tmp = tmp & " Inner Join Mst_State s On l.StateID = s.StateID"
       tmp = tmp & " Where a.YearMonth = '" & mCurrStartDate & "' And (a.ContractID is null Or a.ContractID = 0)"
       tmp = tmp & " Group By s.StateName,l.LocationName,c.CMPName,g.GodownNo,g.Address,a.GodownID"
       tmp = tmp & " Order By s.StateName,l.LocationName,c.CMPName,g.GodownNo,g.Address,a.GodownID"
       
       Call TmpTable
                 
       For i = 1 To TmpRs.RecordCount
           MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
           For C = 0 To MSHFlexGrid3.Cols - 1
               MSHFlexGrid3.TextMatrix(i, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
           Next
           TmpRs.MoveNext
       Next
       
       tmp = "Select s.StateName,l.LocationName,c.CMPName,g.GodownNo,g.Address,g.GodownID,NNNRIncomeAdhoc,NNNRIncome,"
       tmp = tmp & " SevenDaysIncome , DPIncomeAdHoc, DPIncomeInvoice, RejectedIncome, SpotIncomeAdHoc, OtherIncome, SCMIncome"
       tmp = tmp & " from Txn_MonGodownWiseProfitability a"
       tmp = tmp & " Inner Join Mst_Godown g On a.GodownID = g.GodownID"
       tmp = tmp & " Inner Join Mst_CMP c On g.CMPID = c.CMPID"
       tmp = tmp & " Inner Join Mst_Location l On c.LocationID = l.LocationID"
       tmp = tmp & " Inner Join Mst_State s On l.StateID = s.StateID"
       tmp = tmp & " Where MonthYear = '" & mCurrStartDate & "' And (NcdexStock > 0 Or NNStock > 0 )"
       tmp = tmp & " Order By s.StateName,l.LocationName,c.CMPName,g.GodownNo,g.Address,g.GodownID"
     
       Call TmpTable
       For i = 1 To TmpRs.RecordCount
           MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
           For C = 0 To MSHFlexGrid4.Cols - 1
               MSHFlexGrid4.TextMatrix(i, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
           Next
           TmpRs.MoveNext
       Next
    End If
End If


MsgBox "Done"

Exit Sub

msgError:
    
    tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & "') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
    tmp = tmp & " drop table dbo.TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ""
    Call TmpTable
    
    'MsgBox "Some problem occured during process!!!"
    
    MsgBox Err.Description, vbExclamation, "Error"


End Sub
Private Sub GetCoExpense()

'   .TextMatrix(0, 10) = "CoExpense"

   mCountMonth = 1
   tmp = " Select sum(ExpenseAmt * " & mCountMonth & ")/12 As IncomeAmt "
   tmp = tmp & " From Mst_CoExpense MCE"
   tmp = tmp & " Where '" & mCurrStartDate & "' Between MCE.FromDate And MCE.Todate"
   Call Tmp1Table

   'mTotalIncome = Val(MSHFlexGrid1.TextMatrix(0, E.TotalIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(0, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.Salary))
'mTotalCapacity
   If TmpRs1.RecordCount > 0 Then
'      For I = 1 To MSHFlexGrid2.Rows - 1
'         MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format((TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(I, 2))) / mTotalCapacity, "#############.00"))
'      Next
   
      mTotalCapacity = 0
      For i = 2 To MSHFlexGrid1.Rows - 2
          If LCase(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.UnitType)) <> LCase("Stock Management") Then
             mTotalCapacity = mTotalCapacity + Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownCapacity))
          End If
      Next
      
      If mTotalCapacity > 0 Then
         For i = 2 To MSHFlexGrid1.Rows - 2
             If LCase(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.UnitType)) <> LCase("Stock Management") Then
                MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.CoExpense) = Format(Val(TmpRs1!IncomeAmt) * Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownCapacity)) / mTotalCapacity, "########0")
             End If
         Next i
      End If
   End If

End Sub
Private Sub GetNotionalExpense()
'tmp = "Select LocationID,NotionalExp From Txn_LocationwiseNotionalExpense"
tmp = "Select Distinct LocationID From Txn_LocationwiseNotionalExpense "
Call Tmp2Table


If TmpRs2.RecordCount > 0 Then
   For x = 1 To TmpRs2.RecordCount
       mTotalCapacity = 0
       For i = 2 To MSHFlexGrid1.Rows - 2
           If TmpRs2!LocationID = (MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LocationID)) Then
              mTotalCapacity = mTotalCapacity + Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownCapacity))
           End If
       Next
       If mTotalCapacity > 0 Then
          tmp = "Select LocationID,NotionalExp From Txn_LocationwiseNotionalExpense Where LocationID = " & TmpRs2!LocationID & " "
          Call Tmp3Table
          For i = 2 To MSHFlexGrid1.Rows - 2
              TmpRs3.MoveFirst
              TmpRs3.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LocationID))
               'If Val(MSHFlexGrid1.TextMatrix(I, E.LocationID)) = 111 Then MsgBox "stop"
              If Not TmpRs3.EOF Then
                 MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.NotionalRent) = Format(Val(TmpRs3!NotionalExp) * Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownCapacity)) / mTotalCapacity, "########0")
              End If
          Next i
       End If
       TmpRs2.MoveNext
   Next
End If

End Sub

Private Sub GetFixedReservationIncome()
'tmp = "Select LocationID,NotionalExp From Txn_LocationwiseNotionalExpense"
'tmp = "Select Distinct LocationID From Txn_LocationwiseNotionalExpense "
'Call Tmp2Table

tmp = " Select Distinct c.LocationID,Sum(ReservationQty) 'TotResQty' from Txn_SRGodownMapDetails a" ','1' 'ResQty',Max(d.BillingRate) 'Rate'
tmp = tmp & " Inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_StorageContract c on b.ContractID = c.ContractID"
tmp = tmp & " Inner Join Mst_SContractBillingDetail d On C.contractID = d.ContractID"
tmp = tmp & " Where d.BillingCycleID = 5 And c.BookedAsPer = 'Fixed' And c.Flag = 'A' And c.ExtendedPeriod >= '" & TxtFromDate.Value & "' And '" & TxtFromDate.Value & "' Between d.FromDate and d.ToDate and GodownID in " & mGodownIDs & " "
tmp = tmp & " Group By c.LocationID"

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   For x = 1 To TmpRs2.RecordCount
       mTotalCapacity = TmpRs2!TotResQty
       If mTotalCapacity > 0 Then
          'tmp = "Select LocationID,NotionalExp From Txn_LocationwiseNotionalExpense Where LocationID = " & TmpRs2!LocationID & " "
          
          tmp = " Select c.locationid,a.GodownID,SUM(ReservationQty) 'ResQty' ,Max(d.BillingRate) 'Rate' from Txn_SRGodownMapDetails a"
          tmp = tmp & " Inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID"
          tmp = tmp & " Inner Join Mst_StorageContract c on b.ContractID = c.ContractID"
          tmp = tmp & " Inner Join Mst_SContractBillingDetail d On C.contractID = d.ContractID"
          tmp = tmp & " Where c.locationid = " & TmpRs2!LocationID & " And d.BillingCycleID = 5 And c.BookedAsPer = 'Fixed' And c.Flag = 'A' And c.ExtendedPeriod >= '" & TxtFromDate.Value & "' And '" & TxtFromDate.Value & "' Between d.FromDate and d.ToDate and GodownID in " & mGodownIDs & " "
          tmp = tmp & " Group By c.locationid,a.GodownID"
          
          Call Tmp3Table
          
          For i = 2 To MSHFlexGrid1.Rows - 2
              TmpRs3.MoveFirst
              TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownID))
               'If Val(MSHFlexGrid1.TextMatrix(I, E.LocationID)) = 111 Then MsgBox "stop"
              If Not TmpRs3.EOF Then
                    
                 MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.ReservationIncome)) + Format(Val(TmpRs3!Rate) * Val(TmpRs3!ResQty) / mTotalCapacity, "########0")
                 MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.TotalIncome) = Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.TotalIncome)) + Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.ReservationIncome))
              End If
          Next i
       Else
          For i = 2 To MSHFlexGrid1.Rows - 2
              If TmpRs2!LocationID = (MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LocationID)) Then
                 mTotalCapacity = mTotalCapacity + Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownCapacity))
              End If
          Next
          If mTotalCapacity > 0 Then
             tmp = " Select c.locationid,a.GodownID,SUM(ReservationQty) 'ResQty' ,Max(d.BillingRate) 'Rate' from Txn_SRGodownMapDetails a"
             tmp = tmp & " Inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID"
             tmp = tmp & " Inner Join Mst_StorageContract c on b.ContractID = c.ContractID"
             tmp = tmp & " Inner Join Mst_SContractBillingDetail d On C.contractID = d.ContractID"
             tmp = tmp & " Where c.locationid = " & TmpRs2!LocationID & " And d.BillingCycleID = 5 And c.BookedAsPer = 'Fixed' And c.Flag = 'A' And c.ExtendedPeriod >= '" & TxtFromDate.Value & "' And '" & TxtFromDate.Value & "' Between d.FromDate and d.ToDate and GodownID in " & mGodownIDs & " "
             tmp = tmp & " Group By c.locationid,a.GodownID"
             
             Call Tmp3Table
            
             For i = 2 To MSHFlexGrid1.Rows - 2
                 TmpRs3.MoveFirst
                 TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownID))
                 'If Val(MSHFlexGrid1.TextMatrix(I, E.LocationID)) = 111 Then MsgBox "stop"
                 If Not TmpRs3.EOF Then
                    MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.ReservationIncome)) + Format(Val(TmpRs3!Rate) * Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownCapacity)) / mTotalCapacity, "########0")
                    MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.TotalIncome) = Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.TotalIncome)) + Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.ReservationIncome))
                 End If
             Next i
          End If
       End If
       TmpRs2.MoveNext
   Next
End If

End Sub



Private Sub CmdProcess_Click()

Call Grid_HeadIncomeExpense

mYear = Year(TxtFromDate.Value)

tmp = "Select MAX(MonthYear) 'MonthYear' from Txn_MonGodownWiseProfitability"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   
   If MonthName(Month(TmpRs!MonthYear)) = MonthName(1) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(2) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(3) Then
      mYear = mYear - 1
      mWideStartDate = mYear & "-" & 4 & "-1"
   ElseIf MonthName(Month(TmpRs!MonthYear)) = MonthName(4) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(5) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(6) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(7) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(8) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(9) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(10) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(11) Or MonthName(Month(TmpRs!MonthYear)) = MonthName(12) Then
      mWideStartDate = mYear & "-" & 4 & "-1"
   End If
   
   mWiseTotalDays = Day(DateSerial(Year(TmpRs!MonthYear), Month(TmpRs!MonthYear) + 1, 0))
   mWideEndDate = Year(TmpRs!MonthYear) & "-" & Month(TmpRs!MonthYear) & "-" & mWiseTotalDays
   mCurrStartDate = Year(TmpRs!MonthYear) & "-" & Month(TmpRs!MonthYear) & "-1"
   mStartDate = Format(CDate(mWideStartDate), "YYYY-MM-DD") 'mYear & "-" & mMonth & "-1"
   mTotalDays = Day(DateSerial(Year(mStartDate), Month(mStartDate) + 1, 0))
   mEndDate = Year(mStartDate) & "-" & Month(mStartDate) & "-" & mTotalDays
   mTotalMonths = DateDiff("M", mWideStartDate, mWideEndDate) + 1
   'mYearMonth = CmbMonth.Text & "-" & CmbYear.Text
      
   tmp = "Select Distinct GodownID,Capacity,LocationID from Txn_MonGodownWiseProfitability Where MonthYear = '" & mCurrStartDate & " '"
   Call TmpTable
'   .TextMatrix(0, 0) = "GodownID"
   mTotalCapacity = 0
   For i = 1 To TmpRs.RecordCount
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       MSHFlexGrid2.TextMatrix(i, 0) = TmpRs!GodownID
       MSHFlexGrid2.TextMatrix(i, 2) = TmpRs!Capacity
       mTotalCapacity = mTotalCapacity + Val(MSHFlexGrid2.TextMatrix(i, 2))
       MSHFlexGrid2.TextMatrix(i, 12) = TmpRs!LocationID
       TmpRs.MoveNext
   Next
   
   
'   .TextMatrix(0, 1) = "AvgRate"
'   tmp = "Select GodownID,SUM(BalanceWeight) 'TotalStock' from Txn_InvMonGSLDetail Where YearMonth = '" & mCurrStartDate & "' And (ContractID is null Or ContractID = 0)"
   
   'Query for fetching Weigthed Qty from latest previous monthly GSL data data excluding reservation
   tmp = "Select GodownID,SUM(WtQty) 'TotalStock' from Txn_InvMonGSLDetail Where YearMonth = '" & mCurrStartDate & "' And (ContractID is null Or ContractID = 0)"
   tmp = tmp & " Group By GodownID "
   
   Call Tmp1Table
      
   
   'Query for fetching income from latest previous PnL data excluding reservation and CM in CMP
   tmp = " Select LocationID,GodownID,(Sum(NNNRIncomeAdhoc)+Sum(NNNRIncome)+Sum(SevenDaysIncome)+Sum(DPIncomeAdHoc)+Sum(DPIncomeInvoice)+Sum(RejectedIncome)+"
   tmp = tmp & " Sum(SpotIncomeAdHoc)+Sum(OtherIncome)+Sum(SCMIncome)) 'TotalIncome' " '-Sum(PartiallyReversedIncome))  'TotalIncome'"
   tmp = tmp & " from Txn_MonGodownWiseProfitability Where MonthYear = '" & mCurrStartDate & " ' And (NcdexStock > 0 Or NNStock > 0 )  Group By LocationId,GodownID"
 
   Call TmpTable
   
   
   If TmpRs.RecordCount > 0 Then
      For i = 1 To MSHFlexGrid2.Rows - 1
          If Val(MSHFlexGrid2.TextMatrix(i, 0)) > 0 Then
             TmpRs.MoveFirst
             TmpRs.Find "GodownID = " & Val(MSHFlexGrid2.TextMatrix(i, 0))
             If TmpRs.EOF = False Then
                TmpRs1.MoveFirst
                TmpRs1.Find "GodownID = " & TmpRs!GodownID
                If TmpRs1.EOF = False Then
'Calculating Avg rate from Income and WtQty data
                   MSHFlexGrid2.TextMatrix(i, 1) = IIf(IsNull(TmpRs!TotalIncome), 0, Format(TmpRs!TotalIncome / TmpRs1!TotalStock, "########0"))
                   If Val(MSHFlexGrid2.TextMatrix(i, 1)) < 0 Then
'Allocating Location Avg rate to Godown if Godown is not having Income in previous month
                      tmp = "  Select l.LocationID,Avg(AvgRate) 'AvgRate' from Txn_IncomeExpenseForDailyPnL a"
                      tmp = tmp & " inner join Mst_Godown g On a.GodownID = g.GodownID"
                      tmp = tmp & " Inner join Mst_CMP b on g.CMPID = b.CMPID"
                      tmp = tmp & " Inner Join Mst_Location l on b.LocationID = l.LocationID"
                      tmp = tmp & " Where a.GodownID = " & Val(MSHFlexGrid2.TextMatrix(i, 0)) & ""
                      tmp = tmp & " Group by l.LocationID"
                      Call Tmp2Table
                      If TmpRs2.RecordCount > 0 Then
                         MSHFlexGrid2.TextMatrix(i, 1) = IIf(IsNull(TmpRs2!AvgRate), 0, TmpRs2!AvgRate)
                      End If
                   End If
                Else
'                   MSHFlexGrid2.TextMatrix(I, 1) = "0"
'Allocating Location Avg rate to Godown if Godown is not having Income in previous month
                   tmp = "  Select l.LocationID,Avg(AvgRate) 'AvgRate' from Txn_IncomeExpenseForDailyPnL a"
                   tmp = tmp & " inner join Mst_Godown g On a.GodownID = g.GodownID"
                   tmp = tmp & " Inner join Mst_CMP b on g.CMPID = b.CMPID"
                   tmp = tmp & " Inner Join Mst_Location l on b.LocationID = l.LocationID"
                   tmp = tmp & " Where l.LocationID = " & Val(MSHFlexGrid2.TextMatrix(i, 12)) & ""
                   tmp = tmp & " Group by l.LocationID"
                   Call Tmp2Table
                   If TmpRs2.RecordCount > 0 Then
                      MSHFlexGrid2.TextMatrix(i, 1) = IIf(IsNull(TmpRs2!AvgRate), 0, TmpRs2!AvgRate)
                   End If
                End If
             Else
'                MSHFlexGrid2.TextMatrix(I, 1) = "0"
'Allocating Location Avg rate to Godown if Godown is not having Income in previous month
                tmp = "  Select l.LocationID,Avg(AvgRate) 'AvgRate' from Txn_IncomeExpenseForDailyPnL a"
                tmp = tmp & " inner join Mst_Godown g On a.GodownID = g.GodownID"
                tmp = tmp & " Inner join Mst_CMP b on g.CMPID = b.CMPID"
                tmp = tmp & " Inner Join Mst_Location l on b.LocationID = l.LocationID"
                tmp = tmp & " Where l.LocationID = " & Val(MSHFlexGrid2.TextMatrix(i, 12)) & ""
                tmp = tmp & " Group by l.LocationID"
                Call Tmp2Table
                If TmpRs2.RecordCount > 0 Then
                   MSHFlexGrid2.TextMatrix(i, 1) = IIf(IsNull(TmpRs2!AvgRate), 0, TmpRs2!AvgRate)
                End If
             End If
          End If
      Next
   End If


'   .TextMatrix(0, 3) = "CMinCMP"

   tmp = "Select LocationID,GodownID,(Sum(CMUnbilledIncome)+Sum(CMPendingIncome)+Sum(CMApprovedIncome)) 'CMinCMP'"
   tmp = tmp & " from Txn_MonGodownWiseProfitability Where MonthYear = '" & mCurrStartDate & "'"
   tmp = tmp & " Group By LocationId,GodownID"
   tmp = tmp & " Having (Sum(CMUnbilledIncome) + Sum(CMPendingIncome) + Sum(CMApprovedIncome)) > 0"
   
   Call TmpTable

   If TmpRs.RecordCount > 0 Then
      For i = 1 To MSHFlexGrid2.Rows - 1
          If Val(MSHFlexGrid2.TextMatrix(i, 0)) > 0 Then
             TmpRs.MoveFirst
             TmpRs.Find " GodownID = " & Val(MSHFlexGrid2.TextMatrix(i, 0))
             If TmpRs.EOF = False Then
                MSHFlexGrid2.TextMatrix(i, 3) = IIf(IsNull(TmpRs!CMinCMP), 0, Format(TmpRs!CMinCMP, "########0"))
             Else
                MSHFlexGrid2.TextMatrix(i, 3) = "0"
             End If
          End If
      Next
   End If


'   .TextMatrix(0, 4) = "EmployeeClaim" '   .TextMatrix(0, 5) = "VendorClaim" '   .TextMatrix(0, 6) = "FumigationClaim" '   .TextMatrix(0, 7) = "FranchiseeFees"
'Fetching data of expenses (Employee , Vendor , Fumigation, Frachisee as avg from Financial Year Starting till last PnL generated
   tmp = " Select LocationId,GodownID,"
   tmp = tmp & " (Sum(EmpClaimFAG)+ Sum(EmpClaimAccrual))/" & mTotalMonths & " 'EmpClaim',"
   tmp = tmp & " (Sum(VendorExpenses)+ Sum(VendorExpensesAccrued))/" & mTotalMonths & " 'VendorClaim',"
   tmp = tmp & " (Sum(FumigationFeesVendor)+ Sum(FumigationFeesPaid)+ Sum(FumigationFeesAccrual))/" & mTotalMonths & " 'FumigationExp',"
   tmp = tmp & " (Sum(FranchiseeFeesPaid) +  Sum(FranchiseeFeesFAG) + Sum(FranchiseeFeesCMG))/" & mTotalMonths & " 'FranchiseeFees'"
   tmp = tmp & " From Txn_MonGodownWiseProfitability Where MonthYear between '" & mWideStartDate & "' And '" & mWideEndDate & "' Group by LocationId,GodownID"

   Call TmpTable

   If TmpRs.RecordCount > 0 Then
      For i = 1 To MSHFlexGrid2.Rows - 1
          If Val(MSHFlexGrid2.TextMatrix(i, 0)) > 0 Then
             TmpRs.MoveFirst
             TmpRs.Find "GodownID = " & Val(MSHFlexGrid2.TextMatrix(i, 0))
             If TmpRs.EOF = False Then
                MSHFlexGrid2.TextMatrix(i, 4) = IIf(IsNull(TmpRs!EmpClaim), 0, Format(TmpRs!EmpClaim, "########0"))
                MSHFlexGrid2.TextMatrix(i, 5) = IIf(IsNull(TmpRs!VendorClaim), 0, Format(TmpRs!VendorClaim, "########0"))
                MSHFlexGrid2.TextMatrix(i, 6) = IIf(IsNull(TmpRs!FumigationExp), 0, Format(TmpRs!FumigationExp, "########0"))
                MSHFlexGrid2.TextMatrix(i, 7) = IIf(IsNull(TmpRs!FranchiseeFees), 0, Format(TmpRs!FranchiseeFees, "########0"))
             Else
                MSHFlexGrid2.TextMatrix(i, 4) = "0"
                MSHFlexGrid2.TextMatrix(i, 5) = "0"
                MSHFlexGrid2.TextMatrix(i, 6) = "0"
                MSHFlexGrid2.TextMatrix(i, 7) = "0"
             End If
          End If
      Next
   End If

'   .TextMatrix(0, 8) = "SalaryCost" '   .TextMatrix(0, 9) = "SecurityCost"
'Fetching data of expenses (Salary Cost , Security from just previous PnL generated

   tmp = "Select LocationId,GodownID,Sum(SecurityFeesPaid)+Sum(SecurityFeesAccrual) 'SecurityCost',Sum(Salary) + Sum(SalaryFieldStaff)'SalaryCost'"
   tmp = tmp & " From Txn_MonGodownWiseProfitability Where MonthYear = '" & mCurrStartDate & "' Group by LocationId,GodownID"
   
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      For i = 1 To MSHFlexGrid2.Rows - 1
          If Val(MSHFlexGrid2.TextMatrix(i, 0)) > 0 Then
             TmpRs.MoveFirst
             TmpRs.Find " GodownID = " & Val(MSHFlexGrid2.TextMatrix(i, 0))
             If TmpRs.EOF = False Then
                MSHFlexGrid2.TextMatrix(i, 8) = IIf(IsNull(TmpRs!SalaryCost), 0, Format(TmpRs!SalaryCost, "########0"))
                MSHFlexGrid2.TextMatrix(i, 9) = IIf(IsNull(TmpRs!SecurityCost), 0, Format(TmpRs!SecurityCost, "########0"))
             Else
                MSHFlexGrid2.TextMatrix(i, 8) = "0"
                MSHFlexGrid2.TextMatrix(i, 9) = "0"
             End If
          End If
      Next
   End If

End If

tmp = "Delete from Txn_IncomeExpenseForDailyPnL "
Call TmpTable

Call TableTxn_IncomeExpenseForDailyPnL

For i = 1 To MSHFlexGrid2.Rows - 1
    If Val(MSHFlexGrid2.TextMatrix(i, 0)) > 0 Then
       RsTxn_IncomeExpenseForDailyPnL.AddNew
       RsTxn_IncomeExpenseForDailyPnL!GodownID = Val(MSHFlexGrid2.TextMatrix(i, 0))
       RsTxn_IncomeExpenseForDailyPnL!AvgRate = Val(MSHFlexGrid2.TextMatrix(i, 1))
       RsTxn_IncomeExpenseForDailyPnL!CMinCMP = Val(MSHFlexGrid2.TextMatrix(i, 3))
       RsTxn_IncomeExpenseForDailyPnL!EmployeeClaim = Val(MSHFlexGrid2.TextMatrix(i, 4))
       RsTxn_IncomeExpenseForDailyPnL!VendorClaim = Val(MSHFlexGrid2.TextMatrix(i, 5))
       RsTxn_IncomeExpenseForDailyPnL!FumigationClaim = Val(MSHFlexGrid2.TextMatrix(i, 6))
       RsTxn_IncomeExpenseForDailyPnL!FranchiseeFees = Val(MSHFlexGrid2.TextMatrix(i, 7))
       RsTxn_IncomeExpenseForDailyPnL!SalaryCost = Val(MSHFlexGrid2.TextMatrix(i, 8))
       RsTxn_IncomeExpenseForDailyPnL!SecurityCost = Val(MSHFlexGrid2.TextMatrix(i, 9))
       RsTxn_IncomeExpenseForDailyPnL!CreatedBy = Gen_UserLoginID
       RsTxn_IncomeExpenseForDailyPnL!CreatedDate = Now
       RsTxn_IncomeExpenseForDailyPnL.Update
    End If
Next

Call Grid_HeadIncomeExpense
MsgBox "Done"


End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame11.Enabled = False
   Exit Sub
End If

Frame11.Enabled = True

CmdProcess.Visible = False
CmdUpdateBillingRate.Visible = True
If LCase(Gen_UserRole) = "power" Then
   CmdProcess.Visible = True
   TxtFromDate.Enabled = True
    
End If

Call CmdClear_Click

End Sub

Private Sub Grid_HeadIncomeExpense()
CmdExcel.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid2
   .Clear
   .Rows = 2
   .FixedRows = 1
   .Cols = 13
   
   .TextMatrix(0, 0) = "GodownID"
   .TextMatrix(0, 1) = "AvgRate"
   .TextMatrix(0, 2) = "GodownCapacity"
   .TextMatrix(0, 3) = "CMinCMP"
   .TextMatrix(0, 4) = "EmployeeClaim"
   .TextMatrix(0, 5) = "VendorClaim"
   .TextMatrix(0, 6) = "FumigationClaim"
   .TextMatrix(0, 7) = "FranchiseeFees"
   .TextMatrix(0, 8) = "SalaryCost"
   .TextMatrix(0, 9) = "SecurityCost"
   .TextMatrix(0, 10) = "CoExpense"
   .TextMatrix(0, 11) = "NotionalRent"
   .TextMatrix(0, 12) = "LocationID"
   
   .ColWidth(0) = 1500
   .ColWidth(1) = 1800
   .ColWidth(2) = 1800
   .ColWidth(3) = 2800
   .ColWidth(4) = 2800
   .ColWidth(5) = 1800
   .ColWidth(6) = 1000
   .WordWrap = True
   .RowHeight(0) = 700
End With

End Sub

Private Function TableTxn_CreateTempGSLForDailyPnL()


Gen_mTimeStamp = GetTimeStamp

tmp = " Select GodownID,GSLID,SM_Prefix,Qty,TFlag,CDTFDate,CxTFDate,WtNoOfDays,WtStock,RateForDailyPnL,IncomewithoutReservation,ContractID into TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & " From ("
tmp = tmp & " Select GodownID,GSLID,SM_Prefix,SUM(BalanceWeight) 'Qty','O' TFlag,'' CDTFDate,'' CxTFDate,0 WtNoOfDays ,0 WtStock,Isnull(RateForDailyPnL,0) 'RateForDailyPnL', 0 IncomewithoutReservation,ContractID from Mst_GSL Where SpillageFlag = 0 And MadeupFlag = 0 And Flag  not in ('B','Z','E') And CDTFDate < '" & mCurrStartDate & "' "
tmp = tmp & " And GodownID in (Select Distinct MG.GodownID"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & TxtFromDate.Value & "') And MG.StartDate <= '" & TxtFromDate.Value & "')"
'tmp = tmp & " And " & wc
tmp = tmp & " Group By GodownID,GSLID,SM_Prefix,RateForDailyPnL,ContractID"
tmp = tmp & " Union All"
tmp = tmp & " Select c.GodownID,a.GSLID,a.SM_Prefix,SUM(a.DematWeight) 'Qty','D' TFlag,C.CDTFDate 'CDTFDate',b.CxTFDate 'CxTFDate',0 WtNoOfDays,0 WtStock,Isnull(a.RateForDailyPnL,0) 'RateForDailyPnL',0 IncomewithoutReservation,c.ContractID From  Txn_CxTF_GSL a"
tmp = tmp & " Inner Join Txn_CxTF_Details b on a.CxTFNo = b.CxTFNo and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
tmp = tmp & " Inner Join Mst_GSL c On a.GSLID = c.GSLID and a.SM_Prefix = c.SM_Prefix"
tmp = tmp & " Where c.SpillageFlag = 0 And c.MadeupFlag = 0 And b.TxnType = 'D' And b.CxTFDate >= '" & mCurrStartDate & "' And b.CxTFDate <= '" & TxtFromDate.Value & "'"
tmp = tmp & " And GodownID in (Select Distinct MG.GodownID"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & TxtFromDate.Value & "') And MG.StartDate <= '" & TxtFromDate.Value & "')"
'tmp = tmp & " And " & wc
tmp = tmp & " Group By c.GodownID,a.GSLID,a.SM_Prefix,C.CDTFDate,b.CxTFDate,a.RateForDailyPnL,c.ContractID"
tmp = tmp & " Union All"
tmp = tmp & " Select c.GodownID,a.GSLID,a.SM_Prefix,SUM(a.DematWeight) 'Qty','W' TFlag,C.CDTFDate 'CDTFDate',b.CxTFDate 'CxTFDate',0 WtNoOfDays,0 WtStock,Isnull(a.RateForDailyPnL,0) 'RateForDailyPnL',0 IncomewithoutReservation,c.ContractID From  Txn_CxTF_GSL a"
tmp = tmp & " Inner Join Txn_CxTF_Details b on a.CxTFNo = b.CxTFNo and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
tmp = tmp & " Inner Join Mst_GSL c On a.GSLID = c.GSLID and a.SM_Prefix = c.SM_Prefix"
tmp = tmp & " Where c.SpillageFlag = 0 And c.MadeupFlag = 0 And b.TxnType = 'W' And b.CxTFDate >= '" & mCurrStartDate & "' And b.CxTFDate <= '" & TxtFromDate.Value & "'"
tmp = tmp & " And GodownID in (Select Distinct MG.GodownID"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & TxtFromDate.Value & "') And MG.StartDate <= '" & TxtFromDate.Value & "')"
'tmp = tmp & " And " & wc
tmp = tmp & " Group By c.GodownID,a.GSLID,a.SM_Prefix,C.CDTFDate,b.CxTFDate,a.RateForDailyPnL,c.ContractID) A"
tmp = tmp & " Where Qty > 0"
tmp = tmp & " Order By GodownID,TFlag"

Call TmpTable

        'tmp = "Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = " & mTotalDays & " Where TFlag = 'O' And Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.WtNoOfDay is null And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
tmp = "Update TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".WtNoOfDays = " & mTotalDays & " Where TFlag = 'O' And WtNoOfDays = 0 " 'And TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
Call TmpTable

'--6
'tmp = " Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = (" & mTotalDays & " - Day(Txn_InvMonGSLDetail.CxTFDate)) + 1  Where YearMonth = '" & mFromDate & "' And TFlag = 'D' And WtNoOfDay is null And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
tmp = " Update TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".WtNoOfDays = (" & mTotalDays & " - Day(TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".CxTFDate)) + 1  Where TFlag = 'D' And WtNoOfDays = 0 " 'And TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
Call TmpTable

'--7
'tmp = " Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = (Day(Txn_InvMonGSLDetail.CxTFDate))  Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And WtNoOfDay is null And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
tmp = " Update TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".WtNoOfDays = (Day(TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".CxTFDate))  Where TFlag = 'W' And WtNoOfDays = 0 " 'And TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
Call TmpTable

'--8
'tmp = " Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = -(Day(Txn_InvMonGSLDetail.CxTFDate))  Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And DateDiff(Month,Txn_InvMonGSLDetail.CDTFDate,Txn_InvMonGSLDetail.CxTFDate) = 0 And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
tmp = " Update TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".WtNoOfDays = -(Day(TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".CxTFDate))  Where TFlag = 'W' And DateDiff(Month,TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".CDTFDate,TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".CxTFDate) = 0 " 'And TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
Call TmpTable

'--9
'tmp = " Update Txn_InvMonGSLDetail Set WtQty = WtNoOfDay*BalanceWeight Where YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
tmp = " Update TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & " Set WtStock = WtNoOfDays*Qty " 'Where TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
Call TmpTable

tmp = " Update TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & " Set IncomewithoutReservation = WtStock*RateForDailyPnL " 'Where TempGSLForDailyPnL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
Call TmpTable


End Function


Private Sub Grid_Head13()
With MSHFlexGrid13
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 12
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "Opening TI Amount"
   .TextMatrix(1, 3) = "Opening Received Amount"
   .TextMatrix(1, 4) = "Opeing Advance Amount"
   .TextMatrix(1, 5) = "Opening Balance"
   .TextMatrix(1, 6) = "This Month TI Amount"
   .TextMatrix(1, 7) = "This Month Advance"
   .TextMatrix(1, 8) = "This Month Received Amount"
   .TextMatrix(1, 9) = "Closing Balance"
   .TextMatrix(1, 10) = "Interest Amount"
   .TextMatrix(1, 11) = "LocationID"
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 2500
   .ColWidth(5) = 2500
   .ColWidth(6) = 2500
   .ColWidth(7) = 1500
   .ColWidth(8) = 1500
   .ColWidth(9) = 1500
   .ColWidth(10) = 1500
   .ColWidth(11) = 0
      
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub GetRentExpense()


''-- Add On 10 Feb 2016

tmp = " SELECT GodownID,SUM(RentAmount)[RentAmount] FROM Txn_GodownLessorAgreeDetail"
tmp = tmp & " Where Convert(Varchar,GodownID)+'~'+Convert(Varchar,LAID) in ("
tmp = tmp & " Select Convert(Varchar,a.GodownID)+'~'+Convert(Varchar,Max(a.LAID)) 'LAID' " '',Max(b.ExpiryDate) 'ExpDate'
tmp = tmp & " From Txn_GodownLessorAgreeDetail a"
tmp = tmp & " inner Join Txn_LessorAgreement b on a.LAID = b.LAID"
tmp = tmp & " inner join mst_godown g  on a.godownid = g.godownid"
tmp = tmp & " Where (g.closedate is null or g.Closedate > '" & Format(TxtFromDate.Value, "YYYY-MM-DD") & "') And b.ExpiryDate < '" & Format(TxtFromDate.Value, "YYYY-MM-DD") & "' And a.GodownID In " & mGodownIDs & ""
tmp = tmp & " And (a.GodownID not in (Select c.GodownID From Txn_GodownRentDetail c"
tmp = tmp & " Inner Join  Txn_GodownRent d on c.ClaimID = d.claimID"
tmp = tmp & " Where GOdownID In " & mGodownIDs & " And d.ExpenseFrom = '" & Format(TxtFromDate.Value, "YYYY-MM-DD") & "' And d.Flag in ('S','A', 'P')) And "
tmp = tmp & " a.GodownID not in (Select c.GodownID From Txn_GodownLessorAgreeDetail c"
tmp = tmp & " Inner Join  Txn_LessorAgreement d on c.LAID = d.LAID"
tmp = tmp & " Where  d.ExpiryDate >= '" & Format(TxtFromDate.Value, "YYYY-MM-DD") & "' And d.Flag in ('S','A', 'P')))"
tmp = tmp & " Group BY a.GodownID) Group By GodownID"


Call Tmp1Table


If TmpRs1.RecordCount > 0 Then
With MSHFlexGrid1
   For i = 2 To .Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(.TextMatrix(i, ETempDailyPnL.GodownID))
      If Not TmpRs1.EOF Then
         .TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!RentAmount), 0, TmpRs1!RentAmount)
      End If
'      .TextMatrix(0, ETempDailyPnL.LeaseRent) = Val(.TextMatrix(I, ETempDailyPnL.LeaseRent)) + Val(.TextMatrix(0, ETempDailyPnL.LeaseRent))
   Next
End With
End If
''----

tmp = " Select TLAD.GodownID, MG.CLoseDate, TLA.StartDate , TLA.ExpiryDate, sum(TLAD.RentAmount) Amount  from Txn_GodownLessorAgreeDetail TLAD"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TLAD.GodownID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID = TLAD.LAID"
tmp = tmp & " Where TLA.ExpiryDate >= '" & mEndDate & "' and TLA.StartDate <= '" & mCurrStartDate & "'  and  TLA.FLag not in ('R')"
'tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And MG.GOdownID in " & mGodownIDs & " "
tmp = tmp & " Group by TLAD.GodownID, TLA.EXpiryDate, TLA.StartDate , MG.CLoseDate "

Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownID))
      If Not TmpRs1.EOF Then
         If IsNull(TmpRs1!CloseDate) = True And Format(TmpRs1!StartDate, "yyyy-mm-dd") <= Format(mStartDate, "yyyy-mm-dd") Then
            MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
         Else
            If Format(TmpRs1!ExpiryDate, "yyyy-mm-dd") > Format(TmpRs1!CloseDate, "yyyy-mm-dd") Then
               If Format(TmpRs1!ExpiryDate, "yyyy-mm-dd") >= Format(mEndDate, "yyyy-mm-dd") Then
                  If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mCurrStartDate, "yyyy-mm-dd") Then
                     mdays = DateDiff("d", TmpRs1!StartDate, CDate(mEndDate)) + 1
                     MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########0"))
                  Else
                     MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                  End If
               Else
                  If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mCurrStartDate, "yyyy-mm-dd") Then
                     mdays = DateDiff("d", TmpRs1!StartDate, TmpRs1!ExpiryDate) + 1
                     MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########0"))
                  Else
                     mdays = DateDiff("d", CDate(mCurrStartDate), TmpRs1!ExpiryDate) + 1
                     MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########0"))
                  End If
               End If
            Else
               If Format(TmpRs1!CloseDate, "yyyy-mm-dd") >= Format(mEndDate, "yyyy-mm-dd") Then
                  If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mCurrStartDate, "yyyy-mm-dd") Then
                     mdays = DateDiff("d", TmpRs1!StartDate, CDate(mEndDate)) + 1
                     MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########0"))
                  Else
                     MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                  End If
               Else
                  If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mCurrStartDate, "yyyy-mm-dd") Then
                     mdays = DateDiff("d", TmpRs1!StartDate, TmpRs1!CloseDate) + 1
                     MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########0"))
                  Else
                     mdays = DateDiff("d", CDate(mStartDate), TmpRs1!CloseDate) + 1
                     MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LeaseRent) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########0"))
                  End If
               End If
            End If
         End If
      End If
'      MSHFlexGrid1.TextMatrix(0, ETempDailyPnL.LeaseRent) = Val(MSHFlexGrid1.TextMatrix(I, ETempDailyPnL.LeaseRent)) + Val(MSHFlexGrid1.TextMatrix(0, ETempDailyPnL.LeaseRent))
   Next
End If
End Sub

Private Sub GetAMRMName()

tmp = " Select ECM.LocationID, EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMName'"
tmp = tmp & " , EMP3.EmployeeName +'~'+EMP3.EmployeeNo as 'RMName'"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID And EMP.Flag = 'Y'"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP.DirectReportingTo = EMP2.EmployeeID And EMP2.Flag = 'Y'"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID And EMP3.Flag = 'Y'"
tmp = tmp & " Where EMP2.EmployeeName Is Not Null"
tmp = tmp & " Group By ECM.LocationID, EMP2.EmployeeName, EMP2.EmployeeNo"
tmp = tmp & " ,EMP3.EmployeeName ,EMP3.EmployeeNo"
Call Tmp1Table

For x_ = 2 To MSHFlexGrid1.Rows - 2
   TmpRs1.Filter = "LocationID = " & MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.LocationID)
   If TmpRs1.RecordCount > 0 Then
      For j = 1 To TmpRs1.RecordCount
         If MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.AMname) = "" Then
            MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.AMname) = IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname)
         Else
            If InStr(1, MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.AMname), IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname)) <= 0 Then
               MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.AMname) = IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname) & " , " & MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.AMname)
            End If
         End If
         If MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.RMName) = "" Then
            MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.RMName) = IIf(IsNull(TmpRs1!RMName), "", TmpRs1!RMName)
         Else
            If InStr(1, MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.RMName), IIf(IsNull(TmpRs1!RMName), "", TmpRs1!RMName)) <= 0 Then
               MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.RMName) = IIf(IsNull(TmpRs1!RMName), "", TmpRs1!RMName) & " , " & MSHFlexGrid1.TextMatrix(x_, ETempDailyPnL.RMName)
            End If
         End If
         TmpRs1.MoveNext
      Next j
   End If
   TmpRs1.Filter = ""
Next
End Sub

Private Sub GetOutstandingInterest()

Call Grid_Head13

tmp = " Select distinct MS.StateName,ML.LocationName, ML.LocationID ,Sum(GodownCapacity) 'GodownCapacity'"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If wc <> "" Then
   tmp = tmp & wc & " And MG.GodownID in " & mGodownIDs & " "
Else
   tmp = tmp & " Where MG.GodownID in " & mGodownIDs & ""
End If
tmp = tmp & " Group By MS.StateName,ML.LocationName, ML.LocationID "
tmp = tmp & " Order by MS.StateName,ML.LocationName, ML.LocationID "

Call TmpTable

tmp = "Select LocationName,Sum(TIAmount)TIAmount,Sum(RecAmount)+Sum(RecTDSAmount)+Sum(RecOtherDed) RecAmount"
tmp = tmp & " From (Select ML.LocationName,TI.TIID,TI.TIAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate<'" & mCurrStartDate & "'"
tmp = tmp & " )+ isnull(TI.AdjAmountFAG,0) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate<'" & mCurrStartDate & "'"
tmp = tmp & " ) +(Select isnull(Sum(TCRT.RecTDS),0) as RecTDS"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.Flag<>'R'And TIID=TI.TIID"
tmp = tmp & " And TCR.AdjDate<'" & mCurrStartDate & "'"
tmp = tmp & "  )RecTDSAmount,(Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate<'" & mCurrStartDate & "'"
tmp = tmp & " )+"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.Flag<>'R'And TIID=TI.TIID And TCR.AdjDate<'" & mCurrStartDate & "'"
tmp = tmp & " )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Where  TINo<>'' And TI.TIDate<'" & mCurrStartDate & "' )A"
tmp = tmp & " Group By A.LocationNAme"
tmp = tmp & " Order By A.LocationNAme"

Call Tmp1Table

tmp = "Select LocationName,Sum(TIAmount)TIAmount,Sum(RecAmount)+Sum(RecTDSAmount)+Sum(RecOtherDed) RecAmount"
tmp = tmp & " From ("
tmp = tmp & " Select ML.LocationName,TI.TIID,TI.TIAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate>='" & mCurrStartDate & "' And TCR.CRDate<='" & mEndDate & "'"
tmp = tmp & " )+ isnull(TI.AdjAmountFAG,0) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate>='" & mCurrStartDate & "' And TCR.CRDate<='" & mEndDate & "'"
tmp = tmp & " ) +"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.Flag<>'R'And TIID=TI.TIID"
tmp = tmp & " And TCR.AdjDate>='" & mCurrStartDate & "' And TCR.AdjDate<='" & mEndDate & "'"
tmp = tmp & " )RecTDSAmount,(Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate>='" & mCurrStartDate & "' And TCR.CRDate<='" & mEndDate & "'"
tmp = tmp & " )+"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.Flag<>'R'And TIID=TI.TIID And TCR.AdjDate>='" & mCurrStartDate & "' And TCR.AdjDate<='" & mEndDate & "'"
tmp = tmp & " )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Where  TINo<>'' And TI.TIDate<='" & mEndDate & "' )A"
tmp = tmp & " Group By A.LocationNAme"
tmp = tmp & " Order By A.LocationNAme"

Call Tmp6Table

tmp = " Select LocationName,Sum(TIAmount)TIAmount,Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed) RecAmount from"
tmp = tmp & " (select ML.LocationName,TI.TIID,TI.TIAmount, isnull(TI.AdjAmountFAG,0)+"
tmp = tmp & " (select isnull(Sum(TCRT.RecAmount),0)+isnull(Sum(TCRT.RecTDS),0)+isnull(Sum(TCRT.RecOtherDed),0) from Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID = TI.TIID and TCR.ReturnFlag<>'Y'"
tmp = tmp & " And TCR.RejectFlag <> 'Y' And TCR.CRDate between '" & mCurrStartDate & "' and '" & mEndDate & "') RecAmount,"
tmp = tmp & " (select isnull(Sum(TCRTA.RecTDS),0) from Txn_CRTDSAdjustmentDetail TCRTA"
tmp = tmp & " inner Join Txn_CRTDSAdjustment TCRA on TCRA.TDSAID= TCRTA.TDSAID and TCRA.SM_Prefix=TCRTA.SM_Prefix"
tmp = tmp & " Where TCRTA.TIID = TI.TIID and TCRA.Flag<>'R'And TCRTA.TIID=TI.TIID And TCRA.AdjDate between '" & mCurrStartDate & "' And '" & mEndDate & "') as RecTDS,"
tmp = tmp & " (select isnull(Sum(TCRTA.RecOtherDed),0) from Txn_CRTDSAdjustmentDetail TCRTA"
tmp = tmp & " inner Join Txn_CRTDSAdjustment TCRA on TCRA.TDSAID= TCRTA.TDSAID and TCRA.SM_Prefix=TCRTA.SM_Prefix"
tmp = tmp & " Where TCRTA.TIID = TI.TIID and TCRA.Flag<>'R'And TCRTA.TIID=TI.TIID And TCRA.AdjDate between '" & mCurrStartDate & "' And '" & mEndDate & "') as RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Where  TINo<>'' And TI.TIDate between '" & mCurrStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by ML.LocationName,TI.TIID,TI.TIAmount, TI.AdjAmountFAG"
tmp = tmp & " ) a"
tmp = tmp & " group by locationname"
Call Tmp3Table

tmp = "Select LocationName,Sum(AdvanceAmount) AdvanceAmount"
tmp = tmp & " From (Select ML.LocationName,TCRCD.Amount -"
tmp = tmp & " (select isnull(sum(isnull(TCTD.RecAmount,0))+sum(isnull(TCTD.RecTDS,0))+sum(isnull(TCTD.RecOtherDed,0)),0)"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCTD"
tmp = tmp & " Inner Join Txn_TIDetails TI on TI.TIID=TCTD.TIID"
tmp = tmp & " Where TCTD.CRID = CR.CRID And TCTD.SM_Prefix = CR.SM_Prefix And (TCTD.ClientIDRow = TCRCD.ClientIDRow Or TCTD.DPID = TCRCD.DPID) And TCTD.CRAAID > 0"
tmp = tmp & " And TI.TIDate<'" & mCurrStartDate & "'"
tmp = tmp & " ) as AdvanceAmount"
tmp = tmp & " From Txn_NewCashReceipt CR"
tmp = tmp & " Inner Join Mst_Location ML on CR.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Txn_NewCashReceiptClientDetail TCRCD on CR.CRID = TCRCD.CRID And CR.SM_Prefix = TCRCD.SM_Prefix"
tmp = tmp & " where CR.AdvanceFlag = 1 And TCRCD.Amount > 0 And"
tmp = tmp & " (TCRCD.Amount  -"
tmp = tmp & " (select sum(isnull(TCTD.RecAmount,0))+sum(isnull(TCTD.RecTDS,0))+sum(isnull(TCTD.RecOtherDed,0))"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCTD"
tmp = tmp & " Inner Join Txn_TIDetails TI on TI.TIID=TCTD.TIID"
tmp = tmp & " Where TCTD.CRID= CR.CRID And TCTD.SM_Prefix = CR.SM_Prefix And (TCTD.ClientIDrow = TCRCD.ClientIDRow Or TCTD.DPID = TCRCD.DPID) And TCTD.CRAAID>0 )) >0"
tmp = tmp & " And (CR.RejectFlag is Null or CR.RejectFlag = 'N') And (isnull(CR.ReturnFlag,'') <> 'Y')"
tmp = tmp & " And CR.CRDate<'" & mCurrStartDate & "'"
tmp = tmp & " )A"
tmp = tmp & " Group By LocationName"

Call Tmp4Table

tmp = "Select LocationName,Sum(AdvanceAmount) AdvanceAmount"
tmp = tmp & " From (Select ML.LocationName,TCRCD.Amount -"
tmp = tmp & " (select isnull(sum(isnull(TCTD.RecAmount,0))+sum(isnull(TCTD.RecTDS,0))+sum(isnull(TCTD.RecOtherDed,0)),0)"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCTD"
tmp = tmp & " Inner Join Txn_TIDetails TI on TI.TIID=TCTD.TIID"
tmp = tmp & " Where TCTD.CRID = CR.CRID And TCTD.SM_Prefix = CR.SM_Prefix And (TCTD.ClientIDRow = TCRCD.ClientIDRow Or TCTD.DPID = TCRCD.DPID) And TCTD.CRAAID > 0"
tmp = tmp & " And TI.TIDate<='" & mEndDate & "'"
tmp = tmp & " ) as AdvanceAmount"
tmp = tmp & " From Txn_NewCashReceipt CR"
tmp = tmp & " Inner Join Mst_Location ML on CR.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Txn_NewCashReceiptClientDetail TCRCD on CR.CRID = TCRCD.CRID And CR.SM_Prefix = TCRCD.SM_Prefix"
tmp = tmp & " where CR.AdvanceFlag = 1 And TCRCD.Amount > 0 And"
tmp = tmp & " (TCRCD.Amount  -"
tmp = tmp & " (select isnull(sum(TCTD.RecAmount),0)+isnull(sum(TCTD.RecTDS),0)+isnull(sum(TCTD.RecOtherDed),0)"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCTD"
tmp = tmp & " Inner Join Txn_TIDetails TI on TI.TIID=TCTD.TIID"
tmp = tmp & " Where TCTD.CRID= CR.CRID And TCTD.SM_Prefix = CR.SM_Prefix And (TCTD.ClientIDrow = TCRCD.ClientIDRow Or TCTD.DPID = TCRCD.DPID) And CRAAID>0 And TI.TIDate<='" & mEndDate & "')) >0"
tmp = tmp & " And (CR.RejectFlag is Null or CR.RejectFlag = 'N') And (isnull(CR.ReturnFlag,'') <> 'Y')"
tmp = tmp & " And CR.CRDate>='" & mCurrStartDate & "' And CR.CRDate<='" & mEndDate & "'"
tmp = tmp & " )A"
tmp = tmp & " Group By LocationName"

Call Tmp5Table

Call Grid_Head13

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid13.Rows = MSHFlexGrid13.Rows + TmpRs.RecordCount
   With MSHFlexGrid13
      For i = 2 To TmpRs.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
         TmpRs1.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs1.EOF = False Then
            .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!TIAmount), 0, TmpRs1!TIAmount)
            .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!RecAmount), 0, TmpRs1!RecAmount)
         End If
         If TmpRs4.RecordCount > 0 Then TmpRs4.MoveFirst
         TmpRs4.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs4.EOF = False Then
            .TextMatrix(i, 4) = IIf(IsNull(TmpRs4!AdvanceAmount), 0, TmpRs4!AdvanceAmount)
         End If
         .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 2)) - (Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4))), "#######0")
         If TmpRs6.RecordCount > 0 Then TmpRs6.MoveFirst
         TmpRs6.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs6.EOF = False Then
            .TextMatrix(i, 8) = IIf(IsNull(TmpRs6!RecAmount), 0, TmpRs6!RecAmount)
         End If
         If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
         TmpRs3.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs3.EOF = False Then
            .TextMatrix(i, 6) = IIf(IsNull(TmpRs3!TIAmount), 0, TmpRs3!TIAmount)
         End If
         If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
         TmpRs5.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs5.EOF = False Then
            .TextMatrix(i, 7) = IIf(IsNull(TmpRs5!AdvanceAmount), 0, TmpRs5!AdvanceAmount)
         End If
         
         .TextMatrix(i, 9) = Format(Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) - (Val(.TextMatrix(i, 7)) + Val(.TextMatrix(i, 8))), "#######0")
         .TextMatrix(i, 10) = Format(((((Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 9))) / 2) * 45 / 100) * 11 / 100) / 12, "###########0")
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
         .TextMatrix(0, 2) = Format(Val(.TextMatrix(i, 2)) + Val(.TextMatrix(0, 2)), "######0")
         .TextMatrix(0, 3) = Format(Val(.TextMatrix(i, 3)) + Val(.TextMatrix(0, 3)), "######0")
         .TextMatrix(0, 4) = Format(Val(.TextMatrix(i, 4)) + Val(.TextMatrix(0, 4)), "######0")
         .TextMatrix(0, 5) = Format(Val(.TextMatrix(i, 5)) + Val(.TextMatrix(0, 5)), "######0")
         .TextMatrix(0, 6) = Format(Val(.TextMatrix(i, 6)) + Val(.TextMatrix(0, 6)), "######0")
         .TextMatrix(0, 7) = Format(Val(.TextMatrix(i, 7)) + Val(.TextMatrix(0, 7)), "######0")
         .TextMatrix(0, 8) = Format(Val(.TextMatrix(i, 8)) + Val(.TextMatrix(0, 8)), "######0")
         .TextMatrix(0, 9) = Format(Val(.TextMatrix(i, 9)) + Val(.TextMatrix(0, 9)), "######0")
         .TextMatrix(0, 10) = Format(Val(.TextMatrix(i, 10)) + Val(.TextMatrix(0, 10)), "######0")
         TmpRs.MoveNext
      Next
   End With
End If

For i = 2 To MSHFlexGrid1.Rows - 2
   TmpRs.MoveFirst
   TmpRs.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LocationID))
   If Not TmpRs.EOF Then
      For j = 2 To MSHFlexGrid13.Rows - 2
         If MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.LocationID) = MSHFlexGrid13.TextMatrix(j, 11) Then
            MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.InterestCharges) = Format(Val(MSHFlexGrid13.TextMatrix(j, 10)) * Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.GodownCapacity)) / TmpRs!GodownCapacity, "########0")
            If Val(MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.InterestCharges)) < 0 Then
               MSHFlexGrid1.TextMatrix(i, ETempDailyPnL.InterestCharges) = 0
            End If
            Exit For
         End If
      Next j
   End If
Next i
End Sub

Private Sub Grid_HeadGodownStock()
With MSHFlexGrid3
   .Clear
   .Rows = 2
   .FixedRows = 1
   .Cols = 7
  
   .TextMatrix(0, 0) = "State Name"
   .TextMatrix(0, 1) = "Location Name"
   .TextMatrix(0, 2) = "CMP Name"
   .TextMatrix(0, 3) = "Godown No"
   .TextMatrix(0, 4) = "Address"
   .TextMatrix(0, 5) = "Godown ID"
   .TextMatrix(0, 6) = "Total Wt Stock"
   
   .ColWidth(0) = 1800
   .ColWidth(1) = 1800
   .ColWidth(2) = 1800
   .ColWidth(3) = 1500
   .ColWidth(4) = 2800
   .ColWidth(5) = 1800
   .ColWidth(6) = 1000
   
   .WordWrap = True
   .RowHeight(0) = 700
End With

End Sub

Private Sub Grid_HeadGodownIncome()

With MSHFlexGrid4
   .Clear
   .Rows = 2
   .FixedRows = 1
   .Cols = 15
   
   .TextMatrix(0, 0) = "State Name"
   .TextMatrix(0, 1) = "Location Name"
   .TextMatrix(0, 2) = "CMP Name"
   .TextMatrix(0, 3) = "Godown No"
   .TextMatrix(0, 4) = "Address"
   .TextMatrix(0, 5) = "Godown ID"
   .TextMatrix(0, 6) = "NNNR Income Adhoc"
   .TextMatrix(0, 7) = "NNNR Income"
   .TextMatrix(0, 8) = "Seven Days Income"
   .TextMatrix(0, 9) = "DPIncome AdHoc"
   .TextMatrix(0, 10) = "DPIncome Invoice"
   .TextMatrix(0, 11) = "Rejected Income"
   .TextMatrix(0, 12) = "Spot/RRN Income AdHoc"
   .TextMatrix(0, 13) = "Other Income"
   .TextMatrix(0, 14) = "SCM Income"
   
   
   .ColWidth(0) = 1800
   .ColWidth(1) = 1800
   .ColWidth(2) = 1800
   .ColWidth(3) = 1500
   .ColWidth(4) = 2800
   .ColWidth(5) = 1500
   .ColWidth(6) = 1600
   .ColWidth(7) = 1600
   .ColWidth(8) = 1600
   .ColWidth(9) = 1600
   .ColWidth(10) = 1600
   .ColWidth(11) = 1600
   .ColWidth(12) = 1600
   .ColWidth(13) = 1600
   .ColWidth(14) = 1600
   
   .WordWrap = True
   .RowHeight(0) = 700
End With

End Sub


Private Sub CmdUpdateBillingRate_Click()


mFromDate = Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01"
mSMonth = Month(TxtFromDate.Value)
mSYear = Year(TxtFromDate.Value)
mTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))


mContractID = "(0"

tmp = " Select Distinct a.ContractID from  Mst_StorageContract  a"
tmp = tmp & " left join Mst_SContractBillingDetail  d on a.ContractID = d.ContractID"
tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
tmp = tmp & " Where '" & mFromDate & "' Between d.FromDate and d.ToDate and TSGMD.ReservationQty = 0"
Call Tmp1Table

For i = 1 To TmpRs1.RecordCount
    mContractID = mContractID & "," & TmpRs1!ContractID
    TmpRs1.MoveNext
Next

mContractID = mContractID & ")"


If Check1.Value = vbUnchecked Then
   tmp = "Select Distinct SM_Prefix from Mst_GSL Where BillingCycleID is null And Flag not in ('E','Z','B') And SpillageFlag = '0' And MadeupFlag = '0' And (ContractID is null Or ContractID in " & mContractID & ")  " 'ANd SM_Prefix = 'ADONI'"
Else
   tmp = "Select Distinct a.SM_Prefix from Txn_CxTF_GSL a "
   tmp = tmp & " inner join Txn_CxTF_Details d on a.CxTFNo = d.CxTFNo and a.SM_Prefix = d.SM_Prefix and a.TxnType = d.TxnType "
   tmp = tmp & " inner join Mst_GSL g on a.GSLID = g.GSLID and a.SM_Prefix = g.SM_Prefix "
   tmp = tmp & " Where a.BillingCycleID is null and d.CxTFDate >= '" & CDate(mFromDate) & "' And g.SpillageFlag = 0 And g.MadeUpFlag = 0 And (g.ContractID is null Or g.ContractID in " & mContractID & ") " ' ANd g.SM_Prefix = 'ADMPR' "
End If
Call Tmp1Table


ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs1.RecordCount + 5


If TmpRs1.RecordCount > 0 Then
   For i = 1 To TmpRs1.RecordCount
   
       tmp = " Update Mst_GSL Set BagSize = (DepositeWeight/DepositeBags)*1000 where flag = 'O' and SpillageFlag = 0 And DepositeBags <> 0 And SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
       Call TmpTable

       tmp = " Update Mst_GSL Set BagSize = Round(Bagsize,0) Where SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
       Call TmpTable

       If Check1.Value = vbUnchecked Then
           tmp = " Update Mst_GSL Set Mst_GSL.BillingCycleID ="
           tmp = tmp & "  (Select a.BillingCycleID"
           tmp = tmp & "  from  Mst_GSL b"
           tmp = tmp & "  inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
           tmp = tmp & "  inner join mst_commodity c on b.commodityid = c.commodityid"
           tmp = tmp & "  left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
           tmp = tmp & "  And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
           tmp = tmp & "  where (b.ContractID = 0 Or b.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Mst_GSL.GSLID"
           tmp = tmp & "  And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
           tmp = tmp & "  Group by b.GSLID,b.SM_Prefix,"
           tmp = tmp & "  a.SPRCID, a.BillingCycleID,a.BillingUnit,"
           tmp = tmp & "  a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
           tmp = tmp & "  Where  Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingCycleID is null "
           Call TmpTable
            
            
            tmp = " Update Mst_GSL Set Mst_GSL.BillingUnit ="
            tmp = tmp & " (Select a.BillingUnit"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
            tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
            tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
            tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
            tmp = tmp & " Where  Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingCycleID is not null"
            Call TmpTable
            
           
            tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
            tmp = tmp & " (Select a.Amount"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
            tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
            tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And a.BillingUnit = 'Q'"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
            tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
            tmp = tmp & " Where  Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'Q' And Mst_GSL.BillingCycleID is not null "
            Call TmpTable
           
           
           
            tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
            tmp = tmp & " (Select ad.RatePerBag"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
            tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
            tmp = tmp & " left join Mst_SpecialRateCardDetail  ad on a.SPRCID = ad.SPRCID"
            tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And a.BillingUnit = 'B'"
            tmp = tmp & " And b.BagSize >= ad.FromBagSize And b.BagSize <= ad.ToBagSize"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
            tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount,ad.RatePerBag)"
            tmp = tmp & " Where  Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'B'"
            Call TmpTable
           
           
            tmp = " Update Mst_GSL Set RateCardType = 'S' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And ExchangeTypeID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Flag = 'O' And RateCardType is null"
            Call TmpTable
           
           
            tmp = " Update Mst_GSL Set Mst_GSL.BillingCycleID ="
            tmp = tmp & " (Select a.BillingCycleID"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
            tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
            tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SRCID is not null And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.DefaultRate = '1'"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
            tmp = tmp & " a.RateOnQty)"
            tmp = tmp & " Where Mst_GSL.RateCardType is null And  Mst_GSL.BillingCycleID is null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
            Call TmpTable
           
           
            tmp = " Update Mst_GSL Set Mst_GSL.BillingUnit ="
            tmp = tmp & " (Select a.flag"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
            tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
            tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SRCID is not null And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.DefaultRate = '1'"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
            tmp = tmp & " a.RateOnQty)"
            tmp = tmp & " Where Mst_GSL.RateCardType is null And  Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
            Call TmpTable
           
           
            tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
            tmp = tmp & " (Select a.RateOnQty"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
            tmp = tmp & " And  a.commoditygroupid = c.commoditygroupid"
            tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SRCID is not null  And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And a.Flag = 'Q' and a.DefaultRate = '1'"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
            tmp = tmp & " a.RateOnQty)"
            tmp = tmp & " Where  Mst_GSL.RateCardType is null And Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'Q'"
            Call TmpTable
           
           
            tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
            tmp = tmp & " (Select distinct ad.RatePerBag"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
            tmp = tmp & " left join Mst_StorageRateCardDetail  ad on a.SRCID = ad.SRCID"
            tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.DefaultRate = 1 And a.SRCID is not null  And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And a.Flag = 'B'"
            tmp = tmp & " And b.BagSize >= ad.FromBagSize And b.BagSize <= ad.ToBagSize"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
            tmp = tmp & " a.RateOnQty,ad.RatePerBag)"
            tmp = tmp & " Where  Mst_GSL.RateCardType is null And Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'B'"
            Call TmpTable
           
           
            tmp = " Update Mst_GSL Set BillingCycleID =  1 , BillingUnit = 'Q' Where ExchangeTypeID = 2 And Flag = 'O' And BillingCycleID is null And SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
            Call TmpTable
           
           
            tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
            tmp = tmp & " (Select ad.Rate"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_NCDEXStorageRateCard  a on a.locationid = l.locationid And a.CommodityGroupID =c.CommodityGroupID"
            tmp = tmp & " left Join Mst_NCDEXStorageRateCardDetail ad On a.NSRCID = ad.NSRCID And b.GodownID = ad.GodownID "
            tmp = tmp & " Where (b.ContractID = 0 Or b.ContractID Is Null) And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
            tmp = tmp & " Group by b.GodownID,ad.Rate)"
            tmp = tmp & " Where  Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 2 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'Q'"
            Call TmpTable
           
           
            tmp = " Update Mst_GSL Set Mst_GSL.BillingCycleID ="
            tmp = tmp & " (Select Distinct d.BillingCycleID"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And a.commodityid = c.commodityid and b.ContractID = a.ContractID "
            tmp = tmp & " left join Mst_SContractBillingDetail  d on a.ContractID = d.ContractID"
            tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
            tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
            tmp = tmp & " Where (b.ContractID Is Not Null And b.ContractID <> 0) And d.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between d.FromDate and d.ToDate and TSGMD.ReservationQty = 0"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.ContractID, d.BillingCycleID)"
            tmp = tmp & " Where (Mst_GSL.ContractID is not null and Mst_GSL.ContractID <>  0) And  Mst_GSL.BillingCycleID is null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
            Call TmpTable
            
            
            tmp = " Update Mst_GSL Set Mst_GSL.BillingUnit ="
            tmp = tmp & " (Select Distinct 'B'"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And a.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_SContractBillingDetail  d on a.ContractID = d.ContractID"
            tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
            tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
            tmp = tmp & " Where (b.ContractID Is Not Null And b.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between d.FromDate and d.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'No Of Bags'"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.ContractID, d.BillingCycleID,a.Flag,d.BillingRate)"
            tmp = tmp & " Where (Mst_GSL.ContractID is not null and Mst_GSL.ContractID <>  0) And  Mst_GSL.BillingCycleID is not null And  Mst_GSL.BillingUnit is null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
            Call TmpTable
            
            
            tmp = " Update Mst_GSL Set Mst_GSL.BillingUnit ="
            tmp = tmp & " (Select Distinct 'Q'"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And a.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_SContractBillingDetail  d on a.ContractID = d.ContractID"
            tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
            tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
            tmp = tmp & " Where (b.ContractID Is Not Null And b.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between d.FromDate and d.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'Weight In MT'"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.ContractID, d.BillingCycleID,a.Flag,d.BillingRate)"
            tmp = tmp & " Where (Mst_GSL.ContractID is not null and Mst_GSL.ContractID <>  0) And  Mst_GSL.BillingCycleID is not null And  Mst_GSL.BillingUnit is null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
            
            Call TmpTable
            
            
            tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
            tmp = tmp & " (Select Distinct d.BillingRate"
            tmp = tmp & " from  Mst_GSL b"
            tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
            tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
            tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And a.commodityid = c.commodityid and b.ContractID = a.ContractID "
            tmp = tmp & " left join Mst_SContractBillingDetail  d on a.ContractID = d.ContractID"
            tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
            tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
            tmp = tmp & " Where (b.ContractID Is Not Null And b.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Mst_GSL.GSLID"
            tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between d.FromDate and d.ToDate and TSGMD.ReservationQty = 0"
            tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
            tmp = tmp & " a.ContractID, d.BillingCycleID,a.Flag,d.BillingRate)"
            tmp = tmp & " Where (Mst_GSL.ContractID is not null and Mst_GSL.ContractID <>  0) And  Mst_GSL.BillingCycleID is not null  And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
            
            Call TmpTable
           
            tmp = " Update Mst_GSL Set RateCardType = 'N' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And ExchangeTypeID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Flag = 'O' And RateCardType is null"
            Call TmpTable
           
           
            tmp = "Update Mst_GSL Set RateForDailyPnL = (Rate) Where BillingCycleID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
            Call TmpTable
            
            tmp = "Update Mst_GSL Set RateForDailyPnL = Round((((Rate/BagSize)*1000)),0) Where BillingCycleID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
            Call TmpTable
     
            tmp = "Update Mst_GSL Set RateForDailyPnL = (((Rate)*4) / " & mTotalDays & ") Where BillingCycleID = 2 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
            Call TmpTable
            
            tmp = " Update Mst_GSL Set RateForDailyPnL = Round(((((Rate/BagSize)*1000)*4)/ " & mTotalDays & "),0) Where BillingCycleID = 2 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
            Call TmpTable
            
            tmp = " Update Mst_GSL Set RateForDailyPnL = (((Rate)*2)/ " & mTotalDays & ") Where BillingCycleID = 3 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
            Call TmpTable
            
            tmp = " Update Mst_GSL Set RateForDailyPnL = Round(((((Rate/BagSize)*1000)*2)/ " & mTotalDays & "),0) Where BillingCycleID = 3 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
            Call TmpTable
    
            tmp = "Update Mst_GSL Set RateForDailyPnL = ((Rate)/ " & mTotalDays & ") Where BillingCycleID = 4 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
            Call TmpTable
            
            tmp = "Update Mst_GSL Set RateForDailyPnL = ((Rate)/ " & mTotalDays & ") Where BillingCycleID = 5 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
            Call TmpTable
            
            tmp = " Update Mst_GSL Set RateForDailyPnL = Round((((Rate/BagSize)*1000)/ " & mTotalDays & "),0) Where BillingCycleID = 4 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
            Call TmpTable
    
            tmp = " Update Mst_GSL Set RateForDailyPnL = Round((((Rate/BagSize)*1000)/ " & mTotalDays & "),0) Where BillingCycleID = 5 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
            Call TmpTable

       End If
       
       
       
'------------------Transactions apply
       
       
       
       tmp = "  Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
       tmp = tmp & "  (Select a.DepoBillingCycleID"
       tmp = tmp & "  from  Txn_CxTF_GSL b"
       tmp = tmp & "  inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
       tmp = tmp & "  inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
       tmp = tmp & "  Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
       tmp = tmp & "  inner join mst_commodity c on d.commodityid = c.commodityid"
       tmp = tmp & "  left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
       tmp = tmp & "  And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
       tmp = tmp & "  Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Txn_CxTF_GSL.GSLID"
       tmp = tmp & "  And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
       tmp = tmp & "  Group by b.GSLID,b.SM_Prefix,"
       tmp = tmp & "  a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
       tmp = tmp & "  a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
       tmp = tmp & "  Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
       
       Call TmpTable
       
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
        tmp = tmp & " (Select a.BillingCycleID"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And  Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        
        Call TmpTable
 '       ---------------------------------------------------------------------------------------------------
 '       ---------------------------------------------------------------------------------------------------
        
 '       --Billing Unit
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
        tmp = tmp & " (Select a.BillingUnit"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and b.BillingCycleID is not null"
        tmp = tmp & "  Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
        tmp = tmp & " (Select a.BillingUnit"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' And b.BillingCycleID is not null"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
'        ---------------------------------------------------------------------------------------------------
'        ---------------------------------------------------------------------------------------------------
        
'        ---Rate
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select a.Amount"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select a.Amount"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
        
        
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select ad.RatePerBag"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " left join Mst_SpecialRateCardDepoDetail  ad on a.SPRCID = ad.SPRCID"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'B'"
        tmp = tmp & " And g.BagSize >= ad.FromBagSize And g.BagSize <= ad.ToBagSize"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount,ad.RatePerBag)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select ad.RatePerBag"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " left join Mst_SpecialRateCardDetail  ad on a.SPRCID = ad.SPRCID"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'B'"
        tmp = tmp & " And g.BagSize >= ad.FromBagSize And g.BagSize <= ad.ToBagSize"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount,ad.RatePerBag)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
        
        
        tmp = " Update Txn_CxTF_GSL Set RateCardType = 'S' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And TxnType = 'D' And RateCardType is null"
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set RateCardType = 'S' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And TxnType = 'W' And RateCardType is null"
        Call TmpTable
'        ---------------------------------------------------------------------------------------------------
'        ---------------------------------------------------------------------------------------------------
        
'        ---Base Rate Card
        
'        ---Billing Cycle ID
        
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
        tmp = tmp & " (Select a.DepoBillingCycleID"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and a.DefaultRate = '1'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.SRCID, a.DepoBillingCycleID,a.flag,a.RateOnQty)"
        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And  Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
        tmp = tmp & " (Select a.BillingCycleID"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and a.DefaultRate = '1'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.SRCID, a.BillingCycleID,a.flag,a.RateOnQty)"
        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And  Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
'        ---------------------------------------------------------------------------------------------------
'        ---------------------------------------------------------------------------------------------------
        
'        --Billing Unit
        
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
        tmp = tmp & " (Select a.flag"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and a.DefaultRate = '1'"
        tmp = tmp & " And  b.BillingCycleID is not null"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.SRCID, a.DepoBillingCycleID,a.flag,a.RateOnQty)"
        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And  Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
        tmp = tmp & " (Select a.flag"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and a.DefaultRate = '1'"
        tmp = tmp & " And  b.BillingCycleID is not null"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.SRCID, a.BillingCycleID,a.flag,a.RateOnQty)"
        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And  Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
'        ---------------------------------------------------------------------------------------------------
 '       ---------------------------------------------------------------------------------------------------
        
'        --Billing Rate
        
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select a.RateOnQty"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q' and a.DefaultRate = '1'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.SRCID, a.DepoBillingCycleID,a.Flag,a.RateOnQty)"
        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select a.RateOnQty"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q' and a.DefaultRate = '1'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.SRCID, a.DepoBillingCycleID,a.Flag,a.RateOnQty)"
        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
       
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select ad.RatePerBag"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " left join Mst_StorageRateCardDepoDetail  ad on a.SRCID = ad.SRCID"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.DefaultRate = 1 And a.SRCID is not null  And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'B'"
        tmp = tmp & " And g.BagSize >= ad.FromBagSize And g.BagSize <= ad.ToBagSize"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SRCID, a.DepoBillingCycleID,a.Flag,"
        tmp = tmp & " a.RateOnQty,ad.RatePerBag)"
        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select ad.RatePerBag"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " left join Mst_StorageRateCardDetail  ad on a.SRCID = ad.SRCID"
        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.DefaultRate = 1 And a.SRCID is not null  And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'B'"
        tmp = tmp & " And g.BagSize >= ad.FromBagSize And g.BagSize <= ad.ToBagSize"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
        tmp = tmp & " a.RateOnQty,ad.RatePerBag)"
        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        Call TmpTable
       
        
        tmp = " Update Txn_CxTF_GSL Set RateCardType = 'N' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And TxnType = 'D' And RateCardType is null"
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set RateCardType = 'N' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And TxnType = 'W' And RateCardType is null"
        Call TmpTable
        
'        ------------------------------------------------------------------------------------------------------------
'        ------------------------------------------------------------------------------------------------------------
        
'        --NCDEX Rate
        
'        --Billing Cycle and Billing Unit
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
        tmp = tmp & " (Select Distinct 1 'BillingCycleID'"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Where d.ExchangeTypeID = 2 And d.CxTFDate >= '" & CDate(mFromDate) & "')"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null and Txn_CxTF_GSL.TxnType in ('D','W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType in ('D','W') And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 2)"
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
        tmp = tmp & " (Select Distinct 'Q' 'BillingUnit'"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Where d.ExchangeTypeID = 2 And d.CxTFDate >= '" & CDate(mFromDate) & "')"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null and Txn_CxTF_GSL.TxnType in ('D','W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType in ('D','W') And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 2)"
        Call TmpTable
        
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select ad.Rate"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_NCDEXStorageRateCard  a on a.locationid = l.locationid And a.CommodityGroupID =c.CommodityGroupID"
        tmp = tmp & " left Join Mst_NCDEXStorageRateCardDetail ad On a.NSRCID = ad.NSRCID And g.GodownID = ad.GodownID "
        tmp = tmp & " Where b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 2 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,ad.Rate)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 2)"
        
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select ad.Rate"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_NCDEXStorageRateCard  a on a.locationid = l.locationid And a.CommodityGroupID =c.CommodityGroupID"
        tmp = tmp & " left Join Mst_NCDEXStorageRateCardDetail ad On a.NSRCID = ad.NSRCID And g.GodownID = ad.GodownID "
        tmp = tmp & " Where b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 2 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,ad.Rate)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 2)"
        
        Call TmpTable

        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
        tmp = tmp & " (Select Distinct MSB.BillingCycleID"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid And g.ContractID = a.ContractID and d.TxnType = 'D' "
        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID and g.GodownID = TSGMD.GodownID "
        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And MSB.BillingCycleID Is not Null And a.ContractID Is Not Null And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
        tmp = tmp & " And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
        tmp = tmp & " (Select Distinct MSB.BillingCycleID"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid And g.ContractID = a.ContractID and d.TxnType = 'W'"
        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID and g.GodownID = TSGMD.GodownID "
        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And MSB.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0 "
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
        tmp = tmp & " And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        
        Call TmpTable
        
      
        
'        -----------------------------------------------------------------------
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
        tmp = tmp & " (Select Distinct 'Q'"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is not Null And a.ContractID Is Not Null And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'Weight In MT'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,b.Rate)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
        tmp = tmp & " And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
        tmp = tmp & " (Select Distinct 'B'"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'No Of Bags'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,b.Rate)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
        tmp = tmp & " And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
        tmp = tmp & " (Select Distinct 'Q'"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid and d.ContractID = a.ContractID"
        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'Weight In MT'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,b.Rate)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
        tmp = tmp & " And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
        tmp = tmp & " (Select Distinct 'B'"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid and d.ContractID = a.ContractID "
        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'No Of Bags'"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,b.Rate)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
        tmp = tmp & " And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select Distinct MSB.BillingRate"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid And g.ContractID = a.ContractID and d.TxnType = 'D' and d.ContractID = a.ContractID "
        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.ContractID, MSB.BillingRate,a.Flag,b.Rate)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
        tmp = tmp & " And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
        tmp = tmp & " (Select Distinct MSB.BillingRate"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTFNo = d.CxTFNo and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
        tmp = tmp & " Inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid And g.ContractID = a.ContractID and d.TxnType = 'W' "
        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,a.ContractID, MSB.BillingRate,a.Flag,b.Rate)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
        tmp = tmp & " And Txn_CxTF_GSL.CxTFNo in (Select Distinct CxTFNo from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
        
        Call TmpTable

        

'-------------------------------------------------
        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = Rate Where RateForDailyPnL is null and BillingCycleID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = (((Rate)*4)/ " & mTotalDays & ") Where RateForDailyPnL is null and BillingCycleID = 2 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = (((Rate)*2)/ " & mTotalDays & ") Where RateForDailyPnL is null and BillingCycleID = 3 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = (Rate / " & mTotalDays & ") Where RateForDailyPnL is null and BillingCycleID = 4 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
        Call TmpTable

        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = (Rate / " & mTotalDays & ") Where RateForDailyPnL is null and BillingCycleID = 5 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
        Call TmpTable

        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
        tmp = tmp & " (Select Distinct Round((((b.Rate/g.BagSize)*1000)),0)"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " Where b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 1 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 1 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
        
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
        tmp = tmp & " (Select Distinct Round(((((b.Rate/g.BagSize)*1000)*4)/" & mTotalDays & "),0)"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " Where b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 2 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 2 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
        
        Call TmpTable
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
        tmp = tmp & " (Select Distinct Round(((((b.Rate/g.BagSize)*1000)*2)/" & mTotalDays & "),0)"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " Where b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 3 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 3 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
        
        Call TmpTable

        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
        tmp = tmp & " (Select Distinct Round(((((b.Rate/g.BagSize)*1000))/" & mTotalDays & "),0)"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " Where b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 4 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 4 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
        
        Call TmpTable
        
        
        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
        tmp = tmp & " (Select Distinct Round(((((b.Rate/g.BagSize)*1000))/" & mTotalDays & "),0)"
        tmp = tmp & " from  Txn_CxTF_GSL b"
        tmp = tmp & " inner join Mst_GSL g On b.GSLID = g.GSLID and b.SM_Prefix = g.SM_Prefix"
        tmp = tmp & " Where b.GSLID = Txn_CxTF_GSL.GSLID"
        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 5 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 5 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
                
        Call TmpTable

        ProgressBar1.Value = ProgressBar1.Value + 1

        TmpRs1.MoveNext
   Next
End If
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done " & Now & ""

End Sub
