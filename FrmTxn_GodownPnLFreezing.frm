VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_GodownPnLFreezing 
   Caption         =   "Godown wise Profitability Freezing"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8115
      Left            =   120
      TabIndex        =   19
      Top             =   1155
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14314
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consolidated upto selected Month"
      TabPicture(0)   =   "FrmTxn_GodownPnLFreezing.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "For Selected Month from Freezed Data"
      TabPicture(1)   =   "FrmTxn_GodownPnLFreezing.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid2"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7455
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13150
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   7455
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13150
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
      Height          =   1125
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Report for Selected Month"
         Enabled         =   0   'False
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
         Height          =   375
         Left            =   5805
         TabIndex        =   21
         Top             =   420
         Width           =   3165
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "From Freezed Data"
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
         Height          =   375
         Left            =   3240
         TabIndex        =   18
         Top             =   420
         Width           =   2175
      End
      Begin VB.ComboBox CmbYear 
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
         ItemData        =   "FrmTxn_GodownPnLFreezing.frx":0038
         Left            =   1575
         List            =   "FrmTxn_GodownPnLFreezing.frx":003A
         TabIndex        =   15
         Top             =   435
         Width           =   975
      End
      Begin VB.ComboBox CmbMonth 
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
         TabIndex        =   14
         Top             =   435
         Width           =   1410
      End
      Begin VB.CommandButton CmdFreeze 
         Caption         =   "Freeze P&&L"
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
         Left            =   13125
         TabIndex        =   13
         Top             =   600
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
         Left            =   4320
         TabIndex        =   7
         Top             =   240
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
         Left            =   2640
         TabIndex        =   6
         Top             =   195
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
         Left            =   10695
         TabIndex        =   5
         Top             =   600
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
         Left            =   9480
         TabIndex        =   4
         Top             =   600
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
         Left            =   11910
         TabIndex        =   3
         Top             =   600
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
         Left            =   7230
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   75
         Visible         =   0   'False
         Width           =   2640
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
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   75
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label Label2 
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
         Left            =   2400
         TabIndex        =   24
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data freezed upto : - "
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
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Year"
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
         Left            =   1807
         TabIndex        =   17
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   510
         TabIndex        =   16
         Top             =   120
         Width           =   630
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
         Left            =   10080
         TabIndex        =   11
         Top             =   195
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
         Left            =   12480
         TabIndex        =   10
         Top             =   195
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
         Left            =   8295
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
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
         Left            =   5490
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   9270
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmTxn_GodownPnLFreezing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mStateID As Variant 'mTotalDays, mTotalMonth
Dim mStartDate, mEndDate, mMonth, mYear, mTotalDays, wc As Variant
Enum ETemp1
   MonthYear = 0
   StateName = 1
   LocationName = 2
   AMname = 3
   RMName = 4
   CMPName = 5
   GodownNo = 6
   UnitType = 7
   GodownStartDate = 8
   GodownCloseDate = 9
   GodownCapacity = 10
   NcdexStock = 11
   NonNCDEXStock = 12
   NNNRIncomeAdhoc = 13
   NNNRIncome = 14
   ReservationIncomeAdhoc = 15
   ReservationIncome = 16
   SevenDaysIncome = 17
   DPIncomeAdHoc = 18
   FinalDPIncome = 19
   RejectedIncome = 20
   SpotIncomeAdHoc = 21
   OtherIncome = 22
   SCMIncome = 23
   CMUnbilledIncome = 24
   CMPendingIncome = 25
   CMApprovedIncome = 26
   LeftOverStock = 27
   PledgeIncome = 28
   PartiallyReversedIncome = 29
   LeaseRentPaid = 30
   LeaseRentFAG = 31
   LeaseRentDRC = 32
   EmpClaimFAG = 33
   EmpClaimAccrual = 34
   VendorExpenses = 35
   VendorExpensesAccrued = 36
   SecurityFeesPaid = 37
   SecurityFeesAccrual = 38
   FumigationFeesVendor = 39
   FumigationFeesPaid = 40
   FumigationFeesAccrual = 41
   FranchiseeFeesPaid = 42
   FranchiseeFeesFAG = 43
   FranchiseeFeesCMG = 44
   Salary = 45
   SalaryFieldStaff = 46
   NCDEXCharges = 47
   CoExpense = 48
   InsurancePremium = 49
   InterestCharges = 50
   TotalIncome = 51
   TotalExpense = 52
   GrossIncome = 53
   NetIncome = 54
   StateID = 55
   LocationID = 56
   GodownID = 57
   DPIncomeasperGSL = 58
   Difference = 59
   DPIncomeasperInvoice = 60
'   StateID = 49
'   LocationID = 50
'   GodownID = 51
'   DPIncomeasperGSL = 52
'   Difference = 53
'   DPIncomeasperInvoice = 54
   
End Enum


Enum ETemp2
   MonthYear = 0
   StateName = 1
   LocationName = 2
   AMname = 3
   RMName = 4
   CMPName = 5
   GodownNo = 6
   UnitType = 7
   GodownStartDate = 8
   GodownCloseDate = 9
   GodownCapacity = 10
   NcdexStock = 11
   NonNCDEXStock = 12
   NNNRIncomeAdhoc = 13
   NNNRIncome = 14
   ReservationIncomeAdhoc = 15
   ReservationIncome = 16
   SevenDaysIncome = 17
   DPIncomeAdHoc = 18
   FinalDPIncome = 19
   RejectedIncome = 20
   SpotIncomeAdHoc = 21
   OtherIncome = 22
   SCMIncome = 23
   CMUnbilledIncome = 24
   CMPendingIncome = 25
   CMApprovedIncome = 26
   LeftOverStock = 27
   PledgeIncome = 28
   PartiallyReversedIncome = 29
   LeaseRentPaid = 30
   LeaseRentFAG = 31
   LeaseRentDRC = 32
   EmpClaimFAG = 33
   EmpClaimAccrual = 34
   VendorExpenses = 35
   VendorExpensesAccrued = 36
   SecurityFeesPaid = 37
   SecurityFeesAccrual = 38
   FumigationFeesVendor = 39
   FumigationFeesPaid = 40
   FumigationFeesAccrual = 41
   FranchiseeFeesPaid = 42
   FranchiseeFeesFAG = 43
   FranchiseeFeesCMG = 44
   Salary = 45
   SalaryFieldStaff = 46
   NCDEXCharges = 47
   CoExpense = 48
   InsurancePremium = 49
   InterestCharges = 50
   TotalIncome = 51
   TotalExpense = 52
   GrossIncome = 53
   NetIncome = 54
   StateID = 55
   LocationID = 56
   GodownID = 57
   DPIncomeasperGSL = 58
   Difference = 59
   DPIncomeasperInvoice = 60
'   StateID = 49
'   LocationID = 50
'   GodownID = 51
'   DPIncomeasperGSL = 52
'   Difference = 53
'   DPIncomeasperInvoice = 54
   
End Enum

Private Sub Check1_Click()
If CmdFreeze.Visible = False Then Exit Sub
Check2.Value = 0
If Check1.Value = 0 Then
   CmdFreeze.Enabled = True
   Check2.Enabled = False
Else
   CmdFreeze.Enabled = False
   Check2.Enabled = True
End If

End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
Call TableMst_State
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
For I = 1 To RsMst_Location.RecordCount
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

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdClear_Click()

Call Grid_Head
Call Grid_Head2
CmbMonth.Text = ""
CmbYear.Text = ""
CmbStateID.Text = ""
CmbLocationID.Text = ""
If LCase(Gen_UserRole) = "power" Then
   CmdFreeze.Visible = True
Else
   CmdFreeze.Visible = False
End If
Check1.Value = 0
Check2.Value = 0
Check2.Enabled = False
'CmdFreeze.Enabled = False

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_GodownWisePnLFreezing.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_GodownWisePnLFreezing.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Godown Wise Profitability"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
mRow = 1

For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 7
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        If I = 0 Then
           oWS.Cells(mRow, C + 1).Interior.Color = &HD9D9FF
           oWS.Cells(mRow, C + 1).Font.Bold = True
        ElseIf I = 1 Then
           If C >= 13 And C <= 28 Then
              oWS.Cells(mRow, C + 1).Interior.Color = vbGreen
           ElseIf C >= 29 And C <= 46 Then
              oWS.Cells(mRow, C + 1).Interior.Color = vbRed
           ElseIf C >= 47 And C <= 50 Then
              oWS.Cells(mRow, C + 1).Interior.Color = &HC0FF&
           ElseIf C >= 51 And C <= 55 Then
              oWS.Cells(mRow, C + 1).Interior.Color = &HF3EEDB
           Else
              oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF
           End If
           oWS.Cells(mRow, C + 1).Font.Bold = True
           oWS.Cells(mRow, C + 1).RowHeight = 45
        End If
        oWS.Cells(mRow, C + 1).WrapText = True
        oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max

If Check2.Value = 1 Then

   Set oWS = oWB.Worksheets("Sheet2")
   oWS.Name = "Godown PnL"
   oWS.ClearArrows
   ProgressBar1.Value = 0
   ProgressBar1.Max = MSHFlexGrid1.Rows + 2
   mRow = 1

   For I = 0 To MSHFlexGrid2.Rows - 2
       mRow = mRow + 1 ' + 6
       For C = 0 To MSHFlexGrid2.Cols - 7
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
           If I = 0 Then
              oWS.Cells(mRow, C + 1).Interior.Color = &HD9D9FF
              oWS.Cells(mRow, C + 1).Font.Bold = True
           ElseIf I = 1 Then
              If C >= 13 And C <= 28 Then
                 oWS.Cells(mRow, C + 1).Interior.Color = vbGreen
              ElseIf C >= 29 And C <= 46 Then
                 oWS.Cells(mRow, C + 1).Interior.Color = vbRed
              ElseIf C >= 47 And C <= 50 Then
                 oWS.Cells(mRow, C + 1).Interior.Color = &HC0FF&
              ElseIf C >= 51 And C <= 55 Then
                 oWS.Cells(mRow, C + 1).Interior.Color = &HF3EEDB
              Else
                 oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF
              End If
              oWS.Cells(mRow, C + 1).Font.Bold = True
              oWS.Cells(mRow, C + 1).RowHeight = 45
           End If
           oWS.Cells(mRow, C + 1).WrapText = True
           oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
       Next
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next

   ProgressBar1.Value = ProgressBar1.Max
End If

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdFreeze_Click()

If MSHFlexGrid1.Rows <= 4 Then
   MsgBox "No data to freeze !!!"
   Exit Sub
End If


mCurrentStartDate = CmbYear.Text & "-" & mMonth & "-1"
mEndDate = CmbYear.Text & "-" & mMonth & "-" & mTotalDays
mPreviousEndDate = CDate(mCurrentStartDate) - 1

tmp = "Select * From Txn_MonGodownWiseProfitabilityFreezed Where MonthYear = '" & mPreviousEndDate & "' " 'And ReportFlag = '" & mReportFlag & "'"
Call Tmp1Table

If TmpRs1.RecordCount = 0 Then
   MsgBox "Profitability for previous month is not yet freezed . Please freeze previous month's profitability !!!"
   Exit Sub
End If

tmp = "Select * From Txn_MonGodownWiseProfitabilityFreezed Where MonthYear = '" & mEndDate & "' " 'And ReportFlag = '" & mReportFlag & "'"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MsgBox "Profitability is already freezed for this month!!!"
   Exit Sub
End If

If Opt_Capacity.Value = True Then
    mReportFlag = "C"
ElseIf Opt_Stock.Value = True Then
    mReportFlag = "S"
End If

'tmp = "Delete From Txn_MonGodownWiseProfitabilityFreezed Where MonthYear = '" & mEndDate & "' And ReportFlag = '" & mReportFlag & "'"
'Call TmpTable

tmp = "Select * From Txn_MonGodownWiseProfitabilityFreezed Where MonthYear = '" & mEndDate & "' And ReportFlag = '" & mReportFlag & "'"
Call Tmp1Table

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
For I = 2 To MSHFlexGrid1.Rows - 2
   TmpRs1.AddNew
   TmpRs1!MGWPID = GetMGWPID
   TmpRs1!G_UID = GetGUID
   TmpRs1!MonthYear = CDate(mEndDate)
   If MSHFlexGrid1.TextMatrix(I, ETemp1.GodownID) <> "" Then
      TmpRs1!StateID = MSHFlexGrid1.TextMatrix(I, ETemp1.StateID)
      TmpRs1!LocationID = MSHFlexGrid1.TextMatrix(I, ETemp1.LocationID)
      TmpRs1!GodownID = MSHFlexGrid1.TextMatrix(I, ETemp1.GodownID)
   End If
   TmpRs1!AM = MSHFlexGrid1.TextMatrix(I, ETemp1.AMname)
   TmpRs1!RM = MSHFlexGrid1.TextMatrix(I, ETemp1.RMName)
   TmpRs1!Capacity = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.GodownCapacity))
   TmpRs1!NcdexStock = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.NcdexStock))
   TmpRs1!NNStock = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.NonNCDEXStock))
   TmpRs1!NNNRIncomeAdhoc = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.NNNRIncomeAdhoc))
   TmpRs1!NNNRIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.NNNRIncome))
   TmpRs1!ReservationIncomeAdhoc = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.ReservationIncomeAdhoc))
   TmpRs1!ReservationIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.ReservationIncome))
   TmpRs1!SevenDaysIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.SevenDaysIncome))
   TmpRs1!DPIncomeAdHoc = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.DPIncomeAdHoc))
   TmpRs1!FinalDPIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.FinalDPIncome))
   TmpRs1!DPIncomeGSL = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.DPIncomeasperGSL))
   TmpRs1!Diff = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.Difference))
   TmpRs1!DPIncomeInvoice = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.DPIncomeasperInvoice))
   TmpRs1!RejectedIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.RejectedIncome))
   TmpRs1!SpotIncomeAdHoc = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.SpotIncomeAdHoc))
   TmpRs1!OtherIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.OtherIncome))
   TmpRs1!SCMIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.SCMIncome))
   TmpRs1!CMUnbilledIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.CMUnbilledIncome))
   TmpRs1!CMPendingIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.CMPendingIncome))
   TmpRs1!CMApprovedIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.CMApprovedIncome))
   TmpRs1!LeftOverStock = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.LeftOverStock))
   TmpRs1!PledgeIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.PledgeIncome))
   TmpRs1!PartiallyReversedIncome = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.PartiallyReversedIncome))
   TmpRs1!LeaseRentPaid = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.LeaseRentPaid))
   TmpRs1!LeaseRentFAG = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.LeaseRentFAG))
   TmpRs1!LeaseRentDRC = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.LeaseRentDRC))
   TmpRs1!EmpClaimFAG = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.EmpClaimFAG))
   TmpRs1!EmpClaimAccrual = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.EmpClaimAccrual))
   TmpRs1!VendorExpenses = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.VendorExpenses))
   TmpRs1!VendorExpensesAccrued = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.VendorExpensesAccrued))
   TmpRs1!SecurityFeesPaid = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.SecurityFeesPaid))
   TmpRs1!SecurityFeesAccrual = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.SecurityFeesAccrual))
   TmpRs1!FumigationFeesVendor = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.FumigationFeesVendor))
   TmpRs1!FumigationFeesPaid = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.FumigationFeesPaid))
   TmpRs1!FumigationFeesAccrual = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.FumigationFeesAccrual))
   TmpRs1!FranchiseeFeesPaid = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.FranchiseeFeesPaid))
   TmpRs1!FranchiseeFeesFAG = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.FranchiseeFeesFAG))
   TmpRs1!FranchiseeFeesCMG = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.FranchiseeFeesCMG))
   TmpRs1!Salary = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.Salary))
   TmpRs1!SalaryFieldStaff = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.SalaryFieldStaff))
   TmpRs1!NCDEXCharges = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.NCDEXCharges))
   TmpRs1!CoExpense = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.CoExpense))
   TmpRs1!InsurancePremium = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.InsurancePremium))
   TmpRs1!InterestCharges = Val(MSHFlexGrid1.TextMatrix(I, ETemp1.InterestCharges))
   TmpRs1!ReportFlag = mReportFlag
   TmpRs1!CreatedDate = Date
   TmpRs1!CreatedBy = Gen_UserLoginID
   TmpRs1.Update
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
Call Grid_Head
MsgBox "Done!!!"

tmp = "Select Max(MonthYear) MonthYear From Txn_MonGodownWiseProfitabilityFreezed "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   If IsNull(TmpRs1!MonthYear) = False Then
      Label2.Caption = Format(TmpRs1!MonthYear, "MMM-yyyy")
   End If
End If


End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head
Call Grid_Head2

wc = ""

If Trim(CmbMonth.Text) = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If Trim(CmbYear.Text) = "" Then
   MsgBox "Select Year!!!"
   CmbYear.SetFocus
   Exit Sub
End If

'mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
wc = ""
mYearMonth = CmbMonth.Text & "-" & CmbYear.Text
mYear = CmbYear.Text
If Check1.Value = 0 Then
   If CmbMonth = MonthName(1) Or CmbMonth = MonthName(2) Or CmbMonth = MonthName(3) Then
      mYear = mYear - 1
      mStartDate = mYear & "-" & 4 & "-1"
   ElseIf CmbMonth = MonthName(4) Or CmbMonth = MonthName(5) Or CmbMonth = MonthName(6) Or CmbMonth = MonthName(7) Or CmbMonth = MonthName(8) Or CmbMonth = MonthName(9) Or CmbMonth = MonthName(10) Or CmbMonth = MonthName(11) Or CmbMonth = MonthName(12) Then
      mStartDate = mYear & "-" & 4 & "-1"
   End If
Else
   mStartDate = mYear & "-" & mMonth & "-1"
End If

mEndDate = CmbYear.Text & "-" & mMonth & "-" & mTotalDays

'mTotalMonths = DateDiff("M", mWideStartDate, mWideEndDate) + 1
'mStartDate = Format(TxtFromDate.Value, "yyyy-mm-dd")
'mEndDate = Format(TxtToDate.Value, "yyyy-mm-dd")

tmp = " Select"
tmp = tmp & " MS.StateName, ML.LocationName, TMGP.AM, TMGP.RM, MCMP.CMPName, MG.GodownNo, MCMP.CMPName+'_'+MG.GodownNo, MUT.UnitType, MG.StartDate, MG.CloseDate,"
'tmp = tmp & " avg(MG.GodownCapacity) Capacity,  avg(TMGP.NcdexStock) NcdexStock, avg(TMGP.NNStock) NNStock, sum(TMGP.NNNRIncomeAdhoc) NNNRIncomeAdhoc, sum(TMGP.NNNRIncome) NNNRIncome,"
tmp = tmp & " Max(MG.GodownCapacity) Capacity,  sum(TMGP.NcdexStock) NcdexStock, sum(TMGP.NNStock) NNStock, sum(TMGP.NNNRIncomeAdhoc) NNNRIncomeAdhoc, sum(TMGP.NNNRIncome) NNNRIncome,"
tmp = tmp & " sum(TMGP.ReservationIncomeAdhoc) ReservationIncomeAdhoc, sum(TMGP.ReservationIncome) ReservationIncome, sum(TMGP.SevenDaysIncome) SevenDaysIncome, sum(TMGP.DPIncomeAdHoc) DPIncomeAdHoc,"
tmp = tmp & " sum(TMGP.FinalDPIncome) FinalDPIncome,  sum(TMGP.RejectedIncome) RejectedIncome,"
tmp = tmp & " sum(TMGP.SpotIncomeAdHoc) SpotIncomeAdHoc, sum(TMGP.OtherIncome) OtherIncome, sum(TMGP.SCMIncome) SCMIncome, sum(TMGP.CMUnbilledIncome) CMUnbilledIncome,"
tmp = tmp & " sum(TMGP.CMPendingIncome) CMPendingIncome, sum(TMGP.CMApprovedIncome) CMApprovedIncome, sum(TMGP.LeftOverStock) LeftOverStock, sum(TMGP.PledgeIncome) PledgeIncome, "
tmp = tmp & " sum(TMGP.LeaseRentPaid) LeaseRentPaid, sum(TMGP.LeaseRentFAG) LeaseRentFAG, sum(TMGP.LeaseRentDRC) LeaseRentDRC, sum(TMGP.EmpClaimFAG) EmpClaimFAG, sum(TMGP.EmpClaimAccrual) EmpClaimAccrual,"
tmp = tmp & " sum(TMGP.VendorExpenses) VendorExpenses, sum(TMGP.VendorExpensesAccrued) VendorExpensesAccrued, sum(TMGP.SecurityFeesPaid) SecurityFeesPaid, sum(TMGP.SecurityFeesAccrual) SecurityFeesAccrual, sum(TMGP.FumigationFeesVendor) FumigationFeesVendor,"
tmp = tmp & " sum(TMGP.FumigationFeesPaid) FumigationFeesPaid, sum(TMGP.FumigationFeesAccrual) FumigationFeesAccrual, sum(TMGP.FranchiseeFeesPaid) FranchiseeFeesPaid, sum(TMGP.FranchiseeFeesFAG) FranchiseeFeesFAG, sum(TMGP.FranchiseeFeesCMG) FranchiseeFeesCMG, (Sum(TMGP.Salary)) Salary , (Sum(TMGP.SalaryFieldStaff)) SalaryFieldStaff , "
tmp = tmp & " sum(TMGP.NCDEXCharges) NCDEXCharges, sum(TMGP.CoExpense) CoExpense, sum(TMGP.InsurancePremium) InsurancePremium, sum(TMGP.InterestCharges) InterestCharges,sum(TMGP.PartiallyReversedIncome) PartiallyReversedIncome,"
tmp = tmp & " TMGP.StateID , TMGP.LocationID, TMGP.GodownID, Sum(TMGP.DPIncomeGSL) DPIncomeGSL, Sum(TMGP.Diff) Diff, Sum(TMGP.DPIncomeInvoice) DPIncomeInvoice"
If Check1.Value = 1 Then
   tmp = tmp & " From Txn_MonGodownWiseProfitabilityFreezed TMGP"
Else
   tmp = tmp & " From Txn_MonGodownWiseProfitability TMGP"
End If
tmp = tmp & " left Join Mst_Location ML On ML.LocationID = TMGP.LocationID"
tmp = tmp & " left Join Mst_State MS On MS.StateID =  TMGP.StateID"
tmp = tmp & " left Join Mst_Godown MG On MG.GodownID = TMGP.GodownID"
tmp = tmp & " left Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
tmp = tmp & " left Join Mst_CMP MCMP On MCMP.CMPID = MG.CMPID"
If wc <> "" Then
   tmp = tmp & " " & wc & " And TMGP.MonthYear between '" & mStartDate & "' and '" & mEndDate & "'"
Else
   tmp = tmp & " Where TMGP.MonthYear between '" & mStartDate & "' and '" & mEndDate & "'"
End If
If Opt_Capacity.Value = True Then
    tmp = tmp & " And TMGP.ReportFlag = 'C'"
ElseIf Opt_Stock.Value = True Then
    tmp = tmp & " And TMGP.ReportFlag = 'S'"
End If
tmp = tmp & " group by MS.StateName,"
tmp = tmp & " ML.LocationName, TMGP.AM, TMGP.RM, MCMP.CMPName, MG.GodownNo, MUT.UnitType, MG.StartDate, MG.CloseDate,"
tmp = tmp & " TMGP.StateID , TMGP.LocationID, TMGP.GodownID"
tmp = tmp & " Order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   Label5.Caption = TmpRs.RecordCount
   ProgressBar1.Value = 0
   ProgressBar1.Max = MSHFlexGrid1.Rows + TmpRs.RecordCount
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   CmdExcel.Enabled = True
    
   For I = 2 To TmpRs.RecordCount + 1
       With MSHFlexGrid1
           .TextMatrix(I, ETemp1.MonthYear) = mYearMonth
           .TextMatrix(I, ETemp1.StateName) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
           .TextMatrix(I, ETemp1.LocationName) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
           .TextMatrix(I, ETemp1.AMname) = IIf(IsNull(TmpRs!AM), "", TmpRs!AM)
           .TextMatrix(I, ETemp1.RMName) = IIf(IsNull(TmpRs!RM), "", TmpRs!RM)
           .TextMatrix(I, ETemp1.CMPName) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
           .TextMatrix(I, ETemp1.GodownNo) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
           .TextMatrix(I, ETemp1.GodownCapacity) = IIf(IsNull(TmpRs!Capacity), "", TmpRs!Capacity)
           .TextMatrix(I, ETemp1.NcdexStock) = IIf(IsNull(TmpRs!NcdexStock), "", TmpRs!NcdexStock)
           .TextMatrix(I, ETemp1.NonNCDEXStock) = IIf(IsNull(TmpRs!NNStock), "", TmpRs!NNStock)
           .TextMatrix(I, ETemp1.NNNRIncomeAdhoc) = IIf(IsNull(TmpRs!NNNRIncomeAdhoc), "", TmpRs!NNNRIncomeAdhoc)
           .TextMatrix(I, ETemp1.NNNRIncome) = IIf(IsNull(TmpRs!NNNRIncome), "", TmpRs!NNNRIncome)
           .TextMatrix(I, ETemp1.ReservationIncomeAdhoc) = IIf(IsNull(TmpRs!ReservationIncomeAdhoc), "", TmpRs!ReservationIncomeAdhoc)
           .TextMatrix(I, ETemp1.ReservationIncome) = IIf(IsNull(TmpRs!ReservationIncome), "", TmpRs!ReservationIncome)
           .TextMatrix(I, ETemp1.SevenDaysIncome) = IIf(IsNull(TmpRs!SevenDaysIncome), "", TmpRs!SevenDaysIncome)
           .TextMatrix(I, ETemp1.DPIncomeAdHoc) = IIf(IsNull(TmpRs!DPIncomeAdHoc), "", TmpRs!DPIncomeAdHoc)
           .TextMatrix(I, ETemp1.FinalDPIncome) = IIf(IsNull(TmpRs!FinalDPIncome), "", TmpRs!FinalDPIncome)
           .TextMatrix(I, ETemp1.DPIncomeasperGSL) = IIf(IsNull(TmpRs!DPIncomeGSL), "", TmpRs!DPIncomeGSL)
           .TextMatrix(I, ETemp1.Difference) = IIf(IsNull(TmpRs!Diff), "", TmpRs!Diff)
           .TextMatrix(I, ETemp1.DPIncomeasperInvoice) = IIf(IsNull(TmpRs!DPIncomeInvoice), "", TmpRs!DPIncomeInvoice)
           .TextMatrix(I, ETemp1.RejectedIncome) = IIf(IsNull(TmpRs!RejectedIncome), "", TmpRs!RejectedIncome)
           .TextMatrix(I, ETemp1.SpotIncomeAdHoc) = IIf(IsNull(TmpRs!SpotIncomeAdHoc), "", TmpRs!SpotIncomeAdHoc)
           .TextMatrix(I, ETemp1.OtherIncome) = IIf(IsNull(TmpRs!OtherIncome), "", TmpRs!OtherIncome)
           .TextMatrix(I, ETemp1.SCMIncome) = IIf(IsNull(TmpRs!SCMIncome), "", TmpRs!SCMIncome)
           .TextMatrix(I, ETemp1.CMUnbilledIncome) = IIf(IsNull(TmpRs!CMUnbilledIncome), "", TmpRs!CMUnbilledIncome)
           .TextMatrix(I, ETemp1.CMPendingIncome) = IIf(IsNull(TmpRs!CMPendingIncome), "", TmpRs!CMPendingIncome)
           .TextMatrix(I, ETemp1.CMApprovedIncome) = IIf(IsNull(TmpRs!CMApprovedIncome), "", TmpRs!CMApprovedIncome)
           .TextMatrix(I, ETemp1.LeftOverStock) = IIf(IsNull(TmpRs!LeftOverStock), "", TmpRs!LeftOverStock)
           .TextMatrix(I, ETemp1.PledgeIncome) = IIf(IsNull(TmpRs!PledgeIncome), "", TmpRs!PledgeIncome)
           .TextMatrix(I, ETemp1.PartiallyReversedIncome) = IIf(IsNull(TmpRs!PartiallyReversedIncome), "", TmpRs!PartiallyReversedIncome)
           .TextMatrix(I, ETemp1.LeaseRentPaid) = IIf(IsNull(TmpRs!LeaseRentPaid), "", TmpRs!LeaseRentPaid)
           .TextMatrix(I, ETemp1.LeaseRentFAG) = IIf(IsNull(TmpRs!LeaseRentFAG), "", TmpRs!LeaseRentFAG)
           .TextMatrix(I, ETemp1.LeaseRentDRC) = IIf(IsNull(TmpRs!LeaseRentDRC), "", TmpRs!LeaseRentDRC)
           .TextMatrix(I, ETemp1.EmpClaimFAG) = IIf(IsNull(TmpRs!EmpClaimFAG), "", TmpRs!EmpClaimFAG)
           .TextMatrix(I, ETemp1.EmpClaimAccrual) = IIf(IsNull(TmpRs!EmpClaimAccrual), "", TmpRs!EmpClaimAccrual)
           .TextMatrix(I, ETemp1.VendorExpenses) = IIf(IsNull(TmpRs!VendorExpenses), "", TmpRs!VendorExpenses)
           .TextMatrix(I, ETemp1.VendorExpensesAccrued) = IIf(IsNull(TmpRs!VendorExpensesAccrued), "", TmpRs!VendorExpensesAccrued)
           .TextMatrix(I, ETemp1.SecurityFeesPaid) = IIf(IsNull(TmpRs!SecurityFeesPaid), "", TmpRs!SecurityFeesPaid)
           .TextMatrix(I, ETemp1.SecurityFeesAccrual) = IIf(IsNull(TmpRs!SecurityFeesAccrual), "", TmpRs!SecurityFeesAccrual)
           .TextMatrix(I, ETemp1.FumigationFeesVendor) = IIf(IsNull(TmpRs!FumigationFeesVendor), "", TmpRs!FumigationFeesVendor)
           .TextMatrix(I, ETemp1.FumigationFeesPaid) = IIf(IsNull(TmpRs!FumigationFeesPaid), "", TmpRs!FumigationFeesPaid)
           .TextMatrix(I, ETemp1.FumigationFeesAccrual) = IIf(IsNull(TmpRs!FumigationFeesAccrual), "", TmpRs!FumigationFeesAccrual)
           .TextMatrix(I, ETemp1.FranchiseeFeesPaid) = IIf(IsNull(TmpRs!FranchiseeFeesPaid), "", TmpRs!FranchiseeFeesPaid)
           .TextMatrix(I, ETemp1.FranchiseeFeesCMG) = IIf(IsNull(TmpRs!FranchiseeFeesCMG), "", TmpRs!FranchiseeFeesCMG)
           .TextMatrix(I, ETemp1.FranchiseeFeesFAG) = IIf(IsNull(TmpRs!FranchiseeFeesFAG), "", TmpRs!FranchiseeFeesFAG)
           .TextMatrix(I, ETemp1.Salary) = IIf(IsNull(TmpRs!Salary), "", TmpRs!Salary)
           .TextMatrix(I, ETemp1.SalaryFieldStaff) = IIf(IsNull(TmpRs!SalaryFieldStaff), "", TmpRs!SalaryFieldStaff)
           .TextMatrix(I, ETemp1.NCDEXCharges) = IIf(IsNull(TmpRs!NCDEXCharges), "", TmpRs!NCDEXCharges)
           .TextMatrix(I, ETemp1.CoExpense) = IIf(IsNull(TmpRs!CoExpense), "", TmpRs!CoExpense)
           .TextMatrix(I, ETemp1.InsurancePremium) = IIf(IsNull(TmpRs!InsurancePremium), "", TmpRs!InsurancePremium)
           .TextMatrix(I, ETemp1.InterestCharges) = IIf(IsNull(TmpRs!InterestCharges), "", TmpRs!InterestCharges)
           .TextMatrix(I, ETemp1.GodownStartDate) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
           .TextMatrix(I, ETemp1.GodownCloseDate) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
           .TextMatrix(I, ETemp1.UnitType) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
           .TextMatrix(I, ETemp1.StateID) = IIf(IsNull(TmpRs!StateID), "", TmpRs!StateID)
           .TextMatrix(I, ETemp1.LocationID) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
           .TextMatrix(I, ETemp1.GodownID) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
           
           .TextMatrix(0, ETemp1.NcdexStock) = Val(.TextMatrix(I, ETemp1.NcdexStock)) + Val(.TextMatrix(0, ETemp1.NcdexStock))
           .TextMatrix(0, ETemp1.GodownCapacity) = Val(.TextMatrix(I, ETemp1.GodownCapacity)) + Val(.TextMatrix(0, ETemp1.GodownCapacity))
           .TextMatrix(0, ETemp1.NonNCDEXStock) = Val(.TextMatrix(I, ETemp1.NonNCDEXStock)) + Val(.TextMatrix(0, ETemp1.NonNCDEXStock))
           .TextMatrix(0, ETemp1.NNNRIncomeAdhoc) = Val(.TextMatrix(I, ETemp1.NNNRIncomeAdhoc)) + Val(.TextMatrix(0, ETemp1.NNNRIncomeAdhoc))
           .TextMatrix(0, ETemp1.NNNRIncome) = Val(.TextMatrix(I, ETemp1.NNNRIncome)) + Val(.TextMatrix(0, ETemp1.NNNRIncome))
           .TextMatrix(0, ETemp1.ReservationIncomeAdhoc) = Val(.TextMatrix(I, ETemp1.ReservationIncomeAdhoc)) + Val(.TextMatrix(0, ETemp1.ReservationIncomeAdhoc))
           .TextMatrix(0, ETemp1.ReservationIncome) = Val(.TextMatrix(I, ETemp1.ReservationIncome)) + Val(.TextMatrix(0, ETemp1.ReservationIncome))
           .TextMatrix(0, ETemp1.SevenDaysIncome) = Val(.TextMatrix(I, ETemp1.SevenDaysIncome)) + Val(.TextMatrix(0, ETemp1.SevenDaysIncome))
           .TextMatrix(0, ETemp1.DPIncomeAdHoc) = Val(.TextMatrix(I, ETemp1.DPIncomeAdHoc)) + Val(.TextMatrix(0, ETemp1.DPIncomeAdHoc))
           .TextMatrix(0, ETemp1.FinalDPIncome) = Val(.TextMatrix(I, ETemp1.FinalDPIncome)) + Val(.TextMatrix(0, ETemp1.FinalDPIncome))
           .TextMatrix(0, ETemp1.DPIncomeasperGSL) = Val(.TextMatrix(I, ETemp1.DPIncomeasperGSL)) + Val(.TextMatrix(0, ETemp1.DPIncomeasperGSL))
           .TextMatrix(0, ETemp1.Difference) = Val(.TextMatrix(I, ETemp1.Difference)) + Val(.TextMatrix(0, ETemp1.Difference))
           .TextMatrix(0, ETemp1.DPIncomeasperInvoice) = Val(.TextMatrix(I, ETemp1.DPIncomeasperInvoice)) + Val(.TextMatrix(0, ETemp1.DPIncomeasperInvoice))
           .TextMatrix(0, ETemp1.RejectedIncome) = Val(.TextMatrix(I, ETemp1.RejectedIncome)) + Val(.TextMatrix(0, ETemp1.RejectedIncome))
           .TextMatrix(0, ETemp1.SpotIncomeAdHoc) = Val(.TextMatrix(I, ETemp1.SpotIncomeAdHoc)) + Val(.TextMatrix(0, ETemp1.SpotIncomeAdHoc))
           .TextMatrix(0, ETemp1.OtherIncome) = Val(.TextMatrix(I, ETemp1.OtherIncome)) + Val(.TextMatrix(0, ETemp1.OtherIncome))
           .TextMatrix(0, ETemp1.SCMIncome) = Val(.TextMatrix(I, ETemp1.SCMIncome)) + Val(.TextMatrix(0, ETemp1.SCMIncome))
           .TextMatrix(0, ETemp1.CMUnbilledIncome) = Val(.TextMatrix(I, ETemp1.CMUnbilledIncome)) + Val(.TextMatrix(0, ETemp1.CMUnbilledIncome))
           .TextMatrix(0, ETemp1.CMPendingIncome) = Val(.TextMatrix(I, ETemp1.CMPendingIncome)) + Val(.TextMatrix(0, ETemp1.CMPendingIncome))
           .TextMatrix(0, ETemp1.CMApprovedIncome) = Val(.TextMatrix(I, ETemp1.CMApprovedIncome)) + Val(.TextMatrix(0, ETemp1.CMApprovedIncome))
           .TextMatrix(0, ETemp1.LeftOverStock) = Val(.TextMatrix(I, ETemp1.LeftOverStock)) + Val(.TextMatrix(0, ETemp1.LeftOverStock))
           .TextMatrix(0, ETemp1.PledgeIncome) = Val(.TextMatrix(I, ETemp1.PledgeIncome)) + Val(.TextMatrix(0, ETemp1.PledgeIncome))
           .TextMatrix(0, ETemp1.PartiallyReversedIncome) = Val(.TextMatrix(I, ETemp1.PartiallyReversedIncome)) + Val(.TextMatrix(0, ETemp1.PartiallyReversedIncome))
           .TextMatrix(0, ETemp1.LeaseRentPaid) = Val(.TextMatrix(I, ETemp1.LeaseRentPaid)) + Val(.TextMatrix(0, ETemp1.LeaseRentPaid))
           .TextMatrix(0, ETemp1.LeaseRentFAG) = Val(.TextMatrix(I, ETemp1.LeaseRentFAG)) + Val(.TextMatrix(0, ETemp1.LeaseRentFAG))
           .TextMatrix(0, ETemp1.LeaseRentDRC) = Val(.TextMatrix(I, ETemp1.LeaseRentDRC)) + Val(.TextMatrix(0, ETemp1.LeaseRentDRC))
           .TextMatrix(0, ETemp1.EmpClaimFAG) = Val(.TextMatrix(I, ETemp1.EmpClaimFAG)) + Val(.TextMatrix(0, ETemp1.EmpClaimFAG))
           .TextMatrix(0, ETemp1.EmpClaimAccrual) = Val(.TextMatrix(I, ETemp1.EmpClaimAccrual)) + Val(.TextMatrix(0, ETemp1.EmpClaimAccrual))
           .TextMatrix(0, ETemp1.VendorExpenses) = Val(.TextMatrix(I, ETemp1.VendorExpenses)) + Val(.TextMatrix(0, ETemp1.VendorExpenses))
           .TextMatrix(0, ETemp1.VendorExpensesAccrued) = Val(.TextMatrix(I, ETemp1.VendorExpensesAccrued)) + Val(.TextMatrix(0, ETemp1.VendorExpensesAccrued))
           .TextMatrix(0, ETemp1.SecurityFeesPaid) = Val(.TextMatrix(I, ETemp1.SecurityFeesPaid)) + Val(.TextMatrix(0, ETemp1.SecurityFeesPaid))
           .TextMatrix(0, ETemp1.SecurityFeesAccrual) = Val(.TextMatrix(I, ETemp1.SecurityFeesAccrual)) + Val(.TextMatrix(0, ETemp1.SecurityFeesAccrual))
           .TextMatrix(0, ETemp1.FumigationFeesVendor) = Val(.TextMatrix(I, ETemp1.FumigationFeesVendor)) + Val(.TextMatrix(0, ETemp1.FumigationFeesVendor))
           .TextMatrix(0, ETemp1.FumigationFeesPaid) = Val(.TextMatrix(I, ETemp1.FumigationFeesPaid)) + Val(.TextMatrix(0, ETemp1.FumigationFeesPaid))
           .TextMatrix(0, ETemp1.FumigationFeesAccrual) = Val(.TextMatrix(I, ETemp1.FumigationFeesAccrual)) + Val(.TextMatrix(0, ETemp1.FumigationFeesAccrual))
           .TextMatrix(0, ETemp1.FranchiseeFeesPaid) = Val(.TextMatrix(I, ETemp1.FranchiseeFeesPaid)) + Val(.TextMatrix(0, ETemp1.FranchiseeFeesPaid))
           .TextMatrix(0, ETemp1.FranchiseeFeesCMG) = Val(.TextMatrix(I, ETemp1.FranchiseeFeesCMG)) + Val(.TextMatrix(0, ETemp1.FranchiseeFeesCMG))
           .TextMatrix(0, ETemp1.FranchiseeFeesFAG) = Val(.TextMatrix(I, ETemp1.FranchiseeFeesFAG)) + Val(.TextMatrix(0, ETemp1.FranchiseeFeesFAG))
           .TextMatrix(0, ETemp1.Salary) = Val(.TextMatrix(I, ETemp1.Salary)) + Val(.TextMatrix(0, ETemp1.Salary))
           .TextMatrix(0, ETemp1.SalaryFieldStaff) = Val(.TextMatrix(I, ETemp1.SalaryFieldStaff)) + Val(.TextMatrix(0, ETemp1.SalaryFieldStaff))
           .TextMatrix(0, ETemp1.NCDEXCharges) = Val(.TextMatrix(I, ETemp1.NCDEXCharges)) + Val(.TextMatrix(0, ETemp1.NCDEXCharges))
           .TextMatrix(0, ETemp1.CoExpense) = Val(.TextMatrix(I, ETemp1.CoExpense)) + Val(.TextMatrix(0, ETemp1.CoExpense))
           .TextMatrix(0, ETemp1.InsurancePremium) = Val(.TextMatrix(I, ETemp1.InsurancePremium)) + Val(.TextMatrix(0, ETemp1.InsurancePremium))
           .TextMatrix(0, ETemp1.InterestCharges) = Val(.TextMatrix(I, ETemp1.InterestCharges)) + Val(.TextMatrix(0, ETemp1.InterestCharges))
      End With
      TmpRs.MoveNext
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next I
End If
With MSHFlexGrid1
   For I = 2 To MSHFlexGrid1.Rows - 2
      .TextMatrix(I, ETemp1.TotalIncome) = (Val(.TextMatrix(I, ETemp1.NNNRIncomeAdhoc)) + Val(.TextMatrix(I, ETemp1.NNNRIncome)) + Val(.TextMatrix(I, ETemp1.ReservationIncomeAdhoc)) + Val(.TextMatrix(I, ETemp1.ReservationIncome)) + Val(.TextMatrix(I, ETemp1.SevenDaysIncome)) + Val(.TextMatrix(I, ETemp1.DPIncomeAdHoc)) + Val(.TextMatrix(I, ETemp1.FinalDPIncome)) + Val(.TextMatrix(I, ETemp1.RejectedIncome)) + Val(.TextMatrix(I, ETemp1.SpotIncomeAdHoc)) + Val(.TextMatrix(I, ETemp1.OtherIncome)) + Val(.TextMatrix(I, ETemp1.SCMIncome)) + Val(.TextMatrix(I, ETemp1.CMUnbilledIncome)) + Val(.TextMatrix(I, ETemp1.CMPendingIncome)) + Val(.TextMatrix(I, ETemp1.CMApprovedIncome)) + Val(.TextMatrix(I, ETemp1.LeftOverStock)) + Val(.TextMatrix(I, ETemp1.PledgeIncome))) - Val(.TextMatrix(I, ETemp1.PartiallyReversedIncome))
      .TextMatrix(I, ETemp1.TotalIncome) = Format(.TextMatrix(I, ETemp1.TotalIncome), "############.00")
      .TextMatrix(I, ETemp1.TotalExpense) = Val(.TextMatrix(I, ETemp1.LeaseRentPaid)) + Val(.TextMatrix(I, ETemp1.LeaseRentFAG)) + Val(.TextMatrix(I, ETemp1.LeaseRentDRC)) + Val(.TextMatrix(I, ETemp1.EmpClaimFAG)) + Val(.TextMatrix(I, ETemp1.EmpClaimAccrual)) + Val(.TextMatrix(I, ETemp1.VendorExpenses)) + Val(.TextMatrix(I, ETemp1.VendorExpensesAccrued)) + Val(.TextMatrix(I, ETemp1.SecurityFeesPaid)) + Val(.TextMatrix(I, ETemp1.SecurityFeesAccrual)) + Val(.TextMatrix(I, ETemp1.FumigationFeesVendor)) + Val(.TextMatrix(I, ETemp1.FumigationFeesPaid)) + Val(.TextMatrix(I, ETemp1.FumigationFeesAccrual)) + Val(.TextMatrix(I, ETemp1.FranchiseeFeesPaid)) + Val(.TextMatrix(I, ETemp1.FranchiseeFeesCMG)) + Val(.TextMatrix(I, ETemp1.FranchiseeFeesFAG)) + Val(.TextMatrix(I, ETemp1.Salary)) + Val(.TextMatrix(I, ETemp1.SalaryFieldStaff)) + Val(.TextMatrix(I, ETemp1.NCDEXCharges))
      .TextMatrix(I, ETemp1.TotalExpense) = Format(.TextMatrix(I, ETemp1.TotalExpense), "############.00")
      .TextMatrix(I, ETemp1.GrossIncome) = Val(.TextMatrix(I, ETemp1.TotalIncome)) - Val(.TextMatrix(I, ETemp1.TotalExpense))
      .TextMatrix(I, ETemp1.GrossIncome) = Format(.TextMatrix(I, ETemp1.GrossIncome), "############.00")
      .TextMatrix(I, ETemp1.NetIncome) = Val(.TextMatrix(I, ETemp1.GrossIncome)) - Val(.TextMatrix(I, ETemp1.CoExpense)) - Val(.TextMatrix(I, ETemp1.InsurancePremium)) - Val(.TextMatrix(I, ETemp1.InterestCharges))
      .TextMatrix(I, ETemp1.NetIncome) = Format(.TextMatrix(I, ETemp1.NetIncome), "############.00")
      .TextMatrix(0, ETemp1.TotalIncome) = Format(Val(.TextMatrix(0, ETemp1.TotalIncome)) + Val(.TextMatrix(I, ETemp1.TotalIncome)), "############.00")
      .TextMatrix(0, ETemp1.TotalExpense) = Format(Val(.TextMatrix(0, ETemp1.TotalExpense)) + Val(.TextMatrix(I, ETemp1.TotalExpense)), "############.00")
      .TextMatrix(0, ETemp1.GrossIncome) = Format(Val(.TextMatrix(0, ETemp1.GrossIncome)) + Val(.TextMatrix(I, ETemp1.GrossIncome)), "############.00")
      .TextMatrix(0, ETemp1.NetIncome) = Format(Val(.TextMatrix(0, ETemp1.NetIncome)) + Val(.TextMatrix(I, ETemp1.NetIncome)), "############.00")
   Next
End With
ProgressBar1.Value = ProgressBar1.Max


If Check2.Value = 1 Then
   
   Call Grid_Head2
   mYear = CmbYear.Text
   If CmbMonth = MonthName(1) Or CmbMonth = MonthName(2) Or CmbMonth = MonthName(3) Then
      mYear = mYear - 1
      mStartDate = mYear & "-" & 4 & "-1"
   ElseIf CmbMonth = MonthName(4) Or CmbMonth = MonthName(5) Or CmbMonth = MonthName(6) Or CmbMonth = MonthName(7) Or CmbMonth = MonthName(8) Or CmbMonth = MonthName(9) Or CmbMonth = MonthName(10) Or CmbMonth = MonthName(11) Or CmbMonth = MonthName(12) Then
      mStartDate = mYear & "-" & 4 & "-1"
   End If
       
   mCurrentStartDate = CmbYear.Text & "-" & mMonth & "-1"
   mEndDate = CmbYear.Text & "-" & mMonth & "-" & mTotalDays
   mPreviousEndDate = CDate(mCurrentStartDate) - 1
    
    'mTotalMonths = DateDiff("M", mWideStartDate, mWideEndDate) + 1
    'mStartDate = Format(TxtFromDate.Value, "yyyy-mm-dd")
    'mEndDate = Format(TxtToDate.Value, "yyyy-mm-dd")
    
    
   tmp = " Select StateName, LocationName,AM, RM, CMPName, GodownNo, CMPName+'_'+GodownNo'KeyFoGodown',"
   tmp = tmp & " UnitType, StartDate, CloseDate, Capacity,"
   tmp = tmp & " sum(NcdexStock) NcdexStock, sum(NNStock) NNStock, sum(NNNRIncomeAdhoc) NNNRIncomeAdhoc,"
   tmp = tmp & " sum(NNNRIncome) NNNRIncome, sum(ReservationIncomeAdhoc) ReservationIncomeAdhoc,"
   tmp = tmp & " sum(ReservationIncome) ReservationIncome, sum(SevenDaysIncome) SevenDaysIncome,"
   tmp = tmp & " sum(DPIncomeAdHoc) DPIncomeAdHoc, sum(FinalDPIncome) FinalDPIncome,  sum(RejectedIncome) RejectedIncome,"
   tmp = tmp & " sum(SpotIncomeAdHoc) SpotIncomeAdHoc, sum(OtherIncome) OtherIncome, sum(SCMIncome) SCMIncome,"
   tmp = tmp & " sum(CMUnbilledIncome) CMUnbilledIncome, sum(CMPendingIncome) CMPendingIncome,"
   tmp = tmp & " sum(CMApprovedIncome) CMApprovedIncome, sum(LeftOverStock) LeftOverStock,"
   tmp = tmp & " sum(PledgeIncome) PledgeIncome,  sum(LeaseRentPaid) LeaseRentPaid, sum(LeaseRentFAG) LeaseRentFAG,"
   tmp = tmp & " sum(LeaseRentDRC) LeaseRentDRC, sum(EmpClaimFAG) EmpClaimFAG, sum(EmpClaimAccrual) EmpClaimAccrual,"
   tmp = tmp & " sum(VendorExpenses) VendorExpenses, sum(VendorExpensesAccrued) VendorExpensesAccrued,"
   tmp = tmp & " sum(SecurityFeesPaid) SecurityFeesPaid, sum(SecurityFeesAccrual) SecurityFeesAccrual,"
   tmp = tmp & " sum(FumigationFeesVendor) FumigationFeesVendor, sum(FumigationFeesPaid) FumigationFeesPaid,"
   tmp = tmp & " sum(FumigationFeesAccrual) FumigationFeesAccrual, sum(FranchiseeFeesPaid) FranchiseeFeesPaid,"
   tmp = tmp & " sum(FranchiseeFeesFAG) FranchiseeFeesFAG, sum(FranchiseeFeesCMG) FranchiseeFeesCMG,"
   tmp = tmp & " (Sum(Salary)) Salary , (Sum(SalaryFieldStaff)) SalaryFieldStaff ,"
   tmp = tmp & " sum(NCDEXCharges) NCDEXCharges, sum(CoExpense) CoExpense, sum(InsurancePremium) InsurancePremium,"
   tmp = tmp & " sum(InterestCharges) InterestCharges,sum(PartiallyReversedIncome) PartiallyReversedIncome,"
   tmp = tmp & " StateID , LocationID, GodownID, Sum(DPIncomeGSL) DPIncomeGSL, Sum(Diff) Diff,"
   tmp = tmp & " Sum(DPIncomeInvoice) DPIncomeInvoice"
   tmp = tmp & " From"
   tmp = tmp & " ((Select MS.StateName, ML.LocationName, TMGP.AM, TMGP.RM, MCMP.CMPName, MG.GodownNo, MCMP.CMPName+'_'+MG.GodownNo'KeyFoGodown',"
   tmp = tmp & " MUT.UnitType, MG.StartDate, MG.CloseDate, Max(MG.GodownCapacity) Capacity,"
   tmp = tmp & " sum(TMGP.NcdexStock) NcdexStock, sum(TMGP.NNStock) NNStock, sum(TMGP.NNNRIncomeAdhoc) NNNRIncomeAdhoc,"
   tmp = tmp & " sum(TMGP.NNNRIncome) NNNRIncome, sum(TMGP.ReservationIncomeAdhoc) ReservationIncomeAdhoc,"
   tmp = tmp & " sum(TMGP.ReservationIncome) ReservationIncome, sum(TMGP.SevenDaysIncome) SevenDaysIncome,"
   tmp = tmp & " sum(TMGP.DPIncomeAdHoc) DPIncomeAdHoc, sum(TMGP.FinalDPIncome) FinalDPIncome,  sum(TMGP.RejectedIncome) RejectedIncome,"
   tmp = tmp & " sum(TMGP.SpotIncomeAdHoc) SpotIncomeAdHoc, sum(TMGP.OtherIncome) OtherIncome, sum(TMGP.SCMIncome) SCMIncome,"
   tmp = tmp & " sum(TMGP.CMUnbilledIncome) CMUnbilledIncome, sum(TMGP.CMPendingIncome) CMPendingIncome,"
   tmp = tmp & " sum(TMGP.CMApprovedIncome) CMApprovedIncome, sum(TMGP.LeftOverStock) LeftOverStock,"
   tmp = tmp & " sum(TMGP.PledgeIncome) PledgeIncome,  sum(TMGP.LeaseRentPaid) LeaseRentPaid, sum(TMGP.LeaseRentFAG) LeaseRentFAG,"
   tmp = tmp & " sum(TMGP.LeaseRentDRC) LeaseRentDRC, sum(TMGP.EmpClaimFAG) EmpClaimFAG, sum(TMGP.EmpClaimAccrual) EmpClaimAccrual,"
   tmp = tmp & " sum(TMGP.VendorExpenses) VendorExpenses, sum(TMGP.VendorExpensesAccrued) VendorExpensesAccrued,"
   tmp = tmp & " sum(TMGP.SecurityFeesPaid) SecurityFeesPaid, sum(TMGP.SecurityFeesAccrual) SecurityFeesAccrual,"
   tmp = tmp & " sum(TMGP.FumigationFeesVendor) FumigationFeesVendor, sum(TMGP.FumigationFeesPaid) FumigationFeesPaid,"
   tmp = tmp & " sum(TMGP.FumigationFeesAccrual) FumigationFeesAccrual, sum(TMGP.FranchiseeFeesPaid) FranchiseeFeesPaid,"
   tmp = tmp & " sum(TMGP.FranchiseeFeesFAG) FranchiseeFeesFAG, sum(TMGP.FranchiseeFeesCMG) FranchiseeFeesCMG,"
   tmp = tmp & " (Sum(TMGP.Salary)) Salary , (Sum(TMGP.SalaryFieldStaff)) SalaryFieldStaff ,"
   tmp = tmp & " sum(TMGP.NCDEXCharges) NCDEXCharges, sum(TMGP.CoExpense) CoExpense, sum(TMGP.InsurancePremium) InsurancePremium,"
   tmp = tmp & " sum(TMGP.InterestCharges) InterestCharges,sum(TMGP.PartiallyReversedIncome) PartiallyReversedIncome,"
   tmp = tmp & " TMGP.StateID , TMGP.LocationID, TMGP.GodownID, Sum(TMGP.DPIncomeGSL) DPIncomeGSL, Sum(TMGP.Diff) Diff,"
   tmp = tmp & " Sum(TMGP.DPIncomeInvoice) DPIncomeInvoice"
   tmp = tmp & " From Txn_MonGodownWiseProfitabilityFreezed TMGP"
   tmp = tmp & " left Join Mst_Location ML On ML.LocationID = TMGP.LocationID"
   tmp = tmp & " left Join Mst_State MS On MS.StateID =  TMGP.StateID"
   tmp = tmp & " left Join Mst_Godown MG On MG.GodownID = TMGP.GodownID"
   tmp = tmp & " left Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
   tmp = tmp & " left Join Mst_CMP MCMP On MCMP.CMPID = MG.CMPID"
   tmp = tmp & " Where TMGP.MonthYear between '" & mCurrentStartDate & "' and '" & mEndDate & "' And TMGP.ReportFlag = 'C'"
   tmp = tmp & " Group by MS.StateName, ML.LocationName, TMGP.AM, TMGP.RM, MCMP.CMPName, MG.GodownNo, MUT.UnitType, MG.StartDate,"
   tmp = tmp & " MG.CloseDate, TMGP.StateID , TMGP.LocationID, TMGP.GodownID )"
   tmp = tmp & " Union All"
   tmp = tmp & " (Select MS.StateName, ML.LocationName, TMGP.AM, TMGP.RM, MCMP.CMPName, MG.GodownNo, MCMP.CMPName+'_'+MG.GodownNo 'KeyFoGodown',"
   tmp = tmp & " MUT.UnitType, MG.StartDate, MG.CloseDate, Max(MG.GodownCapacity) Capacity,"
   tmp = tmp & " -sum(TMGP.NcdexStock) NcdexStock, -sum(TMGP.NNStock) NNStock, -sum(TMGP.NNNRIncomeAdhoc) NNNRIncomeAdhoc,"
   tmp = tmp & " -sum(TMGP.NNNRIncome) NNNRIncome, -sum(TMGP.ReservationIncomeAdhoc) ReservationIncomeAdhoc,"
   tmp = tmp & " -sum(TMGP.ReservationIncome) ReservationIncome, -sum(TMGP.SevenDaysIncome) SevenDaysIncome,"
   tmp = tmp & " -sum(TMGP.DPIncomeAdHoc) DPIncomeAdHoc, -sum(TMGP.FinalDPIncome) FinalDPIncome,  -sum(TMGP.RejectedIncome) RejectedIncome,"
   tmp = tmp & " -sum(TMGP.SpotIncomeAdHoc) SpotIncomeAdHoc, -sum(TMGP.OtherIncome) OtherIncome, -sum(TMGP.SCMIncome) SCMIncome,"
   tmp = tmp & " -sum(TMGP.CMUnbilledIncome) CMUnbilledIncome, -sum(TMGP.CMPendingIncome) CMPendingIncome,"
   tmp = tmp & " -sum(TMGP.CMApprovedIncome) CMApprovedIncome, -sum(TMGP.LeftOverStock) LeftOverStock,"
   tmp = tmp & " -sum(TMGP.PledgeIncome) PledgeIncome,  -sum(TMGP.LeaseRentPaid) LeaseRentPaid, -sum(TMGP.LeaseRentFAG) LeaseRentFAG,"
   tmp = tmp & " -sum(TMGP.LeaseRentDRC) LeaseRentDRC, -sum(TMGP.EmpClaimFAG) EmpClaimFAG, -sum(TMGP.EmpClaimAccrual) EmpClaimAccrual,"
   tmp = tmp & " -sum(TMGP.VendorExpenses) VendorExpenses, -sum(TMGP.VendorExpensesAccrued) VendorExpensesAccrued,"
   tmp = tmp & " -sum(TMGP.SecurityFeesPaid) SecurityFeesPaid, -sum(TMGP.SecurityFeesAccrual) SecurityFeesAccrual,"
   tmp = tmp & " -sum(TMGP.FumigationFeesVendor) FumigationFeesVendor, -sum(TMGP.FumigationFeesPaid) FumigationFeesPaid,"
   tmp = tmp & " -sum(TMGP.FumigationFeesAccrual) FumigationFeesAccrual, -sum(TMGP.FranchiseeFeesPaid) FranchiseeFeesPaid,"
   tmp = tmp & " -sum(TMGP.FranchiseeFeesFAG) FranchiseeFeesFAG, -sum(TMGP.FranchiseeFeesCMG) FranchiseeFeesCMG,"
   tmp = tmp & " (-Sum(TMGP.Salary)) Salary , (-Sum(TMGP.SalaryFieldStaff)) SalaryFieldStaff ,"
   tmp = tmp & " -sum(TMGP.NCDEXCharges) NCDEXCharges, -sum(TMGP.CoExpense) CoExpense, -sum(TMGP.InsurancePremium) InsurancePremium,"
   tmp = tmp & " -sum(TMGP.InterestCharges) InterestCharges,-sum(TMGP.PartiallyReversedIncome) PartiallyReversedIncome,"
   tmp = tmp & " TMGP.StateID , TMGP.LocationID, TMGP.GodownID, -Sum(TMGP.DPIncomeGSL) DPIncomeGSL, -Sum(TMGP.Diff) Diff,"
   tmp = tmp & " -Sum(TMGP.DPIncomeInvoice) DPIncomeInvoice"
   tmp = tmp & " From Txn_MonGodownWiseProfitabilityFreezed TMGP"
   tmp = tmp & " left Join Mst_Location ML On ML.LocationID = TMGP.LocationID"
   tmp = tmp & " left Join Mst_State MS On MS.StateID =  TMGP.StateID"
   tmp = tmp & " left Join Mst_Godown MG On MG.GodownID = TMGP.GodownID"
   tmp = tmp & " left Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
   tmp = tmp & " left Join Mst_CMP MCMP On MCMP.CMPID = MG.CMPID"
   tmp = tmp & " Where TMGP.MonthYear between '" & mStartDate & "' and '" & mPreviousEndDate & "' And TMGP.ReportFlag = 'C'"
   tmp = tmp & " Group by MS.StateName, ML.LocationName, TMGP.AM, TMGP.RM, MCMP.CMPName, MG.GodownNo, MUT.UnitType, MG.StartDate,"
   tmp = tmp & " MG.CloseDate, TMGP.StateID , TMGP.LocationID, TMGP.GodownID )) A"
   tmp = tmp & " Group by StateName, LocationName, AM, RM, CMPName, GodownNo, UnitType, StartDate,"
   tmp = tmp & " CloseDate , StateID, LocationID, GodownID, Capacity"
   tmp = tmp & " Order by StateName,LocationName, CMPName,GodownNo"

   Call TmpTable
    
   If TmpRs.RecordCount > 0 Then
'      Label5.Caption = TmpRs.RecordCount
      ProgressBar1.Value = 0
      ProgressBar1.Max = MSHFlexGrid2.Rows + TmpRs.RecordCount
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
      CmdExcel.Enabled = True
       
      For I = 2 To TmpRs.RecordCount + 1
          With MSHFlexGrid2
              .TextMatrix(I, ETemp2.MonthYear) = mYearMonth
              .TextMatrix(I, ETemp2.StateName) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
              .TextMatrix(I, ETemp2.LocationName) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
              .TextMatrix(I, ETemp2.AMname) = IIf(IsNull(TmpRs!AM), "", TmpRs!AM)
              .TextMatrix(I, ETemp2.RMName) = IIf(IsNull(TmpRs!RM), "", TmpRs!RM)
              .TextMatrix(I, ETemp2.CMPName) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
              .TextMatrix(I, ETemp2.GodownNo) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
              .TextMatrix(I, ETemp2.GodownCapacity) = IIf(IsNull(TmpRs!Capacity), "", TmpRs!Capacity)
              .TextMatrix(I, ETemp2.NcdexStock) = IIf(IsNull(TmpRs!NcdexStock), "", TmpRs!NcdexStock)
              .TextMatrix(I, ETemp2.NonNCDEXStock) = IIf(IsNull(TmpRs!NNStock), "", TmpRs!NNStock)
              .TextMatrix(I, ETemp2.NNNRIncomeAdhoc) = IIf(IsNull(TmpRs!NNNRIncomeAdhoc), "", TmpRs!NNNRIncomeAdhoc)
              .TextMatrix(I, ETemp2.NNNRIncome) = IIf(IsNull(TmpRs!NNNRIncome), "", TmpRs!NNNRIncome)
              .TextMatrix(I, ETemp2.ReservationIncomeAdhoc) = IIf(IsNull(TmpRs!ReservationIncomeAdhoc), "", TmpRs!ReservationIncomeAdhoc)
              .TextMatrix(I, ETemp2.ReservationIncome) = IIf(IsNull(TmpRs!ReservationIncome), "", TmpRs!ReservationIncome)
              .TextMatrix(I, ETemp2.SevenDaysIncome) = IIf(IsNull(TmpRs!SevenDaysIncome), "", TmpRs!SevenDaysIncome)
              .TextMatrix(I, ETemp2.DPIncomeAdHoc) = IIf(IsNull(TmpRs!DPIncomeAdHoc), "", TmpRs!DPIncomeAdHoc)
              .TextMatrix(I, ETemp2.FinalDPIncome) = IIf(IsNull(TmpRs!FinalDPIncome), "", TmpRs!FinalDPIncome)
              .TextMatrix(I, ETemp2.DPIncomeasperGSL) = IIf(IsNull(TmpRs!DPIncomeGSL), "", TmpRs!DPIncomeGSL)
              .TextMatrix(I, ETemp2.Difference) = IIf(IsNull(TmpRs!Diff), "", TmpRs!Diff)
              .TextMatrix(I, ETemp2.DPIncomeasperInvoice) = IIf(IsNull(TmpRs!DPIncomeInvoice), "", TmpRs!DPIncomeInvoice)
              .TextMatrix(I, ETemp2.RejectedIncome) = IIf(IsNull(TmpRs!RejectedIncome), "", TmpRs!RejectedIncome)
              .TextMatrix(I, ETemp2.SpotIncomeAdHoc) = IIf(IsNull(TmpRs!SpotIncomeAdHoc), "", TmpRs!SpotIncomeAdHoc)
              .TextMatrix(I, ETemp2.OtherIncome) = IIf(IsNull(TmpRs!OtherIncome), "", TmpRs!OtherIncome)
              .TextMatrix(I, ETemp2.SCMIncome) = IIf(IsNull(TmpRs!SCMIncome), "", TmpRs!SCMIncome)
              .TextMatrix(I, ETemp2.CMUnbilledIncome) = IIf(IsNull(TmpRs!CMUnbilledIncome), "", TmpRs!CMUnbilledIncome)
              .TextMatrix(I, ETemp2.CMPendingIncome) = IIf(IsNull(TmpRs!CMPendingIncome), "", TmpRs!CMPendingIncome)
              .TextMatrix(I, ETemp2.CMApprovedIncome) = IIf(IsNull(TmpRs!CMApprovedIncome), "", TmpRs!CMApprovedIncome)
              .TextMatrix(I, ETemp2.LeftOverStock) = IIf(IsNull(TmpRs!LeftOverStock), "", TmpRs!LeftOverStock)
              .TextMatrix(I, ETemp2.PledgeIncome) = IIf(IsNull(TmpRs!PledgeIncome), "", TmpRs!PledgeIncome)
              .TextMatrix(I, ETemp2.PartiallyReversedIncome) = IIf(IsNull(TmpRs!PartiallyReversedIncome), "", TmpRs!PartiallyReversedIncome)
              .TextMatrix(I, ETemp2.LeaseRentPaid) = IIf(IsNull(TmpRs!LeaseRentPaid), "", TmpRs!LeaseRentPaid)
              .TextMatrix(I, ETemp2.LeaseRentFAG) = IIf(IsNull(TmpRs!LeaseRentFAG), "", TmpRs!LeaseRentFAG)
              .TextMatrix(I, ETemp2.LeaseRentDRC) = IIf(IsNull(TmpRs!LeaseRentDRC), "", TmpRs!LeaseRentDRC)
              .TextMatrix(I, ETemp2.EmpClaimFAG) = IIf(IsNull(TmpRs!EmpClaimFAG), "", TmpRs!EmpClaimFAG)
              .TextMatrix(I, ETemp2.EmpClaimAccrual) = IIf(IsNull(TmpRs!EmpClaimAccrual), "", TmpRs!EmpClaimAccrual)
              .TextMatrix(I, ETemp2.VendorExpenses) = IIf(IsNull(TmpRs!VendorExpenses), "", TmpRs!VendorExpenses)
              .TextMatrix(I, ETemp2.VendorExpensesAccrued) = IIf(IsNull(TmpRs!VendorExpensesAccrued), "", TmpRs!VendorExpensesAccrued)
              .TextMatrix(I, ETemp2.SecurityFeesPaid) = IIf(IsNull(TmpRs!SecurityFeesPaid), "", TmpRs!SecurityFeesPaid)
              .TextMatrix(I, ETemp2.SecurityFeesAccrual) = IIf(IsNull(TmpRs!SecurityFeesAccrual), "", TmpRs!SecurityFeesAccrual)
              .TextMatrix(I, ETemp2.FumigationFeesVendor) = IIf(IsNull(TmpRs!FumigationFeesVendor), "", TmpRs!FumigationFeesVendor)
              .TextMatrix(I, ETemp2.FumigationFeesPaid) = IIf(IsNull(TmpRs!FumigationFeesPaid), "", TmpRs!FumigationFeesPaid)
              .TextMatrix(I, ETemp2.FumigationFeesAccrual) = IIf(IsNull(TmpRs!FumigationFeesAccrual), "", TmpRs!FumigationFeesAccrual)
              .TextMatrix(I, ETemp2.FranchiseeFeesPaid) = IIf(IsNull(TmpRs!FranchiseeFeesPaid), "", TmpRs!FranchiseeFeesPaid)
              .TextMatrix(I, ETemp2.FranchiseeFeesCMG) = IIf(IsNull(TmpRs!FranchiseeFeesCMG), "", TmpRs!FranchiseeFeesCMG)
              .TextMatrix(I, ETemp2.FranchiseeFeesFAG) = IIf(IsNull(TmpRs!FranchiseeFeesFAG), "", TmpRs!FranchiseeFeesFAG)
              .TextMatrix(I, ETemp2.Salary) = IIf(IsNull(TmpRs!Salary), "", TmpRs!Salary)
              .TextMatrix(I, ETemp2.SalaryFieldStaff) = IIf(IsNull(TmpRs!SalaryFieldStaff), "", TmpRs!SalaryFieldStaff)
              .TextMatrix(I, ETemp2.NCDEXCharges) = IIf(IsNull(TmpRs!NCDEXCharges), "", TmpRs!NCDEXCharges)
              .TextMatrix(I, ETemp2.CoExpense) = IIf(IsNull(TmpRs!CoExpense), "", TmpRs!CoExpense)
              .TextMatrix(I, ETemp2.InsurancePremium) = IIf(IsNull(TmpRs!InsurancePremium), "", TmpRs!InsurancePremium)
              .TextMatrix(I, ETemp2.InterestCharges) = IIf(IsNull(TmpRs!InterestCharges), "", TmpRs!InterestCharges)
              .TextMatrix(I, ETemp2.GodownStartDate) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
              .TextMatrix(I, ETemp2.GodownCloseDate) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
              .TextMatrix(I, ETemp2.UnitType) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
              .TextMatrix(I, ETemp2.StateID) = IIf(IsNull(TmpRs!StateID), "", TmpRs!StateID)
              .TextMatrix(I, ETemp2.LocationID) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
              .TextMatrix(I, ETemp2.GodownID) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
              
              .TextMatrix(0, ETemp2.NcdexStock) = Val(.TextMatrix(I, ETemp2.NcdexStock)) + Val(.TextMatrix(0, ETemp2.NcdexStock))
              .TextMatrix(0, ETemp2.GodownCapacity) = Val(.TextMatrix(I, ETemp2.GodownCapacity)) + Val(.TextMatrix(0, ETemp2.GodownCapacity))
              .TextMatrix(0, ETemp2.NonNCDEXStock) = Val(.TextMatrix(I, ETemp2.NonNCDEXStock)) + Val(.TextMatrix(0, ETemp2.NonNCDEXStock))
              .TextMatrix(0, ETemp2.NNNRIncomeAdhoc) = Val(.TextMatrix(I, ETemp2.NNNRIncomeAdhoc)) + Val(.TextMatrix(0, ETemp2.NNNRIncomeAdhoc))
              .TextMatrix(0, ETemp2.NNNRIncome) = Val(.TextMatrix(I, ETemp2.NNNRIncome)) + Val(.TextMatrix(0, ETemp2.NNNRIncome))
              .TextMatrix(0, ETemp2.ReservationIncomeAdhoc) = Val(.TextMatrix(I, ETemp2.ReservationIncomeAdhoc)) + Val(.TextMatrix(0, ETemp2.ReservationIncomeAdhoc))
              .TextMatrix(0, ETemp2.ReservationIncome) = Val(.TextMatrix(I, ETemp2.ReservationIncome)) + Val(.TextMatrix(0, ETemp2.ReservationIncome))
              .TextMatrix(0, ETemp2.SevenDaysIncome) = Val(.TextMatrix(I, ETemp2.SevenDaysIncome)) + Val(.TextMatrix(0, ETemp2.SevenDaysIncome))
              .TextMatrix(0, ETemp2.DPIncomeAdHoc) = Val(.TextMatrix(I, ETemp2.DPIncomeAdHoc)) + Val(.TextMatrix(0, ETemp2.DPIncomeAdHoc))
              .TextMatrix(0, ETemp2.FinalDPIncome) = Val(.TextMatrix(I, ETemp2.FinalDPIncome)) + Val(.TextMatrix(0, ETemp2.FinalDPIncome))
              .TextMatrix(0, ETemp2.DPIncomeasperGSL) = Val(.TextMatrix(I, ETemp2.DPIncomeasperGSL)) + Val(.TextMatrix(0, ETemp2.DPIncomeasperGSL))
              .TextMatrix(0, ETemp2.Difference) = Val(.TextMatrix(I, ETemp2.Difference)) + Val(.TextMatrix(0, ETemp2.Difference))
              .TextMatrix(0, ETemp2.DPIncomeasperInvoice) = Val(.TextMatrix(I, ETemp2.DPIncomeasperInvoice)) + Val(.TextMatrix(0, ETemp2.DPIncomeasperInvoice))
              .TextMatrix(0, ETemp2.RejectedIncome) = Val(.TextMatrix(I, ETemp2.RejectedIncome)) + Val(.TextMatrix(0, ETemp2.RejectedIncome))
              .TextMatrix(0, ETemp2.SpotIncomeAdHoc) = Val(.TextMatrix(I, ETemp2.SpotIncomeAdHoc)) + Val(.TextMatrix(0, ETemp2.SpotIncomeAdHoc))
              .TextMatrix(0, ETemp2.OtherIncome) = Val(.TextMatrix(I, ETemp2.OtherIncome)) + Val(.TextMatrix(0, ETemp2.OtherIncome))
              .TextMatrix(0, ETemp2.SCMIncome) = Val(.TextMatrix(I, ETemp2.SCMIncome)) + Val(.TextMatrix(0, ETemp2.SCMIncome))
              .TextMatrix(0, ETemp2.CMUnbilledIncome) = Val(.TextMatrix(I, ETemp2.CMUnbilledIncome)) + Val(.TextMatrix(0, ETemp2.CMUnbilledIncome))
              .TextMatrix(0, ETemp2.CMPendingIncome) = Val(.TextMatrix(I, ETemp2.CMPendingIncome)) + Val(.TextMatrix(0, ETemp2.CMPendingIncome))
              .TextMatrix(0, ETemp2.CMApprovedIncome) = Val(.TextMatrix(I, ETemp2.CMApprovedIncome)) + Val(.TextMatrix(0, ETemp2.CMApprovedIncome))
              .TextMatrix(0, ETemp2.LeftOverStock) = Val(.TextMatrix(I, ETemp2.LeftOverStock)) + Val(.TextMatrix(0, ETemp2.LeftOverStock))
              .TextMatrix(0, ETemp2.PledgeIncome) = Val(.TextMatrix(I, ETemp2.PledgeIncome)) + Val(.TextMatrix(0, ETemp2.PledgeIncome))
              .TextMatrix(0, ETemp2.PartiallyReversedIncome) = Val(.TextMatrix(I, ETemp2.PartiallyReversedIncome)) + Val(.TextMatrix(0, ETemp2.PartiallyReversedIncome))
              .TextMatrix(0, ETemp2.LeaseRentPaid) = Val(.TextMatrix(I, ETemp2.LeaseRentPaid)) + Val(.TextMatrix(0, ETemp2.LeaseRentPaid))
              .TextMatrix(0, ETemp2.LeaseRentFAG) = Val(.TextMatrix(I, ETemp2.LeaseRentFAG)) + Val(.TextMatrix(0, ETemp2.LeaseRentFAG))
              .TextMatrix(0, ETemp2.LeaseRentDRC) = Val(.TextMatrix(I, ETemp2.LeaseRentDRC)) + Val(.TextMatrix(0, ETemp2.LeaseRentDRC))
              .TextMatrix(0, ETemp2.EmpClaimFAG) = Val(.TextMatrix(I, ETemp2.EmpClaimFAG)) + Val(.TextMatrix(0, ETemp2.EmpClaimFAG))
              .TextMatrix(0, ETemp2.EmpClaimAccrual) = Val(.TextMatrix(I, ETemp2.EmpClaimAccrual)) + Val(.TextMatrix(0, ETemp2.EmpClaimAccrual))
              .TextMatrix(0, ETemp2.VendorExpenses) = Val(.TextMatrix(I, ETemp2.VendorExpenses)) + Val(.TextMatrix(0, ETemp2.VendorExpenses))
              .TextMatrix(0, ETemp2.VendorExpensesAccrued) = Val(.TextMatrix(I, ETemp2.VendorExpensesAccrued)) + Val(.TextMatrix(0, ETemp2.VendorExpensesAccrued))
              .TextMatrix(0, ETemp2.SecurityFeesPaid) = Val(.TextMatrix(I, ETemp2.SecurityFeesPaid)) + Val(.TextMatrix(0, ETemp2.SecurityFeesPaid))
              .TextMatrix(0, ETemp2.SecurityFeesAccrual) = Val(.TextMatrix(I, ETemp2.SecurityFeesAccrual)) + Val(.TextMatrix(0, ETemp2.SecurityFeesAccrual))
              .TextMatrix(0, ETemp2.FumigationFeesVendor) = Val(.TextMatrix(I, ETemp2.FumigationFeesVendor)) + Val(.TextMatrix(0, ETemp2.FumigationFeesVendor))
              .TextMatrix(0, ETemp2.FumigationFeesPaid) = Val(.TextMatrix(I, ETemp2.FumigationFeesPaid)) + Val(.TextMatrix(0, ETemp2.FumigationFeesPaid))
              .TextMatrix(0, ETemp2.FumigationFeesAccrual) = Val(.TextMatrix(I, ETemp2.FumigationFeesAccrual)) + Val(.TextMatrix(0, ETemp2.FumigationFeesAccrual))
              .TextMatrix(0, ETemp2.FranchiseeFeesPaid) = Val(.TextMatrix(I, ETemp2.FranchiseeFeesPaid)) + Val(.TextMatrix(0, ETemp2.FranchiseeFeesPaid))
              .TextMatrix(0, ETemp2.FranchiseeFeesCMG) = Val(.TextMatrix(I, ETemp2.FranchiseeFeesCMG)) + Val(.TextMatrix(0, ETemp2.FranchiseeFeesCMG))
              .TextMatrix(0, ETemp2.FranchiseeFeesFAG) = Val(.TextMatrix(I, ETemp2.FranchiseeFeesFAG)) + Val(.TextMatrix(0, ETemp2.FranchiseeFeesFAG))
              .TextMatrix(0, ETemp2.Salary) = Val(.TextMatrix(I, ETemp2.Salary)) + Val(.TextMatrix(0, ETemp2.Salary))
              .TextMatrix(0, ETemp2.SalaryFieldStaff) = Val(.TextMatrix(I, ETemp2.SalaryFieldStaff)) + Val(.TextMatrix(0, ETemp2.SalaryFieldStaff))
              .TextMatrix(0, ETemp2.NCDEXCharges) = Val(.TextMatrix(I, ETemp2.NCDEXCharges)) + Val(.TextMatrix(0, ETemp2.NCDEXCharges))
              .TextMatrix(0, ETemp2.CoExpense) = Val(.TextMatrix(I, ETemp2.CoExpense)) + Val(.TextMatrix(0, ETemp2.CoExpense))
              .TextMatrix(0, ETemp2.InsurancePremium) = Val(.TextMatrix(I, ETemp2.InsurancePremium)) + Val(.TextMatrix(0, ETemp2.InsurancePremium))
              .TextMatrix(0, ETemp2.InterestCharges) = Val(.TextMatrix(I, ETemp2.InterestCharges)) + Val(.TextMatrix(0, ETemp2.InterestCharges))
         End With
         TmpRs.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next I
   End If
   With MSHFlexGrid2
        For I = 2 To MSHFlexGrid2.Rows - 2
            .TextMatrix(I, ETemp2.TotalIncome) = (Val(.TextMatrix(I, ETemp2.NNNRIncomeAdhoc)) + Val(.TextMatrix(I, ETemp2.NNNRIncome)) + Val(.TextMatrix(I, ETemp2.ReservationIncomeAdhoc)) + Val(.TextMatrix(I, ETemp2.ReservationIncome)) + Val(.TextMatrix(I, ETemp2.SevenDaysIncome)) + Val(.TextMatrix(I, ETemp2.DPIncomeAdHoc)) + Val(.TextMatrix(I, ETemp2.FinalDPIncome)) + Val(.TextMatrix(I, ETemp2.RejectedIncome)) + Val(.TextMatrix(I, ETemp2.SpotIncomeAdHoc)) + Val(.TextMatrix(I, ETemp2.OtherIncome)) + Val(.TextMatrix(I, ETemp2.SCMIncome)) + Val(.TextMatrix(I, ETemp2.CMUnbilledIncome)) + Val(.TextMatrix(I, ETemp2.CMPendingIncome)) + Val(.TextMatrix(I, ETemp2.CMApprovedIncome)) + Val(.TextMatrix(I, ETemp2.LeftOverStock)) + Val(.TextMatrix(I, ETemp2.PledgeIncome))) - Val(.TextMatrix(I, ETemp2.PartiallyReversedIncome))
            .TextMatrix(I, ETemp2.TotalIncome) = Format(.TextMatrix(I, ETemp2.TotalIncome), "############.00")
            .TextMatrix(I, ETemp2.TotalExpense) = Val(.TextMatrix(I, ETemp2.LeaseRentPaid)) + Val(.TextMatrix(I, ETemp2.LeaseRentFAG)) + Val(.TextMatrix(I, ETemp2.LeaseRentDRC)) + Val(.TextMatrix(I, ETemp2.EmpClaimFAG)) + Val(.TextMatrix(I, ETemp2.EmpClaimAccrual)) + Val(.TextMatrix(I, ETemp2.VendorExpenses)) + Val(.TextMatrix(I, ETemp2.VendorExpensesAccrued)) + Val(.TextMatrix(I, ETemp2.SecurityFeesPaid)) + Val(.TextMatrix(I, ETemp2.SecurityFeesAccrual)) + Val(.TextMatrix(I, ETemp2.FumigationFeesVendor)) + Val(.TextMatrix(I, ETemp2.FumigationFeesPaid)) + Val(.TextMatrix(I, ETemp2.FumigationFeesAccrual)) + Val(.TextMatrix(I, ETemp2.FranchiseeFeesPaid)) + Val(.TextMatrix(I, ETemp2.FranchiseeFeesCMG)) + Val(.TextMatrix(I, ETemp2.FranchiseeFeesFAG)) + Val(.TextMatrix(I, ETemp2.Salary)) + Val(.TextMatrix(I, ETemp2.SalaryFieldStaff)) + Val(.TextMatrix(I, ETemp2.NCDEXCharges))
            .TextMatrix(I, ETemp2.TotalExpense) = Format(.TextMatrix(I, ETemp2.TotalExpense), "############.00")
            .TextMatrix(I, ETemp2.GrossIncome) = Val(.TextMatrix(I, ETemp2.TotalIncome)) - Val(.TextMatrix(I, ETemp2.TotalExpense))
            .TextMatrix(I, ETemp2.GrossIncome) = Format(.TextMatrix(I, ETemp2.GrossIncome), "############.00")
            .TextMatrix(I, ETemp2.NetIncome) = Val(.TextMatrix(I, ETemp2.GrossIncome)) - Val(.TextMatrix(I, ETemp2.CoExpense)) - Val(.TextMatrix(I, ETemp2.InsurancePremium)) - Val(.TextMatrix(I, ETemp2.InterestCharges))
            .TextMatrix(I, ETemp2.NetIncome) = Format(.TextMatrix(I, ETemp2.NetIncome), "############.00")
            .TextMatrix(0, ETemp2.TotalIncome) = Format(Val(.TextMatrix(0, ETemp2.TotalIncome)) + Val(.TextMatrix(I, ETemp2.TotalIncome)), "############.00")
            .TextMatrix(0, ETemp2.TotalExpense) = Format(Val(.TextMatrix(0, ETemp2.TotalExpense)) + Val(.TextMatrix(I, ETemp2.TotalExpense)), "############.00")
            .TextMatrix(0, ETemp2.GrossIncome) = Format(Val(.TextMatrix(0, ETemp2.GrossIncome)) + Val(.TextMatrix(I, ETemp2.GrossIncome)), "############.00")
            .TextMatrix(0, ETemp2.NetIncome) = Format(Val(.TextMatrix(0, ETemp2.NetIncome)) + Val(.TextMatrix(I, ETemp2.NetIncome)), "############.00")
        Next
   End With
   ProgressBar1.Value = ProgressBar1.Max

End If

MsgBox "Done!!!"
End Sub

Private Sub Grid_Head()
CmdExcel.Enabled = False
Label5.Caption = ""
ProgressBar1.Value = 0
With MSHFlexGrid1
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 61 ' 60 '59
   
   .TextMatrix(1, ETemp1.MonthYear) = "Month Year"
   .TextMatrix(1, ETemp1.StateName) = "State Name"
   .TextMatrix(1, ETemp1.LocationName) = "Location Name"
   .TextMatrix(1, ETemp1.AMname) = "AM Name"
   .TextMatrix(1, ETemp1.RMName) = "RM Name"
   .TextMatrix(1, ETemp1.CMPName) = "CMP Name"
   .TextMatrix(1, ETemp1.GodownNo) = "Godown No"
   .TextMatrix(1, ETemp1.GodownCapacity) = "Godown Capacity"
   .TextMatrix(1, ETemp1.NcdexStock) = "NCDEX Stock "
   .TextMatrix(1, ETemp1.NonNCDEXStock) = "Non NCDEX Stock"
   .TextMatrix(1, ETemp1.NNNRIncome) = "NNNR Income "
   .TextMatrix(1, ETemp1.ReservationIncome) = "Reservation Income "
   .TextMatrix(1, ETemp1.SevenDaysIncome) = "Seven Days Income "
   .TextMatrix(1, ETemp1.FinalDPIncome) = "Final DP Income for the Godown" ''"DP Income "
   .TextMatrix(1, ETemp1.DPIncomeasperGSL) = "DP Income as per GSL" '"DP Invoice Amount From GSL Detail"
   .TextMatrix(1, ETemp1.Difference) = "Difference between DP (GSL - Invoice)" '"Difference Distribution"
   .TextMatrix(1, ETemp1.DPIncomeasperInvoice) = "DP Income as per DP Invoice" '"Amount After Distribution"
   .TextMatrix(1, ETemp1.RejectedIncome) = "Rejected Income " '" NCDEX WithDrawal Income "
   .TextMatrix(1, ETemp1.OtherIncome) = "Other Income "
   .TextMatrix(1, ETemp1.SCMIncome) = "SCM Income "
   .TextMatrix(1, ETemp1.LeftOverStock) = "Left over Stocks"
   .TextMatrix(1, ETemp1.PledgeIncome) = "Pledge Income"
   .TextMatrix(1, ETemp1.PartiallyReversedIncome) = "Partially Reversed Income"
   .TextMatrix(1, ETemp1.EmpClaimFAG) = "Employee Claims - approved by FAG"
   .TextMatrix(1, ETemp1.VendorExpenses) = "Vendor Expenses Paid"
   .TextMatrix(1, ETemp1.VendorExpensesAccrued) = "Vendor Expenses Accrued"
   .TextMatrix(1, ETemp1.SecurityFeesPaid) = "Security Fees - paid "
   .TextMatrix(1, ETemp1.FumigationFeesVendor) = "Fumigation Fees thru Vendor module - paid"
   .TextMatrix(1, ETemp1.LeaseRentPaid) = "Lease Rent Paid"
   .TextMatrix(1, ETemp1.LeaseRentFAG) = "Lease Rent - pending with FAG"
   .TextMatrix(1, ETemp1.LeaseRentDRC) = "Lease Rent - pending with DRC"
   .TextMatrix(1, ETemp1.FranchiseeFeesPaid) = "Franchisee Fees - Paid by FAG"
   .TextMatrix(1, ETemp1.FranchiseeFeesFAG) = "Franchisee Fees - Processed but not paid by FAG"
   .TextMatrix(1, ETemp1.FranchiseeFeesCMG) = "Franchisee Fees - Unprocessed by CMG"
   .TextMatrix(1, ETemp1.Salary) = "Salary Corporate"
   .TextMatrix(1, ETemp1.SalaryFieldStaff) = "Salary Field Staff"
   .TextMatrix(1, ETemp1.CoExpense) = "CO Expenses - Allocated"
   .TextMatrix(1, ETemp1.NCDEXCharges) = "NCDEX Charges - 2% of DP Fees"
   .TextMatrix(1, ETemp1.GodownStartDate) = "Godown Start Date"
   .TextMatrix(1, ETemp1.GodownCloseDate) = "Godown Close Date"
   .TextMatrix(1, ETemp1.UnitType) = "Unit Type"
   .TextMatrix(1, ETemp1.StateID) = "StateID"
   .TextMatrix(1, ETemp1.LocationID) = "LocationID"
   .TextMatrix(1, ETemp1.GodownID) = "GodownID"
   .TextMatrix(1, ETemp1.SecurityFeesAccrual) = "Security Fees - Accrued"
   .TextMatrix(1, ETemp1.FumigationFeesPaid) = "Fumigation & Spraying Fees - paid"
   .TextMatrix(1, ETemp1.FumigationFeesAccrual) = "Fumigation & Spraying Fees - accrued"
   .TextMatrix(1, ETemp1.InsurancePremium) = "Insurance Premium"
   .TextMatrix(1, ETemp1.InterestCharges) = "Interest Charges"
   .TextMatrix(1, ETemp1.EmpClaimAccrual) = "Employee Claim - accrued"
   .TextMatrix(1, ETemp1.CMUnbilledIncome) = "Income from CM Unbilled"
   .TextMatrix(1, ETemp1.CMPendingIncome) = "Income from CM Pending"
   .TextMatrix(1, ETemp1.CMApprovedIncome) = "Income from CM Approved"
   .TextMatrix(1, ETemp1.ReservationIncomeAdhoc) = "Reservation Income through Ad-Hoc"
   .TextMatrix(1, ETemp1.DPIncomeAdHoc) = "DP Income through Ad-Hoc"
   .TextMatrix(1, ETemp1.NNNRIncomeAdhoc) = "NNNR Income through Ad-Hoc"
   .TextMatrix(1, ETemp1.SpotIncomeAdHoc) = "Spot Income through Ad-Hoc"
   .TextMatrix(1, ETemp1.TotalIncome) = "Total Income"
   .TextMatrix(1, ETemp1.TotalExpense) = "Total Expense"
   .TextMatrix(1, ETemp1.GrossIncome) = "Gross Income"
   .TextMatrix(1, ETemp1.NetIncome) = "Net Income"
   
   .ColWidth(0) = 1500
   .ColWidth(1) = 1800
   .ColWidth(2) = 1800
   .ColWidth(3) = 2800
   .ColWidth(4) = 2800
   .ColWidth(5) = 1800
   .ColWidth(6) = 1000
   For j = 7 To 52
       .ColWidth(j) = 1000
   Next
   For j = 53 To 57
       .ColWidth(j) = 0
   Next
   .WordWrap = True
   .RowHeight(1) = 700
End With

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
Call CmdClear_Click

tmp = "Select Max(MonthYear) MonthYear From Txn_MonGodownWiseProfitabilityFreezed "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   If IsNull(TmpRs1!MonthYear) = False Then
      Label2.Caption = Format(TmpRs1!MonthYear, "MMM-yyyy")
   End If
End If

End Sub
Private Sub CmbMonth_GotFocus()
tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I)
Next
CmbMonth.Text = tmp
End Sub

Private Sub CmbMonth_LostFocus()
If Trim(CmbMonth.Text) = "" Then Exit Sub
mStatus = False
For I = 0 To 11
   If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
      CmbMonth.Text = CmbMonth.List(I)
      mStatus = True
      mMonth = I + 1
      Exit For
   End If
Next
If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbYear_GotFocus()
If CmbMonth.Text = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If
tmp = CmbYear.Text
CmbYear.Clear
For I = 2010 To Year(Now)
   CmbYear.AddItem I
Next

If Month(Now) = 12 Then
   CmbYear.AddItem Year(Now) + 1
End If

CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
If Trim(CmbYear.Text) = "" Then Exit Sub
Dim flg As Variant
Dim myr As Variant

flg = "N"

If Month(Now) = 12 Then
   myr = Year(Now) + 1
Else
   myr = Year(Now)
End If

For I = 2010 To myr 'Year(Now)
   If CmbYear.Text = I Then
      flg = "Y"
      Exit For
   End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbYear.SetFocus
   Exit Sub
End If

mYear = CmbYear.Text
End Sub

Private Function GetMGWPID() As Double

Dim Retval As Double
tmp = "Select max(MGWPID) from Txn_MonGodownWiseProfitabilityFreezed "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMGWPID = Retval

End Function

Private Sub Grid_Head2()
CmdExcel.Enabled = False
Label5.Caption = ""
ProgressBar1.Value = 0
With MSHFlexGrid2
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 61 ' 60 '59
   
   .TextMatrix(1, ETemp2.MonthYear) = "Month Year"
   .TextMatrix(1, ETemp2.StateName) = "State Name"
   .TextMatrix(1, ETemp2.LocationName) = "Location Name"
   .TextMatrix(1, ETemp2.AMname) = "AM Name"
   .TextMatrix(1, ETemp2.RMName) = "RM Name"
   .TextMatrix(1, ETemp2.CMPName) = "CMP Name"
   .TextMatrix(1, ETemp2.GodownNo) = "Godown No"
   .TextMatrix(1, ETemp2.GodownCapacity) = "Godown Capacity"
   .TextMatrix(1, ETemp2.NcdexStock) = "NCDEX Stock "
   .TextMatrix(1, ETemp2.NonNCDEXStock) = "Non NCDEX Stock"
   .TextMatrix(1, ETemp2.NNNRIncome) = "NNNR Income "
   .TextMatrix(1, ETemp2.ReservationIncome) = "Reservation Income "
   .TextMatrix(1, ETemp2.SevenDaysIncome) = "Seven Days Income "
   .TextMatrix(1, ETemp2.FinalDPIncome) = "Final DP Income for the Godown" ''"DP Income "
   .TextMatrix(1, ETemp2.DPIncomeasperGSL) = "DP Income as per GSL" '"DP Invoice Amount From GSL Detail"
   .TextMatrix(1, ETemp2.Difference) = "Difference between DP (GSL - Invoice)" '"Difference Distribution"
   .TextMatrix(1, ETemp2.DPIncomeasperInvoice) = "DP Income as per DP Invoice" '"Amount After Distribution"
   .TextMatrix(1, ETemp2.RejectedIncome) = "Rejected Income " '" NCDEX WithDrawal Income "
   .TextMatrix(1, ETemp2.OtherIncome) = "Other Income "
   .TextMatrix(1, ETemp2.SCMIncome) = "SCM Income "
   .TextMatrix(1, ETemp2.LeftOverStock) = "Left over Stocks"
   .TextMatrix(1, ETemp2.PledgeIncome) = "Pledge Income"
   .TextMatrix(1, ETemp2.PartiallyReversedIncome) = "Partially Reversed Income"
   .TextMatrix(1, ETemp2.EmpClaimFAG) = "Employee Claims - approved by FAG"
   .TextMatrix(1, ETemp2.VendorExpenses) = "Vendor Expenses Paid"
   .TextMatrix(1, ETemp2.VendorExpensesAccrued) = "Vendor Expenses Accrued"
   .TextMatrix(1, ETemp2.SecurityFeesPaid) = "Security Fees - paid "
   .TextMatrix(1, ETemp2.FumigationFeesVendor) = "Fumigation Fees thru Vendor module - paid"
   .TextMatrix(1, ETemp2.LeaseRentPaid) = "Lease Rent Paid"
   .TextMatrix(1, ETemp2.LeaseRentFAG) = "Lease Rent - pending with FAG"
   .TextMatrix(1, ETemp2.LeaseRentDRC) = "Lease Rent - pending with DRC"
   .TextMatrix(1, ETemp2.FranchiseeFeesPaid) = "Franchisee Fees - Paid by FAG"
   .TextMatrix(1, ETemp2.FranchiseeFeesFAG) = "Franchisee Fees - Processed but not paid by FAG"
   .TextMatrix(1, ETemp2.FranchiseeFeesCMG) = "Franchisee Fees - Unprocessed by CMG"
   .TextMatrix(1, ETemp2.Salary) = "Salary Corporate"
   .TextMatrix(1, ETemp2.SalaryFieldStaff) = "Salary Field Staff"
   .TextMatrix(1, ETemp2.CoExpense) = "CO Expenses - Allocated"
   .TextMatrix(1, ETemp2.NCDEXCharges) = "NCDEX Charges - 2% of DP Fees"
   .TextMatrix(1, ETemp2.GodownStartDate) = "Godown Start Date"
   .TextMatrix(1, ETemp2.GodownCloseDate) = "Godown Close Date"
   .TextMatrix(1, ETemp2.UnitType) = "Unit Type"
   .TextMatrix(1, ETemp2.StateID) = "StateID"
   .TextMatrix(1, ETemp2.LocationID) = "LocationID"
   .TextMatrix(1, ETemp2.GodownID) = "GodownID"
   .TextMatrix(1, ETemp2.SecurityFeesAccrual) = "Security Fees - Accrued"
   .TextMatrix(1, ETemp2.FumigationFeesPaid) = "Fumigation & Spraying Fees - paid"
   .TextMatrix(1, ETemp2.FumigationFeesAccrual) = "Fumigation & Spraying Fees - accrued"
   .TextMatrix(1, ETemp2.InsurancePremium) = "Insurance Premium"
   .TextMatrix(1, ETemp2.InterestCharges) = "Interest Charges"
   .TextMatrix(1, ETemp2.EmpClaimAccrual) = "Employee Claim - accrued"
   .TextMatrix(1, ETemp2.CMUnbilledIncome) = "Income from CM Unbilled"
   .TextMatrix(1, ETemp2.CMPendingIncome) = "Income from CM Pending"
   .TextMatrix(1, ETemp2.CMApprovedIncome) = "Income from CM Approved"
   .TextMatrix(1, ETemp2.ReservationIncomeAdhoc) = "Reservation Income through Ad-Hoc"
   .TextMatrix(1, ETemp2.DPIncomeAdHoc) = "DP Income through Ad-Hoc"
   .TextMatrix(1, ETemp2.NNNRIncomeAdhoc) = "NNNR Income through Ad-Hoc"
   .TextMatrix(1, ETemp2.SpotIncomeAdHoc) = "Spot Income through Ad-Hoc"
   .TextMatrix(1, ETemp2.TotalIncome) = "Total Income"
   .TextMatrix(1, ETemp2.TotalExpense) = "Total Expense"
   .TextMatrix(1, ETemp2.GrossIncome) = "Gross Income"
   .TextMatrix(1, ETemp2.NetIncome) = "Net Income"
   
   .ColWidth(0) = 1500
   .ColWidth(1) = 1800
   .ColWidth(2) = 1800
   .ColWidth(3) = 2800
   .ColWidth(4) = 2800
   .ColWidth(5) = 1800
   .ColWidth(6) = 1000
   For j = 7 To 52
       .ColWidth(j) = 1000
   Next
   For j = 53 To 57
       .ColWidth(j) = 0
   Next
   .WordWrap = True
   .RowHeight(1) = 700
End With

End Sub

