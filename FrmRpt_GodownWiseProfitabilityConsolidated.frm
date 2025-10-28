VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_GodownWiseProfitabilityConsolidated 
   Caption         =   "Godown Wise Profitability Consolidated"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20220
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      Height          =   1605
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15000
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   5895
         Begin VB.OptionButton OptConsolidated 
            Caption         =   "Consolidated Data"
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
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton OptMonthwise 
            Caption         =   "Month wise Detail"
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
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
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
         TabIndex        =   15
         Top             =   555
         Width           =   2445
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
         Left            =   7350
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   555
         Width           =   2640
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
         Left            =   12750
         TabIndex        =   6
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
         Left            =   10320
         TabIndex        =   5
         Top             =   600
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
         Left            =   11535
         TabIndex        =   4
         Top             =   600
         Width           =   1215
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
         Left            =   3120
         TabIndex        =   2
         Top             =   195
         Value           =   -1  'True
         Width           =   1695
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
         Left            =   3120
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1620
         TabIndex        =   3
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   49086467
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   165
         TabIndex        =   7
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   49086467
         CurrentDate     =   39850
         MinDate         =   38353
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
         TabIndex        =   17
         Top             =   240
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
         Left            =   8295
         TabIndex        =   16
         Top             =   240
         Width           =   765
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
         Left            =   12960
         TabIndex        =   11
         Top             =   195
         Width           =   90
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
         Left            =   10560
         TabIndex        =   10
         Top             =   195
         Width           =   2325
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   1920
         TabIndex        =   9
         Top             =   195
         Width           =   735
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
         TabIndex        =   8
         Top             =   195
         Width           =   945
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   9390
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7575
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13361
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
Attribute VB_Name = "FrmRpt_GodownWiseProfitabilityConsolidated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mStateID As Variant 'mTotalDays, mTotalMonth
Dim mStartDate, mEndDate, mMonth, mYear, mTotalDays, wc As Variant
Enum ETemp
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
   LeaseRentAccrued = 33
   EmpClaimFAG = 34
   EmpClaimAccrual = 35
   VendorExpenses = 36
   VendorExpensesAccrued = 37
   SecurityFeesPaid = 38
   SecurityFeesAccrual = 39
   FumigationFeesVendor = 40
   FumigationFeesPaid = 41
   FumigationFeesAccrual = 42
   FranchiseeFeesPaid = 43
   FranchiseeFeesFAG = 44
   FranchiseeFeesCMG = 45
   Salary = 46
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

''================Listbox Item Sequence
'SelectAll
'StateName
'LocationName
'AM Name
'RM Name
'CMP Name
'Godown No
'Godown Address
'Unit Type
'Godown Start Date
'Godown Close Date
'Capacity
'NCDEX Stock
'Non NCDEX Stock
'NNNR Income through Ad-Hoc
'NNNR Income
'Reservation Income through Ad-Hoc
'Reservation Income
'Seven Days Income
'DP Income through Ad-Hoc
'Final DP Income for the Godown
'Rejected Income
'Spot Income through Ad-Hoc
'Other Income
'SCM Income
'Income from CM Unbilled
'Income from CM Pending
'Income from CM Approved
'Left over Stocks
'Pledge Income
'Lease Rent Paid
'Lease Rent - pending with FAG
'Lease Rent - pending with DRC
'Employee Claims - approved by FAG
'Employee Claim - Accrued
'Vendor Expenses
'Security Fees - Paid
'Security Fees - Accrued
'Fumigation Fees thru Vendor module - paid
'Fumigation & Spraying Fees - paid
'Fumigation & Spraying Fees - accrued
'Franchisee Fees - Paid
'Salary
'NCDEX Charges - 2% of DP Fees
'CO Expenses - Allocated
'Insurance Premium
'Interest Charges


Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState = "" Then
   Call TableMst_State
Else
   Call TableMst_State(" Where " & Gen_WcState & "")
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
   MsgBox "Please Select State Name !!!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text

If Gen_WcLocation <> "" Then
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation & " ")
Else
   Call TableMst_Location("Where StateID=" & mStateID & "")
End If
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
TxtFromDate.Value = Date
TxtToDate.Value = Date
Call TxtFromDate_LostFocus
Call TxtToDate_LostFocus
CmbStateID.Text = ""
CmbLocationID.Text = ""
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWiseProfitability.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWiseProfitability.xls")
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
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Call Grid_Head

If Format(TxtFromDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future Date not allowed!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If
If Format(TxtToDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future Date not allowed!!!"
   TxtToDate.SetFocus
   Exit Sub
End If
If Format(TxtFromDate.Value, "yyyy-mm") > Format(TxtToDate.Value, "yyyy-mm") Then
   MsgBox "From Date should not greater than To Date!!!"
   TxtToDatEtemp.SetFocus
   Exit Sub
End If

wc = ""

If CmbStateID.Text <> "" Then
   wc = " Where MS.StateID = " & mStateID
Else
   If Gen_WcState <> "" Then
      wc = " Where ML." & Gen_WcState
   End If
End If

If CmbLocationID.Text <> "" Then
   wc = " Where ML.LocationID = " & mLocationID
Else
    If Gen_WcLocation <> "" Then
       wc = wc & " And ML." & Gen_WcLocation
    End If
End If
If Month(TxtFromDate.Value) = Month(TxtToDate.Value) Then
   mYearMonth = Format(TxtFromDate.Value, "MMM-yyyy")
Else
   mYearMonth = Format(TxtFromDate.Value, "MMM-yyyy") & " To " & Format(TxtToDate.Value, "MMM-yyyy")
End If
mStartDate = Format(TxtFromDate.Value, "yyyy-mm-dd")
mEndDate = Format(TxtToDate.Value, "yyyy-mm-dd")

tmp = " Select"
If OptConsolidated.Value = True Then
   tmp = tmp & " MS.StateName, ML.LocationName, TMGP.AM, TMGP.RM, MCMP.CMPName, MG.GodownNo, MCMP.CMPName+'_'+MG.GodownNo, MUT.UnitType, MG.StartDate, MG.CloseDate,"
Else
   tmp = tmp & " TMGP.MonthYear,MS.StateName, ML.LocationName, TMGP.AM, TMGP.RM, MCMP.CMPName, MG.GodownNo, MCMP.CMPName+'_'+MG.GodownNo, MUT.UnitType, MG.StartDate, MG.CloseDate,"
End If
'tmp = tmp & " avg(MG.GodownCapacity) Capacity,  avg(TMGP.NcdexStock) NcdexStock, avg(TMGP.NNStock) NNStock, sum(TMGP.NNNRIncomeAdhoc) NNNRIncomeAdhoc, sum(TMGP.NNNRIncome) NNNRIncome,"
tmp = tmp & " Avg(MG.GodownCapacity) Capacity,  sum(TMGP.NcdexStock) NcdexStock, sum(TMGP.NNStock) NNStock, sum(TMGP.NNNRIncomeAdhoc) NNNRIncomeAdhoc, sum(TMGP.NNNRIncome) NNNRIncome,"
tmp = tmp & " sum(TMGP.ReservationIncomeAdhoc) ReservationIncomeAdhoc, sum(TMGP.ReservationIncome) ReservationIncome, sum(TMGP.SevenDaysIncome) SevenDaysIncome, sum(TMGP.DPIncomeAdHoc) DPIncomeAdHoc,"
tmp = tmp & " sum(TMGP.FinalDPIncome) FinalDPIncome,  sum(TMGP.RejectedIncome) RejectedIncome,"
tmp = tmp & " sum(TMGP.SpotIncomeAdHoc) SpotIncomeAdHoc, sum(TMGP.OtherIncome) OtherIncome, sum(TMGP.SCMIncome) SCMIncome, sum(TMGP.CMUnbilledIncome) CMUnbilledIncome,"
tmp = tmp & " sum(TMGP.CMPendingIncome) CMPendingIncome, sum(TMGP.CMApprovedIncome) CMApprovedIncome, sum(TMGP.LeftOverStock) LeftOverStock, sum(TMGP.PledgeIncome) PledgeIncome, "
tmp = tmp & " sum(TMGP.LeaseRentPaid) LeaseRentPaid, sum(TMGP.LeaseRentFAG) LeaseRentFAG, sum(TMGP.LeaseRentDRC) LeaseRentDRC,sum(TMGP.LeaseRentAccrued) LeaseRentAccrued,  sum(TMGP.EmpClaimFAG) EmpClaimFAG, sum(TMGP.EmpClaimAccrual) EmpClaimAccrual,"
tmp = tmp & " sum(TMGP.VendorExpenses) VendorExpenses, sum(TMGP.VendorExpensesAccrued) VendorExpensesAccrued, sum(TMGP.SecurityFeesPaid) SecurityFeesPaid, sum(TMGP.SecurityFeesAccrual) SecurityFeesAccrual, sum(TMGP.FumigationFeesVendor) FumigationFeesVendor,"
tmp = tmp & " sum(TMGP.FumigationFeesPaid) FumigationFeesPaid, sum(TMGP.FumigationFeesAccrual) FumigationFeesAccrual, sum(TMGP.FranchiseeFeesPaid) FranchiseeFeesPaid, sum(TMGP.FranchiseeFeesFAG) FranchiseeFeesFAG, sum(TMGP.FranchiseeFeesCMG) FranchiseeFeesCMG, (Sum(TMGP.Salary) + Sum(TMGP.SalaryFieldStaff)) Salary , "
tmp = tmp & " sum(TMGP.NCDEXCharges) NCDEXCharges, sum(TMGP.CoExpense) CoExpense, sum(TMGP.InsurancePremium) InsurancePremium, sum(TMGP.InterestCharges) InterestCharges,sum(TMGP.PartiallyReversedIncome) PartiallyReversedIncome,"
tmp = tmp & " TMGP.StateID , TMGP.LocationID, TMGP.GodownID, Sum(TMGP.DPIncomeGSL) DPIncomeGSL, Sum(TMGP.Diff) Diff, Sum(TMGP.DPIncomeInvoice) DPIncomeInvoice"
tmp = tmp & " From Txn_MonGodownWiseProfitability TMGP"
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
If OptConsolidated.Value = True Then
   tmp = tmp & " group by MS.StateName,"
Else
   tmp = tmp & " group by TMGP.MonthYear,MS.StateName,"
End If
tmp = tmp & " ML.LocationName, TMGP.AM, TMGP.RM, MCMP.CMPName, MG.GodownNo, MUT.UnitType, MG.StartDate, MG.CloseDate,"
tmp = tmp & " TMGP.StateID , TMGP.LocationID, TMGP.GodownID"
If OptConsolidated.Value = True Then
   tmp = tmp & " Order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
Else
   tmp = tmp & " Order by TMGP.MonthYear,MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
End If
Call TmpTable

If TmpRs.RecordCount > 0 Then
   Label5.Caption = TmpRs.RecordCount
   ProgressBar1.Value = 0
   ProgressBar1.Max = MSHFlexGrid1.Rows + TmpRs.RecordCount
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   CmdExcel.Enabled = True
    
   For I = 2 To TmpRs.RecordCount + 1
      With MSHFlexGrid1
            If OptConsolidated.Value = True Then
               .TextMatrix(I, ETemp.MonthYear) = mYearMonth
            Else
               .TextMatrix(I, ETemp.MonthYear) = IIf(IsNull(TmpRs!MonthYear), "", Format(TmpRs!MonthYear, "MMM-YYYY"))
            End If
            
            .TextMatrix(I, ETemp.StateName) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
            .TextMatrix(I, ETemp.LocationName) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
            .TextMatrix(I, ETemp.AMname) = IIf(IsNull(TmpRs!AM), "", TmpRs!AM)
            .TextMatrix(I, ETemp.RMName) = IIf(IsNull(TmpRs!RM), "", TmpRs!RM)
            .TextMatrix(I, ETemp.CMPName) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
            .TextMatrix(I, ETemp.GodownNo) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
            .TextMatrix(I, ETemp.GodownCapacity) = IIf(IsNull(TmpRs!Capacity), "", TmpRs!Capacity)
            .TextMatrix(I, ETemp.NcdexStock) = IIf(IsNull(TmpRs!NcdexStock), "", TmpRs!NcdexStock)
            .TextMatrix(I, ETemp.NonNCDEXStock) = IIf(IsNull(TmpRs!NNStock), "", TmpRs!NNStock)
            .TextMatrix(I, ETemp.NNNRIncomeAdhoc) = IIf(IsNull(TmpRs!NNNRIncomeAdhoc), "", TmpRs!NNNRIncomeAdhoc)
            .TextMatrix(I, ETemp.NNNRIncome) = IIf(IsNull(TmpRs!NNNRIncome), "", TmpRs!NNNRIncome)
            .TextMatrix(I, ETemp.ReservationIncomeAdhoc) = IIf(IsNull(TmpRs!ReservationIncomeAdhoc), "", TmpRs!ReservationIncomeAdhoc)
            .TextMatrix(I, ETemp.ReservationIncome) = IIf(IsNull(TmpRs!ReservationIncome), "", TmpRs!ReservationIncome)
            .TextMatrix(I, ETemp.SevenDaysIncome) = IIf(IsNull(TmpRs!SevenDaysIncome), "", TmpRs!SevenDaysIncome)
            .TextMatrix(I, ETemp.DPIncomeAdHoc) = IIf(IsNull(TmpRs!DPIncomeAdHoc), "", TmpRs!DPIncomeAdHoc)
            .TextMatrix(I, ETemp.FinalDPIncome) = IIf(IsNull(TmpRs!FinalDPIncome), "", TmpRs!FinalDPIncome)
            .TextMatrix(I, ETemp.DPIncomeasperGSL) = IIf(IsNull(TmpRs!DPIncomeGSL), "", TmpRs!DPIncomeGSL)
            .TextMatrix(I, ETemp.Difference) = IIf(IsNull(TmpRs!Diff), "", TmpRs!Diff)
            .TextMatrix(I, ETemp.DPIncomeasperInvoice) = IIf(IsNull(TmpRs!DPIncomeInvoice), "", TmpRs!DPIncomeInvoice)
            .TextMatrix(I, ETemp.RejectedIncome) = IIf(IsNull(TmpRs!RejectedIncome), "", TmpRs!RejectedIncome)
            .TextMatrix(I, ETemp.SpotIncomeAdHoc) = IIf(IsNull(TmpRs!SpotIncomeAdHoc), "", TmpRs!SpotIncomeAdHoc)
            .TextMatrix(I, ETemp.OtherIncome) = IIf(IsNull(TmpRs!OtherIncome), "", TmpRs!OtherIncome)
            .TextMatrix(I, ETemp.SCMIncome) = IIf(IsNull(TmpRs!SCMIncome), "", TmpRs!SCMIncome)
            .TextMatrix(I, ETemp.CMUnbilledIncome) = IIf(IsNull(TmpRs!CMUnbilledIncome), "", TmpRs!CMUnbilledIncome)
            .TextMatrix(I, ETemp.CMPendingIncome) = IIf(IsNull(TmpRs!CMPendingIncome), "", TmpRs!CMPendingIncome)
            .TextMatrix(I, ETemp.CMApprovedIncome) = IIf(IsNull(TmpRs!CMApprovedIncome), "", TmpRs!CMApprovedIncome)
            .TextMatrix(I, ETemp.LeftOverStock) = IIf(IsNull(TmpRs!LeftOverStock), "", TmpRs!LeftOverStock)
            .TextMatrix(I, ETemp.PledgeIncome) = IIf(IsNull(TmpRs!PledgeIncome), "", TmpRs!PledgeIncome)
            .TextMatrix(I, ETemp.PartiallyReversedIncome) = IIf(IsNull(TmpRs!PartiallyReversedIncome), "", TmpRs!PartiallyReversedIncome)
            .TextMatrix(I, ETemp.LeaseRentPaid) = IIf(IsNull(TmpRs!LeaseRentPaid), "", TmpRs!LeaseRentPaid)
            .TextMatrix(I, ETemp.LeaseRentFAG) = IIf(IsNull(TmpRs!LeaseRentFAG), "", TmpRs!LeaseRentFAG)
            .TextMatrix(I, ETemp.LeaseRentDRC) = IIf(IsNull(TmpRs!LeaseRentDRC), "", TmpRs!LeaseRentDRC)
            .TextMatrix(I, ETemp.LeaseRentAccrued) = IIf(IsNull(TmpRs!LeaseRentAccrued), "", TmpRs!LeaseRentAccrued)  'LeaseRentAccrued
            .TextMatrix(I, ETemp.EmpClaimFAG) = IIf(IsNull(TmpRs!EmpClaimFAG), "", TmpRs!EmpClaimFAG)
            .TextMatrix(I, ETemp.EmpClaimAccrual) = IIf(IsNull(TmpRs!EmpClaimAccrual), "", TmpRs!EmpClaimAccrual)
            .TextMatrix(I, ETemp.VendorExpenses) = IIf(IsNull(TmpRs!VendorExpenses), "", TmpRs!VendorExpenses)
            .TextMatrix(I, ETemp.VendorExpensesAccrued) = IIf(IsNull(TmpRs!VendorExpensesAccrued), "", TmpRs!VendorExpensesAccrued)
            .TextMatrix(I, ETemp.SecurityFeesPaid) = IIf(IsNull(TmpRs!SecurityFeesPaid), "", TmpRs!SecurityFeesPaid)
            .TextMatrix(I, ETemp.SecurityFeesAccrual) = IIf(IsNull(TmpRs!SecurityFeesAccrual), "", TmpRs!SecurityFeesAccrual)
            .TextMatrix(I, ETemp.FumigationFeesVendor) = IIf(IsNull(TmpRs!FumigationFeesVendor), "", TmpRs!FumigationFeesVendor)
            .TextMatrix(I, ETemp.FumigationFeesPaid) = IIf(IsNull(TmpRs!FumigationFeesPaid), "", TmpRs!FumigationFeesPaid)
            .TextMatrix(I, ETemp.FumigationFeesAccrual) = IIf(IsNull(TmpRs!FumigationFeesAccrual), "", TmpRs!FumigationFeesAccrual)
            .TextMatrix(I, ETemp.FranchiseeFeesPaid) = IIf(IsNull(TmpRs!FranchiseeFeesPaid), "", TmpRs!FranchiseeFeesPaid)
            .TextMatrix(I, ETemp.FranchiseeFeesCMG) = IIf(IsNull(TmpRs!FranchiseeFeesCMG), "", TmpRs!FranchiseeFeesCMG)
            .TextMatrix(I, ETemp.FranchiseeFeesFAG) = IIf(IsNull(TmpRs!FranchiseeFeesFAG), "", TmpRs!FranchiseeFeesFAG)
            .TextMatrix(I, ETemp.Salary) = IIf(IsNull(TmpRs!Salary), "", TmpRs!Salary)
            .TextMatrix(I, ETemp.NCDEXCharges) = IIf(IsNull(TmpRs!NCDEXCharges), "", TmpRs!NCDEXCharges)
            .TextMatrix(I, ETemp.CoExpense) = IIf(IsNull(TmpRs!CoExpense), "", TmpRs!CoExpense)
            .TextMatrix(I, ETemp.InsurancePremium) = IIf(IsNull(TmpRs!InsurancePremium), "", TmpRs!InsurancePremium)
            .TextMatrix(I, ETemp.InterestCharges) = IIf(IsNull(TmpRs!InterestCharges), "", TmpRs!InterestCharges)
            .TextMatrix(I, ETemp.GodownStartDate) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
            .TextMatrix(I, ETemp.GodownCloseDate) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
            .TextMatrix(I, ETemp.UnitType) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
            .TextMatrix(I, ETemp.StateID) = IIf(IsNull(TmpRs!StateID), "", TmpRs!StateID)
            .TextMatrix(I, ETemp.LocationID) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
            .TextMatrix(I, ETemp.GodownID) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
            
            .TextMatrix(0, ETemp.NcdexStock) = Val(.TextMatrix(I, ETemp.NcdexStock)) + Val(.TextMatrix(0, ETemp.NcdexStock))
            .TextMatrix(0, ETemp.GodownCapacity) = Val(.TextMatrix(I, ETemp.GodownCapacity)) + Val(.TextMatrix(0, ETemp.GodownCapacity))
            .TextMatrix(0, ETemp.NonNCDEXStock) = Val(.TextMatrix(I, ETemp.NonNCDEXStock)) + Val(.TextMatrix(0, ETemp.NonNCDEXStock))
            .TextMatrix(0, ETemp.NNNRIncomeAdhoc) = Val(.TextMatrix(I, ETemp.NNNRIncomeAdhoc)) + Val(.TextMatrix(0, ETemp.NNNRIncomeAdhoc))
            .TextMatrix(0, ETemp.NNNRIncome) = Val(.TextMatrix(I, ETemp.NNNRIncome)) + Val(.TextMatrix(0, ETemp.NNNRIncome))
            .TextMatrix(0, ETemp.ReservationIncomeAdhoc) = Val(.TextMatrix(I, ETemp.ReservationIncomeAdhoc)) + Val(.TextMatrix(0, ETemp.ReservationIncomeAdhoc))
            .TextMatrix(0, ETemp.ReservationIncome) = Val(.TextMatrix(I, ETemp.ReservationIncome)) + Val(.TextMatrix(0, ETemp.ReservationIncome))
            .TextMatrix(0, ETemp.SevenDaysIncome) = Val(.TextMatrix(I, ETemp.SevenDaysIncome)) + Val(.TextMatrix(0, ETemp.SevenDaysIncome))
            .TextMatrix(0, ETemp.DPIncomeAdHoc) = Val(.TextMatrix(I, ETemp.DPIncomeAdHoc)) + Val(.TextMatrix(0, ETemp.DPIncomeAdHoc))
            .TextMatrix(0, ETemp.FinalDPIncome) = Val(.TextMatrix(I, ETemp.FinalDPIncome)) + Val(.TextMatrix(0, ETemp.FinalDPIncome))
            .TextMatrix(0, ETemp.DPIncomeasperGSL) = Val(.TextMatrix(I, ETemp.DPIncomeasperGSL)) + Val(.TextMatrix(0, ETemp.DPIncomeasperGSL))
            .TextMatrix(0, ETemp.Difference) = Val(.TextMatrix(I, ETemp.Difference)) + Val(.TextMatrix(0, ETemp.Difference))
            .TextMatrix(0, ETemp.DPIncomeasperInvoice) = Val(.TextMatrix(I, ETemp.DPIncomeasperInvoice)) + Val(.TextMatrix(0, ETemp.DPIncomeasperInvoice))
            .TextMatrix(0, ETemp.RejectedIncome) = Val(.TextMatrix(I, ETemp.RejectedIncome)) + Val(.TextMatrix(0, ETemp.RejectedIncome))
            .TextMatrix(0, ETemp.SpotIncomeAdHoc) = Val(.TextMatrix(I, ETemp.SpotIncomeAdHoc)) + Val(.TextMatrix(0, ETemp.SpotIncomeAdHoc))
            .TextMatrix(0, ETemp.OtherIncome) = Val(.TextMatrix(I, ETemp.OtherIncome)) + Val(.TextMatrix(0, ETemp.OtherIncome))
            .TextMatrix(0, ETemp.SCMIncome) = Val(.TextMatrix(I, ETemp.SCMIncome)) + Val(.TextMatrix(0, ETemp.SCMIncome))
            .TextMatrix(0, ETemp.CMUnbilledIncome) = Val(.TextMatrix(I, ETemp.CMUnbilledIncome)) + Val(.TextMatrix(0, ETemp.CMUnbilledIncome))
            .TextMatrix(0, ETemp.CMPendingIncome) = Val(.TextMatrix(I, ETemp.CMPendingIncome)) + Val(.TextMatrix(0, ETemp.CMPendingIncome))
            .TextMatrix(0, ETemp.CMApprovedIncome) = Val(.TextMatrix(I, ETemp.CMApprovedIncome)) + Val(.TextMatrix(0, ETemp.CMApprovedIncome))
            .TextMatrix(0, ETemp.LeftOverStock) = Val(.TextMatrix(I, ETemp.LeftOverStock)) + Val(.TextMatrix(0, ETemp.LeftOverStock))
            .TextMatrix(0, ETemp.PledgeIncome) = Val(.TextMatrix(I, ETemp.PledgeIncome)) + Val(.TextMatrix(0, ETemp.PledgeIncome))
            .TextMatrix(0, ETemp.PartiallyReversedIncome) = Val(.TextMatrix(I, ETemp.PartiallyReversedIncome)) + Val(.TextMatrix(0, ETemp.PartiallyReversedIncome))
            .TextMatrix(0, ETemp.LeaseRentPaid) = Val(.TextMatrix(I, ETemp.LeaseRentPaid)) + Val(.TextMatrix(0, ETemp.LeaseRentPaid))
            .TextMatrix(0, ETemp.LeaseRentFAG) = Val(.TextMatrix(I, ETemp.LeaseRentFAG)) + Val(.TextMatrix(0, ETemp.LeaseRentFAG))
            .TextMatrix(0, ETemp.LeaseRentDRC) = Val(.TextMatrix(I, ETemp.LeaseRentDRC)) + Val(.TextMatrix(0, ETemp.LeaseRentDRC))
            .TextMatrix(0, ETemp.LeaseRentAccrued) = Val(.TextMatrix(I, ETemp.LeaseRentAccrued)) + Val(.TextMatrix(0, ETemp.LeaseRentAccrued))
            .TextMatrix(0, ETemp.EmpClaimFAG) = Val(.TextMatrix(I, ETemp.EmpClaimFAG)) + Val(.TextMatrix(0, ETemp.EmpClaimFAG))
            .TextMatrix(0, ETemp.EmpClaimAccrual) = Val(.TextMatrix(I, ETemp.EmpClaimAccrual)) + Val(.TextMatrix(0, ETemp.EmpClaimAccrual))
            .TextMatrix(0, ETemp.VendorExpenses) = Val(.TextMatrix(I, ETemp.VendorExpenses)) + Val(.TextMatrix(0, ETemp.VendorExpenses))
            .TextMatrix(0, ETemp.VendorExpensesAccrued) = Val(.TextMatrix(I, ETemp.VendorExpensesAccrued)) + Val(.TextMatrix(0, ETemp.VendorExpensesAccrued))
            .TextMatrix(0, ETemp.SecurityFeesPaid) = Val(.TextMatrix(I, ETemp.SecurityFeesPaid)) + Val(.TextMatrix(0, ETemp.SecurityFeesPaid))
            .TextMatrix(0, ETemp.SecurityFeesAccrual) = Val(.TextMatrix(I, ETemp.SecurityFeesAccrual)) + Val(.TextMatrix(0, ETemp.SecurityFeesAccrual))
            .TextMatrix(0, ETemp.FumigationFeesVendor) = Val(.TextMatrix(I, ETemp.FumigationFeesVendor)) + Val(.TextMatrix(0, ETemp.FumigationFeesVendor))
            .TextMatrix(0, ETemp.FumigationFeesPaid) = Val(.TextMatrix(I, ETemp.FumigationFeesPaid)) + Val(.TextMatrix(0, ETemp.FumigationFeesPaid))
            .TextMatrix(0, ETemp.FumigationFeesAccrual) = Val(.TextMatrix(I, ETemp.FumigationFeesAccrual)) + Val(.TextMatrix(0, ETemp.FumigationFeesAccrual))
            .TextMatrix(0, ETemp.FranchiseeFeesPaid) = Val(.TextMatrix(I, ETemp.FranchiseeFeesPaid)) + Val(.TextMatrix(0, ETemp.FranchiseeFeesPaid))
            .TextMatrix(0, ETemp.FranchiseeFeesCMG) = Val(.TextMatrix(I, ETemp.FranchiseeFeesCMG)) + Val(.TextMatrix(0, ETemp.FranchiseeFeesCMG))
            .TextMatrix(0, ETemp.FranchiseeFeesFAG) = Val(.TextMatrix(I, ETemp.FranchiseeFeesFAG)) + Val(.TextMatrix(0, ETemp.FranchiseeFeesFAG))
            .TextMatrix(0, ETemp.Salary) = Val(.TextMatrix(I, ETemp.Salary)) + Val(.TextMatrix(0, ETemp.Salary))
            .TextMatrix(0, ETemp.NCDEXCharges) = Val(.TextMatrix(I, ETemp.NCDEXCharges)) + Val(.TextMatrix(0, ETemp.NCDEXCharges))
            .TextMatrix(0, ETemp.CoExpense) = Val(.TextMatrix(I, ETemp.CoExpense)) + Val(.TextMatrix(0, ETemp.CoExpense))
            .TextMatrix(0, ETemp.InsurancePremium) = Val(.TextMatrix(I, ETemp.InsurancePremium)) + Val(.TextMatrix(0, ETemp.InsurancePremium))
            .TextMatrix(0, ETemp.InterestCharges) = Val(.TextMatrix(I, ETemp.InterestCharges)) + Val(.TextMatrix(0, ETemp.InterestCharges))
      End With
      TmpRs.MoveNext
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next I
End If
With MSHFlexGrid1
   For I = 2 To MSHFlexGrid1.Rows - 2
      .TextMatrix(I, ETemp.TotalIncome) = (Val(.TextMatrix(I, ETemp.NNNRIncomeAdhoc)) + Val(.TextMatrix(I, ETemp.NNNRIncome)) + Val(.TextMatrix(I, ETemp.ReservationIncomeAdhoc)) + Val(.TextMatrix(I, ETemp.ReservationIncome)) + Val(.TextMatrix(I, ETemp.SevenDaysIncome)) + Val(.TextMatrix(I, ETemp.DPIncomeAdHoc)) + Val(.TextMatrix(I, ETemp.FinalDPIncome)) + Val(.TextMatrix(I, ETemp.RejectedIncome)) + Val(.TextMatrix(I, ETemp.SpotIncomeAdHoc)) + Val(.TextMatrix(I, ETemp.OtherIncome)) + Val(.TextMatrix(I, ETemp.SCMIncome)) + Val(.TextMatrix(I, ETemp.CMUnbilledIncome)) + Val(.TextMatrix(I, ETemp.CMPendingIncome)) + Val(.TextMatrix(I, ETemp.CMApprovedIncome)) + Val(.TextMatrix(I, ETemp.LeftOverStock)) + Val(.TextMatrix(I, ETemp.PledgeIncome))) - Val(.TextMatrix(I, ETemp.PartiallyReversedIncome))
      .TextMatrix(I, ETemp.TotalIncome) = Format(.TextMatrix(I, ETemp.TotalIncome), "############.00")
      .TextMatrix(I, ETemp.TotalExpense) = Val(.TextMatrix(I, ETemp.LeaseRentPaid)) + Val(.TextMatrix(I, ETemp.LeaseRentFAG)) + Val(.TextMatrix(I, ETemp.LeaseRentDRC)) + Val(.TextMatrix(I, ETemp.EmpClaimFAG)) + Val(.TextMatrix(I, ETemp.EmpClaimAccrual)) + Val(.TextMatrix(I, ETemp.VendorExpenses)) + Val(.TextMatrix(I, ETemp.VendorExpensesAccrued)) + Val(.TextMatrix(I, ETemp.SecurityFeesPaid)) + Val(.TextMatrix(I, ETemp.SecurityFeesAccrual)) + Val(.TextMatrix(I, ETemp.FumigationFeesVendor)) + Val(.TextMatrix(I, ETemp.FumigationFeesPaid)) + Val(.TextMatrix(I, ETemp.FumigationFeesAccrual)) + Val(.TextMatrix(I, ETemp.FranchiseeFeesPaid)) + Val(.TextMatrix(I, ETemp.FranchiseeFeesCMG)) + Val(.TextMatrix(I, ETemp.FranchiseeFeesFAG)) + Val(.TextMatrix(I, ETemp.Salary)) + Val(.TextMatrix(I, ETemp.NCDEXCharges)) + Val(.TextMatrix(I, ETemp.LeaseRentAccrued))
      .TextMatrix(I, ETemp.TotalExpense) = Format(.TextMatrix(I, ETemp.TotalExpense), "############.00")
      .TextMatrix(I, ETemp.GrossIncome) = Val(.TextMatrix(I, ETemp.TotalIncome)) - Val(.TextMatrix(I, ETemp.TotalExpense))
      .TextMatrix(I, ETemp.GrossIncome) = Format(.TextMatrix(I, ETemp.GrossIncome), "############.00")
      .TextMatrix(I, ETemp.NetIncome) = Val(.TextMatrix(I, ETemp.GrossIncome)) - Val(.TextMatrix(I, ETemp.CoExpense)) - Val(.TextMatrix(I, ETemp.InsurancePremium)) - Val(.TextMatrix(I, ETemp.InterestCharges))
      .TextMatrix(I, ETemp.NetIncome) = Format(.TextMatrix(I, ETemp.NetIncome), "############.00")
      .TextMatrix(0, ETemp.TotalIncome) = Format(Val(.TextMatrix(0, ETemp.TotalIncome)) + Val(.TextMatrix(I, ETemp.TotalIncome)), "############.00")
      .TextMatrix(0, ETemp.TotalExpense) = Format(Val(.TextMatrix(0, ETemp.TotalExpense)) + Val(.TextMatrix(I, ETemp.TotalExpense)), "############.00")
      .TextMatrix(0, ETemp.GrossIncome) = Format(Val(.TextMatrix(0, ETemp.GrossIncome)) + Val(.TextMatrix(I, ETemp.GrossIncome)), "############.00")
      .TextMatrix(0, ETemp.NetIncome) = Format(Val(.TextMatrix(0, ETemp.NetIncome)) + Val(.TextMatrix(I, ETemp.NetIncome)), "############.00")
   Next
End With
ProgressBar1.Value = ProgressBar1.Max
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
   .Cols = 61 '60 '59
   
   .TextMatrix(1, ETemp.MonthYear) = "Month Year"
   .TextMatrix(1, ETemp.StateName) = "State Name"
   .TextMatrix(1, ETemp.LocationName) = "Location Name"
   .TextMatrix(1, ETemp.AMname) = "AM Name"
   .TextMatrix(1, ETemp.RMName) = "RM Name"
   .TextMatrix(1, ETemp.CMPName) = "CMP Name"
   .TextMatrix(1, ETemp.GodownNo) = "Godown No"
   .TextMatrix(1, ETemp.GodownCapacity) = "Godown Capacity"
   .TextMatrix(1, ETemp.NcdexStock) = "NCDEX Stock "
   .TextMatrix(1, ETemp.NonNCDEXStock) = "Non NCDEX Stock"
   .TextMatrix(1, ETemp.NNNRIncome) = "NNNR Income "
   .TextMatrix(1, ETemp.ReservationIncome) = "Reservation Income "
   .TextMatrix(1, ETemp.SevenDaysIncome) = "Seven Days Income "
   .TextMatrix(1, ETemp.FinalDPIncome) = "Final DP Income for the Godown" ''"DP Income "
   .TextMatrix(1, ETemp.DPIncomeasperGSL) = "DP Income as per GSL" '"DP Invoice Amount From GSL Detail"
   .TextMatrix(1, ETemp.Difference) = "Difference between DP (GSL - Invoice)" '"Difference Distribution"
   .TextMatrix(1, ETemp.DPIncomeasperInvoice) = "DP Income as per DP Invoice" '"Amount After Distribution"
   .TextMatrix(1, ETemp.RejectedIncome) = "Rejected Income " '" NCDEX WithDrawal Income "
   .TextMatrix(1, ETemp.OtherIncome) = "Other Income "
   .TextMatrix(1, ETemp.SCMIncome) = "SCM Income "
   .TextMatrix(1, ETemp.LeftOverStock) = "Left over Stocks"
   .TextMatrix(1, ETemp.PledgeIncome) = "Pledge Income"
   .TextMatrix(1, ETemp.PartiallyReversedIncome) = "Partially Reversed Income"
   .TextMatrix(1, ETemp.EmpClaimFAG) = "Employee Claims - approved by FAG"
   .TextMatrix(1, ETemp.VendorExpenses) = "Vendor Expenses Paid"
   .TextMatrix(1, ETemp.VendorExpensesAccrued) = "Vendor Expenses Accrued"
   .TextMatrix(1, ETemp.SecurityFeesPaid) = "Security Fees - paid "
   .TextMatrix(1, ETemp.FumigationFeesVendor) = "Fumigation Fees thru Vendor module - paid"
   .TextMatrix(1, ETemp.LeaseRentPaid) = "Lease Rent Paid"
   .TextMatrix(1, ETemp.LeaseRentFAG) = "Lease Rent - pending with FAG"
   .TextMatrix(1, ETemp.LeaseRentDRC) = "Lease Rent - pending with DRC"
   .TextMatrix(1, ETemp.LeaseRentAccrued) = "Lease Rent - Accrued"
   .TextMatrix(1, ETemp.FranchiseeFeesPaid) = "Franchisee Fees - Paid by FAG"
   .TextMatrix(1, ETemp.FranchiseeFeesFAG) = "Franchisee Fees - Processed but not paid by FAG"
   .TextMatrix(1, ETemp.FranchiseeFeesCMG) = "Franchisee Fees - Unprocessed by CMG"
   .TextMatrix(1, ETemp.Salary) = "Salary"
   .TextMatrix(1, ETemp.CoExpense) = "CO Expenses - Allocated"
   .TextMatrix(1, ETemp.NCDEXCharges) = "NCDEX Charges - 2% of DP Fees"
   .TextMatrix(1, ETemp.GodownStartDate) = "Godown Start Date"
   .TextMatrix(1, ETemp.GodownCloseDate) = "Godown Close Date"
   .TextMatrix(1, ETemp.UnitType) = "Unit Type"
   .TextMatrix(1, ETemp.StateID) = "StateID"
   .TextMatrix(1, ETemp.LocationID) = "LocationID"
   .TextMatrix(1, ETemp.GodownID) = "GodownID"
   .TextMatrix(1, ETemp.SecurityFeesAccrual) = "Security Fees - Accrued"
   .TextMatrix(1, ETemp.FumigationFeesPaid) = "Fumigation & Spraying Fees - paid"
   .TextMatrix(1, ETemp.FumigationFeesAccrual) = "Fumigation & Spraying Fees - accrued"
   .TextMatrix(1, ETemp.InsurancePremium) = "Insurance Premium"
   .TextMatrix(1, ETemp.InterestCharges) = "Interest Charges"
   .TextMatrix(1, ETemp.EmpClaimAccrual) = "Employee Claim - accrued"
   .TextMatrix(1, ETemp.CMUnbilledIncome) = "Income from CM Unbilled"
   .TextMatrix(1, ETemp.CMPendingIncome) = "Income from CM Pending"
   .TextMatrix(1, ETemp.CMApprovedIncome) = "Income from CM Approved"
   .TextMatrix(1, ETemp.ReservationIncomeAdhoc) = "Reservation Income through Ad-Hoc"
   .TextMatrix(1, ETemp.DPIncomeAdHoc) = "DP Income through Ad-Hoc"
   .TextMatrix(1, ETemp.NNNRIncomeAdhoc) = "NNNR Income through Ad-Hoc"
   .TextMatrix(1, ETemp.SpotIncomeAdHoc) = "Spot Income through Ad-Hoc / RRN Withdrawal Income"
   .TextMatrix(1, ETemp.TotalIncome) = "Total Income"
   .TextMatrix(1, ETemp.TotalExpense) = "Total Expense"
   .TextMatrix(1, ETemp.GrossIncome) = "Gross Income"
   .TextMatrix(1, ETemp.NetIncome) = "Net Income"
   
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

End Sub

Private Sub TxtFromDate_LostFocus()
TxtFromDate.Day = 1
End Sub

Private Sub TxtToDate_LostFocus()
TxtToDate.Day = Day(DateSerial(TxtToDate.Year, TxtToDate.Month + 1, 0))
End Sub
