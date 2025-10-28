VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_NFinPledge 
   Caption         =   "N-Fin Pledge Creation And Pledge Release Report"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   12
      Top             =   8880
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   15000
      Begin VB.CheckBox ChkZeroQuantity 
         Caption         =   "With Zero Quantity"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1000
         Width           =   1935
      End
      Begin VB.ComboBox CmbStateID 
         Height          =   360
         Left            =   1740
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   570
         Width           =   4695
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   12720
         TabIndex        =   3
         Top             =   555
         Width           =   1095
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         Height          =   375
         Left            =   11505
         TabIndex        =   2
         Top             =   555
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   13815
         TabIndex        =   4
         Top             =   555
         Width           =   1095
      End
      Begin VB.ComboBox CmbLocation 
         Height          =   360
         Left            =   6540
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   570
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker TxtMonthYear 
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   570
         Width           =   1575
         _ExtentX        =   2778
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   109707267
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Month - Year"
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
         Left            =   247
         TabIndex        =   14
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
         Height          =   240
         Left            =   3795
         TabIndex        =   10
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
         Left            =   13440
         TabIndex        =   9
         Top             =   1005
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
         Left            =   11040
         TabIndex        =   8
         Top             =   1005
         Width           =   2325
      End
      Begin VB.Label Label3 
         Caption         =   "Location"
         Height          =   255
         Left            =   8520
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7515
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   5
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
Attribute VB_Name = "FrmRpt_NFinPledge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mMonthYear As Variant
Dim mStartDate As Variant, mEndDate As Variant
Dim mNFinFundingAmountPercentage As Double, mNFinFeesPA As Double

Private Sub CmdGetReport_Click()
Dim mWC As Variant
Dim mMonth As Variant, mMonthNofoDays As Variant
Dim mR As Double, mLastRow As Double

Call Grid_Head
Label5.Caption = ""

If CDate(Format(TxtMonthYear.Value, "yyyy-MM-01")) > CDate(Format(Date, "yyyy-MM-01")) Then
    MsgBox "Funture Month-Year not allowed."
    Exit Sub
End If
ProgressBar1.Value = 0
MsgBox "Wait till next message."

mMonth = Month(TxtMonthYear.Value)
mMonthYear = Format(TxtMonthYear.Value, "yyyy-mm-01")

mStartDate = Format(TxtMonthYear.Value, "yyyy-mm-01")
mEndDate = Format(CDate(Format(DateAdd("m", 1, TxtMonthYear.Value), "yyyy-mm-01")) - 1, "yyyy-mm-dd")
mMonthNofoDays = Day(CDate(mEndDate))

mWC = ""
If CmbStateID.Text <> "" Then
   mWC = " And L.StateID = " & Val(mStateID) & " "
End If
If CmbLocation.Text <> "" Then
    If mWC = "" Then
        mWC = " And L.LocationID = " & Val(mLocationID) & " "
    Else
        mWC = mWC & " And L.LocationID = " & Val(mLocationID) & " "
    End If
End If

tmp = " Select 'Pledge-O' 'Type',A.LocationName,A.GodownNo, A.CMPID , A.WHRNo,A.PledgeDate,'' As 'PledgeReleaseDate',SUM(A.NoofBags) 'NoofBags' ,SUM(A.Quantity) 'Quantity',Avg(A.Rate) 'Rate'"
tmp = tmp & " From ("
tmp = tmp & "    select 'Pledge-O' 'Type',l.LocationName , g.GodownNo , g.CMPID , a.WHRNo , PledgeDate,'' as 'PledgeReleaseDate',sum(b.PledgeNoOfBags) 'NoofBags' ,sum(b.PledgeQuantity) 'Quantity' , avg(w.rate) 'Rate'"
tmp = tmp & "    from txn_pledgeCreation a"
tmp = tmp & "    inner join txn_pledgeCreation_gsl b on a.Txn_PledgeCreation_ID = b.Txn_PledgeCreation_ID and a.sm_prefix = b.sm_prefix"
tmp = tmp & "    inner join txn_Whr_Detail w on a.WHRNo = w.WHRNo and a.sm_prefix = w.sm_prefix"
tmp = tmp & "    inner join mst_location l on a.sm_prefix = l.sm_prefix"
tmp = tmp & "    inner join mst_gsl gsl on b.sm_prefix = gsl.sm_prefix and b.Mst_GSL_ID = gsl.Mst_GSL_ID"
tmp = tmp & "    inner join mst_godown g on g.GodownID = gsl.GodownID"
tmp = tmp & "    where a.ncmlfinance = 1 and PledgeDate <= '" & mEndDate & "'"
tmp = tmp & "    " & mWC & ""
tmp = tmp & "    group by l.LocationName , g.GodownNo , g.CMPID , a.WHRNo , PledgeDate"
tmp = tmp & "    Union All"
tmp = tmp & "    select 'Release-O' 'Type',l.LocationName , g.GodownNo, g.CMPID , c.WHRNo , c.PledgeDate,a.PledgeReleaseDate ,sum(b.ReleasedNoOfBags) 'NoofBags' ,sum(b.ReleasedQuantity) 'Quantity' , avg(w.rate) 'Rate'"
tmp = tmp & "    from txn_pledgeRelease a"
tmp = tmp & "    inner join txn_pledgeRelease_gsl b on a.Txn_PledgeRelease_ID = b.Txn_PledgeRelease_ID and a.sm_prefix = b.sm_prefix"
tmp = tmp & "    inner join txn_pledgeCreation c on a.Txn_PledgeCreation_ID = c.Txn_PledgeCreation_ID and a.sm_prefix = c.sm_prefix"
tmp = tmp & "    inner join txn_Whr_Detail w on c.WHRNo = w.WHRNo and c.sm_prefix = w.sm_prefix"
tmp = tmp & "    inner join mst_location l on a.sm_prefix = l.sm_prefix"
tmp = tmp & "    inner join mst_gsl gsl on b.sm_prefix = gsl.sm_prefix and b.Mst_GSL_ID = gsl.Mst_GSL_ID"
tmp = tmp & "    inner join mst_godown g on g.GodownID = gsl.GodownID"
tmp = tmp & "    where c.ncmlfinance = 1 and c.PledgeDate <= '" & mEndDate & "' and a.PledgeReleaseDate > '" & mEndDate & "'"
tmp = tmp & "    " & mWC & ""
tmp = tmp & "    group by l.LocationName , g.GodownNo, g.CMPID , c.WHRNo , c.PledgeDate , a.PledgeReleaseDate"
tmp = tmp & " ) A"
tmp = tmp & " Group by A.LocationName,A.GodownNo,A.CMPID,A.WHRNo,A.PledgeDate"
tmp = tmp & " Union All"
tmp = tmp & " select 'Release-C' 'Type',l.LocationName , g.GodownNo,g.CMPID , c.WHRNo , c.PledgeDate, a.PledgeReleaseDate ,sum(b.ReleasedNoOfBags) 'NoofBags' ,sum(b.ReleasedQuantity) 'Quantity' , avg(w.rate) 'Rate'"
tmp = tmp & " from txn_pledgeRelease a"
tmp = tmp & " inner join txn_pledgeRelease_gsl b on a.Txn_PledgeRelease_ID = b.Txn_PledgeRelease_ID and a.sm_prefix = b.sm_prefix"
tmp = tmp & " inner join txn_pledgeCreation c on a.Txn_PledgeCreation_ID = c.Txn_PledgeCreation_ID and a.sm_prefix = c.sm_prefix"
tmp = tmp & " inner join txn_Whr_Detail w on c.WHRNo = w.WHRNo and c.sm_prefix = w.sm_prefix"
tmp = tmp & " inner join mst_location l on a.sm_prefix = l.sm_prefix"
tmp = tmp & " inner join mst_gsl gsl on b.sm_prefix = gsl.sm_prefix and b.Mst_GSL_ID = gsl.Mst_GSL_ID"
tmp = tmp & " inner join mst_godown g on g.GodownID = gsl.GodownID"
tmp = tmp & " where c.ncmlfinance = 1 and a.PledgeReleaseDate >= '" & mStartDate & "' and a.PledgeReleaseDate <= '" & mEndDate & "'"
tmp = tmp & "    " & mWC & ""
tmp = tmp & " group by l.LocationName , g.GodownNo, g.CMPID , c.WHRNo , c.PledgeDate , a.PledgeReleaseDate"
tmp = tmp & " Order By 'Type',A.LocationName,A.PledgeDate "
Call TmpTable
''--'NoofDays','AUM','FundingAmount,'Fees','Forday','Actualfees'
If TmpRs.RecordCount > 0 Then
    ProgressBar1.Max = TmpRs.RecordCount
    With MSHFlexGrid1
        For mR = 1 To TmpRs.RecordCount
            
          If ChkZeroQuantity.Value = vbUnchecked Then
            If IIf(IsNull(TmpRs!NoofBags), 0, TmpRs!NoofBags) = 0 Then GoTo NextMove
          End If
          
            mLastRow = .Rows - 1
            .TextMatrix(mLastRow, 0) = IIf(IsNull(TmpRs!Type), "", TmpRs!Type)
            .TextMatrix(mLastRow, 1) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
            .TextMatrix(mLastRow, 2) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
            .TextMatrix(mLastRow, 3) = TmpRs!CMPID & "-" & ((TmpRs!WHRNo) - TmpRs!CMPID * 1000000) '' IIf(IsNull(TmpRs!WHRNo), "", TmpRs!WHRNo)
            .TextMatrix(mLastRow, 4) = IIf(IsNull(TmpRs!PledgeDate), "", Format(TmpRs!PledgeDate, "DD-MMM-YYYY"))
            .TextMatrix(mLastRow, 5) = IIf(TmpRs!PledgeReleaseDate = "1/1/1900", "", Format(TmpRs!PledgeReleaseDate, "DD-MMM-YYYY"))
            If LCase(.TextMatrix(mLastRow, 0)) = LCase("Release-C") Then
                   .TextMatrix(mLastRow, 6) = IIf(Format(TmpRs!PledgeDate, "YYYY-MM-01") = Format(TmpRs!PledgeReleaseDate, "YYYY-MM-01"), DateDiff("d", TmpRs!PledgeDate, TmpRs!PledgeReleaseDate) + 1, Day(TmpRs!PledgeReleaseDate))
            Else
                   .TextMatrix(mLastRow, 6) = IIf(Format(TmpRs!PledgeDate, "YYYY-MM-01") = mMonthYear, mMonthNofoDays - Day(TmpRs!PledgeDate) + 1, mMonthNofoDays)
            End If
            .TextMatrix(mLastRow, 7) = IIf(IsNull(TmpRs!NoofBags), 0, TmpRs!NoofBags)
            .TextMatrix(mLastRow, 8) = IIf(IsNull(TmpRs!Quantity), 0, TmpRs!Quantity)
            .TextMatrix(mLastRow, 9) = Format(IIf(IsNull(TmpRs!Rate), 0, TmpRs!Rate), "#0.00")
            .TextMatrix(mLastRow, 10) = Format(.TextMatrix(mLastRow, 8) * .TextMatrix(mLastRow, 9), "#0.00")
            .TextMatrix(mLastRow, 11) = Format(.TextMatrix(mLastRow, 10) * (mNFinFundingAmountPercentage / 100), "#0.00")
            .TextMatrix(mLastRow, 12) = Format(.TextMatrix(mLastRow, 11) * (mNFinFeesPA / 100), "#0.00")
            .TextMatrix(mLastRow, 13) = Format(.TextMatrix(mLastRow, 12) / 365, "#0.00")
            .TextMatrix(mLastRow, 14) = Format(.TextMatrix(mLastRow, 13) * .TextMatrix(mLastRow, 6), "#0.00")
            
            ''-- Total on Top
            .TextMatrix(0, 14) = Format(Val(.TextMatrix(0, 14)) + Val(.TextMatrix(mLastRow, 14)), "#0.00")
            .Rows = .Rows + 1
          
NextMove:
          TmpRs.MoveNext
          ProgressBar1.Value = mR
        
        Next mR
    End With
End If

Label5.Caption = MSHFlexGrid1.Rows - 3
MsgBox "Done."
CmdExcel.Enabled = True

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
mNFinFundingAmountPercentage = Gen_NFinPledgeCalculationFactor("NFinFundingAmountPercentage-")
mNFinFeesPA = Gen_NFinPledgeCalculationFactor("NFINFeesPA-")
CmdClear_Click
CmdExcel.Enabled = False
End Sub

Private Sub CmdClear_Click()
Label5.Caption = ""
TxtMonthYear.Value = Date - 30
CmdExcel.Enabled = False
CmbLocation.Text = ""
CmbStateID.Text = ""
ChkZeroQuantity.Value = vbUnchecked
Call Grid_Head
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NFinPledgeCreationAndPledgeReleaseReport.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NFinPledgeCreationAndPledgeReleaseReport.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
       If C = 2 Then
            oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
       Else
            oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
       End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
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
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.Clear
   CmbStateID.SetFocus
   Exit Sub
End If
CmbLocation.Clear
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
    CmbLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   mLocationID = 0
   Exit Sub
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

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Cols = 15
    .Rows = 3
    .FixedRows = 2
    .FixedCols = 0
    .WordWrap = True
    
    .TextMatrix(1, 0) = "Type":                             .ColWidth(0) = 1200
    .TextMatrix(1, 1) = "Location Name":                    .ColWidth(1) = 1500
    .TextMatrix(1, 2) = "Godown No":                        .ColWidth(2) = 1200
    .TextMatrix(1, 3) = "WHR No":                           .ColWidth(3) = 1500
    .TextMatrix(1, 4) = "Pledge Date":                      .ColWidth(4) = 1500
    .TextMatrix(1, 5) = "Pledge Release Date":              .ColWidth(5) = 1500
    .TextMatrix(1, 6) = "No of days":                       .ColWidth(6) = 1200
    .TextMatrix(1, 7) = "No of Bags":                       .ColWidth(7) = 1200
    .TextMatrix(1, 8) = "Quantity":                         .ColWidth(8) = 1200
    .TextMatrix(1, 9) = "Rate":                             .ColWidth(9) = 1500
    .TextMatrix(1, 10) = "AUM":                             .ColWidth(10) = 1800
    .TextMatrix(1, 11) = "Funding Amount":                  .ColWidth(11) = 1800
    .TextMatrix(1, 12) = "Fees " & mNFinFeesPA & " % PA":   .ColWidth(12) = 1800
    .TextMatrix(1, 13) = "For day":                         .ColWidth(13) = 1200
    .TextMatrix(1, 14) = "Actual Fees":                     .ColWidth(14) = 2000
    .RowHeight(1) = 800
    .RowHeight(0) = 400
    .Refresh
End With
End Sub
