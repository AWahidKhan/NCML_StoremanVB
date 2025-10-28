VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_MonthwiseSecAttendanceProcess 
   Caption         =   "Month Wise Security Attendance Process"
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
      Height          =   9780
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   14970
      Begin VB.CheckBox ChkAllLocation 
         Appearance      =   0  'Flat
         Caption         =   "Show All Location"
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
         Height          =   240
         Left            =   1680
         TabIndex        =   19
         Top             =   847
         Width           =   2175
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
         ItemData        =   "Frm_MonthwiseSecAttendanceProcess.frx":0000
         Left            =   3360
         List            =   "Frm_MonthwiseSecAttendanceProcess.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   375
         Width           =   3960
      End
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
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
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   847
         Width           =   1215
      End
      Begin VB.ComboBox CmbSecurityAgency 
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
         Left            =   7395
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   375
         Width           =   7380
      End
      Begin VB.CommandButton CmdGo 
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
         Left            =   9120
         TabIndex        =   5
         Top             =   780
         Width           =   1320
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
         TabIndex        =   1
         Top             =   375
         Width           =   1575
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
         Left            =   1755
         TabIndex        =   2
         Top             =   375
         Width           =   1500
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   13680
         TabIndex        =   9
         Top             =   780
         Width           =   1080
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
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
         Left            =   10440
         TabIndex        =   6
         Top             =   780
         Width           =   1080
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
         Left            =   11520
         TabIndex        =   7
         Top             =   780
         Width           =   1080
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
         Left            =   12600
         TabIndex        =   8
         Top             =   780
         Width           =   1080
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8055
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1200
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   14208
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   9240
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label22 
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
         Left            =   6840
         TabIndex        =   18
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Total Records Found :"
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
         Left            =   4440
         TabIndex        =   17
         Top             =   840
         Width           =   2325
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Security Agency"
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
         Left            =   10358
         TabIndex        =   16
         Top             =   135
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Location"
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
         Left            =   4950
         TabIndex        =   15
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label Label2 
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
         Left            =   592
         TabIndex        =   13
         Top             =   135
         Width           =   630
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
         Left            =   2250
         TabIndex        =   12
         Top             =   135
         Width           =   510
      End
   End
End
Attribute VB_Name = "Frm_MonthwiseSecAttendanceProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mFromDate, mToDate, mMonth, mTotalDays As Variant
Dim mLocationID, mSecurityAgencyID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mSelRow, mSelCol As Variant

Private Sub ChkCheckAll_Click()
If tmp = 0 Then Exit Sub
If MSHFlexGrid1.Rows = 2 Then ChkCheckAll.Value = vbUnchecked: Exit Sub
For I = 1 To MSHFlexGrid1.Rows - 2
    If ChkCheckAll.Value = vbChecked Then
        MSHFlexGrid1.TextMatrix(I, 0) = strChecked
    Else
        MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
    End If
Next
End Sub

Private Sub CmbLocation_GotFocus()
If CmbMonth.Text = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox "Select Year!!!"
   CmbYear.SetFocus
   Exit Sub
End If

tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   If ChkAllLocation.Value = 0 Then
      Call TableMst_Location(" Where LocationID not in (Select Distinct LocationID from Txn_MonthwiseSecLoc Where MonthYear = '" & CmbYear.Text & "-" & mMonth & "-" & "01" & "' ) ")
   Else
      Call TableMst_Location(" ")
   End If
Else
   Call TableMst_Location(" WHERE " & Gen_WcLocation & " And LocationID not in (Select Distinct LocationID from Txn_MonthwiseSecLoc Where MonthYear = '" & CmbYear.Text & "-" & mMonth & "-" & "01" & "' ) ")
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
    MsgBox "Invalid Selection!!!"
    CmbLocation.SetFocus
    Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbMonth_GotFocus()

tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I), I - 1
Next

CmbMonth.Text = tmp

End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
       CmbMonth.Text = CmbMonth.List(I)
       mMonth = I + 1
       mStatus = True
       Exit For
    End If
Next

If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbMonth.Text <> tmp Then
   CmbLocation.Text = ""
   Call Grid_Head
End If

End Sub

Private Sub CmbSecurityAgency_GotFocus()
tmp = CmbSecurityAgency.Text
Call TableMst_SecurityAgency("Where Flag = 'Y' And AgencyID <> 80 ")
CmbSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbSecurityAgency.Text = tmp
End Sub

Private Sub CmbSecurityAgency_LostFocus()
If CmbSecurityAgency.Text = "" Then Exit Sub

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
mSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID
End Sub

Private Sub CmbYear_GotFocus()
If CmbMonth.Text = "" Then
    MsgBox "Select Month first!!!"
    CmbMonth.SetFocus
    Exit Sub
End If
tmp = CmbYear.Text
CmbYear.Clear
For I = 2010 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()

If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
flg = "N"

For I = 2010 To Year(Now)
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
If CmbYear.Text <> tmp Then
   CmbLocation.Text = ""
   Call Grid_Head
End If




End Sub

Private Sub CmdClear_Click()
CmbMonth.Text = ""
CmbYear.Text = ""
CmbLocation.Text = ""
CmbSecurityAgency.Text = ""
Call Grid_Head
ChkAllLocation.Value = 0
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_MonthwiseSecAttendanceProcess" & CmbMonth.Text & "-" & CmbYear.Text & ".xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_MonthwiseSecAttendanceProcess" & CmbMonth.Text & "-" & CmbYear.Text & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid1.Cols - 1
       oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
If CmbMonth.Text = "" Then
    MsgBox "Blank Month not Allowed!!!"
    CmbMonth.SetFocus
    Exit Sub
End If
If CmbYear.Text = "" Then
    MsgBox "Blank Year not Allowed!!!"
    CmbYear.SetFocus
    Exit Sub
End If

Call Grid_Head

MsgBox "Wait till next Message!!!"

If CmbYear.Text = "" Then
   MsgBox "Blank Year not Allowed!!!"
   CmbYear.SetFocus
   Exit Sub
End If


mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mToDate = Format(DateSerial(CmbYear.Text, mMonth + 1, 0), "yyyy-mm-dd")
mTotalDays = Day(mToDate)

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(mFromDate, "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbLocation.Text = "" Then
    MsgBox "Blank Location not Allowed!!!"
    CmbLocation.SetFocus
    Exit Sub
End If

If ChkAttendancenotMarked = False Then
   MsgBox " Mark/Absense Attendance of All Security Guard under this Location!!!"
   Exit Sub
End If

wc = ""
If CmbLocation.Text <> "" Then
    wc = wc & " And MSLM.LocationID = " & mLocationID
Else
    If Gen_WcLocation <> "" Then
        wc = wc & " And MSLM." & Gen_WcLocation
    End If
End If
If CmbSecurityAgency.Text <> "" Then
    wc = wc & " And MSLM.SecurityAgencyID = " & mSecurityAgencyID
End If


Dim mChooseSecurityID As Variant

mChooseSecurityID = ""

'If Gen_DBType = "MDB" Then
'   tmp = " Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = #" & mFromDate & "# "
'Else
'   tmp = " Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' "
'End If


If Gen_DBType = "MDB" Then
   tmp = " select isnull(msld.SecLocMapDetailID,0) 'SecLocMapDetailID' from Txn_MonthwiseSecLocDetail msld"
   tmp = tmp & " inner join Txn_MonthwiseSecLoc msl on msl.monthwisetxnid = msld.monthwisetxnid"
   tmp = tmp & " where monthyear = #" & mFromDate & "# "
Else
   tmp = " select isnull(msld.SecLocMapDetailID,0) 'SecLocMapDetailID' from Txn_MonthwiseSecLocDetail msld"
   tmp = tmp & " inner join Txn_MonthwiseSecLoc msl on msl.monthwisetxnid = msld.monthwisetxnid"
   tmp = tmp & " where monthyear = '" & mFromDate & "' "
End If

Call TmpTable

If TmpRs.RecordCount = 0 Then
   mChooseSecurityID = "(0)"
End If

For I = 1 To TmpRs.RecordCount
    If mChooseSecurityID = "" Then
       mChooseSecurityID = "(" & "'" & TmpRs!SecLocMapDetailID & "'"
    Else
       mChooseSecurityID = mChooseSecurityID & "," & "'" & TmpRs!SecLocMapDetailID & "'"
    End If
    
    If I = TmpRs.RecordCount Then
       mChooseSecurityID = mChooseSecurityID & ")"
    End If
    TmpRs.MoveNext
Next



'Process to Update RatePerMonth
tmp = "select "
tmp = tmp & " mslm.securityagencyid, mslm.locationid, ms.stateid,  mslmd.txndetailid, mslmd.securitytypeid,"
tmp = tmp & " mslmd.securitydesignationid, mslmd.startdate, mslmd.enddate, mslmd.noofgodown, mslmd.ratepermonth,"
tmp = tmp & " mslmd.SalaryPaidUpto, count(attendanceid) noofdays, ml.locationname, msa.agencyname, mss.securityshifts,  msd.securitydesignation"
tmp = tmp & " from txn_securityattendance tsa"
tmp = tmp & " inner join mst_seclocmappingdetail mslmd on mslmd.txndetailid = tsa.txndetailid"
tmp = tmp & " inner join mst_seclocmapping mslm on mslmd.txnid = mslm.txnid"
tmp = tmp & " inner join mst_location ml on ml.locationid = mslm.locationid"
tmp = tmp & " inner join mst_state ms on ms.stateid = ml.stateid"
tmp = tmp & " inner join mst_securityagency msa on msa.SecurityAgencyID = mslm.securityagencyid"
tmp = tmp & " inner join mst_securityshifts mss on mss.SecurityShiftsID = mslmd.SecurityTypeID"
tmp = tmp & " inner join mst_securitydesignation msd on msd.securitydesignationid = mslmd.securitydesignationid"
tmp = tmp & " where tsa.status  = 'A' and tsa.attendancedate between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & wc
'tmp = tmp & " and mslmd.txndetailid not in (select isnull(msld.SecLocMapDetailID,0) from Txn_MonthwiseSecLocDetail msld"
'tmp = tmp & " inner join Txn_MonthwiseSecLoc msl on msl.monthwisetxnid = msld.monthwisetxnid"
'tmp = tmp & " where monthyear = '" & mFromDate & "')"
tmp = tmp & " and mslmd.txndetailid not in " & mChooseSecurityID & " "
tmp = tmp & " group by mslm.securityagencyid, mslm.locationid, ms.stateid,mslmd.txndetailid, mslmd.securitytypeid,"
tmp = tmp & " mslmd.securitydesignationid, mslmd.startdate, mslmd.enddate, mslmd.noofgodown, mslmd.ratepermonth,"
tmp = tmp & " mslmd.SalaryPaidUpto, ml.locationname, msa.agencyname, mss.securityshifts,  msd.securitydesignation "
tmp = tmp & " order by mslm.locationid, mslm.securityagencyid"

Call TmpTable


tmp = "Select Txn_SA.StateID, Txn_SA.SecurityAgencyID, Txn_AD.SecurityTypeID, Txn_AD.SecurityDesignationID, Txn_AD.MonthlyRate,Txn_SA.AgreeExpiryDate,Txn_SA.AgreeStartDate"
tmp = tmp & " From Txn_SecurityAgreement as Txn_SA"
tmp = tmp & " Inner Join Txn_SecurityAgreementDetail as Txn_AD on Txn_AD.SAID = Txn_SA.SAID"
tmp = tmp & " Where Txn_SA.Flag = 'A' And Txn_SA.AgreeStartDate <= '" & mToDate & "' and Txn_SA.AgreeExpiryDate >= '" & mFromDate & "'"
tmp = tmp & " Order by Txn_SA.StateID,Txn_SA.SecurityAgencyID "

Call Tmp1Table
'mRate = ""
Call TableMst_SecLocMappingDetails

If TmpRs.RecordCount > 0 Then
    For I = 1 To TmpRs.RecordCount
        If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
        TmpRs1.Filter = "StateID=" & TmpRs!StateID & " And SecurityAgencyID = " & TmpRs!SecurityAgencyID & " And SecurityTypeID = " & TmpRs!SecurityTypeID & " And SecurityDesignationID= " & TmpRs!SecurityDesignationID & " And AgreeExpiryDate >= '" & TmpRs!SalaryPaidUpto & "' And AgreeStartDate <= '" & TmpRs!SalaryPaidUpto & "'"
        If TmpRs1.RecordCount > 0 Then
           mRate = TmpRs1!MonthlyRate
           RsMst_SecLocMappingDetails.MoveFirst
           RsMst_SecLocMappingDetails.Find "txndetailid =" & TmpRs!TxnDetailID
           If RsMst_SecLocMappingDetails.EOF = False Then
              If mRate <> RsMst_SecLocMappingDetails!RatePerMonth Then
                 RsMst_SecLocMappingDetails!OldRatePerMonth = IIf(IsNull(RsMst_SecLocMappingDetails!OldRatePerMonth), RsMst_SecLocMappingDetails!RatePerMonth, RsMst_SecLocMappingDetails!OldRatePerMonth & "~" & RsMst_SecLocMappingDetails!RatePerMonth)
                 RsMst_SecLocMappingDetails!RatePerMonth = mRate
                 RsMst_SecLocMappingDetails.Update
                 RsMst_SecLocMappingDetails.Requery
              End If
           End If
        End If
        TmpRs1.Filter = ""
        TmpRs.MoveNext
    Next
End If


tmp = "select "
tmp = tmp & " mslm.securityagencyid, mslm.locationid, mslmd.txndetailid, mslmd.securitytypeid,"
tmp = tmp & " mslmd.securitydesignationid, mslmd.startdate, mslmd.enddate, mslmd.noofgodown, mslmd.ratepermonth,"
tmp = tmp & " mslmd.SalaryPaidUpto, count(attendanceid) noofdays, ml.locationname, msa.agencyname, mss.securityshifts,  msd.securitydesignation"
tmp = tmp & " from txn_securityattendance tsa"
tmp = tmp & " inner join mst_seclocmappingdetail mslmd on mslmd.txndetailid = tsa.txndetailid"
tmp = tmp & " inner join mst_seclocmapping mslm on mslmd.txnid = mslm.txnid"
tmp = tmp & " inner join mst_location ml on ml.locationid = mslm.locationid"
tmp = tmp & " inner join mst_securityagency msa on msa.SecurityAgencyID = mslm.securityagencyid"
tmp = tmp & " inner join mst_securityshifts mss on mss.SecurityShiftsID = mslmd.SecurityTypeID"
tmp = tmp & " inner join mst_securitydesignation msd on msd.securitydesignationid = mslmd.securitydesignationid"
tmp = tmp & " where mslm.SecurityAgencyID <> 0 And tsa.status  = 'A' and tsa.attendancedate between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & wc
tmp = tmp & " and mslmd.txndetailid not in (select isnull(msld.SecLocMapDetailID,0) from Txn_MonthwiseSecLocDetail msld"
tmp = tmp & " inner join Txn_MonthwiseSecLoc msl on msl.monthwisetxnid = msld.monthwisetxnid"
tmp = tmp & " where msl.monthyear = '" & mFromDate & "' and msl.locationid = " & mLocationID & " )"
tmp = tmp & " group by mslm.securityagencyid, mslm.locationid, mslmd.txndetailid, mslmd.securitytypeid,"
tmp = tmp & " mslmd.securitydesignationid, mslmd.startdate, mslmd.enddate, mslmd.noofgodown, mslmd.ratepermonth,"
tmp = tmp & " mslmd.SalaryPaidUpto, ml.locationname, msa.agencyname, mss.securityshifts,  msd.securitydesignation "
tmp = tmp & " order by mslm.locationid, mslm.securityagencyid"

Call TmpTable


tmp = " Select DISTINCT ME.EmployeeID, ME.EmployeeName, ME.EmployeeNo, MECM.LocationID"
tmp = tmp & " From Mst_Employee ME"
tmp = tmp & " Inner Join Mst_EmpCMPMapping MECM On MECM.EmployeeID = ME.EmployeeID"
tmp = tmp & " Order by MECM.LocationID, ME.EmployeeID Desc"


Call Tmp1Table

If TmpRs.RecordCount > 0 Then
    Label22.Caption = TmpRs.RecordCount
    With MSHFlexGrid1
        .Rows = .Rows + TmpRs.RecordCount
        ProgressBar1.Max = .Rows
        For I = 1 To TmpRs.RecordCount
            .TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) '"Location Name"
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName) '"Security Agency"
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts) '"Security Type"
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation) '"Security Designation"
            .TextMatrix(I, 5) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy")) '"Start Date"
            .TextMatrix(I, 6) = IIf(IsNull(TmpRs!EndDate), "", Format(TmpRs!EndDate, "dd-MMM-yyyy")) '"End Date"
            .TextMatrix(I, 7) = IIf(IsNull(TmpRs!NoofGodown), "", TmpRs!NoofGodown) '"No of Godown"
            .TextMatrix(I, 8) = IIf(IsNull(TmpRs!RatePerMonth), "", TmpRs!RatePerMonth) '"Rate per month"
            If TmpRs1.RecordCount > 0 Then
                TmpRs1.MoveFirst
                TmpRs1.Find "LocationID = " & IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
                If Not TmpRs1.EOF Then
                    .TextMatrix(I, 9) = IIf(IsNull(TmpRs1!EmployeeName), "", TmpRs1!EmployeeName & "~" & TmpRs1!EmployeeNo) '"CMPM Name"
                    .TextMatrix(I, 18) = IIf(IsNull(TmpRs1!EmployeeID), "", TmpRs1!EmployeeID) '"CMPMID"
                End If
            End If
            .TextMatrix(I, 10) = IIf(IsNull(TmpRs!NoOfDays), "", TmpRs!NoOfDays) '"No of Days"
            .TextMatrix(I, 11) = Format((Val(.TextMatrix(I, 10)) * Val(.TextMatrix(I, 8)) / Val(mTotalDays)), "#########.00") '"Esitmated Amount"
            .TextMatrix(I, 12) = 0 '"Esitmated S.Tax"
            .TextMatrix(I, 13) = .TextMatrix(I, 11) '"Total Est. Amount"
            .TextMatrix(I, 14) = "" ''"Remarks"
            .TextMatrix(I, 15) = IIf(IsNull(TmpRs!SalaryPaidUpto), "", Format(TmpRs!SalaryPaidUpto, "dd-MMM-yyyy")) '"Salary Payable Upto"
            .TextMatrix(I, 16) = IIf(IsNull(TmpRs!SecurityTypeID), "", TmpRs!SecurityTypeID) '"Security TypeID"
            .TextMatrix(I, 17) = IIf(IsNull(TmpRs!SecurityDesignationID), "", TmpRs!SecurityDesignationID) '"Security DesignationID"
            .TextMatrix(I, 19) = IIf(IsNull(TmpRs!TxnDetailID), "", TmpRs!TxnDetailID) '"TxnDetailID"
            .TextMatrix(I, 20) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID) '"locationid"
            .TextMatrix(I, 21) = IIf(IsNull(TmpRs!SecurityAgencyID), "", TmpRs!SecurityAgencyID) '"Security Agency ID"
            .row = I
            .Col = 0
            .CellFontName = "Wingdings"
            .CellFontSize = 10
            .CellAlignment = flexAlignCenterCenter
            .TextMatrix(I, 0) = strUnChecked
            ProgressBar1.Value = ProgressBar1.Value + 1
            TmpRs.MoveNext
        Next
    End With
    ProgressBar1.Value = ProgressBar1.Max
    CmdSave.Enabled = True
    CmdExcel.Enabled = True
    MsgBox "Done!!!"
Else
    MsgBox "No Record Found!!!"
End If

End Sub

Private Sub CmdSave_Click()
ProgressBar1.Value = 0
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
        tmp = "select monthwisetxnid from Txn_MonthwiseSecLoc where monthyear = '" & mFromDate & "' and locationid = " & Val(MSHFlexGrid1.TextMatrix(I, 20)) & " and securityagencyid = " & Val(MSHFlexGrid1.TextMatrix(I, 21))
        Call Tmp1Table
        If TmpRs1.RecordCount > 0 Then
            mMaxTxnID = TmpRs1!MonthWiseTxnID
        Else
            RsTxn_MonthwiseSecLoc.AddNew
            RsTxn_MonthwiseSecLoc!MonthWiseTxnID = GetMaxTxnID
            mMaxTxnID = RsTxn_MonthwiseSecLoc!MonthWiseTxnID
            RsTxn_MonthwiseSecLoc!G_UID = GetGUID
            RsTxn_MonthwiseSecLoc!MonthwiseTxnDate = Date
            RsTxn_MonthwiseSecLoc!MonthYear = mFromDate
            RsTxn_MonthwiseSecLoc!LocationID = Val(MSHFlexGrid1.TextMatrix(I, 20))
            RsTxn_MonthwiseSecLoc!SecurityAgencyID = Val(MSHFlexGrid1.TextMatrix(I, 21))
            RsTxn_MonthwiseSecLoc!CreatedBy = Gen_UserLoginID
            RsTxn_MonthwiseSecLoc!CreatedDate = Now
            RsTxn_MonthwiseSecLoc.Update
        End If
        RsTxn_MonthwiseSecLocDetail.AddNew
        RsTxn_MonthwiseSecLocDetail!MonthwiseTxnDetailID = GetMaxTxnDetails
        RsTxn_MonthwiseSecLocDetail!MonthWiseTxnID = mMaxTxnID
        RsTxn_MonthwiseSecLocDetail!SecurityTypeID = MSHFlexGrid1.TextMatrix(I, 16)
        RsTxn_MonthwiseSecLocDetail!SecurityDesignationID = MSHFlexGrid1.TextMatrix(I, 17)
        RsTxn_MonthwiseSecLocDetail!StartDate = Format(CDate(MSHFlexGrid1.TextMatrix(I, 5)), "yyyy-mm-dd")
        RsTxn_MonthwiseSecLocDetail!EndDate = Format(CDate(MSHFlexGrid1.TextMatrix(I, 6)), "yyyy-mm-dd")
        RsTxn_MonthwiseSecLocDetail!NoofGodown = Val(MSHFlexGrid1.TextMatrix(I, 7))
        RsTxn_MonthwiseSecLocDetail!RatePerMonth = Val(MSHFlexGrid1.TextMatrix(I, 8))
        If MSHFlexGrid1.TextMatrix(I, 18) <> "" Then
           RsTxn_MonthwiseSecLocDetail!CMPMID = Val(MSHFlexGrid1.TextMatrix(I, 18))
        Else
           RsTxn_MonthwiseSecLocDetail!CMPMID = Null
        End If
        'RsTxn_MonthwiseSecLocDetail!AttendanceSheetID = "NA"
        RsTxn_MonthwiseSecLocDetail!NoOfDays = Val(MSHFlexGrid1.TextMatrix(I, 10))
        RsTxn_MonthwiseSecLocDetail!EstimatedAmt = Val(MSHFlexGrid1.TextMatrix(I, 11))
        RsTxn_MonthwiseSecLocDetail!EstimatedSTax = 0 ' Val(MSHFlexGrid2.TextMatrix(I, 12))
        RsTxn_MonthwiseSecLocDetail!TotEstimatedAmt = Val(MSHFlexGrid1.TextMatrix(I, 13))
        'RsTxn_MonthwiseSecLocDetail!Remarks = Null 'Trim(MSHFlexGrid2.TextMatrix(I, 15))
        RsTxn_MonthwiseSecLocDetail!SecLocMapDetailID = Val(MSHFlexGrid1.TextMatrix(I, 19))
        RsTxn_MonthwiseSecLocDetail!SalaryPaidUpto = MSHFlexGrid1.TextMatrix(I, 15)
        RsTxn_MonthwiseSecLocDetail!CreatedBy = Gen_UserLoginID
        RsTxn_MonthwiseSecLocDetail!CreatedDate = Now
        RsTxn_MonthwiseSecLocDetail.Update
    End If
ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!!"
Call CmdClear_Click

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If

Call TableTxn_MonthwiseSecLoc
Call TableTxn_MonthwiseSecLocDetails

Call Grid_Head

ChkAllLocation.Value = 0
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Then
   ChkAllLocation.Visible = True
Else
   ChkAllLocation.Visible = False
End If

End Sub
Private Sub Grid_Head()
ChkCheckAll.Value = vbUnchecked
Label22.Caption = ""
CmdSave.Enabled = False
CmdExcel.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 22
    
    .TextMatrix(0, 0) = "Selection"
    .TextMatrix(0, 1) = "Location Name"
    .TextMatrix(0, 2) = "Security Agency"
    .TextMatrix(0, 3) = "Security Type"
    .TextMatrix(0, 4) = "Security Designation"
    .TextMatrix(0, 5) = "Start Date"
    .TextMatrix(0, 6) = "End Date"
    .TextMatrix(0, 7) = "No of Godown"
    .TextMatrix(0, 8) = "Rate per month"
    .TextMatrix(0, 9) = "CMPM Name"
    .TextMatrix(0, 10) = "No of Days"
    .TextMatrix(0, 11) = "Estimated Amount"
    .TextMatrix(0, 12) = "Estimated S.Tax"
    .TextMatrix(0, 13) = "Total Est. Amount"
    .TextMatrix(0, 14) = "Remarks"
    .TextMatrix(0, 15) = "Salary Payable Upto"
    .TextMatrix(0, 16) = "Security TypeID"
    .TextMatrix(0, 17) = "Security DesignationID"
    .TextMatrix(0, 18) = "CMPMID"
    .TextMatrix(0, 19) = "TxnDetailID"
    .TextMatrix(0, 20) = "LocationID"
    .TextMatrix(0, 21) = "Security Agency ID"
    
    
    .ColWidth(0) = 900
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 2000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    .ColWidth(13) = 1000
    .ColWidth(14) = 1500
    .ColWidth(15) = 1000
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    .ColWidth(20) = 0
    .ColWidth(21) = 0
    
    .RowHeight(0) = 600
    .WordWrap = True
End With
End Sub

Private Function GetMaxTxnID() As Double
Dim Retval As Double
tmp = "Select max(MonthwiseTxnID) from Txn_MonthwiseSecLoc "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTxnID = Retval
End Function

Function GetMaxTxnDetails() As Double
Dim Retval As Double
tmp = "Select max(MonthwiseTxnDetailID) From Txn_MonthwiseSecLocDetail " 'Where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTxnDetails = Retval
End Function


Private Sub MSHFlexGrid1_Click()
mSelRow = MSHFlexGrid1.RowSel
mSelCol = MSHFlexGrid1.ColSel

If mSelRow = 0 Or mSelRow = MSHFlexGrid1.Rows - 1 Then Exit Sub

If mSelCol = 0 Then
    If MSHFlexGrid1.TextMatrix(mSelRow, 0) = strChecked Then
        MSHFlexGrid1.TextMatrix(mSelRow, 0) = strUnChecked
        tmp = 0
        ChkCheckAll.Value = vbUnchecked
        tmp = 1
    Else
        MSHFlexGrid1.TextMatrix(mSelRow, 0) = strChecked
    End If
End If

End Sub

Private Function ChkAttendancenotMarked() As Boolean
'ChkAttendancenotMarked = True
'tmp = " Select"
'tmp = tmp & " MSLD.TxnDetailID, count(TSA.AttendanceID) TotalAttendance,"
'tmp = tmp & " case"
'tmp = tmp & " when MSLD.StartDate < '" & mFromDate & "' and MSLD.EndDate >= '" & mToDate & "' then datediff(d, '" & mFromDate & "', '" & mToDate & "')+1"
'tmp = tmp & " when MSLD.StartDate >= '" & mFromDate & "' and MSLD.EndDate > '" & mToDate & "' then datediff(d, MSLD.StartDate, '" & mToDate & "')+1"
'tmp = tmp & " when MSLD.StartDate >= '" & mFromDate & "' and MSLD.EndDate <= '" & mToDate & "' then datediff(d, MSLD.StartDate, MSLD.EndDate)+1"
'tmp = tmp & " end as 'TotalDays'"
'tmp = tmp & " from  Mst_SecLocMappingDEtail MSLD"
'tmp = tmp & " Inner Join MSt_SecLocMapping MSL on MSL.TxnID = MSLD.TxnID"
'tmp = tmp & " Left Join Txn_securityAttendance TSA on TSA.TxnDetailID = MSLD.TxnDetailID and TSA.AttendanceDate between '" & mFromDate & "' and '" & mToDate & "'"
'tmp = tmp & " Where MSL.LocationID = " & mLocationID & " and MSLD.StartDate <= '" & mToDate & "' And  MSLD.EndDate >= '" & mFromDate & "'"
'tmp = tmp & " group by MSLD.TxnDetailID, MSLD.StartDate, MSLD.EndDate"
'
'
'Call Tmp7Table
'
'If TmpRs7.RecordCount > 0 Then
'   For j = 1 To TmpRs7.RecordCount
'      If IIf(IsNull(TmpRs7!TotalAttendance), 0, TmpRs7!TotalAttendance) <> IIf(IsNull(TmpRs7!TotalDays), 0, TmpRs7!TotalDays) Then
'         ChkAttendancenotMarked = False
'         Exit For
'      End If
'      TmpRs7.MoveNext
'   Next
'End If


ChkAttendancenotMarked = True
tmp = " Select"
tmp = tmp & " MSLD.TxnDetailID, count(TSA.AttendanceID) TotalAttendance,"
tmp = tmp & " case"
tmp = tmp & " when MSLD.StartDate <= '" & mFromDate & "' and MSLD.EndDate >= '" & mToDate & "' then datediff(d, '" & mFromDate & "', '" & mToDate & "')+1"
tmp = tmp & " when MSLD.StartDate <= '" & mFromDate & "' and MSLD.EndDate <= '" & mToDate & "' then datediff(d, '" & mFromDate & "', MSLD.EndDate)+1"
tmp = tmp & " when MSLD.StartDate >= '" & mFromDate & "' and MSLD.EndDate >= '" & mToDate & "' then datediff(d, MSLD.StartDate, '" & mToDate & "')+1"
tmp = tmp & " when MSLD.StartDate >= '" & mFromDate & "' and MSLD.EndDate <= '" & mToDate & "' then datediff(d, MSLD.StartDate, MSLD.EndDate)+1 "
tmp = tmp & " end as 'TotalDays'"
tmp = tmp & " from  Mst_SecLocMappingDEtail MSLD"
tmp = tmp & " Inner Join MSt_SecLocMapping MSL on MSL.TxnID = MSLD.TxnID"
tmp = tmp & " Left Join Txn_securityAttendance TSA on TSA.TxnDetailID = MSLD.TxnDetailID and TSA.AttendanceDate between '" & mFromDate & "' and '" & mToDate & "' "
tmp = tmp & " Where MSL.SecurityAgencyID <> 80 And  MSL.LocationID = " & mLocationID & " and MSLD.StartDate <= '" & mToDate & "' And  MSLD.EndDate >= '" & mFromDate & "'"
tmp = tmp & " group by MSLD.TxnDetailID, MSLD.StartDate, MSLD.EndDate"
Call Tmp7Table

If TmpRs7.RecordCount > 0 Then
   For j = 1 To TmpRs7.RecordCount
      If IIf(IsNull(TmpRs7!TotalAttendance), 0, TmpRs7!TotalAttendance) <> IIf(IsNull(TmpRs7!TotalDays), 0, TmpRs7!TotalDays) Then
         ChkAttendancenotMarked = False
         Exit For
      End If
      TmpRs7.MoveNext
   Next
End If


End Function
