VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_MonthWiseSecurityDetailBilledUnbilled 
   Caption         =   "Security Detail "
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
      Height          =   1095
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   14880
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
         Height          =   360
         Left            =   13530
         TabIndex        =   7
         Top             =   555
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
         Height          =   360
         Left            =   11100
         TabIndex        =   6
         Top             =   555
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
         Height          =   360
         Left            =   12315
         TabIndex        =   5
         Top             =   555
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
         Left            =   2910
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   555
         Width           =   2625
      End
      Begin VB.ComboBox CmbAgencyName 
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
         Left            =   5655
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   555
         Width           =   5280
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   555
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   48300035
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   1500
         TabIndex        =   9
         Top             =   555
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   48300035
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label lblRec 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   11115
         TabIndex        =   15
         Top             =   172
         Width           =   2280
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
         Left            =   13515
         TabIndex        =   14
         Top             =   172
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
         Left            =   1800
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   165
         Width           =   1005
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
         Left            =   3840
         TabIndex        =   11
         Top             =   172
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agency Name"
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
         Left            =   7650
         TabIndex        =   10
         Top             =   172
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8295
      Left            =   45
      TabIndex        =   0
      Top             =   1110
      Width           =   14895
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14655
         _ExtentX        =   25850
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   7800
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_MonthWiseSecurityDetailBilledUnbilled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mAgencyID, mLocationID As Variant
Dim mStartDate, mEndDate As Variant

Private Sub CmbAgencyName_GotFocus()
tmp = CmbAgencyName.Text
Call TableMst_SecurityAgency("")

CmbAgencyName.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Agency found"
   CmbAgencyName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbAgencyName.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbAgencyName.Text = tmp
End Sub

Private Sub CmbAgencyName_LostFocus()
If CmbAgencyName.Text = "" Then
   mAgencyID = 0
   Exit Sub
End If
' Code for clearing Godown combox in case of change

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbAgencyName.Text & "'"
If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Agency Selection "
   CmbAgencyName.SetFocus
   Exit Sub
End If

mAgencyID = RsMst_SecurityAgency!SecurityAgencyID
End Sub


Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location("")
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
If CmbLocationID.Text = "" Then Exit Sub
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

Private Sub CmdClear_Click()
CmbLocationID.Text = ""
CmbAgencyName.Text = ""
TxtFromDate.Value = Date
Call TxtFromDate_LostFocus
TxtToDate.Value = Date
Call TxtToDate_LostFocus
Call Grid_Head
End Sub

Private Sub Grid_Head()
CmdExcel.Enabled = False
lblRecordCount.Caption = ""
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 12
    .Font.Size = 10
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(1, 0) = "Month Year"
    .TextMatrix(1, 1) = "State Name"
    .TextMatrix(1, 2) = "Location Name"
    .TextMatrix(1, 3) = "Bill TxnID/Security Detail ID"
    .TextMatrix(1, 4) = "Agency Name"
    .TextMatrix(1, 5) = "Security Type" 'NCDEX / Non - NCDEX
    .TextMatrix(1, 6) = "Security Designation"
    .TextMatrix(1, 7) = "No of Godown"
    .TextMatrix(1, 8) = "Rate Per Month"
    .TextMatrix(1, 9) = "Description"
    .TextMatrix(1, 10) = "Amount"
    .TextMatrix(1, 11) = "Flag"
    
    .RowHeight(1) = 900
End With
End Sub

Private Sub CmdExcel_Click()

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_LocationwiseClient.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_LocationwiseClient.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1 ' + 6
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant
If Format(TxtFromDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future Date not allowed!!"
   TxtFromDate.SetFocus
   Exit Sub
End If
If Format(TxtToDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future Date not allowed!!"
   TxtToDate.SetFocus
   Exit Sub
End If
If Format(TxtFromDate.Value, "yyyy-mm") > Format(TxtToDate.Value, "yyyy-mm") Then
   MsgBox "From Date should be less than To Date !!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If CmbLocationID.Text <> "" Then
   wc = " and ML.LOcationID = " & mLocationID
ElseIf Gen_WcLocation <> "" Then
   wc = " and ML." & Gen_WcLocation
End If
If CmbAgencyName.Text <> "" Then
   wc = wc & " and MSA.SecurityAgencyID = " & mAgencyID
End If

MsgBox "Wait till next message!!!"
Call Grid_Head

For k = 1 To DateDiff("m", TxtFromDate.Value, TxtToDate.Value) + 1
   
   mStartDate = Format(DateAdd("m", k - 1, TxtFromDate.Value), "yyyy-mm-dd")
   mEndDate = Format(DateSerial(Year(CDate(mStartDate)), Month(CDate(mStartDate)) + 1, 0), "yyyy-mm-dd")
   
   tmp = " Select MS.StateName, ML.LocationName, TSBD.BillTxnID, TSB.MonthYear, MSA.AgencyName, MSS.SecurityShifts,"
   tmp = tmp & " MSD.SecurityDesignation, NoofGodown, RatePerMonth, MST.Description, TSBD.BillDetailAmt Amount, 'B' Flag"
   tmp = tmp & " From Txn_SecurityBilling TSB"
   tmp = tmp & " Inner Join Txn_SecurityBillingDetail TSBD on TSBD.BillTxnID = TSB.BillTxnID"
   tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TSBD.LocationID And TSBD.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
   tmp = tmp & " Inner Join MSt_State MS on MS.StateID = ML.StateID"
   tmp = tmp & " Inner Join Mst_ServiceTax MST on MST.ServiceTaxID =  TSB.STaxID"
   tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.SecurityAgencyID = TSB.SecurityAgencyID"
   tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = TSBD.SecurityDesignationID"
   tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = TSBD.SecurityTypeID"
   tmp = tmp & " Where TSB.Status in ('A', 'P') And TSB.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "'"
   tmp = tmp & wc
   tmp = tmp & " Union All"
   tmp = tmp & " select MS.StateName, ML.LocationName, MSLD.TxnDetailID, '" & mStartDate & "' MonthYear, MSA.AgencyName, MSS.SecurityShifts,"
   tmp = tmp & " MSD.SecurityDesignation, MSLD.NoofGodown, MSLD.RatePerMonth, '' Description, MSLD.RatePerMonth Amount, 'UB' Flag"
   tmp = tmp & " from Mst_SeclocMappingDetail  MSLD"
   tmp = tmp & " Inner Join Mst_SecLocMapping MSL on MSL.TxnID = MSLD.TxnID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MSL.LocationID And MSL.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
   tmp = tmp & " Inner Join MSt_State MS on MS.StateID = ML.StateID"
   tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.SecurityAgencyID = MSL.SecurityAgencyID"
   tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLD.SecurityDesignationID"
   tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLD.SecurityTypeID"
   tmp = tmp & " Where '" & mStartDate & "' Between MSLD.StartDate And MSLD.EndDate"
   tmp = tmp & " And MSLD.TxnDetailID not in (Select SecLocMapDetailID from Txn_MonthwiseSecLocDetail  TMSLD"
   tmp = tmp & " Inner Join Txn_SecurityBillingDetail TSBD on TSBD.MonthwiseTxnDetailID = TMSLD.MonthwiseTxnDetailID"
   tmp = tmp & " Inner Join Txn_SecurityBilling TSB on TSB.BillTxnID = TSBD.BillTxnID"
   tmp = tmp & " Where TSB.Status in ('A', 'P', 'R') And TSB.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "')"
   tmp = tmp & " And MSLD.TxnDetailID not in (select txndetailid from Txn_SecurityAttendance"
   tmp = tmp & " where AttendanceDate between '" & mStartDate & "' And '" & mEndDate & "')"
   tmp = tmp & wc
   tmp = tmp & " Union All"
   tmp = tmp & " select  MS.StateName, ML.LocationName, MSLD.TxnDetailID, '" & mStartDate & "' MonthYear, MSA.AgencyName, MSS.SecurityShifts,"
   tmp = tmp & " MSD.SecurityDesignation, MSLD.NoofGodown, MSLD.RatePerMonth, '' Description, ((msld.ratepermonth * count(attendanceid)) / datediff(d, '" & mStartDate & "','" & mEndDate & "')+1) Amount, 'UB' Flag"
   tmp = tmp & " from mst_seclocmappingdetail msld"
   tmp = tmp & " inner join txn_securityattendance tsa on msld.txndetailid = tsa.txndetailid"
   tmp = tmp & " inner join mst_seclocmapping msl on msld.txnid = msl.txnid"
   tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MSL.LocationID And MSL.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
   tmp = tmp & " Inner Join MSt_State MS on MS.StateID = ML.StateID"
   tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.SecurityAgencyID = MSL.SecurityAgencyID"
   tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLD.SecurityDesignationID"
   tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLD.SecurityTypeID"
   tmp = tmp & " where tsa.status  = 'A' and tsa.attendancedate between '" & mStartDate & "' and '" & mEndDate & "'"
   tmp = tmp & " and msld.txndetailid not in (select isnull(msld.SecLocMapDetailID,0) from Txn_MonthwiseSecLocDetail msld"
   tmp = tmp & " inner join Txn_MonthwiseSecLoc msl on msl.monthwisetxnid = msld.monthwisetxnid"
   tmp = tmp & " where monthyear Between  '" & mStartDate & "' And '" & mEndDate & "')"
   tmp = tmp & wc
   tmp = tmp & " group by MS.StateName, ML.LocationName, MSLD.TxnDetailID, MSA.AgencyName, MSS.SecurityShifts, MSD.SecurityDesignation, MSLD.NoofGodown, MSLD.RatePerMonth"
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      CmdExcel.Enabled = True
      ProgressBar1.Value = 0
      ProgressBar1.Max = TmpRs.RecordCount + 2
      With MSHFlexGrid1
         For j = 1 To TmpRs.RecordCount
            I = MSHFlexGrid1.Rows - 1
            .TextMatrix(I, 0) = Format(mStartDate, "MMM-yyyy")
            .TextMatrix(I, 1) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs!BillTxnID), "", TmpRs!BillTxnID)
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName)
            .TextMatrix(I, 5) = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts)
            .TextMatrix(I, 6) = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation)
            .TextMatrix(I, 7) = IIf(IsNull(TmpRs!NoofGodown), "", TmpRs!NoofGodown)
            .TextMatrix(I, 8) = IIf(IsNull(TmpRs!RatePerMonth), "", TmpRs!RatePerMonth)
            .TextMatrix(I, 9) = IIf(IsNull(TmpRs!Description), "", TmpRs!Description)
            .TextMatrix(I, 10) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
            .TextMatrix(I, 10) = Format(.TextMatrix(I, 10), "#######0.00")
            .TextMatrix(0, 10) = Format(Val(.TextMatrix(I, 10)) + Val(.TextMatrix(0, 10)), "#######0.00")
            .TextMatrix(I, 11) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag)
            TmpRs.MoveNext
            MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
         Next j
      End With
   End If
Next k
ProgressBar1.Value = ProgressBar1.Max
lblRecordCount.Caption = MSHFlexGrid1.Rows - 3
MsgBox "Done!!!"
End Sub
Private Sub Form_Load()
'Dim FrmLoadAccess() As Boolean
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
End Sub


Private Sub TxtFromDate_LostFocus()
TxtFromDate.Day = 1
End Sub

Private Sub TxtToDate_LostFocus()
TxtToDate.Value = DateSerial(TxtToDate.Year, TxtToDate.Month + 1, 0)
End Sub
