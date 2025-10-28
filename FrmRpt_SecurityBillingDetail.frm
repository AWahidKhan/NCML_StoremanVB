VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_SecurityBillingDetail 
   Caption         =   "Security Billing Detail"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   14865
      Begin VB.OptionButton OptAttendanceBillUnprocessed 
         Caption         =   "Attendance Processed and Bill Un Processed "
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
         Left            =   4680
         TabIndex        =   27
         Top             =   210
         Width           =   4485
      End
      Begin VB.OptionButton OptUnProcessed 
         Caption         =   "Un Processed Attendance"
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
         TabIndex        =   26
         Top             =   210
         Width           =   2685
      End
      Begin VB.OptionButton OptProcessed 
         Caption         =   "Bill Processed"
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
         Left            =   150
         TabIndex        =   25
         Top             =   210
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   14895
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7335
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   12938
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
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   14880
      Begin VB.ComboBox CmbAttendance 
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
         Height          =   360
         ItemData        =   "FrmRpt_SecurityBillingDetail.frx":0000
         Left            =   5760
         List            =   "FrmRpt_SecurityBillingDetail.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   937
         Width           =   1305
      End
      Begin VB.OptionButton OptMonthYear 
         Caption         =   "MonthYear"
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
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptBillTxnDate 
         Caption         =   "Bill Txn Date"
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
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmRpt_SecurityBillingDetail.frx":0030
         Left            =   3120
         List            =   "FrmRpt_SecurityBillingDetail.frx":003D
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   937
         Width           =   1305
      End
      Begin VB.TextBox TxtBillNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         MaxLength       =   50
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox CmbForwardTo 
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
         Left            =   11677
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   420
         Width           =   3120
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
         Left            =   7215
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   420
         Width           =   4320
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
         Left            =   4470
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   2625
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
         Left            =   12075
         TabIndex        =   3
         Top             =   885
         Width           =   1335
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
         Left            =   10740
         TabIndex        =   2
         Top             =   885
         Width           =   1335
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
         Height          =   360
         Left            =   13410
         TabIndex        =   1
         Top             =   885
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
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
         Format          =   106233857
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   3060
         TabIndex        =   6
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
         Format          =   106233857
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Attendance"
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
         Left            =   4560
         TabIndex        =   29
         Top             =   990
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   2400
         TabIndex        =   19
         Top             =   990
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bill No"
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
         TabIndex        =   17
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forward To"
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
         Left            =   12720
         TabIndex        =   15
         Top             =   180
         Width           =   1035
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
         Left            =   8730
         TabIndex        =   13
         Top             =   180
         Width           =   1290
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
         Left            =   5400
         TabIndex        =   11
         Top             =   180
         Width           =   765
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
         Left            =   1860
         TabIndex        =   10
         Top             =   165
         Width           =   1005
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
         Left            =   3360
         TabIndex        =   9
         Top             =   165
         Width           =   780
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
         Left            =   9675
         TabIndex        =   8
         Top             =   975
         Width           =   210
      End
      Begin VB.Label lblRec 
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
         Left            =   7275
         TabIndex        =   7
         Top             =   975
         Width           =   2325
      End
   End
End
Attribute VB_Name = "FrmRpt_SecurityBillingDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mAgencyID, mEmployeeID, mStatus As Variant

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

Private Sub CmbAttendance_LostFocus()
If CmbAttendance.Text = "" Then Exit Sub

If LCase(CmbAttendance.Text) <> "attendance" And LCase(CmbAttendance.Text) <> "absence" And LCase(CmbAttendance.Text) <> "rejected" Then
   MsgBox "Invalid Selection!!!"
   CmbAttendance.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbForwardTo_GotFocus()
tmp = CmbForwardTo.Text

Call TableMst_Employee("")

CmbForwardTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee found"
   CmbForwardTo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbForwardTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbForwardTo.Text = tmp
End Sub

Private Sub CmbForwardTo_LostFocus()
If CmbForwardTo.Text = "" Then
   mEmployeeID = 0
   Exit Sub
End If

Dim Str_() As String
Str_ = Split(CmbForwardTo.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection!!! "
   CmbForwardTo.SetFocus
   Exit Sub
End If

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeName = '" & Str_(0) & "'"
If RsMst_Employee.EOF Then
   MsgBox "Invalid Employee Selection "
   CmbForwardTo.SetFocus
   Exit Sub
End If
mEmployeeID = RsMst_Employee!EmployeeID
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
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub

Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If
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

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) = "pending" Then
   mStatus = "P"
ElseIf LCase(CmbStatus.Text) = "approved" Then
   mStatus = "A"
ElseIf LCase(CmbStatus.Text) = "rejected" Then
   mStatus = "R"
Else
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdClear_Click()
OptProcessed.Value = True
Call Grid_Head
TxtFromDate = Date
TxtToDate = Date
CmdExcel.Enabled = False
CmbLocationID.Text = ""
CmbAgencyName.Text = ""
CmbForwardTo.Text = ""
TxtBillNo.Text = ""
CmbStatus.Text = ""
LblRecordCount.Caption = "-"
End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant
LblRecordCount.Caption = ""
If OptProcessed.Value = True Then

   Call Grid_Head
    
   MsgBox " Wait till next message !!!"
    
   tmp = "SELECT TSB.BillTxnID,"
   tmp = tmp & " TSB.BillTxnDate ,TSB.MonthYear,SA.AgencyName,TSB.ServiceTaxID,TSB.BillNo,TSB.BillDate,TSB.BillAmt,TSB.BillSTax,"
   tmp = tmp & " TSB.TotalAmt,TSB.Remarks,TSB.FAGRemarks,TSB.ApprovedDate,TSB.Status,"
   tmp = tmp & " MST.Description, LocationName, SecurityShifts,SecurityDesignation, TSBD.NoofGodown, TSBD.RatePerMonth,"
   tmp = tmp & " ME.EmployeeName +'~'+ ME.EmployeeNo as ForwardTo,"
   tmp = tmp & " Me.EmployeeID, TSBD.AttendanceSheetID, TSBD.NoofDays, TSBD.EstimatedAmt, TSBD.EstimatedSTax, TSBD.TotEstimatedAmt, BillDetailAmt,"
   tmp = tmp & " BillDetailSTax , BillNoOfDays, TSBD.MonthwiseTxnDetailID, TMSLD.SecLocMapDetailID"
   tmp = tmp & " FROM Txn_SecurityBillingDetail TSBD"
   tmp = tmp & " Inner Join Txn_SecurityBilling TSB on TSBD.BillTxnID  = TSB.BillTxnID"
   tmp = tmp & " Inner Join Txn_Monthwiseseclocdetail TMSLD on TMSLD.MonthwiseTxnDetailID  = TSBD.MonthwiseTxnDetailID"
   tmp = tmp & " Inner Join Mst_ServiceTax MST on TSB.ServiceTaxID = MST.ServiceTaxID"
   tmp = tmp & " Inner Join Mst_SecurityAgency SA on TSB.SecurityAgencyID  = SA.SecurityAgencyID"
   tmp = tmp & " Inner Join Mst_SecurityShifts MSS On MSS.SecurityShiftsID = TSBD.SecurityTypeID"
   tmp = tmp & " Inner Join Mst_SecurityDesignation MSD On MSD.SecurityDesignationID = TSBD.SecurityDesignationID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = TSBD.LocationID"
   tmp = tmp & " Left Join Mst_Employee ME On ME.EmployeeID = TSB.ForwardTo "
    'TSBD.CMPMID ""
   
   If OptMonthYear.Value = True Then
      TxtFromDate.Day = 1
      TxtToDate.Day = 1
      If Gen_DBType = "MDB" Then
         tmp = tmp & " And (TSB.MonthYear between #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#)"
      Else
         tmp = tmp & " where TSB.MonthYear between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' and '" & Format(TxtToDate.Value, Gen_DatFormat) & "'" '''   tmp = tmp & " And (CDTFDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'  And '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
      End If
   ElseIf OptBillTxnDate.Value = True Then
      If Gen_DBType = "MDB" Then
         tmp = tmp & " And (TSB.BillTxnDate  between #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#)"
      Else
         tmp = tmp & " where TSB.BillTxnDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' and '" & Format(TxtToDate.Value, Gen_DatFormat) & "'" '''   tmp = tmp & " And (CDTFDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'  And '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
      End If
   End If
    
   If CmbLocationID.Text <> "" Then
      tmp = tmp & " And ML.LocationID=" & mLocationID & ""
   Else
      If Gen_WcLocation <> "" Then
         tmp = tmp & "ML." & Gen_WcLocation
      End If
   End If
   If CmbAgencyName.Text <> "" Then
      tmp = tmp & " And SA.SecurityAgencyID=" & mAgencyID & ""
   End If
   If CmbForwardTo.Text <> "" Then
      tmp = tmp & " And ME.EmployeeID=" & mEmployeeID & " "
   End If
   If CmbStatus.Text <> "" Then
      tmp = tmp & " And TSB.Status ='" & mStatus & "'"
   End If
   If TxtBillNo.Text <> "" Then
      tmp = tmp & " And TSB.BillNo like '%" & TxtBillNo & "%'"
   End If
   
   tmp = tmp & " Order by TSB.BillTxnID"
   
   Call TmpTable
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   If TmpRs.RecordCount > 0 Then
      For I = 1 To TmpRs.RecordCount
          MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!BillTxnID), "", TmpRs!BillTxnID)
          MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!BillTxnDate), "", Format(TmpRs!BillTxnDate, "dd-MMM-yyyy"))
          MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!MonthYear), "", Format(TmpRs!MonthYear, "dd-MMM-yyyy"))
          MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName)
          MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!Description), "", TmpRs!Description)
          MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!BillNo), "", TmpRs!BillNo)
          MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!BillDate), "", Format(TmpRs!BillDate, "dd-MMM-yyyy"))
          MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!BillAmt), "", TmpRs!BillAmt)
          MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!BillSTax), "", TmpRs!BillSTax)
          MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!TotalAmt), "", TmpRs!TotalAmt)
          MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!ForwardTo), "", TmpRs!ForwardTo)
          MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
          MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!FAGRemarks), "", TmpRs!FAGRemarks)
          MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs!ApprovedDate), "", Format(TmpRs!ApprovedDate, "dd-MMM-yyyy"))
          MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!Status), "", TmpRs!Status)
          If MSHFlexGrid1.TextMatrix(I, 14) = "P" Then
             MSHFlexGrid1.TextMatrix(I, 14) = "Pending"
          ElseIf MSHFlexGrid1.TextMatrix(I, 14) = "A" Then
             MSHFlexGrid1.TextMatrix(I, 14) = "Approved"
          ElseIf MSHFlexGrid1.TextMatrix(I, 14) = "R" Then
             MSHFlexGrid1.TextMatrix(I, 14) = "Rejected"
          ElseIf MSHFlexGrid1.TextMatrix(I, 14) = "S" Then
             MSHFlexGrid1.TextMatrix(I, 14) = "Settled"
          End If
          MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
          MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts)
          MSHFlexGrid1.TextMatrix(I, 17) = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation)
          MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs!NoofGodown), "", TmpRs!NoofGodown)
          MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs!RatePerMonth), "", TmpRs!RatePerMonth)
          MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs!AttendanceSheetID), "", TmpRs!AttendanceSheetID)
          MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs!NoOfDays), "", TmpRs!NoOfDays)
          MSHFlexGrid1.TextMatrix(I, 22) = IIf(IsNull(TmpRs!EstimatedAmt), "", TmpRs!EstimatedAmt)
          MSHFlexGrid1.TextMatrix(I, 23) = IIf(IsNull(TmpRs!EstimatedSTax), "", TmpRs!EstimatedSTax)
          MSHFlexGrid1.TextMatrix(I, 24) = IIf(IsNull(TmpRs!TotEstimatedAmt), "", TmpRs!TotEstimatedAmt)
          MSHFlexGrid1.TextMatrix(I, 25) = IIf(IsNull(TmpRs!BillDetailAmt), "", TmpRs!BillDetailAmt)
          MSHFlexGrid1.TextMatrix(I, 26) = IIf(IsNull(TmpRs!BillDetailSTax), "", TmpRs!BillDetailSTax)
          MSHFlexGrid1.TextMatrix(I, 27) = IIf(IsNull(TmpRs!BillNoOfDays), "", TmpRs!BillNoOfDays)
          MSHFlexGrid1.TextMatrix(I, 28) = IIf(IsNull(TmpRs!SecLocMapDetailID), "", TmpRs!SecLocMapDetailID)
            
          TmpRs.MoveNext
      Next
      CmdExcel.Enabled = True
      MsgBox ("Done!!!")
      LblRecordCount.Caption = TmpRs.RecordCount
  Else
      MsgBox "No Record Found !!!"
  End If

ElseIf OptUnProcessed.Value = True Then
  
  mTotalDays = Day(DateSerial(Year(TxtToDate), Month(TxtToDate) + 1, 0))
  
  TxtFromDate.Day = 1
  TxtToDate.Day = mTotalDays
  Call Grid_Head1
  
  MsgBox " Wait till next message !!!"

  tmp = " select ML.LocationName, MA.AgencyName, MSLD.TxnDetailID, MSLD.TxnID, MSS.SecurityShifts, MSD.SecurityDesignation, MSLD.StartDate, MSLD.EndDate,"
  tmp = tmp & " MSLD.NoofGodown , MSLD.RatePerMonth, MSL.LocationID"
  tmp = tmp & " from Mst_SeclocMappingDetail  MSLD"
  tmp = tmp & " Inner Join Mst_SecLocMapping MSL on MSL.TxnID = MSLD.TxnID And MSL.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "') And MG.StartDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' )"
  tmp = tmp & " Inner Join Mst_Location ML on MSL.LocationID = ML.LocationID"
  tmp = tmp & " Inner Join Mst_SecurityAgency MA on MA.SecurityAgencyID = MSL.SecurityAgencyID"
  tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLD.SecurityTypeID"
  tmp = tmp & " Inner JOin Mst_SecurityDesignation MSD on MSD.SecurityDesignationID= MSLD.SecurityDesignationID"
  tmp = tmp & " Where MSLD.StartDate <= '" & TxtToDate.Value & "' And MSLD.EndDate >= '" & TxtFromDate.Value & "' " '" & TxtFromDate.Value & "' Between MSLD.StartDate And MSLD.EndDate"
  tmp = tmp & " And MSLD.TxnDetailID not in (Select SecLocMapDetailID from Txn_MonthwiseSecLocDetail  TMSLD"
  tmp = tmp & " Inner Join Txn_MonthwiseSecLoc TMSL on tmsld.monthwisetxnid = tmsl.monthwisetxnid "
  If Gen_DBType = "MDB" Then
     tmp = tmp & " Where TMSL.MonthYear Between  #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#)"
  Else
     tmp = tmp & " Where TMSL.MonthYear Between  '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
  End If
  If LCase(CmbAttendance.Text) = "attendance" Then
      tmp = tmp & " And MSLD.TxnDetailID in (Select tsa.TxnDetailID from txn_securityattendance tsa"
      If Gen_DBType = "MDB" Then
         tmp = tmp & " Where tsa.attendancedate Between  #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#"
      Else
         tmp = tmp & " Where tsa.attendancedate Between  '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
      End If
      tmp = tmp & " and tsa.status = 'A')"
  ElseIf LCase(CmbAttendance.Text) = "absence" Then
      tmp = tmp & " And MSLD.TxnDetailID in (Select tsa.TxnDetailID from txn_securityattendance tsa"
      If Gen_DBType = "MDB" Then
         tmp = tmp & " Where tsa.attendancedate Between  #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#"
      Else
         tmp = tmp & " Where tsa.attendancedate Between  '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
      End If
      tmp = tmp & " and tsa.status = 'L')"
  ElseIf LCase(CmbAttendance.Text) = "rejected" Then
      tmp = tmp & " And MSLD.TxnDetailID in (Select tsa.TxnDetailID from txn_securityattendance tsa"
      If Gen_DBType = "MDB" Then
         tmp = tmp & " Where tsa.attendancedate Between  #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#"
      Else
         tmp = tmp & " Where tsa.attendancedate Between  '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
      End If
      tmp = tmp & " and tsa.status = 'R')"
  End If
   
  If CmbLocationID.Text <> "" Then
     tmp = tmp & " And ML.LocationID=" & mLocationID & ""
  Else
     If Gen_WcLocation <> "" Then
        tmp = tmp & "ML." & Gen_WcLocation
     End If
  End If
  If CmbAgencyName.Text <> "" Then
     tmp = tmp & " And MA.SecurityAgencyID=" & mAgencyID & ""
  End If
  tmp = tmp & " Order by ML.LocationName, MA.AgencyName"
   
  Call TmpTable
  MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
  If TmpRs.RecordCount > 0 Then
     For I = 1 To TmpRs.RecordCount
         MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
         MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName)
         MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!TxnDetailID), "", TmpRs!TxnDetailID)
         MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!TxnID), "", TmpRs!TxnID)
         MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts)
         MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation)
         MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
         MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!EndDate), "", Format(TmpRs!EndDate, "dd-MMM-yyyy"))
         MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!NoofGodown), "", TmpRs!NoofGodown)
         MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!RatePerMonth), "", TmpRs!RatePerMonth)
           
         TmpRs.MoveNext
     Next
     CmdExcel.Enabled = True
     MsgBox ("Done!!!")
     LblRecordCount.Caption = TmpRs.RecordCount
  Else
     MsgBox "No Record Found !!!"
  End If

ElseIf OptAttendanceBillUnprocessed.Value = True Then
  
  mTotalDays = Day(DateSerial(Year(TxtToDate), Month(TxtToDate) + 1, 0))
  
  TxtFromDate.Day = 1
  TxtToDate.Day = mTotalDays
  Call Grid_Head2
  
  MsgBox " Wait till next message !!!"

  tmp = " select tmsl.monthyear, ml.locationname, msa.agencyname, tmsl.monthwisetxnid, msld.txnid, tmsld.totestimatedamt, tmsld.remarks, tmsld.SecLocMapDetailID"
  tmp = tmp & " from Txn_Monthwisesecloc tmsl"
  tmp = tmp & " inner join Txn_Monthwiseseclocdetail tmsld on tmsld.monthwisetxnid = tmsl.monthwisetxnid"
  tmp = tmp & " inner join mst_location ml on ml.locationid = tmsl.locationid"
  tmp = tmp & " inner join mst_securityagency msa on msa.SecurityAgencyID = tmsl.securityagencyid"
  tmp = tmp & " inner join mst_seclocmappingdetail msld on msld.txndetailid = tmsld.seclocmapdetailid"
  tmp = tmp & " inner join mst_seclocmapping msl on msl.txnid = msld.txnid"
  'tmp = tmp & " where tmsl.monthyear between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
  If Gen_DBType = "MDB" Then
     tmp = tmp & " Where TMSL.MonthYear Between  #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#"
  Else
     tmp = tmp & " Where TMSL.MonthYear Between  '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
  End If
  If CmbLocationID.Text <> "" Then
     tmp = tmp & " And ML.LocationID=" & mLocationID
  Else
     If Gen_WcLocation <> "" Then
        tmp = tmp & "And ML." & Gen_WcLocation
     End If
  End If
  If CmbAgencyName.Text <> "" Then
     tmp = tmp & " And MSA.SecurityAgencyID=" & mAgencyID
  End If
  tmp = tmp & " and tmsld.monthwisetxndetailid not in (select tsbd.monthwisetxndetailid from txn_securitybillingdetail tsbd"
  tmp = tmp & " inner join txn_securitybilling tsb on tsb.billtxnid = tsbd.billtxnid where tsb.status in ('A', 'P','S')"
  'tmp = tmp & " and tsb.monthyear between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "')"
  If Gen_DBType = "MDB" Then
     tmp = tmp & " And TSB.MonthYear Between  #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#)"
  Else
     tmp = tmp & " And TSB.MonthYear Between  '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
  End If
  tmp = tmp & " Order by ML.LocationName, MSA.AgencyName"
   
  Call TmpTable
  MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
  If TmpRs.RecordCount > 0 Then
     For I = 1 To TmpRs.RecordCount
         MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!MonthYear), "", Format(TmpRs!MonthYear, "MMM-yyyy"))
         MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
         MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName)
         MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!MonthWiseTxnID), "", TmpRs!MonthWiseTxnID)
         MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!TxnID), "", TmpRs!TxnID)
         MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!SecLocMapDetailID), "", TmpRs!SecLocMapDetailID)
         MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!TotEstimatedAmt), "", TmpRs!TotEstimatedAmt)
         MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
         'MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!SecLocMapDetailID), "", TmpRs!SecLocMapDetailID)
         TmpRs.MoveNext
     Next
     CmdExcel.Enabled = True
     MsgBox ("Done!!!")
     LblRecordCount.Caption = TmpRs.RecordCount
  Else
     MsgBox "No Record Found !!!"
  End If
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
If OptUnProcessed.Value = True Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_SecurityUnprocessedBillingDetail.xls")
Else
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_SecurityBillingDetail.xls")
End If
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
If OptUnProcessed.Value = True Then
   Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_SecurityUnprocessedBillingDetail.xls")
Else
   Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_SecurityBillingDetail.xls")
End If
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "mm/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 29
    .Font.Size = 10
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "BillTxnID"
    .TextMatrix(0, 1) = "Bill Txn Date "
    .TextMatrix(0, 2) = "Month Year"
    .TextMatrix(0, 3) = "Agency Name"
    .TextMatrix(0, 4) = "Service Tax"
    .TextMatrix(0, 5) = "Bill No" 'NCDEX / Non - NCDEX
    .TextMatrix(0, 6) = "Bill Date"
    .TextMatrix(0, 7) = "Bill Amount"
    .TextMatrix(0, 8) = "Bill S.Tax Amount"
    .TextMatrix(0, 9) = "Total Bill Amt"
    .TextMatrix(0, 10) = "Forward To"
    .TextMatrix(0, 11) = "Remarks"
    .TextMatrix(0, 12) = "FAG Remarks"
    .TextMatrix(0, 13) = "Navision Posting Date"
    .TextMatrix(0, 14) = "Status"    'in the GSL
    .TextMatrix(0, 15) = "Location Name"
    .TextMatrix(0, 16) = "Security Shifts"
    .TextMatrix(0, 17) = "Security Designation "
    .TextMatrix(0, 18) = "No of Godown"
    .TextMatrix(0, 19) = "Rate Per Month"
    .TextMatrix(0, 20) = "Attendance Sheet ID "
    .TextMatrix(0, 21) = "No of Days "
    .TextMatrix(0, 22) = "Estimated Amt "
    .TextMatrix(0, 23) = "Estimated S.Tax "
    .TextMatrix(0, 24) = "Tot Estimated Amt "
    .TextMatrix(0, 25) = "Bill Detail Amt"
    .TextMatrix(0, 26) = "Bill Detail STax"
    .TextMatrix(0, 27) = "Bill No Of Days "    'in the GSL
    .TextMatrix(0, 28) = "Security Detail ID"

    .RowHeight(0) = 900
End With

End Sub
Private Sub Grid_Head1()

With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 10
     .Font.Size = 10
     .WordWrap = True
    
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(0, 0) = "Location Name"
     .TextMatrix(0, 1) = "Agency Name "
     .TextMatrix(0, 2) = "Txn Detail ID"
     .TextMatrix(0, 3) = "Txn ID"
     .TextMatrix(0, 4) = "Security Type"
     .TextMatrix(0, 5) = "Security Designation" 'NCDEX / Non - NCDEX
     .TextMatrix(0, 6) = "Start Date"
     .TextMatrix(0, 7) = "End Date"
     .TextMatrix(0, 8) = "No Of Godown"
     .TextMatrix(0, 9) = "Rate per Month"
         
     .ColWidth(0) = 1600
     .ColWidth(1) = 1600
     .ColWidth(2) = 1600
     .ColWidth(3) = 1000
     .ColWidth(4) = 1000
     .ColWidth(5) = 1400
     .ColWidth(6) = 1200
     .ColWidth(7) = 1200
     .ColWidth(7) = 1200
     .ColWidth(7) = 1200
     
     .RowHeight(0) = 900
    
End With

End Sub

Private Sub Grid_Head2()

With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 8
     .Font.Size = 10
     .WordWrap = True
    
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(0, 0) = "Month Year"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "Agency Name "
     .TextMatrix(0, 3) = "Security Attendance ID"
     .TextMatrix(0, 4) = "Security Txn ID"
     .TextMatrix(0, 5) = "Security Detail ID"
     .TextMatrix(0, 6) = "Bill Amount"
     .TextMatrix(0, 7) = "Remarks"
     '.TextMatrix(0, 8) = "Security ID"
     
     .ColWidth(0) = 1500
     .ColWidth(1) = 1900
     .ColWidth(2) = 1900
     .ColWidth(3) = 1500
     .ColWidth(4) = 1500
     .ColWidth(5) = 1500
     .ColWidth(6) = 1500
     .ColWidth(7) = 4500
     
     .RowHeight(0) = 900
End With

End Sub

Private Sub Combo1_Change()

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
CmdExcel.Enabled = False
End Sub

Private Sub OptAttendanceBillUnprocessed_Click()
Call OptProcessed_Click
End Sub

Private Sub OptProcessed_Click()
LblRecordCount.Caption = ""
TxtToDate.Enabled = True
If OptUnProcessed.Value = True Then
   OptBillTxnDate.Enabled = False
   TxtBillNo.Locked = True
   CmbForwardTo.Enabled = False
   CmbStatus.Enabled = False
   OptMonthYear.Value = True
   TxtFromDate.Day = 1
   TxtToDate.Value = DateSerial(TxtFromDate.Year, TxtFromDate.Month + 1, 0)
   TxtToDate.Enabled = False
   CmbAttendance.Enabled = True
   Call Grid_Head1
ElseIf OptProcessed.Value = True Then
   OptBillTxnDate.Enabled = True
   TxtBillNo.Locked = False
   CmbForwardTo.Enabled = True
   CmbStatus.Enabled = True
   CmbAttendance.Enabled = False
   CmbAttendance.Text = ""
   Call Grid_Head
ElseIf OptAttendanceBillUnprocessed.Value = True Then
   OptBillTxnDate.Enabled = False
   TxtBillNo.Locked = True
   CmbForwardTo.Enabled = False
   CmbStatus.Enabled = False
   OptMonthYear.Value = True
   CmbAttendance.Enabled = False
   CmbAttendance.Text = ""
   Call Grid_Head2
End If
End Sub

Private Sub OptUnProcessed_Click()
Call OptProcessed_Click
End Sub

Private Sub TxtFromDate_LostFocus()
If OptUnProcessed.Value = True Then
   TxtFromDate.Day = 1
   TxtToDate.Value = DateSerial(TxtFromDate.Year, TxtFromDate.Month + 1, 0)
End If
End Sub

Private Sub TxtToDate_LostFocus()
'If OptUnProcessed.Value = True Then
'   TxtFromDate.Day = 1
'   TxtToDate.Value = DateSerial(TxtFromDate.Year, TxtFromDate.Month + 1, 0)
'End If
End Sub
