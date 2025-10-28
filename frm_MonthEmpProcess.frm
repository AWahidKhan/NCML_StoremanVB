VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_MonthEmpProcess 
   Caption         =   "Month-wise Employee Process"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8340
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   14610
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
         Left            =   4200
         TabIndex        =   13
         Top             =   360
         Width           =   1800
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
         Left            =   10920
         TabIndex        =   12
         Top             =   360
         Width           =   1560
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
         Left            =   9360
         TabIndex        =   11
         Top             =   360
         Width           =   1560
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
         Left            =   7800
         TabIndex        =   10
         Top             =   360
         Width           =   1560
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
         Left            =   12480
         TabIndex        =   9
         Top             =   360
         Width           =   1560
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
         ItemData        =   "frm_MonthEmpProcess.frx":0000
         Left            =   2595
         List            =   "frm_MonthEmpProcess.frx":0002
         TabIndex        =   5
         Top             =   375
         Width           =   1500
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Show Employee"
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
         Left            =   6000
         TabIndex        =   4
         Top             =   360
         Width           =   1800
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
         ItemData        =   "frm_MonthEmpProcess.frx":0004
         Left            =   120
         List            =   "frm_MonthEmpProcess.frx":0006
         TabIndex        =   3
         Top             =   375
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7095
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   12515
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   3090
         TabIndex        =   8
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   1027
         TabIndex        =   7
         Top             =   135
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   120
      TabIndex        =   0
      Top             =   8280
      Width           =   14610
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   75
         TabIndex        =   1
         Top             =   210
         Width           =   14355
         _ExtentX        =   25321
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "Frm_monthEmpProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth, mFromDate, mMonthListindex, mSM_Prefix, mTotalDays, mToDate, mNoofdays As Variant
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
       mMonthListindex = I '+ 1
       mStatus = True
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

End Sub

Private Sub CmdClear_Click()
CmbMonth.Text = ""
CmbYear.Text = ""
Call Grid_Head
CmdExcel.Enabled = False
CmdSave.Enabled = False
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdGetReport_Click()
Dim wc As String

wc = ""

If CmbMonth.Text = "" Then
   MsgBox "Please select Month"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox "Please select Year"
   CmbYear.SetFocus
   Exit Sub
End If

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

Call Grid_Head

tmp = "Select EMP.MonthEmpID,EMP.EmployeeNo,EMP.EmployeeName,MD.Designation,EMP.Flag,Emp.EmployeeID,EMP1.EmployeeName+'~'+ EMP1.EmployeeNo as ReportingTo,MD1.Designation as RDesignation"
tmp = tmp & "  From Txn_MonthEmp_History EMP "
tmp = tmp & "  Inner Join Mst_Designation MD On EMP.DesignationID = MD.DesignationID"
tmp = tmp & "  left Outer Join Txn_MonthEmp_History EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID And EMP.TxnDate=EMP1.TxnDate"
tmp = tmp & "  Inner Join Mst_Designation MD1 On EMP1.DesignationID = MD1.DesignationID"
tmp = tmp & "  where EMP.TxnDate='" & Format(mToDate, Gen_DatFormat) & "' And EMP.DesignationID not in (1,7) And EMP.Flag='Y'"
tmp = tmp & "  order by EMP.EmployeeName"

Call TmpTable
ProgressBar1.Value = 0
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount
   For o = 1 To TmpRs.RecordCount
       DoEvents
       Call dispdata
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   CmdExcel.Enabled = True
   CmdSave.Enabled = True
   MsgBox "Done!!"
Else
   CmdExcel.Enabled = False
   CmdSave.Enabled = False
   MsgBox "Record not found"
End If
'ProgressBar1.Value = 0
End Sub
Public Sub dispdata()
Dim MaxR, x, mFlag As Double
x = 1
MaxR = 0

tmp = "Select distinct EGM.BaselocationID,ML.LocationName"
tmp = tmp & " From Txn_MonthEmp_History EGM"
tmp = tmp & " inner Join Mst_Location ML On EGM.BaseLocationID = ML.LocationID"
tmp = tmp & " left Outer Join Txn_MonthEmp_History EMP On EGM.EmployeeID = EMP.EmployeeID And  EGM.TxnDate=EMP.TxnDate"
tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID And  EMP.TxnDate=EMP1.TxnDate"
tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID and EMP1.TxnDate=EMP2.TxnDate"
tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID and EMP2.TxnDate=EMP3.TxnDate"
tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID And EMP3.TxnDate=EMP4.TxnDate"
tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP5 On EMP.ConnectedToRAG = EMP5.EmployeeID And EMP.TxnDate=EMP5.TxnDate"
tmp = tmp & " Where (EGM.EmployeeID =" & TmpRs!EmployeeID & " or EMP.EmployeeID=" & TmpRs!EmployeeID & " or EMP1.EmployeeID=" & TmpRs!EmployeeID & " or EMP2.EmployeeID=" & TmpRs!EmployeeID & ""
tmp = tmp & " or EMP3.EmployeeID=" & TmpRs!EmployeeID & " or EMP4.EmployeeID=" & TmpRs!EmployeeID & " or EMP5.EmployeeID=" & TmpRs!EmployeeID & " ) and EGM.BaselocationID is not null And EGM.TxnDate='" & Format(mToDate, Gen_DatFormat) & "'"

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       x = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(x, 0) = IIf(IsNull(TmpRs!MonthEmpID), "", TmpRs!MonthEmpID)
       MSHFlexGrid1.TextMatrix(x, 1) = IIf(IsNull(TmpRs!EmployeeNo), "", TmpRs!EmployeeNo)
       MSHFlexGrid1.TextMatrix(x, 2) = IIf(IsNull(TmpRs!EmployeeName), "", TmpRs!EmployeeName)
       MSHFlexGrid1.TextMatrix(x, 3) = IIf(IsNull(TmpRs!Designation), "", TmpRs!Designation)
       If TmpRs!Flag = "Y" Then
          MSHFlexGrid1.TextMatrix(x, 4) = "Active"
       Else
          MSHFlexGrid1.TextMatrix(x, 4) = "De-Active"
       End If
       MSHFlexGrid1.TextMatrix(x, 5) = IIf(IsNull(TmpRs2.Fields(0)), "", TmpRs2.Fields(0))
       MSHFlexGrid1.TextMatrix(x, 6) = IIf(IsNull(TmpRs2.Fields(1)), "", TmpRs2.Fields(1))
       MSHFlexGrid1.TextMatrix(x, 7) = IIf(IsNull(TmpRs!ReportingTo), "", TmpRs!ReportingTo)
       MSHFlexGrid1.TextMatrix(x, 8) = IIf(IsNull(TmpRs!RDesignation), "", TmpRs!RDesignation)
       
       TmpRs2.MoveNext
   Next
End If

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_MonthEmpAllocation" & CmbMonth.Text & "-" & CmbYear.Text & ".xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_MonthEmpAllocation" & CmbMonth.Text & "-" & CmbYear.Text & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdProcess_Click()

If CmbMonth.Text = "" Then
   MsgBox "Please select Month"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox "Please select Year"
   CmbYear.SetFocus
   Exit Sub
End If

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

Call TableTxn_MonthEmp_History(" Where TxnDate='" & Format(mToDate, Gen_DatFormat) & "'")

If RsTxn_MonthEmp_History.RecordCount > 0 Then
   MsgBox "Data already exist"
   Exit Sub
End If
CmdProcess.Enabled = False
MsgBox "Please wait till next message!"

tmp = "select Max(MonthEmpID) From Txn_MonthEmp_History"
Call TmpTable
mMaxID = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)

Call TableMst_Employee

tmp = "Select * from Mst_Employee"
Call TmpTableHRPro
Dim mDOJ, mDOL As Variant

For I = 1 To RsMst_Employee.RecordCount
    mDOJ = Null
    mDOL = Null
    If TmpRsHRPro.RecordCount > 0 Then
       TmpRsHRPro.MoveFirst
       TmpRsHRPro.Find "EmployeeNo='" & RsMst_Employee!EmployeeNo & "'"
       If TmpRsHRPro.EOF = False Then
          mDOL = TmpRsHRPro!DOL
          mDOJ = TmpRsHRPro!JoinDate
       End If
    End If
    tmp = "Insert into Txn_MonthEmp_History"
    tmp = tmp & " Select " & mMaxID & ",'" & mToDate & "',*,0,0," & IIf(IsNull(mDOL), "Null", "'" & mDOL & "'") & "," & IIf(IsNull(mDOJ), "Null", "'" & mDOJ & "'") & " From Mst_Employee where EmployeeID=" & RsMst_Employee!EmployeeID
    Call TmpTable
    RsMst_Employee.MoveNext
    mMaxID = mMaxID + 1
Next
'Call TmpTable

MsgBox "Done!!"
CmdProcess.Enabled = True
End Sub

Private Sub CmdSave_Click()
If MsgBox("Are you wan't to save?", vbYesNo) = vbNo Then
   Exit Sub
End If
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows - 2
Call TableTxn_MonthEmp_HistoryDetail("")
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
       RsTxn_MonthEmp_HistoryDetail.Filter = "MonthEmpID=" & Val(MSHFlexGrid1.TextMatrix(I, 0)) & " And LocationID=" & Val(MSHFlexGrid1.TextMatrix(I, 5))
       If RsTxn_MonthEmp_HistoryDetail.RecordCount = 0 Then
          RsTxn_MonthEmp_HistoryDetail.AddNew
          RsTxn_MonthEmp_HistoryDetail!MonthEmpDetailID = GetMaxMonthEmpDetailID
          RsTxn_MonthEmp_HistoryDetail!MonthEmpID = Val(MSHFlexGrid1.TextMatrix(I, 0))
          RsTxn_MonthEmp_HistoryDetail!LocationID = Val(MSHFlexGrid1.TextMatrix(I, 5))
          RsTxn_MonthEmp_HistoryDetail!CreatedBy = Gen_UserLoginID
          RsTxn_MonthEmp_HistoryDetail!CreatedDate = Now
          RsTxn_MonthEmp_HistoryDetail.Update
       End If
       RsTxn_MonthEmp_HistoryDetail.Filter = ""
       ProgressBar1.Value = ProgressBar1.Value + 1
    End If
Next
MsgBox "Done!!"
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
'FrmLoadAccess = GetFrmLoadAccess(Me.Name)
'If FrmLoadAccess(0) = False Then
'   'Call GButtonLockAD(Me)
'   MsgBox "Access Denied !!!!!!!!!"
'   CmdGetReport.Enabled = False
'   CmdSave.Enabled = False
'   CmdExcel.Enabled = False
'   Exit Sub
'End If

Call Grid_Head
CmdGetReport.Enabled = True
CmdExcel.Enabled = False
CmdSave.Enabled = False
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 9

MSHFlexGrid1.TextMatrix(0, 0) = "MonthEmpID"
MSHFlexGrid1.TextMatrix(0, 1) = "Employee No"
MSHFlexGrid1.TextMatrix(0, 2) = "Employee Name "
MSHFlexGrid1.TextMatrix(0, 3) = "Designation"
MSHFlexGrid1.TextMatrix(0, 4) = "Status"
MSHFlexGrid1.TextMatrix(0, 5) = "LocationID"
MSHFlexGrid1.TextMatrix(0, 6) = "Location"
MSHFlexGrid1.TextMatrix(0, 7) = "Reporting To "
MSHFlexGrid1.TextMatrix(0, 8) = "Reporting To Designation"

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 0
MSHFlexGrid1.ColWidth(6) = 3000
MSHFlexGrid1.ColWidth(7) = 3000
MSHFlexGrid1.ColWidth(8) = 2200

End Sub

Private Function GetMaxMonthEmpDetailID() As Double
Dim Retval As Double
tmp = "Select max(MonthEmpDetailID) from Txn_MonthEmp_HistoryDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxMonthEmpDetailID = Retval
End Function
