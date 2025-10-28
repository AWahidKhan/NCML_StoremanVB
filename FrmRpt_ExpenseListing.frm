VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_ExpenseListing 
   Caption         =   "Expense Listing"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   15075
      Begin VB.ComboBox CmbSExpenseType 
         Height          =   315
         ItemData        =   "FrmRpt_ExpenseListing.frx":0000
         Left            =   2835
         List            =   "FrmRpt_ExpenseListing.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   915
         Width           =   3990
      End
      Begin VB.ComboBox CmbSLocation 
         Height          =   315
         Left            =   6885
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   345
         Width           =   2325
      End
      Begin VB.ComboBox CmbSEmployee 
         Height          =   315
         Left            =   2835
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   345
         Width           =   3990
      End
      Begin VB.ComboBox CmbSStatus 
         Height          =   315
         ItemData        =   "FrmRpt_ExpenseListing.frx":0004
         Left            =   9240
         List            =   "FrmRpt_ExpenseListing.frx":0014
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   345
         Width           =   1995
      End
      Begin VB.ComboBox CmbSForwardTo 
         Height          =   315
         Left            =   11280
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   345
         Width           =   3630
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   9390
         TabIndex        =   9
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         Height          =   375
         Left            =   6960
         TabIndex        =   7
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   8175
         TabIndex        =   8
         Top             =   885
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   345
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61603841
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   75
         TabIndex        =   0
         Top             =   345
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61603841
         CurrentDate     =   36494
      End
      Begin VB.Label Label3 
         Caption         =   "Expense Type"
         Height          =   225
         Left            =   4200
         TabIndex        =   21
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label44 
         Caption         =   "Expense From Date"
         Height          =   255
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label Label36 
         Caption         =   "Location"
         Height          =   300
         Left            =   7680
         TabIndex        =   19
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label37 
         Caption         =   "Employee Name"
         Height          =   225
         Left            =   4200
         TabIndex        =   18
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Status"
         Height          =   225
         Left            =   9930
         TabIndex        =   16
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label31 
         Caption         =   "Forwarded To"
         Height          =   225
         Left            =   12585
         TabIndex        =   15
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label1 
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
         Left            =   10920
         TabIndex        =   14
         Top             =   960
         Width           =   2415
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
         Left            =   13440
         TabIndex        =   13
         Top             =   960
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7935
      Left            =   120
      TabIndex        =   11
      Top             =   1335
      Width           =   15075
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7575
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14865
         _ExtentX        =   26220
         _ExtentY        =   13361
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_ExpenseListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSStatus As Variant
Private Sub CmdClear_Click()


SFrom.Value = CDate(From_GenDate) - 30
STo.Value = Date 'To_GenDate
CmbSLocation.Text = ""
CmbSEmployee.Text = ""
CmbSStatus.Text = ""
CmbSForwardTo.Text = ""
CmbSExpenseType.Text = ""
Label2.Caption = ""

tmp = "Select EmployeeID,EmployeeName +'~'+ EmployeeNo as Employee,DesignationID,EmployeeEmailID From Mst_Employee Where EmployeeNo = '" & Gen_UserLoginID & "'"
Call TmpTable

If TmpRs.RecordCount = 1 Then
   CmbSForwardTo.Text = TmpRs!Employee
   mSForwardToID = TmpRs!EmployeeID
Else
   CmbSForwardTo.Text = ""
   mSForwardToID = Null
End If

Call Grid_Head
CmdExcel.Enabled = False


End Sub

Private Sub CmdGetReport_Click()

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim wc As Variant

Label2.Caption = ""
If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

Call Grid_Head
MsgBox ("Wait till Next Message")

wc = ""


tmp = "SELECT     Txn_ClaimSubmission.ClaimRowID, Txn_ClaimSubmission.ClaimDate, Txn_ClaimSubmission.ExpenseFrom,"
tmp = tmp & "     Txn_ClaimSubmission.ExpenseTo, EMP.EmployeeName, Mst_Location.LocationName,"
tmp = tmp & "     Txn_ClaimSubmission.SM_Prefix, Txn_ClaimSubmission.Amount,"
tmp = tmp & "     Txn_ClaimSubmission.Remarks, Mst_ExpenseType.ExpenseType, Txn_ClaimSubmissionDetail.Particulars,"
tmp = tmp & "     Txn_ClaimSubmissionDetail.Amount, Txn_ClaimSubmissionDetail.RecoverableAmount, Txn_ClaimSubmission.EmployeeID,Txn_ClaimSubmission.LocationID, "
tmp = tmp & "     EMP1.EmployeeName as 'Forwarded To',Txn_ClaimSubmission.Flag,EMP.AllocateExpense "
tmp = tmp & "FROM Txn_ClaimSubmission INNER JOIN"
tmp = tmp & "     Mst_Employee EMP ON Txn_ClaimSubmission.EmployeeID = EMP.EmployeeID INNER JOIN"
tmp = tmp & "     Mst_Employee EMP1 ON Txn_ClaimSubmission.ForwardToID = EMP1.EmployeeID INNER JOIN"
tmp = tmp & "     Mst_Location ON Txn_ClaimSubmission.LocationID = Mst_Location.LocationID INNER JOIN"
tmp = tmp & "     Txn_ClaimSubmissionDetail ON Txn_ClaimSubmission.ClaimRowID = Txn_ClaimSubmissionDetail.ClaimRowID"
'tmp = tmp & "     Txn_ClaimSubmission.SM_Prefix = Txn_ClaimSubmissionDetail.SM_Prefix INNER JOIN"
tmp = tmp & "     INNER JOIN"
tmp = tmp & "     Mst_ExpenseType ON Txn_ClaimSubmissionDetail.ExpenseTypeID = Mst_ExpenseType.ExpenseTypeID"


If Gen_DBType = "MDB" Then
   wc = " Where ExpenseFrom Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   wc = " Where ExpenseFrom Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSLocation.Text <> "" Then
   wc = wc & " AND Mst_Location.LocationName = '" & CmbSLocation.Text & "' "
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And Mst_Location." & Gen_WcLocation
   End If
End If

If CmbSEmployee.Text <> "" Then
   strSplitString = Split(CmbSEmployee, "~")
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
   wc = wc & " AND EMP.EmployeeNo = '" & EmpNo_ & "' "
'Else
'   Wc = Wc & " And EMP." & Gen_WcEmployeeID
Else
   If Gen_WcEmployeeID <> "" And LCase(Gen_UserRole) <> "power" And LCase(Gen_DesignationID) <> 1 Then
      wc = wc & " AND EMP." & Gen_WcEmployeeID
   End If
End If

If CmbSStatus.Text <> "" Then
   wc = wc & " AND Txn_ClaimSubmission.Flag = '" & mSStatus & "' "
Else
   wc = wc & " AND Txn_ClaimSubmission.Flag <> 'R' "
End If

If CmbSForwardTo.Text <> "" Then
   strSplitString = Split(CmbSForwardTo, "~")
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
   wc = wc & " AND EMP1.EmployeeNo = '" & EmpNo_ & "' "
'Else
'   If Gen_WcEmployeeID <> "" And LCase(Gen_UserRole) <> "power" Then
'      Wc = Wc & " AND EMP1." & Gen_WcEmployeeID
'   End If
End If

If CmbSExpenseType.Text <> "" Then
   wc = wc & " And Mst_ExpenseType.ExpenseType = '" & CmbSExpenseType.Text & "' "
End If

tmp = tmp & wc

Call TmpTable

Call dispdata



End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ExpenseListing.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ExpenseListing.xls")
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

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetReport.Enabled = False
   CmdExcel.Enabled = False
   Exit Sub
End If

Call CmdClear_Click

End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 18

MSHFlexGrid1.TextMatrix(0, 0) = "Claim ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Claim Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Expense From"
MSHFlexGrid1.TextMatrix(0, 3) = "Expense To"
MSHFlexGrid1.TextMatrix(0, 4) = "Employee Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 6) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 7) = "Total Amount"
MSHFlexGrid1.TextMatrix(0, 8) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 9) = "Expense Type"
MSHFlexGrid1.TextMatrix(0, 10) = "Particulars"
MSHFlexGrid1.TextMatrix(0, 11) = "Expense Amount"
MSHFlexGrid1.TextMatrix(0, 12) = "Recoverable Amount"
MSHFlexGrid1.TextMatrix(0, 13) = "EmployeeID"
MSHFlexGrid1.TextMatrix(0, 14) = "LocationID"
MSHFlexGrid1.TextMatrix(0, 15) = "Forwarded to"
MSHFlexGrid1.TextMatrix(0, 16) = "Status"
MSHFlexGrid1.TextMatrix(0, 17) = "Allocate Expenses"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 0
MSHFlexGrid1.ColWidth(7) = 800
MSHFlexGrid1.ColWidth(8) = 1800
MSHFlexGrid1.ColWidth(9) = 1800
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 0
MSHFlexGrid1.ColWidth(14) = 0
MSHFlexGrid1.ColWidth(15) = 2000
MSHFlexGrid1.ColWidth(16) = 1200
MSHFlexGrid1.ColWidth(17) = 1200

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True

End Sub
Public Sub dispdata()
Dim x As Double

x = 1

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(I, 0) = I
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
           If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
              If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
                 MSHFlexGrid1.TextMatrix(I, c) = Format(MSHFlexGrid1.TextMatrix(I, c), "dd-mmm-yyyy")
              Else
                 MSHFlexGrid1.TextMatrix(I, c) = MSHFlexGrid1.TextMatrix(I, c)
              End If
           Else
              MSHFlexGrid1.TextMatrix(I, c) = MSHFlexGrid1.TextMatrix(I, c)
           End If
       Next
       If LCase(MSHFlexGrid1.TextMatrix(I, 16)) = "p" Then
          MSHFlexGrid1.TextMatrix(I, 16) = "Pending"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 16)) = "a" Then
          MSHFlexGrid1.TextMatrix(I, 16) = "Approved"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 16)) = "r" Then
          MSHFlexGrid1.TextMatrix(I, 16) = "Rejected"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 16)) = "f" Then
          MSHFlexGrid1.TextMatrix(I, 16) = "Frozen"
       End If
       If MSHFlexGrid1.TextMatrix(I, 17) = "0" Then
          MSHFlexGrid1.TextMatrix(I, 17) = "No"
       Else
          MSHFlexGrid1.TextMatrix(I, 17) = "Yes"
       End If
       TmpRs.MoveNext
   Next
   Label2.Refresh
   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
   MsgBox "Done!!!"
Else
   Call Grid_Head
   MsgBox "No record found!!!"
End If

End Sub
Private Sub CmbSEmployee_GotFocus()
tmp = CmbSEmployee.Text
CmbSEmployee.Clear
Call GetSearchEmployeeDataonGotFocus
CmbSEmployee.Text = tmp
End Sub

Private Sub GetSearchEmployeeDataonGotFocus()

If Gen_WcEmployeeID <> "" And LCase(Gen_UserRole) <> "power" And LCase(Gen_DesignationID) <> 1 Then
   Call TableMst_Employee(" Where " & Gen_WcEmployeeID)  '("Where Flag = 'Y'")
Else
   Call TableMst_Employee '("Where Flag = 'Y'")
End If
CmbSEmployee.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSEmployee.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSEmployee.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next

End Sub

Private Sub CmbSEmployee_LostFocus()

Call GetSearchEmployeeDataOnLostFocus

End Sub


Private Sub CmbSForwardTo_GotFocus()
tmp = CmbSForwardTo.Text
'If Gen_WcEmployeeID <> "" And LCase(Gen_UserRole) <> "power" Then
'   Call TableMst_Employee("Where " & Gen_WcEmployeeID)  '("Where Flag = 'Y'")
'Else
   Call TableMst_Employee '("Where Flag = 'Y'")
'End If
CmbSForwardTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSForwardTo.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSForwardTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSForwardTo.Text = tmp

End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) = "approved" Then
   mSStatus = "A"
ElseIf LCase(CmbSStatus.Text) = "pending" Then
   mSStatus = "P"
ElseIf LCase(CmbSStatus.Text) = "rejected" Then
   mSStatus = "R"
ElseIf LCase(CmbSStatus.Text) = "frozen" Then
   mSStatus = "F"
Else
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationId = RsMst_Location!LocationID
End Sub


Private Sub GetSearchEmployeeDataOnLostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSEmployee.Text = "" Then
   mSEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbSEmployee, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSEmployee.SetFocus
   Exit Sub
End If
mSEmployeeID = RsMst_Employee!EmployeeID

End Sub


Private Sub CmbSForwardTo_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSForwardTo.Text = "" Then
   mSForwardToID = Null
   Exit Sub
End If

strSplitString = Split(CmbSForwardTo, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSForwardTo.SetFocus
   Exit Sub
End If
mSForwardToID = RsMst_Employee!EmployeeID

End Sub

Private Sub CmbSExpenseType_GotFocus()
Dim wc As String
tmp = CmbSExpenseType.Text
Call TableMst_ExpenseType(" Where AllocateTo = 'B' Or AllocateTo = 'E'")
CmbSExpenseType.Clear
If RsMst_ExpenseType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExpenseType.RecordCount
    CmbSExpenseType.AddItem RsMst_ExpenseType!ExpenseType
    RsMst_ExpenseType.MoveNext
Next
CmbSExpenseType.Text = tmp
End Sub
'----
Private Sub CmbSExpenseType_LostFocus()
If CmbSExpenseType.Text = "" Then
   mExpenseTypeID = Null
   Exit Sub
End If
RsMst_ExpenseType.MoveFirst
RsMst_ExpenseType.Find "ExpenseType = '" & CmbSExpenseType.Text & "'"
If RsMst_ExpenseType.EOF Then
   MsgBox "Invalid selection "
   CmbSExpenseType.SetFocus
   Exit Sub
End If
mExpenseTypeID = RsMst_ExpenseType!ExpenseTypeID
End Sub
