VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_EmployeeExpenseMIS 
   Caption         =   "Employee Expense MIS"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   15015
      Begin VB.OptionButton OptAll 
         Caption         =   "All"
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
         Left            =   3525
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton optEmp 
         Caption         =   "Employee"
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
         Left            =   5820
         TabIndex        =   15
         Top             =   615
         Width           =   1335
      End
      Begin VB.OptionButton optDesign 
         Caption         =   "Designation"
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
         Left            =   4215
         TabIndex        =   14
         Top             =   615
         Width           =   1695
      End
      Begin VB.ComboBox CmbEmpName 
         Height          =   315
         ItemData        =   "FrmRpt_EmployeeExpenseMIS.frx":0000
         Left            =   7200
         List            =   "FrmRpt_EmployeeExpenseMIS.frx":0002
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   3885
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
         Left            =   13590
         TabIndex        =   11
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
         Left            =   11160
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
         Left            =   12375
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   61210625
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   61210625
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Employee Name"
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
         Left            =   7200
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   3720
      End
      Begin VB.Label Label3 
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
         Left            =   13590
         TabIndex        =   10
         Top             =   255
         Width           =   90
      End
      Begin VB.Label Label2 
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
         Left            =   11190
         TabIndex        =   9
         Top             =   255
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Expense To  Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   8
         Top             =   135
         Width           =   1110
      End
      Begin VB.Label Label6 
         Caption         =   "Expense From  Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   495
         TabIndex        =   7
         Top             =   135
         Width           =   1110
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_EmployeeExpenseMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GunVer As String
Dim gmTotCap As Double
Dim TotExp As Double
Dim mDesID As Double
Dim mEmpID As Double
Private Sub Grid_Head()

tmp = "Select distinct ExpT.ExpenseTypeID,ExpT.ExpenseType from Txn_ClaimSubmissionDetail TCSD " & _
"Inner Join Mst_ExpenseType ExpT on TCSD.ExpenseTypeID  = ExpT.ExpenseTypeID " & _
"order by ExpT.ExpenseTypeID,ExpT.ExpenseType"
Call TmpTable

If TmpRs.RecordCount = 0 Then Exit Sub

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.Cols = TmpRs.RecordCount + 6


MSHFlexGrid1.TextMatrix(1, 0) = "Managment/RM"
MSHFlexGrid1.TextMatrix(1, 1) = "RM/AM"
'MSHFlexGrid1.TextMatrix(0, 2) = "State"
'MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(1, 2) = "EmployeeID"
MSHFlexGrid1.TextMatrix(1, 3) = "Employee"
MSHFlexGrid1.TextMatrix(1, 4) = "Designation"
MSHFlexGrid1.TextMatrix(1, 5) = "Total Expense"

Dim I As Variant
I = 6
For x = 1 To TmpRs.RecordCount
    MSHFlexGrid1.TextMatrix(1, I) = TmpRs!ExpenseType
    TmpRs.MoveNext
    I = I + 1
Next

MSHFlexGrid1.ColWidth(0) = 2200
MSHFlexGrid1.ColWidth(1) = 2200
MSHFlexGrid1.ColWidth(2) = 0
MSHFlexGrid1.ColWidth(3) = 2400
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
'MSHFlexGrid1.ColWidth(6) = 2000
'MSHFlexGrid1.ColWidth(7) = 1800
'MSHFlexGrid1.ColWidth(4) = 2600
'MSHFlexGrid1.ColWidth(5) = 2000
'MSHFlexGrid1.ColWidth(6) = 2200
'MSHFlexGrid1.ColWidth(7) = 2200
''MSHFlexGrid1.ColWidth(8) = 1600

End Sub

Private Sub CmbEmpName_GotFocus()

tmp = CmbEmpName.Text

If optEmp.Value = True Then
   If Gen_WcEmployeeID = "" Then
      Call TableMst_Employee
   Else
      Call TableMst_Employee(" Where " & Gen_WcEmployeeID)
   End If
   CmbEmpName.Clear
   If RsMst_Employee.RecordCount = 0 Then
      MsgBox "No Employee found"
      CmbEmpName.SetFocus
      Exit Sub
   End If
   For I = 1 To RsMst_Employee.RecordCount
       If RsMst_Employee!DesignationID <> 7 Then
          CmbEmpName.AddItem RsMst_Employee!EmployeeName
       End If
       RsMst_Employee.MoveNext
   Next
End If
If optDesign.Value = True Then

   Call TableMst_Designation
   CmbEmpName.Clear
   If RsMst_Designation.RecordCount = 0 Then
      MsgBox "No Designation found"
      CmbEmpName.SetFocus
      Exit Sub
   End If
   For I = 1 To RsMst_Designation.RecordCount
       If LCase(RsMst_Designation!Designation) <> "rag" Then
          CmbEmpName.AddItem RsMst_Designation!Designation
       End If
       RsMst_Designation.MoveNext
   Next
End If
   CmbEmpName.Text = tmp
   
   
End Sub

Private Sub CmbEmpName_LostFocus()
If CmbEmpName.Text = "" Then
   mEmpID = 0
   mDesID = 0
   Exit Sub
End If
If optDesign.Value = True Then
   RsMst_Designation.MoveFirst
   RsMst_Designation.Find "Designation = '" & CmbEmpName.Text & "'"
   If RsMst_Designation.EOF Then
      MsgBox "Invalid Designation selection "
      CmbLocationID.SetFocus
      Exit Sub
   End If
   mDesID = RsMst_Designation!DesignationID
End If
If optEmp.Value = True Then
   RsMst_Employee.MoveFirst
   RsMst_Employee.Find "EmployeeName = '" & CmbEmpName.Text & "'"
   If RsMst_Employee.EOF Then
      MsgBox "Invalid Employee Name selection "
      CmbEmpName.SetFocus
      Exit Sub
   End If
   mEmpID = RsMst_Employee!EmployeeID
End If

End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmdExcel.Enabled = False
CmbEmpName.Visible = False
Label4.Visible = False
CmbEmpName.Clear
OptAll.Value = True
Label3.Caption = ""
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeExpenseMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeExpenseMIS.xls")
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


Private Sub CmdGetReport_Click()
Dim mCol As Integer
Dim mTotExp As Double

If TxtFromDate > TxtToDate Then
   MsgBox "FromDate Should be Less Than Todate."
   TxtFromDate.SetFocus
   Exit Sub
End If
If optDesign.Value = True Then
   If CmbEmpName.Text = "" Then
      MsgBox " Please Select Designation."
      Exit Sub
   End If
End If
If optEmp.Value = True Then
   If CmbEmpName.Text = "" Then
      MsgBox " Please Select Employee Name."
      Exit Sub
   End If
End If
Label3.Caption = ""
Call Grid_Head

MsgBox "Wait till next msg"
'
'tmp = "Select EMP2.EmployeeName as 'AM/RM' "
'tmp = tmp & " ,EMP1.EmployeeName 'CC/AM'"
'tmp = tmp & " ,ME.EmployeeID,ME.EmployeeName + '~' + ME.EmployeeNo as Employee,MD.Designation"
'tmp = tmp & " From Mst_Employee ME"
'tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP1 On ME.DirectReportingTo = EMP1.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
'tmp = tmp & " Inner Join Mst_Designation MD1 On EMP1.DesignationID = MD1.DesignationID"
'tmp = tmp & " Inner Join Mst_Designation MD2 On EMP2.DesignationID = MD2.DesignationID"

tmp = "Select EMP2.EmployeeName ,EMP1.EmployeeName,EMP.EmployeeID,EMP.EmployeeName + ' ~ ' +EMP.EmployeeNo as 'Employee', "
tmp = tmp & " MD.Designation"
tmp = tmp & " From Mst_Employee EMP"
tmp = tmp & " Left Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID "
tmp = tmp & " Left Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID "
tmp = tmp & " Left Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID "
tmp = tmp & " Inner Join Mst_Designation MD On EMP.DesignationID = MD.DesignationID "
tmp = tmp & " Where MD.Designation <> 'RAG' "

If optDesign.Value = True Then
   tmp = tmp & " And MD.DesignationID=" & mDesID & " "
End If

If optEmp.Value = True Then
   tmp = tmp & " And EMP.EmployeeID=" & mEmpID & " "
Else
   If Gen_WcEmployeeID <> "" Then
      tmp1 = Replace(Gen_WcEmployeeID, "EmployeeID", "EMP.EmployeeID")
      tmp = tmp & " And " & tmp1
   End If
End If

Call TmpTable

If TmpRs.RecordCount > 0 Then
   x = 2
   
   tmp = "Select ME.EmployeeID,ExpT.ExpenseType , Sum(TCSD.Amount) As TotalExpenseAmount"
   tmp = tmp & " From  Txn_ClaimSubmission TCS"
   tmp = tmp & " Inner Join Mst_Employee ME On TCS.EmployeeID = ME.EmployeeID"
   tmp = tmp & " Inner Join Txn_ClaimSubmissionDetail TCSD on TCS.ClaimID  = TCSD.ClaimID  And TCS.SM_Prefix = TCSD.SM_Prefix"
   tmp = tmp & " Inner Join Mst_ExpenseType ExpT on TCSD.ExpenseTypeID  = ExpT.ExpenseTypeID"
   'tmp = tmp & " Where TCS.ClaimDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TCS.ClaimDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
   tmp = tmp & " Where TCS.ExpenseFrom >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TCS.ExpenseFrom <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
   tmp = tmp & " And TCS.FLag <> 'R' "
   tmp = tmp & " Group By ME.EmployeeID,ExpT.ExpenseType"
   tmp = tmp & " order by ExpT.ExpenseType"
       
   Call Tmp1Table
   
   For I = 1 To TmpRs.RecordCount
       mTotExp = 0
       '- 1

       For c = 0 To 4
           MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       TmpRs1.Filter = "EmployeeID= " & Val(MSHFlexGrid1.TextMatrix(x, 2)) & "  "
       If TmpRs1.RecordCount > 0 Then
          For r = 1 To TmpRs1.RecordCount
              mCol = GetExpenseCol(TmpRs1!ExpenseType)
              If mCol <> 0 Then
                 MSHFlexGrid1.TextMatrix(x, mCol) = IIf(IsNull(TmpRs1!TotalExpenseAmount), "0", TmpRs1!TotalExpenseAmount)
                 MSHFlexGrid1.TextMatrix(0, mCol) = Val(MSHFlexGrid1.TextMatrix(0, mCol)) + Val(MSHFlexGrid1.TextMatrix(x, mCol))
                 mTotExp = mTotExp + Val(MSHFlexGrid1.TextMatrix(x, mCol))
              End If
              TmpRs1.MoveNext
          Next
       End If
       TmpRs1.Filter = ""
       MSHFlexGrid1.TextMatrix(x, 5) = mTotExp
       MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 5))
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       x = x + 1
       TmpRs.MoveNext
   Next
   CmdExcel.Enabled = True
End If
Label3.Caption = TmpRs.RecordCount
Label3.Refresh
MsgBox "Done!!!"

End Sub
Private Function GetExpenseCol(mExpType As Variant) As Integer
Dim mCol As Variant
mCol = 0
For o = 6 To MSHFlexGrid1.Cols - 1
    If mExpType = MSHFlexGrid1.TextMatrix(1, o) Then
       mCol = o
       Exit For
    End If
Next
GetExpenseCol = mCol
End Function
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
Call Grid_Head
CmdExcel.Enabled = False
TxtFromDate.Value = Date - 365
TxtToDate.Value = Date
End Sub

Private Sub OptAll_Click()

Call optDesign_Click


End Sub

Private Sub optDesign_Click()

CmbEmpName.Visible = True
Label4.Visible = True
CmbEmpName.Clear

If optDesign.Value = True Then
   Label4.Caption = "Designation Name"
ElseIf optEmp.Value = True Then
   Label4.Caption = "Employee Name"
End If

If OptAll.Value = True Then
   CmbEmpName.Visible = False
   Label4.Visible = False
'   CmbEmpName.Clear
'   Label4.Caption = "Designation Name"
    
End If
End Sub

Private Sub optEmp_Click()

Call optDesign_Click
End Sub
