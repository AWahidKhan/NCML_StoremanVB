VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_EmployeeLedger 
   Caption         =   "Employee Ledger"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   5
      Top             =   945
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   6
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
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   9510
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CmbEmployee 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   8295
         TabIndex        =   2
         Top             =   233
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         Height          =   375
         Left            =   7080
         TabIndex        =   1
         Top             =   233
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   293
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "_"
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
         TabIndex        =   4
         Top             =   300
         Width           =   135
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
         TabIndex        =   3
         Top             =   300
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRpt_EmployeeLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSEmployeeID As Variant
Private Sub CmbEmployee_GotFocus()

tmp = CmbEmployee.Text
Call TableMst_Employee("Where " & Gen_WcEmployeeID)
CmbEmployee.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbEmployee.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Employee.RecordCount
    CmbEmployee.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next

CmbEmployee.Text = tmp

End Sub

Private Sub CmbEmployee_LostFocus()

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbEmployee.Text = "" Then
   mSEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbEmployee, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbEmployee.SetFocus
   Exit Sub
End If

mSEmployeeID = RsMst_Employee!EmployeeID

End Sub

Private Sub CmdClear_Click()

Call Grid_Head
CmdExcel.Enabled = False
CmbEmployee.Text = ""
Label2.Caption = ""

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

Frame1.Enabled = True
Call CmdClear_Click
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 6

MSHFlexGrid1.TextMatrix(0, 0) = "Document Type"
MSHFlexGrid1.TextMatrix(0, 1) = "Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Narration"
MSHFlexGrid1.TextMatrix(0, 3) = "Amount"
MSHFlexGrid1.TextMatrix(0, 4) = "Entry No"
MSHFlexGrid1.TextMatrix(0, 5) = "Employee Name"

MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 7600
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 2000


End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeLedger.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeLedger.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim tmp1 As String
Dim Wc As String
Wc = ""

Call Grid_Head

If CmbEmployee.Text = "" Then
   MsgBox "Please select atleast one employee!!!"
   CmbEmployee.SetFocus
   Exit Sub
End If

tmp = "Select MEL.DocumentType,MEL.AcutalPostingDate,MEL.Narration,MEL.TransactionAmount,MEL.EntryNo,ME.EmployeeName "
tmp = tmp & " From  Mst_EmpLedger MEL"
tmp = tmp & " Inner Join Mst_Employee ME On MEL.EmployeeID = ME.EmployeeID "
'If CmbEmployee.Text = "" Then
   'tmp = tmp & " Where  & " "
'   tmp1 = Replace(Gen_WcEmployeeID, "EmployeeID", "MEL.EmployeeID")
'   Wc = GenWc(Wc, "", tmp1, "N", "I")
'Else
tmp = tmp & " Where MEL.EmployeeID = " & mSEmployeeID & ""
'End If



tmp1 = " Order By ME.EmployeeName,MEL.AcutalPostingDate"
tmp = tmp & Wc & tmp1 '" Order By ME.EmployeeName"

Call TmpTable

Call DispData

End Sub

Public Sub DispData()
Dim x As Double
MsgBox ("Wait till Next Message")
x = 1

If TmpRs.RecordCount > 0 Then
   For i = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       'MSHFlexGrid1.TextMatrix(i, 0) = i
    
       MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(TmpRs.Fields(0)), "", TmpRs.Fields(0))
       MSHFlexGrid1.TextMatrix(i, 1) = Format(TmpRs.Fields(1), "dd-MMM-yyyy") 'IIf(IsNull(TmpRs.Fields(1)), "", TmpRs.Fields(1))
       MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(TmpRs.Fields(2)), "", TmpRs.Fields(2))
       MSHFlexGrid1.TextMatrix(i, 3) = Format(TmpRs.Fields(3), "########0.00") 'IIf(IsNull(TmpRs.Fields(3)), "", TmpRs.Fields(3))
       MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(TmpRs.Fields(4)), "", TmpRs.Fields(4))
       MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(TmpRs.Fields(5)), "", TmpRs.Fields(5))
       
       TmpRs.MoveNext
   Next
   MsgBox "Done!!!"
   CmdExcel.Enabled = True
Else
   MsgBox "No record found!!!"
   Call Grid_Head
End If

Label2.Refresh
Label2.Caption = TmpRs.RecordCount

End Sub

