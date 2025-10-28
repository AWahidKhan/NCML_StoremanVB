VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_ImportMonthWiseSalary 
   Caption         =   "Import Monthwise Salary"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   150
      TabIndex        =   8
      Top             =   795
      Width           =   14775
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   14775
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
         ItemData        =   "Frm_ImportMonthWiseSalary.frx":0000
         Left            =   120
         List            =   "Frm_ImportMonthWiseSalary.frx":0002
         TabIndex        =   11
         Top             =   375
         Width           =   2415
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
         ItemData        =   "Frm_ImportMonthWiseSalary.frx":0004
         Left            =   2595
         List            =   "Frm_ImportMonthWiseSalary.frx":0006
         TabIndex        =   10
         Top             =   375
         Width           =   1500
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
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
         Height          =   350
         Left            =   5820
         TabIndex        =   5
         Top             =   375
         Width           =   1620
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Height          =   350
         Left            =   10680
         TabIndex        =   4
         Top             =   375
         Width           =   1620
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
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
         Height          =   350
         Left            =   9060
         TabIndex        =   3
         Top             =   375
         Width           =   1620
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
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
         Height          =   350
         Left            =   7440
         TabIndex        =   2
         Top             =   375
         Width           =   1620
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4200
         TabIndex        =   1
         Top             =   375
         Width           =   1620
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
         Left            =   1020
         TabIndex        =   13
         Top             =   135
         Width           =   600
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
         TabIndex        =   12
         Top             =   120
         Width           =   495
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   150
      TabIndex        =   6
      Top             =   1425
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7935
      Left            =   150
      TabIndex        =   7
      Top             =   1800
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   13996
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_ImportMonthWiseSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth, mFromDate, mMonthListindex, mSM_Prefix, mTotalDays, mToDate, mNoofdays As Variant
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim ErrMsg As Variant
Dim Status As Variant
Dim WrkSheet As String
Dim RowCount As Double

Private Sub CmdExit_Click()
On Error Resume Next
oXL.Quit
Unload Me
End Sub

Private Sub CmdReset_Click()
Call Grid_Head
'CmdImport.Enabled = True
CmdExit.Enabled = True
CmdSave.Enabled = False
CmdGetfile.Enabled = True
CmdImport.Enabled = False
End Sub
Private Sub CmdGetFile_Click()
'ListView1.ListItems.Clear
On Error GoTo errortrap
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

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls)|*.xls"
     .CancelError = True
procReOpen:
     .ShowOpen
End With
Path = CommonDialog1.FileName
lblFile.Caption = Path
'CmdUpload.Enabled = True
CmdImport.Enabled = True
CmdReset.Enabled = False
CmdExit.Enabled = False
CmdSave.Enabled = False
  
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"

End Sub
Private Sub UploadFile()
Status = 0
If lblFile.Caption <> "" And LCase(lblFile.Caption) <> "filename" Then
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(lblFile.Caption)
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   CmdExit.Enabled = True
   Status = 0
Else
   MsgBox "Please Select File", vbInformation, "Information"
   Status = 1
End If
End Sub
Private Sub CmdImport_Click()
Dim mulCheck As Integer
Dim isSelected As Boolean

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

Call UploadFile
If Status = 1 Then Exit Sub
   'CmdImport.Enabled = False
   'cmdExit.Enabled = False
   'CmdReset.Enabled = False
   CmdGetfile.Enabled = False
   
   '-----------------------------
   ErrMsg = ""
   Status = 0
   RowCount = 0
   I = 1
   '-----------------------------
   Call Validation
   If Status = 1 Then
       MsgBox ErrMsg, vbExclamation, "Message"
       CmdGetfile.Enabled = True
   Else
       Call Grid_Head
       Call FillData
       CmdSave.Enabled = True
       CmdReset.Enabled = True
   End If
End Sub
Private Sub Validation()

   Set oWS = oWB.Worksheets("Sheet1")
   WrkSheet = "Sheet1"
   If LCase(Trim(oWS.Cells(1, 1))) <> "sr. no" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A(Sr. No) " & vbCrLf
      Status = 1
   End If
   If LCase(Trim(oWS.Cells(1, 2))) <> "year month" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of B(Year Month) " & vbCrLf
      Status = 1
   End If
   If LCase(Trim(oWS.Cells(1, 3))) <> "employee code" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of C(Employee Code) " & vbCrLf
      Status = 1
   End If
   If LCase(Trim(oWS.Cells(1, 4))) <> "employee name" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of D(Employee Name) " & vbCrLf
      Status = 1
   End If
   If UCase(Trim(oWS.Cells(1, 5))) <> "CTC-MONTHLY" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of E(CTC-monthly) " & vbCrLf
      Status = 1
   End If
   If UCase(Trim(oWS.Cells(1, 6))) <> "GROSS-MONTHLY" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of F(Gross-Monthly) " & vbCrLf
      Status = 1
   End If
' Check Column Name In Excel Sheet
''Set oWS = oWB.Worksheets("Sheet1")
''WrkSheet = "Sheet1"
''If LCase(Trim(oWS.Cells(1, 1))) <> "txn id" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of A='Txn ID'" & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 2))) <> "txn date" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of B='Txn Date'" & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 3))) <> "sm's employee id" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of C='SM's Employee Id' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 4))) <> "hr employee id" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of D='HR Employee Id' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 5))) <> "employee name" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of E='Employee Name' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 6))) <> "designation" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of F='Designation' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 7))) <> "direct repotee" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of G='Direct Repotee' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 8))) <> "connected drc" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of H='Connected DRC' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 9))) <> "connected rag" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of I='Connected RAG' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 10))) <> "email id" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of J='Email ID' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 11))) <> "mobile no" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of K='Mobile No' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 12))) <> "address" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of L='Address' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 13))) <> "base location" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of M='Base Location' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 14))) <> "status" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of N='Status' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 15))) <> "self reporting" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of O='Self Reporting' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 16))) <> "allocation expense" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of P='Allocation Expense' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 17))) <> "cmp allocation" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of Q='CMP Allocation' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 18))) <> "ctc monthly" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of R='CTC Monthly' " & vbCrLf
''   Status = 1
''End If
''If LCase(Trim(oWS.Cells(1, 19))) <> "gross pay" Then
''   ErrMsg = ErrMsg + "Please Set Column Name of S='Gross Pay' " & vbCrLf
''   Status = 1
''End If
''If CDate(Trim(oWS.Cells(2, 2))) <> CDate(mToDate) Then
''   ErrMsg = ErrMsg + "Please Check Txn Date" & vbCrLf
''   Status = 1
''End If
End Sub
Private Sub FillData()
Dim intNumberRows As Double

Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
RowCount = oWB.Worksheets(1).UsedRange.Rows.Count
ProgressBar1.Value = 0
ProgressBar1.Max = RowCount

Do
   If oWS.Cells(mRow, 1) = "" Then
      'ProgressBar1.Value = ProgressBar1.Value + 1
      Exit Do
   End If
  
   For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(mRow, c) = Trim(oWS.Cells(mRow + 1, c + 1).Value)
   Next
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   ProgressBar1.Value = ProgressBar1.Value + 1
   mRow = mRow + 1
Loop Until False

oXL.Quit
MsgBox "Data Imported Successfully ", vbInformation
CmdImport.Enabled = False
ProgressBar1.Value = 0
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.FixedRows = 1

With MSHFlexGrid1
    .Rows = 2
    .Cols = 6
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Employee ID"
    .TextMatrix(0, 1) = "Year Month"
    .TextMatrix(0, 2) = "Employee Code"
    .TextMatrix(0, 3) = "Employee Name"
    .TextMatrix(0, 4) = "CTC-MONTHLY"
    .TextMatrix(0, 5) = "CTC-MONTHLY"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 1600
    .ColWidth(2) = 1800
    .ColWidth(3) = 2800
    .ColWidth(4) = 1800
    .ColWidth(5) = 1800
End With
End Sub

'Private Sub Grid_Head()
'
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Rows = 2
'MSHFlexGrid1.Cols = 19
'
'MSHFlexGrid1.TextMatrix(0, 0) = "Txn ID"
'MSHFlexGrid1.TextMatrix(0, 1) = "Txn Date"
'MSHFlexGrid1.TextMatrix(0, 2) = "SM's Employee Id"
'MSHFlexGrid1.TextMatrix(0, 3) = "HR Employee Id"
'MSHFlexGrid1.TextMatrix(0, 4) = "Employee Name"
'MSHFlexGrid1.TextMatrix(0, 5) = "Designation"
'MSHFlexGrid1.TextMatrix(0, 6) = "Direct Repotee"
'MSHFlexGrid1.TextMatrix(0, 7) = "Connected DRC"
'MSHFlexGrid1.TextMatrix(0, 8) = "Connected RAG"
'MSHFlexGrid1.TextMatrix(0, 9) = "Email ID"
'MSHFlexGrid1.TextMatrix(0, 10) = "Mobile No"
'MSHFlexGrid1.TextMatrix(0, 11) = "Address"
'MSHFlexGrid1.TextMatrix(0, 12) = "Base Location"
'MSHFlexGrid1.TextMatrix(0, 13) = "Status"
'MSHFlexGrid1.TextMatrix(0, 14) = "Self Reporting"
'MSHFlexGrid1.TextMatrix(0, 15) = "Allocation Expense"
'MSHFlexGrid1.TextMatrix(0, 16) = "CMP Allocation"
'MSHFlexGrid1.TextMatrix(0, 17) = "CTC"
'MSHFlexGrid1.TextMatrix(0, 18) = "Gross Pay"
'
'MSHFlexGrid1.ColWidth(0) = 1000
'MSHFlexGrid1.ColWidth(1) = 1600
'MSHFlexGrid1.ColWidth(2) = 1200
'MSHFlexGrid1.ColWidth(3) = 1400
'MSHFlexGrid1.ColWidth(4) = 3000
'MSHFlexGrid1.ColWidth(5) = 1200
'MSHFlexGrid1.ColWidth(6) = 1800
'MSHFlexGrid1.ColWidth(7) = 1800
'MSHFlexGrid1.ColWidth(8) = 1800
'MSHFlexGrid1.ColWidth(9) = 1400
'MSHFlexGrid1.ColWidth(0) = 1000
'MSHFlexGrid1.ColWidth(11) = 2000
'MSHFlexGrid1.ColWidth(12) = 1000
'MSHFlexGrid1.ColWidth(13) = 1000
'MSHFlexGrid1.ColWidth(14) = 1000
'MSHFlexGrid1.ColWidth(15) = 1000
'MSHFlexGrid1.ColWidth(16) = 1000
'MSHFlexGrid1.ColWidth(17) = 1200
'MSHFlexGrid1.ColWidth(18) = 1200
'
'End Sub

Private Sub CmdSave_Click()
Dim vAddNew As Variant
ErrMsg = ""
On Error GoTo errortrap
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows - 1

Call TableTxn_MonthEmp_History("Where TxnDate='" & Format(mToDate, Gen_DatFormat) & "' Order by MonthEmpID ")
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 2) <> "" And (Val(MSHFlexGrid1.TextMatrix(I, 4)) > 0 Or Val(MSHFlexGrid1.TextMatrix(I, 5)) > 0) Then
       If mToDate = CDate(MSHFlexGrid1.TextMatrix(I, 1)) Then
          If RsTxn_MonthEmp_History.RecordCount > 0 Then RsTxn_MonthEmp_History.MoveFirst
          RsTxn_MonthEmp_History.Find "EmployeeNo='" & Trim(MSHFlexGrid1.TextMatrix(I, 2)) & "'"
          If RsTxn_MonthEmp_History.EOF = False Then
             RsTxn_MonthEmp_History!CTC = Val(MSHFlexGrid1.TextMatrix(I, 4))
             RsTxn_MonthEmp_History!GrossPay = Val(MSHFlexGrid1.TextMatrix(I, 5))
             RsTxn_MonthEmp_History.Update
          End If
       Else
          MSHFlexGrid1.Row = I
          MSHFlexGrid1.Col = 0
          MSHFlexGrid1.CellBackColor = vbRed
       End If
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
MsgBox "Data Saved Successfully ", vbInformation
CmdSave.Enabled = False
CmdImport.Enabled = False
CmdGetfile.Enabled = True
ProgressBar1.Value = 0
Exit Sub
errortrap:
MsgBox "System Error Occured!!!", vbExclamation
End Sub
Private Sub Form_Load()
'LblFile.Caption = ""
'Frame1.Enabled = False
'Frame2.Enabled = False
'Frame3.Enabled = False
'CmdSave.Enabled = False
'
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Call CmdReset_Click

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

