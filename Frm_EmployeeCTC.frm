VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_EmployeeCTC 
   Caption         =   "Employee CTC"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   14775
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
         Left            =   360
         TabIndex        =   6
         Top             =   240
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
         Left            =   5100
         TabIndex        =   5
         Top             =   240
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
         Left            =   3540
         TabIndex        =   4
         Top             =   240
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
         Left            =   6720
         TabIndex        =   3
         Top             =   240
         Width           =   1620
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
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1620
      End
      Begin VB.CommandButton CmdUpload 
         Caption         =   "Upload"
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
         Left            =   12540
         TabIndex        =   1
         Top             =   240
         Width           =   2100
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2985
      TabIndex        =   15
      Top             =   870
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6255
      Left            =   60
      TabIndex        =   14
      Top             =   3465
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   11033
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
   Begin VB.ComboBox CmbSheetName 
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
      Left            =   12600
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "List of Sheets "
      Height          =   255
      Left            =   12600
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
      Begin MSComctlLib.ListView ListView1 
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         View            =   2
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Error Details"
      Height          =   1935
      Left            =   105
      TabIndex        =   9
      Top             =   1485
      Width           =   14775
      Begin VB.TextBox TxtErrorMsg 
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   240
         Width           =   14535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   90
      TabIndex        =   7
      Top             =   750
      Width           =   14775
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   660
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_EmployeeCTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
ListView1.ListItems.Clear
Call Grid_Head
CmdUpload.Enabled = True
'CmdImport.Enabled = True
CmdExit.Enabled = True
CmdSave.Enabled = False
CmdGetfile.Enabled = True
CmdImport.Enabled = False
TxtErrorMsg.Text = ""
End Sub
Private Sub CmdGetFile_Click()
'ListView1.ListItems.Clear
TxtErrorMsg.Text = ""
On Error GoTo errortrap
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
TxtErrorMsg.Text = ""
If lblFile.Caption <> "" And LCase(lblFile.Caption) <> "filename" Then
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(lblFile.Caption)
   ' Code For Ceating List of Sheet in Excel Sheet
'   vTemp = Split(LblFile.Caption, "\")
'   ReDim FName(UBound(vTemp))
'   Namef = vTemp(UBound(vTemp))
   Dim I As Integer
   For I = 1 To oWB.Worksheets.Count
       Set oWS = oWB.Worksheets(I)
       If I = 1 Then
          'CmbSheetName.AddItem oWS.Name
          ListView1.ListItems.Add I, , oWS.Name
       End If
   Next
   ListView1.ListItems.Item(1).Checked = True
   If Status = 1 Then
      ListView1.ListItems.Clear
      MsgBox ErrMsg, vbExclamation
   End If

   CmdUpload.Enabled = False
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   CmdExit.Enabled = True
   Status = 0
   'oXL.Quit

Else
   MsgBox "Please Select File", vbInformation, "Information"
   Status = 1
End If

End Sub
Private Sub CmdImport_Click()
Dim mulCheck As Integer

Dim isSelected As Boolean
Call UploadFile
If Status = 1 Then Exit Sub
TxtErrorMsg.Text = ""
isSelected = False
For j = 1 To ListView1.ListItems.Count
    ErrMsg = ListView1.ListItems.Item(j).Text
    If ListView1.ListItems.Item(j).Checked Then
       isSelected = True
       mulCheck = mulCheck + 1
    End If
Next j
If isSelected = True Then
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
   'Set oXL = New Excel.Application
   
   'Set oWB = oXL.Workbooks.Open(LblFile.Caption)
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
   'CmdImport.Enabled = True

Else
   MsgBox "Please Select WorkSheet To Import !...", vbInformation
End If
End Sub
Private Sub Validation()
For j = 1 To ListView1.ListItems.Count
' Check Column Name In Excel Sheet
    If ListView1.ListItems.Item(j).Checked Then
       If LCase(ListView1.ListItems.Item(j).Text) = "sheet1" Then
          Set oWS = oWB.Worksheets(ListView1.ListItems.Item(j).Text)
          WrkSheet = "Sheet1"
          If LCase(Trim(oWS.Cells(1, 1))) <> "category" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 2))) <> "employee code" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of B " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 3))) <> "employee name" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of C " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 4))) <> "department" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of D " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 5))) <> "professional fees _ monthly" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of E " & vbCrLf
             Status = 1
          End If
          If UCase(Trim(oWS.Cells(1, 6))) <> "CTC-YEARLY" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of F " & vbCrLf
             Status = 1
          End If
          If UCase(Trim(oWS.Cells(1, 7))) <> "CTC-MONTHLY" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of G " & vbCrLf
             Status = 1
          End If
          
       Else
          ErrMsg = ErrMsg + "Please Select Sheet1 "
          Status = 1
       End If
    End If
Next
End Sub
Private Sub FillData()
Dim intNumberRows As Double
For j = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(j).Checked Then
       Set oWS = oWB.Worksheets(ListView1.ListItems.Item(j).Text)
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
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          For c = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(mRow, c) = Trim(oWS.Cells(mRow + 1, c + 1).Value)
          Next
          
          ProgressBar1.Value = ProgressBar1.Value + 1
          mRow = mRow + 1
       Loop Until False
       
       oXL.Quit
    
    End If
Next
ProgressBar1.Value = 0
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.FixedRows = 1

With MSHFlexGrid1
    .Rows = 2
    .Cols = 7
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Category"
    .TextMatrix(0, 1) = "Employee Code"
    .TextMatrix(0, 2) = "Employee Name"
    .TextMatrix(0, 3) = "Department"
    .TextMatrix(0, 4) = "Professional Fees _ Monthly "
    .TextMatrix(0, 5) = "CTC-YEARLY"
    .TextMatrix(0, 6) = "CTC-MONTHLY"
    .ColWidth(0) = 1500
    .ColWidth(1) = 1600
    .ColWidth(2) = 3000
    .ColWidth(3) = 1800
    .ColWidth(4) = 2300
    .ColWidth(5) = 1800
    .ColWidth(6) = 1800
End With
 

End Sub

Private Sub CmdSave_Click()
Dim vAddNew As Variant
Call TableTxn_EmpCTC
ErrMsg = ""
On Error GoTo errortrap
Set oXL = New Excel.Application
Set oWB = oXL.Workbooks.Open(lblFile.Caption)
Set oWS = oWB.Worksheets(WrkSheet)
oWS.ClearArrows
mRow = 2
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows - 2
If WrkSheet = "Sheet1" Then
   Status = 0
   ErrMsg = ""
   Do
     vAddNew = 0
     If oWS.Cells(mRow, 1) = "" Then
     ProgressBar1.Value = ProgressBar1.Value + 1
         Exit Do
     End If
     
     If RsTxn_EmpCTC.RecordCount > 0 Then
        RsTxn_EmpCTC.MoveFirst
     End If
     
     RsTxn_EmpCTC.Find "EmployeeCode = '" & Trim(oWS.Cells(mRow, 2).Value) & " ' "
     If RsTxn_EmpCTC.EOF Then
        RsTxn_EmpCTC.AddNew
        vAddNew = 1
     Else
        RsTxn_EmpCTC.Find "EmployeeCode = '" & Trim(oWS.Cells(mRow, 2).Value) & " ' "
        'ErrMsg = ErrMsg + "Data Over Write of Employee Code " & Trim(oWS.Cells(mRow, 2).Value) & vbCrLf
     End If
     
     If Trim(oWS.Cells(mRow, 1).Value) <> "" Then
        RsTxn_EmpCTC!Category = Trim(oWS.Cells(mRow, 1).Value)
     Else
        ErrMsg = ErrMsg + "Please Check Data of Row " & mRow & vbCrLf
        Status = 1
     End If
                             
     If Trim(oWS.Cells(mRow, 2).Value) <> "" Then
        RsTxn_EmpCTC!EmployeeCode = Trim(oWS.Cells(mRow, 2).Value)
     Else
        ErrMsg = ErrMsg + "Please Check Data of Row " & mRow & vbCrLf
        Status = 1
     End If
                     
     If Trim(oWS.Cells(mRow, 4).Value) <> "" Then
        RsTxn_EmpCTC!Department = Trim(oWS.Cells(mRow, 4).Value)
     Else
        ErrMsg = ErrMsg + "Please Check Data of Row " & mRow & vbCrLf
        Status = 1
     End If
                  
     If Trim(oWS.Cells(mRow, 5).Value) <> "" Then
        RsTxn_EmpCTC!ProfessionalFees_Monthly = Trim(oWS.Cells(mRow, 5).Value)
     Else
        ErrMsg = ErrMsg + "Please Check Data of Row " & mRow & vbCrLf
        Status = 1
     End If
     
     If Trim(oWS.Cells(mRow, 6).Value) <> "" Then
        RsTxn_EmpCTC!CTC_YEARLY = Trim(oWS.Cells(mRow, 6).Value)
     Else
        ErrMsg = ErrMsg + "Please Check Data of Row " & mRow & vbCrLf
        Status = 1
     End If
     
     If Trim(oWS.Cells(mRow, 7).Value) <> "" Then
        RsTxn_EmpCTC!CTC_Monthly = Trim(oWS.Cells(mRow, 7).Value)
     Else
        ErrMsg = ErrMsg + "Please Check Data of Row " & mRow & vbCrLf
        Status = 1
     End If
     
     If Status = 0 Then
        RsTxn_EmpCTC.Update
        If vAddNew = 0 Then
           ErrMsg = ErrMsg + "Data Over Write of Employee Code " & Trim(oWS.Cells(mRow, 2).Value) & vbCrLf
        End If
     Else
        RsTxn_EmpCTC.CancelUpdate
     End If
     
     ProgressBar1.Value = ProgressBar1.Value + 1
     mRow = mRow + 1
        
   Loop Until False
End If
TxtErrorMsg.Text = ErrMsg
MsgBox "Data Saved To DataBase Successfully except shown in Error details......", vbInformation
CmdSave.Enabled = False
CmdImport.Enabled = False
CmdGetfile.Enabled = True
oXL.Quit
ProgressBar1.Value = 0
Exit Sub
errortrap:
    MsgBox "System Error Occured!!!", vbExclamation
oXL.Quit
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Index <> 1 Then
   Item.Checked = False
End If
End Sub
Private Sub Form_Load()
'LblFile.Caption = ""
'Frame1.Enabled = False
'Frame2.Enabled = False
'Frame3.Enabled = False
'CmdSave.Enabled = False
'
'Dim FrmLoadAccess() As Boolean
'FrmLoadAccess = GetFrmLoadAccess(Me.Name)
'If FrmLoadAccess(0) = False Then
'   MsgBox "Access denied !!!!!!!!!"
'   Exit Sub
'End If
'
'Frame1.Enabled = True
'Frame2.Enabled = True
'Frame3.Enabled = True
'CmdSave.Enabled = False
'CmdImport.Enabled = False
'CmdExit.Enabled = False
'CmdReset.Enabled = False

End Sub


