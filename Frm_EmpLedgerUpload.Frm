VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_EmpLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Ledger File Import"
   ClientHeight    =   6405
   ClientLeft      =   150
   ClientTop       =   330
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   9780
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   15060
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6435
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetfile 
         Caption         =   "Get File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
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
         Left            =   10560
         TabIndex        =   18
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "NAV Code not mapped in Employee Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11205
         TabIndex        =   17
         Top             =   1770
         Width           =   3735
      End
      Begin VB.Label Label10 
         Caption         =   "Entry No Already Exist "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   16
         Top             =   1770
         Width           =   1965
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
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
         Left            =   7560
         TabIndex        =   15
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lines :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblLines 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   9
         Top             =   1680
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   435
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   1
         Top             =   1320
         Width           =   60
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   4920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6840
      Left            =   120
      TabIndex        =   11
      Top             =   2340
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   12065
      _Version        =   393216
      WordWrap        =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   9240
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "Frm_EmpLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstemp As New ADODB.Recordset
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream

Dim objFScp As Scripting.FileSystemObject
Dim objFilecp As File
Dim objTScp As TextStream

Dim sFileLine As String
Dim sFileLineNo As Long

Dim temp As Variant
Dim arecordset As String
Dim strFilename As String
Dim strsql As String
Dim strdata As String
Dim t_line As String
Dim strfile As String
Dim strqry As String
Dim sArray(0 To 10) As String
Dim mEmployeeID As Variant

Private Sub CmdCancel_Click()
lblFile.Caption = ""
lblStatus.Caption = ""
lblLines.Caption = ""
strFilename = ""
CmdGetfile.Enabled = True
CmdUpdate.Enabled = False
CmdClose.Enabled = True
CmdCancel.Enabled = True
CmdSave.Enabled = False
CmdExcel.Enabled = False
Call Grid_Head
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub CmdGetFile_Click()
    
CmdGetfile.Enabled = False
CmdUpdate.Enabled = False
CmdClose.Enabled = False
CmdCancel.Enabled = True

lblStatus.Refresh
lblStatus.Caption = ""
lblFile.Refresh
lblFile.Caption = ""


With cdialog
   .InitDir = App.Path
   .DialogTitle = "Select Employee Ledger File"
   .Filter = "TxT files(*.txt)|*.txt"
   .ShowOpen
End With
strFilename = cdialog.FileName
lblFile.Refresh
lblFile.Caption = strFilename
If cdialog.CancelError Then
   lblStatus.Refresh
   lblStatus.Caption = ""
Else
   lblStatus.Refresh
   lblStatus.Caption = "File Dectected..." & GetFileFromPath(strFilename)
'   If Left(GetFileFromPath(strFilename), 10) = "EMP LEDGER" Then
   CmdUpdate.Enabled = True
'   Else
'      lblStatus.Caption = "Please Select the correct file "
'      CmdGetfile.Enabled = True
'   End If
                  
End If
End Sub
Function GetFileFromPath(vPath As String)
    Dim Items() As String
    Items = Split(vPath, "\")
    If UBound(Items) = -1 Then Exit Function
    GetFileFromPath = Items(UBound(Items))
End Function

Private Sub CmdSave_Click()

CmdSave.Enabled = False

If MSHFlexGrid1.Rows < 3 Then
   MsgBox "No data to save!!!"
   Exit Sub
End If

If MSHFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
MsgBox "Wait till Next Message....."
 
Call TableMst_EmpLedger
Call TableMst_Employee

ErrMsg = ""
RowCount = MSHFlexGrid1.Rows
ProgressBar1.Value = 0
ProgressBar1.Max = RowCount
lblStatus.Caption = "Please wait saving data..."
For I = 1 To RowCount - 2
    Wc = ""
    mStatus = 0
    If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
       If RsMst_EmpLedger.RecordCount > 0 Then
          RsMst_EmpLedger.MoveFirst
       End If
       RsMst_EmpLedger.Find "EntryNo = '" & MSHFlexGrid1.TextMatrix(I, 10) & "'"
       If RsMst_EmpLedger.EOF Then
          RsMst_Employee.MoveFirst
          RsMst_Employee.Find "NAVCode  = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'  "
          If Not RsMst_Employee.EOF Then
             mEmployeeID = RsMst_Employee!EmployeeID
          Else
             mEmployeeID = Null
             ErrMsg = ErrMsg & "Employee not found in Employee Master At Row " & I + 1 & vbCrLf
             MSHFlexGrid1.Col = 1
             MSHFlexGrid1.Row = I
             MSHFlexGrid1.CellBackColor = vbRed
          End If
           
          RsMst_EmpLedger.AddNew
          RsMst_EmpLedger!VendorCode = Trim(MSHFlexGrid1.TextMatrix(I, 0))
          RsMst_EmpLedger!VendorName = Trim(MSHFlexGrid1.TextMatrix(I, 1))
          RsMst_EmpLedger!DocumentType = Trim(MSHFlexGrid1.TextMatrix(I, 2))
          RsMst_EmpLedger!DocumentNo = Trim(MSHFlexGrid1.TextMatrix(I, 3))
          RsMst_EmpLedger!TransactionDate = MSHFlexGrid1.TextMatrix(I, 4)
          RsMst_EmpLedger!Narration = Trim(MSHFlexGrid1.TextMatrix(I, 5))
          RsMst_EmpLedger!TransactionAmount = Val(Replace(MSHFlexGrid1.TextMatrix(I, 6), ",", ""))
          RsMst_EmpLedger!RemainingAmount = Val(Replace(MSHFlexGrid1.TextMatrix(I, 7), ",", ""))
          RsMst_EmpLedger!AcutalPostingDate = MSHFlexGrid1.TextMatrix(I, 8)
          RsMst_EmpLedger!AcutalPostingTime = MSHFlexGrid1.TextMatrix(I, 9)
          RsMst_EmpLedger!EntryNo = MSHFlexGrid1.TextMatrix(I, 10)
          RsMst_EmpLedger!EmployeeID = mEmployeeID 'MSHFlexGrid1.TextMatrix(i, 0)
          RsMst_EmpLedger!CreatedBy = Gen_UserLoginID
          RsMst_EmpLedger!CreatedOn = Now 'MSHFlexGrid1.TextMatrix(i, 0)
          RsMst_EmpLedger.Update
       Else
           ErrMsg = ErrMsg & "Entry No already exist Employee Ledger At Row " & I + 1 & vbCrLf
           MSHFlexGrid1.Col = 0
           MSHFlexGrid1.Row = I
           MSHFlexGrid1.CellBackColor = vbGreen
       End If
    Else
       ErrMsg = ErrMsg & "Entry No Not Found in Grid " & MSHFlexGrid1.TextMatrix(r, 0) & "  At Row " & I + 1 & vbCrLf
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.Row = I
       MSHFlexGrid1.CellBackColor = vbRed
          'mStatus = 1
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max
'TxtErrorMsg = ErrMsg
lblStatus.Caption = "Done"
MsgBox "Done!!!"
CmdExcel.Enabled = True

End Sub

Private Sub CmdUpdate_Click()
Call Grid_Head
'CmdGetFile.Enabled = False
'cmdUpdate.Enabled = False
'cmdClose.Enabled = False
'CmdCancel.Enabled = False
              
lblStatus.Refresh
Call ReadFiles(strFilename)
lblStatus.Caption = ""
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
'   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Call CmdCancel_Click
End Sub

Public Sub ReadFiles(file_name As String)
   
Dim lngFreefile As Long
'Call TableMst_EmpLedger

Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(file_name)
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
  
If CheckFile(file_name) = False Then
   MsgBox "Please Select proper file!!!"
   Exit Sub
End If
  
sFileLineNo = 0

lblStatus.Refresh
'lblStatus.Caption = "Reading File These will take few mins..Please Wait"
ProgressBar1.Value = 0
ProgressBar1.Max = 100000 'objTS.AtEndOfStream
Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   tmp = Chr(34) + "~" & Chr(34)
   sFileLine = Replace(sFileLine, tmp, "~")
   sFileLine = Replace(sFileLine, Chr(34), "")
  
   sFileLineNo = sFileLineNo + 1
    'tmp = Chr(59)
   If Left(sFileLine, 3) = "EMP" Then
      sFileLine = Replace(sFileLine, "'", " ")
      ps = Split(sFileLine, "~")
      If UBound(ps) = 10 Then
         'For i = 0 To 10
         '    sArray(i) = Trim(ps(i))
         'Next i
         l = l + 1
         For c = 0 To MSHFlexGrid1.Cols - 1
             If InStr(Trim(ps(c)), "/") > 0 Then
                If IsDate(Trim(ps(c))) Then
                   mdate = Split(Trim(ps(c)), "/")
                   'dd/nn/yy
                   'StrDate = mdate(0) & "/" & mdate(1) & "/" & "20" & mdate(2)
                   StrDate = "20" & mdate(2) & "/" & mdate(1) & "/" & mdate(0)
                   
                   MSHFlexGrid1.TextMatrix(l, c) = StrDate 'Format(Trim(ps(c)), "dd/mm/yyyy")
                Else
                   MSHFlexGrid1.TextMatrix(l, c) = Trim(ps(c))
                End If
             Else
                MSHFlexGrid1.TextMatrix(l, c) = Trim(ps(c))
             End If
         'MSHFlexGrid1.TextMatrix(L, c) = Trim(ps(c)) 'Trim(oWS.Cells(mRow, c + 1))
         Next
         MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
      End If
         lblLines.Caption = sFileLineNo
   End If
   ProgressBar1.Value = ProgressBar1.Value + 1
Loop
ProgressBar1.Value = ProgressBar1.Max
lblStatus.Refresh
MsgBox "File Imported Sucessfully...."
ProgressBar1.Value = 0
CmdGetfile.Enabled = False
CmdUpdate.Enabled = False
CmdClose.Enabled = True
CmdCancel.Enabled = True
CmdSave.Enabled = True
CmdExcel.Enabled = True

End Sub

Private Function CheckFile(file_name) As Boolean
Dim Retval As Boolean
Retval = True

'Set objFS = New Scripting.FileSystemObject
'Set objFile = objFS.GetFile(file_name)
'Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
  
sFileLineNo = 0

lblStatus.Refresh
lblStatus.Caption = "Reading File These will take few mins..Please Wait"

sFileLine = objTS.ReadLine
tmp = Chr(34) + "~" & Chr(34)
sFileLine = Replace(sFileLine, tmp, "~")
sFileLine = Replace(sFileLine, Chr(34), "")

sFileLine = Replace(sFileLine, "'", " ")
ps = Split(sFileLine, "~")
If UBound(ps) = 10 Then
   For I = 0 To 10
       sArray(I) = Trim(ps(I))
   Next I
     
Else
   Retval = False
End If     'If RsDistrict.RecordCount > 0 Then

'sFileLineNo = sFileLineNo + 1
  
If LCase(sArray(0)) <> "vendor code" Then
   Retval = False
ElseIf LCase(sArray(1)) <> "vendor name" Then
   Retval = False
ElseIf LCase(sArray(2)) <> "document type" Then
   Retval = False
ElseIf LCase(sArray(3)) <> "document no." Then
   Retval = False
ElseIf LCase(sArray(4)) <> "transaction date" Then
   Retval = False
ElseIf LCase(sArray(5)) <> "narration" Then
   Retval = False
ElseIf LCase(sArray(6)) <> "transaction amount" Then
   Retval = False
ElseIf LCase(sArray(7)) <> "remaining amount" Then
   Retval = False
ElseIf LCase(sArray(8)) <> "acutal posting date" Then
   Retval = False
ElseIf LCase(sArray(9)) <> "acutal posting time" Then
   Retval = False
ElseIf LCase(sArray(10)) <> "entry no." Then
   Retval = False
End If

CheckFile = Retval

End Function

Private Sub Grid_Head()
    
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.AllowUserResizing = flexResizeColumns

MSHFlexGrid1.TextMatrix(0, 0) = "Vendor Code"
MSHFlexGrid1.TextMatrix(0, 1) = "Vendor Name"
MSHFlexGrid1.TextMatrix(0, 2) = "DocumentType"
MSHFlexGrid1.TextMatrix(0, 3) = "Document No"
MSHFlexGrid1.TextMatrix(0, 4) = "Transaction Date"
MSHFlexGrid1.TextMatrix(0, 5) = "Narration"
MSHFlexGrid1.TextMatrix(0, 6) = "Transaction Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "Remaining Amount"
MSHFlexGrid1.TextMatrix(0, 8) = "AcutalPosting Date"
MSHFlexGrid1.TextMatrix(0, 9) = "AcutalPosting Time"
MSHFlexGrid1.TextMatrix(0, 10) = "Entry No"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 2400
MSHFlexGrid1.ColWidth(3) = 2200
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 1600
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
   
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "EmployeeLedger.xls")
Call Xls_Detail_Rpt
CmdExcel.Enabled = False
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "EmployeeLedger.xls")
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
        
        MSHFlexGrid1.Row = I
        MSHFlexGrid1.Col = c
        
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbRed
        ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbYellow
        ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbGreen
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
