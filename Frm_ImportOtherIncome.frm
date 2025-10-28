VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_ImportOtherIncome 
   Caption         =   "Import Other Income"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Details"
      Height          =   7875
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7515
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13256
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1170
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000C000&
         Height          =   375
         Left            =   10440
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Record Already Exist."
         Height          =   195
         Left            =   11010
         TabIndex        =   17
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Location Not found"
         Height          =   195
         Left            =   13260
         TabIndex        =   15
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   12690
         TabIndex        =   14
         Top             =   255
         Width           =   495
      End
      Begin VB.Label lblProgress 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   45
      End
      Begin VB.Label lblProcess 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   7440
         TabIndex        =   12
         Top             =   480
         Width           =   45
      End
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
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbIncomeType 
         Height          =   315
         ItemData        =   "Frm_ImportOtherIncome.frx":0000
         Left            =   1320
         List            =   "Frm_ImportOtherIncome.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   273
         Width           =   2505
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
         Left            =   5403
         TabIndex        =   2
         Top             =   255
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
         Height          =   350
         Left            =   9960
         TabIndex        =   5
         Top             =   255
         Width           =   1500
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
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
         Left            =   8439
         TabIndex        =   4
         Top             =   255
         Width           =   1500
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
         Left            =   6921
         TabIndex        =   3
         Top             =   255
         Width           =   1500
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
         Left            =   3885
         TabIndex        =   1
         Top             =   255
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Income Type"
         Height          =   195
         Left            =   255
         TabIndex        =   16
         Top             =   330
         Width           =   930
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_ImportOtherIncome"
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
Dim mIETypeID As Variant

Private Sub CmbIncomeType_GotFocus()
tmp = CmbIncomeType.Text
Call TableMst_IncomeExpenseType(" WHERE IEFlag='I' AND Flag='A'")
CmbIncomeType.Clear
If RsMst_IncomeExpenseType.RecordCount = 0 Then
   MsgBox "No Income Type found"
   CmbIncomeType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_IncomeExpenseType.RecordCount
    CmbIncomeType.AddItem RsMst_IncomeExpenseType!IEType
    RsMst_IncomeExpenseType.MoveNext
Next
CmbIncomeType.Text = tmp
End Sub

Private Sub CmbIncomeType_LostFocus()
If Trim(CmbIncomeType.Text) = "" Then
   mSegmentTypeID = Null
   Exit Sub
End If

RsMst_IncomeExpenseType.MoveFirst
RsMst_IncomeExpenseType.Find "IEType = '" & CmbIncomeType.Text & "'"

If RsMst_IncomeExpenseType.EOF Then
   MsgBox "Invalid Income Type Selection "
   CmbIncomeType.SetFocus
   Exit Sub
End If
mIETypeID = RsMst_IncomeExpenseType!IETypeID
If tmp <> CmbIncomeType.Text Then
   Call Grid_Head
   lblFile.Caption = ""
   ProgressBar1.Value = 0
   CmdSave.Enabled = False
   CmdImport.Enabled = False
   CmdGetfile.Enabled = True
End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdGetFile_Click()
On Error GoTo errortrap

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls)|*.xls" '|Text files(*.txt)|*.txt"
     .MaxFileSize = &H7FFF ' 32KB filename buffer
    'same as .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNLongNames or cdlOFNExplorer --Or cdlOFNAllowMultiselect
     .Flags = &H4 Or &H800 Or &H40000 Or &H80000 Or &H200
     .CancelError = True
procReOpen:
     .ShowOpen
End With
Path = CommonDialog1.FileName
mFilenames = CommonDialog1.FileName
If lblFile.Caption <> Path Then
 '  Call Clear
End If
lblFile.Caption = Path
If lblFile.Caption <> "" Then
   CmdImport.Enabled = True
End If
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub CmdImport_Click()
Dim fn, CFlag               As Variant
Dim intCol                  As Variant
Dim iCount                  As Variant
Dim mdate                   As Variant
fn = strFilename
Dim objFileSystem           As Object
Dim FileToCheck             As String

Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set oXL = New Excel.Application
Dim Pat As String
Set oWB = oXL.Workbooks.Open(lblFile.Caption) '(Pat & "\excel\" & fn & ".xls", , , , "")
Set oWS = oWB.Worksheets(1)

If CmbIncomeType.Text = "" Then
   MsgBox "Blank Income Type Not Allowed !!! "
   CmbIncomeType.SetFocus
   Exit Sub
End If

If LCase(oWS.Name) <> LCase(CmbIncomeType.Text) Then
    MsgBox "Please Select proper file for uploading!!!"
    CmdImport.Enabled = False
    Exit Sub
Else
   If CheckFile = False Then
      MsgBox "Please Select proper file for uploading!!!"
      CmdImport.Enabled = False
      Exit Sub
   End If
End If

Call Grid_Head
Dim I, mRows As Variant
ProgressBar1.Value = 0
mRows = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row
ProgressBar1.Max = mRows - 1

MSHFlexGrid3.Rows = mRows + 1
For I = 1 To mRows - 1
    DoEvents
    If Trim(oWS.Cells(I, 1)) = "" Then Exit For
    intCol = 1
    For c = 0 To MSHFlexGrid3.Cols - 1
        If Trim(oWS.Cells(I + 1, 2)) <> "" Then
            If c = 0 Then
               mdate = ""
               mdate = "01-" & Format(oWS.Cells(I + 1, intCol), "mm-yyyy")
               MSHFlexGrid3.TextMatrix(I, c) = Format(mdate, "dd-mm-yyyy")
            ElseIf c = 1 Then
               MSHFlexGrid3.TextMatrix(I, c) = Format(oWS.Cells(I + 1, intCol), "mmm-yy")
               intCol = intCol + 1
            ElseIf c = 3 Then
               If Val(GetLocationID(Trim(oWS.Cells(I + 1, intCol - 1)))) = 0 Then
                  MSHFlexGrid3.TextMatrix(I, c) = 0
                  For iCount = 0 To MSHFlexGrid3.Cols - 1
                      MSHFlexGrid3.Col = iCount
                      MSHFlexGrid3.Row = I
                      MSHFlexGrid3.CellBackColor = vbRed
                  Next
               Else
                  MSHFlexGrid3.TextMatrix(I, c) = GetLocationID(Trim(oWS.Cells(I + 1, intCol - 1)))
               End If
            Else
               MSHFlexGrid3.TextMatrix(I, c) = Trim(oWS.Cells(I + 1, intCol))
               intCol = intCol + 1
            End If
        End If
    Next
    'MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + Val(MSHFlexGrid3.TextMatrix(I, 5))
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
oWB.Close
ProgressBar1.Value = 0
MsgBox "Done!!"
CmdSave.Enabled = True
End Sub
Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True
 
If LCase(oWS.Cells(1, 1)) <> "month" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 2)) <> "location name" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 3)) <> "client name" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 4)) <> "amount" Then
   Retval = False
End If
CheckFile = Retval
End Function

Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 6
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Txn Date"
    .TextMatrix(0, 1) = "Month-Year"
    .TextMatrix(0, 2) = "Location"
    .TextMatrix(0, 3) = "Location ID"
    .TextMatrix(0, 4) = "Client Name"
    .TextMatrix(0, 5) = "Amount"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 2200
    .ColWidth(3) = 0
    .ColWidth(4) = 3500
    .ColWidth(5) = 1200
End With
End Sub

Private Sub CmdReset_Click()
Call Grid_Head
lblFile.Caption = ""
ProgressBar1.Value = 0
CmdSave.Enabled = False
CmdImport.Enabled = False
CmdGetfile.Enabled = True
CmbIncomeType.Text = ""
End Sub

Private Sub CmdSave_Click()
MsgBox "Please Wait till Next Message."
Call SaveDetails
MsgBox "Done !!"
ProgressBar1.Value = 0
End Sub

Private Sub SaveDetails()
Call TableTxn_OtherIncome(" Where IETypeID=" & mIETypeID & "")
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid3.Rows - 2
If CmbIncomeType.Text = "" Then
   MsgBox "Blank Income Type Not Allowed !!! "
   CmbIncomeType.SetFocus
   Exit Sub
End If
For I = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit For
    RsTxn_OtherIncome.Filter = "MonthYear='" & MSHFlexGrid3.TextMatrix(I, 1) & "' And LocationID=" & Val(MSHFlexGrid3.TextMatrix(I, 3)) & " And ClientName='" & Trim(MSHFlexGrid3.TextMatrix(I, 4)) & "'"
    If RsTxn_OtherIncome.RecordCount = 0 Then
       If Val(MSHFlexGrid3.TextMatrix(I, 3)) <> 0 Then
          RsTxn_OtherIncome.AddNew
          RsTxn_OtherIncome!OtherIncomeID = GetOtherIncomeID
          RsTxn_OtherIncome!IETypeID = mIETypeID
          RsTxn_OtherIncome!TxnDate = Format(MSHFlexGrid3.TextMatrix(I, 0), Gen_DatFormat)
          RsTxn_OtherIncome!MonthYear = Trim(MSHFlexGrid3.TextMatrix(I, 1))
          RsTxn_OtherIncome!LocationID = Val(MSHFlexGrid3.TextMatrix(I, 3))
          RsTxn_OtherIncome!ClientName = Trim(MSHFlexGrid3.TextMatrix(I, 4))
          RsTxn_OtherIncome!Amount = Val(MSHFlexGrid3.TextMatrix(I, 5))
         ' RsTxn_OtherIncome!Flag = CmbIncomeType.Text
          RsTxn_OtherIncome!CreatedBy = Gen_UserLoginID
          RsTxn_OtherIncome!CreatedDate = Now
          RsTxn_OtherIncome.Update
          RsTxn_OtherIncome.Filter = ""
       End If
    Else
       MSHFlexGrid3.Row = I
       MSHFlexGrid3.Col = 0
       MSHFlexGrid3.CellBackColor = vbGreen
       RsTxn_OtherIncome.Filter = ""
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
End Sub

Private Function GetOtherIncomeID() As Double
Dim Retval As Double
tmp = "Select max(OtherIncomeID) from Txn_OtherIncome"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetOtherIncomeID = Retval
End Function

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetfile.Enabled = False
   CmdImport.Enabled = False
   CmdSave.Enabled = False
   CmdReset.Enabled = False
   Exit Sub
End If

Call CmdReset_Click
End Sub
