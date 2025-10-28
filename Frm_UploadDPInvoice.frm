VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_UploadDPInvoice 
   Caption         =   "Upload DP Invoice"
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
      Caption         =   "Non Taxable"
      Height          =   3705
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   3360
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5927
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Taxable"
      Height          =   3705
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   3360
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5927
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
         TabIndex        =   13
         Top             =   720
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Record Already Exist."
         Height          =   195
         Left            =   13260
         TabIndex        =   17
         Top             =   315
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         Height          =   375
         Left            =   12690
         TabIndex        =   16
         Top             =   255
         Width           =   495
      End
      Begin VB.Label lblProgress 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   45
      End
      Begin VB.Label lblProcess 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   7440
         TabIndex        =   14
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
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   12720
         TabIndex        =   6
         Top             =   240
         Width           =   2100
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
         TabIndex        =   5
         Top             =   255
         Width           =   1620
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
         Left            =   6780
         TabIndex        =   4
         Top             =   255
         Width           =   1620
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
         Left            =   5160
         TabIndex        =   3
         Top             =   255
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
         Left            =   3540
         TabIndex        =   2
         Top             =   255
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
         Left            =   285
         TabIndex        =   1
         Top             =   255
         Width           =   1620
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
Attribute VB_Name = "Frm_UploadDPInvoice"
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
If LblFile.Caption <> Path Then
 '  Call Clear
End If
LblFile.Caption = Path
If LblFile.Caption <> "" Then
   CmdImport.Enabled = True
End If
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub CmdImport_Click()
Dim fn, CFlag As Variant

fn = strFilename
Dim objFileSystem As Object
Dim FileToCheck As String
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set oXL = New Excel.Application
Dim Pat As String
Set oWB = oXL.Workbooks.Open(LblFile.Caption) '(Pat & "\excel\" & fn & ".xls", , , , "")
Set oWS = oWB.Worksheets(1)
If oWS.Name <> "Non Taxable" Then
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
Set oWS = oWB.Worksheets(2)
If oWS.Name <> "Taxable" Then
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
Set oWS = oWB.Worksheets(1)
ProgressBar1.Value = 0
mRows = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
ProgressBar1.Max = mRows - 1

MSHFlexGrid3.Rows = mRows + 1
For I = 2 To mRows - 1
    If Trim(oWS.Cells(I, 1)) = "" Then Exit For
    For c = 0 To MSHFlexGrid3.Cols - 1
        MSHFlexGrid3.TextMatrix(I, c) = Trim(oWS.Cells(I + 1, c + 1))
    Next
    MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + Val(MSHFlexGrid3.TextMatrix(I, 5))
    MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(I, 6))
    MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 7))
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

Set oWS = oWB.Worksheets(2)
ProgressBar1.Value = 0
mRows = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
ProgressBar1.Max = mRows - 1

MSHFlexGrid4.Rows = mRows + 1
For I = 2 To mRows - 1
    If Trim(oWS.Cells(I, 1)) = "" Then Exit For
    For c = 0 To MSHFlexGrid4.Cols - 1
        MSHFlexGrid4.TextMatrix(I, c) = Trim(oWS.Cells(I + 1, c + 1))
    Next
    MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(I, 5))
    MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(I, 6))
    MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(I, 7))
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = 0
CmdSave.Enabled = True
End Sub
Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True
 '

If oWS.Cells(2, 1) <> "Month-Year" Then
   Retval = False
ElseIf oWS.Cells(2, 2) <> "Location" Then
   Retval = False
ElseIf oWS.Cells(2, 3) <> "Location ID" Then
   Retval = False
ElseIf oWS.Cells(2, 4) <> "DP Nav Code" Then
   Retval = False
ElseIf oWS.Cells(2, 5) <> "DP Name" Then
   Retval = False
ElseIf oWS.Cells(2, 6) <> "Amount" Then
   Retval = False
ElseIf oWS.Cells(2, 7) <> "ST. Amount" Then
   Retval = False
ElseIf oWS.Cells(2, 8) <> "Total Amount" Then
   Retval = False
End If
CheckFile = Retval
End Function

Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 8
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Month-Year"
    .TextMatrix(1, 1) = "Location"
    .TextMatrix(1, 2) = "Location ID"
    .TextMatrix(1, 3) = "DP Nav Code"
    .TextMatrix(1, 4) = "DP Name"
    .TextMatrix(1, 5) = "Amount"
    .TextMatrix(1, 6) = "ST. Amount"
    .TextMatrix(1, 7) = "Total Amount"
    
    .ColWidth(0) = 1600
    .ColWidth(1) = 2200
    .ColWidth(2) = 1200
    .ColWidth(3) = 1400
    .ColWidth(4) = 3000
    .ColWidth(5) = 1400
    .ColWidth(6) = 1400
    .ColWidth(7) = 1400
End With
With MSHFlexGrid4
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 8
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Month-Year"
    .TextMatrix(1, 1) = "Location"
    .TextMatrix(1, 2) = "Location ID"
    .TextMatrix(1, 3) = "DP Nav Code"
    .TextMatrix(1, 4) = "DP Name"
    .TextMatrix(1, 5) = "Amount"
    .TextMatrix(1, 6) = "ST. Amount"
    .TextMatrix(1, 7) = "Total Amount"
    
    .ColWidth(0) = 1600
    .ColWidth(1) = 2200
    .ColWidth(2) = 1200
    .ColWidth(3) = 1400
    .ColWidth(4) = 3000
    .ColWidth(5) = 1400
    .ColWidth(6) = 1400
    .ColWidth(7) = 1400
End With
End Sub

Private Sub CmdReset_Click()
Call Grid_Head
LblFile.Caption = ""
ProgressBar1.Value = 0
CmdSave.Enabled = False
CmdImport.Enabled = False
CmdGetFile.Enabled = True
End Sub

Private Sub CmdSave_Click()

MsgBox "Please Wait till Next Message."
Call TableMst_DPInvoiceUpload
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid3.Rows + MSHFlexGrid4.Rows - 6
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit For
    RsMst_DPInvoiceUpload.Filter = "MonthYear='" & Format(MSHFlexGrid3.TextMatrix(I, 0), Gen_DatFormat) & "' And LocationID=" & Val(MSHFlexGrid3.TextMatrix(I, 2)) & " And DPNavCode='" & Trim(MSHFlexGrid3.TextMatrix(I, 3)) & "' And Flag = 'N'"
    If RsMst_DPInvoiceUpload.RecordCount = 0 Then
       RsMst_DPInvoiceUpload.AddNew
       RsMst_DPInvoiceUpload!DPInvoiceID = GetDPInvoiceID
       RsMst_DPInvoiceUpload!MonthYear = Format(MSHFlexGrid3.TextMatrix(I, 0), Gen_DatFormat)
       RsMst_DPInvoiceUpload!LocationID = Val(MSHFlexGrid3.TextMatrix(I, 2))
       RsMst_DPInvoiceUpload!DPNAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
       RsMst_DPInvoiceUpload!Amt = Val(MSHFlexGrid3.TextMatrix(I, 5))
       RsMst_DPInvoiceUpload!STAmt = Val(MSHFlexGrid3.TextMatrix(I, 6))
       RsMst_DPInvoiceUpload!TotalAmt = Val(MSHFlexGrid3.TextMatrix(I, 7))
       RsMst_DPInvoiceUpload!Flag = "N"
       RsMst_DPInvoiceUpload!PINo = GetDPInvoiceID
       RsMst_DPInvoiceUpload!CreatedBy = Gen_UserLoginID
       RsMst_DPInvoiceUpload!CreatedDate = Now
       RsMst_DPInvoiceUpload.Update
       RsMst_DPInvoiceUpload.Requery
       RsMst_DPInvoiceUpload.Filter = ""
    Else
       MSHFlexGrid3.row = I
       MSHFlexGrid3.Col = 0
       MSHFlexGrid3.CellBackColor = vbGreen
       RsMst_DPInvoiceUpload.Filter = ""
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 0) = "" Then Exit For
    RsMst_DPInvoiceUpload.Filter = "MonthYear='" & Format(MSHFlexGrid4.TextMatrix(I, 0), Gen_DatFormat) & "' And LocationID=" & Val(MSHFlexGrid4.TextMatrix(I, 2)) & " And DPNavCode='" & Trim(MSHFlexGrid4.TextMatrix(I, 3)) & "' And Flag = 'T'"
    If RsMst_DPInvoiceUpload.RecordCount = 0 Then
    
       RsMst_DPInvoiceUpload.AddNew
       RsMst_DPInvoiceUpload!DPInvoiceID = GetDPInvoiceID
       RsMst_DPInvoiceUpload!MonthYear = Format(MSHFlexGrid4.TextMatrix(I, 0), Gen_DatFormat)
       RsMst_DPInvoiceUpload!LocationID = Val(MSHFlexGrid4.TextMatrix(I, 2))
       RsMst_DPInvoiceUpload!DPNAVCode = MSHFlexGrid4.TextMatrix(I, 3)
       RsMst_DPInvoiceUpload!Amt = Val(MSHFlexGrid4.TextMatrix(I, 5))
       RsMst_DPInvoiceUpload!STAmt = Val(MSHFlexGrid4.TextMatrix(I, 6))
       RsMst_DPInvoiceUpload!TotalAmt = Val(MSHFlexGrid4.TextMatrix(I, 7))
       RsMst_DPInvoiceUpload!Flag = "T"
       RsMst_DPInvoiceUpload!PINo = GetDPInvoiceID
       RsMst_DPInvoiceUpload!CreatedBy = Gen_UserLoginID
       RsMst_DPInvoiceUpload!CreatedDate = Now
       RsMst_DPInvoiceUpload.Update
       RsMst_DPInvoiceUpload.Requery
       RsMst_DPInvoiceUpload.Filter = ""
    Else
       MSHFlexGrid4.row = I
       MSHFlexGrid4.Col = 0
       MSHFlexGrid4.CellBackColor = vbGreen
        RsMst_DPInvoiceUpload.Filter = ""
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
MsgBox "Done !!"
 ProgressBar1.Value = 0
End Sub
Private Function GetDPInvoiceID() As Double
Dim Retval As Double
tmp = "Select max(DPInvoiceID) from Mst_DPInvoiceUpload"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetDPInvoiceID = Retval
End Function


Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetFile.Enabled = False
   CmdImport.Enabled = False
   CmdSave.Enabled = False
   CmdReset.Enabled = False
   Exit Sub
End If

Call CmdReset_Click
End Sub
