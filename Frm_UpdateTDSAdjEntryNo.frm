VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_UpdateTDSAdjEntryNo 
   Caption         =   "Update TDS Adj Entry No"
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8025
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7680
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13547
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   1905
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
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
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
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   660
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_UpdateTDSAdjEntryNo"
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

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdGetFile_Click()
On Error GoTo errortrap

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "CSV files(*.csv)|*.csv" '|Text files(*.txt)|*.txt"
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
Call ReadFiles(LblFile.Caption)
CmdUpdate.Enabled = True
End Sub

Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 16
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(1, 0) = "Posting Date"
    .TextMatrix(1, 1) = "Document Type"
    .TextMatrix(1, 2) = "Account Type"
    .TextMatrix(1, 3) = "Account No."
    .TextMatrix(1, 4) = "Amount"
    .TextMatrix(1, 5) = "Invoice No"
    .TextMatrix(1, 6) = "Narration"
    .TextMatrix(1, 7) = "RegionDim"
    .TextMatrix(1, 8) = "DeptDim"
    .TextMatrix(1, 9) = "EmpDim"
    .TextMatrix(1, 10) = "Cheque No."
    .TextMatrix(1, 11) = "Cheque Date"
    .TextMatrix(1, 12) = "Deposit Date"
    .TextMatrix(1, 13) = "PINo"
    .TextMatrix(1, 14) = "PostInvoiceNo"
    .TextMatrix(1, 15) = "EntryNo"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1600
    .ColWidth(5) = 1600
    .ColWidth(6) = 1600
    .ColWidth(7) = 2000
    .ColWidth(8) = 1600
    .ColWidth(9) = 1400
    .ColWidth(10) = 1400
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    
End With

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
LblFile.Caption = ""
CmdUpdate.Enabled = False
CmdImport.Enabled = False
CmdGetFile.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Dim mCRID, mSMPrefix As Variant
Dim mstr() As String
MsgBox "Please Wait till Next Message."
For i = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(i, 13) <> "" Then
       mstr = Split(MSHFlexGrid3.TextMatrix(i, 13), "-")
       If UBound(mstr) = 1 Then
          mCRID = mstr(0)
          mSMPrefix = mstr(1)
          
          If LCase(MSHFlexGrid3.TextMatrix(i, 1)) = "payment" Then
             tmp = "Update Txn_CRTDSAdjustmentDetail Set TDSDocNo='" & Trim(MSHFlexGrid3.TextMatrix(i, 14)) & "',TDSEntryNo=" & Val(MSHFlexGrid3.TextMatrix(i, 15)) & ""
             tmp = tmp & " Where TIID in( Select TIID from  Txn_TIDetails Where TIID in (select TIID from Txn_CRTDSAdjustmentDetail Where TDSAID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "') and NAVCode='" & Trim(MSHFlexGrid3.TextMatrix(i, 3)) & "') And RecTDS>0 And TDSAID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "'"
             
             Call TmpTable
          ElseIf LCase(MSHFlexGrid3.TextMatrix(i, 1)) = "credit memo" Then
             If InStr(1, LCase(MSHFlexGrid3.TextMatrix(i, 6)), "other deduction") > 0 Then
                tmp = "Update Txn_CRTDSAdjustmentDetail Set OtherDedDocNo='" & Trim(MSHFlexGrid3.TextMatrix(i, 14)) & "',OtherDedEntryNo=" & Val(MSHFlexGrid3.TextMatrix(i, 15)) & ""
                tmp = tmp & " Where TIID in( Select TIID from  Txn_TIDetails Where TIID in (select TIID from Txn_CRTDSAdjustmentDetail Where TDSAID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "') and NAVCode='" & Trim(MSHFlexGrid3.TextMatrix(i, 3)) & "') And RecOtherDed>0 And TDSAID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "'"

                Call TmpTable
             End If
          End If
         
       End If
    End If
Next
MsgBox "Done !!"
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
Call CmdReset_Click
End Sub

Public Sub ReadFiles(file_name As String)
Dim lngFreefile As Long
Dim mRow As Variant
MsgBox "Wait till next message"

Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(file_name)
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
sFileLineNo = 1
mRow = 2

Do While Not objTS.AtEndOfStream
    sFileLine = objTS.ReadLine
    sFileLineNo = sFileLineNo + 1
    ps = Split(sFileLine, ",")
    If UBound(ps) <> 15 Then
       MsgBox "Invalid file format."
       Exit Sub
    End If
    For i = 0 To UBound(ps)
         MSHFlexGrid3.TextMatrix(mRow, i) = Trim(ps(i))
    Next i
    mRow = mRow + 1
    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
Loop
MsgBox "Import Successfully"
End Sub


