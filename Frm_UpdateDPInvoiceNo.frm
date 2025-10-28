VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_UpdateDPInvoiceNo 
   Caption         =   "Update DP Invoice No."
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   10830
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   15015
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   3
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   660
      End
   End
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
      TabIndex        =   0
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7680
         Left            =   120
         TabIndex        =   1
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
End
Attribute VB_Name = "Frm_UpdateDPInvoiceNo"
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
    .Cols = 26
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Date"
    .TextMatrix(1, 1) = "Invoice"
    .TextMatrix(1, 2) = "Customer"
    .TextMatrix(1, 3) = "DPNavCode"
    .TextMatrix(1, 4) = "Amount"
    .TextMatrix(1, 5) = "G/L Account"
    .TextMatrix(1, 6) = "GL A/C NAVCode "
    .TextMatrix(1, 12) = "Narration"
    .TextMatrix(1, 13) = "Location NAV Code"
    .TextMatrix(1, 14) = "CMP"
    .TextMatrix(1, 18) = "PI No"
    .TextMatrix(1, 19) = "ST Code"
    .TextMatrix(1, 20) = "Invoice No"
    .TextMatrix(1, 21) = "Invoice Date"
    .TextMatrix(1, 22) = "ST Amt"
    .TextMatrix(1, 23) = "Invoice Amt"
    .TextMatrix(1, 24) = "Post Invoice No"
    .TextMatrix(1, 25) = "Entry No"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1600
    .ColWidth(5) = 1600
    .ColWidth(6) = 1600
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 2000
    .ColWidth(13) = 1600
    .ColWidth(14) = 1400
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 1400
    .ColWidth(19) = 1000
    .ColWidth(20) = 1200
    .ColWidth(21) = 1200
    .ColWidth(22) = 1200
    .ColWidth(23) = 1200
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

MsgBox "Please Wait till Next Message."
Call TableMst_DPInvoiceUpload
Call TableTxn_TIDetails '(" Where PIFlag='DP'")
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 20) = "" Or MSHFlexGrid3.TextMatrix(I, 21) = "" Then Exit For
    If RsMst_DPInvoiceUpload.RecordCount > 0 Then
       RsMst_DPInvoiceUpload.MoveFirst
    End If
    mInvoice = Replace(MSHFlexGrid3.TextMatrix(I, 18), "DP", "")
    RsMst_DPInvoiceUpload.Find "PINo= " & Val(mInvoice) & " " '" & Val(MSHFlexGrid3.TextMatrix(I, 18)) & ""
    If Not RsMst_DPInvoiceUpload.EOF Then
       If RsTxn_TIDetails.RecordCount > 0 Then
          RsTxn_TIDetails.MoveFirst
          RsTxn_TIDetails.Find "EntryNo=" & Val(MSHFlexGrid3.TextMatrix(I, 25)) 'RsTxn_TIDetails.Find "PIID=" & RsMst_DPInvoiceUpload!DPInvoiceID
          If RsTxn_TIDetails.EOF = True Then
             RsTxn_TIDetails.AddNew
             RsTxn_TIDetails!TIID = getTIDetailMax
             RsTxn_TIDetails!TINo = Trim(MSHFlexGrid3.TextMatrix(I, 20))
             RsTxn_TIDetails!TIDate = Format(MSHFlexGrid3.TextMatrix(I, 21), Gen_DatFormat)
             RsTxn_TIDetails!TIAmount = Val(MSHFlexGrid3.TextMatrix(I, 23))
             RsTxn_TIDetails!TISTAmount = Val(MSHFlexGrid3.TextMatrix(I, 22))
             RsTxn_TIDetails!PIID = RsMst_DPInvoiceUpload!DPInvoiceID
             RsTxn_TIDetails!LocationID = RsMst_DPInvoiceUpload!LocationID
             RsTxn_TIDetails!DepositoryID = GetDepositoryIDFromNavCode(RsMst_DPInvoiceUpload!DPNAVCode)
             RsTxn_TIDetails!NAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
             RsTxn_TIDetails!ExchangeTypeID = 2
             RsTxn_TIDetails!PIFlag = "DP"
             RsTxn_TIDetails!SMID = MSHFlexGrid3.TextMatrix(I, 18)
             RsTxn_TIDetails!RecAmount = 0
             RsTxn_TIDetails!RecTDSAmount = 0
             RsTxn_TIDetails!RecOtherDed = 0
'             RsTxn_TIDetails!EntryFlag = "System"
             RsTxn_TIDetails!G_UID = GetGUID
             RsTxn_TIDetails!CreatedDate = Now
             RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
             RsTxn_TIDetails!PostInvoiceNo = Trim(MSHFlexGrid3.TextMatrix(I, 24))
             RsTxn_TIDetails!EntryNo = Val(MSHFlexGrid3.TextMatrix(I, 25))
             RsTxn_TIDetails.Update
             
             RsMst_DPInvoiceUpload!InvoiceDate = Format(MSHFlexGrid3.TextMatrix(I, 21), Gen_DatFormat)
             RsMst_DPInvoiceUpload!InvoiceNo = Trim(MSHFlexGrid3.TextMatrix(I, 20))
             RsMst_DPInvoiceUpload!InvSTAmt = Val(MSHFlexGrid3.TextMatrix(I, 22))
             RsMst_DPInvoiceUpload!InvoiceAmt = Val(MSHFlexGrid3.TextMatrix(I, 23))
             RsMst_DPInvoiceUpload!UpdatedBy = Gen_UserLoginID
             RsMst_DPInvoiceUpload!UpdatedDate = Now
             RsMst_DPInvoiceUpload.Update
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
    If UBound(ps) <> 25 Then
       MsgBox "Invalid file format."
       Exit Sub
    End If
    For I = 0 To UBound(ps)
         MSHFlexGrid3.TextMatrix(mRow, I) = Trim(ps(I))
    Next I
    mRow = mRow + 1
    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
Loop
MsgBox "Import Successfully"
End Sub

Function getTIDetailMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(TIID) from Txn_TIDetails"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getTIDetailMax = Retval

End Function


