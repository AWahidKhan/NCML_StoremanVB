VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_ImportGodClaimPayDetails 
   Caption         =   "Import Expenses Payment Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   15015
      Begin VB.OptionButton OptFumigationDetails 
         Caption         =   "Fumigation Payment Details"
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
         Left            =   9240
         TabIndex        =   11
         Top             =   240
         Width           =   2970
      End
      Begin VB.OptionButton OptSecurityDetails 
         Caption         =   "Security Payment Details"
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
         Left            =   6240
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
      Begin VB.OptionButton OptFranchiseeClaim 
         Caption         =   "Franchisee Payment Details"
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
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
      Begin VB.OptionButton OptGodownRent 
         Caption         =   "Godown Payment Details"
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2655
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
         Height          =   435
         Left            =   8160
         TabIndex        =   6
         Top             =   960
         Width           =   1350
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12210
         TabIndex        =   5
         Top             =   960
         Width           =   1350
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
         Height          =   435
         Left            =   10860
         TabIndex        =   4
         Top             =   960
         Width           =   1350
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
         Height          =   435
         Left            =   13560
         TabIndex        =   3
         Top             =   960
         Width           =   1350
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
         Height          =   435
         Left            =   9510
         TabIndex        =   2
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   1080
         Width           =   660
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   1545
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   13361
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
End
Attribute VB_Name = "Frm_ImportGodClaimPayDetails"
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

If OptGodownRent.Value = True Then
   If LCase(oWS.Name) <> "godownrent" Then
      MsgBox "Please provide Sheet Name as 'GodownRent'"
      CmdImport.Enabled = False
      Exit Sub
   End If
ElseIf OptFumigationDetails.Value = True Then
   If LCase(oWS.Name) <> "fumigation" Then
      MsgBox "Please provide Sheet Name as 'Fumigation'"
      CmdImport.Enabled = False
      Exit Sub
   End If
ElseIf OptFranchiseeClaim.Value = True Then
   If LCase(oWS.Name) <> "franchisee" Then
      MsgBox "Please provide Sheet Name as 'Franchisee'"
      CmdImport.Enabled = False
      Exit Sub
   End If
ElseIf OptSecurityDetails.Value = True Then
   If LCase(oWS.Name) <> "security" Then
      MsgBox "Please provide Sheet Name as 'Security'"
      CmdImport.Enabled = False
      Exit Sub
   End If
End If

If CheckFile = False Then
   MsgBox "Please Select file with colums Date,NAV code,Claim ID,Bank Name,ch.no.,Amount & TDS"
   CmdImport.Enabled = False
   Exit Sub
End If

Call Grid_Head
Dim I, mRows As Variant
Set oWS = oWB.Worksheets(1)

mRows = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row

MSHFlexGrid1.Rows = mRows + 1
For I = 1 To mRows - 1
    If Trim(oWS.Cells(I, 1)) = "" Then Exit For
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(I, c) = Trim(oWS.Cells(I + 1, c + 1))
    Next
Next
CmdSave.Enabled = True
End Sub
Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True

If LCase(oWS.Cells(1, 1)) <> "date" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 2)) <> "nav code" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 3)) <> "claim id" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 4)) <> "bank name" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 5)) <> "ch.no." Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 6)) <> "amount" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 7)) <> "tds" Then
   Retval = False
End If
CheckFile = Retval
End Function

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 7
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Date"
    .TextMatrix(0, 1) = "NAV code"
    .TextMatrix(0, 2) = "Claim ID"
    .TextMatrix(0, 3) = "Bank Name"
    .TextMatrix(0, 4) = "ch.no."
    .TextMatrix(0, 5) = "Amount"
    .TextMatrix(0, 6) = "TDS"

    .ColWidth(0) = 1500
    .ColWidth(1) = 1800
    .ColWidth(2) = 1400
    .ColWidth(3) = 2800
    .ColWidth(4) = 2200
    .ColWidth(5) = 1600
    .ColWidth(6) = 1600
End With
End Sub

Private Sub CmdReset_Click()
Call Grid_Head
LblFile.Caption = ""
CmdSave.Enabled = False
CmdImport.Enabled = False
CmdGetFile.Enabled = True
OptGodownRent.Value = True
End Sub

Private Sub CmdSave_Click()
MsgBox "Please Wait till Next Message."
If OptGodownRent.Value = True Then
   Call TableTxn_GodownRent(" Where Flag in ('S', 'A') And BatchID > 0 ")
   For I = 1 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit For
       If RsTxn_GodownRent.RecordCount > 0 Then
          RsTxn_GodownRent.MoveFirst
       End If
       RsTxn_GodownRent.Find "ClaimID=" & Val(MSHFlexGrid1.TextMatrix(I, 2))
       If Not RsTxn_GodownRent.EOF Then
          RsTxn_GodownRent!Bankname = Trim(MSHFlexGrid1.TextMatrix(I, 3))
          RsTxn_GodownRent!ChequeNo = Trim(MSHFlexGrid1.TextMatrix(I, 4))
          RsTxn_GodownRent!PaidAmount = Val(MSHFlexGrid1.TextMatrix(I, 5))
          RsTxn_GodownRent!TDS = Val(MSHFlexGrid1.TextMatrix(I, 6))
          RsTxn_GodownRent!PaidDate = CDate(MSHFlexGrid1.TextMatrix(I, 0))
          RsTxn_GodownRent!PaymentdetailsUpdatedBy = Gen_UserLoginID
          RsTxn_GodownRent!PaymentdetailsUpdatedOn = Now
          RsTxn_GodownRent.Update
          RsTxn_GodownRent.Requery
       End If
   Next
ElseIf OptFranchiseeClaim.Value = True Then
   Call TableTxn_FranchiseeRent(" Where Flag in ('S', 'A') And BatchID > 0 ")
   For I = 1 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit For
       If RsTxn_FranchiseeRent.RecordCount > 0 Then
          RsTxn_FranchiseeRent.MoveFirst
       End If
       RsTxn_FranchiseeRent.Find "ClaimID=" & Val(MSHFlexGrid1.TextMatrix(I, 2))
       If Not RsTxn_FranchiseeRent.EOF Then
          RsTxn_FranchiseeRent!Bankname = Trim(MSHFlexGrid1.TextMatrix(I, 3))
          RsTxn_FranchiseeRent!ChequeNo = Trim(MSHFlexGrid1.TextMatrix(I, 4))
          RsTxn_FranchiseeRent!PaidAmount = Val(MSHFlexGrid1.TextMatrix(I, 5))
          RsTxn_FranchiseeRent!TDS = Val(MSHFlexGrid1.TextMatrix(I, 6))
          RsTxn_FranchiseeRent!PaidDate = CDate(MSHFlexGrid1.TextMatrix(I, 0))
          RsTxn_FranchiseeRent!PaymentdetailsUpdatedBy = Gen_UserLoginID
          RsTxn_FranchiseeRent!PaymentdetailsUpdatedOn = Now
          RsTxn_FranchiseeRent.Update
          RsTxn_FranchiseeRent.Requery
       End If
   Next
ElseIf OptSecurityDetails.Value = True Then
   Call TableTxn_SecurityBilling(" Where Status in ('S', 'A') And BatchID > 0 ")
   For I = 1 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit For
       If RsTxn_SecurityBilling.RecordCount > 0 Then
          RsTxn_SecurityBilling.MoveFirst
       End If
       RsTxn_SecurityBilling.Find "BillTxnID =" & Val(MSHFlexGrid1.TextMatrix(I, 2))
       If Not RsTxn_SecurityBilling.EOF Then
          RsTxn_SecurityBilling!Bankname = Trim(MSHFlexGrid1.TextMatrix(I, 3))
          RsTxn_SecurityBilling!ChequeNo = Trim(MSHFlexGrid1.TextMatrix(I, 4))
          RsTxn_SecurityBilling!PaidAmount = Val(MSHFlexGrid1.TextMatrix(I, 5))
          RsTxn_SecurityBilling!TDS = Val(MSHFlexGrid1.TextMatrix(I, 6))
          RsTxn_SecurityBilling!PaidDate = CDate(MSHFlexGrid1.TextMatrix(I, 0))
          RsTxn_SecurityBilling!PaymentdetailsUpdatedBy = Gen_UserLoginID
          RsTxn_SecurityBilling!PaymentdetailsUpdatedOn = Now
          RsTxn_SecurityBilling.Update
          RsTxn_SecurityBilling.Requery
       End If
   Next
ElseIf OptFumigationDetails.Value = True Then
   Call TableTxn_FumigationApproval(" Where Flag in ('S', 'A') And BatchID > 0 ")
   For I = 1 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit For
       If RsTxn_FumigationApproval.RecordCount > 0 Then
          RsTxn_FumigationApproval.MoveFirst
       End If
       RsTxn_FumigationApproval.Find "MappingID =" & Val(MSHFlexGrid1.TextMatrix(I, 2))
       If Not RsTxn_FumigationApproval.EOF Then
          RsTxn_FumigationApproval!Bankname = Trim(MSHFlexGrid1.TextMatrix(I, 3))
          RsTxn_FumigationApproval!ChequeNo = Trim(MSHFlexGrid1.TextMatrix(I, 4))
          RsTxn_FumigationApproval!PaidAmount = Val(MSHFlexGrid1.TextMatrix(I, 5))
          RsTxn_FumigationApproval!TDS = Val(MSHFlexGrid1.TextMatrix(I, 6))
          RsTxn_FumigationApproval!PaidDate = CDate(MSHFlexGrid1.TextMatrix(I, 0))
          RsTxn_FumigationApproval!PaymentdetailsUpdatedBy = Gen_UserLoginID
          RsTxn_FumigationApproval!PaymentdetailsUpdatedOn = Now
          RsTxn_FumigationApproval.Update
          RsTxn_FumigationApproval.Requery
       End If
       
   Next
End If

MsgBox "Done !!"
Call Grid_Head

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetFile.Enabled = False
   CmdImport.Enabled = False
   CmdSave.Enabled = False
   CmdReset.Enabled = False
   Exit Sub
End If

Call CmdReset_Click
End Sub

Private Sub OptFranchiseeClaim_Click()
Call OptGodownRent_Click
End Sub

Private Sub OptFumigationDetails_Click()
Call OptGodownRent_Click
End Sub

Private Sub OptGodownRent_Click()
Call Grid_Head
End Sub

Private Sub OptSecurityDetails_Click()
Call OptGodownRent_Click
End Sub
