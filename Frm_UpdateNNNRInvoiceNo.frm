VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_UpdateNNNRInvoiceNo 
   Caption         =   "Update NNNR / 7 Days Invoice No"
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
      Top             =   1200
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbInvoiceType 
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
         ItemData        =   "Frm_UpdateNNNRInvoiceNo.frx":0000
         Left            =   1560
         List            =   "Frm_UpdateNNNRInvoiceNo.frx":0016
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   255
         Width           =   2160
      End
      Begin VB.CommandButton CmdUpdateInvoice 
         Caption         =   "Update Invoice"
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
         Left            =   12600
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   2220
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
         Left            =   5520
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
         Left            =   10380
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
         Left            =   8760
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
         Left            =   7140
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
         Left            =   3885
         TabIndex        =   1
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Invoice Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   11
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   795
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
Attribute VB_Name = "Frm_UpdateNNNRInvoiceNo"
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

If CmbInvoiceType.Text = "" Then
   MsgBox "Please select Invoice Type"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

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
If CmbInvoiceType.Text = "" Then
   MsgBox "Please select Invoice Type"
   CmbInvoiceType.SetFocus
   Exit Sub
End If
Call ReadFiles(LblFile.Caption)
CmdUpdate.Enabled = True
End Sub

Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 31 '27
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Date"
    .TextMatrix(1, 1) = "Invoice"
    .TextMatrix(1, 2) = "Customer"
    .TextMatrix(1, 3) = "NavCode"
    .TextMatrix(1, 4) = "Amount"
    .TextMatrix(1, 5) = "G/L Account"
    .TextMatrix(1, 6) = "GL A/C NAVCode "
    .TextMatrix(1, 12) = "Narration"
    .TextMatrix(1, 13) = "Location NAV Code"
    .TextMatrix(1, 14) = "CMP"
    .TextMatrix(1, 18) = "PI No"
    .TextMatrix(1, 19) = "ST Code"
    .TextMatrix(1, 20) = "Agree/Non Agree"
    .TextMatrix(1, 21) = "GST Code"
    .TextMatrix(1, 22) = "HSN No"
    .TextMatrix(1, 23) = "Invoice Type"
    .TextMatrix(1, 24) = "Number Series Flag"
    .TextMatrix(1, 25) = "Invoice No"
    .TextMatrix(1, 26) = "Invoice Date"
    .TextMatrix(1, 27) = "ST Amt"
    .TextMatrix(1, 28) = "Invoice Amt"
    .TextMatrix(1, 29) = "PostInvoice No"
    .TextMatrix(1, 30) = "Entry No"
    
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
    .ColWidth(19) = 0
    .ColWidth(20) = 1200
    .ColWidth(21) = 1200
    .ColWidth(22) = 1200
    .ColWidth(23) = 1200
    .ColWidth(24) = 1200
    .ColWidth(25) = 1200
    .ColWidth(26) = 1200
    .ColWidth(27) = 1200
    .ColWidth(28) = 1200
    .ColWidth(29) = 1200
    .ColWidth(30) = 1200
End With

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
LblFile.Caption = ""
CmdUpdate.Enabled = False
CmdImport.Enabled = False
CmdGetFile.Enabled = True
CmbInvoiceType.Text = ""
End Sub

Private Sub CmdUpdate_Click()

If CmbInvoiceType.Text = "" Then
   MsgBox "Please select Invoice Type"
   CmbInvoiceType.SetFocus
   Exit Sub
End If
MsgBox "Please Wait till Next Message."

Call TableTxn_TIDetails '(" Where PIFlag='NNNR'")

If LCase(CmbInvoiceType.Text) = "nnnr" Then
    Call TableTxn_ProInvoice
    
    For I = 2 To MSHFlexGrid3.Rows - 1
        If MSHFlexGrid3.TextMatrix(I, 26) = "" Or MSHFlexGrid3.TextMatrix(I, 25) = "" Then Exit For 'If MSHFlexGrid3.TextMatrix(I, 22) = "" Or MSHFlexGrid3.TextMatrix(I, 21) = "" Then Exit For
        If RsTxn_ProInvoice.RecordCount > 0 Then
           RsTxn_ProInvoice.MoveFirst
        End If
        RsTxn_ProInvoice.Find "InvoiceNo=" & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "NN", "")) & ""
        If Not RsTxn_ProInvoice.EOF Then
           If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
           RsTxn_TIDetails.Find "EntryNo=" & Val(MSHFlexGrid3.TextMatrix(I, 30)) 'RsTxn_TIDetails.Find "PIID=" & RsTxn_ProInvoice!InvoiceNo
           If RsTxn_TIDetails.EOF = True Then
              RsTxn_TIDetails.AddNew
              RsTxn_TIDetails!TIID = getTIDetailMax
              RsTxn_TIDetails!TINo = Trim(MSHFlexGrid3.TextMatrix(I, 25)) 'Trim(MSHFlexGrid3.TextMatrix(I, 21))
              RsTxn_TIDetails!TIDate = Format(MSHFlexGrid3.TextMatrix(I, 26), Gen_DatFormat) 'Format(MSHFlexGrid3.TextMatrix(I, 22), Gen_DatFormat)
              RsTxn_TIDetails!TIAmount = Val(MSHFlexGrid3.TextMatrix(I, 28)) 'Val(MSHFlexGrid3.TextMatrix(I, 24))
              RsTxn_TIDetails!TISTAmount = Val(MSHFlexGrid3.TextMatrix(I, 27)) 'Val(MSHFlexGrid3.TextMatrix(I, 23))
              RsTxn_TIDetails!PostInvoiceNo = Trim(MSHFlexGrid3.TextMatrix(I, 29)) 'Trim(MSHFlexGrid3.TextMatrix(I, 25))
              RsTxn_TIDetails!EntryNo = Val(MSHFlexGrid3.TextMatrix(I, 30)) 'Val(MSHFlexGrid3.TextMatrix(I, 26))
              RsTxn_TIDetails!PIID = RsTxn_ProInvoice!InvoiceNo
              RsTxn_TIDetails!LocationID = RsTxn_ProInvoice!LocationID
              RsTxn_TIDetails!ClientIDRow = RsTxn_ProInvoice!ClientIDRow
              RsTxn_TIDetails!NAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
              RsTxn_TIDetails!ExchangeTypeID = RsTxn_ProInvoice!ExchangeTypeID
              RsTxn_TIDetails!RecAmount = 0
              RsTxn_TIDetails!RecTDSAmount = 0
              RsTxn_TIDetails!RecOtherDed = 0
              RsTxn_TIDetails!PIFlag = "NNNR"
              RsTxn_TIDetails!SMID = MSHFlexGrid3.TextMatrix(I, 18) '"NN" & RsTxn_ProInvoice!InvoiceNo
    '          RsTxn_TIDetails!EntryFlag = "System"
              RsTxn_TIDetails!G_UID = GetGUID
              RsTxn_TIDetails!CreatedDate = Now
              RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
              RsTxn_TIDetails.Update
           End If
        End If
    Next

ElseIf LCase(CmbInvoiceType.Text) = "7 days" Then
   
    Call TableTxn_CDTFBOE
    For I = 2 To MSHFlexGrid3.Rows - 1
        If MSHFlexGrid3.TextMatrix(I, 26) = "" Or MSHFlexGrid3.TextMatrix(I, 25) = "" Then Exit For 'If MSHFlexGrid3.TextMatrix(I, 22) = "" Or MSHFlexGrid3.TextMatrix(I, 21) = "" Then Exit For
        If RsTxn_CDTFBOE.RecordCount > 0 Then
           RsTxn_CDTFBOE.MoveFirst
        End If
        RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "SD", "")) & ""
        If Not RsTxn_CDTFBOE.EOF Then
           If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
           RsTxn_TIDetails.Find "EntryNo=" & Val(MSHFlexGrid3.TextMatrix(I, 30)) 'RsTxn_TIDetails.Find "PIID=" & RsTxn_ProInvoice!InvoiceNo
           If RsTxn_TIDetails.EOF = True Then
              RsTxn_TIDetails.AddNew
              RsTxn_TIDetails!TIID = getTIDetailMax
              RsTxn_TIDetails!TINo = Trim(MSHFlexGrid3.TextMatrix(I, 25)) 'Trim(MSHFlexGrid3.TextMatrix(I, 21))
              RsTxn_TIDetails!TIDate = Format(MSHFlexGrid3.TextMatrix(I, 26), Gen_DatFormat) 'Format(MSHFlexGrid3.TextMatrix(I, 22), Gen_DatFormat)
              RsTxn_TIDetails!TIAmount = Val(MSHFlexGrid3.TextMatrix(I, 28)) 'Val(MSHFlexGrid3.TextMatrix(I, 24))
              RsTxn_TIDetails!TISTAmount = Val(MSHFlexGrid3.TextMatrix(I, 27)) 'Val(MSHFlexGrid3.TextMatrix(I, 23))
              RsTxn_TIDetails!PostInvoiceNo = Trim(MSHFlexGrid3.TextMatrix(I, 29)) 'Trim(MSHFlexGrid3.TextMatrix(I, 25))
              RsTxn_TIDetails!EntryNo = Val(MSHFlexGrid3.TextMatrix(I, 30)) 'Val(MSHFlexGrid3.TextMatrix(I, 26))
              RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
              RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
              RsTxn_TIDetails!ClientIDRow = RsTxn_CDTFBOE!ClientIDRow
              RsTxn_TIDetails!NAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
              RsTxn_TIDetails!ExchangeTypeID = 2 'RsTxn_CDTFBOE!ExchangeTypeID
              RsTxn_TIDetails!RecAmount = 0
              RsTxn_TIDetails!RecTDSAmount = 0
              RsTxn_TIDetails!RecOtherDed = 0
              RsTxn_TIDetails!PIFlag = "SD"
              RsTxn_TIDetails!SMID = MSHFlexGrid3.TextMatrix(I, 18) '"NN" & RsTxn_ProInvoice!InvoiceNo
    '          RsTxn_TIDetails!EntryFlag = "System"
              RsTxn_TIDetails!G_UID = GetGUID
              RsTxn_TIDetails!CreatedDate = Now
              RsTxn_TIDetails!CreatedBy = Gen_UserLoginID

              RsTxn_TIDetails.Update
           End If
        End If
    Next
ElseIf LCase(CmbInvoiceType.Text) = "rejected" Then
   
    Call TableTxn_CDTFBOE
    For I = 2 To MSHFlexGrid3.Rows - 1
        If MSHFlexGrid3.TextMatrix(I, 26) = "" Or MSHFlexGrid3.TextMatrix(I, 25) = "" Then Exit For 'If MSHFlexGrid3.TextMatrix(I, 22) = "" Or MSHFlexGrid3.TextMatrix(I, 21) = "" Then Exit For
        If RsTxn_CDTFBOE.RecordCount > 0 Then
           RsTxn_CDTFBOE.MoveFirst
        End If
        RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "RJ", "")) & ""
        If Not RsTxn_CDTFBOE.EOF Then
           If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
           RsTxn_TIDetails.Find "EntryNo=" & Val(MSHFlexGrid3.TextMatrix(I, 30)) 'RsTxn_TIDetails.Find "PIID=" & RsTxn_ProInvoice!InvoiceNo
           If RsTxn_TIDetails.EOF = True Then
              RsTxn_TIDetails.AddNew
              RsTxn_TIDetails!TIID = getTIDetailMax
              RsTxn_TIDetails!TINo = Trim(MSHFlexGrid3.TextMatrix(I, 25)) 'Trim(MSHFlexGrid3.TextMatrix(I, 21))
              RsTxn_TIDetails!TIDate = Format(MSHFlexGrid3.TextMatrix(I, 26), Gen_DatFormat) 'Format(MSHFlexGrid3.TextMatrix(I, 22), Gen_DatFormat)
              RsTxn_TIDetails!TIAmount = Val(MSHFlexGrid3.TextMatrix(I, 28)) 'Val(MSHFlexGrid3.TextMatrix(I, 24))
              RsTxn_TIDetails!TISTAmount = Val(MSHFlexGrid3.TextMatrix(I, 27)) 'Val(MSHFlexGrid3.TextMatrix(I, 23))
              RsTxn_TIDetails!PostInvoiceNo = Trim(MSHFlexGrid3.TextMatrix(I, 29)) 'Trim(MSHFlexGrid3.TextMatrix(I, 25))
              RsTxn_TIDetails!EntryNo = Val(MSHFlexGrid3.TextMatrix(I, 30)) 'Val(MSHFlexGrid3.TextMatrix(I, 26))
              
              RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
              RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
              RsTxn_TIDetails!ClientIDRow = RsTxn_CDTFBOE!ClientIDRow
              RsTxn_TIDetails!NAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
              RsTxn_TIDetails!ExchangeTypeID = 2 ' RsTxn_CDTFBOE!ExchangeTypeID
              RsTxn_TIDetails!RecAmount = 0
              RsTxn_TIDetails!RecTDSAmount = 0
              RsTxn_TIDetails!RecOtherDed = 0
              RsTxn_TIDetails!PIFlag = "RJ"
              RsTxn_TIDetails!SMID = MSHFlexGrid3.TextMatrix(I, 18) '"NN" & RsTxn_ProInvoice!InvoiceNo
    '          RsTxn_TIDetails!EntryFlag = "System"
              RsTxn_TIDetails!G_UID = GetGUID
              RsTxn_TIDetails!CreatedDate = Now
              RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
              RsTxn_TIDetails.Update
           End If
        End If
    Next
ElseIf LCase(CmbInvoiceType.Text) = "rrn" Then
    Call TableTxn_CDTFBOE
    For I = 2 To MSHFlexGrid3.Rows - 1
        If MSHFlexGrid3.TextMatrix(I, 26) = "" Or MSHFlexGrid3.TextMatrix(I, 25) = "" Then Exit For 'If MSHFlexGrid3.TextMatrix(I, 22) = "" Or MSHFlexGrid3.TextMatrix(I, 21) = "" Then Exit For
        If RsTxn_CDTFBOE.RecordCount > 0 Then
           RsTxn_CDTFBOE.MoveFirst
        End If
        RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "RRN", "")) & ""
        If Not RsTxn_CDTFBOE.EOF Then
           If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
           RsTxn_TIDetails.Find "EntryNo=" & Val(MSHFlexGrid3.TextMatrix(I, 30)) 'RsTxn_TIDetails.Find "PIID=" & RsTxn_ProInvoice!InvoiceNo
           If RsTxn_TIDetails.EOF = True Then
              RsTxn_TIDetails.AddNew
              RsTxn_TIDetails!TIID = getTIDetailMax
              RsTxn_TIDetails!TINo = Trim(MSHFlexGrid3.TextMatrix(I, 25)) 'Trim(MSHFlexGrid3.TextMatrix(I, 21))
              RsTxn_TIDetails!TIDate = Format(MSHFlexGrid3.TextMatrix(I, 26), Gen_DatFormat) 'Format(MSHFlexGrid3.TextMatrix(I, 22), Gen_DatFormat)
              RsTxn_TIDetails!TIAmount = Val(MSHFlexGrid3.TextMatrix(I, 28)) 'Val(MSHFlexGrid3.TextMatrix(I, 24))
              RsTxn_TIDetails!TISTAmount = Val(MSHFlexGrid3.TextMatrix(I, 27)) 'Val(MSHFlexGrid3.TextMatrix(I, 23))
              RsTxn_TIDetails!PostInvoiceNo = Trim(MSHFlexGrid3.TextMatrix(I, 29)) 'Trim(MSHFlexGrid3.TextMatrix(I, 25))
              RsTxn_TIDetails!EntryNo = Val(MSHFlexGrid3.TextMatrix(I, 30)) 'Val(MSHFlexGrid3.TextMatrix(I, 26))
              RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
              RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
              RsTxn_TIDetails!ClientIDRow = RsTxn_CDTFBOE!ClientIDRow
              RsTxn_TIDetails!NAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
              RsTxn_TIDetails!ExchangeTypeID = 2 ' RsTxn_CDTFBOE!ExchangeTypeID
              RsTxn_TIDetails!RecAmount = 0
              RsTxn_TIDetails!RecTDSAmount = 0
              RsTxn_TIDetails!RecOtherDed = 0
              RsTxn_TIDetails!PIFlag = "RRN"
              RsTxn_TIDetails!SMID = MSHFlexGrid3.TextMatrix(I, 18) '"NN" & RsTxn_ProInvoice!InvoiceNo
    '          RsTxn_TIDetails!EntryFlag = "System"
              RsTxn_TIDetails!G_UID = GetGUID
              RsTxn_TIDetails!CreatedDate = Now
              RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
              RsTxn_TIDetails.Update
           End If
        End If
    Next

ElseIf LCase(CmbInvoiceType.Text) = "nnnryearly" Then
   
    Call TableTxn_CDTFBOE
    For I = 2 To MSHFlexGrid3.Rows - 1
        If MSHFlexGrid3.TextMatrix(I, 26) = "" Or MSHFlexGrid3.TextMatrix(I, 25) = "" Then Exit For 'If MSHFlexGrid3.TextMatrix(I, 22) = "" Or MSHFlexGrid3.TextMatrix(I, 21) = "" Then Exit For
        If RsTxn_CDTFBOE.RecordCount > 0 Then
           RsTxn_CDTFBOE.MoveFirst
        End If
        RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "NNNRY", "")) & ""
        If Not RsTxn_CDTFBOE.EOF Then
           If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
           RsTxn_TIDetails.Find "EntryNo=" & Val(MSHFlexGrid3.TextMatrix(I, 30)) 'RsTxn_TIDetails.Find "PIID=" & RsTxn_ProInvoice!InvoiceNo
           If RsTxn_TIDetails.EOF = True Then
              RsTxn_TIDetails.AddNew
              RsTxn_TIDetails!TIID = getTIDetailMax
              RsTxn_TIDetails!TINo = Trim(MSHFlexGrid3.TextMatrix(I, 25)) 'Trim(MSHFlexGrid3.TextMatrix(I, 21))
              RsTxn_TIDetails!TIDate = Format(MSHFlexGrid3.TextMatrix(I, 26), Gen_DatFormat) 'Format(MSHFlexGrid3.TextMatrix(I, 22), Gen_DatFormat)
              RsTxn_TIDetails!TIAmount = Val(MSHFlexGrid3.TextMatrix(I, 28)) 'Val(MSHFlexGrid3.TextMatrix(I, 24))
              RsTxn_TIDetails!TISTAmount = Val(MSHFlexGrid3.TextMatrix(I, 27)) 'Val(MSHFlexGrid3.TextMatrix(I, 23))
              RsTxn_TIDetails!PostInvoiceNo = Trim(MSHFlexGrid3.TextMatrix(I, 29)) 'Trim(MSHFlexGrid3.TextMatrix(I, 25))
              RsTxn_TIDetails!EntryNo = Val(MSHFlexGrid3.TextMatrix(I, 30)) 'Val(MSHFlexGrid3.TextMatrix(I, 26))
              RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
              RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
              RsTxn_TIDetails!ClientIDRow = RsTxn_CDTFBOE!ClientIDRow
              RsTxn_TIDetails!NAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
              RsTxn_TIDetails!ExchangeTypeID = 1 'RsTxn_CDTFBOE!ExchangeTypeID
              RsTxn_TIDetails!RecAmount = 0
              RsTxn_TIDetails!RecTDSAmount = 0
              RsTxn_TIDetails!RecOtherDed = 0
              RsTxn_TIDetails!PIFlag = "NNNRY"
              RsTxn_TIDetails!SMID = MSHFlexGrid3.TextMatrix(I, 18) '"NN" & RsTxn_ProInvoice!InvoiceNo
    '          RsTxn_TIDetails!EntryFlag = "System"
              RsTxn_TIDetails!G_UID = GetGUID
              RsTxn_TIDetails!CreatedDate = Now
              RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
              RsTxn_TIDetails.Update
           End If
        End If
    Next
ElseIf LCase(CmbInvoiceType.Text) = "pwr" Then
   
    Call TableTxn_CDTFBOE
    For I = 2 To MSHFlexGrid3.Rows - 1
        If MSHFlexGrid3.TextMatrix(I, 26) = "" Or MSHFlexGrid3.TextMatrix(I, 25) = "" Then Exit For 'If MSHFlexGrid3.TextMatrix(I, 22) = "" Or MSHFlexGrid3.TextMatrix(I, 21) = "" Then Exit For
        If RsTxn_CDTFBOE.RecordCount > 0 Then
           RsTxn_CDTFBOE.MoveFirst
        End If
        RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "PWR", "")) & ""
        If Not RsTxn_CDTFBOE.EOF Then
           If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
           RsTxn_TIDetails.Find "EntryNo=" & Val(MSHFlexGrid3.TextMatrix(I, 30)) 'RsTxn_TIDetails.Find "PIID=" & RsTxn_ProInvoice!InvoiceNo
           If RsTxn_TIDetails.EOF = True Then
              RsTxn_TIDetails.AddNew
              RsTxn_TIDetails!TIID = getTIDetailMax
              RsTxn_TIDetails!TINo = Trim(MSHFlexGrid3.TextMatrix(I, 25)) 'Trim(MSHFlexGrid3.TextMatrix(I, 21))
              RsTxn_TIDetails!TIDate = Format(MSHFlexGrid3.TextMatrix(I, 26), Gen_DatFormat) 'Format(MSHFlexGrid3.TextMatrix(I, 22), Gen_DatFormat)
              RsTxn_TIDetails!TIAmount = Val(MSHFlexGrid3.TextMatrix(I, 28)) 'Val(MSHFlexGrid3.TextMatrix(I, 24))
              RsTxn_TIDetails!TISTAmount = Val(MSHFlexGrid3.TextMatrix(I, 27)) 'Val(MSHFlexGrid3.TextMatrix(I, 23))
              RsTxn_TIDetails!PostInvoiceNo = Trim(MSHFlexGrid3.TextMatrix(I, 29)) 'Trim(MSHFlexGrid3.TextMatrix(I, 25))
              RsTxn_TIDetails!EntryNo = Val(MSHFlexGrid3.TextMatrix(I, 30)) 'Val(MSHFlexGrid3.TextMatrix(I, 26))
              RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
              RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
              RsTxn_TIDetails!ClientIDRow = RsTxn_CDTFBOE!ClientIDRow
              RsTxn_TIDetails!NAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
              RsTxn_TIDetails!ExchangeTypeID = 1 'RsTxn_CDTFBOE!ExchangeTypeID
              RsTxn_TIDetails!RecAmount = 0
              RsTxn_TIDetails!RecTDSAmount = 0
              RsTxn_TIDetails!RecOtherDed = 0
              RsTxn_TIDetails!PIFlag = "PWR"
              RsTxn_TIDetails!SMID = MSHFlexGrid3.TextMatrix(I, 18) '"NN" & RsTxn_ProInvoice!InvoiceNo
    '          RsTxn_TIDetails!EntryFlag = "System"
              RsTxn_TIDetails!G_UID = GetGUID
              RsTxn_TIDetails!CreatedDate = Now
              RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
              RsTxn_TIDetails.Update
           End If
        End If
    Next
End If
MsgBox "Done !!"
End Sub

Private Sub CmdUpdateInvoice_Click()

Dim mInvoiceID As Double

MsgBox " Please wait till next message!!"

tmp = "Update Txn_TIDetails set PIFlag='NNNR' ,PIID=convert (decimal,replace(SMID,'NN','')) Where SMID like 'NN%'"

Call TmpTable

tmp = "Update TTI Set TTI.LocationID= TTD.LocationID,TTI.ExchangeTypeID=TTD.ExchangeTypeID from Txn_TIDetails TTI"
tmp = tmp & " inner join Txn_ProInvoice TTD on TTD.InvoiceNo=TTI.PIID And TTI.PIFlag='NNNR'"

Call TmpTable

tmp = "Update Txn_TIDetails set PIFlag='Reservation' ,PIID=convert (decimal,replace(SMID,'RI','')) Where SMID like 'RI%'"

Call TmpTable

tmp = "Update TTI Set TTI.LocationID= TTD.LocationID,TTI.ExchangeTypeID=1 from Txn_TIDetails TTI"
tmp = tmp & " inner join Txn_ReservationInvoice TTD on TTD.InvoiceNo=TTI.PIID and TTI.PIFlag='Reservation'"

Call TmpTable

tmp = "Select * from Txn_TIDetails Where SMID like '%DP%' or"
tmp = tmp & " SMID in ((Select SMID from Txn_TIDetails Where  isnumeric(SMID)=1 and convert (decimal,SMID)<4903))"
tmp = tmp & " order By SMID"

Call TmpTable

Call TableMst_DPInvoiceUpload

For I = 1 To TmpRs.RecordCount
    If Val(TmpRs!SMID) > 0 Then
       If InStr(1, TmpRs!NAVCode, "CSWD") Then
          mInvoiceID = Val(Replace(TmpRs!SMID, "DP", ""))
          If RsMst_DPInvoiceUpload.RecordCount > 0 Then RsMst_DPInvoiceUpload.MoveFirst
          RsMst_DPInvoiceUpload.Find "DPInvoiceID=" & mInvoiceID
          If RsMst_DPInvoiceUpload.EOF = False Then
             RsMst_DPInvoiceUpload!InvoiceNo = TmpRs!TINo
             RsMst_DPInvoiceUpload!InvoiceDate = TmpRs!TIDate
             RsMst_DPInvoiceUpload!InvSTAmt = TmpRs!TISTAmount
             RsMst_DPInvoiceUpload!InvoiceAmt = TmpRs!TIAmount
             RsMst_DPInvoiceUpload.Update

             TmpRs!PIID = RsMst_DPInvoiceUpload!DPInvoiceID
             TmpRs!PIFlag = "DP"
             TmpRs!LocationID = RsMst_DPInvoiceUpload!LocationID
             TmpRs!ExchangeTypeID = 2
             TmpRs.Update
             
          End If
       End If
    End If
    TmpRs.MoveNext
Next
MsgBox "Done!!"


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
'CmdUpdateInvoice.Visible = False
'If LCase(Gen_UserRole) = "power" Then
'   CmdUpdateInvoice.Visible = True
'End If

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
    If UBound(ps) <> 30 Then 'If UBound(ps) <> 26 Then
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



Private Sub CmbInvoiceType_LostFocus()
If CmbInvoiceType.Text = "" Then Exit Sub

If LCase(CmbInvoiceType.Text) <> "nnnr" And LCase(CmbInvoiceType.Text) <> "7 days" And LCase(CmbInvoiceType.Text) <> "rejected" And LCase(CmbInvoiceType.Text) <> "rrn" And LCase(CmbInvoiceType.Text) <> "nnnryearly" And LCase(CmbInvoiceType.Text) <> "pwr" Then
   MsgBox "Invalid Selection!!!"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

End Sub

