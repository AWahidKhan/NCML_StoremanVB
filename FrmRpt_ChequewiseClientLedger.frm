VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_ChequewiseClientLedger 
   Caption         =   "Cheque No wise Client Ledger"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.ComboBox CmbDocumentType 
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
         Left            =   3840
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   435
         Width           =   3750
      End
      Begin VB.TextBox TxtInstrumentNo 
         Appearance      =   0  'Flat
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
         Left            =   120
         MaxLength       =   25
         TabIndex        =   4
         Top             =   435
         Width           =   3675
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
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
         Left            =   11895
         TabIndex        =   2
         Top             =   428
         Width           =   1335
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
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
         Left            =   13230
         TabIndex        =   1
         Top             =   428
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Document Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4988
         TabIndex        =   9
         Top             =   135
         Width           =   1455
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
         Left            =   10560
         TabIndex        =   7
         Top             =   495
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Total Invoices Found :-"
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
         Left            =   7920
         TabIndex        =   6
         Top             =   488
         Width           =   2400
      End
      Begin VB.Label Label5 
         Caption         =   "Cheque No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1350
         TabIndex        =   5
         Top             =   135
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7695
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   13573
      _Version        =   393216
      AllowUserResizing=   3
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   8760
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmRpt_ChequewiseClientLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 12
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.FixedCols = 1
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "Customer No"
MSHFlexGrid1.TextMatrix(0, 1) = "Customer Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Invoice Amount"
MSHFlexGrid1.TextMatrix(0, 3) = "Narration"
MSHFlexGrid1.TextMatrix(0, 4) = "Actual Date Posted"
MSHFlexGrid1.TextMatrix(0, 5) = "Actual Time Posted"
MSHFlexGrid1.TextMatrix(0, 6) = "Entry No"
MSHFlexGrid1.TextMatrix(0, 7) = "EHOP Invoice No"
MSHFlexGrid1.TextMatrix(0, 8) = "Document Type"
MSHFlexGrid1.TextMatrix(0, 9) = "Document No"
MSHFlexGrid1.TextMatrix(0, 10) = "Voucher No"
MSHFlexGrid1.TextMatrix(0, 11) = "Posting Date"


MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 4000
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1400


MSHFlexGrid1.RowHeight(0) = 600

End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub CmdGetReport_Click()

ProgressBar1.Value = 0

Call Grid_Head

MsgBox "Wait till next message!!!"

If TxtInstrumentNo = "" Then
   MsgBox "Blank Cheque No not allowed!!!"
   TxtInstrumentNo.SetFocus
   Exit Sub
End If

tmp = "Select CustomerNo,CustomerName,InvoiceAmount,Narration,ActualDatePosted,ActualTimePosted,EntryNo,EHOPInvoiceNo,DocumentType,DocumentNo,VoucherNo,PostingDate from Txn_ClientLEdger Where Narration like '%" & TxtInstrumentNo & "%' "

If CmbDocumentType.Text <> "" Then
   tmp = tmp & " And DocumentType = '" & CmbDocumentType.Text & "'"
End If

Call Tmp1Table

ProgressBar1.Max = TmpRs1.RecordCount + 1

For i = 1 To TmpRs1.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs1.Fields(c)), "", TmpRs1.Fields(c))
    Next
    TmpRs1.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
Label2.Caption = TmpRs1.RecordCount

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

Call ClearData

End Sub
Private Sub ClearData()

TxtInstrumentNo = ""
Call Grid_Head
Label2.Caption = "_"
ProgressBar1.Value = 0
CmbDocumentType.Text = "Payment"

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ChequewiseClientLedger.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ChequewiseClientLedger.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub

Private Sub CmbDocumentType_GotFocus()

tmp1 = CmbDocumentType.Text

tmp = "Select Distinct DocumentType From Txn_ClientLedger "
Call TmpTable

CmbDocumentType.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To TmpRs.RecordCount
    CmbDocumentType.AddItem TmpRs!DocumentType
    TmpRs.MoveNext
Next
CmbDocumentType.Text = tmp1
End Sub

Private Sub CmbDocumentType_LostFocus()
If CmbDocumentType.Text = "" Then
   Exit Sub
End If
TmpRs.MoveFirst
TmpRs.Find "DocumentType = '" & CmbDocumentType.Text & "'"
If TmpRs.EOF Then
   MsgBox "Invalid selection "
   CmbDocumentType.SetFocus
   Exit Sub
End If

End Sub
