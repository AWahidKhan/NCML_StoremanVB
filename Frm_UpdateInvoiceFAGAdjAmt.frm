VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_UpdateInvoiceFAGAdjAmt 
   Caption         =   "Update Invoice FAG Adjustment Amount"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   10005
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   14985
      Begin VB.TextBox TxtErrMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1755
         Width           =   7935
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
         Height          =   375
         Left            =   9930
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
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
         Height          =   375
         Left            =   12360
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
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
         Height          =   375
         Left            =   13575
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
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
         Height          =   375
         Left            =   11145
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton CmdGo 
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
         Left            =   7500
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
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
         Left            =   8715
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtSTINo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2940
         TabIndex        =   3
         Top             =   495
         Width           =   2535
      End
      Begin VB.TextBox TxtSPostInvoiceNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5520
         TabIndex        =   4
         Top             =   495
         Width           =   2535
      End
      Begin VB.TextBox TxtSEntryNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   8100
         TabIndex        =   5
         Top             =   495
         Width           =   2655
      End
      Begin VB.ComboBox CmbSLocation 
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
         ItemData        =   "Frm_UpdateInvoiceFAGAdjAmt.frx":0000
         Left            =   10815
         List            =   "Frm_UpdateInvoiceFAGAdjAmt.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   495
         Width           =   3240
      End
      Begin VB.ComboBox CmbSEntryType 
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
         ItemData        =   "Frm_UpdateInvoiceFAGAdjAmt.frx":0004
         Left            =   120
         List            =   "Frm_UpdateInvoiceFAGAdjAmt.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   3240
      End
      Begin VB.ComboBox CmbSReverseFlag 
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
         ItemData        =   "Frm_UpdateInvoiceFAGAdjAmt.frx":0022
         Left            =   5160
         List            =   "Frm_UpdateInvoiceFAGAdjAmt.frx":002C
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2400
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12726
         _Version        =   393216
         WordWrap        =   -1  'True
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
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   345
         Left            =   135
         TabIndex        =   1
         Top             =   510
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
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
         Format          =   47448065
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   1590
         TabIndex        =   2
         Top             =   495
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Format          =   47448065
         CurrentDate     =   39884
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   90
         TabIndex        =   26
         Top             =   9720
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label8 
         Caption         =   "Invoices in red below are not saved due to incomplete data."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8805
         TabIndex        =   30
         Top             =   1815
         Width           =   5190
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
         Left            =   8160
         TabIndex        =   29
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   3600
         TabIndex        =   27
         Top             =   1050
         Width           =   705
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "TI No (Voucher No)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3345
         TabIndex        =   25
         Top             =   195
         Width           =   1725
      End
      Begin VB.Label Label16 
         Caption         =   "Total Records Found :"
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
         Left            =   11280
         TabIndex        =   24
         Top             =   1020
         Width           =   2415
      End
      Begin VB.Label Label15 
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
         Left            =   13920
         TabIndex        =   23
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   22
         Top             =   195
         Width           =   945
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1890
         TabIndex        =   21
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Post Invoice No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6082
         TabIndex        =   20
         Top             =   195
         Width           =   1410
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Entry No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9045
         TabIndex        =   19
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12053
         TabIndex        =   18
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Entry Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1253
         TabIndex        =   17
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Reverse Flag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5205
         TabIndex        =   16
         Top             =   1020
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_UpdateInvoiceFAGAdjAmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mLocationID As Variant
Dim mErrorMsg As String

Private Sub CmbSEntryType_LostFocus()

If CmbSEntryType.Text = "" Then Exit Sub

If LCase(CmbSEntryType.Text) <> "system" And LCase(CmbSEntryType.Text) <> "manual" Then
   MsgBox "Invalid Selection!!!"
   CmbSEntryType.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
'Call TableMst_Location(" Where LocationId > 0 ")

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSReverseFlag_LostFocus()
If CmbSReverseFlag.Text = "" Then Exit Sub
If LCase(CmbSReverseFlag.Text) <> "yes" And LCase(CmbSReverseFlag.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSReverseFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "UpdateInvoiceFAGAdjAmt.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "UpdateInvoiceFAGAdjAmt.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1 ' + 6
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
      MSHFlexGrid1.row = I
      MSHFlexGrid1.Col = C
      If MSHFlexGrid1.CellBackColor = vbRed Then
      'oWS.Range(oWS.Cells(mRow, c), oWS.Cells(mRow, c)).Interior.Color = &H8080FF
         oWS.Cells(mRow, C + 1).Interior.Color = vbRed
      End If
   Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetFile_Click()

On Error GoTo errortrap
CmdImport.Enabled = False
CmdSave.Enabled = False
With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx"
     .CancelError = True
procReOpen:
     .ShowOpen
End With

Path = CommonDialog1.FileName
LblFile.Caption = Path

If LblFile.Caption <> "" Then
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   Dim I As Integer
   Dim runApp As Boolean
   CmdImport.Enabled = True
   oXL.Quit
   MsgBox "Select Import To Import Data From Excel Sheet!!!"
   CmdImport.SetFocus
End If
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Grid_Head()

Label15.Caption = ""
TxtErrMsg.Text = ""
ProgressBar1.Value = 0
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 20

MSHFlexGrid1.TextMatrix(0, 0) = "TI ID"
MSHFlexGrid1.TextMatrix(0, 1) = "TI Date"
MSHFlexGrid1.TextMatrix(0, 2) = "TI No"
MSHFlexGrid1.TextMatrix(0, 3) = "Post Invoice No"
MSHFlexGrid1.TextMatrix(0, 4) = "Entry No"
MSHFlexGrid1.TextMatrix(0, 5) = "TI Amount"
MSHFlexGrid1.TextMatrix(0, 6) = "TI Service Tax Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "Received Amount"
MSHFlexGrid1.TextMatrix(0, 8) = "Received TDS"
MSHFlexGrid1.TextMatrix(0, 9) = "Received Other Ded"
MSHFlexGrid1.TextMatrix(0, 10) = "Client NAV Code"
MSHFlexGrid1.TextMatrix(0, 11) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 12) = "Entry Flag"
MSHFlexGrid1.TextMatrix(0, 13) = "Adjustment Entry by FAG"
MSHFlexGrid1.TextMatrix(0, 14) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 15) = "Invoice Type"
MSHFlexGrid1.TextMatrix(0, 16) = "SM ID"
MSHFlexGrid1.TextMatrix(0, 17) = "Reverse Flag"
MSHFlexGrid1.TextMatrix(0, 18) = "Adjustment Amount"
MSHFlexGrid1.TextMatrix(0, 19) = "Remarks"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1800
MSHFlexGrid1.ColWidth(9) = 1800
MSHFlexGrid1.ColWidth(10) = 3400
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 1400
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1400

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600

End Sub

Private Sub CmdGo_Click()

If TxtFromDate.Value > Date Then
    MsgBox "Future Date not Allowed!!!"
    TxtFromDate.SetFocus
    Exit Sub
End If
If TxtToDate.Value > Date Then
    MsgBox "Future Date not Allowed!!!"
    TxtToDate.SetFocus
    Exit Sub
End If

If TxtFromDate.Value > TxtToDate.Value Then
    MsgBox "From Date should not greater than To Date!!!"
    TxtToDate.SetFocus
    Exit Sub
End If

MsgBox "Wait till Next Message!!!"
Call Grid_Head

tmp = " Select TI.TIID,TI.TIDate,TI.TINo,TI.PostInvoiceNo,TI.EntryNo,TIAmount,TI.TISTAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDSAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecOtherDed"
tmp = tmp & " ,TI.NAVCode,ML.LocationName,TI.EntryFlag,TI.AdjAmountFAG,TI.Remarks,TI.PIFlag,TI.PIID,TI.ReverseFlag"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Where TINo <> '' "
tmp = tmp & " And TIAmount > TI.AdjAmountFAG+"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) +"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )+"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID ) "

If Gen_DBType = "MDB" Then
   tmp = tmp & " And TI.TIDate Between #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "# "
Else
   tmp = tmp & " And TI.TIDate Between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
End If

If TxtSTINo <> "" Then
   tmp = tmp & " And TI.TINO like '%" & TxtSTINo & "%'"
End If

If TxtSPostInvoiceNo <> "" Then
   tmp = tmp & " And TI.PostInvoiceNo like '%" & TxtSPostInvoiceNo & "%'"
End If

If TxtSEntryNo <> "" Then
   tmp = tmp & " And TI.EntryNo like '%" & TxtSEntryNo & "%'"
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " And ML.LocationName = '" & CmbSLocation.Text & "'"
End If

If CmbSEntryType.Text <> "" Then
   If LCase(CmbSEntryType.Text) = "system" Then
      tmp = tmp & " And (TI.EntryFlag = '" & CmbSEntryType.Text & "' Or TI.EntryFlag is Null) "
   Else
      tmp = tmp & " And TI.EntryFlag = '" & CmbSEntryType.Text & "'"
   End If
End If

If LCase(CmbSReverseFlag.Text) = "yes" Then
   tmp = tmp & " And isnull(TI.ReverseFlag,'N') = 'Y'"
ElseIf LCase(CmbSReverseFlag.Text) = "no" Then
   tmp = tmp & " And isnull(TI.ReverseFlag,'N') = 'N'"
End If

Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount + 2
ProgressBar1.Value = 0
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 3
       MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 1) = Format(MSHFlexGrid1.TextMatrix(iu, 1), "dd-mmm-yyyy")
    If MSHFlexGrid1.TextMatrix(iu, 12) = "" Then
       MSHFlexGrid1.TextMatrix(iu, 12) = "System"
    End If
    
    If MSHFlexGrid1.TextMatrix(iu, 17) = "" Or MSHFlexGrid1.TextMatrix(iu, 17) = "N" Then
       MSHFlexGrid1.TextMatrix(iu, 17) = "No"
    Else
       MSHFlexGrid1.TextMatrix(iu, 17) = "Yes"
    End If
    
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max
Label15.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdImport_Click()

MsgBox "Wait till next message!!!"
TxtErrorMsg = ""
ErrMsg = ""
Status = 0
RowCount = 0
Set oXL = New Excel.Application
Set oWB = oXL.Workbooks.Open(LblFile.Caption)
Set oWS = oWB.Worksheets(1)
Call Grid_Head
If CheckFile = False Then
   oXL.Quit
   Exit Sub
End If
Call FillData
CmdSave.Enabled = True
CmdImport.Enabled = False
CmdSave.SetFocus
MsgBox "Import Succesful!!!"
LblFile.Caption = ""
End Sub

Private Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True

If LCase(Trim(oWS.Cells(1, 1))) <> "ti id" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of A(TI ID) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 2))) <> "ti date" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of B(TI DATE) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 3))) <> "ti no" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of C(TI NO) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 4))) <> "post invoice no" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of D(POST INVOICE NO) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 5))) <> "entry no" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of E(ENTRY NO) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 6))) <> "ti amount" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of F(TI AMOUNT) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 7))) <> "ti service tax amount" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of G(TI SERVICE TAX AMOUNT) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 8))) <> "received amount" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of H(RECEIVED AMOUNT) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 9))) <> "received tds" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of I(RECEIVED TDS) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 10))) <> "received other ded" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of J(RECEIVED OTHER DED) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 11))) <> "client nav code" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of K(CLIENT NAV CODE) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 12))) <> "location name" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of L(LOCATION NAME) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 13))) <> "entry flag" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of M(ENTRY FLAG) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 14))) <> "adjustment entry by fag" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of N(ADJUSTMENT ENTRY BY FAG) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 15))) <> "remarks" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of O(REMARKS) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 16))) <> "invoice type" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of P(INVOICE TYPE) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 17))) <> "sm id" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of Q(SM ID) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 18))) <> "reverse flag" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of R(REVERSE FLAG) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 19))) <> "adjustment amount" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of S(ADJUSTMENT AMOUNT) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 20))) <> "remarks" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of T(REMARKS) "
   Retval = False
End If
CheckFile = Retval
End Function

Private Sub FillData()

'Set oWS = oWB.Worksheets(CmbSheetName.Text)
oWS.ClearArrows
mRow = 1
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
Label15.Caption = RowCount - 1
For l = 1 To RowCount - 1
    mRow = mRow + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(l, C) = Trim(oWS.Cells(mRow, C + 1))
    Next
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
Next
oXL.Quit
End Sub

Private Sub CmdReset_Click()
Call Grid_Head
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
TxtSTINo.Text = ""
TxtSPostInvoiceNo.Text = ""
TxtSEntryNo.Text = ""
CmbSEntryType.Text = ""
CmbSLocation.Text = ""
CmbSReverseFlag.Text = ""
CmdImport.Enabled = False
CmdSave.Enabled = False
LblFile.Caption = ""
End Sub

Private Sub CmdSave_Click()
MsgBox "Please Wait till Next Message!!!"
Dim mTIAmount As Double
Dim mReceiveAmount As Double

tmp = " Select TI.TIID, TI.EntryNo,TIAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDSAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Where TINo <> '' "
Call TmpTable

mErrorMsg = ""
Call TableTxn_TIDetails
If RsTxn_TIDetails.RecordCount > 0 Then
   For I = 1 To MSHFlexGrid1.Rows - 2
      If MSHFlexGrid1.TextMatrix(I, 18) <> "" And MSHFlexGrid1.TextMatrix(I, 19) <> "" Then
         RsTxn_TIDetails.MoveFirst
         RsTxn_TIDetails.Find "TIID=" & Val(MSHFlexGrid1.TextMatrix(I, 0))
         If Not RsTxn_TIDetails.EOF Then
            If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
            TmpRs.Find "TIID=" & Val(MSHFlexGrid1.TextMatrix(I, 0))
            If Not TmpRs.EOF Then
               mTIAmount = TmpRs!TIAmount
               mReceiveAmount = TmpRs!RecAmount + TmpRs!RecTDSAmount + TmpRs!RecOtherDed + Val(MSHFlexGrid1.TextMatrix(I, 18))
               If mTIAmount >= mReceiveAmount Then
                  RsTxn_TIDetails!AdjAmountFAG = RsTxn_TIDetails!AdjAmountFAG + Val(MSHFlexGrid1.TextMatrix(I, 18))
                  RsTxn_TIDetails!Remarks = MSHFlexGrid1.TextMatrix(I, 19)
                  RsTxn_TIDetails!UpdatedDate = Now
                  RsTxn_TIDetails!UpdatedBy = Gen_UserLoginID
                  RsTxn_TIDetails.Update
               Else
                  mErrorMsg = mErrorMsg & " Received Amount + Adjustment Amount is greater than TI Amount for row number " & I & " " & vbCrLf
                  MSHFlexGrid1.row = I
                  For j = 0 To MSHFlexGrid1.Cols - 1
                     MSHFlexGrid1.Col = j
                     MSHFlexGrid1.CellBackColor = vbRed
                  Next
               End If
            End If
         Else
            mErrorMsg = mErrorMsg & " TIID not found in Invoices for row number " & I & " " & vbCrLf
            MSHFlexGrid1.row = I
            For j = 0 To MSHFlexGrid1.Cols - 1
               MSHFlexGrid1.Col = j
               MSHFlexGrid1.CellBackColor = vbRed
            Next
         End If
      Else
         mErrorMsg = mErrorMsg & " Blank Adjustment Amount Or Remarks for row number " & I & " " & vbCrLf
         MSHFlexGrid1.row = I
         For j = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.Col = j
            MSHFlexGrid1.CellBackColor = vbRed
         Next
      End If
   Next
End If
TxtErrMsg.Text = mErrorMsg
CmdImport.Enabled = False
CmdSave.Enabled = False
MsgBox "Done!!!"
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame0.Enabled = False
   Exit Sub
End If
Frame0.Enabled = True
Call CmdReset_Click
End Sub

