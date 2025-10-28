VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Frm_UploadTIInvoicebyFAG 
   Caption         =   "Upload TI Invoice by FAG"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   10005
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   14985
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
         Left            =   12015
         TabIndex        =   4
         Top             =   840
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
         Left            =   9585
         TabIndex        =   2
         Top             =   840
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
         Left            =   13230
         TabIndex        =   5
         Top             =   840
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
         Left            =   10800
         TabIndex        =   3
         Top             =   840
         Width           =   1215
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
         Left            =   8370
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
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
         Height          =   960
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   7215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1680
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13996
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   9720
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Error Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   90
         TabIndex        =   14
         Top             =   773
         Width           =   975
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
         Left            =   13200
         TabIndex        =   13
         Top             =   540
         Width           =   135
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
         Left            =   10680
         TabIndex        =   12
         Top             =   540
         Width           =   2415
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   705
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
         Left            =   8400
         TabIndex        =   10
         Top             =   1320
         Width           =   495
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
         Left            =   9045
         TabIndex        =   9
         Top             =   1335
         Width           =   5190
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
Attribute VB_Name = "Frm_UploadTIInvoicebyFAG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mLocationID, mClientID, mExchangeTypeID, mServiceTaxID As Variant
Dim mErrorMsg As String

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "UploadTIInvoicebyFAG.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "UploadTIInvoicebyFAG.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
   mRow = i + 1 ' + 6
   For c = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
      MSHFlexGrid1.Row = i
      MSHFlexGrid1.Col = c
      If MSHFlexGrid1.CellBackColor = vbRed Then
         oWS.Cells(mRow, c + 1).Interior.Color = vbRed
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
   Dim i As Integer
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
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 11
   
   .TextMatrix(0, 0) = "TI DATE"
   .TextMatrix(0, 1) = "TI NO"
   .TextMatrix(0, 2) = "POST INVOICE NO"
   .TextMatrix(0, 3) = "ENTRY NO"
   .TextMatrix(0, 4) = "TI AMOUNT"
   .TextMatrix(0, 5) = "TI SERVICE TAX AMOUNT"
   .TextMatrix(0, 6) = "CLIENT/DEPOSITOR"
   .TextMatrix(0, 7) = "CLIENT NAVCODE"
   .TextMatrix(0, 8) = "LOCATION NAME"
   .TextMatrix(0, 9) = "EXCHANGE TYPE"
   .TextMatrix(0, 10) = "SERVICE TAX CODE"
   
   .ColWidth(0) = 1500
   .ColWidth(1) = 1500
   .ColWidth(2) = 1500
   .ColWidth(3) = 1500
   .ColWidth(4) = 1500
   .ColWidth(5) = 1500
   .ColWidth(6) = 1500
   .ColWidth(7) = 1500
   .ColWidth(8) = 1500
   .ColWidth(9) = 1500
   .ColWidth(10) = 1500
   
   .WordWrap = True
   .RowHeight(0) = 600
End With

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

If LCase(Trim(oWS.Cells(1, 1))) <> "ti date" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of A(TI DATE) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 2))) <> "ti no" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of B(TI NO) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 3))) <> "post invoice no" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of C(POST INVOICE NO) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 4))) <> "entry no" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of D(ENTRY NO) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 5))) <> "ti amount" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of E(TI AMOUNT) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 6))) <> "ti service tax amount" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of F(TI SERVICE TAX AMOUNT) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 7))) <> "client/depositor" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of G(CLIENT/DEPOSITOR) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 8))) <> "client navcode" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of H(CLIENT NAVCODE) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 9))) <> "location name" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of I(LOCATION NAME) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 10))) <> "exchange type" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of J(EXCHANGE TYPE) "
   Retval = False
ElseIf LCase(Trim(oWS.Cells(1, 11))) <> "service tax code" Then
   MsgBox "Please Check Sheet " & oWS.Name & " Column Name of K(SERVICE TAX CODE) "
   Retval = False
End If
CheckFile = Retval
End Function

Private Sub FillData()

'Set oWS = oWB.Worksheets(CmbSheetName.Text)
oWS.ClearArrows
mRow = 1
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row
Label15.Caption = RowCount - 1
For L = 1 To RowCount - 1
    mRow = mRow + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(L, c) = Trim(oWS.Cells(mRow, c + 1))
    Next
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
Next
oXL.Quit
End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmdImport.Enabled = False
CmdSave.Enabled = False
LblFile.Caption = ""
End Sub

Private Function CheckCorrectData(i_ As Long) As Boolean
CheckCorrectData = True
For j = 1 To MSHFlexGrid1.Cols - 1
   If Trim(MSHFlexGrid1.TextMatrix(i_, j)) = "" Then
      CheckCorrectData = False
      mErrorMsg = mErrorMsg & " All Data are not filled at Row " & i_ & vbCrLf
      Exit Function
   End If
Next
If IsDate(MSHFlexGrid1.TextMatrix(i_, 0)) = False Then
   CheckCorrectData = False
   mErrorMsg = mErrorMsg & " TIDate is not in date format at Row " & i_ & vbCrLf
   Exit Function
End If
If RsTxn_TIDetails.RecordCount > 0 Then
   RsTxn_TIDetails.MoveFirst
   RsTxn_TIDetails.Find "TINo='" & MSHFlexGrid1.TextMatrix(i_, 1) & "'"
   If Not RsTxn_TIDetails.EOF Then
'      CheckCorrectData = False
      mErrorMsg = mErrorMsg & " Duplicate TINo at Row " & i_ & vbCrLf
'      Exit Function
   End If
   
   RsTxn_TIDetails.MoveFirst
   RsTxn_TIDetails.Find "PostInvoiceNo='" & MSHFlexGrid1.TextMatrix(i_, 2) & "'"
   If Not RsTxn_TIDetails.EOF Then
'      CheckCorrectData = False
      mErrorMsg = mErrorMsg & " Duplicate PostInvoiceNo at Row " & i_ & vbCrLf
'      Exit Function
   End If
   
   RsTxn_TIDetails.MoveFirst
   RsTxn_TIDetails.Find "EntryNo='" & MSHFlexGrid1.TextMatrix(i_, 3) & "'"
   If Not RsTxn_TIDetails.EOF Then
      CheckCorrectData = False
      mErrorMsg = mErrorMsg & " Duplicate EntryNo at Row " & i_ & vbCrLf
      Exit Function
   End If
End If

If IsNumeric(MSHFlexGrid1.TextMatrix(i_, 4)) = False Then
   CheckCorrectData = False
   mErrorMsg = mErrorMsg & " TI Amount is not numeric at Row " & i_ & vbCrLf
   Exit Function
Else
   If Val(MSHFlexGrid1.TextMatrix(i_, 4)) < 0 Then
      CheckCorrectData = False
      mErrorMsg = mErrorMsg & " TI Amount is negative at Row " & i_ & vbCrLf
      Exit Function
   End If
End If

If IsNumeric(MSHFlexGrid1.TextMatrix(i_, 5)) = False Then
   CheckCorrectData = False
   mErrorMsg = mErrorMsg & " TI ServiceTax Amount is not numeric at Row " & i_ & vbCrLf
   Exit Function
Else
   If Val(MSHFlexGrid1.TextMatrix(i_, 5)) < 0 Then
      CheckCorrectData = False
      mErrorMsg = mErrorMsg & " TI ServiceTax Amount is negative at Row " & i_ & vbCrLf
      Exit Function
   End If
End If

If LCase(Trim(MSHFlexGrid1.TextMatrix(i_, 6))) <> "client" And LCase(Trim(MSHFlexGrid1.TextMatrix(i_, 6))) <> "depositor" Then
   CheckCorrectData = False
   mErrorMsg = mErrorMsg & " Column Client/Depositor must be 'Client' or 'Depositor' at Row " & i_ & vbCrLf
   Exit Function
ElseIf LCase(Trim(MSHFlexGrid1.TextMatrix(i_, 6))) = "client" Then
   If RsMst_Client.RecordCount > 0 Then
   RsMst_Client.MoveFirst
   RsMst_Client.Find "NAVCode='" & MSHFlexGrid1.TextMatrix(i_, 7) & "'"
      If Not RsMst_Client.EOF Then
         mClientID = RsMst_Client!ClientIDRow
      Else
         CheckCorrectData = False
         mErrorMsg = mErrorMsg & " Client NAVCode not found in Client Master at Row " & i_ & vbCrLf
         Exit Function
      End If
   End If
ElseIf LCase(Trim(MSHFlexGrid1.TextMatrix(i_, 6))) = "depositor" Then
   If RsMst_Depository.RecordCount > 0 Then
   RsMst_Depository.MoveFirst
   RsMst_Depository.Find "DepositoryName='" & MSHFlexGrid1.TextMatrix(i_, 7) & "'"
      If Not RsMst_Location.EOF Then
         mClientID = RsMst_Depository!DepositoryID
      Else
         CheckCorrectData = False
         mErrorMsg = mErrorMsg & " Depositor NAVCode not found in Depository Master at Row " & i_ & vbCrLf
         Exit Function
      End If
   End If
End If

If RsMst_Location.RecordCount > 0 Then
   RsMst_Location.MoveFirst
   RsMst_Location.Find "LocationName='" & MSHFlexGrid1.TextMatrix(i_, 8) & "'"
   If Not RsMst_Location.EOF Then
      mLocationID = RsMst_Location!LocationID
   Else
      CheckCorrectData = False
      mErrorMsg = mErrorMsg & " LocationName not found in Location Master at Row " & i_ & vbCrLf
      Exit Function
   End If
End If

If RsMst_ExchangeType.RecordCount > 0 Then
   RsMst_ExchangeType.MoveFirst
   RsMst_ExchangeType.Find "ExchangeType='" & MSHFlexGrid1.TextMatrix(i_, 9) & "'"
   If Not RsMst_ExchangeType.EOF Then
      mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
   Else
      CheckCorrectData = False
      mErrorMsg = mErrorMsg & " ExchangeType not found in ExchangeType Master at Row " & i_ & vbCrLf
      Exit Function
   End If
End If

If RsMst_ServiceTax.RecordCount > 0 Then
   RsMst_ServiceTax.MoveFirst
   RsMst_ServiceTax.Find "Description='" & MSHFlexGrid1.TextMatrix(i_, 10) & "'"
   If Not RsMst_ServiceTax.EOF Then
      mServiceTaxID = RsMst_ServiceTax!ServiceTaxID
   Else
      CheckCorrectData = False
      mErrorMsg = mErrorMsg & " ServiceTax not found in ServiceTax Master at Row " & i_ & vbCrLf
      Exit Function
   End If
End If
End Function

Private Sub CmdSave_Click()
MsgBox "Please Wait till Next Message!!!"
mErrorMsg = ""
Dim i As Long
Call TableMst_Location
Call TableTxn_TIDetails
Call TableMst_ExchangeType
Call TableMst_ServiceTax
Call TableMst_Client
Call TableMst_Depository

For i = 1 To MSHFlexGrid1.Rows - 2
   If CheckCorrectData(i) = True Then
      RsTxn_TIDetails.AddNew
      RsTxn_TIDetails!TIID = GetMaxTIID
      RsTxn_TIDetails!G_UID = GetGUID
      RsTxn_TIDetails!EntryFlag = "Manual"
      RsTxn_TIDetails!TIDate = CDate(MSHFlexGrid1.TextMatrix(i, 0))
      RsTxn_TIDetails!TINo = MSHFlexGrid1.TextMatrix(i, 1)
      RsTxn_TIDetails!PostInvoiceNo = MSHFlexGrid1.TextMatrix(i, 2)
      RsTxn_TIDetails!EntryNo = MSHFlexGrid1.TextMatrix(i, 3)
      RsTxn_TIDetails!TIAmount = Val(MSHFlexGrid1.TextMatrix(i, 4))
      RsTxn_TIDetails!TISTAmount = Val(MSHFlexGrid1.TextMatrix(i, 5))
      If LCase(MSHFlexGrid1.TextMatrix(i, 6)) = "client" Then
         RsTxn_TIDetails!ClientIDRow = mClientID
      Else
         RsTxn_TIDetails!DepositoryID = mClientID
      End If
      RsTxn_TIDetails!NAVCode = MSHFlexGrid1.TextMatrix(i, 7)
      RsTxn_TIDetails!LocationID = mLocationID
      RsTxn_TIDetails!ExchangeTypeID = mExchangeTypeID
      RsTxn_TIDetails!ServiceTaxID = mServiceTaxID
      RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
      RsTxn_TIDetails!CreatedDate = Now
      RsTxn_TIDetails.Update
   Else
      MSHFlexGrid1.Row = i
      For j = 0 To MSHFlexGrid1.Cols - 1
         MSHFlexGrid1.Col = j
         MSHFlexGrid1.CellBackColor = vbRed
      Next
   End If
Next

TxtErrMsg.Text = mErrorMsg
CmdImport.Enabled = False
CmdSave.Enabled = False
MsgBox "Done!!!"

End Sub

Private Function GetMaxTIID() As Double
Dim Retval As Double
tmp = "Select max(TIID) from Txn_TIDetails"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTIID = Retval
End Function

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


