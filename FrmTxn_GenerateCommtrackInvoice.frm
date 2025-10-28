VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_GenerateCommtrackInvoice 
   Caption         =   "Generate Commtrack Invoice"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.OptionButton OptCMSE 
         Caption         =   "CMS E Invoice"
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
         Left            =   3120
         TabIndex        =   24
         Top             =   2160
         Width           =   2175
      End
      Begin VB.OptionButton OptCMSP 
         Caption         =   "CMS P Invoice"
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
         TabIndex        =   23
         Top             =   2160
         Value           =   -1  'True
         Width           =   2175
      End
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
         ItemData        =   "FrmTxn_GenerateCommtrackInvoice.frx":0000
         Left            =   1560
         List            =   "FrmTxn_GenerateCommtrackInvoice.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   3255
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
         Left            =   13170
         TabIndex        =   9
         Top             =   365
         Width           =   1335
      End
      Begin VB.ComboBox CmbLocationID 
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
         Left            =   6120
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   3060
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
         Left            =   5160
         TabIndex        =   3
         Top             =   365
         Width           =   1335
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
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
         Left            =   6495
         TabIndex        =   4
         Top             =   365
         Width           =   1335
      End
      Begin VB.CommandButton CmdValidate 
         Caption         =   "Validate"
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
         Left            =   7830
         TabIndex        =   5
         Top             =   365
         Width           =   1335
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
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
         Left            =   9165
         TabIndex        =   6
         Top             =   365
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
         Height          =   350
         Left            =   10500
         TabIndex        =   7
         Top             =   365
         Width           =   1335
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
         Height          =   350
         Left            =   11835
         TabIndex        =   8
         Top             =   365
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker TxtMonth 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   48103427
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtError 
         Height          =   825
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1230
         Width           =   14415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   2595
         TabIndex        =   21
         Top             =   135
         Width           =   1170
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   885
         Width           =   705
      End
      Begin VB.Label LblLocationID 
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
         Left            =   3240
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   510
         TabIndex        =   12
         Top             =   135
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7350
      Left            =   120
      TabIndex        =   13
      Top             =   2565
      Width           =   14655
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6285
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   11086
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   6900
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label LblTotalRecord 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Left            =   12240
         TabIndex        =   16
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label7 
         Caption         =   "Total Record Found :"
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
         Left            =   9960
         TabIndex        =   15
         Top             =   240
         Width           =   2235
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   6765
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   11933
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmTxn_GenerateCommtrackInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim mFilenames As String
'Dim strFileNames  As Variant
'Dim mMonthYear As Variant
'Dim mYear, mMonthNo As Variant
'Dim mChkYear, mChkMonthNo As Variant

Dim mFileName As String
Dim mMonth, mYear As Variant
Dim mStartDate, mEndDate As Variant
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Double, mCol As Double
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mError As String
Dim mToDate As Variant


Private Sub ChkCheckAll_Click()
With MSHFlexGrid1
   If ChkCheckAll.Value = vbChecked Then
      For I = 1 To .Rows - 2
         .row = I
         .Col = 0
         .CellFontName = "Wingdings"
         .CellFontSize = 12
         .CellAlignment = flexAlignCenterCenter
         .Text = strChecked
      Next
   Else
      For I = 1 To .Rows - 2
         .row = I
         .Col = 0
         .CellFontName = "Wingdings"
         .CellFontSize = 12
         .CellAlignment = flexAlignCenterCenter
         .Text = strUnChecked
      Next
   End If
End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_GenerateCommtrackInvoice.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_GenerateCommtrackInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1 ' + 6
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
   Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetFile_Click()
On Error GoTo errortrap
TxtError.Text = ""
With CommonDialog1
   .DialogTitle = "Open Excel File..."
   .Filter = "Excel files(*.xls)|*.xlsx"
   .CancelError = True
procReOpen:
   .ShowOpen
End With
Path = CommonDialog1.FileName
If LblFile.Caption <> Path Then
   If OptCMSE.Value = True Then
      Call Grid_Head
   Else
      Call Grid_HeadCMSP
   End If
End If
LblFile.Caption = Path
If LblFile.Caption <> "" Then
   CmdImport.Enabled = True
End If
CmdValidate.Enabled = False
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub CmdGo_Click()

TxtError.Text = ""
mError = ""
Call Grid_Head2

mStartDate = Format(TxtMonth.Value, "yyyy-mm-dd")

'mTotalDays = Day(DateSerial(Year(TxtMonth.Value), Month(TxtMonth.Value), 0))

mToDate = DateSerial(Year(TxtMonth.Value), Month(TxtMonth.Value) + 1, 0)
tmp = "Select top 1 InvoiceID from Txn_MonthlyCommtrackInvoice Where MonthYear = '" & Format(mToDate, "yyyy-mm-dd") & "'"
If LCase(CmbInvoiceType.Text) = "non taxable" Then
   tmp = tmp & " AND Flag ='N'"
ElseIf LCase(CmbInvoiceType.Text) = "taxable" Then
   tmp = tmp & " AND Flag ='T'"
End If
Call Tmp7Table
If TmpRs7.RecordCount > 0 Then
   MsgBox "Invoice is already generated for month " & Format(TxtMonth.Value, "MMM-yyyy") & "!!!"
   Exit Sub
End If

If Format(TxtMonth.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future month not allowed!!!"
   TxtMonth.SetFocus
   Exit Sub
End If
If CmbInvoiceType.Text = "" Then
   MsgBox "Select Invoice Type!!!"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

MsgBox "Please wait till next Message!!!"

mStartDate = Format(TxtMonth.Value, "yyyy-mm-dd")

'mTotalDays = Day(DateSerial(Year(TxtMonth.Value), Month(TxtMonth.Value), 0))

mToDate = DateSerial(Year(TxtMonth.Value), Month(TxtMonth.Value) + 1, 0) 'Format(Year(TxtMonth.Value) & "-" & Month(TxtMonth.Value) & "-" & mTotalDays, "yyyy-mm-dd")


tmp = "Select ML.LocationName, ML.LocationID, TMCD.CPID, TMCD.CPName, MD.NAVCode, sum(TMCD.Amount) Amount, sum(TMCD.AmtIncST) AmountIncST"
tmp = tmp & " From Txn_MonthlyCommtrackFileDetail TMCD"
tmp = tmp & " Inner Join Mst_ISIN MISIN on MISIN.ISINID = TMCD.LOTID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MISIN.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " INNER JOIN Mst_depository MD ON MD.DepositoryID = TMCD.CPID"
tmp = tmp & " Where TMCD.MonthYear = '" & mStartDate & "'"
If LCase(CmbInvoiceType.Text) = "non taxable" Then
   tmp = tmp & " AND TMCD.AmtIncST = TMCD.Amount"
ElseIf LCase(CmbInvoiceType.Text) = "taxable" Then
   tmp = tmp & " AND TMCD.AmtIncST <> TMCD.Amount"
End If
tmp = tmp & " group by ML.LocationName, ML.LocationID, TMCD.CPID, TMCD.CPName, MD.NAVCode"
tmp = tmp & " order by ML.LocationName, TMCD.CPName"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   LblTotalRecord.Caption = TmpRs.RecordCount
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      For I = 1 To TmpRs.RecordCount
         '.TextMatrix(i, 0) = "SELECT"
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!CPID), "", TmpRs!CPID)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs!CPName), "", TmpRs!CPName)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs!Amount), 0, TmpRs!Amount)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs!AmountIncST), 0, TmpRs!AmountIncST)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
         If .TextMatrix(I, 7) = "" Then
            mError = mError & "NAVCode not found for CPID " & MSHFlexGrid1.TextMatrix(I, 2) & "!!!" & vbCrLf
         End If
         .row = I
         .Col = 0
         .CellFontName = "Wingdings"
         .CellFontSize = 12
         .CellAlignment = flexAlignCenterCenter
         .Text = strUnChecked
         TmpRs.MoveNext
      Next
   End With
   TxtError.Text = mError
   If TxtError.Text = "" Then
      CmdSave.Enabled = True
   End If
Else
   MsgBox "Record not found!!!"
End If
End Sub

Private Sub CmdImport_Click()
'If CmbMonth.Text = "" Then
'   MsgBox "Blank Month not Allowed"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox "Blank Year not Allowed!!!"
'   CmbYear.SetFocus
'   Exit Sub
'End If
'MsgBox "Please wait till next Message!!!"
'
'
'mStartDate = mYear & "-" & mMonth & "-01"
'
'tmp = "SELECT TOP 1 MonthYear FROM Txn_MonthlyCommtrack WHERE MonthYear = '" & mStartDate & "'"
'Call TmpTable
'If Not TmpRs.EOF Then
'   MsgBox "Data for '" & CmbMonth.Text & "-" & CmbYear.Text & "' already exists!!!"
'   Exit Sub
'End If
TxtError.Text = ""

Set oXL = New Excel.Application
Set oWB = oXL.Workbooks.Open(LblFile.Caption) '(Pat & "\excel\" & fn & ".xls", , , , "")
Set oWS = oWB.Worksheets(1)

If OptCMSE.Value = True Then
   If CheckFile("CMSE") = False Then
      MsgBox "Please Select proper file for uploading!!!"
      CmdImport.Enabled = False
      Exit Sub
   End If
   Call Grid_Head
Else
   If CheckFile("CMSP") = False Then
      MsgBox "Please Select proper file for uploading!!!"
      CmdImport.Enabled = False
      Exit Sub
   End If
   Call Grid_HeadCMSP
End If

Dim I, mRows As Variant
Set oWS = oWB.Worksheets(1)
ProgressBar1.Value = 0
mRows = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
LblTotalRecord.Caption = mRows - 1
ProgressBar1.Max = mRows + 1

MSHFlexGrid1.Rows = mRows + 1
For I = 2 To mRows
    If Trim(oWS.Cells(I, 1)) = "" Then Exit For
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(I - 1, C) = Trim(oWS.Cells(I, C + 1))
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = 0
CmdValidate.Enabled = True
CmdImport.Enabled = False
oXL.Quit
LblFile.Caption = ""
MsgBox "Import Successfully!!!"
End Sub

Private Sub CmdReset_Click()
TxtMonth.Value = Date
TxtError.Text = ""
Call TxtMonth_LostFocus
CmbInvoiceType.Text = ""
LblFile.Caption = ""
CmdImport.Enabled = False
CmdValidate.Enabled = False
If OptCMSE.Value = True Then
   Call Grid_Head
Else
   Call Grid_HeadCMSP
End If
End Sub

Private Sub CmdSave_Click()
TxtError.Text = ""

'tmp = "Select top 1 InvoiceID from Txn_MonthlyCommtrackInvoice Where MonthYear = '" & Format(TxtMonth.Value, "yyyy-mm-dd") & "'"

If OptCMSP.Value = True Then
   MsgBox "Only CMS E Invoice can be generated from This Screen!!!"
   Exit Sub
End If

tmp = "Select top 1 InvoiceID from Txn_MonthlyCommtrackInvoice Where MonthYear = '" & Format(mToDate, "yyyy-mm-dd") & "'"

If LCase(CmbInvoiceType.Text) = "non taxable" Then
   tmp = tmp & " AND Flag ='N'"
ElseIf LCase(CmbInvoiceType.Text) = "taxable" Then
   tmp = tmp & " AND Flag ='T'"
End If
Call Tmp7Table
If TmpRs7.RecordCount > 0 Then
   MsgBox "Invoice is already generated for month " & Format(TxtMonth.Value, "MMM-yyyy") & "!!!"
       Exit Sub
End If

Call TableTxn_MonthlyCommtrackInvoice(" Where MonthYear = '" & Format(mToDate, "yyyy-mm-dd") & "' ")
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows
For I = 1 To MSHFlexGrid1.Rows - 2

'CPID Varchar(50) ,CPName Varchar(200)

   If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
      RsTxn_MonthlyCommtrackInvoice.AddNew
      RsTxn_MonthlyCommtrackInvoice!InvoiceID = GetMaxInvoiceID
      RsTxn_MonthlyCommtrackInvoice!G_UID = GetGUID
      RsTxn_MonthlyCommtrackInvoice!MonthYear = CDate(mToDate) 'TxtMonth.Value 'CDate(mToDate) 'TxtMonth.Value
      RsTxn_MonthlyCommtrackInvoice!LocationID = Val(MSHFlexGrid1.TextMatrix(I, 6))
      RsTxn_MonthlyCommtrackInvoice!CPNAVCode = MSHFlexGrid1.TextMatrix(I, 7)
      RsTxn_MonthlyCommtrackInvoice!Amt = Val(MSHFlexGrid1.TextMatrix(I, 4))
      RsTxn_MonthlyCommtrackInvoice!STAmt = Val(MSHFlexGrid1.TextMatrix(I, 5)) - Val(MSHFlexGrid1.TextMatrix(I, 4))
      RsTxn_MonthlyCommtrackInvoice!TotalAmt = Val(MSHFlexGrid1.TextMatrix(I, 5))
      RsTxn_MonthlyCommtrackInvoice!PINo = RsTxn_MonthlyCommtrackInvoice!InvoiceID
      'RsTxn_MonthlyCommtrackInvoice!BatchID = 0
      RsTxn_MonthlyCommtrackInvoice!CPID = MSHFlexGrid1.TextMatrix(I, 2)
      RsTxn_MonthlyCommtrackInvoice!CPName = MSHFlexGrid1.TextMatrix(I, 3)
      RsTxn_MonthlyCommtrackInvoice!Flag = Left$(UCase(CmbInvoiceType.Text), 1)
      RsTxn_MonthlyCommtrackInvoice!CreatedBy = Gen_UserLoginID
      RsTxn_MonthlyCommtrackInvoice!CreatedDate = Now
      RsTxn_MonthlyCommtrackInvoice.Update
   End If
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Data Saved Successfully!!!"
Call CmdReset_Click
End Sub

Private Sub CmdValidate_Click()
'On Error GoTo errortrap
Call Grid_Head1
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Gen_mTimeStamp = GetTimeStamp
If OptCMSE.Value = True Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "CMSETxn_GenerateCommtrackInvoiceError.xls")
Else
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "CMSPTxn_GenerateCommtrackInvoiceError.xls")
End If
Set oXL = New Excel.Application
If OptCMSE.Value = True Then
   Set oWB = oXL.Workbooks.Open(App.Path & "\excel\" & Gen_mTimeStamp & Gen_UserName & "CMSETxn_GenerateCommtrackInvoiceError.xls")
Else
   Set oWB = oXL.Workbooks.Open(App.Path & "\excel\" & Gen_mTimeStamp & Gen_UserName & "CMSPTxn_GenerateCommtrackInvoiceError.xls")
End If

Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Error Detail"
oWS.ClearArrows
mRow = 1
TxtError.Text = ""
mError = ""
If Format(TxtMonth.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future month not allowed!!!"
   TxtMonth.SetFocus
   Exit Sub
End If
MsgBox "Please wait till next Message!!!"

mStartDate = Format(TxtMonth.Value, "yyyy-mm-dd")

If OptCMSE.Value = True Then
  
   tmp = "SELECT TOP 1 MonthYear FROM Txn_MonthlyCommtrackFileDetail WHERE MonthYear = '" & mStartDate & "'"
   Call TmpTable
   

   
   If TmpRs.RecordCount > 0 Then
      MsgBox "Data for month '" & Format(TxtMonth.Value, "MMM-yyyy") & "' already exists!!!"
      Exit Sub
   End If
Else
   tmp = "SELECT TOP 1 MonthYear FROM Txn_MonthlyCommtrackFileDetailCMSP WHERE MonthYear = '" & mStartDate & "'"
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "Data for month '" & Format(TxtMonth.Value, "MMM-yyyy") & "' already exists!!!"
      Exit Sub
   End If
End If

If OptCMSE.Value = True Then
   tmp = "Select MI.ISINID, MI.LocationID,ML.LocationName,MG.CommodityGroup,MC.CommodityGroupID from Mst_ISIN MI"
   tmp = tmp & " Left Join Mst_Commodity MC On MC.CommodityID=MI.CommodityID"
   tmp = tmp & " Left Join Mst_CommodityGroup MG On MG.CommodityGroupID=MC.CommodityGroupID"
   tmp = tmp & " Left Join Mst_Location ML On ML.LocationID=MI.LocationID"
Else
   tmp = "Select Distinct KRNo 'ISINID' , ML.LocationID,ML.LocationName,MCG.CommodityGroup,MCG.CommodityGroupID from Mst_GSL MG"
   tmp = tmp & " Left Join Mst_Commodity MC On MC.CommodityID=MG.CommodityID And MC.CommTrackFlag = '1'"
   tmp = tmp & " Left Join Mst_CommodityGroup MCG On MCG.CommodityGroupID=MC.CommodityGroupID"
   tmp = tmp & " Left Join Mst_Location ML On ML.SM_Prefix=MG.SM_Prefix"
   tmp = tmp & " Where MG.ExchangeTypeID = 2 And MG.Flag Not in ('Z') And MG.KRNo is not null And MG.KRNo > 0"
End If

Call Tmp1Table

If OptCMSE.Value = True Then
   tmp = "Select DepositoryID from MST_Depository"
   Call Tmp2Table
Else
   tmp = "Select CMSPCode 'Key',ClientIDRow from MST_Client Where ExchangeTypeID = 2 And CMSPCode is not null And CMSPCode <> ''"
   Call Tmp2Table
End If

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows
If TmpRs1.RecordCount > 0 And TmpRs2.RecordCount > 0 Then
   For I = 1 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "ISINID = '" & Trim(MSHFlexGrid1.TextMatrix(I, 0)) & "'"
      If TmpRs1.EOF Then
         'MsgBox "ISINID " & MSHFlexGrid1.TextMatrix(i, 0) & " not found in ISIN Master!!!"
         'ProgressBar1.Value = 0
         'Exit Sub
         If OptCMSE.Value = True Then
            mError = mError & "LotID " & MSHFlexGrid1.TextMatrix(I, 0) & " not found in ISIN Master!!!" & vbCrLf
            oWS.Cells(mRow, 1) = "LotID " & MSHFlexGrid1.TextMatrix(I, 0) & " not found in ISIN Master!!!"
         Else
            mError = mError & "LotID " & MSHFlexGrid1.TextMatrix(I, 0) & " not found in GSL Master!!!" & vbCrLf
            oWS.Cells(mRow, 1) = "LotID " & MSHFlexGrid1.TextMatrix(I, 0) & " not found in GSL Master!!!"
         End If
         mRow = mRow + 1
      Else
         'If Val(MSHFlexGrid1.TextMatrix(i, 5)) * Val(MSHFlexGrid1.TextMatrix(i, 8)) > 0 Then
         '   mRate = Format(Val(MSHFlexGrid1.TextMatrix(i, 10)) / (Val(MSHFlexGrid1.TextMatrix(i, 5)) * Val(MSHFlexGrid1.TextMatrix(i, 8))), "#####0.00")
         'Else
         '   mRate = 0
         'End If
         If OptCMSE.Value = True Then
            If Val(MSHFlexGrid1.TextMatrix(I, 8)) * Val(MSHFlexGrid1.TextMatrix(I, 15)) > 0 Then
               mRate = Format(Val(MSHFlexGrid1.TextMatrix(I, 20)) / (Val(MSHFlexGrid1.TextMatrix(I, 8)) * Val(MSHFlexGrid1.TextMatrix(I, 15))), "#####0.00") 'mRate = Format(Val(MSHFlexGrid1.TextMatrix(I, 19)) / (Val(MSHFlexGrid1.TextMatrix(I, 8)) * Val(MSHFlexGrid1.TextMatrix(I, 15))), "#####0.00")
            Else
               mRate = 0
            End If
         Else
            If Val(MSHFlexGrid1.TextMatrix(I, 6)) * Val(MSHFlexGrid1.TextMatrix(I, 13)) > 0 Then
               mRate = Format(Val(MSHFlexGrid1.TextMatrix(I, 17)) / (Val(MSHFlexGrid1.TextMatrix(I, 6)) * Val(MSHFlexGrid1.TextMatrix(I, 13))), "#####0.00")
            Else
               mRate = 0
            End If
         End If
         
         tmp = " Select distinct LocationID,CommodityGroupID,Rate From Mst_NCDEXStorageRateCard MNR"
         tmp = tmp & " Inner join Mst_NCDEXStorageRateCardDetail MNRD on MNR.NSRCID=MNRD.NSRCID"
         tmp = tmp & " Where LocationID=" & TmpRs1!LocationID & " And CommodityGroupID=" & TmpRs1!CommodityGroupID & " And Rate >0"
         Call Tmp5Table

         If TmpRs5.RecordCount = 0 Then
            mError = mError & "NCDEX Rate not found for Location:" & TmpRs1!Locationname & "  Commodity Group:" & TmpRs1!COmmodityGroup & vbCrLf
            oWS.Cells(mRow, 1) = "NCDEX Rate not found for Location:" & TmpRs1!Locationname & "  Commodity Group:" & TmpRs1!COmmodityGroup
            mRow = mRow + 1
            
         Else
            If TmpRs5.RecordCount = 1 Then
               If Val(mRate) <> TmpRs5!Rate Then
                  r = MSHFlexGrid2.Rows - 1
                  MSHFlexGrid2.TextMatrix(r, 0) = Format(TxtMonth.Value, "MMM-yyyy")
                  MSHFlexGrid2.TextMatrix(r, 1) = IIf(IsNull(TmpRs1!ISINID), "", TmpRs1!ISINID)
                  MSHFlexGrid2.TextMatrix(r, 2) = IIf(IsNull(TmpRs1!Locationname), "", TmpRs1!Locationname)
                  MSHFlexGrid2.TextMatrix(r, 3) = IIf(IsNull(TmpRs1!COmmodityGroup), "", TmpRs1!COmmodityGroup)
                  MSHFlexGrid2.TextMatrix(r, 4) = mRate
                  MSHFlexGrid2.TextMatrix(r, 5) = IIf(IsNull(TmpRs5!Rate), "", TmpRs5!Rate)
                  MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
               End If
            Else
               For Ln = 1 To TmpRs5.RecordCount
                   r = MSHFlexGrid2.Rows - 1
                   MSHFlexGrid2.TextMatrix(r, 0) = Format(TxtMonth.Value, "MMM-yyyy")
                   MSHFlexGrid2.TextMatrix(r, 1) = IIf(IsNull(TmpRs1!ISINID), "", TmpRs1!ISINID)
                   MSHFlexGrid2.TextMatrix(r, 2) = IIf(IsNull(TmpRs1!Locationname), "", TmpRs1!Locationname)
                   MSHFlexGrid2.TextMatrix(r, 3) = IIf(IsNull(TmpRs1!COmmodityGroup), "", TmpRs1!COmmodityGroup)
                   MSHFlexGrid2.TextMatrix(r, 4) = mRate
                   MSHFlexGrid2.TextMatrix(r, 5) = IIf(IsNull(TmpRs5!Rate), "", TmpRs5!Rate)
                   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
                   TmpRs5.MoveNext
               Next
            End If
         End If
      End If
      If OptCMSE.Value = True Then
         TmpRs2.MoveFirst
         TmpRs2.Find "DepositoryID = '" & Trim(MSHFlexGrid1.TextMatrix(I, 2)) & "'"
         If TmpRs2.EOF Then
         'MsgBox "Depositor " & MSHFlexGrid1.TextMatrix(i, 2) & " not found in Depository Master!!!"
         'ProgressBar1.Value = 0
         'Exit Sub
            mError = mError & "CPID " & MSHFlexGrid1.TextMatrix(I, 2) & " not found in Depository Master!!!" & vbCrLf
            oWS.Cells(mRow, 1) = "CPID " & MSHFlexGrid1.TextMatrix(I, 2) & " not found in Depository Master!!!"
            mRow = mRow + 1
         End If
      Else
         TmpRs2.MoveFirst
         TmpRs2.Find "Key = '" & Trim(MSHFlexGrid1.TextMatrix(I, 2)) & "' "   '& "~" & Trim(MSHFlexGrid1.TextMatrix(I, 3)) & "'"
         If TmpRs2.EOF Then
         'MsgBox "Depositor " & MSHFlexGrid1.TextMatrix(i, 2) & " not found in Depository Master!!!"
         'ProgressBar1.Value = 0
         'Exit Sub
            mError = mError & "CMS P Code  " & MSHFlexGrid1.TextMatrix(I, 2) & " not found for Client " & MSHFlexGrid1.TextMatrix(I, 3) & "  in Client Master!!!" & vbCrLf
            oWS.Cells(mRow, 1) = "CMS P Code   " & MSHFlexGrid1.TextMatrix(I, 2) & " not found for Client " & MSHFlexGrid1.TextMatrix(I, 3) & " in Client Master!!!"
            mRow = mRow + 1
         Else
            mClientIDRow = TmpRs2!ClientIDRow
         End If
      End If
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
Else
   MsgBox "Depositor OR ISINID not found, data not validated!!!"
   ProgressBar1.Value = 0
   Exit Sub
End If
ProgressBar1.Value = ProgressBar1.Max
TxtError.Text = mError
Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Rate Card Details"
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
       oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
    Next
Next

oWB.Save

If mError <> "" Then
   oXL.Visible = True
   MsgBox "Data not validated due to some error!!!"
   Exit Sub
End If

ProgressBar1.Value = 0

If OptCMSE.Value = True Then
   Call TableTxn_MonthlyCommtrackFileDetail(" WHERE MonthYear = '" & mStartDate & "' ")
   For I = 1 To MSHFlexGrid1.Rows - 2
       RsTxn_MonthlyCommtrackFileDetail.AddNew
       RsTxn_MonthlyCommtrackFileDetail!G_UID = GetGUID
       RsTxn_MonthlyCommtrackFileDetail!MonthYear = CDate(mStartDate)
       RsTxn_MonthlyCommtrackFileDetail!LotId = MSHFlexGrid1.TextMatrix(I, 0)
       RsTxn_MonthlyCommtrackFileDetail!EDD = CDate(MSHFlexGrid1.TextMatrix(I, 1))
       RsTxn_MonthlyCommtrackFileDetail!CPID = MSHFlexGrid1.TextMatrix(I, 2)
       RsTxn_MonthlyCommtrackFileDetail!CPName = MSHFlexGrid1.TextMatrix(I, 3)
       RsTxn_MonthlyCommtrackFileDetail!ClientID = MSHFlexGrid1.TextMatrix(I, 4)
       RsTxn_MonthlyCommtrackFileDetail!ClientName = MSHFlexGrid1.TextMatrix(I, 5)
       RsTxn_MonthlyCommtrackFileDetail!Commodity = MSHFlexGrid1.TextMatrix(I, 6)
       RsTxn_MonthlyCommtrackFileDetail!CommodityCode = MSHFlexGrid1.TextMatrix(I, 7)
       RsTxn_MonthlyCommtrackFileDetail!Qty = Val(MSHFlexGrid1.TextMatrix(I, 8))
       RsTxn_MonthlyCommtrackFileDetail!ChargeableQty = Val(MSHFlexGrid1.TextMatrix(I, 9))
       RsTxn_MonthlyCommtrackFileDetail!TranxType = Trim(MSHFlexGrid1.TextMatrix(I, 10))
       RsTxn_MonthlyCommtrackFileDetail!TranxNo = Trim(MSHFlexGrid1.TextMatrix(I, 11))
       RsTxn_MonthlyCommtrackFileDetail!TranxDate = CDate(MSHFlexGrid1.TextMatrix(I, 12))
       RsTxn_MonthlyCommtrackFileDetail!StartDate = CDate(MSHFlexGrid1.TextMatrix(I, 13))
       RsTxn_MonthlyCommtrackFileDetail!EndDate = CDate(MSHFlexGrid1.TextMatrix(I, 14))
       RsTxn_MonthlyCommtrackFileDetail!HoldDuration = Val(MSHFlexGrid1.TextMatrix(I, 15))
       RsTxn_MonthlyCommtrackFileDetail!ChargeableDuration = Val(MSHFlexGrid1.TextMatrix(I, 16))
       RsTxn_MonthlyCommtrackFileDetail!Rate = Val(MSHFlexGrid1.TextMatrix(I, 17))
       RsTxn_MonthlyCommtrackFileDetail!MultiplyingFactor = Val(MSHFlexGrid1.TextMatrix(I, 18))
       RsTxn_MonthlyCommtrackFileDetail!GST = Val(MSHFlexGrid1.TextMatrix(I, 19))
       RsTxn_MonthlyCommtrackFileDetail!Amount = Val(MSHFlexGrid1.TextMatrix(I, 20))
       RsTxn_MonthlyCommtrackFileDetail!AmtIncST = Val(MSHFlexGrid1.TextMatrix(I, 21))
       RsTxn_MonthlyCommtrackFileDetail!WarehouseCode = MSHFlexGrid1.TextMatrix(I, 22)
       RsTxn_MonthlyCommtrackFileDetail!WarehouseName = MSHFlexGrid1.TextMatrix(I, 23)
       RsTxn_MonthlyCommtrackFileDetail!WSPCode = MSHFlexGrid1.TextMatrix(I, 24)
       RsTxn_MonthlyCommtrackFileDetail!WSPName = MSHFlexGrid1.TextMatrix(I, 25)
      
       RsTxn_MonthlyCommtrackFileDetail!CreatedBy = Gen_UserLoginID
       RsTxn_MonthlyCommtrackFileDetail!CreatedDate = Now
       RsTxn_MonthlyCommtrackFileDetail.Update
       ProgressBar1.Value = ProgressBar1.Value + 1
    Next
Else
   Call TableTxn_MonthlyCommtrackFileDetailCMSP(" WHERE MonthYear = '" & mStartDate & "' ")
   For I = 1 To MSHFlexGrid1.Rows - 2
       RsTxn_MonthlyCommtrackFileDetailCMSP.AddNew
       RsTxn_MonthlyCommtrackFileDetailCMSP!G_UID = GetGUID
       RsTxn_MonthlyCommtrackFileDetailCMSP!MonthYear = CDate(mStartDate)
       RsTxn_MonthlyCommtrackFileDetailCMSP!LotId = MSHFlexGrid1.TextMatrix(I, 0)
       RsTxn_MonthlyCommtrackFileDetailCMSP!EDD = CDate(MSHFlexGrid1.TextMatrix(I, 1))
       RsTxn_MonthlyCommtrackFileDetailCMSP!ClientID = MSHFlexGrid1.TextMatrix(I, 2)
       RsTxn_MonthlyCommtrackFileDetailCMSP!ClientName = MSHFlexGrid1.TextMatrix(I, 3)
       RsTxn_MonthlyCommtrackFileDetailCMSP!Commodity = MSHFlexGrid1.TextMatrix(I, 4)
       RsTxn_MonthlyCommtrackFileDetailCMSP!CommodityCode = MSHFlexGrid1.TextMatrix(I, 5)
       RsTxn_MonthlyCommtrackFileDetailCMSP!Qty = Val(MSHFlexGrid1.TextMatrix(I, 6))
       RsTxn_MonthlyCommtrackFileDetailCMSP!ChargeableQty = Val(MSHFlexGrid1.TextMatrix(I, 7))
       RsTxn_MonthlyCommtrackFileDetailCMSP!TranxType = Trim(MSHFlexGrid1.TextMatrix(I, 8))
       RsTxn_MonthlyCommtrackFileDetailCMSP!TranxNo = Trim(MSHFlexGrid1.TextMatrix(I, 9))
       RsTxn_MonthlyCommtrackFileDetailCMSP!TranxDate = CDate(MSHFlexGrid1.TextMatrix(I, 10))
       RsTxn_MonthlyCommtrackFileDetailCMSP!StartDate = CDate(MSHFlexGrid1.TextMatrix(I, 11))
       RsTxn_MonthlyCommtrackFileDetailCMSP!EndDate = CDate(MSHFlexGrid1.TextMatrix(I, 12))
       RsTxn_MonthlyCommtrackFileDetailCMSP!HoldDuration = Val(MSHFlexGrid1.TextMatrix(I, 13))
       RsTxn_MonthlyCommtrackFileDetailCMSP!ChargeableDuration = Val(MSHFlexGrid1.TextMatrix(I, 14))
       RsTxn_MonthlyCommtrackFileDetailCMSP!Rate = Val(MSHFlexGrid1.TextMatrix(I, 15))
       RsTxn_MonthlyCommtrackFileDetailCMSP!MultiplyingFactor = Val(MSHFlexGrid1.TextMatrix(I, 16))
       RsTxn_MonthlyCommtrackFileDetailCMSP!Amount = Val(MSHFlexGrid1.TextMatrix(I, 17))
       RsTxn_MonthlyCommtrackFileDetailCMSP!AmtIncST = Val(MSHFlexGrid1.TextMatrix(I, 18))
       RsTxn_MonthlyCommtrackFileDetailCMSP!WarehouseCode = MSHFlexGrid1.TextMatrix(I, 19)
       RsTxn_MonthlyCommtrackFileDetailCMSP!WarehouseName = MSHFlexGrid1.TextMatrix(I, 20)
       RsTxn_MonthlyCommtrackFileDetailCMSP!WSPCode = MSHFlexGrid1.TextMatrix(I, 21)
       RsTxn_MonthlyCommtrackFileDetailCMSP!WSPName = MSHFlexGrid1.TextMatrix(I, 22)
       RsTxn_MonthlyCommtrackFileDetailCMSP!ClientIDRow = mClientIDRow
       RsTxn_MonthlyCommtrackFileDetailCMSP!CreatedBy = Gen_UserLoginID
       RsTxn_MonthlyCommtrackFileDetailCMSP!CreatedDate = Now
       RsTxn_MonthlyCommtrackFileDetailCMSP.Update
       ProgressBar1.Value = ProgressBar1.Value + 1
    Next
End If

ProgressBar1.Value = ProgressBar1.Max
MsgBox "Validated Successfully!!!"
Call Grid_Head

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"

End Sub

Private Sub CmbInvoiceType_LostFocus()
If CmbInvoiceType.Text = "" Then Exit Sub
If LCase(CmbInvoiceType.Text) <> "non taxable" And LCase(CmbInvoiceType.Text) <> "taxable" Then
   MsgBox "Invalid Selection!!!"
   
     
   CmbInvoiceType.SetFocus
   Exit Sub
End If
End Sub

Function CheckFile(Flg_ As Variant) As Boolean
Dim Retval As Boolean
Retval = True

If LCase(Flg_) = "cmse" Then
   If UCase(oWS.Cells(1, 1)) <> "LOT ID" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 2)) <> "EDD" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 3)) <> "CP ID" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 4)) <> "CP NAME" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 5)) <> "CLIENT ID" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 6)) <> "CLIENT NAME" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 7)) <> "COMMODITY NAME" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 8)) <> "COMMODITY CODE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 9)) <> "ACTUAL QTY" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 10)) <> "CHARGEABLE QTY" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 11)) <> "TRANX TYPE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 12)) <> "TRANX NO" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 13)) <> "TRANX DATE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 14)) <> "START DATE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 15)) <> "END DATE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 16)) <> "HOLDING DURATION" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 17)) <> "CHARGEABLE DURATION" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 18)) <> "RATE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 19)) <> "MULTIPLYING FACTOR" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 20)) <> "GST" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 21)) <> "AMOUNT" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 22)) <> "AMT. INC ST" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 23)) <> "WAREHOUSE CODE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 24)) <> "WAREHOUSE NAME" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 25)) <> "WSP CODE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 26)) <> "WSP NAME" Then
      Retval = False
   End If
Else
   If UCase(oWS.Cells(1, 1)) <> "LOT ID" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 2)) <> "EDD" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 3)) <> "CLIENT ID" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 4)) <> "CLIENT NAME" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 5)) <> "COMMODITY NAME" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 6)) <> "COMMODITY CODE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 7)) <> "ACTUAL QTY" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 8)) <> "CHARGEABLE QTY" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 9)) <> "TRANX TYPE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 10)) <> "TRANX NO" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 11)) <> "TRANX DATE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 12)) <> "START DATE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 13)) <> "END DATE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 14)) <> "HOLDING DURATION" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 15)) <> "CHARGEABLE DURATION" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 16)) <> "RATE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 17)) <> "MULTIPLYING FACTOR" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 18)) <> "AMOUNT" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 19)) <> "AMT. INC ST" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 20)) <> "WAREHOUSE CODE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 21)) <> "WAREHOUSE NAME" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 22)) <> "WSP CODE" Then
      Retval = False
   ElseIf UCase(oWS.Cells(1, 23)) <> "WSP NAME" Then
      Retval = False
   End If
End If

CheckFile = Retval

End Function

Public Sub Grid_Head()
LblTotalRecord.Caption = ""
ChkCheckAll.Enabled = False
ChkCheckAll.Value = vbUnchecked
CmdSave.Enabled = False
CmdImport.Enabled = False
CmdValidate.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .FixedRows = 1
   .Cols = 26 '25 '23 '22 '15
   .AllowUserResizing = flexResizeColumns
'   .TextMatrix(0, 0) = "LOT ID"
'   .TextMatrix(0, 1) = "CP ID"
'   .TextMatrix(0, 2) = "CP NAME"
'   .TextMatrix(0, 3) = "CLIENT ID"
'   .TextMatrix(0, 4) = "CLIENT NAME"
'   .TextMatrix(0, 5) = "QTY"
'   .TextMatrix(0, 6) = "START DATE"
'   .TextMatrix(0, 7) = "END DATE"
'   .TextMatrix(0, 8) = "HOLD DURATION"
'   .TextMatrix(0, 9) = "AMOUNT"
'   .TextMatrix(0, 10) = "AMT.INC. ST"
'   .TextMatrix(0, 11) = "WAREHOUSE NAME"
'   .TextMatrix(0, 12) = "COMMODITY"
'   .TextMatrix(0, 13) = "WAREHOUSE CODE"
'   .TextMatrix(0, 14) = "WSP NAME"
   
   .TextMatrix(0, 0) = "LOT ID"
   .TextMatrix(0, 1) = "EDD"
   .TextMatrix(0, 2) = "CP Id"
   .TextMatrix(0, 3) = "CP Name"
   .TextMatrix(0, 4) = "Client Id"
   .TextMatrix(0, 5) = "Client Name"
   .TextMatrix(0, 6) = "Commodity Name"
   .TextMatrix(0, 7) = "Commodity Code"
   .TextMatrix(0, 8) = "Actual Qty"
   .TextMatrix(0, 9) = "Chargeable Qty"
   .TextMatrix(0, 10) = "Tranx Type"
   .TextMatrix(0, 11) = "Tranx No"
   .TextMatrix(0, 12) = "Tranx Date"
   .TextMatrix(0, 13) = "Start Date"
   .TextMatrix(0, 14) = "End Date"
   .TextMatrix(0, 15) = "Holding Duration"
   .TextMatrix(0, 16) = "Chargeable Duration"
   .TextMatrix(0, 17) = "Rate"
   .TextMatrix(0, 18) = "Multiplying Factor"
   .TextMatrix(0, 19) = "GST"
   .TextMatrix(0, 20) = "Amount"
   .TextMatrix(0, 21) = "Amt. Inc ST"
   .TextMatrix(0, 22) = "Warehouse Code"
   .TextMatrix(0, 23) = "Warehouse Name"
   .TextMatrix(0, 24) = "WSP Code"
   .TextMatrix(0, 25) = "WSP Name"
   
   
   
   
   .ColWidth(0) = 1200
   .ColWidth(1) = 1200
   .ColWidth(2) = 1200
   .ColWidth(3) = 1200
   .ColWidth(4) = 1200
   .ColWidth(5) = 1200
   .ColWidth(6) = 1200
   .ColWidth(7) = 1200
   .ColWidth(8) = 1200
   .ColWidth(9) = 1200
   .ColWidth(10) = 1200
   .ColWidth(11) = 1200
   .ColWidth(12) = 1200
   .ColWidth(13) = 1200
   .ColWidth(14) = 1200
   .ColWidth(15) = 1200
   .ColWidth(16) = 1200
   .ColWidth(17) = 1200
   .ColWidth(18) = 1200
   .ColWidth(19) = 1200
   .ColWidth(20) = 1200
   .ColWidth(21) = 1200
   .ColWidth(22) = 1200
   .ColWidth(23) = 1200
   .ColWidth(24) = 1200
   .ColWidth(25) = 1200
      
   .RowHeight(0) = 500
   .WordWrap = True
End With
End Sub

Public Sub Grid_HeadCMSP()
LblTotalRecord.Caption = ""
ChkCheckAll.Enabled = False
ChkCheckAll.Value = vbUnchecked
CmdSave.Enabled = False
CmdImport.Enabled = False
CmdValidate.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .FixedRows = 1
   .Cols = 23 '22 '15
   .AllowUserResizing = flexResizeColumns
'   .TextMatrix(0, 0) = "LOT ID"
'   .TextMatrix(0, 1) = "CP ID"
'   .TextMatrix(0, 2) = "CP NAME"
'   .TextMatrix(0, 3) = "CLIENT ID"
'   .TextMatrix(0, 4) = "CLIENT NAME"
'   .TextMatrix(0, 5) = "QTY"
'   .TextMatrix(0, 6) = "START DATE"
'   .TextMatrix(0, 7) = "END DATE"
'   .TextMatrix(0, 8) = "HOLD DURATION"
'   .TextMatrix(0, 9) = "AMOUNT"
'   .TextMatrix(0, 10) = "AMT.INC. ST"
'   .TextMatrix(0, 11) = "WAREHOUSE NAME"
'   .TextMatrix(0, 12) = "COMMODITY"
'   .TextMatrix(0, 13) = "WAREHOUSE CODE"
'   .TextMatrix(0, 14) = "WSP NAME"
   
   .TextMatrix(0, 0) = "LOT ID"
   .TextMatrix(0, 1) = "EDD"
   .TextMatrix(0, 2) = "Client Id"
   .TextMatrix(0, 3) = "Client Name"
   .TextMatrix(0, 4) = "Commodity Name"
   .TextMatrix(0, 5) = "Commodity Code"
   .TextMatrix(0, 6) = "Actual Qty"
   .TextMatrix(0, 7) = "Chargeable Qty"
   .TextMatrix(0, 8) = "Tranx Type"
   .TextMatrix(0, 9) = "Tranx No"
   .TextMatrix(0, 10) = "Tranx Date"
   .TextMatrix(0, 11) = "Start Date"
   .TextMatrix(0, 12) = "End Date"
   .TextMatrix(0, 13) = "Holding Duration"
   .TextMatrix(0, 14) = "Chargeable Duration"
   .TextMatrix(0, 15) = "Rate"
   .TextMatrix(0, 16) = "Multiplying Factor"
   .TextMatrix(0, 17) = "Amount"
   .TextMatrix(0, 18) = "Amt. Inc ST"
   .TextMatrix(0, 19) = "Warehouse Code"
   .TextMatrix(0, 20) = "Warehouse Name"
   .TextMatrix(0, 21) = "WSP Code"
   .TextMatrix(0, 22) = "WSP Name"
   
   
   
   
   .ColWidth(0) = 1200
   .ColWidth(1) = 1200
   .ColWidth(2) = 1200
   .ColWidth(3) = 1200
   .ColWidth(4) = 1200
   .ColWidth(5) = 1200
   .ColWidth(6) = 1200
   .ColWidth(7) = 1200
   .ColWidth(8) = 1200
   .ColWidth(9) = 1200
   .ColWidth(10) = 1200
   .ColWidth(11) = 1200
   .ColWidth(12) = 1200
   .ColWidth(13) = 1200
   .ColWidth(14) = 1200
   .ColWidth(15) = 1200
   .ColWidth(16) = 1200
   .ColWidth(17) = 1200
   .ColWidth(18) = 1200
   .ColWidth(19) = 1200
   .ColWidth(20) = 1200
   .ColWidth(21) = 1200
   
   .RowHeight(0) = 500
   .WordWrap = True
End With
End Sub


Public Sub Grid_Head2()
LblTotalRecord.Caption = ""
ProgressBar1.Value = 0
ChkCheckAll.Enabled = True
ChkCheckAll.Value = vbUnchecked
CmdValidate.Enabled = False
CmdImport.Enabled = False
CmdSave.Enabled = False
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .FixedRows = 1
   .Cols = 8
   .AllowUserResizing = flexResizeColumns
   .TextMatrix(0, 0) = "SELECT"
   .TextMatrix(0, 1) = "LOCATION NAME"
   .TextMatrix(0, 2) = "CP ID"
   .TextMatrix(0, 3) = "CP NAME"
   .TextMatrix(0, 4) = "AMOUNT"
   .TextMatrix(0, 5) = "AMOUNT INC TAX"
   .TextMatrix(0, 6) = "LOCATIONID"
   .TextMatrix(0, 7) = "CPNAVCODE"
    
   .ColWidth(0) = 1200
   .ColWidth(1) = 2500
   .ColWidth(2) = 2500
   .ColWidth(3) = 4500
   .ColWidth(5) = 0
   .ColWidth(6) = 0
   .ColWidth(7) = 0
   
   .RowHeight(0) = 500
   .WordWrap = True
End With
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True

OptCMSP.Value = True
Call CmdReset_Click

End Sub

Private Sub OptCMSE_Click()
Call CmdReset_Click
End Sub

Private Sub OptCMSP_Click()
Call CmdReset_Click
End Sub

Private Sub TxtMonth_GotFocus()
tmp = TxtMonth.Value
End Sub

Private Sub TxtMonth_LostFocus()
TxtMonth.Day = 1
If tmp <> TxtMonth.Value Then
   Call Grid_Head
End If
End Sub

Private Function GetMaxInvoiceID() As Double
Dim Retval As Double
tmp = "Select MAX(InvoiceID) from Txn_MonthlyCommtrackInvoice "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxInvoiceID = Retval
End Function

Private Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 6
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Month-Year"
    .TextMatrix(0, 1) = "ISIN"
    .TextMatrix(0, 2) = "Location"
    .TextMatrix(0, 3) = "Commodity Group"
    .TextMatrix(0, 4) = "NCDEX Rate"
    .TextMatrix(0, 5) = "Store-Man Rate"
        
    .ColWidth(0) = 1600
    .ColWidth(1) = 1600
    .ColWidth(2) = 2200
    .ColWidth(3) = 2200
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
End With
End Sub

