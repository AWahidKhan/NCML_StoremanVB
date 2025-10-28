VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_RRNUpload 
   Caption         =   "RRN Upload"
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
      Height          =   1245
      Left            =   90
      TabIndex        =   2
      Top             =   0
      Width           =   15015
      Begin VB.CheckBox ChkCommTrack 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3195
         TabIndex        =   25
         Top             =   180
         Width           =   195
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
         Height          =   345
         Left            =   6165
         TabIndex        =   22
         Top             =   525
         Width           =   1575
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
         Height          =   345
         Left            =   75
         TabIndex        =   7
         Top             =   525
         Width           =   1455
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
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
         Height          =   345
         Left            =   4455
         TabIndex        =   6
         Top             =   525
         Width           =   1695
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
         Height          =   345
         Left            =   7740
         TabIndex        =   5
         Top             =   525
         Width           =   1455
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
         Height          =   345
         Left            =   1530
         TabIndex        =   4
         Top             =   525
         Width           =   1455
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
         Height          =   345
         Left            =   2985
         TabIndex        =   3
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Commtrack RRN Upload"
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
         Left            =   75
         TabIndex        =   26
         Top             =   165
         Width           =   2700
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   11685
         TabIndex        =   15
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   9360
         TabIndex        =   14
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   1005
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Error Details"
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Top             =   1350
      Width           =   15000
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1455
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox TxtErrorMsg 
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   240
         Width           =   9045
      End
      Begin VB.Label Label10 
         Caption         =   "Already in RRN Master but Mismatch in RRN Qty."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10320
         TabIndex        =   24
         Top             =   615
         Width           =   4245
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FFFF&
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
         Left            =   9555
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C000&
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
         Left            =   9555
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
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
         Left            =   9555
         TabIndex        =   20
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Already in RRN Master "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10320
         TabIndex        =   19
         Top             =   255
         Width           =   2325
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
         Left            =   9555
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Client Created in Client Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   17
         Top             =   960
         Width           =   2730
      End
      Begin VB.Label Label8 
         Caption         =   "Records not in Master Tables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   16
         Top             =   1335
         Width           =   2655
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5775
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   10186
      _Version        =   393216
      WordWrap        =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   9075
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   8160
      TabIndex        =   11
      Top             =   0
      Width           =   90
      Begin VB.ComboBox CmbSheetName 
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
         Left            =   2145
         TabIndex        =   12
         Top             =   360
         Width           =   4740
      End
      Begin VB.Label Label1 
         Caption         =   "Select Sheet Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   13
         Top             =   360
         Width           =   2025
      End
   End
End
Attribute VB_Name = "Frm_RRNUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet

Dim ErrMsg As Variant
Dim Status As Variant
Dim mSheetFlg, mExchangeTypeID, mDPID As Variant
Dim mClientIDRow As Double
Dim mISINID As Variant
Dim mDepositoryID As Variant
Dim mErrorMsg As Variant
Dim mRRNUpdated As Boolean
'Dim ErrMsg As Variant

Private Sub ChkCommTrack_Click()
Call Grid_HeadError
If ChkCommTrack.Value = 0 Then
   Call Grid_Head
   MSHFlexGrid2.Visible = False
   CmdExit.Enabled = True
   CmdSave.Enabled = False
   CmdGetFile.Enabled = True
   CmdImport.Enabled = False
   LblFile.Caption = ""
   LblTotalRecords.Caption = ""
Else
   Call Grid_HeadCommtrack
   MSHFlexGrid2.Visible = True
   CmdExit.Enabled = True
   CmdSave.Enabled = False
   CmdGetFile.Enabled = True
   CmdImport.Enabled = False
   LblFile.Caption = ""
   LblTotalRecords.Caption = ""
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "PendingRRNUpload.xls")
Call Xls_Detail_Rpt
CmdExcel.Enabled = False
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "PendingRRNUpload.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
        
        MSHFlexGrid1.row = I
        MSHFlexGrid1.Col = C
        
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
        ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbGreen
        ElseIf MSHFlexGrid1.CellBackColor = vbBlue Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbBlue
        ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbYellow
        End If
    Next
Next

If ChkCommTrack.Value = 1 Then
   Set oWS = oWB.Worksheets("Sheet2")
   oWS.ClearArrows
   mRow = 1
   For I = 0 To MSHFlexGrid2.Rows - 1
       mRow = I + 1 ' + 6
       For C = 0 To MSHFlexGrid2.Cols - 1
           If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(I, C), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
              End If
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
           End If
       Next
   Next
End If

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetFile_Click()

On Error GoTo errortrap

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx"
     .CancelError = True
procReOpen:
     .ShowOpen
End With

Path = CommonDialog1.FileName
LblFile.Caption = Path

'CmbSheetName.Clear

If LblFile.Caption <> "" Then
   CmdExit.Enabled = True
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   Dim I As Integer
   Dim runApp As Boolean
'   runApp = False
   For I = 1 To oWB.Worksheets.Count
       Set oWS = oWB.Worksheets(I)
       CmbSheetName.AddItem oWS.Name
'       If LCase(oWS.Name) = "sheet1" Then
'          runApp = True
'       End If
   Next
'   If runApp = False Then
'      MsgBox "Please Select Proper Excel Sheet!!!"
'      Exit Sub
'   End If
'
'   CmbSheetName.SelText = "Sheet1"
   Frame2.Enabled = True
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   oXL.Quit
   MsgBox "Select Import To Import Data From Excel Sheet!!!"
   CmdImport.SetFocus
End If

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub CmbSheetName_LostFocus()
If CmbSheetName.Text = "" Then
   mSheetFlg = Null
   Exit Sub
End If
Call Grid_Head

End Sub
Private Sub CmdExit_Click()
On Error Resume Next
oXL.Quit
Unload Me
End Sub

Private Sub CmdImport_Click()

'If CmbSheetName.Text = "" Then
'   MsgBox "Name Of Sheet Can't be blank!!!"
'   CmbSheetName.SetFocus
'   Exit Sub
'End If

'If CmbSheetName.Text <> "" Then
   MsgBox "Wait till next message!!!"
   TxtErrorMsg = ""
   ErrMsg = ""
   Status = 0
   RowCount = 0
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   Set oWS = oWB.Worksheets(1)
   If ChkCommTrack.Value = 0 Then
      Call Grid_Head
      If CheckFile = False Then
      'MsgBox "Please select proper file !!!"
         oXL.Quit
         Exit Sub
      End If
      Call FillData
   ElseIf ChkCommTrack.Value = 1 Then
      Call Grid_HeadCommtrack
      If CheckFileCommtrack = False Then
      'MsgBox "Please select proper file !!!"
         oXL.Quit
         Exit Sub
      End If
      Call FillData
   End If
   
   CmdSave.Enabled = True
   CmdReset.Enabled = True
   CmdSave.SetFocus
   MsgBox "Import Succesful!!!"
'End If

End Sub
Private Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True
'Set oWS = oWB.Worksheets("Sheet1")
If LCase(Trim(oWS.Cells(1, 1))) <> "rrn" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 1 "
ElseIf LCase(Trim(oWS.Cells(1, 2))) <> "rr date" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 2 "
ElseIf LCase(Trim(oWS.Cells(1, 3))) <> "isin" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 3 "
ElseIf LCase(Trim(oWS.Cells(1, 4))) <> "warehouse" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 4 "
ElseIf LCase(Trim(oWS.Cells(1, 5))) <> "commodity" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 5 "
ElseIf LCase(Trim(oWS.Cells(1, 6))) <> "dpid" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 6 "
ElseIf LCase(Trim(oWS.Cells(1, 7))) <> "clientid" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 7 "
ElseIf LCase(Trim(oWS.Cells(1, 8))) <> "clientname" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 8 "
ElseIf LCase(Trim(oWS.Cells(1, 9))) <> "remat" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 9 "
ElseIf LCase(Trim(oWS.Cells(1, 10))) <> "confirmed qty at dp" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 10 "
ElseIf LCase(Trim(oWS.Cells(1, 11))) <> "confirmed qty at ncom" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 11 "
ElseIf LCase(Trim(oWS.Cells(1, 12))) <> "actual pending qty" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 12 "
End If

CheckFile = Retval

End Function

Private Function CheckFileCommtrack() As Boolean
Dim Retval As Boolean
Retval = True
'Set oWS = oWB.Worksheets("Sheet1")
'RRN No  RRN Date    ISIN No Client ID   RRN Qty

If LCase(Trim(oWS.Cells(1, 1))) <> "rrn no" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 1 "
ElseIf LCase(Trim(oWS.Cells(1, 2))) <> "rrn date" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 2 "
ElseIf LCase(Trim(oWS.Cells(1, 3))) <> "isin no" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 3 "
ElseIf LCase(Trim(oWS.Cells(1, 4))) <> "client id" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 4 "
ElseIf LCase(Trim(oWS.Cells(1, 5))) <> "rrn qty" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 5 "
End If

CheckFileCommtrack = Retval

End Function


Private Sub CmdReset_Click()

ChkCommTrack.Value = 0
Call Grid_Head
CmdExit.Enabled = True
CmdSave.Enabled = False
CmdGetFile.Enabled = True
CmdImport.Enabled = False
LblFile.Caption = ""
LblTotalRecords.Caption = ""
mSheetFlg = ""
TxtErrorMsg = ""
CmbSheetName.Text = ""
ProgressBar1.Value = 0

End Sub

Private Sub CmdSave_Click()

Call Grid_HeadError
CmdSave.Enabled = False

If MSHFlexGrid1.Rows < 3 Then
   MsgBox "No data to save!!!"
   Exit Sub
End If

If MSHFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub

MsgBox "Wait till Next Message....."


Call TableMst_RRN
Call TableMst_ISIN
Call TableMst_Client(" Where ExchangeTypeID = 2 And Status = 'A' ")



If ChkCommTrack.Value = 0 Then
   Call TableMst_Depository
End If
If ChkCommTrack.Value = 1 Then
   Call TableMst_RRN_LotDetail
   tmp = "Select * from Mst_RRN "
   Call Tmp1Table
   tmp = "Select Distinct b.RRNNo,b.RRNID,a.ISINID from Mst_RRN_LotDetail a Inner Join Mst_RRN b on a.RRNID = b.RRNID "
   Call Tmp2Table

'
'   tmp = "Select NCDEXCode,TransNo, Convert(Varchar,CARNo) 'CARNo' from Txn_ImportNcdexLotPleadge Where CARNo Is not Null"
'   'tmp = tmp & " Union All"
'   'tmp = tmp & " Select NCDEXCode,TransNo, '' CARNo from Txn_ImportNcdexLotPleadge Where BatchID1 <> 0 And BatchID2 <> 0 And CARNo Is Null ANd Flag = 'P'"
'   tmp = tmp & " Union All"
'   tmp = tmp & " Select NCDEXCode,TransNo, '' CARNo from Txn_ImportNcdexLotPleadge Where CARNo Is Null ANd BatchID1 = 0"
'   tmp = tmp & " Union All"
'   tmp = tmp & " Select NCDEXCode,TransNo, '' CARNo from Txn_ImportNcdexLotPleadge Where CARNo Is Null ANd BatchID2 = 0"
   
tmp = " Select NCDEXCode,TransNo, Convert(Varchar,CARNo) 'CARNo',Code,CMSELotId from Txn_ImportNcdexLotPleadge Where CARNo Is not Null"
tmp = tmp & " Union All"
tmp = tmp & " Select NCDEXCode,TransNo, '' CARNo,Code,CMSELotId from Txn_ImportNcdexLotPleadge Where CARNo Is Null ANd BatchID1 = 0"
tmp = tmp & " Union All"
tmp = tmp & " Select NCDEXCode,TransNo, '' CARNo,Code,CMSELotId from Txn_ImportNcdexLotPleadge Where CARNo Is Null ANd BatchID2 = 0"
tmp = tmp & " Union All"
tmp = tmp & " Select NCDEXCode,TransNo, '' CARNo,Code,CMSELotId from Txn_ImportNcdexLotPleadge Where CARNo Is Null And BatchID1 <> 0 And BatchID2 = 0"
tmp = tmp & " Union All"
'tmp = tmp & " Select NCDEXCode,TransNo, '' CARNo,Code,CMSELotId from Txn_ImportNcdexLotPleadge Where CARNo Is Null And Code in (Select Distinct ISINID from Txn_PledgeCreation a"
'tmp = tmp & " inner join Txn_PledgeCreation_GSL b On a.SM_Prefix = b.SM_Prefix And a.PledgeID = b.PledgeID"
'tmp = tmp & " WHere b.PledgeNoOfBags > 0 And isnull(a.ISINID,0) <> '0')"
tmp = tmp & " Select NCDEXCode,TransNo, '' CARNo,Code,CMSELotId from Txn_ImportNcdexLotPleadge Where CARNo Is Null And CMSELotId in (Select Distinct b.ISINID from Txn_PledgeCreation_GSL a inner join Mst_GSL b on a.SM_Prefix = b.SM_Prefix And a.Mst_GSL_ID = b.Mst_GSL_ID Where  a.PledgeNoOfBags > 0 And isnull(b.ISINID,0) <> '0')   "
   
   
Call Tmp3Table

End If

ErrMsg = ""
RowCount = MSHFlexGrid1.Rows
ProgressBar1.Value = 0
ProgressBar1.Max = RowCount
Dim r As Integer

If ChkCommTrack.Value = 0 Then

   For r = 1 To RowCount - 2
       wc = ""
       mStatus = 0
       RsMst_RRN.Filter = ""
       If RsMst_RRN.RecordCount > 0 Then
          RsMst_RRN.MoveFirst
       End If
       RsMst_RRN.Filter = "RRNNo=" & MSHFlexGrid1.TextMatrix(r, 0) & " and DepositoryID='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
       If RsMst_RRN.RecordCount = 0 Then
          RsMst_RRN.Filter = ""
          If MSHFlexGrid1.TextMatrix(r, 2) <> "" Then
             If RsMst_ISIN.RecordCount > 0 Then
                RsMst_ISIN.MoveFirst
                RsMst_ISIN.Find "ISINID = '" & MSHFlexGrid1.TextMatrix(r, 2) & "'"
             End If
      
             If RsMst_ISIN.EOF = True Then
                ErrMsg = ErrMsg & "ISIN Not Found in ISIN Master " & MSHFlexGrid1.TextMatrix(r, 2) & "  At Row " & r + 1 & vbCrLf
                MSHFlexGrid1.Col = 2
                MSHFlexGrid1.row = r
                MSHFlexGrid1.CellBackColor = vbRed
                mStatus = 1
             End If
          Else
             ErrMsg = ErrMsg & "ISIN Not Found in Grid At Row " & r + 1 & vbCrLf
             MSHFlexGrid1.Col = 2
             MSHFlexGrid1.row = r
             MSHFlexGrid1.CellBackColor = vbRed
             mStatus = 1
          End If
       
              
          RsMst_Depository.Filter = ""
              
          If MSHFlexGrid1.TextMatrix(r, 5) <> "" Then
             If RsMst_Depository.RecordCount > 0 Then
                RsMst_Depository.MoveFirst
                RsMst_Depository.Filter = "DepositoryID = '" & MSHFlexGrid1.TextMatrix(r, 5) & "' And ExchangeTypeID = " & mExchangeTypeID & ""
             End If
      
             If RsMst_Depository.RecordCount > 0 Then
                mDPID = RsMst_Depository!DepositoryID
                RsMst_Depository.Filter = ""
             Else
                ErrMsg = ErrMsg & "DP ID Not Fount for DP in DP Master " & MSHFlexGrid1.TextMatrix(r, 5) & "  At Row " & r + 1 & vbCrLf
                MSHFlexGrid1.Col = 5
                MSHFlexGrid1.row = r
                MSHFlexGrid1.CellBackColor = vbRed
                mStatus = 1
             End If
          Else
             ErrMsg = ErrMsg & "DP ID Not Found in Grid At Row " & r + 1 & vbCrLf
             MSHFlexGrid1.Col = 5
             MSHFlexGrid1.row = r
             MSHFlexGrid1.CellBackColor = vbRed
             mStatus = 1
          End If
       
          If MSHFlexGrid1.TextMatrix(r, 6) <> "" Then
             RsMst_Client.Filter = "ClientID='" & MSHFlexGrid1.TextMatrix(r, 6) & "' And DepositoryID='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
             If RsMst_Client.RecordCount > 0 Then
                mClientIDRow = RsMst_Client!ClientIDRow
                RsMst_Client.Filter = ""
                RsMst_Client.MoveFirst
             Else
                ErrMsg = ErrMsg & "Client Not Found for Client ID in Client Master " & MSHFlexGrid1.TextMatrix(C, 6) & "  At Row " & r + 1 & vbCrLf
                MSHFlexGrid1.Col = 6
                MSHFlexGrid1.row = r
                MSHFlexGrid1.CellBackColor = vbRed
                If RsMst_Depository.RecordCount > 0 Then
                   Call AddClientToMaster(r)
                   MSHFlexGrid1.Col = 6
                   MSHFlexGrid1.row = r
                   MSHFlexGrid1.CellBackColor = vbBlue
                   ErrMsg = ErrMsg & "Client " & MSHFlexGrid1.TextMatrix(r, 6) & "  At Row " & r + 1 & " was not in Masters so created " & vbCrLf
                End If
             End If
          Else
             ErrMsg = ErrMsg & "Client Not Found Grid At Row " & r + 1 & vbCrLf
             MSHFlexGrid1.Col = 6
             MSHFlexGrid1.row = r
             MSHFlexGrid1.CellBackColor = vbRed
             mStatus = 1
          End If
          If mStatus = 0 Then
             RsMst_RRN.AddNew
             RsMst_RRN!RRNID = GetRRNID
             RsMst_RRN!RRNNo = MSHFlexGrid1.TextMatrix(r, 0)
             RsMst_RRN!RRNDate = MSHFlexGrid1.TextMatrix(r, 1)
             RsMst_RRN!ISINID = MSHFlexGrid1.TextMatrix(r, 2)
             RsMst_RRN!DepositoryID = mDPID 'MSHFlexGrid1.TextMatrix(c, 5)
             RsMst_RRN!ClientIDRow = mClientIDRow
             RsMst_RRN!RRNQty = Val(MSHFlexGrid1.TextMatrix(r, 8))
             RsMst_RRN!CommTrackFlag = 0
             RsMst_RRN!G_UID = GetGUID
             RsMst_RRN!CreatedBy = Gen_UserLoginID
             RsMst_RRN!CreatedDate = Now 'Gen_UserLoginID
             RsMst_RRN.Update
          End If
       Else
          If Val(RsMst_RRN!RRNQty) <> Val(MSHFlexGrid1.TextMatrix(r, 8)) Then
             ErrMsg = ErrMsg & "Record Already Present but Mismatch in RRNQty of RRN No " & MSHFlexGrid1.TextMatrix(r, 0) & " At Row " & r + 1 & vbCrLf
             MSHFlexGrid1.Col = 0
             MSHFlexGrid1.row = r
             MSHFlexGrid1.CellBackColor = vbYellow
          Else
             ErrMsg = ErrMsg & "Record Already Present of RRN No " & MSHFlexGrid1.TextMatrix(r, 0) & " At Row " & r + 1 & vbCrLf
             MSHFlexGrid1.Col = 0
             MSHFlexGrid1.row = r
             MSHFlexGrid1.CellBackColor = vbGreen
          End If
       End If
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next

   ProgressBar1.Value = ProgressBar1.Max
   TxtErrorMsg = ErrMsg

Else
   y = 0
   For I = 1 To MSHFlexGrid1.Rows - 2
       mADDMasterID = True
       y = y + 1
'       If MSHFlexGrid1.TextMatrix(y, 0) = "TWH14615" Then
'          parag = parag
'       End If
       If MSHFlexGrid1.Rows > 3 Then
          If y >= MSHFlexGrid1.Rows - 1 Then Exit For
       End If
       'If MSHFlexGrid1.TextMatrix(y, E.ESMTestResultID) = "" Or MSHFlexGrid1.TextMatrix(y, E.ESMTestResultDate) = "" Or MSHFlexGrid1.TextMatrix(y, E.SampleReceiptDate) = "" Or MSHFlexGrid1.TextMatrix(y, E.SampleDrawnDate) = "" Or MSHFlexGrid1.TextMatrix(y, E.AnalysisCompletiondate) = "" Or MSHFlexGrid1.TextMatrix(y, E.IssueDate) = "" Or MSHFlexGrid1.TextMatrix(y, E.Quantity) = "" Or MSHFlexGrid1.TextMatrix(y, E.Location) = "" Or MSHFlexGrid1.TextMatrix(y, E.ClientName) = "" Or MSHFlexGrid1.TextMatrix(y, E.Commodity) = "" Or MSHFlexGrid1.TextMatrix(y, E.Variety) = "" Or MSHFlexGrid1.TextMatrix(y, E.FarmarName) = "" Or MSHFlexGrid1.TextMatrix(y, E.Farm_PlotNo) = "" Or MSHFlexGrid1.TextMatrix(y,0) = "" Or MSHFlexGrid1.TextMatrix(y, E.SampleDrawanBy) = "" Or MSHFlexGrid1.TextMatrix(y, E.ExpGCode) = "" Or MSHFlexGrid1.TextMatrix(y, E.PesticidesID) = "" Or MSHFlexGrid1.TextMatrix(y, E.Pesticides) = "" Or Val(MSHFlexGrid1.TextMatrix(y, E.Result)) = 0 Then
       If MSHFlexGrid1.TextMatrix(y, 0) = "" Or MSHFlexGrid1.TextMatrix(y, 1) = "" Or MSHFlexGrid1.TextMatrix(y, 2) = "" Or MSHFlexGrid1.TextMatrix(y, 3) = "" Or Val(MSHFlexGrid1.TextMatrix(y, 4)) = 0 Then
          mSalary = False
          If mSalary = False Then
             For k = 0 To MSHFlexGrid1.Cols - 1
                 MSHFlexGrid1.row = y
                 MSHFlexGrid1.Col = k
                 MSHFlexGrid1.CellBackColor = vbRed
             Next
             
             mErrorMsg = mErrorMsg & " All/Some of component are bank for RRN at Row no " & y & " " & vbCrLf
             MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
             MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "All/Some of component are bank for RRN at Row no " & y & ""
             GoTo NextRow
          End If
      Else
          If RsMst_RRN.RecordCount > 0 Then
             RsMst_RRN.MoveFirst
             RsMst_RRN.Find "RRNNo = '" & MSHFlexGrid1.TextMatrix(y, 0) & "'"
          End If
          If RsMst_RRN.EOF = False Then
             For k = 0 To MSHFlexGrid1.Cols - 1
                 MSHFlexGrid1.row = y
                 MSHFlexGrid1.Col = k
                 MSHFlexGrid1.CellBackColor = vbGreen
             Next
             If InStr(mErrorMsg, "RRN No " & MSHFlexGrid1.TextMatrix(y, 0) & " already exist in RRN Master") <= 0 Then
                mErrorMsg = mErrorMsg & " RRN No " & MSHFlexGrid1.TextMatrix(y, 0) & " already exist in RRN Master" & vbCrLf
                MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
                MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "RRN No " & MSHFlexGrid1.TextMatrix(y, 0) & " already exist in RRN Master"
             End If
             GoTo NextRow
          Else
             If ChkRowValidation(y) = False Then
                GoTo NextRow
             End If
         End If
         If mADDMasterID = True Then
            RsMst_RRN.AddNew
            RsMst_RRN!RRNID = GetRRNID
            mRRNID = RsMst_RRN!RRNID
            RsMst_RRN!G_UID = GetGUID
            
            RsMst_RRN!RRNDate = CDate(MSHFlexGrid1.TextMatrix(y, 1))
            RsMst_RRN!RRNQty = Val(MSHFlexGrid1.TextMatrix(y, 4))
            RsMst_RRN!ClientIDRow = mClientIDRow
            RsMst_RRN!ISINID = mISINID
            RsMst_RRN!RRNNo = Trim(MSHFlexGrid1.TextMatrix(y, 0))
            RsMst_RRN!DeliveredQty = 0
            RsMst_RRN!EntryDate = Date
            RsMst_RRN!DepositoryID = mDepositoryID
            RsMst_RRN!CA_Qty = 0
            RsMst_RRN!OutOfSystemQty = 0
            RsMst_RRN!Remarks = "Uploaded by " & Gen_UserName & " And uploaded on " & Format(Date, "dd-mmm-yyyy") & "   "
            RsMst_RRN!CommTrackFlag = 1
            RsMst_RRN!CreatedBy = Gen_UserLoginID
            RsMst_RRN!CreatedDate = Now
            RsMst_RRN.Update
            mADDMasterID = False
            mRRNNo = MSHFlexGrid1.TextMatrix(y, 0)
            mTotalRRNQty = 0
            mRRNUpdated = False
         End If
         
         For xy = y To MSHFlexGrid1.Rows - 2
             If mRRNNo = MSHFlexGrid1.TextMatrix(xy, 0) Then
                mRRNNo = MSHFlexGrid1.TextMatrix(xy, 0)
                RsMst_RRNLotDetail.AddNew
                RsMst_RRNLotDetail!LotDetailID = GetRRN_LotDetailID
                RsMst_RRNLotDetail!RRNID = mRRNID
                RsMst_RRNLotDetail!ISINID = Trim(MSHFlexGrid1.TextMatrix(xy, 2))
                RsMst_RRNLotDetail!RematQty = Val(MSHFlexGrid1.TextMatrix(xy, 4))
                RsMst_RRNLotDetail!G_UID = GetGUID
                RsMst_RRNLotDetail!CreatedBy = Gen_UserLoginID
                RsMst_RRNLotDetail!CreatedDate = Now
                RsMst_RRNLotDetail.Update
                'RsMst_RRNLotDetail.Requery
                mTotalRRNQty = Val(mTotalRRNQty) + Val(MSHFlexGrid1.TextMatrix(xy, 4))
             Else
                mADDMasterID = True
                y = xy - 1
                tmp = "Update Mst_RRN Set RRNQty = " & mTotalRRNQty & " Where RRNID = " & mRRNID & " "
                Call TmpTable
                mRRNUpdated = True
                GoTo NextRow
             End If
         Next
         If mRRNUpdated = False And mTotalRRNQty > 0 Then
            tmp = "Update Mst_RRN Set RRNQty = " & mTotalRRNQty & " Where RRNID = " & mRRNID & " "
            Call TmpTable
         End If
       End If
NextRow:
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
   TxtErrorMsg = mErrorMsg
End If

MsgBox "Done!!!"

CmdExcel.Enabled = True

End Sub

Private Function ChkRowValidation(mRowNo_ As Variant) As Boolean
Dim x As Variant
ChkRowValidation = True

x = mRowNo_
mRRNNo = MSHFlexGrid1.TextMatrix(x, 0)
mISINID = ""
mClientIDRow = 0

For a = x To MSHFlexGrid1.Rows - 1
    If mRRNNo <> MSHFlexGrid1.TextMatrix(a, 0) Then
       If ChkRowValidation = False Then
          y = a - 1
       End If
       Exit For
    End If
        
    If MSHFlexGrid1.TextMatrix(a, 0) = "" Or MSHFlexGrid1.TextMatrix(a, 1) = "" Or MSHFlexGrid1.TextMatrix(a, 2) = "" Or MSHFlexGrid1.TextMatrix(a, 3) = "" Or Val(MSHFlexGrid1.TextMatrix(a, 4)) = 0 Then
       mErrorMsg = mErrorMsg & " Blank data not allowed At Row " & a & " " & vbCrLf
       
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "Blank data not allowed At Row " & a & ""
       
       ChkRowValidation = False
    
       For k = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.row = a
           MSHFlexGrid1.Col = k
           MSHFlexGrid1.CellBackColor = vbRed
       Next
    
    End If
    If CDate(MSHFlexGrid1.TextMatrix(a, 1)) > Date Then
       If InStr(mErrorMsg, "Future RRN Date Not allowed For RRN No " & Trim(MSHFlexGrid1.TextMatrix(a, 0)) & "") <= 0 Then
          mErrorMsg = mErrorMsg & " Future RRN Date Not allowed For RRN No " & Trim(MSHFlexGrid1.TextMatrix(a, 0)) & " " & vbCrLf
          MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
          MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "Future RRN Date Not allowed For RRN No " & Trim(MSHFlexGrid1.TextMatrix(a, 0)) & ""
       End If
       ChkRowValidation = False
       For k = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.row = a
           MSHFlexGrid1.Col = k
           MSHFlexGrid1.CellBackColor = vbRed
       Next
    End If
    If RsMst_ISIN.RecordCount > 0 Then
       RsMst_ISIN.MoveFirst
       RsMst_ISIN.Find "ISINID = '" & Trim(MSHFlexGrid1.TextMatrix(a, 2)) & "'"
       If RsMst_ISIN.EOF = True Then
          If InStr(mErrorMsg, "ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " not exist in Store-Man") <= 0 Then
             mErrorMsg = mErrorMsg & " ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " not exist in Store-Man" & vbCrLf
             MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
             MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " not exist in Store-Man"
          End If
          ChkRowValidation = False
          For k = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.row = a
              MSHFlexGrid1.Col = k
              MSHFlexGrid1.CellBackColor = vbRed
          Next
       Else
          If mISINID = "" Then
             mISINID = Trim(MSHFlexGrid1.TextMatrix(a, 2))
          End If
       End If
    End If
    
'    tmp = "Select * from Mst_RRN Where ISINID = '" & Trim(MSHFlexGrid1.TextMatrix(a, 2)) & "'"
'    Call Tmp1Table
'
'    tmp = "Select Distinct b.RRNNo,b.RRNID from Mst_RRN_LotDetail a Inner Join Mst_RRN b on a.RRNID = b.RRNID Where a.ISINID = '" & Trim(MSHFlexGrid1.TextMatrix(a, 2)) & "' "
'    Call Tmp2Table
    
    
    If TmpRs1.RecordCount > 0 Then
       TmpRs1.MoveFirst
       TmpRs1.Find "ISINID = '" & Trim(MSHFlexGrid1.TextMatrix(a, 2)) & "'"
       If TmpRs1.EOF = False Then
          If InStr(mErrorMsg, "ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " already mapped to RRN No " & TmpRs1!RRNNo & "  in Store-Man") <= 0 Then
             mErrorMsg = mErrorMsg & " ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " already mapped to RRN No " & TmpRs1!RRNNo & "  in Store-Man" & vbCrLf
             MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
             MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " already mapped to RRN No " & TmpRs1!RRNNo & "  in Store-Man"
          End If
          ChkRowValidation = False
          For k = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.row = a
              MSHFlexGrid1.Col = k
              MSHFlexGrid1.CellBackColor = vbRed
          Next
       End If
    End If
    
'    tmp = "Select Distinct b.RRNNo,b.RRNID,a.ISINID from Mst_RRN_LotDetail a Inner Join Mst_RRN b on a.RRNID = b.RRNID " 'Where a.ISINID = '" & Trim(MSHFlexGrid1.TextMatrix(a, 2)) & "' "
'    Call Tmp2Table
    If TmpRs2.RecordCount > 0 Then
       TmpRs2.MoveFirst
       TmpRs2.Find "ISINID = '" & Trim(MSHFlexGrid1.TextMatrix(a, 2)) & "'"
       If TmpRs2.EOF = False Then
          If InStr(mErrorMsg, "ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " already mapped to RRN No " & TmpRs2!RRNNo & "  in Store-Man") <= 0 Then
             mErrorMsg = mErrorMsg & " ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " already mapped to RRN No " & TmpRs2!RRNNo & "  in Store-Man" & vbCrLf
             MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
             MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " already mapped to RRN No " & TmpRs2!RRNNo & "  in Store-Man"
          End If
          ChkRowValidation = False
          For k = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.row = a
              MSHFlexGrid1.Col = k
              MSHFlexGrid1.CellBackColor = vbRed
          Next
       End If
    End If
    
    
    If TmpRs3.RecordCount > 0 Then
       TmpRs3.MoveFirst
       TmpRs3.Find "CMSELotId = '" & Trim(MSHFlexGrid1.TextMatrix(a, 2)) & "'"
       If TmpRs3.EOF = False Then
          If InStr(mErrorMsg, "ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " already mapped for demat funding against Trans No " & TmpRs3!TransNo & "  in Store-Man") <= 0 Then
             mErrorMsg = mErrorMsg & " ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " already mapped for demat funding against Trans No " & TmpRs3!TransNo & "  in Store-Man" & vbCrLf
             MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
             MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "ISIN No " & MSHFlexGrid1.TextMatrix(a, 2) & " already mapped for demat funding against Trans No " & TmpRs3!TransNo & "  in Store-Man"
          End If
          ChkRowValidation = False
          For k = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.row = a
              MSHFlexGrid1.Col = k
              MSHFlexGrid1.CellBackColor = vbRed
          Next
       End If
    End If
    
    
    If RsMst_Client.RecordCount > 0 Then
       RsMst_Client.MoveFirst
       RsMst_Client.Find "ClientID  = '" & MSHFlexGrid1.TextMatrix(a, 3) & "'"
       If RsMst_Client.EOF = True Then
          If InStr(mErrorMsg, "Client ID " & MSHFlexGrid1.TextMatrix(a, 3) & " not exist in Store-Man") <= 0 Then
             mErrorMsg = mErrorMsg & " Client ID " & MSHFlexGrid1.TextMatrix(a, 3) & " not exist in Store-Man" & vbCrLf
             MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
             MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "Client ID " & MSHFlexGrid1.TextMatrix(a, 3) & " not exist in Store-Man"
          End If
          ChkRowValidation = False
          For k = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.row = a
              MSHFlexGrid1.Col = k
              MSHFlexGrid1.CellBackColor = vbRed
          Next
       
       Else
          If mClientIDRow = 0 Then
             mClientIDRow = RsMst_Client!ClientIDRow ' GetClientID(MSHFlexGrid1.TextMatrix(a, E.ClientName))
             mDepositoryID = RsMst_Client!DepositoryID
          End If
       End If
    End If
Next

For a = x To MSHFlexGrid1.Rows - 1
    mISINNo = MSHFlexGrid1.TextMatrix(a, 2)
    If mRRNNo <> MSHFlexGrid1.TextMatrix(a, 0) Then
       Exit For
    End If
    For ab = x To MSHFlexGrid1.Rows - 1
        If mRRNNo <> MSHFlexGrid1.TextMatrix(ab, 0) Then
           Exit For
        End If
        If a <> ab Then
           If mISINNo = MSHFlexGrid1.TextMatrix(ab, 2) Then
              If InStr(mErrorMsg, "Duplicate ISIN " & MSHFlexGrid1.TextMatrix(a, 2) & " not allowed") <= 0 Then
                 mErrorMsg = mErrorMsg & " Duplicate ISIN " & MSHFlexGrid1.TextMatrix(a, 2) & " not allowed" & vbCrLf
                 MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
                 MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = "Duplicate ISIN " & MSHFlexGrid1.TextMatrix(a, 2) & " not allowed"
              End If
              ChkRowValidation = False
              For k = 0 To MSHFlexGrid1.Cols - 1
                  MSHFlexGrid1.row = a
                  MSHFlexGrid1.Col = k
                  MSHFlexGrid1.CellBackColor = vbRed
              Next
           End If
        End If
    Next
Next

End Function

Private Function GetRRNID() As Double
Dim Retval As Double
tmp = "Select max(RRNID) from Mst_RRN"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetRRNID = Retval
End Function

Private Function GetRRN_LotDetailID() As Double
Dim Retval As Double
tmp = "Select max(LotDetailID) from Mst_RRN_LotDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetRRN_LotDetailID = Retval
End Function


Private Sub Form_Load()
LblFile.Caption = ""
Frame1.Enabled = False
Frame2.Enabled = False
ChkCommTrack.Value = 0
Call Grid_Head
ProgressBar1.Value = 0

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

mExchangeTypeID = GetExchangeID("NCDEX")
Frame1.Enabled = True
CmdImport.Enabled = False
CmdSave.Enabled = False
CmdReset.Enabled = False
CmdExcel.Enabled = False

End Sub

Private Sub FillData()
'mDate = Split(mNE_Date_, "/")
'StrDate = mDate(2) & "/" & mDate(1) & "/" & mDate(0)
''Set oWS = oWB.Worksheets(CmbSheetName.Text)
oWS.ClearArrows
mRow = 1
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
LblTotalRecords.Caption = RowCount - 1
If ChkCommTrack.Value = 0 Then
   For l = 1 To RowCount
       mRow = mRow + 1
       If oWS.Cells(mRow, 1) <> "" Then
          For C = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(l, C) = Trim(oWS.Cells(mRow, C + 1).Value)
          Next
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       End If
   Next
Else

   For l = 1 To RowCount
       mRow = mRow + 1
       If oWS.Cells(mRow, 1) <> "" Then
          For C = 0 To MSHFlexGrid1.Cols - 1
            If InStr(Trim(oWS.Cells(mRow, C + 1).Value), "-") > 0 Then
'               If CDate(Trim(oWS.Cells(mRow, c + 1).Value)) = True Then
               mDate = Split(Trim(oWS.Cells(mRow, C + 1).Value), "-")
               If mDate(0) > 12 Then
                  mDate = Format(Trim(oWS.Cells(mRow, C + 1).Value), "mm/dd/yyyy")
               Else
                  mDate = Format(Trim(oWS.Cells(mRow, C + 1).Value), "dd/mm/yyyy")
               End If
'                  mDate = Split(mDate, "/")
'                  StrDate = mDate(2) & "/" & mDate(1) & "/" & mDate(0)
               MSHFlexGrid1.TextMatrix(l, C) = CDate(mDate) ' Trim(oWS.Cells(mRow, c + 1).Value) 'CDate(StrDate)  '
'               Else
'                  MSHFlexGrid1.TextMatrix(L, c) = Trim(oWS.Cells(mRow, c + 1).Value)
'               End If
            Else
              MSHFlexGrid1.TextMatrix(l, C) = Trim(oWS.Cells(mRow, C + 1).Value)
            End If
          Next
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       End If
   Next
End If
oXL.Quit
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 12
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "RRN No"
    .TextMatrix(0, 1) = "RRN Date"
    .TextMatrix(0, 2) = "ISIN"
    .TextMatrix(0, 3) = "CMP"
    .TextMatrix(0, 4) = "Commodity"
    .TextMatrix(0, 5) = "DP ID"
    .TextMatrix(0, 6) = "Client ID"
    .TextMatrix(0, 7) = "Client Name"
    .TextMatrix(0, 8) = "Remat"
    .TextMatrix(0, 9) = "Confirmed Qty at DP"
    .TextMatrix(0, 10) = "Confirmed Qty at NCOM"
    .TextMatrix(0, 11) = "Actual Pending Qty"
   
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
    .ColWidth(2) = 2000
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
    .ColWidth(5) = 2000
    .ColWidth(6) = 1600
    .ColWidth(7) = 2500
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    
    
End With
End Sub
Private Sub Grid_HeadCommtrack()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 5
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "RRN No"
    .TextMatrix(0, 1) = "RRN Date"
    .TextMatrix(0, 2) = "ISIN No"
    .TextMatrix(0, 3) = "Client ID"
    .TextMatrix(0, 4) = "RRN Qty"
   
    .ColWidth(0) = 2000
    .ColWidth(1) = 2000
    .ColWidth(2) = 2000
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
   
    
End With
End Sub

Private Sub Grid_HeadError()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .Cols = 1
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Error Msg"
    .ColWidth(0) = 6000
End With
End Sub

Private Sub AddClientToMaster(Row_ As Integer)

RsMst_Client.AddNew
RsMst_Client!ClientIDRow = GetRowID
mClientIDRow = RsMst_Client!ClientIDRow
RsMst_Client!G_UID = GetGUID
RsMst_Client!ClientID = MSHFlexGrid1.TextMatrix(Row_, 6) '  .Text
RsMst_Client!ClientName = MSHFlexGrid1.TextMatrix(Row_, 7) 'TxtDepositorName.Text
RsMst_Client!ExchangeTypeID = mExchangeTypeID
RsMst_Client!DepositoryID = MSHFlexGrid1.TextMatrix(Row_, 5) 'mDepositoryID
RsMst_Client!Flag = "0" 'TxtFlag.Text
RsMst_Client!CreatedBy = Gen_UserLoginID
RsMst_Client!CreatedDate = Now
RsMst_Client.Update


End Sub

Private Function GetRowID() As Double
Dim Retval As Double
tmp = "Select max(ClientIDRow) from Mst_Client "
Call Tmp1Table
Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
Retval = Retval + 1
TmpRs1.Close
GetRowID = Retval
End Function

Private Function GetCARowID() As Double
Dim Retval As Double
tmp = "Select max(RowID) from Mst_CA "
Call Tmp1Table
Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
Retval = Retval + 1
TmpRs1.Close
GetCARowID = Retval
End Function
