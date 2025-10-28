VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_CAUpload 
   Caption         =   "CA Upload"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Error Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   990
      Width           =   15000
      Begin VB.TextBox TxtErrorMsg 
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   240
         Width           =   10065
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
         Left            =   10800
         TabIndex        =   19
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Mismatch in ISIN in RRN Master and Selected File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11520
         TabIndex        =   18
         Top             =   765
         Width           =   3285
      End
      Begin VB.Label Label8 
         Caption         =   "Records not in RRN Master Table"
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
         Left            =   11565
         TabIndex        =   16
         Top             =   1335
         Width           =   3255
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
         Left            =   10800
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Already in CA Master "
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
         Left            =   11565
         TabIndex        =   14
         Top             =   375
         Width           =   2325
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
         Left            =   10800
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   6150
         TabIndex        =   17
         Top             =   165
         Width           =   1575
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
         TabIndex        =   5
         Top             =   165
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
         Top             =   165
         Width           =   1455
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
         Left            =   7725
         TabIndex        =   3
         Top             =   165
         Width           =   1695
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
         TabIndex        =   2
         Top             =   165
         Width           =   1695
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
         TabIndex        =   1
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   645
         Width           =   705
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
         Left            =   10260
         TabIndex        =   7
         Top             =   195
         Width           =   2295
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   12585
         TabIndex        =   6
         Top             =   255
         Width           =   675
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   11033
      _Version        =   393216
      WordWrap        =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   9240
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_CAUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet

Dim ErrMsg As Variant
Dim Status As Variant
Dim mSheetFlg, mExchangeTypeID, mRRNID As Variant

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "CAForRRNUpload.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "CAForRRNUpload.xls")
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
        End If
    Next
Next
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

If LblFile.Caption <> "" Then
   CmdExit.Enabled = True
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   Dim I As Integer
   Dim runApp As Boolean
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   oXL.Quit
   MsgBox "Select Import To Import Data From Excel Sheet!!!"
   CmdImport.SetFocus
End If

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub CmdExit_Click()
On Error Resume Next
oXL.Quit
Unload Me
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
'   MsgBox "Please select proper file !!!"
   oXL.Quit
   Exit Sub
End If
Call FillData
CmdSave.Enabled = True
CmdReset.Enabled = True
CmdSave.SetFocus
MsgBox "Import Succesful!!!"

End Sub
Private Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True

If LCase(Trim(oWS.Cells(1, 1))) <> "rrn" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 1 "
ElseIf LCase(Trim(oWS.Cells(1, 2))) <> "isin" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 2 "
ElseIf LCase(Trim(oWS.Cells(1, 3))) <> "isin description" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 3 "
ElseIf LCase(Trim(oWS.Cells(1, 4))) <> "depository" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 4 "
ElseIf LCase(Trim(oWS.Cells(1, 5))) <> "dpid" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 5 "
ElseIf LCase(Trim(oWS.Cells(1, 6))) <> "clientid" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 6 "
ElseIf LCase(Trim(oWS.Cells(1, 7))) <> "quantity" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 7 "
ElseIf LCase(Trim(oWS.Cells(1, 8))) <> "partial acceptance" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 8 "
ElseIf LCase(Trim(oWS.Cells(1, 9))) <> "ca no." Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 9 "
ElseIf LCase(Trim(oWS.Cells(1, 10))) <> "ca date" Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 10 "
End If

CheckFile = Retval

End Function

Private Sub CmdReset_Click()

Call Grid_Head
CmdExit.Enabled = True
CmdSave.Enabled = False
CmdGetFile.Enabled = True
CmdImport.Enabled = False
LblFile.Caption = ""
LblTotalRecords.Caption = ""
mSheetFlg = ""
TxtErrorMsg = ""
ProgressBar1.Value = 0

End Sub

Private Sub CmdSave_Click()
Dim ErrMsg As Variant

CmdSave.Enabled = False

If MSHFlexGrid1.Rows < 3 Then
   MsgBox "No data to save!!!"
   Exit Sub
End If

If MSHFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
MsgBox "Wait till Next Message....."
 
Call TableMst_CA
Call TableMst_RRN
Call TableMst_RRN_LotDetail
'Call TableMst_ISIN
'Call TableMst_Client
'Call TableMst_Depository

ErrMsg = ""
RowCount = MSHFlexGrid1.Rows
ProgressBar1.Value = 0
ProgressBar1.Max = RowCount
Dim r As Integer
For r = 1 To RowCount - 3
    wc = ""
    mStatus = 0
    
    If MSHFlexGrid1.TextMatrix(r, 0) <> "" Then
       RsMst_RRN.Filter = ""
       If RsMst_RRN.RecordCount > 0 Then
          RsMst_RRN.MoveFirst
          RsMst_RRN.Filter = "RRNNo = '" & MSHFlexGrid1.TextMatrix(r, 0) & "' And DepositoryID = '" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
       End If
   
       If RsMst_RRN.RecordCount > 0 Then
          RsMst_RRN.MoveFirst
          mRRNID = RsMst_RRN!RRNID
          'RsMst_RRN.Filter "RRNNo = '" & MSHFlexGrid1.TextMatrix(r, 0) & "' And DepositoryID = '" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
          If RsMst_RRN!CommTrackFlag <> "1" Then
             If RsMst_RRN!ISINID = MSHFlexGrid1.TextMatrix(r, 1) Then
                RsMst_CA.Filter = ""
                If RsMst_CA.RecordCount > 0 Then
                   RsMst_CA.MoveFirst
                End If
                If MSHFlexGrid1.TextMatrix(r, 8) <> "" Then
                   RsMst_CA.Filter = "RRNID= " & mRRNID & " and CANo= " & MSHFlexGrid1.TextMatrix(r, 8) & ""
                   If RsMst_CA.RecordCount = 0 Then
                      If mStatus = 0 Then
                         RsMst_CA.AddNew
                         RsMst_CA!RowID = GetRowID
                         RsMst_CA!RRNID = mRRNID
                         RsMst_CA!CANo = MSHFlexGrid1.TextMatrix(r, 8)
                         RsMst_CA!CADate = MSHFlexGrid1.TextMatrix(r, 9)
                         RsMst_CA!Quantity = Val(MSHFlexGrid1.TextMatrix(r, 6))
                         RsMst_CA!PartialAcceptance = Val(MSHFlexGrid1.TextMatrix(r, 7))
                         RsMst_CA!G_UID = GetGUID
                         RsMst_CA!CreatedBy = Gen_UserLoginID
                         RsMst_CA!CreatedDate = Now
                         RsMst_CA.Update
                
                 
                         RsMst_RRN.Find "RRNID = " & mRRNID & ""
                         RsMst_RRN!CA_Qty = Val(RsMst_RRN!CA_Qty) + Val(MSHFlexGrid1.TextMatrix(r, 7))
                         RsMst_RRN!UpdatedBy = Gen_UserLoginID
                         RsMst_RRN!UpdatedDate = Now
                         RsMst_RRN.Update
                      End If
                   Else
                      ErrMsg = ErrMsg & "Record Already Present of CA No " & MSHFlexGrid1.TextMatrix(r, 8) & " At Row " & r + 1 & vbCrLf
                      MSHFlexGrid1.Col = 8
                      MSHFlexGrid1.row = r
                      MSHFlexGrid1.CellBackColor = vbGreen
                   End If
               Else
                   ErrMsg = ErrMsg & "CA No. not found in Grid At Row " & r + 1 & vbCrLf
                   MSHFlexGrid1.Col = 8
                   MSHFlexGrid1.row = r
                   MSHFlexGrid1.CellBackColor = vbRed
               End If
             Else
               ErrMsg = ErrMsg & "Mismatch in ISIN found in Grid At Row " & r + 1 & vbCrLf
               MSHFlexGrid1.Col = 1
               MSHFlexGrid1.row = r
               MSHFlexGrid1.CellBackColor = vbYellow
            End If
          Else
             RsMst_RRNLotDetail.Filter = ""
             RsMst_RRNLotDetail.MoveFirst
             RsMst_RRNLotDetail.Filter = " RRNID = " & mRRNID & " And ISINID = '" & MSHFlexGrid1.TextMatrix(r, 1) & "'"
             If RsMst_RRNLotDetail.RecordCount > 0 Then
                RsMst_CA.Filter = ""
                If RsMst_CA.RecordCount > 0 Then
                   RsMst_CA.MoveFirst
                End If
                If MSHFlexGrid1.TextMatrix(r, 8) <> "" Then
                   RsMst_CA.Filter = "RRNID= " & mRRNID & " and CANo= " & MSHFlexGrid1.TextMatrix(r, 8) & " and ISINID = '" & MSHFlexGrid1.TextMatrix(r, 1) & "' "
                   If RsMst_CA.RecordCount = 0 Then
                      If mStatus = 0 Then
                         RsMst_CA.AddNew
                         RsMst_CA!RowID = GetRowID
                         RsMst_CA!RRNID = mRRNID
                         RsMst_CA!CANo = MSHFlexGrid1.TextMatrix(r, 8)
                         RsMst_CA!CADate = MSHFlexGrid1.TextMatrix(r, 9)
                         RsMst_CA!Quantity = Val(MSHFlexGrid1.TextMatrix(r, 6))
                         RsMst_CA!PartialAcceptance = Val(MSHFlexGrid1.TextMatrix(r, 7))
                         RsMst_CA!ISINID = MSHFlexGrid1.TextMatrix(r, 1)
                         RsMst_CA!G_UID = GetGUID
                         RsMst_CA!CreatedBy = Gen_UserLoginID
                         RsMst_CA!CreatedDate = Now
                         RsMst_CA.Update
                
                 
                         RsMst_RRN.Find "RRNID = " & mRRNID & ""
                         RsMst_RRN!CA_Qty = Val(RsMst_RRN!CA_Qty) + Val(MSHFlexGrid1.TextMatrix(r, 7))
                         RsMst_RRN!UpdatedBy = Gen_UserLoginID
                         RsMst_RRN!UpdatedDate = Now
                         RsMst_RRN.Update
                      End If
                   Else
                      ErrMsg = ErrMsg & "Record Already Present of CA No " & MSHFlexGrid1.TextMatrix(r, 8) & " At Row " & r + 1 & vbCrLf
                      MSHFlexGrid1.Col = 8
                      MSHFlexGrid1.row = r
                      MSHFlexGrid1.CellBackColor = vbGreen
                   End If
               Else
                   ErrMsg = ErrMsg & "CA No. not found in Grid At Row " & r + 1 & vbCrLf
                   MSHFlexGrid1.Col = 8
                   MSHFlexGrid1.row = r
                   MSHFlexGrid1.CellBackColor = vbRed
               End If
            Else
               ErrMsg = ErrMsg & "Mismatch in ISIN found in Grid At Row " & r + 1 & vbCrLf
               MSHFlexGrid1.Col = 1
               MSHFlexGrid1.row = r
               MSHFlexGrid1.CellBackColor = vbYellow
            End If
          End If
       Else
          ErrMsg = ErrMsg & "RRN No Not Found in RRN Master " & MSHFlexGrid1.TextMatrix(r, 0) & "  At Row " & r + 1 & vbCrLf
          MSHFlexGrid1.Col = 0
          MSHFlexGrid1.row = r
          MSHFlexGrid1.CellBackColor = vbRed
          'mStatus = 1
       End If
    Else
       ErrMsg = ErrMsg & "RRN No Not Found in Grid At Row " & r + 1 & vbCrLf
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.row = r
       MSHFlexGrid1.CellBackColor = vbRed
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max
TxtErrorMsg = ErrMsg
MsgBox "Done!!!"
CmdExcel.Enabled = True

End Sub

Private Sub Form_Load()
LblFile.Caption = ""
Frame1.Enabled = False
'Frame2.Enabled = False
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

'Set oWS = oWB.Worksheets(CmbSheetName.Text)
oWS.ClearArrows
mRow = 1
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
LblTotalRecords.Caption = RowCount - 1
For l = 1 To RowCount
    mRow = mRow + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(l, C) = Trim(oWS.Cells(mRow, C + 1))
    Next
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
Next
oXL.Quit
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1

    .Clear
    .Rows = 2
    .Cols = 10
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(0, 0) = "RRN No"
    .TextMatrix(0, 1) = "ISIN"
    .TextMatrix(0, 2) = "ISIN Description"
    .TextMatrix(0, 3) = "Depository"
    .TextMatrix(0, 4) = "DP ID"
    .TextMatrix(0, 5) = "Client ID"
    .TextMatrix(0, 6) = "Quantity"
    .TextMatrix(0, 7) = "Partial Acceptance"
    .TextMatrix(0, 8) = "CA No."
    .TextMatrix(0, 9) = "CA Date"
       
    .ColWidth(0) = 1000
    .ColWidth(1) = 1400
    .ColWidth(2) = 2400
    .ColWidth(3) = 2200
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
    .ColWidth(6) = 1000
    .ColWidth(7) = 1600
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
   
    
End With
End Sub

Private Function GetRowID() As Double
Dim Retval As Double
tmp = "Select max(RowID) from Mst_CA "
Call Tmp1Table
Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
Retval = Retval + 1
TmpRs1.Close
GetRowID = Retval
End Function
