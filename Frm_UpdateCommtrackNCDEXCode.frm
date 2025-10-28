VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_UpdateCommtrackNCDEXCode 
   Caption         =   "Update Commtrack NCDEX Code"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.CheckBox Check2 
         Caption         =   "All Godowns "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   5
         Top             =   300
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pending NCDEX Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         Top             =   300
         Width           =   2415
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
         Left            =   13635
         TabIndex        =   11
         Top             =   720
         Width           =   1095
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
         Left            =   10350
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptGodownCode 
         Caption         =   "Godown Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptCommodity 
         Caption         =   "Commodity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton OptLocation 
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
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   300
         Width           =   1095
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
         Height          =   375
         Left            =   8160
         TabIndex        =   6
         Top             =   720
         Width           =   1095
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
         Left            =   9255
         TabIndex        =   7
         Top             =   720
         Width           =   1095
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
         Height          =   375
         Left            =   11445
         TabIndex        =   9
         Top             =   720
         Width           =   1095
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
         Height          =   375
         Left            =   12540
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7095
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   12515
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
         Left            =   120
         TabIndex        =   15
         Top             =   8760
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label23 
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
         Left            =   9960
         TabIndex        =   14
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label Label22 
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
         Left            =   12480
         TabIndex        =   13
         Top             =   307
         Width           =   90
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
Attribute VB_Name = "Frm_UpdateCommtrackNCDEXCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim Status, ErrMsg As Variant

Private Sub CmdGetFile_Click()
On Error GoTo errortrap
With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls)|*.xls"
     .CancelError = True
procReOpen:
     .ShowOpen
End With
Path = CommonDialog1.FileName
LblFile.Caption = Path
CmdImport.Enabled = True
'CmdSave.Enabled = False
Exit Sub
errortrap:
   MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub UploadFile()
Status = 0
If LblFile.Caption <> "" And LCase(LblFile.Caption) <> "filename" Then
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   Status = 0
Else
   MsgBox "Please Select File", vbInformation, "Information"
   Status = 1
End If
End Sub


Private Sub CmdGo_Click()

Call Grid_Head
MsgBox "Wait till next message!!!"
If OptLocation.Value = True Then
   tmp = " Select upper(LocationName) LocationName, upper(NCDEXCODE) NCDEXCODE, upper(CommtrackNCDEXCODE) CommtrackNCDEXCODE, upper(NComNCDEXCODE) NComNCDEXCODE from Mst_Location"
ElseIf OptCommodity.Value = True Then
   tmp = " Select upper(Commodity) CommodityName, upper(NCDEXCODE) NCDEXCODE, upper(CommtrackNCDEXCODE) CommtrackNCDEXCODE, upper(NComNCDEXCODE) NComNCDEXCODE  from Mst_Commodity"
ElseIf OptGodownCode.Value = True Then
   tmp = " Select upper(cmp.cmpname)+'_'+upper(mg.godownno) GodownCode, upper(MG.NCDEXCODE) NCDEXCODE, upper(CommtrackNCDEXCODE) CommtrackNCDEXCODE, upper(NComNCDEXCODE) NComNCDEXCODE from Mst_Godown MG"
   tmp = tmp & " Inner join mst_Cmp cmp on cmp.cmpid = mg.cmpid "
End If
If Check1.Value = vbChecked Then
   tmp = tmp & " where (isnull(NCDEXCODE, '') = '' or isnull(CommtrackNCDEXCODE, '') = '' or isnull(NComNCDEXCODE, '') = '')"
End If
If OptLocation.Value = True Then
   tmp = tmp & " order by LocationName"
ElseIf OptCommodity.Value = True Then
   tmp = tmp & " order by CommodityName"
ElseIf OptGodownCode.Value = True Then
   If Check2.Value = vbUnchecked Then
      If Check1.Value = vbChecked Then
         tmp = tmp & " and MG.CloseDate is null"
      Else
         tmp = tmp & " where MG.CloseDate is null"
      End If
   End If
   tmp = tmp & " order by CMPName, GodownNo"
End If
Call TmpTable

If TmpRs.RecordCount > 0 Then
   'CmdExcel.Enabled = True
   ProgressBar1.Max = TmpRs.RecordCount + 3
   Label22.Caption = TmpRs.RecordCount
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      For I = 1 To TmpRs.RecordCount
         For C = 0 To MSHFlexGrid1.Cols - 1
            .TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
         Next
         TmpRs.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
      ProgressBar1.Value = ProgressBar1.Max
   End With
   MsgBox "Done!!!"
Else
   'CmdExcel.Enabled = False
   MsgBox "No record found!!!"
End If
End Sub

Private Sub Grid_Head()
Label22.Caption = ""
LblFile.Caption = ""
ProgressBar1.Value = 0
'CmdExcel.Enabled = False
CmdImport.Enabled = False
CmdSave.Enabled = False
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 4
   .WordWrap = True
   
   If OptLocation.Value = True Then
      .TextMatrix(0, 0) = "LOCATION NAME"
      '.TextMatrix(0, 2) = "LocationID"
   ElseIf OptCommodity.Value = True Then
      .TextMatrix(0, 0) = "COMMODITY NAME"
      '.TextMatrix(0, 2) = "CommodityID"
   ElseIf OptGodownCode.Value = True Then
      .TextMatrix(0, 0) = "GODOWN CODE"
      '.TextMatrix(0, 2) = "GodownID"
   End If
   .TextMatrix(0, 1) = "NON NCDEX COMMTRACK CODE"
   .TextMatrix(0, 2) = "NCDEX COMMTRACK CODE"
   .TextMatrix(0, 3) = "NCOM NCDEX CODE"
   
   .ColWidth(0) = 4500
   .ColWidth(1) = 3500
   .ColWidth(2) = 2500
   .ColWidth(3) = 2500
End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "UpdateNCDEXCODE.xls")
Call Xls_Detail_Rpt
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "UpdateNCDEXCODE.xls")

Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
      'MSHFlexGrid1.Row = I
      'MSHFlexGrid1.Col = c
      'If MSHFlexGrid1.BackColor = vbRed Then
      '   oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbRed
      'End If
   Next
   
   ProgressBar1.Value = ProgressBar1.Value + 1
Next

oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created!!!")
oXL.Visible = True
End Sub

Private Sub CmdImport_Click()
On Error GoTo msgError

Call UploadFile

If Status = 1 Then Exit Sub

   'CmdGetFile.Enabled = False
   '-----------------------------
   ErrMsg = ""
   Status = 0
   RowCount = 0
   I = 1
   '-----------------------------
   Call Validation
   If Status = 1 Then
      MsgBox ErrMsg, vbExclamation, "Message"
      'CmdGetFile.Enabled = True
   Else
      Call Grid_Head
      Call FillData
      'CmdExcel.Enabled = True
      MsgBox "Import Successfully!!!"
      CmdSave.Enabled = True
   End If

Exit Sub
msgError:
   MsgBox "Excel File Missing!!!"
End Sub

Private Sub FillData()
oWS.ClearArrows
mRow = 1
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
Label22.Caption = RowCount - 1
For l = 1 To RowCount
    mRow = mRow + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(l, C) = Trim(oWS.Cells(mRow, C + 1))
    Next
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
Next
oXL.Quit
End Sub

Private Sub Validation()
   Set oWS = oWB.Worksheets("Sheet1")
   WrkSheet = "Sheet1"
   If OptLocation.Value = True Then
      If UCase(Trim(oWS.Cells(1, 1))) <> "LOCATION NAME" Then
         ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A(LOCATION NAME) " & vbCrLf
         Status = 1
      End If
   ElseIf OptCommodity.Value = True Then
      If UCase(Trim(oWS.Cells(1, 1))) <> "COMMODITY NAME" Then
         ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A(COMMODITY NAME) " & vbCrLf
         Status = 1
      End If
   ElseIf OptGodownCode.Value = True Then
      If UCase(Trim(oWS.Cells(1, 1))) <> "GODOWN CODE" Then
         ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A(GODOWN CODE) " & vbCrLf
         Status = 1
      End If
   End If
   If UCase(Trim(oWS.Cells(1, 2))) <> "NON NCDEX COMMTRACK CODE" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of B(NON NCDEX COMMTRACK CODE) " & vbCrLf
      Status = 1
   End If
   If UCase(Trim(oWS.Cells(1, 3))) <> "NCDEX COMMTRACK CODE" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of C(NCDEX COMMTRACK CODE) " & vbCrLf
      Status = 1
   End If
   If UCase(Trim(oWS.Cells(1, 4))) <> "NCOM NCDEX CODE" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of D(NCOM NCDEX CODE) " & vbCrLf
      Status = 1
   End If
End Sub

Private Sub CmdReset_Click()
Call Grid_Head
'CmdSave.Enabled = False
'CmdGetFile.Enabled = True
'CmdImport.Enabled = False
End Sub

Private Sub CmdSave_Click()
ErrMsg = ""
On Error GoTo errortrap
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows - 1
If OptCommodity.Value = True Then
   Call TableMst_Commodity
ElseIf OptGodownCode.Value = True Then
   tmp = " Select upper(cmp.cmpname)+'_'+upper(mg.godownno) GodownCode, upper(MG.NCDEXCODE) NCDEXCODE, upper(MG.CommtrackNCDEXCODE) CommtrackNCDEXCODE, upper(MG.NComNCDEXCODE) NComNCDEXCODE, MG.GodownID from Mst_Godown MG"
   tmp = tmp & " Inner join mst_Cmp cmp on cmp.cmpid = mg.cmpid "
   Call Tmp5Table
   
   Call TableMst_Godown
ElseIf OptLocation.Value = True Then
   Call TableMst_Location
End If

For I = 1 To MSHFlexGrid1.Rows - 1
   If MSHFlexGrid1.TextMatrix(I, 0) <> "" And (MSHFlexGrid1.TextMatrix(I, 1) <> "" Or MSHFlexGrid1.TextMatrix(I, 2) <> "" Or MSHFlexGrid1.TextMatrix(I, 3) <> "") Then
      If OptCommodity.Value = True Then
         RsMst_Commodity.MoveFirst
         RsMst_Commodity.Find "Commodity = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
         If RsMst_Commodity.EOF = True Then
            MSHFlexGrid1.row = I
            MSHFlexGrid1.Col = 0
            'MSHFlexGrid1.CellBackColor = vbRed
            GoTo NextRow
         'Else
         '   tmp = " Select CommodityID from Mst_Commodity where NCDEXCODE = '" & MSHFlexGrid1.TextMatrix(I, 1) & "' and CommodityID not in (" & RsMst_Commodity!CommodityID & ")"
         '   Call Tmp4Table
         '   If TmpRs4.RecordCount > 0 Then
         '      MSHFlexGrid1.Row = I
         '      MSHFlexGrid1.Col = 0
         '    '  MSHFlexGrid1.CellBackColor = vbRed
         '      GoTo nextrow
         '   End If
         End If
         RsMst_Commodity!NCDEXCODE = MSHFlexGrid1.TextMatrix(I, 1)
         RsMst_Commodity!CommtrackNCDEXCode = MSHFlexGrid1.TextMatrix(I, 2)
         RsMst_Commodity!NComNCDEXCODE = MSHFlexGrid1.TextMatrix(I, 3)
         RsMst_Commodity.Update
      ElseIf OptGodownCode.Value = True Then
         TmpRs5.MoveFirst
         TmpRs5.Find "GodownCode = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
         If TmpRs5.EOF = True Then
            MSHFlexGrid1.row = I
            MSHFlexGrid1.Col = 0
            'MSHFlexGrid1.CellBackColor = vbRed
            GoTo NextRow
         'Else
         '   tmp = " Select GodownID from Mst_Godown where NCDEXCODE = '" & MSHFlexGrid1.TextMatrix(I, 1) & "' and GodownID not in (" & TmpRs5!GodownID & ")"
         '   Call Tmp4Table
         '   If TmpRs4.RecordCount > 0 Then
         '      MSHFlexGrid1.Row = I
         '      MSHFlexGrid1.Col = 0
         '    '  MSHFlexGrid1.CellBackColor = vbRed
         '      GoTo nextrow
         '   End If
         End If
         RsMst_Godown.MoveFirst
         RsMst_Godown.Find "GodownID = " & TmpRs5!GodownID
         If RsMst_Godown.EOF = True Then
            MSHFlexGrid1.row = I
            MSHFlexGrid1.Col = 0
            'MSHFlexGrid1.CellBackColor = vbRed
            GoTo NextRow
         End If
         RsMst_Godown!NCDEXCODE = MSHFlexGrid1.TextMatrix(I, 1)
         RsMst_Godown!CommtrackNCDEXCode = MSHFlexGrid1.TextMatrix(I, 2)
         RsMst_Godown!NComNCDEXCODE = MSHFlexGrid1.TextMatrix(I, 3)
         RsMst_Godown.Update
      ElseIf OptLocation.Value = True Then
         RsMst_Location.MoveFirst
         RsMst_Location.Find "LocationName = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
         If RsMst_Location.EOF = True Then
            MSHFlexGrid1.row = I
            MSHFlexGrid1.Col = 0
            'MSHFlexGrid1.CellBackColor = vbRed
            GoTo NextRow
         Else
            tmp = " Select LocationID from Mst_Location where (NCDEXCODE = '" & MSHFlexGrid1.TextMatrix(I, 1) & "' Or CommtrackNCDEXCODE = '" & MSHFlexGrid1.TextMatrix(I, 2) & "' Or NComNCDEXCODE = '" & MSHFlexGrid1.TextMatrix(I, 3) & "' ) and LocationID not in (" & RsMst_Location!LocationID & ")"
            Call Tmp4Table
            If TmpRs4.RecordCount > 0 Then
               MSHFlexGrid1.row = I
               MSHFlexGrid1.Col = 0
               'MSHFlexGrid1.CellBackColor = vbRed
               GoTo NextRow
            End If
         End If
         RsMst_Location!NCDEXCODE = MSHFlexGrid1.TextMatrix(I, 1)
         RsMst_Location!CommtrackNCDEXCode = MSHFlexGrid1.TextMatrix(I, 2)
         RsMst_Location!NComNCDEXCODE = MSHFlexGrid1.TextMatrix(I, 3)
         RsMst_Location.Update
      End If
   Else
      MSHFlexGrid1.row = I
      MSHFlexGrid1.Col = 0
      'MSHFlexGrid1.CellBackColor = vbRed
   End If
NextRow:
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
MsgBox "Data Saved Successfully ", vbInformation
ProgressBar1.Value = 0
oXL.Quit
Call Grid_Head
Exit Sub
errortrap:
MsgBox "System Error Occured!!!", vbExclamation
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call Grid_Head
End Sub

Private Sub OptCommodity_Click()
Check2.Visible = False
Call Grid_Head
End Sub

Private Sub OptGodownCode_Click()
Check2.Visible = True
Check2.Value = vbUnchecked
Call Grid_Head
End Sub

Private Sub OptLocation_Click()
Check2.Visible = False
Call Grid_Head
End Sub
