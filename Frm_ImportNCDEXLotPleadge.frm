VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_ImportNCDEXLotPleadge 
   Caption         =   "Import NCDEX Lot Pleadge"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1530
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   14775
      Begin VB.ComboBox CmbBranchName 
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
         Left            =   8985
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   660
         Width           =   5745
      End
      Begin VB.ComboBox CmbBankName 
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
         Left            =   1365
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   660
         Width           =   5505
      End
      Begin VB.OptionButton OptExportPledge 
         Appearance      =   0  'Flat
         Caption         =   "Export For Pledge"
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
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   255
         Width           =   1935
      End
      Begin VB.ComboBox CmbBatch 
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
         Left            =   5310
         TabIndex        =   19
         Top             =   225
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Pledge Creation"
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
         Height          =   375
         Left            =   9795
         TabIndex        =   18
         Top             =   1050
         Width           =   1815
      End
      Begin VB.CommandButton CmdGetData 
         Caption         =   "Get Data"
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
         Left            =   6720
         TabIndex        =   17
         Top             =   1080
         Width           =   1500
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export"
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
         Left            =   8220
         TabIndex        =   16
         Top             =   1080
         Width           =   1500
      End
      Begin VB.OptionButton OptExport 
         Appearance      =   0  'Flat
         Caption         =   "Export"
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
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   255
         Width           =   975
      End
      Begin VB.OptionButton OptImport 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
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
         Left            =   4635
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton CmdValidate 
         Caption         =   "Validate"
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
         Left            =   5880
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.ComboBox CmbSelect 
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
         ItemData        =   "Frm_ImportNCDEXLotPleadge.frx":0000
         Left            =   5310
         List            =   "Frm_ImportNCDEXLotPleadge.frx":000A
         TabIndex        =   10
         Top             =   225
         Width           =   3135
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
         Left            =   135
         TabIndex        =   7
         Top             =   1080
         Width           =   1500
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
         Height          =   350
         Left            =   3135
         TabIndex        =   6
         Top             =   1080
         Width           =   1500
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
         Height          =   350
         Left            =   11670
         TabIndex        =   5
         Top             =   1080
         Width           =   1500
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   13170
         TabIndex        =   4
         Top             =   1080
         Width           =   1500
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
         Left            =   1635
         TabIndex        =   3
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label39 
         Caption         =   "Branch Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label38 
         Caption         =   "Bank Name"
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
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Type"
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
         Left            =   4560
         TabIndex        =   11
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   150
      TabIndex        =   0
      Top             =   1635
      Width           =   14775
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   660
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   150
      TabIndex        =   8
      Top             =   2265
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6975
      Left            =   150
      TabIndex        =   9
      Top             =   2640
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   12303
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_ImportNCDEXLotPleadge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth, mFromDate, mMonthListindex, mSM_Prefix, mTotalDays, mToDate, mNoofDays As Variant
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim ErrMsg As Variant
Dim Status As Variant
Dim WrkSheet As String
Dim RowCount As Double
Dim I As Long
Dim mFlag, mMaxID, mStatus As Variant
Dim iCounter As Double
Dim mPreMsg As Variant
Dim mPostMsg As Variant
Dim mBankID, mBranchID As Variant

Private Sub Check1_Click()
If CmbSelect.Text <> "received from ncdex" Then
   Check1.Value = 0
   Exit Sub
End If
End Sub

Private Sub Check2_Click()
If OptExportPledge.Value = False Then
   Check2.Value = 0
   Exit Sub
End If
End Sub

Private Sub CmbBatch_GotFocus()

If OptExportPledge.Value = True Then
   tmp = "Select distinct BatchID2 As BatchID2 from Txn_ImportNcdexLotPleadge Where Isnull(Pledge,'N') = 'Y' And BatchId2 Not in (Select distinct BatchID2 From Txn_ImportNcdexLotPleadge Where PledgeID is Not Null) order by BatchID2"
Else
   tmp = "Select distinct BatchID1 As BatchID2 from Txn_ImportNcdexLotPleadge Where PledgeID is Null order by BatchID1"
End If
Call Tmp1Table
CmbBatch.Clear
If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
       CmbBatch.AddItem TmpRs1.Fields(0)
       TmpRs1.MoveNext
   Next
End If

End Sub


Private Sub CmbBatch_LostFocus()

If CmbBatch.Text = "" Then Exit Sub
If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
TmpRs1.Find "BatchID2 = " & Val(CmbBatch.Text) & ""
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbBatch.SetFocus
   Exit Sub
End If

If OptExportPledge.Value = True Then
   
   tmp = "Select Distinct MB.BankName,MB.BankID,MBr.BranchName,MBr.BranchID from Txn_ImportNcdexLotPleadge a"
   tmp = tmp & " Inner Join Mst_Branch MBr On a.BranchID = Mbr.BranchID"
   tmp = tmp & " Inner Join Mst_Bank MB on MBr.BankID = MB.BankID"
   tmp = tmp & " Where a.BatchID2 = " & Val(CmbBatch.Text) & " "
    
   Call Tmp2Table
   
   If TmpRs2.RecordCount > 0 Then
      CmbBankName.Text = TmpRs2!BankName
      CmbBranchName.Text = TmpRs2!BranchName
      mBankID = TmpRs2!BankID
      mBranchID = TmpRs2!BranchID
   Else
      CmbBankName.Text = ""
      CmbBranchName.Text = ""
   End If
    
End If


End Sub


Private Sub CmbSelect_LostFocus()
If CmbSelect.Text = "" Then Exit Sub

If LCase(CmbSelect.Text) <> "received from bank" And LCase(CmbSelect.Text) <> "received from ncdex" Then 'And LCase(CmbSelect.Text) <> "pleage creation"
   MsgBox "Invalid Selection"
   CmbSelect.SetFocus
   Exit Sub
End If


If LCase(CmbSelect.Text) = "received from bank" Then
   CmbBankName.Text = ""
   CmbBankName.Enabled = True
   CmbBranchName.Text = ""
   CmbBranchName.Enabled = True
Else
   CmbBankName.Text = ""
   CmbBankName.Enabled = False
   CmbBranchName.Text = ""
   CmbBranchName.Enabled = False
End If

'If LCase(CmbSelect.Text) = "pleage creation" Then
'   CmdGetFile.Enabled = False
'   CmdImport.Enabled = False
'   CmdSave.Enabled = True
'   CmdExcel.Enabled = True
'
'Else
'   CmdGetFile.Enabled = True
'   CmdImport.Enabled = False
'   CmdSave.Enabled = False
'   CmdExcel.Enabled = False
'End If

End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ImportNCDEXLotPleadge.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Export()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportNCDEXLotPleadge.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        If C <> 18 And C <> 19 And C <> 20 And C <> 21 And C <> 22 Then
           If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
           Else
             oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
           MSHFlexGrid1.Col = C
           MSHFlexGrid1.row = I
           If MSHFlexGrid1.CellBackColor = vbRed Then
              oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
           End If
        End If
    Next
Next

oWB.Save
MsgBox ("Excel file created !!! ")
Call Grid_Head
oXL.Visible = True
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ImportNCDEXLotPleadge.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

If LCase(CmbSelect.Text) = "received from ncdex" Then
   For I = 0 To MSHFlexGrid1.Rows - 1
       mRow = I + 1
       For C = 0 To MSHFlexGrid1.Cols - 1
'           If c <> 18 And c <> 19 And c <> 20 And c <> 21 Then
              If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
                 If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
                    oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
                 Else
                    oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
                 End If
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
              MSHFlexGrid1.Col = C
              MSHFlexGrid1.row = I
              If MSHFlexGrid1.CellBackColor = vbRed Then
                 oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
              End If
'           End If
       Next
   Next
ElseIf LCase(CmbSelect.Text) = "received from bank" Then
   For I = 0 To MSHFlexGrid1.Rows - 1
       mRow = I + 1
       For C = 0 To MSHFlexGrid1.Cols - 1
'           If c <> 18 And c <> 19 Then
              If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
                 If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
                    oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
                 Else
                    oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
                 End If
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
              MSHFlexGrid1.Col = C
              MSHFlexGrid1.row = I
              If MSHFlexGrid1.CellBackColor = vbRed Then
                 oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
              End If
'           End If
       Next
   Next
End If


oWB.Save
MsgBox ("Excel file created !!! ")
Call Grid_Head
oXL.Visible = True
End Sub

Private Sub CmdExit_Click()
On Error Resume Next
oXL.Quit
Unload Me
End Sub

Private Sub CmdExport_Click()
If MSHFlexGrid1.Rows <= 2 Then Exit Sub

If OptExportPledge.Value = True Then
   CmdExport.Enabled = False
   If Check2.Value = 1 Then
      tmp = " Select TransNo,GodownId,Fed From Txn_ImportNcdexLotPleadge"
      tmp = tmp & " Where PreviousColumnsDet='Y' And Pledge='Y' And Code IS Null And Flag='A' And BatchID2=" & Val(CmbBatch.Text) & " "
      tmp = tmp & " Group by TransNo,GodownId,Fed"
      tmp = tmp & " order by TransNo"
      Call TmpTable
      If TmpRs.RecordCount > 0 Then
         TmpRs.MoveFirst
         
         For I = 1 To TmpRs.RecordCount
             Call GenerateCode(TmpRs!GodownID, TmpRs!FED, TmpRs!TransNo)
             TmpRs.MoveNext
         Next
          
         tmp = " Select TransNo,NCDEXCode,Fed,SerialNo,[Code],CMSEClientId,SM_Prefix "
         tmp = tmp & " From Txn_ImportNcdexLotPleadge"
         tmp = tmp & " Where PreviousColumnsDet='Y' And Flag='A' And BatchID2=" & Val(CmbBatch.Text) & " And Txn_PledgeCreation_ID IS Null And "
         tmp = tmp & " TransNo not in (Select distinct TransNo from Txn_ImportNcdexLotPleadge where Flag='R') "
         tmp = tmp & " Group by TransNo,NCDEXCode,Fed,SerialNo,Code,CMSEClientId,SM_Prefix"
         tmp = tmp & " Order by Code"
         Call Tmp2Table
         If TmpRs2.RecordCount > 0 Then
            TmpRs2.MoveFirst
            For I = 1 To TmpRs2.RecordCount
                Call InsertTxn_PleadgeCreation(TmpRs2!Code, TmpRs2!SM_Prefix, TmpRs2!CMSEClientID)
                TmpRs2.MoveNext
            Next
         Else
            MsgBox "Pledge ID Not Generated !!!"
            Exit Sub
         End If
      Else
         MsgBox "No record found to pledge creation"
         Exit Sub
      End If
   End If
End If

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportNCDEXLotPleadge.xls")
Call Xls_Detail_Export

Call Grid_Head
CmbBatch.Text = ""


End Sub

Private Sub CmdGetData_Click()

If CmbBatch.Text = "" Then
   MsgBox "Please Enter Batch ID!!!"
   CmbBatch.SetFocus
   Exit Sub
End If
CmdExport.Enabled = False
Call Grid_Head
If OptExportPledge.Value = True Then
   tmp = " Select NCDEXCode,CMSELotId,TransNo,CMSEClientId,CMSEClientName,CMSPLotId,ActualQuantity,UOM,RatePerUOM,"
   tmp = tmp & " ValueCommodity,ValidityDate,Fed,ActualGrade,GradePremium,LotRefNo,PreviousColumnsDet,Pledge,ErrorMsg,SM_Prefix,Mst_GSL_ID,Bags,Quantity,GodownId,Type,Flag"
   tmp = tmp & " From Txn_ImportNcdexLotPleadge Where BatchId2=" & Val(CmbBatch.Text) & ""
Else
   tmp = " Select NCDEXCode,CMSELotId,TransNo,CMSEClientId,CMSEClientName,CMSPLotId,ActualQuantity,UOM,RatePerUOM,"
   tmp = tmp & " ValueCommodity,ValidityDate,Fed,ActualGrade,GradePremium,LotRefNo,PreviousColumnsDet,Pledge,ErrorMsg,SM_Prefix,Mst_GSL_ID,Bags,Quantity,GodownId,Type,Flag"
   tmp = tmp & " From Txn_ImportNcdexLotPleadge Where BatchId1=" & Val(CmbBatch.Text) & " And Isnull(BatchID2,0) = 0"
End If
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   CmdExport.Enabled = True
   mStatus = IIf(IsNull(TmpRs1!Type), "", TmpRs1!Type)
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
           
           If LCase(TmpRs1!Flag) = "r" Or LCase(TmpRs1!PreviousColumnsDet) <> "y" Or LCase(TmpRs1!Pledge) <> "y" Then
              MSHFlexGrid1.Col = C
              MSHFlexGrid1.row = I
              MSHFlexGrid1.CellBackColor = vbRed
           End If
       Next
       MSHFlexGrid1.TextMatrix(I, 15) = IIf(LCase(MSHFlexGrid1.TextMatrix(I, 15)) = "y", "Yes", "No")
       MSHFlexGrid1.TextMatrix(I, 16) = IIf(LCase(MSHFlexGrid1.TextMatrix(I, 16)) = "y", "Yes", "No")
       TmpRs1.MoveNext
   Next
MsgBox "Done!!!"
Else
MsgBox "No record found!!!"
End If

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmbSelect.Text = ""

CmdGetFile.Enabled = True
CmdImport.Enabled = False
CmdSave.Enabled = False
CmdGetData.Enabled = False
CmdExport.Enabled = False
CmdExit.Enabled = True

CmdValidate.Enabled = False
OptImport.Value = True
CmbBatch.Visible = False
CmbSelect.Text = ""
Label2.Caption = "Type"
LblFile.Caption = ""
Check2.Value = 0

CmbBankName.Text = ""
CmbBankName.Enabled = False
CmbBranchName.Text = ""
CmbBranchName.Enabled = False

End Sub
Private Sub CmdGetFile_Click()
On Error GoTo errortrap

If CmbSelect.Text = "" Then
   MsgBox "Blank Type Not Allowed"
   CmbSelect.SetFocus
   Exit Sub
End If

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .InitDir = App.Path
     .Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx"
     .CancelError = True
procReOpen:
     .ShowOpen
End With

Path = CommonDialog1.FileName
LblFile.Caption = Path

CmdImport.Enabled = True
CmdReset.Enabled = False
CmdExit.Enabled = False
CmdSave.Enabled = False

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub UploadFile()
Status = 0
If LblFile.Caption <> "" And LCase(LblFile.Caption) <> "filename" Then
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   CmdExit.Enabled = True
   Status = 0
Else
   MsgBox "Please Select File", vbInformation, "Information"
   Status = 1
End If
End Sub
Private Sub CmdImport_Click()
Dim mulCheck As Integer
Dim isSelected As Boolean

If CmbSelect.Text = "" Then
   MsgBox "Blank Type Not Allowed"
   CmbSelect.SetFocus
   Exit Sub
End If

Call UploadFile
If Status = 1 Then Exit Sub
   CmdGetFile.Enabled = False
   
   '-----------------------------
   ErrMsg = ""
   Status = 0
   RowCount = 0
   I = 1
   '-----------------------------
   Call Validation
   If Status = 1 Then
       MsgBox ErrMsg, vbExclamation, "Message"
       CmdGetFile.Enabled = True
   Else
       Call Grid_Head
       If LCase(CmbSelect.Text) = "received from bank" Then
          Call FillData
       ElseIf LCase(CmbSelect.Text) = "received from ncdex" Then
          Call FillDataNcdex
       End If
       CmdReset.Enabled = True
       CmdValidate.Enabled = True
       Call CmdValidate_Click
   End If
End Sub
Private Sub Validation()

'.TextMatrix(0, 0) = "WAREHOUSE CODE"
'.TextMatrix(0, 1) = "CMSE LOT ID"
'.TextMatrix(0, 2) = "Transaction No."
'.TextMatrix(0, 3) = "CMSE CLIENT ID"
'.TextMatrix(0, 4) = "CMSE CLIENT NAME"
'.TextMatrix(0, 5) = "CMSP LOT ID"
'.TextMatrix(0, 6) = "ACTUAL QUANTITY"
'.TextMatrix(0, 7) = "UOM"
'.TextMatrix(0, 8) = "Rate Per UOM"
'.TextMatrix(0, 9) = "Value of Commodity"
'.TextMatrix(0, 10) = "VALIDITY DATE"
'.TextMatrix(0, 11) = "FED"
'.TextMatrix(0, 12) = "ACTUAL GRADE"
'.TextMatrix(0, 13) = "GRADE Premium/Discount PERCENTAGE"
'.TextMatrix(0, 14) = "LOT REFERENCE NO"
'.TextMatrix(0, 15) = "All the details given in the previous columns are correct (YES/NO)"
'.TextMatrix(0, 16) = "This Lot is Pledged to the client mentioned in CommTrack (Yes / No)"

Set oWS = oWB.Worksheets("Sheet1")
WrkSheet = "Sheet1"
If LCase(Trim(oWS.Cells(1, 1))) <> "warehouse code" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A(WAREHOUSE CODE) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 2))) <> "cmse lot id" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of B(CMSE LOT ID) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 3))) <> "transaction no." Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of C(Transaction no.) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 4))) <> "cmse client id" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of D(CMSE CLIENT ID) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 5))) <> "cmse client name" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of E(CMSE CLIENT NAME) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 6))) <> "cmsp lot id" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of F(CMSP LOT ID) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 7))) <> "actual quantity" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of G(ACTUAL QUANTITY) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 8))) <> "uom" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of H(UOM) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 9))) <> "rate per uom" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of I(Rate Per UOM) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 10))) <> "value of commodity" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of J(Value of Commodity) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 11))) <> "validity date" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of K(Validity Date) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 12))) <> "fed" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of L(FED) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 13))) <> "actual grade" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of M(Actual Grade) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 14))) <> LCase("GRADE Premium/Discount PERCENTAGE") Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of N(Grade Premium/Discount Percentage) " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 15))) <> "lot reference no" Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of O(Lot Reference No) " & vbCrLf
   Status = 1
End If
If LCase(CmbSelect.Text) = "received from ncdex" Then
   If LCase(Trim(oWS.Cells(1, 16))) <> "all the details given in the previous columns are correct (yes/no)" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of P(All the details given in the previous columns are correct (yes/no)) " & vbCrLf
      Status = 1
   End If
   If LCase(Trim(oWS.Cells(1, 17))) <> "this lot is pledged to the client mentioned in commtrack (yes / no)" Then
      ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of Q(This lot is pledged to the client mentioned in commTrack (yes/no)) " & vbCrLf
      Status = 1
   End If
End If
End Sub
Private Sub FillData()
Dim intNumberRows As Double

Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
RowCount = oWB.Worksheets(1).UsedRange.Rows.Count
ProgressBar1.Value = 0
ProgressBar1.Max = RowCount

Do
   If oWS.Cells(mRow, 1) = "" Then
      Exit Do
   End If
  
   For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(mRow, C) = Trim(oWS.Cells(mRow + 1, C + 1).Value)
   Next
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   ProgressBar1.Value = ProgressBar1.Value + 1
   mRow = mRow + 1
Loop Until False

MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1

ProgressBar1.Value = ProgressBar1.Max
oXL.Quit
MsgBox "Data Imported Successfully ", vbInformation
CmdImport.Enabled = False
CmdValidate.Enabled = True
ProgressBar1.Value = 0
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.FixedRows = 1

With MSHFlexGrid1
    .Rows = 2
    .Cols = 23
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(0, 0) = "WAREHOUSE CODE"
    .TextMatrix(0, 1) = "CMSE LOT ID"
    .TextMatrix(0, 2) = "Transaction No."
    .TextMatrix(0, 3) = "CMSE CLIENT ID"
    .TextMatrix(0, 4) = "CMSE CLIENT NAME"
    .TextMatrix(0, 5) = "CMSP LOT ID"
    .TextMatrix(0, 6) = "ACTUAL QUANTITY"
    .TextMatrix(0, 7) = "UOM"
    .TextMatrix(0, 8) = "Rate Per UOM"
    .TextMatrix(0, 9) = "Value of Commodity"
    .TextMatrix(0, 10) = "VALIDITY DATE"
    .TextMatrix(0, 11) = "FED"
    .TextMatrix(0, 12) = "ACTUAL GRADE"
    .TextMatrix(0, 13) = "GRADE Premium/Discount PERCENTAGE"
    .TextMatrix(0, 14) = "LOT REFERENCE NO"
    .TextMatrix(0, 15) = "All the details given in the previous columns are correct (YES/NO)"
    .TextMatrix(0, 16) = "This Lot is Pledged to the client mentioned in CommTrack (Yes / No)"
    .TextMatrix(0, 17) = "Error Msg"
    .TextMatrix(0, 18) = "SM_Prefix"
    .TextMatrix(0, 19) = "GSLID"
    .TextMatrix(0, 20) = "Balancebags"
    .TextMatrix(0, 21) = "BalanceQuantity"
    .TextMatrix(0, 22) = "GodownID"
    
'    .TextMatrix(0, 0) = "Sr.No."
'    .TextMatrix(0, 1) = "Warehouse Code"
'    .TextMatrix(0, 2) = "CMSE Lot Id"
'    .TextMatrix(0, 3) = "Approved in Comtrack"
'    .TextMatrix(0, 4) = "Transaction no."
'    .TextMatrix(0, 5) = "CMSE Client Id"
'    .TextMatrix(0, 6) = "CMSE Client Name"
'    .TextMatrix(0, 7) = "CMG Confirmation(Current Beneficiary)"
'    .TextMatrix(0, 8) = "CMSP Lot id"
'    .TextMatrix(0, 9) = "Actual Quantity"
'    .TextMatrix(0, 10) = "QUANTITY CONFIRMATION(YES / NO)" 'n
'    .TextMatrix(0, 11) = "UOM"
'    .TextMatrix(0, 12) = "FED"
'    .TextMatrix(0, 13) = "SECURITY 24x7 (YES/NO)" 'n
'    .TextMatrix(0, 14) = "FFE (YES/NO)" 'n
'    .TextMatrix(0, 15) = "Insurace taken by" 'n
'    .TextMatrix(0, 16) = "Insurance policy details" 'n
'    .TextMatrix(0, 17) = "DUNNAGE (YES/NO)" 'n
'    .TextMatrix(0, 18) = "RECORDS MAINTAINED (YES/NO)" 'n
'    .TextMatrix(0, 19) = "PLEDGE CONFIRMATION (ON PHYSICAL LOTS) (YES/NO)" 'n
'    .TextMatrix(0, 20) = "Godown situated in factory premises (Yes/NO)" 'n
'    .TextMatrix(0, 21) = "Factory is in working condition (Yes/NO)" 'n
'    .TextMatrix(0, 22) = "Type of warehouse (Cold Storage/Dry Warehouse/Open Plinth)" 'n
'    .TextMatrix(0, 23) = "Lock & Key with WSP (Yes/No)" 'n
'    .TextMatrix(0, 24) = "Remarks" 'n
'    .TextMatrix(0, 25) = "Error Message" 'n
'    .TextMatrix(0, 26) = "Ncdex Confirmation" 'n

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
    .ColWidth(11) = 1500
    .ColWidth(12) = 1500
    .ColWidth(13) = 1500
    .ColWidth(14) = 1500
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 1500
    .ColWidth(18) = 1500
    .ColWidth(19) = 1500
    .ColWidth(20) = 1500
    .ColWidth(21) = 1500
    
    If LCase(CmbSelect.Text) = "received from bank" Then
       .ColWidth(15) = 0
       .ColWidth(16) = 0
    ElseIf LCase(CmbSelect.Text) = "received from ncdex" Then
       .ColWidth(15) = 1000
       .ColWidth(16) = 1000
    End If
    
    If OptExport.Value = True Then
       .ColWidth(15) = 1000
       .ColWidth(16) = 1000
    End If
    
    .RowHeight(0) = 800
    .WordWrap = True
End With
End Sub


Private Sub CmdSave_Click()

'On Error GoTo errortrap
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows - 1

If CmbSelect.Text = "" Then
   MsgBox "Blank Type Not Allowed"
   CmbSelect.SetFocus
   Exit Sub
End If
   
If LCase(CmbSelect.Text) = "received from bank" Then
   If CmbBankName.Text = "" Then
      MsgBox "Blank Bank Not allowed!!!"
      CmbBankName.SetFocus
      Exit Sub
   End If
   If CmbBranchName.Text = "" Then
      MsgBox "Blank Branch Not allowed!!!"
      CmbBranchName.SetFocus
      Exit Sub
   End If
End If
   
tmp = " Select * From Txn_ImportNcdexLotPleadge "
Call Tmp1Table

If LCase(CmbSelect.Text) = "received from bank" Then
   mFlag = "B"
   
ElseIf LCase(CmbSelect.Text) = "received from ncdex" Then
   mFlag = "N"
End If
    
If LCase(CmbSelect.Text) = "received from bank" Then
   mMaxID = GetMaxBatchID
   For I = 1 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(I, 1) <> "" And MSHFlexGrid1.TextMatrix(I, 2) <> "" Then
          If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
          mkey1 = MSHFlexGrid1.TextMatrix(I, 1) & "~" & MSHFlexGrid1.TextMatrix(I, 2)
          TmpRs5.Find "Key1='" & mkey1 & "'"
          If TmpRs5.EOF = True Then
             TmpRs1.AddNew
             TmpRs1!SrNo = GetMaxID("SrNo", "Txn_ImportNcdexLotPleadge", "")
             TmpRs1!NCDEXCODE = MSHFlexGrid1.TextMatrix(I, 0)
             TmpRs1!CMSELotId = MSHFlexGrid1.TextMatrix(I, 1)
             TmpRs1!TransNo = MSHFlexGrid1.TextMatrix(I, 2)
             TmpRs1!CMSEClientID = MSHFlexGrid1.TextMatrix(I, 3)
             TmpRs1!CMSEClientName = MSHFlexGrid1.TextMatrix(I, 4)
             TmpRs1!CMSPLotId = MSHFlexGrid1.TextMatrix(I, 5)
             TmpRs1!ActualQuantity = MSHFlexGrid1.TextMatrix(I, 6)
             TmpRs1!UOM = MSHFlexGrid1.TextMatrix(I, 7)
             TmpRs1!RatePerUOM = Val(MSHFlexGrid1.TextMatrix(I, 8))
             TmpRs1!ValueCommodity = Val(MSHFlexGrid1.TextMatrix(I, 9))
             TmpRs1!ValidityDate = MSHFlexGrid1.TextMatrix(I, 10)
             TmpRs1!FED = MSHFlexGrid1.TextMatrix(I, 11)
             TmpRs1!ACTUALGRADE = MSHFlexGrid1.TextMatrix(I, 12)
             TmpRs1!GradePremium = MSHFlexGrid1.TextMatrix(I, 13)
             TmpRs1!LotRefNo = MSHFlexGrid1.TextMatrix(I, 14)
             TmpRs1!Remarks = Null
             TmpRs1!Type = mFlag
             TmpRs1!BatchId1 = mMaxID
             TmpRs1!BatchID2 = Null
             TmpRs1!Flag = IIf(MSHFlexGrid1.TextMatrix(I, 17) <> "", "R", "A")
             TmpRs1!ErrorMsg = MSHFlexGrid1.TextMatrix(I, 17)
             TmpRs1!SM_Prefix = MSHFlexGrid1.TextMatrix(I, 18)
             TmpRs1!Mst_GSL_ID = Val(MSHFlexGrid1.TextMatrix(I, 19))
             TmpRs1!Bags = Val(MSHFlexGrid1.TextMatrix(I, 20))
             TmpRs1!Quantity = Val(MSHFlexGrid1.TextMatrix(I, 21))
             TmpRs1!CreatedBy = Gen_UserLoginID
             TmpRs1!CreatedDate = Now
             TmpRs1!ImportDate = Date
             TmpRs1!GodownID = Val(MSHFlexGrid1.TextMatrix(I, 22))
             TmpRs1!BranchID = Val(mBranchID)
             TmpRs1.Update
          End If
       End If
   Next
   MsgBox "Data from bank imported successfuly, Batch Id is " & mMaxID & "!!!"
   Call CmdExcel_Click
ElseIf LCase(CmbSelect.Text) = "received from ncdex" Then
    tmp = " Select Isnull(Max(Convert(integer,BatchId2))+1,1) AS BatchID From Txn_ImportNcdexLotPleadge"
    Call Tmp6Table
    If TmpRs6.RecordCount > 0 Then
       mMaxID = TmpRs6!BatchID
    End If
    For I = 1 To MSHFlexGrid1.Rows - 1
        If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
        mkey1 = MSHFlexGrid1.TextMatrix(I, 1) & "~" & MSHFlexGrid1.TextMatrix(I, 2)
        TmpRs5.Find "Key1='" & mkey1 & "'"
        If TmpRs5.EOF = False Then
'           If Check1.Value = 1 Then
              TmpRs5!BatchID2 = mMaxID
'           End If
           TmpRs5!PreviousColumnsDet = UCase(Left(MSHFlexGrid1.TextMatrix(I, 15), 1))
           TmpRs5!Pledge = UCase(Left(MSHFlexGrid1.TextMatrix(I, 16), 1))
           TmpRs5!Type = mFlag
           TmpRs1!UpdatedBy = Gen_UserLoginID
           TmpRs1!UpdatedDate = Now
           TmpRs5!ErrorMsg = MSHFlexGrid1.TextMatrix(I, 17)
           If MSHFlexGrid1.TextMatrix(I, 17) <> "" Then
              TmpRs5!Flag = "R"
           End If
           TmpRs5.Update
        End If
    Next
    MsgBox "Data from bank imported successfuly, Batch Id is " & mMaxID & "!!!"
    Call CmdExcel_Click
End If

'MsgBox "Data Saved Successfully ", vbInformation
CmdSave.Enabled = False
CmdExcel.Enabled = True
CmdImport.Enabled = False
CmdGetFile.Enabled = True
ProgressBar1.Value = 0
Exit Sub
'errortrap:
'MsgBox "System Error Occured!!!", vbExclamation
End Sub

Private Sub CmdValidate_Click()
Dim mSM_Prefix As Variant
iCounter = 0

mSM_Prefix = ""

If CmbSelect.Text = "" Then
   MsgBox "Blank Type Not Allowed"
   CmbSelect.SetFocus
   Exit Sub
End If

CmdSave.Enabled = True
MsgBox "Please Wait Till Validate Data !!!"

tmp = " Select ISINID,FED from Mst_ISIN"
Call Tmp1Table

tmp = " Select ISINID,RRNID from Mst_RRN_lotDetail"
Call Tmp2Table

tmp = " Select ISINID,SM_Prefix from Mst_ISIN MG"
tmp = tmp & " inner join Mst_Location ML on MG.LocationID=ML.LocationID"
Call Tmp3Table

tmp = "Select SM_Prefix+'~'+ISINID as Key1,Pledgebags,PledgeWeight,BalanceWeight,Mst_GSL_ID,BalanceBags,GodownID "
tmp = tmp & " from Mst_GSL Where ExchangeTypeID = 2 And Status Not in ('E','Z')"
Call Tmp4Table

tmp = " Select CMSELotId+'~'+TransNo As Key1,* From Txn_ImportNcdexLotPleadge "
Call Tmp5Table

tmp = " Select ClientIDRow,ClientID,CMSPCode,Status,ExchangeTypeID from Mst_client Where Status='A' and ExchangeTypeID=2"
Call Tmp6Table

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
iCounter = 0

If LCase(CmbSelect.Text) = "received from ncdex" Then
   For j = 1 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(j, 2) <> "" Then
          iCounter = 0
          For k = 1 To MSHFlexGrid1.Rows - 1
              If MSHFlexGrid1.TextMatrix(j, 2) = MSHFlexGrid1.TextMatrix(k, 2) Then
                 iCounter = iCounter + 1
              End If
          Next
          ProgressBar1.Value = ProgressBar1.Value + 1
          tmp = " Select Isnull(count(TransNo),0) AS CountTransNo from Txn_ImportNcdexLotPleadge where TransNo='" & MSHFlexGrid1.TextMatrix(j, 2) & "'"
          Call TmpTable
          If Not IsNull(TmpRs!CountTransNo) Then
             If Val(iCounter) <> Val(TmpRs!CountTransNo) Then
                ProgressBar1.Value = ProgressBar1.Max
                MsgBox "Count of lot mismatch!!!"
                For m = 1 To MSHFlexGrid1.Rows - 1
                    If MSHFlexGrid1.TextMatrix(m, 0) <> "" Then
                       For C = 0 To MSHFlexGrid1.Cols - 1
                           MSHFlexGrid1.Col = C
                           MSHFlexGrid1.row = m
                           MSHFlexGrid1.CellBackColor = vbRed
                       Next
                    End If
                Next
                Exit Sub
             End If
          End If
      End If
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2

For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
       If MSHFlexGrid1.TextMatrix(I, 0) = "" And MSHFlexGrid1.TextMatrix(I, 1) = "" Then
          MSHFlexGrid1.TextMatrix(I, 17) = "Blank Not Allowed"
          GoTo mNextRow
       End If

       If LCase(CmbSelect.Text) = "received from bank" Then
          If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
          TmpRs5.Find "TransNo = '" & MSHFlexGrid1.TextMatrix(I, 2) & "'"
          If TmpRs5.EOF = False Then
             MSHFlexGrid1.TextMatrix(I, 17) = "Already exist transaction no"
             CmdSave.Enabled = False
             GoTo mNextRow
          End If
       Else
          If LCase(MSHFlexGrid1.TextMatrix(I, 15)) <> "yes" Then
             MSHFlexGrid1.TextMatrix(I, 17) = "Ncdex Confirmation (Yes/No) column is blank or No"
             GoTo mNextRow
          End If
          
          If TmpRs5.RecordCount > 0 Then
             If CheckDataChangeOrNot(I) = False Then
                MSHFlexGrid1.TextMatrix(I, 17) = "Previous data and current data is mismatch"
                GoTo mNextRow
             End If
          Else
                MSHFlexGrid1.TextMatrix(I, 17) = "Data not received from bank"
                GoTo mNextRow
          End If
       End If
       
       If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
       TmpRs2.Find "ISINID='" & MSHFlexGrid1.TextMatrix(I, 1) & "'"
       If TmpRs2.EOF = False Then
          MSHFlexGrid1.TextMatrix(I, 17) = "ISIN already mapped to RRN ID " & TmpRs2!RRNID & ""
          GoTo mNextRow
       End If
       
       
       If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
       TmpRs1.Find "ISINID='" & MSHFlexGrid1.TextMatrix(I, 1) & "'"
       If TmpRs1.EOF = True Then
          MSHFlexGrid1.TextMatrix(I, 17) = "ISIN ID not exist in master"
          GoTo mNextRow
       Else
          If CDate(TmpRs1!FED) < CDate(Date) Then
             MSHFlexGrid1.TextMatrix(I, 17) = "FED date should be greater than or equal to current date"
             GoTo mNextRow
          End If
       End If
       
       If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
       TmpRs3.Find "ISINID='" & MSHFlexGrid1.TextMatrix(I, 1) & "'"
       If TmpRs3.EOF = False Then
          mSM_Prefix = TmpRs3!SM_Prefix
          MSHFlexGrid1.TextMatrix(I, 18) = mSM_Prefix
       Else
          mSM_Prefix = ""
          MSHFlexGrid1.TextMatrix(I, 17) = "Warehouse code not exist in master"
          GoTo mNextRow
       End If
        
       If TmpRs6.RecordCount > 0 Then TmpRs6.MoveFirst
       TmpRs6.Find "ClientIDRow='" & MSHFlexGrid1.TextMatrix(I, 3) & "'"
       If TmpRs6.EOF = True Then
          MSHFlexGrid1.TextMatrix(I, 17) = "Client id not exist in master"
          GoTo mNextRow
'       ElseIf LCase(TmpRs6!Status) <> "a" Then
'          MSHFlexGrid1.TextMatrix(i, 17) = "Client is de-active"
'          GoTo mNextRow
'       ElseIf Val(TmpRs6!ExchangeTypeID) <> 2 Then
'          MSHFlexGrid1.TextMatrix(i, 17) = "Client is not NCDEX"
'          GoTo mNextRow
       End If
       
       If TmpRs4.RecordCount > 0 Then TmpRs4.MoveFirst
       mkey1 = mSM_Prefix & "~" & Trim(MSHFlexGrid1.TextMatrix(I, 1))
       TmpRs4.Find "Key1 ='" & mkey1 & "'"
       If TmpRs4.EOF = True Then
          MSHFlexGrid1.TextMatrix(I, 17) = "ISIN Id not exist in GSL"
          GoTo mNextRow
       ElseIf TmpRs4!PledgeBags > 0 And TmpRs4!PledgeWeight > 0 Then
          MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs4!Mst_GSL_ID), 0, TmpRs4!Mst_GSL_ID)
          MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs4!BalanceBags), "", TmpRs4!BalanceBags)
          MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs4!BalanceWeight), "", TmpRs4!BalanceWeight)
          MSHFlexGrid1.TextMatrix(I, 22) = IIf(IsNull(TmpRs4!GodownID), "", TmpRs4!GodownID)
          MSHFlexGrid1.TextMatrix(I, 17) = "Lot already pledged!!!"
          GoTo mNextRow
       Else
          MSHFlexGrid1.TextMatrix(I, 19) = TmpRs4!Mst_GSL_ID
          MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs4!BalanceBags), "", TmpRs4!BalanceBags)
          MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs4!BalanceWeight), "", TmpRs4!BalanceWeight)
          MSHFlexGrid1.TextMatrix(I, 22) = IIf(IsNull(TmpRs4!GodownID), "", TmpRs4!GodownID)
          
          If Val(MSHFlexGrid1.TextMatrix(I, 6)) > Val(TmpRs4!BalanceWeight) Then
             MSHFlexGrid1.TextMatrix(I, 17) = "Actual quantity is greater than balance quantity (GSL Master)"
             GoTo mNextRow
          End If
          'Val(Val(TmpRs4!BalanceWeight) - (Val(TmpRs4!BalanceWeight) * 5 / 100))
'          mMax = Val(Val(TmpRs4!BalanceWeight) + (Val(TmpRs4!BalanceWeight) * 5 / 100)) 'And Not Val(MSHFlexGrid1.TextMatrix(i, 6)) < mMax
          mMin = Val(Val(TmpRs4!BalanceWeight) - (Val(TmpRs4!BalanceWeight) * 5 / 100))
          If Val(MSHFlexGrid1.TextMatrix(I, 6)) < mMin Then
             MSHFlexGrid1.TextMatrix(I, 17) = "Actual quantity is less than 5 % of balance quantity"
             GoTo mNextRow
          End If
       End If
    End If
    
    If MSHFlexGrid1.TextMatrix(I, 11) <> "" Then
       If CDate(MSHFlexGrid1.TextMatrix(I, 11)) < Date Then
          MSHFlexGrid1.TextMatrix(I, 17) = "FED date should be greater than or equal to current date"
          GoTo mNextRow
       End If
       If CDate(MSHFlexGrid1.TextMatrix(I, 10)) < Date Then
          MSHFlexGrid1.TextMatrix(I, 17) = "Validity date should be greater than or equal to current date"
          GoTo mNextRow
       End If
       If CDate(MSHFlexGrid1.TextMatrix(I, 10)) < CDate(MSHFlexGrid1.TextMatrix(I, 11)) Then
          MSHFlexGrid1.TextMatrix(I, 17) = "Validity date should be greater than or equal to RRN date"
          GoTo mNextRow
       End If
    End If
    
mNextRow:

If MSHFlexGrid1.TextMatrix(I, 17) <> "" And MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
'   MSHFlexGrid1.TextMatrix(i, 15) = "no"
'   MSHFlexGrid1.TextMatrix(i, 16) = "no"
   For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.Col = C
       MSHFlexGrid1.row = I
       MSHFlexGrid1.CellBackColor = vbRed
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
End If


Next
      
ProgressBar1.Value = ProgressBar1.Max
CmdExcel.Enabled = True
MsgBox "Done !!!"

End Sub

Private Sub Form_Load()
'LblFile.Caption = ""
'Frame1.Enabled = False
'Frame2.Enabled = False
'Frame3.Enabled = False
'CmdSave.Enabled = False
'
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True

Call CmdReset_Click

End Sub

Private Sub CmbMonth_GotFocus()

tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I), I - 1
Next

CmbMonth.Text = tmp

End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
       CmbMonth.Text = CmbMonth.List(I)
       mMonthListindex = I '+ 1
       mStatus = True
       Exit For
    End If
Next

If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For I = 2010 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()

If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
flg = "N"
For I = 2010 To Year(Now)
    If CmbYear.Text = I Then
       flg = "Y"
       Exit For
    End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbYear.SetFocus
   Exit Sub
End If

End Sub



Public Function CheckDataChangeOrNot(iRow As Long) As Boolean
Dim Retval As Boolean
Retval = True

If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
mkey1 = MSHFlexGrid1.TextMatrix(iRow, 1) & "~" & MSHFlexGrid1.TextMatrix(iRow, 2)
TmpRs5.Find "Key1='" & mkey1 & "'"
If TmpRs5.EOF = False Then
   If TmpRs5!NCDEXCODE <> MSHFlexGrid1.TextMatrix(iRow, 0) Then
      Retval = False
   End If
   If TmpRs5!CMSEClientID <> MSHFlexGrid1.TextMatrix(iRow, 3) Then
      Retval = False
   End If
   If TmpRs5!CMSEClientName <> MSHFlexGrid1.TextMatrix(iRow, 4) Then
      Retval = False
   End If
   If TmpRs5!CMSPLotId <> MSHFlexGrid1.TextMatrix(iRow, 5) Then
      Retval = False
   End If
   If TmpRs5!ActualQuantity <> MSHFlexGrid1.TextMatrix(iRow, 6) Then
      Retval = False
   End If
   If TmpRs5!UOM <> MSHFlexGrid1.TextMatrix(iRow, 7) Then
      Retval = False
   End If
   If TmpRs5!RatePerUOM <> MSHFlexGrid1.TextMatrix(iRow, 8) Then
      Retval = False
   End If
   If TmpRs5!ValueCommodity <> MSHFlexGrid1.TextMatrix(iRow, 9) Then
      Retval = False
   End If
   If CDate(TmpRs5!ValidityDate) <> CDate(MSHFlexGrid1.TextMatrix(iRow, 10)) Then
      Retval = False
   End If
   If CDate(TmpRs5!FED) <> CDate(MSHFlexGrid1.TextMatrix(iRow, 11)) Then
      Retval = False
   End If
   If TmpRs5!ACTUALGRADE <> MSHFlexGrid1.TextMatrix(iRow, 12) Then
      Retval = False
   End If
   If TmpRs5!GradePremium <> MSHFlexGrid1.TextMatrix(iRow, 13) Then
      Retval = False
   End If
   If TmpRs5!LotRefNo <> MSHFlexGrid1.TextMatrix(iRow, 14) Then
      Retval = False
   End If
End If
CheckDataChangeOrNot = Retval
End Function

Public Sub FillDataNcdex()
Dim intNumberRows As Double

Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
RowCount = oWB.Worksheets(1).UsedRange.Rows.Count
ProgressBar1.Value = 0
ProgressBar1.Max = RowCount

Do
   If oWS.Cells(mRow, 1) = "" Then
      Exit Do
   End If
  
   For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(mRow, C) = Trim(oWS.Cells(mRow + 1, C + 1).Value)
   Next
   If MSHFlexGrid1.TextMatrix(mRow, 17) <> "" And MSHFlexGrid1.TextMatrix(mRow, 0) <> "" Then
      For C = 0 To MSHFlexGrid1.Cols - 1
          MSHFlexGrid1.Col = C
          MSHFlexGrid1.row = mRow
          MSHFlexGrid1.CellBackColor = vbRed
      Next
   End If
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   ProgressBar1.Value = ProgressBar1.Value + 1
   mRow = mRow + 1
Loop Until False
MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1

ProgressBar1.Value = ProgressBar1.Max
oXL.Quit
MsgBox "Data Imported Successfully ", vbInformation
CmdImport.Enabled = False
CmdValidate.Enabled = True
ProgressBar1.Value = 0

End Sub




Private Sub OptExport_Click()
If OptImport.Value = True Then
   CmbBatch.Visible = False
   CmbSelect.Text = ""
   CmbSelect.Visible = True
   Label2.Caption = "Type"
   CmdGetFile.Enabled = True
   CmdImport.Enabled = False
   CmdSave.Enabled = False
   CmdGetData.Enabled = False
   CmdExport.Enabled = False
   CmdExit.Enabled = True
   CmdReset.Enabled = True
Else
   CmbBatch.Visible = True
   CmbSelect.Text = ""
   CmbSelect.Visible = False
   Label2.Caption = "Batch ID"
   CmdGetData.Enabled = True
   CmdExport.Enabled = False
   
   CmdGetFile.Enabled = False
   CmdImport.Enabled = False
   CmdSave.Enabled = False
   CmdExit.Enabled = True
   CmdReset.Enabled = True
   CmdExcel.Enabled = False
   Call Grid_Head
   CmbBankName.Text = ""
   CmbBankName.Enabled = False
   CmbBranchName.Text = ""
   CmbBranchName.Enabled = False
End If
End Sub

Private Sub OptExportPledge_Click()
If OptExportPledge.Value = False Then
   CmbBatch.Visible = False
   CmbSelect.Text = ""
   CmbSelect.Visible = True
   Label2.Caption = "Type"
   CmdGetFile.Enabled = True
   CmdImport.Enabled = False
   CmdSave.Enabled = False
   CmdGetData.Enabled = False
   CmdExport.Enabled = False
   CmdExit.Enabled = True
   CmdReset.Enabled = True
   Call Grid_Head
Else
   CmbBatch.Visible = True
   CmbSelect.Text = ""
   CmbSelect.Visible = False
   Label2.Caption = "Batch ID"
   CmdGetData.Enabled = True
   CmdExport.Enabled = False
   
   CmdGetFile.Enabled = False
   CmdImport.Enabled = False
   CmdSave.Enabled = False
   CmdExit.Enabled = True
   CmdReset.Enabled = True
   CmdExcel.Enabled = False
   Call Grid_Head
   CmbBankName.Text = ""
   CmbBankName.Enabled = False
   CmbBranchName.Text = ""
   CmbBranchName.Enabled = False

End If
End Sub


Private Sub OptImport_Click()
If OptImport.Value = True Then
   CmbBatch.Visible = False
   CmbSelect.Text = ""
   CmbSelect.Visible = True
   Label2.Caption = "Type"
   
   CmdGetFile.Enabled = True
   CmdImport.Enabled = False
   CmdSave.Enabled = False
   CmdExport.Enabled = False
   CmdGetData.Enabled = False
   CmdExcel.Enabled = False
   CmdExit.Enabled = True
   Call Grid_Head
   If LCase(CmbSelect.Text) = "received from bank" Then
      CmbBankName.Text = ""
      CmbBankName.Enabled = True
      CmbBranchName.Text = ""
      CmbBranchName.Enabled = True
   Else
      CmbBankName.Text = ""
      CmbBankName.Enabled = False
      CmbBranchName.Text = ""
      CmbBranchName.Enabled = False
   End If
   
Else
   CmbBatch.Visible = True
   CmbSelect.Text = ""
   CmbSelect.Visible = False
   Label2.Caption = "Batch ID"
   
   CmdGetFile.Enabled = False
   CmdImport.Enabled = False
   CmdSave.Enabled = False
   CmdExport.Enabled = True
   CmdGetData.Enabled = True
   CmdExit.Enabled = True
   Call Grid_Head
   CmbBankName.Text = ""
   CmbBankName.Enabled = False
   CmbBranchName.Text = ""
   CmbBranchName.Enabled = False
End If

End Sub



Public Function GetMaxBatchID() As Variant
tmp = " Select Isnull(Max(Convert(integer,BatchId1))+1,1) AS BatchID From Txn_ImportNcdexLotPleadge"
Call Tmp6Table
If TmpRs6.RecordCount > 0 Then
   GetMaxBatchID = TmpRs6!BatchID
End If
End Function

Public Sub GenerateCode(mCode As Variant, mFed As Variant, mTrans As Variant)

tmp = " Select * From Txn_ImportNcdexLotPleadge Where SerialNo Is null And GodownID='" & mCode & "' And TransNo='" & mTrans & "' And Fed='" & mFed & "' And PreviousColumnsDet='Y' And Flag='A'"
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then

   tmp = " Select Isnull(Max(SerialNo),0)+ 1 As MaxID"
   tmp = tmp & " From Txn_ImportNcdexLotPleadge Where TransNo='" & mTrans & "' And PreviousColumnsDet='Y' And Flag='A'"
   Call Tmp3Table
   If TmpRs3.RecordCount > 0 Then
      tmp = " Update Txn_ImportNcdexLotPleadge Set Type='P', SerialNo=" & TmpRs3!MaxID & ",Code=TransNo+'/'+'" & TmpRs3!MaxID & "'"
      tmp = tmp & " Where PreviousColumnsDet='Y' And Flag='A' And SerialNo Is null And GodownID='" & mCode & "' And Fed='" & mFed & "' And TransNo='" & mTrans & "'"
      Call Tmp4Table
   End If
   
End If
End Sub

Public Sub InsertTxn_PleadgeCreation(Code_ As Variant, mSM_Prefix_ As Variant, mCMSEClientID_ As Variant)  '(Rs As Recordset)
Dim mPledgeID As Variant

tmp = " Select SM_Prefix+'~'+ cast(Mst_GSL_ID As varchar) As mKey1,PledgeBags,PledgeWeight,CMPID,PledgeNo from Mst_GSL where SM_Prefix+'~'+ cast(Mst_GSL_ID  As varchar) in (Select SM_Prefix+'~'+ cast(Mst_GSL_ID  As varchar) From Txn_ImportNcdexLotPleadge)"
Call Tmp5Table

Call TableTxn_PledgeCreation(" Where ISINID='" & Code_ & "'")
If RsTxn_PledgeCreation.EOF = True Then
   RsTxn_PledgeCreation.AddNew
   RsTxn_PledgeCreation!PledgeId = GetMaxID("PledgeID", "Txn_PledgeCreation", " Where SM_Prefix = '" & mSM_Prefix_ & "'")
   mPledgeID = RsTxn_PledgeCreation!PledgeId
   RsTxn_PledgeCreation!PledgeDate = Date
   RsTxn_PledgeCreation!SM_Prefix = mSM_Prefix_ 'Rs!SM_Prefix
   RsTxn_PledgeCreation!BankID = mBankID ' 58
   RsTxn_PledgeCreation!BranchID = mBranchID '348
   RsTxn_PledgeCreation!PledgeReqNo = mPledgeID
   RsTxn_PledgeCreation!PledgeFees = 500
   RsTxn_PledgeCreation!PledgeReqDate = Date
   RsTxn_PledgeCreation!Remarks = "Import LOT NCDEX Pleadge"
   RsTxn_PledgeCreation!ISINID = Code_
   RsTxn_PledgeCreation!ExchangeTypeID = 2
   RsTxn_PledgeCreation!ClientID = GetClientIDRowFromClientID(mCMSEClientID_) '(Rs!CMSEClientID)
   RsTxn_PledgeCreation!G_UID = GetGUID
   mGUIDPC = RsTxn_PledgeCreation!G_UID
   RsTxn_PledgeCreation!CreatedBy = Gen_UserLoginID
   RsTxn_PledgeCreation!CreatedDate = Now
   RsTxn_PledgeCreation!ActualFees = 500
   RsTxn_PledgeCreation!Discount = 0
   RsTxn_PledgeCreation!RecFees = 0
   RsTxn_PledgeCreation!SCMFlag = "0"
   RsTxn_PledgeCreation.Update
   
   
   
    tmp = "Select * from  Txn_PledgeCreation Where G_UID = '" & mGUIDPC & "'"
    Call Tmp1Table
    
    If TmpRs1.RecordCount > 0 Then
       mPledgeCreation_ID = TmpRs1!Txn_PledgeCreation_ID
    End If
    
   Call TableTxn_PledgeCreation_GSL(" Where Txn_PledgeCreation_ID=" & mPledgeCreation_ID & " And SM_Prefix = '" & mSM_Prefix_ & "' ")
   If RsTxn_PledgeCreation_GSL.RecordCount = 0 Then
      Call TableTxn_ImportNcdexLotPleadge("Where Code='" & Code_ & "'")
      If RsTxn_ImportNcdexLotPleadge.RecordCount > 0 Then
         For j = 1 To RsTxn_ImportNcdexLotPleadge.RecordCount
             RsTxn_PledgeCreation_GSL.AddNew
             RsTxn_PledgeCreation_GSL!PledgeNo = GetMaxID("PledgeNo", "Txn_PledgeCreation_GSL", " Where SM_Prefix = '" & mSM_Prefix_ & "'")
             RsTxn_PledgeCreation_GSL!SM_Prefix = RsTxn_ImportNcdexLotPleadge!SM_Prefix
             RsTxn_PledgeCreation_GSL!PledgeId = mPledgeID
             RsTxn_PledgeCreation_GSL!Txn_PledgeCreation_ID = mPledgeCreation_ID
             RsTxn_PledgeCreation_GSL!Mst_GSL_ID = RsTxn_ImportNcdexLotPleadge!Mst_GSL_ID
             RsTxn_PledgeCreation_GSL!PledgeNoOfBags = RsTxn_ImportNcdexLotPleadge!Bags
             RsTxn_PledgeCreation_GSL!PledgeQuantity = RsTxn_ImportNcdexLotPleadge!ActualQuantity
             RsTxn_PledgeCreation_GSL!Remarks = "Import LOT NCDEX Pleadge"
             RsTxn_PledgeCreation_GSL!G_UID = GetGUID
             RsTxn_PledgeCreation_GSL.Update
             
             RsTxn_ImportNcdexLotPleadge!Txn_PledgeCreation_ID = mPledgeCreation_ID
             RsTxn_ImportNcdexLotPleadge.Update
             
             If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
             mkey1 = RsTxn_ImportNcdexLotPleadge!SM_Prefix & "~" & RsTxn_ImportNcdexLotPleadge!Mst_GSL_ID
             TmpRs5.Find "mKey1='" & mkey1 & "'"
             If TmpRs5.EOF = False Then
                mCMPID = IIf(IsNull(TmpRs5!CMPID), 0, TmpRs5!CMPID)
                If IsNull(TmpRs5!PledgeNo) = True Or TmpRs5!PledgeNo = "" Then
                   TmpRs5!PledgeNo = mPledgeCreation_ID
                Else
                   TmpRs5!PledgeNo = TmpRs5!PledgeNo & "~" & mPledgeCreation_ID
                End If
                
                TmpRs5!PledgeBags = IIf(IsNull(RsTxn_ImportNcdexLotPleadge!Bags), 0, RsTxn_ImportNcdexLotPleadge!Bags)
                TmpRs5!PledgeWeight = IIf(IsNull(RsTxn_ImportNcdexLotPleadge!ActualQuantity), 0, RsTxn_ImportNcdexLotPleadge!ActualQuantity)
                TmpRs5.Update
                TmpRs5.Requery
             End If
             RsTxn_ImportNcdexLotPleadge.MoveNext
         Next
      End If
   End If
   Call Gen_Email(mPledgeCreation_ID, mSM_Prefix_)
End If
End Sub
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "10.0.0.13"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT> " & mPreMsg & " <table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  .MailFrom = Gen_UserEmailID
'  'message
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "amol.pa@ncml.com" '"sr.ccg@ncml.com"
'  End If
'  .SendTo = ToADD_
'
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_ & "," & Gen_UserEmailID
'  End If
'
'  .BCC = "parag.r@ncml.com"
'
'  .MessageSubject = MsgSubject_ '& " e-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "parag.r@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ '& " (from HR-Pro)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If


End Function


Private Sub Gen_Email(mmPledgeId As Variant, mmSM_Prefix_ As Variant)

Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim mSumBags As Double
Dim mSumQty As Double

mMsg = ""
mPreMsg = ""
mPostMsg = ""

tmp = " Select TPC.Txn_PledgeCreation_ID ,ML.LocationName,MS.StateName,ML.StateID,ML.SM_Prefix,MCM.CMPID,MCM.CMPNAME,"
tmp = tmp & " MGo.Address,INLP.Code as WHRNoNew,INLP.ImportDate as WHRDate,INLP.ValidityDate,MCL.ClientName,"
tmp = tmp & " MC.Commodity,Sum(INLP.ValueCommodity) 'ValueCommodity',MB.BankName,MBR.BranchName,MCL.ClientID"
tmp = tmp & " From Txn_ImportNcdexLotPleadge INLP"
tmp = tmp & " Inner join Mst_GSL MG on INLP.SM_Prefix=MG.SM_Prefix And INLP.Mst_GSL_ID=MG.Mst_GSL_ID"
tmp = tmp & " Inner join Txn_PledgeCreation TPC on INLP.Txn_PledgeCreation_ID=TPC.Txn_PledgeCreation_ID ANd INLP.SM_Prefix=TPC.SM_Prefix"
tmp = tmp & " Inner join Mst_Godown MGo on MG.GodownID=MGo.GodownID"
tmp = tmp & " Inner join Mst_Commodity MC on MG.CommodityID=MC.CommodityID"
tmp = tmp & " Inner join Mst_CMP MCM on MGo.CMPID=MCM.CMPID"
tmp = tmp & " Inner join Mst_Location ML on MCM.LocationId=ML.LocationId"
tmp = tmp & " Inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Inner join Mst_Client MCL on COnvert(VArchar,INLP.CMSEClientId)=MCL.ClientIDRow And MCL.Status='A'"
tmp = tmp & " Inner join Mst_Bank MB on TPC.BankID=MB.BankID"
tmp = tmp & " Inner join Mst_Branch MBR on TPC.BranchID=MBR.BranchID"
tmp = tmp & " Where TPC.Txn_PledgeCreation_ID = " & Val(mmPledgeId) & " ANd TPC.SM_Prefix = '" & mmSM_Prefix_ & "'"
tmp = tmp & " Group By TPC.Txn_PledgeCreation_ID ,ML.LocationName,MS.StateName,ML.StateID,ML.SM_Prefix,MCM.CMPID,MCM.CMPNAME,"
tmp = tmp & " MGo.Address,INLP.Code,INLP.ImportDate,INLP.ValidityDate,MCL.ClientName,"
tmp = tmp & " MC.Commodity,MB.BankName,MBR.BranchName,MCL.ClientID"
Call TmpTable
If TmpRs.RecordCount > 0 Then

   mPreMsg = "<P> Dear SR Team , <br> "
   mPreMsg = mPreMsg & " The below mentioned WHR/ISIN is pledged today in mentioned CMP warehouse. The details of the pledge are given below. Kindly generate CAR No for the mentioned WHR to the lending Bank at the earliest."
   mPreMsg = mPreMsg & "<br>"
   
   mMsg = mMsg & "<tr><td> Pledge ID </td><td> " & TmpRs!SM_Prefix & " - " & mmPledgeId & " </td></tr>"
   mMsg = mMsg & "<tr><td> CMP ID </td><td> " & TmpRs!CMPID & " </td></tr>"
   mMsg = mMsg & "<tr><td> CMP Name </td><td> " & TmpRs!CMPName & "  </td></tr>"
   mMsg = mMsg & "<tr><td> Godown Address </td><td> " & TmpRs!Address & "  </td></tr>"
   mMsg = mMsg & "<tr><td> Location Name </td><td> " & TmpRs!Locationname & "  </td></tr>"
   mMsg = mMsg & "<tr><td> State Name </td><td> " & TmpRs!StateName & "  </td></tr>"
   
   mMsg = mMsg & "<tr><td> WHR No </td><td> " & TmpRs!WHRNoNew & "   </td></tr>"
   mMsg = mMsg & "<tr><td> WHR Date </td><td> " & Format(TmpRs!WHRDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
   mMsg = mMsg & "<tr><td> WHR Valid Date </td><td> " & Format(TmpRs!ValidityDate, "dd-mmm-yyyy") & "  </td></tr>"
   
   mMsg = mMsg & "<tr><td> Borrower Name and ID </td><td> " & TmpRs!ClientName & " - " & TmpRs!ClientID & " </td></tr>"
   mMsg = mMsg & "<tr><td> Commodity </td><td> " & TmpRs!Commodity & " </td></tr>"
   
   mMsg = mMsg & "<tr><td> Value of Commodity </td><td> " & Val(TmpRs!ValueCommodity) & " </td></tr>"
   
   mMsg = mMsg & "<tr><td> Bank Name </td><td> " & TmpRs!BankName & " </td></tr>"
   mMsg = mMsg & "<tr><td> Branch Name </td><td> " & TmpRs!BranchName & " </td></tr>"
   mMsg = mMsg & "<tr><td> Pledge Created by  </td><td> " & Gen_UserName & " - " & Gen_UserLoginID & " (" & GetDesignaion(Gen_DesignationID) & ")</td></tr>"
   
   tmp = " Select ME.EmployeeName,ME.EmployeeNo,ME.Flag,ME.EmployeeEmailID from Mst_EmpCMPMapping MEC "
   tmp = tmp & " Inner Join Mst_Employee ME On MEC.EMployeeID = ME.EmployeeID And ME.Flag = 'Y'"
   tmp = tmp & " Where MEC.CMPID = " & TmpRs!CMPID & ""
   
   Call Tmp1Table
   TOCC = ""
   If TmpRs1.RecordCount > 0 Then
      For j = 1 To TmpRs1.RecordCount
          mMsg = mMsg & "<tr><td> CMP Manager  </td><td> " & TmpRs1!EmployeeName & " ~ " & TmpRs1!EmployeeNo & " ( Status - " & TmpRs1!Flag & ")</td></tr>"
          If TOCC <> "" Then
             If InStr(1, TOCC, TmpRs1!EmployeeEmailID) = 0 Then
                TOCC = TOCC & "," & TmpRs1!EmployeeEmailID
             End If
          Else
             TOCC = TmpRs1!EmployeeEmailID
          End If
          TmpRs1.MoveNext
      Next
   End If
   
   mGridMsg = ""
   
   mGridMsg = mGridMsg & "<tr><td>Godown No</td><td>Stack No</td><td>Lot No</td><td>Pledged No of Bags</td><td>Pledged Qty</td></tr>"
   mSumBags = 0: mSumQty = 0
   
   tmp = " Select MGo.GodownNo,MG.StackNo,MG.LotNo,PledgeNoOfBags,PledgeQuantity From Txn_PledgeCreation_GSL TPG "
   tmp = tmp & " Inner join Mst_GSL MG on TPG.Mst_GSL_ID=MG.Mst_GSL_ID and TPG.SM_Prefix=MG.SM_Prefix"
   tmp = tmp & " Inner join Mst_Godown MGo on MG.Godownid=MGo.Godownid"
   tmp = tmp & " Where TPG.PledgeID = " & mmPledgeId & " And TPG.SM_Prefix = '" & mmSM_Prefix_ & "'"
   Call Tmp1Table
   For j = 1 To TmpRs1.RecordCount
       mGridMsg = mGridMsg & "<tr><td>" & TmpRs1!GodownNo & "</td><td>" & TmpRs1!StackNo & "</td><td>" & TmpRs1!LotNo & "</td><td>" & TmpRs1!PledgeNoOfBags & "</td><td>" & TmpRs1!PledgeQuantity & "</td></tr>"
       mSumBags = mSumBags + TmpRs1!PledgeNoOfBags
       mSumQty = mSumQty + TmpRs1!PledgeQuantity
       TmpRs1.MoveNext
   Next
   mGridMsg = mGridMsg & "<tr><td> &nbsp </td><td> &nbsp </td><td> Total Pledge </td><td> " & mSumBags & " </td><td> " & mSumQty & " </td></tr>"
   mMsg = mMsg & "<tr><td> Details </td><td><table border = 1>" & mGridMsg & "</table></td></tr>"
   
'   tmp = "Select * from Mst_StateDRCMapping Where StateID = " & TmpRs!StateID & " And Flag = 'A' "
'   Call Tmp1Table
'   TOADD = ""
'    If TmpRs1.RecordCount > 0 Then
'      For j = 1 To TmpRs1.RecordCount
'          If TOADD = "" Then
'             TOADD = TmpRs1!EmailID
'          Else
'             TOADD = TOADD & "," & TmpRs1!EmailID
'          End If
'          TmpRs1.MoveNext
'      Next
'   End If
   
   TOADD = "sr.ccg@ncml.com"
   
'   MsgBox TOADD
'   MsgBox TOCC & "," & Gen_UserEmailID
   TOCC = TOCC & "," & Gen_UserEmailID
   mMsgSubject = "New Pledge created in StoreMan - PledgeID is " & mmPledgeId & " in Location  " & TmpRs!Locationname & ""
   
'   TOADD = "parag.r@ncml.com"
'   TOCC = "prajakta.p@ncml.com"
   
   Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)
End If

End Sub

Public Function GetMaxID(mColumn_ As Variant, mTable_ As Variant, Optional wc_ As Variant)
Dim Retval As Double
If wc_ <> "" Then
   tmp = "Select Max(" & mColumn_ & ") from " & mTable_ & "  " & wc_ & " "
Else
   tmp = "Select Max(" & mColumn_ & ") from " & mTable_ & ""
End If
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetMaxID = Retval
End Function

Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_Bank("")

CmbBankName.Clear
If RsMst_Bank.RecordCount = 0 Then
   MsgBox "No Bank found!!!"
   'CmdInsuranceCo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Bank.RecordCount
    CmbBankName.AddItem RsMst_Bank!BankName
    RsMst_Bank.MoveNext
Next
CmbBankName.Text = tmp
End Sub
Private Sub CmbBankName_LostFocus()
If CmbBankName.Text = "" Then
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankName = '" & CmbBankName.Text & "'"
If RsMst_Bank.EOF Then
   MsgBox "Invalid Bank selection!!!"
   CmbBankName.SetFocus
   Exit Sub
End If
mBankID = RsMst_Bank!BankID
If CmbBankName.Text <> tmp Then
   CmbBranchName.Text = ""
End If
   
End Sub

Private Sub CmbBranchName_GotFocus()
If CmbBankName.Text = "" Then Exit Sub
If IsNull(mBankID) Then Exit Sub
tmp = CmbBranchName.Text
Call TableMst_Branch(" Where BankID =  " & mBankID & " Order by  BranchID")
CmbBranchName.Clear
For I = 1 To RsMst_Branch.RecordCount
    CmbBranchName.AddItem RsMst_Branch!BranchName
    RsMst_Branch.MoveNext
Next
CmbBranchName.Text = tmp
End Sub
Private Sub CmbBranchName_LostFocus()
If CmbBranchName.Text = "" Then
   Exit Sub
End If
RsMst_Branch.MoveFirst
RsMst_Branch.Find "BranchName = '" & CmbBranchName.Text & "'"
If RsMst_Branch.EOF Then
   MsgBox "Invalid selection!!!"
   CmbBranchName.SetFocus
   Exit Sub
End If
mBranchID = RsMst_Branch!BranchID
End Sub


