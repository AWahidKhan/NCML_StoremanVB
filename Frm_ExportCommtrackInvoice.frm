VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportCommtrackInvoice 
   Caption         =   "Export Commtrack Invoice"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
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
      Height          =   8145
      Left            =   90
      TabIndex        =   14
      Top             =   1170
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7680
         Left            =   120
         TabIndex        =   15
         Top             =   360
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
      Height          =   1095
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   15015
      Begin VB.ComboBox CmbJournalCode 
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
         Left            =   5280
         TabIndex        =   4
         Top             =   600
         Width           =   3240
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
         Height          =   350
         Left            =   8760
         TabIndex        =   5
         Top             =   240
         Width           =   1260
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
         Left            =   12540
         TabIndex        =   10
         Top             =   240
         Width           =   1260
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
         Left            =   11280
         TabIndex        =   9
         Top             =   240
         Width           =   1260
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
         Left            =   10020
         TabIndex        =   7
         Top             =   240
         Width           =   1260
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pending"
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
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Old"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox CmbBatch 
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
         Height          =   360
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1575
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
         ItemData        =   "Frm_ExportCommtrackInvoice.frx":0000
         Left            =   5280
         List            =   "Frm_ExportCommtrackInvoice.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3255
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
         Left            =   13800
         TabIndex        =   11
         Top             =   240
         Width           =   1140
      End
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
         Left            =   8880
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Journal Code"
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
         Left            =   3960
         TabIndex        =   16
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Batch No"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   270
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Frm_ExportCommtrackInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCode, mSTCode As Variant
Dim ExpFlag, MaxID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mJournalBatchName, mJournalType As Variant
Dim mIncomeTypeID, mHSNNo, mAccountType, mInvoiceType As Variant

Private Sub ChkCheckAll_Click()
If MSHFlexGrid3.Rows < 3 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.row = I
       If MSHFlexGrid3.TextMatrix(I, 3) <> "" And MSHFlexGrid3.TextMatrix(I, 6) <> "" And MSHFlexGrid3.TextMatrix(I, 13) <> "" And Val(MSHFlexGrid3.TextMatrix(I, 4)) > 0 Then
          MSHFlexGrid3.row = I
          MSHFlexGrid3.TextMatrix(I, 25) = strChecked
       End If
   Next
Else
   For I = 2 To MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.row = I
       MSHFlexGrid3.TextMatrix(I, 25) = strUnChecked
   Next
End If
End Sub

Private Sub CmbBatch_GotFocus()
CmbInvoiceType.Text = ""
tmp = "select distinct BatchID As BatchID from Txn_MonthlyCommtrackInvoice Where BatchID > 0  order by BatchID"
Call Tmp1Table
CmbBatch.Clear
If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
       CmbBatch.AddItem TmpRs1.Fields(0)
       MaxID = TmpRs1.Fields(0) + 1
       TmpRs1.MoveNext
   Next
Else
   MaxID = 1
End If
End Sub

Private Sub CmbBatch_LostFocus()
If CmbBatch.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "BatchID = " & Val(CmbBatch.Text) & ""
If TmpRs1.EOF Then
   MsgBox "Invalid selection!!!"
   CmbBatch.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbInvoiceType_LostFocus()
If CmbInvoiceType.Text = "" Then Exit Sub
If CmbInvoiceType.Text <> "Non Taxable" And CmbInvoiceType.Text <> "Taxable" Then
   MsgBox "Invalid Selection"
   CmbInvoiceType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbJournalCode_GotFocus()
Dim mR As Variant
tmp1 = CmbJournalCode
CmbJournalCode.Clear

'tmp = " SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]" '''JT.JournalTypeName,JCM.JournalCode
'tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
'tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where ME.EmployeeNo = '" & Gen_UserLoginID & "' Order by JCM.JournalTypeID"

tmp = " SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]" '''JT.JournalTypeName,JCM.JournalCode
tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where (ME.EmployeeNo = '" & Gen_UserLoginID & "' Or ME.EmployeeNo = '4474') And JCM.JournalTypeID = 2  Order by JCM.JournalTypeID"

Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
For mR = 1 To TmpRsHRPro.RecordCount
   CmbJournalCode.AddItem IIf(IsNull(TmpRsHRPro!JournalCode), "", TmpRsHRPro!JournalCode)
   TmpRsHRPro.MoveNext
Next
Else
   MsgBox "No Record found !!!! "
   Exit Sub
End If
CmbJournalCode = tmp1
End Sub

Private Sub CmbJournalCode_LostFocus()
Dim mSplit As Variant
If CmbJournalCode.Text = "" Then
   mJournalBatchName = ""
   mJournalType = ""
   Exit Sub
End If
mSplit = Split(CmbJournalCode.Text, "~")
If UBound(mSplit) <> 1 Then
   MsgBox "Invalid selection!!! "
   CmbJournalCode.SetFocus
   Exit Sub
End If
TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "JournalCode = '" & CmbJournalCode.Text & "'"
If TmpRsHRPro.EOF Then
   MsgBox "Invalid selection "
   CmbJournalCode.SetFocus
   Exit Sub
End If
mJournalType = UCase(mSplit(0))
mJournalBatchName = UCase(mSplit(1))
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Export_CommtrackInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Export_CommtrackInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    If Option1.Value = True Then
       For C = 0 To MSHFlexGrid3.Cols - 2
           If InStr(MSHFlexGrid3.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid3.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid3.TextMatrix(I, C), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
              End If
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
           End If
        Next
    Else
        For C = 0 To MSHFlexGrid3.Cols - 1
            If InStr(MSHFlexGrid3.TextMatrix(I, C), "/") > 0 Then
               If IsDate(MSHFlexGrid3.TextMatrix(I, C)) Then
                  oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid3.TextMatrix(I, C), "MM/dd/yyyy")
               Else
                  oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
               End If
            Else
               oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
            End If
        Next
    End If
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdExport_Click()
Dim Path, strFilName, strFilLocation As Variant

''-- On 05 Aug 2016
Dim mUserBatchID, mInvoiceType As Variant
'mJournalType = "SALES"
'mUserBatchID = GetGUID
'mJournalBatchName = GetJournalDetailsFromHRPro(Gen_UserLoginID, mJournalType)
'If Trim(mJournalBatchName) = "" Then
'   MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'   Exit Sub
'End If
''-- On 05 Aug 2016
'Commented on 12 Aug 2017 during GST effect implementation
If Option1.Value = True Then
    If CmbJournalCode.Text = "" Then
       MsgBox "Journal Code not selected"
       CmbJournalCode.SetFocus
       Exit Sub
    End If
    mUserBatchID = GetGUID
'    If Trim(mJournalType) = "" Or Trim(mJournalBatchName) = "" Then
'       MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'       Exit Sub
'    End If
End If

Path = App.Path
MaxID = GetMaxBatchID
If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_CommtrackInvoice" & "_" & CmbInvoiceType.Text & "_" & MaxID & " .csv"
Else
   strFilName = "\Navision\NavUpload_CommtrackInvoice" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
'Commented on 12 Aug 2017 during GST effect implementation
    Call TableTxn_InvoicesForNAVISION(" Where G_UID Is Null ")
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1

End If

If Option1.Value = True Then
   If CmbInvoiceType.Text = "Non Taxable" Then
      tmp = "Select * From Txn_MonthlyCommtrackInvoice Where isnull(BatchID,0) = 0 And Flag = 'N' "
   Else
      tmp = "Select * From Txn_MonthlyCommtrackInvoice Where isnull(BatchID,0) = 0 And Flag = 'T' "
   End If
   Call Tmp1Table
End If

With MSHFlexGrid3
For I = 2 To .Rows - 1
    tmp = ""
    If Option1.Value = True Then
       If .TextMatrix(I, 25) = strChecked Then
       
'Commented on 12 Aug 2017 during GST effect implementation

            RsTxn_InvoicesForNAVISION.AddNew
            RsTxn_InvoicesForNAVISION!Date = Trim(.TextMatrix(I, 0))
            RsTxn_InvoicesForNAVISION!Invoice = Trim(.TextMatrix(I, 1))
            RsTxn_InvoicesForNAVISION!Customer = Trim(.TextMatrix(I, 2))
            RsTxn_InvoicesForNAVISION!NavCode = Trim(.TextMatrix(I, 3))
            RsTxn_InvoicesForNAVISION!Amount = Trim(.TextMatrix(I, 4))
            RsTxn_InvoicesForNAVISION!GLAccount = Trim(.TextMatrix(I, 5))
            RsTxn_InvoicesForNAVISION!GLACNAVCode = Trim(.TextMatrix(I, 6))
            RsTxn_InvoicesForNAVISION!Narration = Trim(.TextMatrix(I, 12))
            RsTxn_InvoicesForNAVISION!LocationNAVCode = Trim(.TextMatrix(I, 13))
            RsTxn_InvoicesForNAVISION!DeptDim = Trim(.TextMatrix(I, 14))
            RsTxn_InvoicesForNAVISION!PINo = Trim(.TextMatrix(I, 18))
            RsTxn_InvoicesForNAVISION!STCode = Trim(.TextMatrix(I, 19))
            RsTxn_InvoicesForNAVISION!AgreeNonAgree = Trim(.TextMatrix(I, 20))
            
            RsTxn_InvoicesForNAVISION!GSTGroupCode = Trim(.TextMatrix(I, 21))
            RsTxn_InvoicesForNAVISION!HSNNo = Trim(.TextMatrix(I, 22))
            RsTxn_InvoicesForNAVISION!AccountType = Trim(.TextMatrix(I, 23))
            RsTxn_InvoicesForNAVISION!NumberSeriesFlag = Trim(.TextMatrix(I, 24))
            
            RsTxn_InvoicesForNAVISION!CreatedBy = Gen_UserLoginID
            RsTxn_InvoicesForNAVISION!G_UID = GetGUID
            RsTxn_InvoicesForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
            'RsTxn_InvoicesForNAVISION!DataTransfer = Null
            RsTxn_InvoicesForNAVISION!InvoiceType = "COMMTRACK"
            RsTxn_InvoicesForNAVISION!JournalBatchType = Trim(mJournalType)
            RsTxn_InvoicesForNAVISION!JournalBatchName = Trim(mJournalBatchName)
            'RsTxn_InvoicesForNAVISION!ReverseInvoice = Null
            'RsTxn_InvoicesForNAVISION!EntryNoTransfer = Null
            'RsTxn_InvoicesForNAVISION!EntryNoTransferRemark = Null
            RsTxn_InvoicesForNAVISION!UserBatchID = mUserBatchID
            
            ' For integration with Business Central - Start - 10-10-2023
            RsTxn_InvoicesForNAVISION!DataTransfer = "T"
            ' For integration with Business Central - End - 10-10-2023
            
            RsTxn_InvoicesForNAVISION.Update
       
'          For c = 0 To .Cols - 2
'              If tmp = "" Then
'                 tmp = tmp & .TextMatrix(i, c)
'              Else
'                 tmp = tmp & "," & .TextMatrix(i, c)
'              End If
'          Next
'          Print #1, tmp
          
          If Option1.Value = True Then
            If TmpRs1.RecordCount > 0 Then
               TmpRs1.MoveFirst
               TmpRs1.Find "PINo = " & Val(Replace(.TextMatrix(I, 18), "Commtrack", "")) & " "
               If TmpRs1.EOF = False Then
                  TmpRs1!BatchID = MaxID
                  ''-- Ullhas Patil 09 Nov 2017 -- Save the GSTGroupCode, HSN Code, Invoice Type
                    TmpRs1!GSTGroupCode = .TextMatrix(I, 21)
                    TmpRs1!HSNNo = .TextMatrix(I, 22)
                    TmpRs1!AccountType = .TextMatrix(I, 23)
                  ''----------------------------
                  TmpRs1.Update
               End If
            End If
          End If
          
       End If
    Else
        For C = 0 To .Cols - 1
            If tmp = "" Then
               tmp = tmp & .TextMatrix(I, C)
            Else
               tmp = tmp & "," & .TextMatrix(I, C)
            End If
        Next
        Print #1, tmp
    End If
Next
End With

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
End Sub

Private Function GetMaxBatchID()
tmp = "select Max(BatchID) from Txn_MonthlyCommtrackInvoice"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   GetMaxBatchID = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0)) + 1
Else
   GetMaxBatchID = 1
End If
End Function

Private Sub CmdGetReport_Click()
Dim wc As Variant
ExpFlag = ""
If Option1.Value = True Then
   If CmbInvoiceType.Text = "" Then
      MsgBox "Select Invoice Type."
      CmbInvoiceType.SetFocus
      Exit Sub
   End If
   If CmbInvoiceType.Text = "Non Taxable" Then
      wc = " Where  TMC.Flag='N' And isnull(TMC.BatchID,0) = 0"
   ElseIf CmbInvoiceType.Text = "Taxable" Then
      wc = " Where  TMC.Flag='T' And isnull(TMC.BatchID,0) = 0"
   End If
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
    wc = " Where TMC.BatchID=" & Val(CmbBatch.Text) & ""
End If
Call Grid_Head



If CmbInvoiceType.Text = "Non Taxable" Then
   mGLAccNAVCode = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "G")
   mHSNNo = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "H")
   mSTCode = "GSTS00"
   mAccountType = "AGRI"
   mInvoiceType = "Bill of Supply"
Else
   mGLAccNAVCode = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "G")
   mSTCode = "GSTS18"
   mHSNNo = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "H")
   mAccountType = "NAGRI"
   mInvoiceType = "Taxable"
End If


'tmp = "Select TMC.InvoiceID, TMC.MonthYear, TMC.LocationID, TMC.CPNavCode, TMC.Amt,"
'tmp = tmp & " TMC.STAmt , TMC.TotalAmt, TMC.PINo, TMC.BatchID, ML.NAVCode, TMC.Flag, ML.LocationName"
'tmp = tmp & " From Txn_MonthlyCommtrackInvoice TMC"
'tmp = tmp & " Left join Mst_Location ML on TMC.LocationID=ML.LocationID "
tmp = " Select TMC.InvoiceID, TMC.MonthYear, TMC.LocationID, MDGST.GSTNavCode 'ClientNAVCode', TMC.Amt,"
tmp = tmp & " TMC.STAmt , TMC.TotalAmt, TMC.PINo, TMC.BatchID, ML.NAVCode, TMC.Flag, ML.LocationName"
tmp = tmp & " From Txn_MonthlyCommtrackInvoice TMC"
tmp = tmp & " Left join Mst_Location ML on TMC.LocationID=ML.LocationID"
tmp = tmp & " LEft Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " LEft Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Left Join Mst_Depository MD on TMC.CPID = MD.DepositoryID And MD.ExchangeTypeID = 2"
tmp = tmp & " Left Join Mst_DepositoryStateWiseGST MDGST On MD.DepositoryID = MDGST.DepositoryID and MSI.StateIndiaID = MDGST.StateIndiaID "
tmp = tmp & wc

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
   If TmpRs!Flag = "T" Then
      CmbInvoiceType.Text = "Taxable"
   Else
      CmbInvoiceType.Text = "Non Taxable"
   End If
   Frame5.Caption = CmbInvoiceType.Text
   For I = 1 To TmpRs.RecordCount
       If Option1.Value = True Then
          MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!MonthYear, "mm/dd/yyyy") 'MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!MonthYear + 1, "mm/dd/yyyy") ', "dd/mm/yyyy")
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!MonthYear, "mm/dd/yyyy") ', "dd/mm/yyyy")
       End If
       MSHFlexGrid3.TextMatrix(I + 1, 1) = "Invoice"
       MSHFlexGrid3.TextMatrix(I + 1, 2) = "Customer"
       MSHFlexGrid3.TextMatrix(I + 1, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
       MSHFlexGrid3.TextMatrix(I + 1, 4) = TmpRs!Amt
       MSHFlexGrid3.TextMatrix(I + 1, 5) = "G/L Account"
       MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCode
       MSHFlexGrid3.TextMatrix(I + 1, 21) = mSTCode
       If CmbInvoiceType.Text = "Non Taxable" Then
          MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being Storage & Warehouse charges for the NERL/NCDEX clients at " & TmpRs!Locationname & " for the month of " & MonthName(Month(TmpRs!MonthYear)) & " " & Year(TmpRs!MonthYear) & " (Without GST) "
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being Storage & Warehouse charges for the NERL/NCDEX clients at " & TmpRs!Locationname & " for the month of " & MonthName(Month(TmpRs!MonthYear)) & " " & Year(TmpRs!MonthYear) & " (With GST) "
       End If
       MSHFlexGrid3.TextMatrix(I + 1, 13) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
       MSHFlexGrid3.TextMatrix(I + 1, 14) = "CMP"
       MSHFlexGrid3.TextMatrix(I + 1, 20) = mAccountType
       MSHFlexGrid3.TextMatrix(I + 1, 22) = mHSNNo
       MSHFlexGrid3.TextMatrix(I + 1, 23) = mInvoiceType
       MSHFlexGrid3.TextMatrix(I + 1, 24) = "M"
       
       If Option1.Value = True Then
          MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!PINo), "", "Commtrack" & TmpRs!PINo)
       Else
         If Val(CmbBatch.Text) >= 33 Then
            MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!PINo), "", "Commtrack" & TmpRs!PINo)
         Else
            MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!PINo), "", TmpRs!PINo)
         End If
       End If
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I + 1, 4))
       
       If IsNull(TmpRs!ClientNAVCode) = True Or TmpRs!ClientNAVCode = "" Then
          ExpFlag = "N"
       End If
       If IsNull(TmpRs!NavCode) = True Or TmpRs!NavCode = "" Then
          ExpFlag = "N"
       End If
       If mGLAccNAVCode = "" Then
          ExpFlag = "N"
       End If
       
       If Option1.Value = True Then
          MSHFlexGrid3.row = I + 1
          MSHFlexGrid3.Col = 25
          MSHFlexGrid3.CellFontName = "Wingdings"
          MSHFlexGrid3.CellFontSize = 12
          MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid3.Text = strUnChecked
       End If
       
       TmpRs.MoveNext
   Next
End If

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmbBatch.Text = ""
CmbInvoiceType.Text = ""
CmdExport.Enabled = True
Option1.Value = True
Call Option1_Click
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame5.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
mIncomeTypeID = 12
Frame1.Enabled = True
Frame5.Enabled = True
Call Grid_Head

tmp = "SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]"
tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where ME.EmployeeNo = '4474' And JCM.JournalTypeID = 2 "
tmp = tmp & " Order by JCM.JournalTypeID"

Call TmpTableHRPro

If TmpRsHRPro.RecordCount = 1 Then
   CmbJournalCode.Text = TmpRsHRPro!JournalCode
'   mSForwardToID = TmpRs!EmployeeID
Else
   CmbJournalCode.Text = ""
'   mSForwardToID = Null
End If


mSplit = Split(CmbJournalCode.Text, "~")
mJournalType = UCase(mSplit(0))
mJournalBatchName = UCase(mSplit(1))



End Sub
Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    If Option1.Value = True Then
       .Cols = 26 '25
    Else
       .Cols = 25 '24
    End If
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Date"
    .TextMatrix(1, 1) = "Invoice"
    .TextMatrix(1, 2) = "Customer"
    .TextMatrix(1, 3) = "CPNavCode"
    .TextMatrix(1, 4) = "Amount"
    .TextMatrix(1, 5) = "G/L Account"
    .TextMatrix(1, 6) = "GL A/C NAVCode "
    .TextMatrix(1, 12) = "Narration"
    .TextMatrix(1, 13) = "Location NAV Code"
    .TextMatrix(1, 14) = "CMP"
    .TextMatrix(1, 18) = "PI No"
    .TextMatrix(1, 20) = "Agri/Non Agri"
    .TextMatrix(1, 21) = "GST Code"
    .TextMatrix(1, 22) = "HSN No"
    .TextMatrix(1, 23) = "Invoice Type"
    .TextMatrix(1, 24) = "Number Series Flag"
    
    If Option1.Value = True Then
       .TextMatrix(1, 25) = "Selection"
    End If
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
    .ColWidth(20) = 1000
    If Option1.Value = True Then
       .ColWidth(24) = 1000
    End If
End With

End Sub

Private Sub MSHFlexGrid3_Click()
With MSHFlexGrid3
     If .Col = 25 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        .Col = 25
        If .TextMatrix(.row, .Col) = strUnChecked Then
           If .TextMatrix(.row, 3) <> "" And .TextMatrix(.row, 6) <> "" And .TextMatrix(.row, 13) <> "" And Val(.TextMatrix(.row, 4)) > 0 Then
              .TextMatrix(.row, .Col) = strChecked
           End If
        Else
           .TextMatrix(.row, .Col) = strUnChecked
        End If
     End If
End With
End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Private Sub Option2_Click()
CmbInvoiceType.Text = ""
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbInvoiceType.Enabled = False
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = False
   Call Grid_Head
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   CmbInvoiceType.Enabled = True
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = True
   Call Grid_Head
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub

