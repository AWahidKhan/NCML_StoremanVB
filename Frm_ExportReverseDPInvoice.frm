VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_ExportReverseDPInvoice 
   Caption         =   "Export Reverse DP Invoices"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
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
      Height          =   7185
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   6840
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12065
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   0
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
         Left            =   3240
         TabIndex        =   4
         Top             =   720
         Width           =   5400
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
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
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1095
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
         Left            =   3240
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
         ItemData        =   "Frm_ExportReverseDPInvoice.frx":0000
         Left            =   6240
         List            =   "Frm_ExportReverseDPInvoice.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2415
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
         TabIndex        =   9
         Top             =   240
         Width           =   1140
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
         Left            =   1800
         TabIndex        =   15
         Top             =   780
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
         Left            =   1800
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
         Left            =   4920
         TabIndex        =   12
         Top             =   270
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Frm_ExportReverseDPInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCode, mSTCode As Variant
Dim ExpFlag, MaxID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mJournalBatchName, mJournalType As Variant

Private Sub CmbBatch_GotFocus()
CmbInvoiceType.Text = ""
tmp = "select distinct RBatchID As RBatchID from Mst_DPInvoiceUpload Where RBatchID > 0  order by RBatchID"
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
TmpRs1.Find "RBatchID = " & Val(CmbBatch.Text) & ""
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
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
tmp = " SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]" '''JT.JournalTypeName,JCM.JournalCode
tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where ME.EmployeeNo = '" & Gen_UserLoginID & "' Order by JCM.JournalTypeID"
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Export_DPInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Export_DPInvoice.xls")
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
If Option1.Value = True Then
    If CmbJournalCode.Text = "" Then
       MsgBox "Journal Code not selected"
       CmbJournalCode.SetFocus
       Exit Sub
    End If
    mUserBatchID = GetGUID
    If Trim(mJournalType) = "" Or Trim(mJournalBatchName) = "" Then
       MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
       Exit Sub
    End If
End If

Path = App.Path
MaxID = GetMaxBatchID
If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_DPInvoice" & "_" & CmbInvoiceType.Text & "_" & MaxID & " .csv"
Else
   strFilName = "\Navision\NavUpload_DPInvoice" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
End If
''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
   Call TableTxn_InvoicesForNAVISION(" Where G_UID Is Null ")
End If

If Option1.Value = True Then
   If CmbInvoiceType.Text = "Non Taxable" Then
      tmp = "Select * From Mst_DPInvoiceUpload Where RBatchID = 0 And Flag = 'N' And ReverseFlag = 'Y' "
   Else
      tmp = "Select * From Mst_DPInvoiceUpload Where RBatchID = 0  And Flag = 'T' And ReverseFlag = 'Y' "
   End If
   Call Tmp1Table
End If

For I = 2 To MSHFlexGrid3.Rows - 2
    tmp = ""
    If Option1.Value = False Then
'       For c = 0 To MSHFlexGrid3.Cols - 1
'           If tmp = "" Then
'              tmp = tmp & MSHFlexGrid3.TextMatrix(I, c)
'           Else
'              tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, c)
'           End If
'       Next

'        For c = 0 To MSHFlexGrid3.Cols - 1
'            If tmp = "" Then
'               tmp = tmp & MSHFlexGrid3.TextMatrix(I, c)
'            Else
'               tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, c)
'            End If
'        Next
'        Print #1, tmp

        RsTxn_InvoicesForNAVISION.AddNew
        RsTxn_InvoicesForNAVISION!Date = Trim(MSHFlexGrid3.TextMatrix(I, 0))
        RsTxn_InvoicesForNAVISION!Invoice = Trim(MSHFlexGrid3.TextMatrix(I, 1))
        RsTxn_InvoicesForNAVISION!Customer = Trim(MSHFlexGrid3.TextMatrix(I, 2))
        RsTxn_InvoicesForNAVISION!NavCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
        RsTxn_InvoicesForNAVISION!Amount = Trim(MSHFlexGrid3.TextMatrix(I, 4))
        RsTxn_InvoicesForNAVISION!Narration = Trim(MSHFlexGrid3.TextMatrix(I, 6))
        RsTxn_InvoicesForNAVISION!LocationNAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 7))
        RsTxn_InvoicesForNAVISION!DeptDim = Trim(MSHFlexGrid3.TextMatrix(I, 8))
        RsTxn_InvoicesForNAVISION!EmpDim = Trim(MSHFlexGrid3.TextMatrix(I, 9))
        RsTxn_InvoicesForNAVISION!PINo = Trim(MSHFlexGrid3.TextMatrix(I, 13))
        'RsTxn_InvoicesForNAVISION!PostInvoiceNo = Null
        RsTxn_InvoicesForNAVISION!EntryNo = Trim(MSHFlexGrid3.TextMatrix(I, 15))
        RsTxn_InvoicesForNAVISION!STCode = Trim(MSHFlexGrid3.TextMatrix(I, 16))
        RsTxn_InvoicesForNAVISION!GLAccount = Trim(MSHFlexGrid3.TextMatrix(I, 17))
        RsTxn_InvoicesForNAVISION!GLACNAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 18))
        'RsTxn_InvoicesForNAVISION!AgreeNonAgree = Null
        'RsTxn_InvoicesForNAVISION!TINo = Null
        'RsTxn_InvoicesForNAVISION!TIDate = Null
        'RsTxn_InvoicesForNAVISION!STAmt = Null
        'RsTxn_InvoicesForNAVISION!TIAmt = Null
        RsTxn_InvoicesForNAVISION!CreatedBy = Gen_UserLoginID
        RsTxn_InvoicesForNAVISION!G_UID = GetGUID
        RsTxn_InvoicesForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
        'RsTxn_InvoicesForNAVISION!DataTransfer = Null
        RsTxn_InvoicesForNAVISION!InvoiceType = "DPINVOICE"
        RsTxn_InvoicesForNAVISION!JournalBatchType = Trim(mJournalType)
        RsTxn_InvoicesForNAVISION!JournalBatchName = Trim(mJournalBatchName)
        RsTxn_InvoicesForNAVISION!ReverseInvoice = "Y"
        'RsTxn_InvoicesForNAVISION!EntryNoTransfer = Null
        'RsTxn_InvoicesForNAVISION!EntryNoTransferRemark = Null
        RsTxn_InvoicesForNAVISION!UserBatchID = mUserBatchID
        RsTxn_InvoicesForNAVISION.Update

        If Option1.Value = True Then
            If TmpRs1.RecordCount > 0 Then
               TmpRs1.MoveFirst
               TmpRs1.Find "PINo = " & Val(Replace(MSHFlexGrid3.TextMatrix(I, 13), "RDP", "")) & " "
               If TmpRs1.EOF = False Then
                  TmpRs1!RBatchID = MaxID
                  TmpRs1.Update
               End If
            End If
        End If
    Else
        For C = 0 To MSHFlexGrid3.Cols - 1
            If tmp = "" Then
               tmp = tmp & MSHFlexGrid3.TextMatrix(I, C)
            Else
               tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, C)
            End If
        Next
        Print #1, tmp
    End If
Next

'If Option1.Value = True Then
'
'   If CmbInvoiceType.Text = "Non Taxable" Then
'      tmp = "Update Mst_DPInvoiceUpload set BatchID = " & MaxID & " Where BatchID is Null And Flag='N' "
'   Else
'      tmp = "Update Mst_DPInvoiceUpload set BatchID = " & MaxID & " Where BatchID is Null And Flag='T' "
'   End If
'   Call TmpTable
'End If

'For I = 2 To MSHFlexGrid3.Rows - 1
'    tmp = ""
'    For c = 0 To MSHFlexGrid3.Cols - 1
'        If tmp = "" Then
'           tmp = tmp & MSHFlexGrid3.TextMatrix(I, c)
'        Else
'           tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, c)
'        End If
'    Next
'    Print #1, tmp
'Next
'If Option1.Value = True Then
'
'   If CmbInvoiceType.Text = "Non Taxable" Then
'      tmp = "Update Mst_DPInvoiceUpload set BatchID = " & MaxID & " Where BatchID is Null And Flag='N' "
'   Else
'      tmp = "Update Mst_DPInvoiceUpload set BatchID = " & MaxID & " Where BatchID is Null And Flag='T' "
'   End If
'   Call TmpTable
'End If
'MsgBox "Done !!!"

MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
End Sub

Private Function GetMaxBatchID()
tmp = "select Max(RBatchID) from Mst_DPInvoiceUpload"
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
      wc = " Where  MD.Flag='N' And MD.RBatchID = 0 And ReverseFlag = 'Y'"
   ElseIf CmbInvoiceType.Text = "Taxable" Then
      wc = " Where  MD.Flag='T' And MD.RBatchID = 0 And ReverseFlag = 'Y'"
   End If
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
    wc = " Where  MD.RBatchID=" & Val(CmbBatch.Text) & ""
'   If CmbInvoiceType.Text = "Non Taxable" Then
'      tmp = tmp & " Where  MD.Flag='N' And MD.BatchID=" & Val(CmbBatch.Text) & ""
'   ElseIf CmbInvoiceType.Text = "Taxable" Then
'      tmp = tmp & " Where  MD.Flag='T' And MD.BatchID=" & Val(CmbBatch.Text) & ""
'   End If
End If

Call Grid_Head

mGLAccNAVCode = GetGLAccNAVCode("DP")
mSTCode = GetGLAccNAVCode("DPSTCode")

tmp = "Select MD.DPInvoiceID,MD.MonthYear,MD.LocationID,MD.DPNavCode,MD.Amt,"
tmp = tmp & " MD.STAmt , MD.TotalAmt, MD.PINo, MD.BatchID, ML.NAVCode, MD.Flag,ML.LocationName,MD.ReverseDate,MD.ReverseRemark"
tmp = tmp & " From Mst_DPInvoiceUpload MD"
tmp = tmp & " Left join Mst_Location ML on MD.LocationID=ML.LocationID "
tmp = tmp & wc

Call TmpTable
'If TmpRs.RecordCount > 0 Then
'   MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
'   If TmpRs!Flag = "T" Then
'      CmbInvoiceType.Text = "Taxable"
'   Else
'      CmbInvoiceType.Text = "Non Taxable"
'   End If
'   Frame5.Caption = CmbInvoiceType.Text
'   For I = 1 To TmpRs.RecordCount
'       MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!ReverseDate, "mm/dd/yyyy") ', "dd/mm/yyyy")
'       MSHFlexGrid3.TextMatrix(I + 1, 1) = "Credit Memo"
'       MSHFlexGrid3.TextMatrix(I + 1, 2) = "Customer"
'       MSHFlexGrid3.TextMatrix(I + 1, 3) = IIf(IsNull(TmpRs!DPNAVCode), "", TmpRs!DPNAVCode)
'       MSHFlexGrid3.TextMatrix(I + 1, 4) = "-" & TmpRs!Amt
'       MSHFlexGrid3.TextMatrix(I + 1, 5) = "G/L Account"
'       MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCode
'       If CmbInvoiceType.Text = "Non Taxable" Then
'          'Commented On 05 May 2012
'          'MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being Storage & Warehousing Charges  (DP) for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!MonthYear)) & " " & Year(TmpRs!MonthYear) & " (Without ST) "
'
'          MSHFlexGrid3.TextMatrix(I + 1, 12) = IIf(IsNull(TmpRs!ReverseRemark), "", Trim(TmpRs!ReverseRemark)) '"Being Reversal of Storage & Warehouse charges for the DP/NCDEX clients at " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!MonthYear)) & " " & Year(TmpRs!MonthYear) & " (Without ST) "
'       Else
'          MSHFlexGrid3.TextMatrix(I + 1, 19) = mSTCode
'          'Commented On 05 May 2012
'          'MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being Storage & Warehousing Charges  (DP) for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!MonthYear)) & " " & Year(TmpRs!MonthYear) & " (With ST) "
'
'          MSHFlexGrid3.TextMatrix(I + 1, 12) = IIf(IsNull(TmpRs!ReverseRemark), "", Trim(TmpRs!ReverseRemark)) '"Being Reversal of Storage & Warehouse charges for the DP/NCDEX clients at " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!MonthYear)) & " " & Year(TmpRs!MonthYear) & " (With ST) "
'          'Being Storage & Warehouse charges for the DP/NCDEX clients at _____(location) for the month of _____2012 (with ST)
'       End If
'       MSHFlexGrid3.TextMatrix(I + 1, 13) = IIf(IsNull(TmpRs!NAVCode), "", TmpRs!NAVCode)
'       MSHFlexGrid3.TextMatrix(I + 1, 14) = "CMP"
'
'       If Option1.Value = True Then
'          MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!PINo), "", "DP" & TmpRs!PINo)
'       Else
'         If Val(CmbBatch.Text) >= 33 Then
'            MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!PINo), "", "DP" & TmpRs!PINo)
'         Else
'            MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!PINo), "", TmpRs!PINo)
'         End If
'       End If
'       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I + 1, 4))
'
'       If IsNull(TmpRs!DPNAVCode) = True Or TmpRs!DPNAVCode = "" Then
'          ExpFlag = "N"
'       End If
'       If IsNull(TmpRs!NAVCode) = True Or TmpRs!NAVCode = "" Then
'          ExpFlag = "N"
'       End If
'       If mGLAccNAVCode = "" Then
'          ExpFlag = "N"
'       End If
'
'       If Option1.Value = True Then
'          MSHFlexGrid3.Row = I + 1
'          MSHFlexGrid3.Col = 20
'          MSHFlexGrid3.CellFontName = "Wingdings"
'          MSHFlexGrid3.CellFontSize = 12
'          MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
'          MSHFlexGrid3.Text = strUnChecked
'       End If
'
'       TmpRs.MoveNext
'   Next
'
''   If ExpFlag = "N" Then
''      CmdExport.Enabled = False
''   Else
''      CmdExport.Enabled = True
''   End If
''Else
''   CmdExport.Enabled = False
'
'End If

If TmpRs.RecordCount > 0 Then
   If TmpRs!Flag = "T" Then
      CmbInvoiceType.Text = "Taxable"
   Else
      CmbInvoiceType.Text = "Non Taxable"
   End If
   For r = 1 To TmpRs.RecordCount

       I = MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.TextMatrix(I, 0) = Format(TmpRs!ReverseDate, "mm/dd/yyyy") ', "dd/mm/yyyy")
       MSHFlexGrid3.TextMatrix(I, 1) = "Credit Memo"
       MSHFlexGrid3.TextMatrix(I, 2) = "Customer"
       MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs!ReverseRemark), "", TmpRs!ReverseRemark)
       MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!DPNAVCode), "", TmpRs!DPNAVCode)
       MSHFlexGrid3.TextMatrix(I, 17) = "G/L Account"
       MSHFlexGrid3.TextMatrix(I, 18) = mGLAccNAVCode
       If MSHFlexGrid3.TextMatrix(I, 3) = "" Then ExpFlag = "N"
       MSHFlexGrid3.TextMatrix(I, 4) = -(TmpRs!Amt)
       'MSHFlexGrid3.TextMatrix(I, 5) = TmpRs!InvoiceNo
       MSHFlexGrid3.TextMatrix(I, 7) = TmpRs!NavCode
       MSHFlexGrid3.TextMatrix(I, 8) = "CMP"
       MSHFlexGrid3.TextMatrix(I, 9) = "EMP-1000"
       
       MSHFlexGrid3.TextMatrix(I, 13) = "RDP" & TmpRs!PINo
       tmp = "Select EntryNo from Txn_TIDetails Where  PIFlag='DP' And PIID=" & TmpRs!PINo & ""
       Call Tmp3Table
       If TmpRs3.RecordCount > 0 Then
          MSHFlexGrid3.TextMatrix(I, 15) = IIf(IsNull(TmpRs3!EntryNo), "", TmpRs3!EntryNo)
       End If
       If CmbInvoiceType.Text = "Non Taxable" Then
          MSHFlexGrid3.TextMatrix(I, 16) = ""
       Else
          MSHFlexGrid3.TextMatrix(I, 16) = mSTCode
       End If
       If MSHFlexGrid3.TextMatrix(I, 15) = "" Then ExpFlag = "N"
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       TmpRs.MoveNext
   Next
   If ExpFlag = "N" Then
      CmdExport.Enabled = False
   Else
      CmdExport.Enabled = True
   End If
Else
   CmdExport.Enabled = False
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
Frame1.Enabled = True
Frame5.Enabled = True
Call Grid_Head

End Sub
Public Sub Grid_Head()
'With MSHFlexGrid3
'    .Clear
'    .Rows = 3
'    .FixedRows = 2
'    If Option1.Value = True Then
'       .Cols = 21
'    Else
'       .Cols = 20
'    End If
'    .AllowUserResizing = flexResizeColumns
'    .TextMatrix(1, 0) = "Date"
'    .TextMatrix(1, 1) = "Invoice"
'    .TextMatrix(1, 2) = "Customer"
'    .TextMatrix(1, 3) = "DPNavCode"
'    .TextMatrix(1, 4) = "Amount"
'    .TextMatrix(1, 5) = "G/L Account"
'    .TextMatrix(1, 6) = "GL A/C NAVCode "
'    .TextMatrix(1, 12) = "Narration"
'    .TextMatrix(1, 13) = "Location NAV Code"
'    .TextMatrix(1, 14) = "CMP"
'    .TextMatrix(1, 18) = "PI No"
'    .TextMatrix(1, 19) = "ST Code"
'    If Option1.Value = True Then
'       .TextMatrix(1, 20) = "Selection"
'    End If
'    .ColWidth(0) = 1200
'    .ColWidth(1) = 1200
'    .ColWidth(2) = 1400
'    .ColWidth(3) = 1400
'    .ColWidth(4) = 1600
'    .ColWidth(5) = 1600
'    .ColWidth(6) = 1600
'    .ColWidth(7) = 0
'    .ColWidth(8) = 0
'    .ColWidth(9) = 0
'    .ColWidth(10) = 0
'    .ColWidth(11) = 0
'    .ColWidth(12) = 2000
'    .ColWidth(13) = 1600
'    .ColWidth(14) = 1400
'    .ColWidth(15) = 0
'    .ColWidth(16) = 0
'    .ColWidth(17) = 0
'    .ColWidth(18) = 1400
'    .ColWidth(19) = 1000
'    If Option1.Value = True Then
'       .ColWidth(20) = 1000
'    End If
'End With


With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 19
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(1, 0) = "Posting Date"
    .TextMatrix(1, 1) = "Document Type"
    .TextMatrix(1, 2) = "Account Type"
    .TextMatrix(1, 3) = "Account No."
    .TextMatrix(1, 4) = "Amount"
    .TextMatrix(1, 5) = "Invoice No"
    .TextMatrix(1, 6) = "Narration"
    .TextMatrix(1, 7) = "RegionDim"
    .TextMatrix(1, 8) = "DeptDim"
    .TextMatrix(1, 9) = "EmpDim"
    .TextMatrix(1, 10) = "Cheque No."
    .TextMatrix(1, 11) = "Cheque Date"
    .TextMatrix(1, 12) = "Deposit Date"
    .TextMatrix(1, 13) = "PINo"
    .TextMatrix(1, 14) = "PostInvoiceNo"
    .TextMatrix(1, 15) = "EntryNo"
    .TextMatrix(1, 16) = "STCode"
    .TextMatrix(1, 17) = "Bal Acc Type"
    .TextMatrix(1, 18) = "Bal Acc No"
     
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1600
    .ColWidth(5) = 1600
    .ColWidth(6) = 1600
    .ColWidth(7) = 2000
    .ColWidth(8) = 1600
    .ColWidth(9) = 1400
    .ColWidth(10) = 1400
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    
End With


End Sub

Private Sub MSHFlexGrid3_Click()
With MSHFlexGrid3
     If .Col = 20 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        .Col = 20
        If .TextMatrix(.row, .Col) = strUnChecked Then
           If .TextMatrix(.row, 3) <> "" And .TextMatrix(.row, 6) <> "" And .TextMatrix(.row, 13) <> "" And Val(.TextMatrix(.row, 4)) < 0 Then
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
'CmdExport.Enabled = False
CmbInvoiceType.Text = ""
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbInvoiceType.Enabled = False
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   CmbInvoiceType.Enabled = True
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub

