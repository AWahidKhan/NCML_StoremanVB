VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ExportTempInvoice 
   Caption         =   "Export Ad-Hoc Invoice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
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
      Height          =   8025
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7560
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13335
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbJournalCode 
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
         Left            =   8280
         TabIndex        =   8
         Top             =   720
         Width           =   2400
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
         Left            =   13695
         TabIndex        =   22
         Top             =   765
         Width           =   1215
      End
      Begin VB.TextBox TxtSearch 
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
         Height          =   375
         Left            =   10815
         TabIndex        =   21
         Top             =   690
         Width           =   1605
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "FAG Posting Date"
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
         Height          =   480
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Invoice Date"
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
         Height          =   240
         Left            =   1200
         TabIndex        =   2
         Top             =   308
         Width           =   1455
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
         Left            =   12480
         TabIndex        =   20
         Top             =   720
         Width           =   1140
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
         Left            =   10800
         TabIndex        =   9
         Top             =   240
         Width           =   1140
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
         Left            =   13980
         TabIndex        =   13
         Top             =   240
         Width           =   900
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
         Left            =   12960
         TabIndex        =   12
         Top             =   240
         Width           =   1020
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
         Left            =   11940
         TabIndex        =   11
         Top             =   240
         Width           =   1020
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
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   240
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
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   675
         Width           =   855
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
         Left            =   6210
         TabIndex        =   7
         Top             =   690
         Width           =   1320
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
         ItemData        =   "Frm_ExportTempInvoice.frx":0000
         Left            =   6195
         List            =   "Frm_ExportTempInvoice.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   3825
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   113573889
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   113573889
         CurrentDate     =   39884
      End
      Begin VB.Label Label3 
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
         Height          =   495
         Left            =   7560
         TabIndex        =   23
         Top             =   653
         Width           =   1095
      End
      Begin VB.Label Label14 
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
         Height          =   225
         Left            =   2760
         TabIndex        =   19
         Top             =   780
         Width           =   810
      End
      Begin VB.Label Label23 
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
         Height          =   225
         Left            =   2760
         TabIndex        =   18
         Top             =   315
         Width           =   990
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
         Left            =   5280
         TabIndex        =   15
         Top             =   735
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
         Height          =   495
         Left            =   5280
         TabIndex        =   14
         Top             =   173
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_ExportTempInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCode, mSTCode As Variant
Dim ExpFlag, MaxID As Variant
Dim mSMonth, mSYear, mSTotalDays, mForwardToID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mJournalBatchName, mJournalType As Variant

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Check2.Value = 0
'Else
'   Check1.Value = 1
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then
   Check1.Value = 0
'Else
'   Check1.Value = 1
End If
End Sub

Private Sub ChkCheckAll_Click()

If MSHFlexGrid3.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.row = I
       If MSHFlexGrid3.TextMatrix(I, 3) <> "" And MSHFlexGrid3.TextMatrix(I, 6) <> "" And MSHFlexGrid3.TextMatrix(I, 13) <> "" Then
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
tmp = "select distinct BatchID As BatchID from Txn_TempInvoice Where BatchID > 0  order by BatchID"
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Export_AdhocInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Export_AdhocInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
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
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
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

'Commented on 14 Aug 2017 during GST Effect implementaion
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

If MSHFlexGrid3.TextMatrix(1, 0) = "" Then Exit Sub
    
Path = App.Path
MaxID = GetMaxBatchID
If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_AdHocInvoice" & "_" & CmbInvoiceType.Text & "_" & MaxID & " .csv"
Else
   strFilName = "\Navision\NavUpload_AdHocInvoice" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
'Commeneted on 14 Aug 2017 during GST effect implementation
    Call TableTxn_InvoicesForNAVISION(" Where G_UID Is Null ")

    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1

End If

If Option1.Value = True Then
   'If CmbApprovedBy.Text <> "" Then
      'If CmbInvoiceType.Text = "Non Taxable" Then
       tmp = "Select * From Txn_TempInvoice Where Isnull(BatchID,0) = 0  And Flag='A' " 'And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'  "
      'Else
      '   tmp = "Select * From Txn_TempInvoice Where BatchID = 0  And Flag='A' " 'And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
      'End If
  ' Else
      
  ' End If
   Call Tmp1Table
End If

With MSHFlexGrid3
For I = 2 To .Rows - 1
    tmp = ""
    If Option1.Value = True Then
       If .TextMatrix(I, 25) = strChecked Then
          
'          For c = 0 To .Cols - 2
'              If tmp = "" Then
'                 tmp = tmp & .TextMatrix(i, c)
'              Else
'                 tmp = tmp & "," & .TextMatrix(i, c)
'              End If
'          Next
'          Print #1, tmp


'Commeneted on 14 Aug 2017 during GST effect implementation
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
            'RsTxn_InvoicesForNAVISION!STCode = Trim(.TextMatrix(i, 19))
            RsTxn_InvoicesForNAVISION!AgreeNonAgree = Trim(.TextMatrix(I, 20))
            
            RsTxn_InvoicesForNAVISION!GSTGroupCode = Trim(.TextMatrix(I, 21))
            RsTxn_InvoicesForNAVISION!HSNNo = Trim(.TextMatrix(I, 22))
            RsTxn_InvoicesForNAVISION!AccountType = Trim(.TextMatrix(I, 23))
            RsTxn_InvoicesForNAVISION!NumberSeriesFlag = Trim(.TextMatrix(I, 24))
            
            
            
            RsTxn_InvoicesForNAVISION!CreatedBy = Gen_UserLoginID
            RsTxn_InvoicesForNAVISION!G_UID = GetGUID
            RsTxn_InvoicesForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text) ''MaxID  '' CmbBatch.Text
            RsTxn_InvoicesForNAVISION!InvoiceType = "ADHOC"
            RsTxn_InvoicesForNAVISION!JournalBatchType = Trim(mJournalType)
            RsTxn_InvoicesForNAVISION!JournalBatchName = Trim(mJournalBatchName)
            RsTxn_InvoicesForNAVISION!UserBatchID = mUserBatchID
            
            ' For integration with Business Central - Start - 10-10-2023
            RsTxn_InvoicesForNAVISION!DataTransfer = "T"
            ' For integration with Business Central - End - 10-10-2023
            
            RsTxn_InvoicesForNAVISION.Update

          If Option1.Value = True Then
            If TmpRs1.RecordCount > 0 Then
               TmpRs1.MoveFirst
               TmpRs1.Find "PINo = " & Val(Replace(.TextMatrix(I, 18), "AH", "")) & " "
               If TmpRs1.EOF = False Then
                  TmpRs1!BatchID = MaxID
                  ''-- Ullhas Patil 09 Nov 2017 -- Save the GSTGroupCode, HSN Code, Invoice Type
                  TmpRs1!HSNNo = .TextMatrix(I, 22)
                  TmpRs1!AccountType = .TextMatrix(I, 23)
                  ''-------------------------------------------
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
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
ChkCheckAll.Value = 0
End Sub

Private Function GetMaxBatchID()
tmp = "select Max(BatchID) from Txn_TempInvoice"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   GetMaxBatchID = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0)) + 1
Else
   GetMaxBatchID = 1
End If
End Function
Private Sub CmdGetReport_Click()
Dim wc As Variant
Dim mGLAccNAVCodeResr, mGLAccNAVCodeSevDays, mGLAccNAVCodeNNNR, mGLAccNAVCodeRej, mGLAccNAVCodeSpot, mGLAccNAVCodeSpil, mGLAccNAVCodeOther, mGLAccNAVCodeDPClient, mGLAccNAVCodeCRC, mGLAccNAVCodeLoadUnload, mGLAccNAVCodeSCMIncome, mGLAccNAVCodeCommtrack, mGLAccNAVCodeRRN, mGLAccNAVCodeSOC, mGLAccNAVCodeWBIncome As Variant
ExpFlag = ""
If Option1.Value = True Then
      
   If Check1.Value = 0 And Check2.Value = 0 Then
      MsgBox "Select Date filter."
      Check1.SetFocus
      Exit Sub
   End If
   
   If CmbInvoiceType.Text = "" Then
      MsgBox "Select Invoice Type."
      CmbInvoiceType.SetFocus
      Exit Sub
   End If
   
   If CmbInvoiceType.Text = "Non Taxable" Then
      wc = " Where  MST.NAVCode = 'GSTS00' And Isnull(TPI.BatchID,0) = 0"
   ElseIf CmbInvoiceType.Text = "Taxable" Then
      wc = " Where  MST.NAVCode = 'GSTS18' And Isnull(TPI.BatchID,0) = 0"
   End If
   If Check1.Value = 1 Then
      If Gen_DBType = "MDB" Then
         wc = wc & "And TPI.InvoiceDate >= #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And TPI.InvoiceDate <= #" & Format(TxtToDate.Value, Gen_DatFormat) & "# "
      Else
         wc = wc & "And TPI.InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TPI.InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
      End If
   ElseIf Check2.Value = 1 Then
      If Gen_DBType = "MDB" Then
         wc = wc & "And TPI.FAGPostingDate >= #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And TPI.FAGPostingDate <= #" & Format(TxtToDate.Value, Gen_DatFormat) & "# "
      Else
         wc = wc & "And TPI.FAGPostingDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TPI.FAGPostingDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
      End If
   End If
   
   wc = wc & " And TPI.Flag='A'"
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
    wc = " Where  TPI.BatchID=" & Val(CmbBatch.Text) & ""
End If

Call Grid_Head


tmp = " Select TPI.PINo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode 'ClientNAVCode',MCSG.RGSTNavCode as 'RClientNAVCode',TPI.Amount,TPI.BatchID,MST.NAVCode, TPI.FAGPostingDate,TPI.FAGRemarks,TPI.IETypeID 'InvoiceTypeID',TPI.InvoiceDate,MDSG.GSTNavCode as DPNAVCode "
tmp = tmp & " ,MCom.HSNNo,MCom.AccountType ,MST.InvoiceType"
tmp = tmp & " From Txn_TempInvoice TPI"
tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_Commodity MCom On TPI.CommodityID = MCom.CommodityID"
tmp = tmp & " Left Join Mst_Client MC On TPI.ClientID = MC.ClientID"
tmp = tmp & " Left Join Mst_Depository MD On TPI.DPID = MD.DepositoryID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Left Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"
tmp = tmp & " Left Join Mst_DepositoryStateWiseGST MDSG on MD.DepositoryID = MDSG.DepositoryID And MSI.StateIndiaID = MDSG.StateIndiaID"
tmp = tmp & " Inner Join Mst_ServiceTax MST on MST.ServiceTaxID=TPI.ServiceTaxID"
tmp = tmp & wc & " Order By TPI.PIDate"


Call TmpTable
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
   If TmpRs!NavCode <> "GSTS00" Then
      CmbInvoiceType.Text = "Taxable"
   Else
      CmbInvoiceType.Text = "Non Taxable"
   End If
   
   Frame5.Caption = CmbInvoiceType.Text
   
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!FAGPostingDate, "mm/dd/yyyy") 'Format(TmpRs!InvoiceDate, "dd/mm/yyyy")
       MSHFlexGrid3.TextMatrix(I + 1, 1) = "Invoice"
       MSHFlexGrid3.TextMatrix(I + 1, 2) = "Customer"
       
       If TmpRs!InvoiceTypeID <> 8 And TmpRs!InvoiceTypeID <> 12 Then
          If TmpRs!InvoiceTypeID = 4 Or TmpRs!InvoiceTypeID = 11 Then
             MSHFlexGrid3.TextMatrix(I + 1, 3) = IIf(IsNull(TmpRs!RClientNAVCode), "", TmpRs!RClientNAVCode)
          Else
             MSHFlexGrid3.TextMatrix(I + 1, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
          End If
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 3) = IIf(IsNull(TmpRs!DPNAVCode), "", TmpRs!DPNAVCode)
       End If
       
       MSHFlexGrid3.TextMatrix(I + 1, 4) = TmpRs!Amount
       MSHFlexGrid3.TextMatrix(I + 1, 5) = "G/L Account"
       MSHFlexGrid3.TextMatrix(I + 1, 22) = IIf(IsNull(TmpRs!HSNNo), "", TmpRs!HSNNo)
       If CmbInvoiceType.Text = "Non Taxable" Then
          MSHFlexGrid3.TextMatrix(I + 1, 6) = GetGLAccNAVCodeFromInvoiceType(TmpRs!InvoiceTypeID, "AGRI", "G")
          If MSHFlexGrid3.TextMatrix(I + 1, 22) = "" Then
             MSHFlexGrid3.TextMatrix(I + 1, 22) = GetGLAccNAVCodeFromInvoiceType(TmpRs!InvoiceTypeID, "AGRI", "H")
          End If
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 6) = GetGLAccNAVCodeFromInvoiceType(TmpRs!InvoiceTypeID, "NAGRI", "G")
          If MSHFlexGrid3.TextMatrix(I + 1, 22) = "" Then
             MSHFlexGrid3.TextMatrix(I + 1, 22) = GetGLAccNAVCodeFromInvoiceType(TmpRs!InvoiceTypeID, "NAGRI", "H")
          End If
       End If
       
       If MSHFlexGrid3.TextMatrix(I + 1, 6) = "" Then ExpFlag = "N"
       If CmbInvoiceType.Text = "Non Taxable" Then
          MSHFlexGrid3.TextMatrix(I + 1, 12) = Trim(TmpRs!FAGRemarks) '"Being Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (Without ST) having PI No :- " & TmpRs!pino & " "
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 12) = Trim(TmpRs!FAGRemarks) '"Being Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (With ST) having PI No :- " & TmpRs!pino & " "
       End If
       MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!PINo), "", "AH" & TmpRs!PINo)
       
       MSHFlexGrid3.TextMatrix(I + 1, 13) = IIf(IsNull(TmpRs!LocatioNAVCode), "", TmpRs!LocatioNAVCode)
       MSHFlexGrid3.TextMatrix(I + 1, 14) = "CMP"
       
       If LCase(CmbInvoiceType) = "non taxable" Then
          MSHFlexGrid3.TextMatrix(I + 1, 20) = "AGRI"
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 20) = "NAGRI"
       End If
       
       MSHFlexGrid3.TextMatrix(I + 1, 21) = TmpRs!NavCode
       
       MSHFlexGrid3.TextMatrix(I + 1, 23) = IIf(IsNull(TmpRs!InvoiceType), "", TmpRs!InvoiceType)
       MSHFlexGrid3.TextMatrix(I + 1, 24) = "D"
       
       
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I + 1, 4))
       
       If IsNull(TmpRs!ClientNAVCode) = True Or TmpRs!ClientNAVCode = "" Then
          ExpFlag = "N"
       End If
       If IsNull(TmpRs!LocatioNAVCode) = True Or TmpRs!LocatioNAVCode = "" Then
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
   CmdExport.Enabled = True
Else
   CmdExport.Enabled = False
End If

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmbBatch.Text = ""
CmdExport.Enabled = False
CmdExcel.Enabled = True
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
mJournalType = ""
mJournalBatchName = ""
Call CmdReset_Click

tmp = "SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]"
tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where ME.EmployeeNo = '4474' And JCM.JournalTypeID = 2 "
tmp = tmp & " Order by JCM.JournalTypeID"
Call TmpTableHRPro

If TmpRsHRPro.RecordCount = 1 Then
   CmbJournalCode.Text = TmpRsHRPro!JournalCode
   mSplit = Split(CmbJournalCode.Text, "~")
   mJournalType = UCase(mSplit(0))
   mJournalBatchName = UCase(mSplit(1))
'   mSForwardToID = TmpRs!EmployeeID
Else
   CmbJournalCode.Text = ""
'   mSForwardToID = Null
End If



End Sub
Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    If Option1.Value = True Then
       .Cols = 26 '25 '22
    Else
       .Cols = 25 '24 '21
    End If
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Date"
    .TextMatrix(1, 1) = "Invoice"
    .TextMatrix(1, 2) = "Customer"
    .TextMatrix(1, 3) = "DPNavCode"
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
    .ColWidth(19) = 0
    .ColWidth(20) = 1400
    .ColWidth(21) = 1400
    .ColWidth(22) = 1400
    .ColWidth(23) = 1400
    If Option1.Value = True Then
       .ColWidth(24) = 1400
    End If
End With

End Sub

Private Sub MSHFlexGrid3_Click()

With MSHFlexGrid3
     If .Col = 25 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        .Col = 25
        If .TextMatrix(.row, .Col) = strUnChecked Then
           If .TextMatrix(.row, 3) <> "" And .TextMatrix(.row, 6) <> "" And .TextMatrix(.row, 13) <> "" Then
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
CmdExport.Enabled = False
CmbInvoiceType.Text = ""
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbInvoiceType.Enabled = False
   TxtFromDate.Enabled = False
   TxtToDate.Enabled = False
   Check1.Enabled = False
   Check2.Enabled = False
   ChkCheckAll.Enabled = False
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   CmbInvoiceType.Enabled = True
   TxtFromDate.Enabled = True
   TxtToDate.Enabled = True
   'TxtFromDate.Value = Date
   TxtFromDate.Value = GetFirstDateOfMonth(Date)
   TxtToDate.Value = GetLastDateOfMonth(Date)
   Check1.Enabled = True
   Check2.Enabled = True
   Check1.Value = 1
   ChkCheckAll.Enabled = True
   'CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
   'CmbJournalCode_GotFocus
   'CmbJournalCode_GotFocus
End If
End Sub


Private Function GetFirstDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_) - 1
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
If mSMonth > 0 Then
   GetFirstDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & "01")
Else
   GetFirstDateOfMonth = CDate(mSYear - 1 & "-12-" & "01")
End If
End Function



Private Sub TxtFromDate_LostFocus()
If Check1.Value = 1 Then
   TxtFromDate.Value = (Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01")
End If

End Sub

Private Sub TxtSearch_LostFocus()



For I = 2 To MSHFlexGrid3.Rows - 1
    MSHFlexGrid3.row = I
    MSHFlexGrid3.Col = 18
    If LCase(TxtSearch) = LCase(MSHFlexGrid3.TextMatrix(I, 18)) Then
       MSHFlexGrid3.TopRow = I
       Exit Sub
    End If
Next

End Sub

Private Sub TxtToDate_LostFocus()
If Check1.Value = 1 Then
   TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
End If
End Sub
Private Function GetLastDateOfMonth(mDate_ As Date)

mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function

