VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ExportNonNCDEXNonCorporateInvoice 
   Caption         =   "Export Non NCDEX/Non Corporate Invoice"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11100
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstLocationName 
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
      Height          =   1110
      ItemData        =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":0000
      Left            =   2400
      List            =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":0007
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
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
         Left            =   7320
         TabIndex        =   28
         Top             =   1080
         Width           =   2880
      End
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":001C
         Left            =   9375
         List            =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":0026
         TabIndex        =   26
         Top             =   660
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.PictureBox picCombo 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2280
         ScaleHeight     =   300
         ScaleWidth      =   3555
         TabIndex        =   22
         Top             =   720
         Width           =   3615
         Begin VB.CommandButton cmdDropDown 
            Height          =   375
            Left            =   3240
            Picture         =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":003E
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   -30
            Width           =   330
         End
         Begin VB.Label lblSelections 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   90
            TabIndex        =   24
            Top             =   30
            Width           =   45
         End
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
         Left            =   10920
         TabIndex        =   21
         Top             =   675
         Width           =   1605
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
         Height          =   375
         Left            =   13680
         TabIndex        =   20
         Top             =   1080
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
         Height          =   350
         Left            =   12600
         TabIndex        =   19
         Top             =   675
         Width           =   1140
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
         ItemData        =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":0188
         Left            =   960
         List            =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":018A
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   390
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
         ItemData        =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":018C
         Left            =   7320
         List            =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":0196
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2895
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
         Left            =   7320
         TabIndex        =   9
         Top             =   675
         Width           =   1320
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
         Left            =   90
         TabIndex        =   8
         Top             =   750
         Width           =   1095
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
         Left            =   90
         TabIndex        =   7
         Top             =   285
         Value           =   -1  'True
         Width           =   1095
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
         Left            =   11460
         TabIndex        =   6
         Top             =   240
         Width           =   1140
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
         Left            =   12600
         TabIndex        =   5
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
         Left            =   13740
         TabIndex        =   4
         Top             =   240
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
         Left            =   10320
         TabIndex        =   3
         Top             =   240
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   2310
         TabIndex        =   13
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
         Left            =   4605
         TabIndex        =   14
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
         Left            =   6000
         TabIndex        =   29
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Status"
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
         Left            =   8700
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label LblCMPName 
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
         Height          =   210
         Left            =   1245
         TabIndex        =   18
         Top             =   795
         Width           =   900
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
         Left            =   1245
         TabIndex        =   16
         Top             =   315
         Width           =   990
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
         Left            =   3750
         TabIndex        =   15
         Top             =   300
         Width           =   810
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
         Left            =   6000
         TabIndex        =   12
         Top             =   270
         Width           =   1455
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
         Left            =   6000
         TabIndex        =   11
         Top             =   735
         Width           =   1095
      End
   End
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
      Height          =   7665
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7320
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12912
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "Frm_ExportNonNCDEXNonCorporateInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCode, mSTCode As Variant
Dim ExpFlag, MaxID As Variant
Dim mSMonth, mSYear, mSTotalDays As Variant
Dim mSM_Prefix As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Private mblnListLostFocus   As Boolean
Private Const mstr_DELIMITER = ","
Dim mJournalBatchName, mJournalType, mIncomeTypeID As Variant

Private Sub ChkAllLocation_Click()
If ChkAllLocation.Value = vbChecked Then
   For I = 0 To lstLocationName.ListCount - 1
       lstLocationName.Selected(I) = True
   Next
ElseIf ChkAllLocation.Value = False Then
   For I = 0 To lstLocationName.ListCount - 1
       lstLocationName.Selected(I) = False
   Next
End If
End Sub

Private Sub CmbBatch_GotFocus()
CmbInvoiceType.Text = ""
tmp = "select distinct BatchID As BatchID from Txn_ProInvoice Where BatchID > 0  order by BatchID"
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

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "confirmed" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportNNNRInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportNNNRInvoice.xls")
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
mUserBatchID = GetGUID
'mJournalBatchName = GetJournalDetailsFromHRPro(Gen_UserLoginID, mJournalType)
'If Trim(mJournalBatchName) = "" Then
'   MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'   Exit Sub
'End If
'-- On 05 Aug 2016

'Commented on 12th Aug 2017 during GST Implementaion
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
   strFilName = "\Navision\NavUpload_NonNCDEXNonCorporate" & "_" & CmbInvoiceType.Text & "_" & MaxID & " .csv"
Else
   strFilName = "\Navision\NavUpload_NonNCDEXNonCorporate" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
   Call TableTxn_InvoicesForNAVISION(" Where G_UID Is Null ")
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
End If

If Option1.Value = True Then
   If CmbInvoiceType.Text = "Non Taxable" Then
      tmp = "Select * From Txn_ProInvoice Where Isnull(BatchID,0) = 0  And ServiceTaxNAVCode = 'GSTS00' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'  "
   Else
      tmp = "Select * From Txn_ProInvoice Where Isnull(BatchID,0) = 0 And ServiceTaxNAVCode  = 'GSTS18' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
   End If
   Call Tmp1Table
End If

With MSHFlexGrid3
For I = 2 To MSHFlexGrid3.Rows - 1
    tmp = ""
    If Option1.Value = True Then
       If .TextMatrix(I, 25) = strChecked Then
       
'Commented on 12 Aug 2017 during GST Implementaion
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
            RsTxn_InvoicesForNAVISION!InvoiceType = "NNNR"
            RsTxn_InvoicesForNAVISION!JournalBatchType = Trim(mJournalType)
            RsTxn_InvoicesForNAVISION!JournalBatchName = Trim(mJournalBatchName)
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
               TmpRs1.Find "InvoiceNo = " & Val(Replace(.TextMatrix(I, 18), "NN", "")) & " "
               If TmpRs1.EOF = False Then
                  TmpRs1!BatchID = MaxID
                  TmpRs1!FAGPostingDate = Date
                  TmpRs1!Flag = "C"
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
    
'    For c = 0 To MSHFlexGrid3.Cols - 1
'        If tmp = "" Then
'           tmp = tmp & MSHFlexGrid3.TextMatrix(I, c)
'        Else
'           tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, c)
'        End If
'    Next
'    Print #1, tmp
Next
End With
'If Option1.Value = True Then
'
''   If CmbInvoiceType.Text = "Non Taxable" Then
''      tmp = "Update Txn_ProInvoice set BatchID = " & MaxID & ",FAGPostingDate = '" & Date & "' Where BatchID = 0  And ServiceTaxNAVCode = '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
''   Else
''      tmp = "Update Txn_ProInvoice set BatchID = " & MaxID & ",FAGPostingDate = '" & Date & "' Where BatchID = 0 And ServiceTaxNAVCode <> '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
''   End If
'
'   If CmbInvoiceType.Text = "Non Taxable" Then
'      If CmbLocationID.Text = "" Then
'         tmp = "Update Txn_ProInvoice set BatchID = " & MaxID & ",FAGPostingDate = '" & Date & "' Where BatchID = 0  And ServiceTaxNAVCode = '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
'      Else
'         tmp = "Update Txn_ProInvoice set BatchID = " & MaxID & ",FAGPostingDate = '" & Date & "' Where BatchID = 0  And ServiceTaxNAVCode = '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' And SM_Prefix = '" & mSM_Prefix & "' "
'      End If
'   Else
'      If CmbLocationID.Text = "" Then
'         tmp = "Update Txn_ProInvoice set BatchID = " & MaxID & ",FAGPostingDate = '" & Date & "' Where BatchID = 0 And ServiceTaxNAVCode <> '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
'      Else
'         tmp = "Update Txn_ProInvoice set BatchID = " & MaxID & ",FAGPostingDate = '" & Date & "' Where BatchID = 0 And ServiceTaxNAVCode <> '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'  And SM_Prefix = '" & mSM_Prefix & "' "
'      End If
'   End If
'
'   Call TmpTable
'End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
End Sub

Private Function GetMaxBatchID()
tmp = "select Max(BatchID) from Txn_ProInvoice"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   GetMaxBatchID = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0)) + 1
Else
   GetMaxBatchID = 1
End If
End Function
Private Sub CmdGetReport_Click()

Dim wc As Variant
Dim StrText     As Variant
Dim astrTemp()  As String
Dim lngX        As Long
Dim mStartDate As Date
ExpFlag = ""

mStartDate = (Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01")

If Option1.Value = True Then
   If CmbInvoiceType.Text = "Non Taxable" Then
      mGLAccNAVCode = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "G")
      wc = " Where  TPI.ServiceTaxNAVCode = 'GSTS00' And Isnull(TPI.BatchID,0) = 0 And TPI.Flag = 'C' " 'TPI.ServiceTaxNAVCode = '' And TPI.BatchID = 0 And TPI.Flag = 'C' "
   ElseIf CmbInvoiceType.Text = "Taxable" Then
      mGLAccNAVCode = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "G")
      wc = " Where  TPI.ServiceTaxNAVCode = 'GSTS18' And Isnull(TPI.BatchID,0) = 0 And TPI.Flag = 'C' " 'TPI.ServiceTaxNAVCode <> '' And TPI.BatchID = 0 And TPI.Flag = 'C' "
   End If
End If

If Option1.Value = True Then
   
   If TxtFromDate.Value > TxtToDate.Value Then
      MsgBox "'To Date' must be greater than 'From Date' always!!!"
      TxtToDate.SetFocus
      Exit Sub
   End If
   If CmbInvoiceType.Text = "" Then
      MsgBox "Select Invoice Type."
      CmbInvoiceType.SetFocus
      Exit Sub
   End If
   
   
   If Gen_DBType = "MDB" Then
      wc = wc & " And TPI.InvoiceDate >= #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And TPI.InvoiceDate <= #" & Format(TxtToDate.Value, Gen_DatFormat) & "# "
   Else
      wc = wc & " And TPI.InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TPI.InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
   End If

'   If CmbLocationID.Text <> "" Then
'      Wc = Wc & " And  TPI.SM_Prefix = '" & mSM_Prefix & "'"
'   End If
   If lblSelections.Caption <> "" Then
      astrTemp = Split(Trim(lblSelections.Caption), ",")
      For lngX = 0 To UBound(astrTemp)
          astrTemp(lngX) = "'" & Replace$(astrTemp(lngX), "'", "''") & "'"
      Next
      StrText = Join(astrTemp, ",")
      
      wc = wc & " And ML.LocationName in ( " & StrText & " ) "
   End If
   
'   If LCase(CmbStatus.Text) = "pending" Then
'      wc = wc & " And  TPI.Flag = 'P'"
'   ElseIf LCase(CmbStatus.Text) = "confirmed" Then
'      wc = wc & " And  TPI.Flag = 'C'"
'   End If
   
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
   wc = " Where  TPI.BatchID=" & Val(CmbBatch.Text) & " "
End If

Call Grid_Head


'Commented on 12 Aug 2017 During Implementation of GST Effect
'tmp = "Select Distinct TPI.InvoiceNo,TPI.SM_Prefix,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MC.NAVCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,TPI.ServiceTaxNAVCode,"
'tmp = tmp & " TPI.InvoiceDate , TPI.Discount, TPI.Premium, TPI.FAGPostingDate"
'tmp = tmp & " From Txn_ProInvoice TPI"
'tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
'tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientIDRow = MC.ClientIDRow"
'''tmp = tmp & " Inner Join Txn_InvMonGSLDetail TIMGD On TPI.InvoiceNo = TIMGD.PINo "
'tmp = tmp & wc & " And TPI.Flag = 'C' Order By TPI.InvoiceDate"



tmp = "Select Distinct TPI.InvoiceNo,TPI.SM_Prefix,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,TPI.ServiceTaxNAVCode,"
tmp = tmp & " TPI.InvoiceDate , TPI.Discount, TPI.Premium, TPI.FAGPostingDate,TPI.HSNNo ,TPI.AccountType,MST.InvoiceType"
tmp = tmp & " From Txn_ProInvoice TPI"
tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"
tmp = tmp & " Inner Join Mst_ServiceTax MST on TPI.ServiceTaxNAVCode = MST.NavCode"
tmp = tmp & wc & " And TPI.Flag = 'C' Order By TPI.InvoiceDate"


Call TmpTable

'Commented on 09 Nov 2012
'tmp = "Select distinct TIG.PINo from Txn_InvMonGSLDetail TIG"

tmp = "Select TIG.PINo, sum(InvoiceAmount) GSLAmount from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner join Mst_CMP MC on TIG.CMPID=MC.CMPID"
tmp = tmp & " Inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where TIG.PINo in ( Select Distinct InvoiceNo From Txn_ProInvoice TPI"
tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
tmp = tmp & wc & " And TPI.Flag = 'C') "
If lblSelections.Caption <> "" Then
   astrTemp = Split(Trim(lblSelections.Caption), ",")
   For lngX = 0 To UBound(astrTemp)
       astrTemp(lngX) = "'" & Replace$(astrTemp(lngX), "'", "''") & "'"
   Next
   StrText = Join(astrTemp, ",")
   
   tmp = tmp & " And ML.LocationName in ( " & StrText & " ) "
End If

If Option1.Value = True Then
   If Gen_DBType = "MDB" Then
      tmp = tmp & " And YearMonth >= #" & Format(mStartDate, Gen_DatFormat) & "# And YearMonth<= #" & Format(TxtToDate.Value, Gen_DatFormat) & "# "
   Else
      tmp = tmp & " And YearMonth >= '" & Format(mStartDate, Gen_DatFormat) & "' And YearMonth<= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
   End If
End If

tmp = tmp & " Group by TIG.PINo"

Call Tmp1Table

If TmpRs.RecordCount > 0 Then
   
   If TmpRs!ServiceTaxNAVCode = "GSTS18" Then
      CmbInvoiceType.Text = "Taxable"
   Else
      CmbInvoiceType.Text = "Non Taxable"
   End If
   If Option2.Value = True Then
      If CmbInvoiceType.Text = "Non Taxable" Then
         mGLAccNAVCode = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI", "G")
      ElseIf CmbInvoiceType.Text = "Taxable" Then
         mGLAccNAVCode = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "NAGRI", "G")
      End If
   End If
   Frame5.Caption = CmbInvoiceType.Text
   I = 2
   For x = 1 To TmpRs.RecordCount
       If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
       TmpRs1.Find "PINo=" & TmpRs!InvoiceNo
       If TmpRs1.EOF = False Then
          'If Format(TmpRs!Amount, "#######") <> Format(TmpRs1!GSLAmount, "#######") Then GoTo nextrow
          If Abs(TmpRs!Amount - TmpRs1!GSLAmount) > 1 Then GoTo NextRow
          I = MSHFlexGrid3.Rows - 2
          If Option2.Value = True Then
             MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!InvoiceDate, "mm/dd/yyyy") ', "dd/mm/yyyy")
          Else
             MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!InvoiceDate, "mm/dd/yyyy") 'MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!InvoiceDate + 1, "mm/dd/yyyy") ', "dd/mm/yyyy")
          End If
          MSHFlexGrid3.TextMatrix(I + 1, 1) = "Invoice"
          MSHFlexGrid3.TextMatrix(I + 1, 2) = "Customer"
          MSHFlexGrid3.TextMatrix(I + 1, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
          
          'If TmpRs!Discount <> 0 And IsNull(TmpRs!Discount) = Flase Then
          MSHFlexGrid3.TextMatrix(I + 1, 4) = Round(TmpRs!Amount + TmpRs!Premium - IIf(IsNull(TmpRs!Discount), 0, TmpRs!Discount))
          'ElseIf TmpRs!Premium <> 0 And IsNull(TmpRs!Premium) = Flase Then
          '   MSHFlexGrid3.TextMatrix(i + 1, 4) = TmpRs!Amount + TmpRs!Premium
          'Else
          '   MSHFlexGrid3.TextMatrix(i + 1, 4) = TmpRs!Amount
          'End If
          MSHFlexGrid3.TextMatrix(I + 1, 5) = "G/L Account"
          MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCode

          If CmbInvoiceType.Text = "Non Taxable" Then
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being Invoice for " & TmpRs!Locationname & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (Without GST) having PI No :- NN" & TmpRs!InvoiceNo & " "
          Else
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being Invoice for " & TmpRs!Locationname & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (With GST) having PI No :- NN" & TmpRs!InvoiceNo & " "
          End If
          MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!InvoiceNo), "", "NN" & TmpRs!InvoiceNo)
          
          MSHFlexGrid3.TextMatrix(I + 1, 13) = IIf(IsNull(TmpRs!LocatioNAVCode), "", TmpRs!LocatioNAVCode)
          MSHFlexGrid3.TextMatrix(I + 1, 14) = "CMP"
'          If LCase(CmbInvoiceType) = "non taxable" Then
'             MSHFlexGrid3.TextMatrix(I + 1, 20) = "AGRI"
'          Else
'             MSHFlexGrid3.TextMatrix(I + 1, 20) = "NAGRI"
'          End If
          MSHFlexGrid3.TextMatrix(I + 1, 20) = IIf(IsNull(TmpRs!AccountType), "", TmpRs!AccountType)
          MSHFlexGrid3.TextMatrix(I + 1, 21) = TmpRs!ServiceTaxNAVCode
          MSHFlexGrid3.TextMatrix(I + 1, 22) = IIf(IsNull(TmpRs!HSNNo), "", TmpRs!HSNNo)
          MSHFlexGrid3.TextMatrix(I + 1, 23) = IIf(IsNull(TmpRs!InvoiceType), "", TmpRs!InvoiceType)
          MSHFlexGrid3.TextMatrix(I + 1, 24) = "M"
          
          MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I + 1, 4))
          
          If IsNull(TmpRs!ClientNAVCode) = True Or TmpRs!ClientNAVCode = "" Then
             ExpFlag = "N"
          End If
          If IsNull(TmpRs!LocatioNAVCode) = True Or TmpRs!LocatioNAVCode = "" Then
             ExpFlag = "N"
          End If
          If IsNull(mGLAccNAVCode) = True Or mGLAccNAVCode = "" Then
             ExpFlag = "N"
          End If
          
   '       If mGLAccNAVCode = "" Then
   '          ExpFlag = "N"
   '       End If
          
          If Option1.Value = True Then
             MSHFlexGrid3.row = I + 1
             MSHFlexGrid3.Col = 25 '21
             MSHFlexGrid3.CellFontName = "Wingdings"
             MSHFlexGrid3.CellFontSize = 12
             MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
             MSHFlexGrid3.Text = strUnChecked
           End If
           MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       End If
NextRow:
       TmpRs.MoveNext
   Next
'   If ExpFlag = "N" Then
'      CmdExport.Enabled = False
'   Else
'      CmdExport.Enabled = True
'   End If
'Else
'   CmdExport.Enabled = False
End If

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmbBatch.Text = ""
CmbInvoiceType.Text = ""
CmdExport.Enabled = True
CmbStatus.Text = ""
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
mIncomeTypeID = 7
Frame1.Enabled = True
Frame5.Enabled = True
Call Grid_Head
Call CmdReset_Click

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
'    .Cols = 21
    
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
'CmdExport.Enabled = False
CmbInvoiceType.Text = ""
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbStatus.Text = ""
   CmbStatus.Enabled = False
   CmbInvoiceType.Enabled = False
   TxtFromDate.Enabled = False
   TxtToDate.Enabled = False
   CmbLocationID.Enabled = False
   cmdDropDown.Enabled = False
   lstLocationName.Visible = False
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = False
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   CmbInvoiceType.Enabled = True
   TxtFromDate.Enabled = True
   TxtToDate.Enabled = True
   TxtFromDate.Value = GetFirstDateOfMonth(Date)
   TxtToDate.Value = Date
   CmbLocationID.Enabled = True
   cmdDropDown.Enabled = True
   lstLocationName.Visible = False
   CmbStatus.Enabled = True
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = True
   'CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub

Private Sub TxtFromDate_LostFocus()
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
End Sub

Private Sub TxtSearch_LostFocus()
'
'For I = 2 To MSHFlexGrid3.Rows - 1
'    If LCase(TxtSearch) = LCase(MSHFlexGrid3.TextMatrix(I, 18)) Then
'       MSHFlexGrid3.RowSel = I
'       MSHFlexGrid3.Row = MSHFlexGrid3.RowSel
'       MSHFlexGrid3.ColSel = 21
'       'MSHFlexGrid3.RowPosition (I)
'       Exit For
'    End If
'Next


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
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
End Sub
Private Function GetLastDateOfMonth(mDate_ As Date)

mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function


Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub ChkCheckAll_Click()

If MSHFlexGrid3.Rows = 2 Then
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


Private Sub cmdDropDown_Click()
If (lstLocationName.Visible = True) Or (mblnListLostFocus) Then
   lstLocationName.Visible = False
   mblnListLostFocus = False
Else
   Call GetLocationNameforList
   lstLocationName.Visible = True
   lstLocationName.SetFocus
End If
End Sub

Private Sub lstLocationName_LostFocus()
lstLocationName.Visible = False
If Screen.ActiveControl.Name = "cmdDropDown" Then
   mblnListLostFocus = True
Else
   mblnListLostFocus = False
End If
End Sub
Private Sub lstLocationName_Click()

Dim intX        As Integer
Dim blnFirst    As Boolean
'Dim IntCount    As Integer

'IntCount = 0
lblSelections.Caption = ""
lblSelections.Tag = ""
lblSelections.ToolTipText = ""
blnFirst = True

If lstLocationName.SelCount = 0 Then Exit Sub

For intX = 0 To lstLocationName.ListCount - 1
    If lstLocationName.Selected(intX) Then
       If blnFirst Then
          blnFirst = False
       Else
          lblSelections.Caption = lblSelections.Caption & mstr_DELIMITER
          lblSelections.Tag = lblSelections.Tag & mstr_DELIMITER
       End If
'      IntCount = IntCount + 1
       lblSelections.Caption = lblSelections.Caption & Trim(lstLocationName.List(intX))
       lblSelections.Tag = lblSelections.Tag & lstLocationName.ItemData(intX)
    End If
Next

lblSelections.ToolTipText = lblSelections.Caption
End Sub

Private Sub GetLocationNameforList()

lblSelections.Caption = ""
lstLocationName.Clear


tmp = " Select * from Mst_Location Order By LocationName"
Call Tmp2Table


If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       lstLocationName.AddItem TmpRs2!Locationname
       lstLocationName.ItemData(lstLocationName.NewIndex) = I
       TmpRs2.MoveNext
   Next
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

