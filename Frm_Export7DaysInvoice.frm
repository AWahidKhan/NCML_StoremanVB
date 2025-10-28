VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Export7DaysInvoice 
   Caption         =   "Export 7 Days Invoice"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   11565
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
      Height          =   7425
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   6960
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12277
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
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
         Left            =   7440
         TabIndex        =   29
         Top             =   1200
         Width           =   2760
      End
      Begin VB.Frame Frame2 
         Height          =   540
         Left            =   3840
         TabIndex        =   23
         Top             =   630
         Width           =   11025
         Begin VB.OptionButton OptPWR 
            Caption         =   "PWR"
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
            Left            =   6840
            TabIndex        =   28
            Top             =   150
            Width           =   1935
         End
         Begin VB.OptionButton OptYearly 
            Caption         =   "NNNR Yearly"
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
            Left            =   4800
            TabIndex        =   27
            Top             =   150
            Width           =   1935
         End
         Begin VB.OptionButton OptRRN 
            Caption         =   "RRN Withdrawal"
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
            Left            =   2760
            TabIndex        =   26
            Top             =   150
            Width           =   1935
         End
         Begin VB.OptionButton OptRej 
            Caption         =   "Rejected"
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
            Left            =   1245
            TabIndex        =   25
            Top             =   150
            Width           =   1215
         End
         Begin VB.OptionButton Opt7Days 
            Caption         =   "7 Days"
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
            Left            =   105
            TabIndex        =   24
            Top             =   150
            Value           =   -1  'True
            Width           =   1095
         End
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   240
         Width           =   1140
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
         TabIndex        =   10
         Top             =   240
         Width           =   1140
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
         TabIndex        =   9
         Top             =   285
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
         Height          =   300
         Left            =   90
         TabIndex        =   8
         Top             =   750
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
         Left            =   2400
         TabIndex        =   7
         Top             =   675
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
         ItemData        =   "Frm_Export7DaysInvoice.frx":0000
         Left            =   7440
         List            =   "Frm_Export7DaysInvoice.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   2775
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
         ItemData        =   "Frm_Export7DaysInvoice.frx":0024
         Left            =   240
         List            =   "Frm_Export7DaysInvoice.frx":0026
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   390
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
         TabIndex        =   4
         Top             =   1275
         Width           =   900
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
         Height          =   495
         Left            =   13590
         TabIndex        =   3
         Top             =   1203
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
         Left            =   10440
         TabIndex        =   2
         Top             =   1275
         Width           =   2085
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
         ItemData        =   "Frm_Export7DaysInvoice.frx":0028
         Left            =   9495
         List            =   "Frm_Export7DaysInvoice.frx":0032
         TabIndex        =   1
         Top             =   675
         Visible         =   0   'False
         Width           =   1440
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   2430
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
         Format          =   113901569
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   4725
         TabIndex        =   15
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
         Format          =   113901569
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
         Left            =   6120
         TabIndex        =   30
         Top             =   1260
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
         Left            =   1365
         TabIndex        =   20
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
         Height          =   255
         Left            =   6120
         TabIndex        =   19
         Top             =   270
         Width           =   1455
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
         Left            =   3870
         TabIndex        =   18
         Top             =   300
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
         Left            =   1365
         TabIndex        =   17
         Top             =   315
         Width           =   990
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
         Left            =   8820
         TabIndex        =   16
         Top             =   728
         Visible         =   0   'False
         Width           =   615
      End
   End
End
Attribute VB_Name = "Frm_Export7DaysInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCode, mGLAccNAVCodeRej, mSTCode, mGLAccNAVCodeRRN, mGLAccNAVCodeYearly As Variant
Dim ExpFlag, MaxID As Variant
Dim mSMonth, mSYear, mSTotalDays As Variant
Dim mSM_Prefix As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Private mblnListLostFocus   As Boolean
Private Const mstr_DELIMITER = ","
Dim mJournalBatchName, mJournalType, mInvoiceType As Variant

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
tmp = "select distinct BatchID As BatchID from Txn_CDTFBOE Where BatchID > 0  order by BatchID"
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
''Private Sub CmbStatus_LostFocus()
''If CmbStatus.Text = "" Then Exit Sub
''
''If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "confirmed" Then
''   MsgBox "Invalid Selection!!!"
''   CmbStatus.SetFocus
''   Exit Sub
''End If
''
''End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportNCDEXInvoices.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportNCDEXInvoices.xls")
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

''-- On 05 Aug 2016
mInvoiceType = ""
If Opt7Days.Value = True Then
    mInvoiceType = "7Days"
ElseIf OptRej.Value = True Then
    mInvoiceType = "Rejected"
ElseIf OptRRN.Value = True Then
    mInvoiceType = "RRN"
ElseIf OptYearly.Value = True Then
    mInvoiceType = "NNNRY"
ElseIf OptPWR.Value = True Then
    mInvoiceType = "PWR"
End If

Path = App.Path
MaxID = GetMaxBatchID
If Opt7Days.Value = True Then
   If Option1.Value = True Then
      strFilName = "\Navision\NavUpload_7DaysInvoice" & "_" & CmbInvoiceType.Text & "_" & MaxID & ".csv"
   Else
      strFilName = "\Navision\NavUpload_7DaysInvoice" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
   End If
ElseIf OptRRN.Value = True Then
   If Option1.Value = True Then
      strFilName = "\Navision\NavUpload_RRNInvoice" & "_" & CmbInvoiceType.Text & "_" & MaxID & ".csv"
   Else
      strFilName = "\Navision\NavUpload_RRNInvoice" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
   End If
ElseIf OptYearly.Value = True Then
   If Option1.Value = True Then
      strFilName = "\Navision\NavUpload_NNNRYearlyInvoice" & "_" & CmbInvoiceType.Text & "_" & MaxID & ".csv"
   Else
      strFilName = "\Navision\NavUpload_NNNRYearlyInvoice" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
   End If
ElseIf OptPWR.Value = True Then
   If Option1.Value = True Then
      strFilName = "\Navision\NavUpload_PWRInvoice" & "_" & CmbInvoiceType.Text & "_" & MaxID & ".csv"
   Else
      strFilName = "\Navision\NavUpload_PWRInvoice" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
   End If
Else
   If Option1.Value = True Then
      strFilName = "\Navision\NavUpload_RejectedInvoice" & "_" & CmbInvoiceType.Text & "_" & MaxID & ".csv"
   Else
      strFilName = "\Navision\NavUpload_RejectedInvoice" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
   End If
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
   'Commented on 14 Aug 2017 during GST Effect implementaion
    Call TableTxn_InvoicesForNAVISION(" Where G_UID Is Null ")
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
End If

If Option1.Value = True Then
   If CmbInvoiceType.Text = "Non Taxable" Then
      ''tmp = "Select * From Txn_ProInvoice Where BatchID = 0  And ServiceTaxNAVCode = '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'  "
      tmp = "Select * From Txn_CDTFBOE Where BOEID in ("
      tmp = tmp & " Select Distinct TPI.BOEID From Txn_CDTFBOE TPI"
      If Opt7Days.Value = True Then
         tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TPI.TxnType = 'D' "
      ElseIf OptRRN.Value = True Then
         tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TPI.TxnType = 'R' "
      ElseIf OptYearly.Value = True Then
         tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TPI.TxnType = 'Y' "
      ElseIf OptPWR.Value = True Then
         tmp = tmp & " Inner Join Txn_RRNPWR TC On TPI.CxTF_DetailsID = TC.ConvertedID and TPI.TxnType = 'P' "
      Else
         tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TPI.TxnType = 'W' "
      End If
      tmp = tmp & " inner join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
      tmp = tmp & " inner join Mst_ServiceTax MS On MComm.ServiceTaxID = MS.ServiceTaxID "
      tmp = tmp & " Where TPI.Flag='C' And Isnull(TPI.BatchID,0) = 0  "
      If Opt7Days.Value = True Then
         tmp = tmp & " And TPI.TxnType='D'"
      ElseIf OptRRN.Value = True Then
         tmp = tmp & " And TPI.TxnType='R'"
      ElseIf OptYearly.Value = True Then
         tmp = tmp & " And TPI.TxnType='Y'"
      ElseIf OptPWR.Value = True Then
         tmp = tmp & " And TPI.TxnType='P'"
      Else
         tmp = tmp & " And TPI.TxnType='W'"
      End If
      
      tmp = tmp & " And MS.NAVCode = 'GSTS00' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"

   Else
      ''tmp = "Select * From Txn_ProInvoice Where BatchID = 0 And ServiceTaxNAVCode  <> '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
      tmp = "Select * From Txn_CDTFBOE Where BOEID in ("
      tmp = tmp & " Select Distinct TPI.BOEID From Txn_CDTFBOE TPI"
'      tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CDTFNo = TC.CxTFNo And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType=TPI.TxnType "
      If Opt7Days.Value = True Then
         tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TPI.TxnType = 'D' "
      ElseIf OptRRN.Value = True Then
         tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TPI.TxnType = 'R' "
      ElseIf OptYearly.Value = True Then
         tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TPI.TxnType = 'Y' "
      ElseIf OptPWR.Value = True Then
         tmp = tmp & " Inner Join Txn_RRNPWR TC On TPI.CxTF_DetailsID = TC.ConvertedID And TPI.TxnType = 'P' "
      Else
         tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TPI.TxnType = 'W' "
      End If
      
      tmp = tmp & " inner join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
      tmp = tmp & " inner join Mst_ServiceTax MS On MComm.ServiceTaxID = MS.ServiceTaxID "
      tmp = tmp & " Where TPI.Flag='C' And Isnull(TPI.BatchID,0) = 0  And MS.NAVCode = 'GSTS18'  "
      If Opt7Days.Value = True Then
         tmp = tmp & " And TPI.TxnType='D'"
      ElseIf OptRRN.Value = True Then
         tmp = tmp & " And TPI.TxnType='R'"
      ElseIf OptYearly.Value = True Then
         tmp = tmp & " And TPI.TxnType='Y'"
      ElseIf OptPWR.Value = True Then
         tmp = tmp & " And TPI.TxnType='P'"
      Else
         tmp = tmp & " And TPI.TxnType='W'"
      End If
      tmp = tmp & " And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
   End If
   Call Tmp1Table
End If

For I = 2 To MSHFlexGrid3.Rows - 1
    tmp = ""
    If Option1.Value = True Then
       
       If MSHFlexGrid3.TextMatrix(I, 25) = strChecked Then
          'Commented on 14 Aug 2017 during GST Effect implementaion
          With MSHFlexGrid3
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
'            RsTxn_InvoicesForNAVISION!TINo = Null
'            RsTxn_InvoicesForNAVISION!TIDate = Null
'            RsTxn_InvoicesForNAVISION!STAmt = Null
'            RsTxn_InvoicesForNAVISION!TIAmt = Null
'            RsTxn_InvoicesForNAVISION!PostInvoiceNo = Null
'            RsTxn_InvoicesForNAVISION!EntryNo = Null
            
            RsTxn_InvoicesForNAVISION!GSTGroupCode = Trim(.TextMatrix(I, 21))
            RsTxn_InvoicesForNAVISION!HSNNo = Trim(.TextMatrix(I, 22))
            RsTxn_InvoicesForNAVISION!AccountType = Trim(.TextMatrix(I, 23))
            RsTxn_InvoicesForNAVISION!NumberSeriesFlag = Trim(.TextMatrix(I, 24))
            
            RsTxn_InvoicesForNAVISION!CreatedBy = Gen_UserLoginID
            RsTxn_InvoicesForNAVISION!G_UID = GetGUID
            RsTxn_InvoicesForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
            
            RsTxn_InvoicesForNAVISION!InvoiceType = mInvoiceType
            RsTxn_InvoicesForNAVISION!JournalBatchType = Trim(mJournalType)
            RsTxn_InvoicesForNAVISION!JournalBatchName = Trim(mJournalBatchName)
            'RsTxn_InvoicesForNAVISION!ReverseInvoice = Null
            'RsTxn_InvoicesForNAVISION!EntryNoTransfer = Null
            'RsTxn_InvoicesForNAVISION!EntryNoTransferRemark = Null
            RsTxn_InvoicesForNAVISION!UserBatchID = mUserBatchID
            
            'RsTxn_InvoicesForNAVISION!DataTransfer = "X"
            
            RsTxn_InvoicesForNAVISION.Update
          End With

'          For c = 0 To MSHFlexGrid3.Cols - 2
'              If tmp = "" Then
'                 tmp = tmp & MSHFlexGrid3.TextMatrix(i, c)
'              Else
'                 tmp = tmp & "," & MSHFlexGrid3.TextMatrix(i, c)
'              End If
'          Next
'          Print #1, tmp
          
          If Option1.Value = True Then
            If TmpRs1.RecordCount > 0 Then
               TmpRs1.MoveFirst
               If Opt7Days.Value = True Then
                  TmpRs1.Find "BOEID = " & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "SD", "")) & " "
               ElseIf OptRRN.Value = True Then
                  TmpRs1.Find "BOEID = " & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "RRN", "")) & " "
               ElseIf OptYearly.Value = True Then
                  TmpRs1.Find "BOEID = " & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "NNNRY", "")) & " "
               ElseIf OptPWR.Value = True Then
                  TmpRs1.Find "BOEID = " & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "PWR", "")) & " "
               Else
                  TmpRs1.Find "BOEID = " & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "RJ", "")) & " "
               End If
               If TmpRs1.EOF = False Then
                  TmpRs1!BatchID = MaxID
                  TmpRs1!FAGPostingDate = Date
                  ''-- Ullhas Patil 09 Nov 2017 -- Save the GSTGroupCode, HSN Code, Invoice Type
                    TmpRs1!GSTGroupCode = MSHFlexGrid3.TextMatrix(I, 21)
                    TmpRs1!HSNNo = MSHFlexGrid3.TextMatrix(I, 22)
                    TmpRs1!AccountType = MSHFlexGrid3.TextMatrix(I, 23)
                  ''----------------------------
                  TmpRs1.Update
               End If
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
    
    
'    For c = 0 To MSHFlexGrid3.Cols - 1
'        If tmp = "" Then
'           tmp = tmp & MSHFlexGrid3.TextMatrix(I, c)
'        Else
'           tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, c)
'        End If
'    Next
'    Print #1, tmp
Next

MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Close #1
Call Grid_Head
End Sub
Private Function GetMaxBatchID()
tmp = "select Max(BatchID) from Txn_CDTFBOE"
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

ExpFlag = ""
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
   
   If CmbInvoiceType.Text = "Non Taxable" Then
      wc = " Where  MST.NAVCode = 'GSTS00' And Isnull(TPI.BatchID,0) = 0"
   ElseIf CmbInvoiceType.Text = "Taxable" Then
      wc = " Where  MST.NAVCode = 'GSTS18' And Isnull(TPI.BatchID,0) = 0"
   End If
   
   If Gen_DBType = "MDB" Then
      wc = wc & " And TPI.InvoiceDate >= #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And TPI.InvoiceDate <= #" & Format(TxtToDate.Value, Gen_DatFormat) & "# "
   Else
      wc = wc & " And TPI.InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TPI.InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
   End If
   
   If Opt7Days.Value = True Then
      wc = wc & " And TPI.TxnType='D'"
   ElseIf OptRRN.Value = True Then
      wc = wc & " And TPI.TxnType='R'"
   ElseIf OptYearly.Value = True Then
      wc = wc & " And TPI.TxnType='Y'"
   ElseIf OptPWR.Value = True Then
      wc = wc & " And TPI.TxnType='P'"
   Else
      wc = wc & " And TPI.TxnType='W'"
   End If
''   If LCase(CmbStatus.Text) = "pending" Then
''      wc = wc & " And  TPI.Flag = 'P'"
''   ElseIf LCase(CmbStatus.Text) = "confirmed" Then
''      wc = wc & " And  TPI.Flag = 'C'"
''   End If
     wc = wc & " And  TPI.Flag = 'C'"
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
   wc = " Where  TPI.BatchID=" & Val(CmbBatch.Text) & " "
''   If Opt7Days.Value = True Then
''      wc = wc & " And TPI.TxnType='D'"
''   Else
''      wc = wc & " And TPI.TxnType='W'"
''   End If
End If

Call Grid_Head

'Commented on 14 Aug 2017 during GST Effect implementaion
'mGLAccNAVCode = GetGLAccNAVCode("7 Days")
'mGLAccNAVCodeRej = GetGLAccNAVCode("Rejected")
'mGLAccNAVCodeRRN = GetGLAccNAVCode("RRN Withdrawal")
'mGLAccNAVCodeYearly = GetGLAccNAVCode("NNNC")

tmp = "Select Distinct TPI.BOEID,TPI.SM_Prefix,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MCSG.GSTNavCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,MST.NAVCode,"
tmp = tmp & " TPI.InvoiceDate , isnull(TPI.Discount,0)Discount, isnull(TPI.Premium,0)Premium, TPI.FAGPostingDate,MComm.Commodity,TPI.TxnType,MComm.HSNNo ,MComm.AccountType,MST.InvoiceType,"
If OptPWR.Value = True Then
   tmp = tmp & " TC.ConvertedID 'CxTF_DetailsID' "
Else
   tmp = tmp & " MA.AgentName,TC.CxTF_DetailsID 'CxTF_DetailsID' "
End If

If OptRRN.Value = True Then
   tmp = tmp & " , Sum(isnull(TCG.NCDEXSpaceCharges,0)) 'NCDEXSpaceCharges' "
End If

tmp = tmp & " From Txn_CDTFBOE TPI"
tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID"
'Commented on 07 Sept 2012
'tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CDTFNo = TC.CxTFNo And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType='D' "
If Opt7Days.Value = True Then
   tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix And TC.TxnType = 'D'  and TPI.TxnType = 'D' "
ElseIf OptRRN.Value = True Then
   tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType = 'W'  and TPI.TxnType = 'R' "
ElseIf OptYearly.Value = True Then
   tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType = 'W'  and TPI.TxnType = 'Y' "
ElseIf OptPWR.Value = True Then
   tmp = tmp & " Inner Join Txn_RRNPWR TC On TPI.CDTFNo = TC.ConvertedID And TPI.TxnType = 'P' "
Else
   tmp = tmp & " Inner Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType = 'W'  and TPI.TxnType = 'W' "
End If
If OptPWR.Value = False Then
   tmp = tmp & " left Join Mst_Agent MA On TC.AgentID = MA.AgentID"
End If
tmp = tmp & " inner join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MC.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"

tmp = tmp & " inner join Mst_ServiceTax MST On MComm.ServiceTaxID = MST.ServiceTaxID "
If OptRRN.Value = True Then
   tmp = tmp & " Inner join Txn_CDTFBOE_GSL TCG on TPI.BOEID = TCG.BOEID "
End If

tmp = tmp & wc '& "  " '" Order By TPI.InvoiceDate"
If OptRRN.Value = True Then
   tmp = tmp & " Group By TPI.BOEID,TPI.SM_Prefix,ML.LocationName,ML.NAVCode,MC.ClientName,MCSG.GSTNavCode,TPI.Amount,"
   tmp = tmp & " TPI.BatchID,MST.NAVCode, TPI.InvoiceDate , TPI.Discount,TPI.Premium, TPI.FAGPostingDate,"
   tmp = tmp & " MComm.Commodity , TPI.TxnType, MA.AgentName, TC.CxTF_DetailsID,MComm.HSNNo ,MComm.AccountType,MST.InvoiceType"
End If
tmp = tmp & " Order By TPI.BOEID "

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
   If TmpRs!NavCode = "GSTS18" Then
      CmbInvoiceType.Text = "Taxable"
   ElseIf TmpRs!NavCode = "GSTS00" Then
      CmbInvoiceType.Text = "Non Taxable"
   End If
   
   If TmpRs!TxnType = "D" Then
      Opt7Days.Value = True
      mInvoiceType = 1
   ElseIf TmpRs!TxnType = "R" Then
      OptRRN.Value = True
      mInvoiceType = 13
   ElseIf TmpRs!TxnType = "Y" Then
      OptYearly.Value = True
      mInvoiceType = 7
   ElseIf TmpRs!TxnType = "P" Then
      OptPWR.Value = True
      mInvoiceType = 13
   Else
      OptRej.Value = True
      mInvoiceType = 2
   End If
   
   Frame5.Caption = CmbInvoiceType.Text
   
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!InvoiceDate, "mm/dd/yyyy")  ', "dd/mm/yyyy")
       MSHFlexGrid3.TextMatrix(I + 1, 1) = "Invoice"
       MSHFlexGrid3.TextMatrix(I + 1, 2) = "Customer"
       MSHFlexGrid3.TextMatrix(I + 1, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
       
       'If TmpRs!Discount <> 0 And IsNull(TmpRs!Discount) = Flase Then
       If OptRRN.Value = True Then
          MSHFlexGrid3.TextMatrix(I + 1, 4) = IIf(IsNull(TmpRs!Amount), 0, TmpRs!Amount) + TmpRs!Premium - TmpRs!Discount
          If Val(MSHFlexGrid3.TextMatrix(I + 1, 4)) > Val(TmpRs!NCDEXSpaceCharges) Then
             MSHFlexGrid3.TextMatrix(I + 1, 4) = Val(MSHFlexGrid3.TextMatrix(I + 1, 4)) - Val(TmpRs!NCDEXSpaceCharges)
          End If
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 4) = IIf(IsNull(TmpRs!Amount), 0, TmpRs!Amount) + TmpRs!Premium - TmpRs!Discount
       End If
       'ElseIf TmpRs!Premium <> 0 And IsNull(TmpRs!Premium) = Flase Then
       '   MSHFlexGrid3.TextMatrix(i + 1, 4) = TmpRs!Amount + TmpRs!Premium
       'Else
       '   MSHFlexGrid3.TextMatrix(i + 1, 4) = TmpRs!Amount
       'End If
       MSHFlexGrid3.TextMatrix(I + 1, 5) = "G/L Account"
       MSHFlexGrid3.TextMatrix(I + 1, 20) = IIf(IsNull(TmpRs!AccountType), "", TmpRs!AccountType)
       MSHFlexGrid3.TextMatrix(I + 1, 21) = TmpRs!NavCode
       MSHFlexGrid3.TextMatrix(I + 1, 22) = IIf(IsNull(TmpRs!HSNNo), "", TmpRs!HSNNo)
       MSHFlexGrid3.TextMatrix(I + 1, 23) = IIf(IsNull(TmpRs!InvoiceType), "", TmpRs!InvoiceType)
       MSHFlexGrid3.TextMatrix(I + 1, 24) = "D"
       
       If CmbInvoiceType.Text = "Non Taxable" Then
          MSHFlexGrid3.TextMatrix(I + 1, 6) = GetGLAccNAVCodeFromInvoiceType(mInvoiceType, "AGRI", "G")
          If MSHFlexGrid3.TextMatrix(I + 1, 22) = "" Then
             MSHFlexGrid3.TextMatrix(I + 1, 22) = GetGLAccNAVCodeFromInvoiceType(mInvoiceType, "AGRI", "H")
          End If
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 6) = GetGLAccNAVCodeFromInvoiceType(mInvoiceType, "NAGRI", "G")
          If MSHFlexGrid3.TextMatrix(I + 1, 22) = "" Then
             MSHFlexGrid3.TextMatrix(I + 1, 22) = GetGLAccNAVCodeFromInvoiceType(mInvoiceType, "NAGRI", "H")
          End If
       End If
       If Opt7Days.Value = True Then
          'MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCode
          If CmbInvoiceType.Text = "Non Taxable" Then
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NCDEX pre demat storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: CDTF No: " & TmpRs!CxTF_DetailsID & " Agent : " & IIf(IsNull(TmpRs!AgentName), "-", TmpRs!AgentName) & "" '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (Without ST) having PI No :- SD" & TmpRs!BOEID & " "
          Else
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NCDEX pre demat storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: CDTF No: " & TmpRs!CxTF_DetailsID & " Agent : " & IIf(IsNull(TmpRs!AgentName), "-", TmpRs!AgentName) & "" '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (With ST) having PI No :- SD" & TmpRs!BOEID & " "
          End If
       ElseIf OptRRN.Value = True Then
'          MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCodeRRN
          If CmbInvoiceType.Text = "Non Taxable" Then
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NCDEX RRN Withdrawal storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: CDTF No: " & TmpRs!CxTF_DetailsID & " Agent : " & IIf(IsNull(TmpRs!AgentName), "-", TmpRs!AgentName) & "" '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (Without ST) having PI No :- SD" & TmpRs!BOEID & " "
          Else
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NCDEX RRN Withdrawal storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: CDTF No: " & TmpRs!CxTF_DetailsID & " Agent : " & IIf(IsNull(TmpRs!AgentName), "-", TmpRs!AgentName) & "" '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (With ST) having PI No :- SD" & TmpRs!BOEID & " "
          End If
       ElseIf OptYearly.Value = True Then
'          MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCodeYearly
          
          If CmbInvoiceType.Text = "Non Taxable" Then
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NNNR Yearly storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: CITF No: " & TmpRs!CxTF_DetailsID & " Agent : " & IIf(IsNull(TmpRs!AgentName), "-", TmpRs!AgentName) & "" '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (Without ST) having PI No :- SD" & TmpRs!BOEID & " "
          Else
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NNNR Yearly storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: CITF No: " & TmpRs!CxTF_DetailsID & " Agent : " & IIf(IsNull(TmpRs!AgentName), "-", TmpRs!AgentName) & "" '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (With ST) having PI No :- SD" & TmpRs!BOEID & " "
          End If
       ElseIf OptPWR.Value = True Then
'          MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCodeRRN
          
          If CmbInvoiceType.Text = "Non Taxable" Then
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NCDEX RRN Withdrawal storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: PWR No: " & TmpRs!CxTF_DetailsID & " " '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (Without ST) having PI No :- SD" & TmpRs!BOEID & " "
          Else
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NCDEX RRN Withdrawal storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: PWR No: " & TmpRs!CxTF_DetailsID & " " '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (With ST) having PI No :- SD" & TmpRs!BOEID & " "
          End If
       
       Else
'          MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCodeRej
          
          If CmbInvoiceType.Text = "Non Taxable" Then
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NCDEX rejected withdrawal storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: CITF No: " & TmpRs!CxTF_DetailsID & " Agent : " & IIf(IsNull(TmpRs!AgentName), "-", TmpRs!AgentName) & "" '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (Without ST) having PI No :- SD" & TmpRs!BOEID & " "
          Else
             MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being NCDEX rejected withdrawal storage charges for " & TmpRs!Commodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!Locationname & " - ref: CITF No: " & TmpRs!CxTF_DetailsID & " Agent : " & IIf(IsNull(TmpRs!AgentName), "-", TmpRs!AgentName) & "" '"Being 7 Days Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (With ST) having PI No :- SD" & TmpRs!BOEID & " "
          End If
       End If
       
       'Being NCDEX pre demat storage charges for " & Tmprs!COmmodity & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " at " & TmpRs!LocationName & " "
       
       
       
       If Opt7Days.Value = True Then
          MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!BOEID), "", "SD" & TmpRs!BOEID)
       ElseIf OptRRN.Value = True Then
          MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!BOEID), "", "RRN" & TmpRs!BOEID)
       ElseIf OptYearly.Value = True Then
          MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!BOEID), "", "NNNRY" & TmpRs!BOEID)
       ElseIf OptPWR.Value = True Then
          MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!BOEID), "", "PWR" & TmpRs!BOEID)
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!BOEID), "", "RJ" & TmpRs!BOEID)
       End If
       
       MSHFlexGrid3.TextMatrix(I + 1, 13) = IIf(IsNull(TmpRs!LocatioNAVCode), "", TmpRs!LocatioNAVCode)
       MSHFlexGrid3.TextMatrix(I + 1, 14) = "CMP"
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I + 1, 4))
       
       If MSHFlexGrid3.TextMatrix(I + 1, 6) = "" Then
          ExpFlag = "N"
       End If
              
       If IsNull(TmpRs!ClientNAVCode) = True Or TmpRs!ClientNAVCode = "" Then
          ExpFlag = "N"
       End If
       If IsNull(TmpRs!LocatioNAVCode) = True Or TmpRs!LocatioNAVCode = "" Then
          ExpFlag = "N"
       End If
'       If mGLAccNAVCode = "" Then
'          ExpFlag = "N"
'       End If
       
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
    
    .TextMatrix(1, 0) = "Date":                 .ColWidth(0) = 1200
    .TextMatrix(1, 1) = "Invoice":              .ColWidth(1) = 1200
    .TextMatrix(1, 2) = "Customer":             .ColWidth(2) = 1400
    .TextMatrix(1, 3) = "DPNavCode":            .ColWidth(3) = 1400
    .TextMatrix(1, 4) = "Amount":               .ColWidth(4) = 1600
    .TextMatrix(1, 5) = "G/L Account":          .ColWidth(5) = 1600
    .TextMatrix(1, 6) = "GL A/C NAVCode":       .ColWidth(6) = 1600
    .TextMatrix(1, 7) = "":                     .ColWidth(7) = 0
    .TextMatrix(1, 8) = "":                     .ColWidth(8) = 0
    .TextMatrix(1, 9) = "":                     .ColWidth(9) = 0
    .TextMatrix(1, 10) = "":                    .ColWidth(10) = 0
    .TextMatrix(1, 11) = "":                    .ColWidth(11) = 0
    .TextMatrix(1, 12) = "Narration":           .ColWidth(12) = 2000
    .TextMatrix(1, 13) = "Location NAV Code":   .ColWidth(13) = 1600
    .TextMatrix(1, 14) = "CMP":                 .ColWidth(14) = 1400
    .TextMatrix(1, 15) = "":                    .ColWidth(15) = 0
    .TextMatrix(1, 16) = "":                    .ColWidth(16) = 0
    .TextMatrix(1, 17) = "":                    .ColWidth(17) = 0
    .TextMatrix(1, 18) = "PI No":               .ColWidth(18) = 1400
    .TextMatrix(1, 19) = "":                    .ColWidth(19) = 0
    .TextMatrix(1, 20) = "Agri/Non Agri":       .ColWidth(20) = 1400
    .TextMatrix(1, 21) = "GST Code":            .ColWidth(21) = 1400
    .TextMatrix(1, 22) = "HSN No":              .ColWidth(22) = 1400
    .TextMatrix(1, 23) = "Invoice Type":        .ColWidth(23) = 1400
    .TextMatrix(1, 24) = "Number Series Flag":  .ColWidth(24) = 1400
    
    If Option1.Value = True Then
       .TextMatrix(1, 25) = "Selection":        .ColWidth(24) = 1400
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
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = False
   Frame2.Enabled = False
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
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = True
   Frame2.Enabled = True
   'CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub
Private Sub TxtFromDate_LostFocus()
'TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
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
'TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
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
