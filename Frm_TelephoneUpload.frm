VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_TelephoneUpload 
   Caption         =   "Telephone Upload"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   10890
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Select Upload File Type "
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   15015
      Begin VB.OptionButton OptDataCard 
         Caption         =   "Data Card Details"
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
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton OptMobile 
         Caption         =   "Mobile Details"
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
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label LblRecordCount 
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
         Left            =   11160
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblTotalRecord 
         Caption         =   "Total record Found :"
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
         Left            =   8880
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   7815
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
         Left            =   1530
         TabIndex        =   11
         Top             =   165
         Width           =   1335
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
         Left            =   75
         TabIndex        =   10
         Top             =   165
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
         Height          =   350
         Left            =   4320
         TabIndex        =   9
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
         Height          =   350
         Left            =   5760
         TabIndex        =   8
         Top             =   165
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
         Height          =   350
         Left            =   2865
         TabIndex        =   7
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   570
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   8040
      TabIndex        =   3
      Top             =   720
      Width           =   7110
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   360
         Width           =   2025
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13785
      _Version        =   393216
      WordWrap        =   -1  'True
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
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_TelephoneUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim ErrMsg As Variant
Dim Status As Variant
Dim WrkSheet As String
Dim FlagType As Variant
Dim NotExists As Variant

Private Sub CmdGetFile_Click()
On Error GoTo errortrap
If OptMobile.Value = False And OptDataCard.Value = False Then
   MsgBox "Please Select Upload File Type !!!", vbExclamation
   Exit Sub
End If

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls)|*.xls"
     .CancelError = False
procReOpen:
     .ShowOpen
End With

Path = CommonDialog1.FileName
LblFile.Caption = Path

Call UploadFile

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub CmbSheetName_LostFocus()
If CmbSheetName.Text = "" Then
   mSheetFlg = Null
   Exit Sub
End If
LblRecordCount.Caption = ""
mSheetFlg = CmbSheetName.Text
End Sub
Private Sub UploadFile()
CmbSheetName.Clear
If LblFile.Caption <> "" Then
   CmdExit.Enabled = True
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
 ' Code For Creating List of Sheet in Excel Sheet
   Dim I As Integer
   For I = 1 To oWB.Worksheets.Count
       Set oWS = oWB.Worksheets(I)
          CmbSheetName.AddItem oWS.Name
   Next
   
   Set oWS = oWB.Worksheets(1)
   
   If FlagType = "T" Then
      If Trim(oWS.Cells(5, 5)) <> "CURRENT MOBILE" Then
         MsgBox "Please Check Selected File !!!" & vbCrLf & "Select Telephone Details File", vbExclamation
         LblFile.Caption = ""
         CmbSheetName.Clear
         oXL.Quit
         Exit Sub
      End If
   Else
      If Trim(oWS.Cells(3, 5)) <> "CURRENT CARD" Then
         MsgBox "Please Check Selected File !!!" & vbCrLf & "Select Internet Card Details File", vbExclamation
         LblFile.Caption = ""
         CmbSheetName.Clear
         oXL.Quit
         Exit Sub
      End If
   End If
   
   MsgBox "Click Import after Selecting Sheet Name !!!"
   CmbSheetName.SetFocus
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   oXL.Quit
End If
End Sub
Private Sub CmdExit_Click()
On Error Resume Next
oXL.Quit
Unload Me
End Sub

Private Sub CmdImport_Click()
If CmbSheetName.Text = "" Then
   MsgBox "Name Of Sheet Can't be blank!!!"
   CmbSheetName.SetFocus
   Exit Sub
End If

If CmbSheetName.Text <> "" Then
   MSHFlexGrid1.Clear
   LblRecordCount.Caption = ""
   Call GridHead
   MsgBox "Please Wait Till Next Message !!!"
   CmdImport.Enabled = False
   CmdReset.Enabled = False
   CmdExit.Enabled = True
   CmdGetFile.Enabled = False
   ErrMsg = ""
   Status = 0
   RowCount = 0
   I = 1
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)

'   If Chk_FileHeads = False Then
'      Exit Sub
'   End If
   Call FillData
Else
   MsgBox "Please Select WorkSheet To Import !...", vbInformation
End If
End Sub

Private Sub CmdReset_Click()
CmbSheetName.Clear
CmdExit.Enabled = True
CmdGetFile.Enabled = True
CmdSave.Enabled = False
CmdImport.Enabled = False
LblRecordCount.Caption = ""
Call GridHead
End Sub
Private Sub GridHead()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    If FlagType = "T" Then
       .Cols = 9
    Else
       .Cols = 7
    End If
    
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "EMP NO"
    .TextMatrix(0, 1) = "EMP Name"
    .TextMatrix(0, 2) = "Designation"
    
    If FlagType = "T" Then
       .TextMatrix(0, 3) = "Mobile No"
       .TextMatrix(0, 7) = "Excess"
       .TextMatrix(0, 8) = "NCMSL Bill Amount"
       .ColWidth(7) = 1500
       .ColWidth(8) = 1900
    Else
       .TextMatrix(0, 3) = "Data Card No"
    End If
    
    .TextMatrix(0, 4) = "Account No"
    .TextMatrix(0, 5) = "Eligibility"
    .TextMatrix(0, 6) = "Monthly Bill Amount"
    .ColWidth(0) = 950
    .ColWidth(1) = 2500
    .ColWidth(2) = 2000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1150
    .ColWidth(6) = 1500
End With
End Sub
Private Sub FillData()
Set oWS = oWB.Worksheets(CmbSheetName.Text)
oWS.ClearArrows

If FlagType = "T" Then
   mRow = 8
Else
   mRow = 5
End If

On Error GoTo errortrap
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row
mCol = 0
mshfrow = 1
For l = 1 To RowCount
    mCol = 0
    If oWS.Cells(mRow, 1).Value = "" And (oWS.Cells(mRow, 3).Value) <> "" Then
       mRow = mRow + 1
    ElseIf oWS.Cells(mRow, 1).Value = "" And (oWS.Cells(mRow, 3).Value) = "" Then
       Exit For
    Else
       For c = 1 To 13
           If FlagType = "T" Then
              If c = 2 Or c = 3 Or c = 4 Or c = 5 Or c = 8 Or c = 9 Or c = 10 Or c = 11 Then
                 MSHFlexGrid1.TextMatrix(mshfrow, mCol) = Trim(oWS.Cells(mRow, c).Value)
                 mCol = mCol + 1
              End If
              If mCol = 8 Then
                 MSHFlexGrid1.TextMatrix(mshfrow, mCol) = Val(MSHFlexGrid1.TextMatrix(mshfrow, 6)) - Val(MSHFlexGrid1.TextMatrix(mshfrow, 7))
              End If
           Else
              If c = 2 Or c = 3 Or c = 4 Or c = 5 Or c = 6 Or c = 7 Or c = 11 Then
                 MSHFlexGrid1.TextMatrix(mshfrow, mCol) = Trim(oWS.Cells(mRow, c).Value)
                 mCol = mCol + 1
              End If
           End If
       Next
       mRow = mRow + 1
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       mshfrow = mshfrow + 1
    End If
Next
LblRecordCount.Caption = MSHFlexGrid1.Rows - 2
CmdImport.Enabled = True
CmdSave.Enabled = True
CmdExit.Enabled = True
CmdReset.Enabled = True
MsgBox "Import Successful !!"
oXL.Quit
Exit Sub
errortrap:
      MsgBox "Please Check Your Excel Sheet Data !!!", vbCritical
      oXL.Quit
End Sub

Private Sub CmdSave_Click()
Dim r, I As Integer
Call TableTxn_TelephoneDetails
Call TableTxn_TelephoneTemp
''****Data Store in Temp Table
For r = 1 To MSHFlexGrid1.Rows - 2
    RsTxn_TelephoneTemp.AddNew
    RsTxn_TelephoneTemp!BillMonthYear = CmbSheetName.Text
    RsTxn_TelephoneTemp!EmpNo = MSHFlexGrid1.TextMatrix(r, 0)
    RsTxn_TelephoneTemp!Name = MSHFlexGrid1.TextMatrix(r, 1)
    RsTxn_TelephoneTemp!Designation = MSHFlexGrid1.TextMatrix(r, 2)
    RsTxn_TelephoneTemp!MobNo = MSHFlexGrid1.TextMatrix(r, 3)
    RsTxn_TelephoneTemp!AccountNo = MSHFlexGrid1.TextMatrix(r, 4)
    RsTxn_TelephoneTemp!Eligibility = MSHFlexGrid1.TextMatrix(r, 5)
    RsTxn_TelephoneTemp!MonthlyBillAmt = Val(MSHFlexGrid1.TextMatrix(r, 6))
    If FlagType = "T" Then
       RsTxn_TelephoneTemp!Excess = Val(MSHFlexGrid1.TextMatrix(r, 7))
       RsTxn_TelephoneTemp!NCMSLBillAmt = Val(MSHFlexGrid1.TextMatrix(r, 8))
    End If
    RsTxn_TelephoneTemp!Flag = FlagType
    RsTxn_TelephoneTemp.Update
Next
''****Data Store In Permanant Table
tmp = "Select BillMonthYear,EmpNo,MobNo,Eligibility,AccountNo,Flag, Sum(MonthlyBillAmt) as MonthlyBillAmt ,Sum(Excess) as Excess ,Sum(MonthlyBillAmt - Excess) as NCMSLBillAmt"
tmp = tmp & " From Txn_TelephoneTemp"
tmp = tmp & " Group by EmpNo,BillMonthYear,MobNo,Eligibility,AccountNo,Flag"

Call TmpTable

For I = 0 To TmpRs.RecordCount - 1
    RsTxn_TelephoneDetail.Filter = ""
    If RsTxn_TelephoneDetail.RecordCount > 0 Then
       RsTxn_TelephoneDetail.MoveFirst
    End If
    RsTxn_TelephoneDetail.Filter = " EmpNo ='" & TmpRs!EmpNo & "' And MobNo='" & TmpRs!MobNo & "' And BillMonthYear ='" & TmpRs!BillMonthYear & "'And Flag='" & FlagType & "' "
    If RsTxn_TelephoneDetail.RecordCount = 0 Then
       RsTxn_TelephoneTemp.MoveFirst
       RsTxn_TelephoneTemp.Find " EmpNo='" & TmpRs!EmpNo & "' "
       'RsTxn_TelephoneTemp.Find " MobNo='" & TmpRs!MobNo & "' "
       RsTxn_TelephoneDetail.AddNew
       RsTxn_TelephoneDetail!BillMonthYear = TmpRs!BillMonthYear
       RsTxn_TelephoneDetail!EmpNo = TmpRs!EmpNo
       RsTxn_TelephoneDetail!Name = RsTxn_TelephoneTemp!Name
       RsTxn_TelephoneDetail!Designation = RsTxn_TelephoneTemp!Designation
       RsTxn_TelephoneDetail!MobNo = TmpRs!MobNo
       If FlagType = "T" Then
          RsTxn_TelephoneDetail!Excess = TmpRs!Excess
          RsTxn_TelephoneDetail!NCMSLBillAmt = TmpRs!NCMSLBillAmt
       End If
       RsTxn_TelephoneDetail!AccountNo = TmpRs!AccountNo
       RsTxn_TelephoneDetail!Eligibility = TmpRs!Eligibility
       RsTxn_TelephoneDetail!MonthlyBillAmt = TmpRs!MonthlyBillAmt
       RsTxn_TelephoneDetail!Flag = FlagType
       RsTxn_TelephoneDetail!G_UID = GetGUID
       RsTxn_TelephoneDetail.Update
    End If
    TmpRs.MoveNext
Next

tmp = "Delete from Txn_TelephoneTemp"
Call TmpTable

MsgBox "Data Saved To DataBase Successfully !!!", vbInformation
CmdSave.Enabled = False
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If
   Frame1.Enabled = True
LblFile.Caption = ""
End Sub

Private Sub OptDataCard_Click()
FlagType = "DC"
Call RefreshBtn
Call GridHead
End Sub

Private Sub OptMobile_Click()
FlagType = "T"
Call RefreshBtn
Call GridHead
End Sub

Private Sub RefreshBtn()
MSHFlexGrid1.Clear
CmbSheetName.Clear
CmdImport.Enabled = False
CmdSave.Enabled = False
CmdGetFile.Enabled = True
LblFile.Caption = ""
LblRecordCount.Caption = ""
End Sub

