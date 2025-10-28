VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_IncomeDetails 
   Caption         =   "Income Summary Upload Details FAG format"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7815
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
         TabIndex        =   11
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
         Left            =   4440
         TabIndex        =   10
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
         Left            =   5895
         TabIndex        =   9
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
         Height          =   350
         Left            =   1530
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
         Left            =   2985
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
         Top             =   630
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   8040
      TabIndex        =   3
      Top             =   0
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
   Begin VB.Frame Frame4 
      Caption         =   "Error Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   15000
      Begin VB.TextBox TxtErrorMsg 
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   240
         Width           =   14535
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
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
      TabIndex        =   13
      Top             =   9360
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
End
Attribute VB_Name = "Frm_IncomeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim ErrMsg As Variant
Dim Status As Variant
Dim mSheetFlg As Variant

Private Sub CmdGetFile_Click()
'ListView1.ListItems.Clear
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

CmbSheetName.Clear
If LblFile.Caption <> "" Then
   CmdExit.Enabled = True
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
 ' Code For Creating List of Sheet in Excel Sheet
   Dim i As Integer

   For i = 1 To oWB.Worksheets.Count
       Set oWS = oWB.Worksheets(i)
       'ListView1.ListItems.Add i, , oWS.Name
       Status = 1
       Frame2.Enabled = True
       If LCase(oWS.Name) = "location corporate" Or LCase(oWS.Name) = "location s & w" Or LCase(oWS.Name) = "location dp" Then
          CmbSheetName.AddItem oWS.Name
       Else
          Status = 0
          Exit For
       End If
   Next
     
   If Status = 0 Then
      CmbSheetName.Clear
      Frame2.Enabled = False
      MsgBox "Invalid File!!!"
      Exit Sub
   End If
   MsgBox "Click Import after Selecting Sheet Name !!!"
   CmbSheetName.SetFocus
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   oXL.Quit
End If

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub CmbSheetName_LostFocus()
If CmbSheetName.Text = "" Then
   mSheetFlg = Null
   Exit Sub
End If
'If LCase(oWS.Name) = "non-ncdex op.bal" Or LCase(oWS.Name) = "non-ncdex deposits" Or LCase(oWS.Name) = "non-ncdex withdrawals" Or LCase(oWS.Name) = "non-ncdex payments" Or LCase(oWS.Name) = "ncdex withdrawals" Then
If LCase(CmbSheetName.Text) = "location corporate" Then
   mSheetFlg = "lc"
ElseIf LCase(CmbSheetName.Text) = "location s & w" Then
   mSheetFlg = "lsw"
ElseIf LCase(CmbSheetName.Text) = "location dp" Then
   mSheetFlg = "ldp"
Else
    MsgBox "Invalid Selection!!!"
    CmbSheetName.SetFocus
    Exit Sub
End If

Call GridHead

End Sub
Private Sub CmdExit_Click()
'If oXL.ActiveWorkbook = True Then oXL.Quit
Unload Me
End Sub

Private Sub CmdImport_Click()
If CmbSheetName.Text = "" Then
   MsgBox "Name Of Sheet Can't be blank!!!"
   CmbSheetName.SetFocus
   Exit Sub
End If
'-----------------------------
If CmbSheetName.Text <> "" Then
   MsgBox "Wait till next message!!!"
   CmdImport.Enabled = False
   CmdReset.Enabled = False
   CmdExit.Enabled = True
   CmdGetFile.Enabled = False
   Call GridHead
   TxtErrorMsg = ""
   ErrMsg = ""
   Status = 0
   RowCount = 0
   i = 1
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)

   Call GridHead
   Call FillData
   CmdImport.Enabled = True
   CmdSave.Enabled = True
   CmdExit.Enabled = True
   CmdReset.Enabled = True
   MsgBox "Import Succesful!!!"
End If

End Sub

Private Sub cmdReset_Click()

Call GridHead
'CmdImport.Enabled = True
CmdExit.Enabled = True
CmdSave.Enabled = False
CmdGetFile.Enabled = True
If LblFile.Caption <> "" Then
   CmdImport.Enabled = True
End If
mSheetFlg = ""
TxtErrorMsg = ""
ProgressBar1.Value = 0

End Sub

Private Sub CmdSave_Click()
Dim mLocationID As Double
Dim mStartDate, mEndDate, mTxnType, ErrMsg As Variant

CmdSave.Enabled = False
If MSHFlexGrid1.Rows < 3 Then
   MsgBox "No data to save!!!"
   Exit Sub
End If

If MSHFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
MsgBox "Wait till Next Message....."

Call TableTxn_IncomeDetails

ErrMsg = ""
mRow = 0
  
RowCount = MSHFlexGrid1.Rows + 1 'oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row

ProgressBar1.Value = 0
ProgressBar1.Max = RowCount

If mSheetFlg = "lc" Then
   mTxnType = "Corporate"
ElseIf mSheetFlg = "lsw" Then
   mTxnType = "S & W"
ElseIf mSheetFlg = "ldp" Then
   mTxnType = "DP"
End If

For c = 1 To RowCount
    If MSHFlexGrid1.TextMatrix(c, 0) = "" Then Exit For
    mRow = mRow + 1
    Status = 0
    If MSHFlexGrid1.TextMatrix(c, 0) <> "" Then
       mLocationID = GetLocationID(Trim(MSHFlexGrid1.TextMatrix(c, 0)))
       wc = wc & " and LocationID = " & mLocationID & " "
    End If
    If mLocationID <> 0 Then
       For j = 1 To MSHFlexGrid1.Cols - 1
           mStartDate = Null
           mEndDate = Null
           If j = 1 Then
              mStartDate = "04/01/2009"
              mEndDate = "04/30/2009"
           ElseIf j = 2 Then
              mStartDate = "05/01/2009"
              mEndDate = "05/31/2009"
           ElseIf j = 3 Then
              mStartDate = "06/01/2009"
              mEndDate = "06/30/2009"
           ElseIf j = 4 Then
              mStartDate = "07/01/2009"
              mEndDate = "07/31/2009"
           ElseIf j = 5 Then
              mStartDate = "08/01/2009"
              mEndDate = "08/31/2009"
           ElseIf j = 6 Then
              mStartDate = "09/01/2009"
              mEndDate = "09/30/2009"
           End If
           RsTxn_IncomeDetails.AddNew
           RsTxn_IncomeDetails!LocationID = mLocationID
           RsTxn_IncomeDetails!BillingMonthStartDate = mStartDate
           RsTxn_IncomeDetails!BillingMonthEndDate = mEndDate
           RsTxn_IncomeDetails!Mode = CmbSheetName.Text
           RsTxn_IncomeDetails!TxnType = mTxnType
           RsTxn_IncomeDetails!ClientIDRow = 1096
           RsTxn_IncomeDetails!AmtInclOfSerTax = Val(Format(MSHFlexGrid1.TextMatrix(c, j), "########0.00"))
           RsTxn_IncomeDetails.Update
       Next
    Else
       ErrMsg = ErrMsg & "Location Name " & MSHFlexGrid1.TextMatrix(c, 0) & " not found in database" & vbCrLf
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.Row = c
       MSHFlexGrid1.CellBackColor = vbRed
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
TxtErrorMsg = ErrMsg
MsgBox "Done!!!"
 

End Sub

Private Sub Form_Load()

LblFile.Caption = ""
Frame1.Enabled = False
Frame2.Enabled = False
Call GridHead
ProgressBar1.Value = 0

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True


End Sub

Private Sub GridHead()
   
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 7
     .AllowUserResizing = flexResizeColumns
     .TextMatrix(0, 0) = "Location Name"
     .TextMatrix(0, 1) = "April 09"
     .TextMatrix(0, 2) = "May 09"
     .TextMatrix(0, 3) = "June 09"
     .TextMatrix(0, 4) = "July 09"
     .TextMatrix(0, 5) = "August 09"
     .TextMatrix(0, 6) = "September 09"
End With
  
MSHFlexGrid1.ColWidth(0) = 3200
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 2000
MSHFlexGrid1.ColWidth(5) = 2200

End Sub

Private Sub FillData()

Set oWS = oWB.Worksheets(CmbSheetName.Text)
oWS.ClearArrows
mRow = 1
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row
'ProgressBar1.Value = 0
'ProgressBar1.Max = rowCount
For L = 1 To RowCount
    mRow = mRow + 1
    If oWS.Cells(mRow, 1) <> "" Then
        For c = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.TextMatrix(L, c) = Trim(oWS.Cells(L + 1, c + 1).Value)
         Next
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    End If
Next
oXL.Quit
End Sub

