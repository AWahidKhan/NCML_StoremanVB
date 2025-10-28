VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_IncomeUpload 
   Caption         =   "Import Income Detail Upload CMG Format"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   14655
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   9480
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11033
      _Version        =   393216
      WordWrap        =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame4 
      Caption         =   "Error Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   15015
      Begin VB.TextBox TxtErrorMsg 
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   240
         Width           =   14775
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   8040
      TabIndex        =   11
      Top             =   120
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
         TabIndex        =   2
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
         TabIndex        =   12
         Top             =   360
         Width           =   2025
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List of Sheets "
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
      Begin VB.PictureBox ListView1 
         BackColor       =   &H80000005&
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
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1635
         ScaleWidth      =   3195
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   165
         Width           =   1455
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
         TabIndex        =   1
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   630
         Width           =   705
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblMsg 
      Height          =   810
      Left            =   10440
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "Frm_IncomeUpload"
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
Dim mClientIDRow, mCommodityID, mExchangeTypeID As Variant
Dim mMonth, mYear As Variant
Dim mStartDate, mEndDate As Variant

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
lblFile.Caption = Path

CmbSheetName.Clear
If lblFile.Caption <> "" Then
   CmdExit.Enabled = True
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(lblFile.Caption)
 ' Code For Creating List of Sheet in Excel Sheet
   Dim I As Integer

   For I = 1 To oWB.Worksheets.Count
       Set oWS = oWB.Worksheets(I)
       'ListView1.ListItems.Add i, , oWS.Name
       Status = 1
       Frame2.Enabled = True
       If LCase(oWS.Name) = "non-ncdex op.bal" Or LCase(oWS.Name) = "non-ncdex deposits" Or LCase(oWS.Name) = "non-ncdex withdrawals" Or LCase(oWS.Name) = "non-ncdex payments" Or LCase(oWS.Name) = "ncdex withdrawals" Then
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
If LCase(CmbSheetName.Text) = "non-ncdex op.bal" Then
   mSheetFlg = "nnob"
ElseIf LCase(CmbSheetName.Text) = "non-ncdex deposits" Then
   mSheetFlg = "nnd"
ElseIf LCase(CmbSheetName.Text) = "non-ncdex withdrawals" Then
   mSheetFlg = "nnw"
ElseIf LCase(CmbSheetName.Text) = "non-ncdex payments" Then
   mSheetFlg = "nnp"
ElseIf LCase(CmbSheetName.Text) = "ncdex withdrawals" Then
   mSheetFlg = "nw"
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
   CmdGetfile.Enabled = False
   Call GridHead
   ErrMsg = ""
   Status = 0
   RowCount = 0
   I = 1
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(lblFile.Caption)
   TxtErrorMsg = ""
   ProgressBar1.Value = 0
   If Chk_FileHeads = False Then
      Exit Sub
   End If
   Call GridHead
   Call FillData
   CmdImport.Enabled = True
   CmdSave.Enabled = True
   CmdExit.Enabled = True
   CmdReset.Enabled = True
   MsgBox "Import Succesful!!!"
End If

End Sub


Private Function Chk_FileHeads() As Boolean
Dim Retval As Boolean
Set oWS = oWB.Worksheets(CmbSheetName.Text)
Retval = True
If mSheetFlg = "nnob" Then
   If LCase(Trim(oWS.Cells(1, 1))) <> "cmp id" Then
      'ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A " & vbCrLf
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of A "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 2))) <> "cmp name" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of B "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 3))) <> "billing month start date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of C "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 4))) <> "billing month end date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of D "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 5))) <> "client name" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of E "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 6))) <> "commodity" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of F "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 7))) <> "opening number of bags (excluding made up bags)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of G "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 8))) <> "bags withdrawn during the month (excluding madeup bags)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of H "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 9))) <> "closing no. of bags (excluding made up bags)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of I "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 10))) <> "opening number of made up bags" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of J "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 11))) <> "new made up bags created during the billing month" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of K "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 12))) <> "withdrawn made up bags during the billing month" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of L "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 13))) <> "closing number of made up bags" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of M "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 14))) <> "rate (per bag per month)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of N "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 15))) <> "storage charges" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of O "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 16))) <> "apply service tax?(y/n)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of P "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 17))) <> "service tax" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of Q "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 18))) <> "insurance charges" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of R "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 19))) <> "amount incl of service tax & insurance" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of S "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 20))) <> "pi number" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of T "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 21))) <> "remarks" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of U "
      Chk_FileHeads = Retval
      Exit Function
   End If
                 
End If
        
If mSheetFlg = "nnd" Then
   If LCase(Trim(oWS.Cells(1, 1))) <> "cmp id" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of A "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 2))) <> "cmp name" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of B "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 3))) <> "billing month end date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of C "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 4))) <> "client name" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of D "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 5))) <> "commodity" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of E "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 6))) <> "cdtf no" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of F "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 7))) <> "cdtf date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of G "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 8))) <> "no. of bags deposited" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of H "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 9))) <> "invoiced   upto date (optional)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of I "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 10))) <> "rate" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of J "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 11))) <> "no. of days" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of K "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 12))) <> "storage charges" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of L "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 13))) <> "apply service tax?(y/n)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of M "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 14))) <> "service tax" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of N "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 15))) <> "insurance charges" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of O "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 16))) <> "amount incl of service tax & insurance" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of P "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 17))) <> "pi number" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of Q "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 18))) <> "whr no" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of R "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 19))) <> "remarks" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of S "
      Chk_FileHeads = Retval
      Exit Function
   End If
             
End If
    
If mSheetFlg = "nnw" Then
   If LCase(Trim(oWS.Cells(1, 1))) <> "cmp id" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of A "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 2))) <> "cmp name" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of B "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 3))) <> "billing month end date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of C "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 4))) <> "citf no." Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of D "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 5))) <> "citf date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of E "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 6))) <> "client name" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of F "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 7))) <> "commodity" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of G "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 8))) <> "corr. cdtf no" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of H "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 9))) <> "cdtf date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of I "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 10))) <> "no. of bags withdrawn (excl. made up bags)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of J "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 11))) <> "no. of madup bags withdrawn" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of K "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 12))) <> "billing cycle (d)aily / (w)eekly / (f)ortnightly / (m)onthly" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of L "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 13))) <> "rate per bag per month" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of M "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 14))) <> "daily" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of N "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 15))) <> "monthly" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of O "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 16))) <> "weekly" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of P "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 17))) <> "fortnightly" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of Q "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 18))) <> "applicable number of days" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of R "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 19))) <> "storage charges" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of S "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 20))) <> "apply service tax?(y/n)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of T "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 21))) <> "service tax" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of U "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 22))) <> "insurance charges" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of V "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 23))) <> "amount incl of service tax & insurance" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of W "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 24))) <> "pi number" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of X "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 25))) <> "whr no" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of Y "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 26))) <> "remarks" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of Z "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 27))) <> "d" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of AA "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 28))) <> "w" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of AB "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 29))) <> "f" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of AC "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 30))) <> "m" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of AD "
      Chk_FileHeads = Retval
      Exit Function
   End If
                
End If
    
If mSheetFlg = "nw" Then
   If LCase(Trim(oWS.Cells(1, 1))) <> "cmp id" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of A "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 2))) <> "cmp name" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of B "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 3))) <> "citf no" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of C "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 4))) <> "citf date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of D "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 5))) <> "client name" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of E "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 6))) <> "commodity" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of F "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 7))) <> "rrn number (for non-rejected withdrawals)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of G "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 8))) <> "cdtf no (for rejected withdrawals)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of H "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 9))) <> "cdtf date (for rejected withdrawals)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of I "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 10))) <> "withdrawn no. of bags" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of J "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 11))) <> "withdrawn quantity in mt" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of K "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 12))) <> "rate per mt per day" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of L "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 13))) <> "no. of days for which charges are already collected" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of M "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 14))) <> "no. of days" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of N "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 15))) <> "storage charges" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of O "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 16))) <> "apply service tax?(y/n)" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of P "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 17))) <> "service tax" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of Q "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 18))) <> "amount incl of service tax" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of R "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 19))) <> "pi num ber" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of S "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 20))) <> "cash receipt no." Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of T "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 21))) <> "cash receipt date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of U "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 22))) <> "cheque no." Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of V "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 23))) <> "cheque date" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of W "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 24))) <> "cheque / cash amount" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of X "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 25))) <> "tds amount" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of Y "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 26))) <> "date of deposited in bank" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of Z "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 27))) <> "deposited bank name" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of AA "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(1, 28))) <> "deposited account no." Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of AB "
      Chk_FileHeads = Retval
      Exit Function
   End If
             
   If LCase(Trim(oWS.Cells(1, 29))) <> "remarks" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of AC "
      Chk_FileHeads = Retval
      Exit Function
   End If
    
End If

Chk_FileHeads = Retval

End Function

Private Sub CmdReset_Click()

Call GridHead
'CmdImport.Enabled = True
CmdExit.Enabled = True
CmdSave.Enabled = False
CmdGetfile.Enabled = True
If lblFile.Caption <> "" Then
   CmdImport.Enabled = True
End If
mSheetFlg = ""
TxtErrorMsg = ""
ProgressBar1.Value = 0

End Sub
Private Sub GridHead()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 2
'MSHFlexGrid1.RowHeight = 200

If mSheetFlg = "nnob" Then
   With MSHFlexGrid1
        .Clear
        .Rows = 2
        .Cols = 21
        .AllowUserResizing = flexResizeColumns
        .TextMatrix(0, 0) = "CMP ID"
        .TextMatrix(0, 1) = "CMP Name"
        .TextMatrix(0, 2) = "Billing Month Start Date"
        .TextMatrix(0, 3) = "Billing Month End Date"
        .TextMatrix(0, 4) = "Client Name"
        .TextMatrix(0, 5) = "Commodity"
        .TextMatrix(0, 6) = "Opening Number of Bags (excluding made up bags)"
        .TextMatrix(0, 7) = "Bags Withdrawn during the month (excluding madeup bags)"
        .TextMatrix(0, 8) = "Closing No. of Bags (excluding made up bags)"
        .TextMatrix(0, 9) = "Opening Number of Made up Bags"
        .TextMatrix(0, 10) = "New made up bags created during the billing month"
        .TextMatrix(0, 11) = "Withdrawn made up bags during the billing month"
        .TextMatrix(0, 12) = "Closing number of made up bags"
        .TextMatrix(0, 13) = "Rate (per bag per month)"
        .TextMatrix(0, 14) = "Storage Charges"
        .TextMatrix(0, 15) = "Apply Service Tax?(y/n)"
        .TextMatrix(0, 16) = "Service Tax"
        .TextMatrix(0, 17) = "Insurance Charges"
        .TextMatrix(0, 18) = "Amount incl of Service Tax & Insurance"
        .TextMatrix(0, 19) = "PI Number"
        .TextMatrix(0, 20) = "Remarks"

   End With
End If
    
If mSheetFlg = "nnd" Then
   With MSHFlexGrid1
        .Clear
        .Rows = 2
        .Cols = 19
        .AllowUserResizing = flexResizeColumns
        .TextMatrix(0, 0) = "CMP ID"
        .TextMatrix(0, 1) = "CMP Name"
        .TextMatrix(0, 2) = "Billing Month End Date"
        .TextMatrix(0, 3) = "Client Name"
        .TextMatrix(0, 4) = "Commodity"
        .TextMatrix(0, 5) = "CDTF No"
        .TextMatrix(0, 6) = "CDTF Date"
        .TextMatrix(0, 7) = "No of Bags Deposited"
        .TextMatrix(0, 8) = "Invoiced upto Date(Opt.)"
        .TextMatrix(0, 9) = "Rate"
        .TextMatrix(0, 10) = "No. of Days"
        .TextMatrix(0, 11) = "Storage Charges"
        .TextMatrix(0, 12) = "Apply service Tax?(y/n)"
        .TextMatrix(0, 13) = "Service Tax"
        .TextMatrix(0, 14) = "Insurance Charges"
        .TextMatrix(0, 15) = "Amt incl of Service Tax & Insurance"
        .TextMatrix(0, 16) = "PI Number"
        .TextMatrix(0, 17) = "WHR No"
        .TextMatrix(0, 18) = "Remarks"
           
   End With
End If
    
If mSheetFlg = "nnw" Then
   With MSHFlexGrid1
        .Clear
        .Rows = 2
        .Cols = 30
        .AllowUserResizing = flexResizeColumns
        .TextMatrix(0, 0) = "CMP ID"
        .TextMatrix(0, 1) = "CMP Name"
        .TextMatrix(0, 2) = "Billing Month End Date"
        .TextMatrix(0, 3) = "CITF No."
        .TextMatrix(0, 4) = "CITF Date"
        .TextMatrix(0, 5) = "Client Name"
        .TextMatrix(0, 6) = "Commodity"
        .TextMatrix(0, 7) = "Corr. CDTF No"
        .TextMatrix(0, 8) = "CDTF Date"
        .TextMatrix(0, 9) = "No. of Bags Withdrawn (excl. made up bags)"
        .TextMatrix(0, 10) = "No. of madup bags withdrawn"
        .TextMatrix(0, 11) = "Billing Cycle(D)aily/(W)eekly/(F)ortnightly/(M)onthly"
        .TextMatrix(0, 12) = "Rate per bag per month"
        .TextMatrix(0, 13) = "Daily"
        .TextMatrix(0, 14) = "Monthly"
        .TextMatrix(0, 15) = "Weekly"
        .TextMatrix(0, 16) = "Fortnightly"
        .TextMatrix(0, 17) = "Applicable No of days"
        .TextMatrix(0, 18) = "Storage Charges"
        .TextMatrix(0, 19) = "Apply Service Tax?(y/n)"
        .TextMatrix(0, 20) = "Service Tax"
        .TextMatrix(0, 21) = "Insurance Charges"
        .TextMatrix(0, 22) = "Amount incl of Service Tax & Insurance"
        .TextMatrix(0, 23) = "PI Number"
        .TextMatrix(0, 24) = "WHR No"
        .TextMatrix(0, 25) = "Remarks"
        .TextMatrix(0, 26) = "D"
        .TextMatrix(0, 27) = "W"
        .TextMatrix(0, 28) = "F"
        .TextMatrix(0, 29) = "M"
        
   End With
End If
    
If mSheetFlg = "nw" Then
   With MSHFlexGrid1
        .Clear
        .Rows = 2
        .Cols = 29
        
        .AllowUserResizing = flexResizeColumns
        .TextMatrix(0, 0) = "CMP ID"
        .TextMatrix(0, 1) = "CMP Name"
        .TextMatrix(0, 2) = "CITF No"
        .TextMatrix(0, 3) = "CITF Date"
        .TextMatrix(0, 4) = "Client Name"
        .TextMatrix(0, 5) = "Commodity"
        .TextMatrix(0, 6) = "RRN Number (for Non-Rejected Withdrawals)"
        .TextMatrix(0, 7) = "CDTF No (for rejected withdrawals)"
        .TextMatrix(0, 8) = "CDTF Date (for rejected withdrawals)"
        .TextMatrix(0, 9) = "Withdrawn No. of Bags "
        .TextMatrix(0, 10) = "Withdrawn Quantity in MT"
        .TextMatrix(0, 11) = "Rate per MT per day"
        .TextMatrix(0, 12) = "No. of Days for which charges are already collected"
        .TextMatrix(0, 13) = "No. of Days"
        .TextMatrix(0, 14) = "Storage Charges"
        .TextMatrix(0, 15) = "Apply Service Tax?(y/n)"
        .TextMatrix(0, 16) = "Service Tax"
        .TextMatrix(0, 17) = "Amt incl of Service Tax"
        .TextMatrix(0, 18) = "PI Num ber"
        .TextMatrix(0, 19) = "Cash Receipt No."
        .TextMatrix(0, 20) = "Cash Receipt Date"
        .TextMatrix(0, 21) = "Cheque No."
        .TextMatrix(0, 22) = "Cheque Date"
        .TextMatrix(0, 23) = "Cheque / Cash Amount"
        .TextMatrix(0, 24) = "TDS Amount"
        .TextMatrix(0, 25) = "Date of Deposited in Bank"
        .TextMatrix(0, 26) = "Deposited Bank Name"
        .TextMatrix(0, 27) = "Deposited Account No."
        .TextMatrix(0, 28) = "Remarks"
   
   
   End With
End If


End Sub
Private Sub FillData()

Set oWS = oWB.Worksheets(CmbSheetName.Text)
oWS.ClearArrows
mRow = 2
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row
'ProgressBar1.Value = 0
'ProgressBar1.Max = rowCount
For l = 1 To RowCount
    mRow = mRow + 1
    If oWS.Cells(mRow, 1) <> "" Then
        For c = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.TextMatrix(l, c) = Trim(oWS.Cells(l + 1, c + 1).Value)
            If MSHFlexGrid1.ColWidth(c) < Len(Trim(MSHFlexGrid1.TextMatrix(l, c))) * 130 Then
                            'MSHFlexGrid1.ColWidth(c) = Len(Trim(oWS.Cells(L, c + 1).Value)) * 130
               If l <> 1 Then
                  If Len(MSHFlexGrid1.TextMatrix(l, c)) = 0 Or Len(MSHFlexGrid1.TextMatrix(l, c)) = 1 Then
                     MSHFlexGrid1.ColWidth(c) = 800
                  ElseIf Len(MSHFlexGrid1.TextMatrix(l, c)) < 15 Then
                     MSHFlexGrid1.ColWidth(c) = Len(Trim(MSHFlexGrid1.TextMatrix(l, c))) * 130
                  End If
               End If
               MSHFlexGrid1.RowHeight(l) = 500
            End If
            
         Next
         MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    End If
Next
oXL.Quit
End Sub

Private Sub CmdSave_Click()
Dim Wc As String
Dim j As Integer
Dim CMP_ID(), mTemp, mCommodityID, mLocationID As Variant
Dim mTxnType, mPINo, mRowSaved As Variant

Dim mCMPID As Double

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
mRowSaved = 0
'mClientIdRow = 1096
RowCount = MSHFlexGrid1.Rows + 1

ProgressBar1.Value = 0
ProgressBar1.Max = RowCount

If mSheetFlg = "nnob" Then
   mRow = mRow + 1
   mExchangeTypeID = GetExchangeID("Non NCDEX")
   If MSHFlexGrid1.TextMatrix(mRow, 0) <> "" Then
      For c = 1 To RowCount - 1
          If MSHFlexGrid1.TextMatrix(c, 0) = "" Then Exit For
          mRow = mRow + 1
          Status = 0
          mTemp = Split(Trim(MSHFlexGrid1.TextMatrix(c, 0)), ";")
          ReDim CMP_ID(UBound(mTemp))
          j = 0
          'If InStr(0, MSHFlexGrid1.TextMatrix(c, 0), ",") = False Then
          Wc = ""
          mCMPID = GetCMPIDEHOPID(mTemp(j))
           'IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
          If mCMPID <> 0 Then
             mLocationID = GetLocationIDFromCMP(mCMPID)
          Else
             ErrMsg = ErrMsg + "CMP " & MSHFlexGrid1.TextMatrix(c, 0) & " At Row " & c & " not found in CMP Master" & vbCrLf
             Status = 1
          End If
          If (Trim(MSHFlexGrid1.TextMatrix(c, 5))) <> "" Then
             mCommodityID = GetCommodityID(Trim(MSHFlexGrid1.TextMatrix(c, 5)))
             If mCommodityID = 0 Then
                ErrMsg = ErrMsg + "Commodity " & MSHFlexGrid1.TextMatrix(c, 5) & " At Row " & c & " not found in Commodity Master" & vbCrLf
                Status = 1
             End If
          Else
             mCommodityID = Null
          End If
                       
          If (Trim(MSHFlexGrid1.TextMatrix(c, 4))) <> "" Then
             mClientIDRow = GetClientIDRowByNameAndExchange(Trim(MSHFlexGrid1.TextMatrix(c, 4)), 1)
             If mClientIDRow = 0 Then
                ErrMsg = ErrMsg + "Client " & MSHFlexGrid1.TextMatrix(c, 4) & " At Row " & c & " not found in Client Master" & vbCrLf
                Status = 1
             End If
          Else
             mClientIDRow = Null
          End If
              
          If Trim(MSHFlexGrid1.TextMatrix(c, 19)) <> "" Then
             mPINo = Trim(MSHFlexGrid1.TextMatrix(c, 19))
          Else
             mPINo = GetMaxNewPINo
          End If                        'PI Number=""
          If Status = 0 Then
             RsTxn_IncomeDetails.AddNew
             RsTxn_IncomeDetails!Mode = "Non-NCDEX Op.Bal"
             RsTxn_IncomeDetails!TxnType = "D"
             RsTxn_IncomeDetails!LocationID = mLocationID
             RsTxn_IncomeDetails!CommodityID = mCommodityID
             RsTxn_IncomeDetails!ClientIDRow = mClientIDRow
             RsTxn_IncomeDetails!ExchangeTypeID = mExchangeTypeID
             RsTxn_IncomeDetails!BillingMonthStartDate = mStartDate 'Trim(MSHFlexGrid1.TextMatrix(c, 2))
             RsTxn_IncomeDetails!BillingMonthEndDate = mEndDate '(Trim(MSHFlexGrid1.TextMatrix(c, 3)))
             RsTxn_IncomeDetails!OpeningNoofBags = Val(Trim(MSHFlexGrid1.TextMatrix(c, 6)))
             RsTxn_IncomeDetails!Rate = Val(Trim(MSHFlexGrid1.TextMatrix(c, 13)))
             RsTxn_IncomeDetails!StorageCharges = Val(Trim(MSHFlexGrid1.TextMatrix(c, 14))) + Val(Trim(MSHFlexGrid1.TextMatrix(c, 17)))
             RsTxn_IncomeDetails!ServiceTax = Val(Trim(MSHFlexGrid1.TextMatrix(c, 16)))
             RsTxn_IncomeDetails!InsuranceCharges = Val(Trim(MSHFlexGrid1.TextMatrix(c, 17)))
             RsTxn_IncomeDetails!AmtinclofSerTaxNInsurance = Val(Trim(MSHFlexGrid1.TextMatrix(c, 18)))
             RsTxn_IncomeDetails!AmtInclOfSerTax = Val(Trim(MSHFlexGrid1.TextMatrix(c, 18)))
             If Trim(MSHFlexGrid1.TextMatrix(c, 19)) <> "" Then
                RsTxn_IncomeDetails!OldPINo = mPINo 'Trim(MSHFlexGrid1.TextMatrix(c, 19)) 'PI_No
             Else
                RsTxn_IncomeDetails!NewPINo = mPINo
             End If
             RsTxn_IncomeDetails!Remarks = Trim(MSHFlexGrid1.TextMatrix(c, 20))
             RsTxn_IncomeDetails!G_UID = GetGUID
             RsTxn_IncomeDetails!CreatedBy = Gen_UserLoginID
             RsTxn_IncomeDetails!CreatedDate = Now
             RsTxn_IncomeDetails.Update
             mRowSaved = mRowSaved + 1
          End If
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End If
End If

'=========================================================================================
    
If mSheetFlg = "nnd" Then
   mExchangeTypeID = GetExchangeID("Non NCDEX")
   If MSHFlexGrid1.TextMatrix(mRow, 0) <> "" Then
      For c = 1 To RowCount - 1
          If MSHFlexGrid1.TextMatrix(c, 0) = "" Then Exit For
          mRow = mRow + 1
          Status = 0
          Wc = ""
          mTxnType = "D"
          mCMPID = GetCMPIDEHOPID(Trim(MSHFlexGrid1.TextMatrix(c, 0)))
           'IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
          If mCMPID <> 0 Then
             mLocationID = GetLocationIDFromCMP(mCMPID)
          Else
             ErrMsg = ErrMsg + "CMP " & MSHFlexGrid1.TextMatrix(c, 0) & " At Row " & c & " not found in CMP Master" & vbCrLf
             Status = 1
          End If
                             
          If (Trim(MSHFlexGrid1.TextMatrix(c, 3))) <> "" Then
             mClientIDRow = GetClientIDRowByNameAndExchange(Trim(MSHFlexGrid1.TextMatrix(c, 3)), 1)
             If mClientIDRow = 0 Then
                ErrMsg = ErrMsg + "Client " & MSHFlexGrid1.TextMatrix(c, 3) & " At Row " & c & " not found in Client Master" & vbCrLf
                Status = 1
             End If
          Else
             mClientIDRow = Null
          End If

          If (Trim(MSHFlexGrid1.TextMatrix(c, 4))) <> "" Then
             mCommodityID = GetCommodityID(Trim(MSHFlexGrid1.TextMatrix(c, 4)))
             If mCommodityID = 0 Then
                ErrMsg = ErrMsg + "Commodity " & MSHFlexGrid1.TextMatrix(c, 4) & " At Row " & c & " not found in Commodity Master" & vbCrLf
                Status = 1
             End If
          Else
             mCommodityID = Null
          End If
            
          If Trim(MSHFlexGrid1.TextMatrix(c, 16)) <> "" Then
             mPINo = Trim(MSHFlexGrid1.TextMatrix(c, 16))
          Else
             mPINo = GetMaxNewPINo
          End If
         
          If Status = 0 Then
             RsTxn_IncomeDetails.AddNew
             RsTxn_IncomeDetails!Mode = "Non-NCDEX Deposits"
             RsTxn_IncomeDetails!TxnType = "D"
             RsTxn_IncomeDetails!LocationID = mLocationID
             RsTxn_IncomeDetails!ExchangeTypeID = mExchangeTypeID
             RsTxn_IncomeDetails!CommodityID = mCommodityID
             RsTxn_IncomeDetails!ClientIDRow = mClientIDRow
             RsTxn_IncomeDetails!BillingMonthStartDate = mStartDate  'Trim(MSHFlexGrid1.TextMatrix(c, 2))
             RsTxn_IncomeDetails!BillingMonthEndDate = mEndDate 'Trim(MSHFlexGrid1.TextMatrix(c, 2))
             RsTxn_IncomeDetails!CxTFNo = Trim(MSHFlexGrid1.TextMatrix(c, 5))
             RsTxn_IncomeDetails!CxtfDate = Trim(MSHFlexGrid1.TextMatrix(c, 6))  'CDTF Date
             RsTxn_IncomeDetails!NoOfBagsDeposited = Val(Trim(MSHFlexGrid1.TextMatrix(c, 7)))
             RsTxn_IncomeDetails!Rate = Val(Trim(MSHFlexGrid1.TextMatrix(c, 9)))
             RsTxn_IncomeDetails!StorageCharges = Val(Trim(MSHFlexGrid1.TextMatrix(c, 11))) + Val(Trim(MSHFlexGrid1.TextMatrix(c, 14)))
             RsTxn_IncomeDetails!ServiceTax = Val(Trim(MSHFlexGrid1.TextMatrix(c, 13)))
             RsTxn_IncomeDetails!InsuranceCharges = Val(Trim(MSHFlexGrid1.TextMatrix(c, 14)))
             RsTxn_IncomeDetails!AmtinclofSerTaxNInsurance = Val(Trim(MSHFlexGrid1.TextMatrix(c, 15)))
             RsTxn_IncomeDetails!AmtInclOfSerTax = Val(Trim(MSHFlexGrid1.TextMatrix(c, 15)))
             If Trim(MSHFlexGrid1.TextMatrix(c, 16)) <> "" Then
                RsTxn_IncomeDetails!OldPINo = mPINo 'Trim(MSHFlexGrid1.TextMatrix(c, 19)) 'PI_No
             Else
                RsTxn_IncomeDetails!NewPINo = mPINo
             End If
             RsTxn_IncomeDetails!Remarks = (Trim(MSHFlexGrid1.TextMatrix(c, 18)))
             RsTxn_IncomeDetails!G_UID = GetGUID
             RsTxn_IncomeDetails!CreatedBy = Gen_UserLoginID
             RsTxn_IncomeDetails!CreatedDate = Now
             RsTxn_IncomeDetails.Update
             mRowSaved = mRowSaved + 1
          End If
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End If
End If
''========================================================
If mSheetFlg = "nnw" Then
   mRow = mRow + 1
   mExchangeTypeID = GetExchangeID("Non NCDEX")
   If MSHFlexGrid1.TextMatrix(mRow, 0) <> "" Then
      For c = 1 To RowCount - 1
          If MSHFlexGrid1.TextMatrix(c, 0) = "" Then Exit For
          Status = 0
          Wc = ""
          mCMPID = GetCMPIDEHOPID(Trim(MSHFlexGrid1.TextMatrix(c, 0)))
           'IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
          If mCMPID <> 0 Then
             mLocationID = GetLocationIDFromCMP(mCMPID)
          Else
             ErrMsg = ErrMsg + "CMP " & MSHFlexGrid1.TextMatrix(c, 0) & " At Row " & c & " not found in CMP Master" & vbCrLf
             Status = 1
          End If
                       
          If (Trim(MSHFlexGrid1.TextMatrix(c, 5))) <> "" Then
             mClientIDRow = GetClientIDRowByNameAndExchange(Trim(MSHFlexGrid1.TextMatrix(c, 5)), 1)
             If mClientIDRow = 0 Then
                ErrMsg = ErrMsg + "Client " & MSHFlexGrid1.TextMatrix(c, 5) & " At Row " & c & " not found in Client Master" & vbCrLf
                Status = 1
             End If
          Else
             mClientIDRow = Null
          End If

          If (Trim(MSHFlexGrid1.TextMatrix(c, 6))) <> "" Then
             mCommodityID = GetCommodityID(Trim(MSHFlexGrid1.TextMatrix(c, 6)))
             If mCommodityID = 0 Then
                ErrMsg = ErrMsg + "Commodity " & MSHFlexGrid1.TextMatrix(c, 6) & " At Row " & c & " not found in Commodity Master" & vbCrLf
                Status = 1
             End If
          Else
             mCommodityID = Null
          End If
           
          If Trim(MSHFlexGrid1.TextMatrix(c, 23)) <> "" Then
             mPINo = Trim(MSHFlexGrid1.TextMatrix(c, 23))
          Else
             mPINo = GetMaxNewPINo
          End If
          
                  
          If Status = 0 Then
             RsTxn_IncomeDetails.AddNew
             RsTxn_IncomeDetails!Mode = "Non-NCDEX Withdrawals"
             RsTxn_IncomeDetails!TxnType = "W"
             RsTxn_IncomeDetails!LocationID = mLocationID
             RsTxn_IncomeDetails!CommodityID = mCommodityID
             RsTxn_IncomeDetails!ClientIDRow = mClientIDRow
             RsTxn_IncomeDetails!ExchangeTypeID = mExchangeTypeID
             RsTxn_IncomeDetails!BillingMonthStartDate = mStartDate 'Trim(MSHFlexGrid1.TextMatrix(c, 2))
             RsTxn_IncomeDetails!BillingMonthEndDate = mEndDate
             RsTxn_IncomeDetails!CxTFNo = Trim(MSHFlexGrid1.TextMatrix(c, 3))
             RsTxn_IncomeDetails!CxtfDate = (Trim(MSHFlexGrid1.TextMatrix(c, 4)))
             RsTxn_IncomeDetails!NoofBagsWithdrawn = Val(Trim(MSHFlexGrid1.TextMatrix(c, 9)))
             RsTxn_IncomeDetails!StorageCharges = Val(Trim(MSHFlexGrid1.TextMatrix(c, 18))) + Val(Trim(MSHFlexGrid1.TextMatrix(c, 21)))
             RsTxn_IncomeDetails!ServiceTax = Val(Trim(MSHFlexGrid1.TextMatrix(c, 20)))
             RsTxn_IncomeDetails!InsuranceCharges = Val(Trim(MSHFlexGrid1.TextMatrix(c, 21)))
             RsTxn_IncomeDetails!AmtinclofSerTaxNInsurance = Val(Trim(MSHFlexGrid1.TextMatrix(c, 22)))
             RsTxn_IncomeDetails!AmtInclOfSerTax = Val(Trim(MSHFlexGrid1.TextMatrix(c, 22)))
             If Trim(MSHFlexGrid1.TextMatrix(c, 23)) <> "" Then
                RsTxn_IncomeDetails!OldPINo = mPINo 'Trim(MSHFlexGrid1.TextMatrix(c, 19)) 'PI_No
             Else
                RsTxn_IncomeDetails!NewPINo = mPINo
             End If
             RsTxn_IncomeDetails!Remarks = (Trim(MSHFlexGrid1.TextMatrix(c, 25)))
             RsTxn_IncomeDetails!G_UID = GetGUID
             RsTxn_IncomeDetails!CreatedBy = Gen_UserLoginID
             RsTxn_IncomeDetails!CreatedDate = Now
             RsTxn_IncomeDetails.Update
             mRowSaved = mRowSaved + 1
          End If
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
  End If
End If
'=========================================================================================
If mSheetFlg = "nw" Then
   mExchangeTypeID = GetExchangeID("NCDEX")
   mRow = mRow + 1
   If MSHFlexGrid1.TextMatrix(mRow, 1) <> "" Then
      For c = 1 To RowCount
          If MSHFlexGrid1.TextMatrix(c, 0) = "" Then Exit For
          Wc = ""
          Status = 0
          mCMPID = GetCMPIDEHOPID(Trim(MSHFlexGrid1.TextMatrix(c, 0)))
           'IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
          If mCMPID <> 0 Then
             mLocationID = GetLocationIDFromCMP(mCMPID)
          Else
             ErrMsg = ErrMsg + "CMP " & MSHFlexGrid1.TextMatrix(c, 0) & " At Row " & c & " not found in CMP Master" & vbCrLf
             Status = 1
          End If

          If (Trim(MSHFlexGrid1.TextMatrix(c, 4))) <> "" Then
             mClientIDRow = GetClientIDRowByNameAndExchange(Trim(MSHFlexGrid1.TextMatrix(c, 4)), 1)
             If mClientIDRow = 0 Then
                ErrMsg = ErrMsg + "Client " & MSHFlexGrid1.TextMatrix(c, 4) & " At Row " & c & " not found in Client Master" & vbCrLf
                Status = 1
             End If
          Else
             mClientIDRow = Null
          End If
          
          If (Trim(MSHFlexGrid1.TextMatrix(c, 5))) <> "" Then
             mCommodityID = GetCommodityID(Trim(MSHFlexGrid1.TextMatrix(c, 5)))
             If mCommodityID = 0 Then
                ErrMsg = ErrMsg + "Commodity " & MSHFlexGrid1.TextMatrix(c, 5) & " At Row " & c & " not found in Commodity Master" & vbCrLf
                Status = 1
             End If
          Else
             mCommodityID = Null
          End If
            
        
          If Trim(MSHFlexGrid1.TextMatrix(c, 18)) <> "" Then
             mPINo = Trim(MSHFlexGrid1.TextMatrix(c, 18))
          Else
             mPINo = GetMaxNewPINo
          End If
          
          If Status = 0 Then
             RsTxn_IncomeDetails.AddNew
             RsTxn_IncomeDetails!Mode = "NCDEX Withdrawals"
             RsTxn_IncomeDetails!TxnType = "W"
             RsTxn_IncomeDetails!LocationID = mLocationID
             RsTxn_IncomeDetails!CommodityID = mCommodityID
             RsTxn_IncomeDetails!ClientIDRow = mClientIDRow
             RsTxn_IncomeDetails!ExchangeTypeID = mExchangeTypeID
             RsTxn_IncomeDetails!BillingMonthStartDate = mStartDate 'Trim(MSHFlexGrid1.TextMatrix(c, 2))
             RsTxn_IncomeDetails!BillingMonthEndDate = mEndDate
             RsTxn_IncomeDetails!CxTFNo = (Trim(MSHFlexGrid1.TextMatrix(c, 2)))
             RsTxn_IncomeDetails!CxtfDate = (Trim(MSHFlexGrid1.TextMatrix(c, 3)))
             RsTxn_IncomeDetails!Rate = Val(Trim(MSHFlexGrid1.TextMatrix(c, 11)))
             RsTxn_IncomeDetails!StorageCharges = Val(Trim(MSHFlexGrid1.TextMatrix(c, 14)))
             RsTxn_IncomeDetails!ServiceTax = Val(Trim(MSHFlexGrid1.TextMatrix(c, 16)))
             RsTxn_IncomeDetails!AmtInclOfSerTax = Val(Trim(MSHFlexGrid1.TextMatrix(c, 17)))
             RsTxn_IncomeDetails!Remarks = (Trim(MSHFlexGrid1.TextMatrix(c, 28)))
             If Trim(MSHFlexGrid1.TextMatrix(c, 18)) <> "" Then
                RsTxn_IncomeDetails!OldPINo = mPINo 'Trim(MSHFlexGrid1.TextMatrix(c, 19)) 'PI_No
             Else
                RsTxn_IncomeDetails!NewPINo = mPINo
             End If
             RsTxn_IncomeDetails!G_UID = GetGUID
             RsTxn_IncomeDetails!CreatedBy = Gen_UserLoginID
             RsTxn_IncomeDetails!CreatedDate = Now
             RsTxn_IncomeDetails.Update
             mRowSaved = mRowSaved + 1
          End If
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
  End If
End If
ProgressBar1.Value = ProgressBar1.Max
TxtErrorMsg = ErrMsg  ', vbCritical
MsgBox "Data saved To database Successfully for " & mRowSaved & " rows except rows in error message!!!", vbInformation
oXL.Quit

End Sub

Private Sub Form_Load()

lblFile.Caption = ""
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

tmp = "Select * from Mst_Config"
Call TmpTable

mStartDate = TmpRs!IncomeStartDate
mMonth = Month(mStartDate)
mEndDate = GetEndDate(mMonth, mStartDate)

End Sub

Private Function GetMaxNewPINo() As Double
Dim Retval As Double
tmp = "Select Max(NewPINo) From Txn_IncomeDetails "
Call Tmp1Table
Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
Retval = Retval + 1
TmpRs1.Close
GetMaxNewPINo = Retval
End Function


