VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_KRPendingUpload 
   Caption         =   "KR Pending Upload"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   7560
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   3045
         TabIndex        =   6
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   4290
         TabIndex        =   4
         Top             =   240
         Width           =   1245
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   765
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   45
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label lblProcess 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   6855
      End
      Begin VB.Label lblProgress 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   6855
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_KRPendingUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilename As String
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Integer
Dim mCol As Integer
Private Sub CmdExport_Click()
Call Import_Data
End Sub
Private Sub Import_Data()
Ans = Check_File
If Ans = True Then
   MsgBox "Import Successfull !!!! "
Else
   MsgBox "Error in Import !!!! "
End If
End Sub

Function Check_File() As Boolean
Dim fn, CFlag As Variant
fn = strFilename
Dim objFileSystem As Object
Dim FileToCheck As String
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
CFlag = True
FileToCheck = strFilename 'App.Path & "\Excel\" & fn & ".xls"
If objFileSystem.FileExists(FileToCheck) Then
Else
    MsgBox "File not found !!!! "
    CFlag = False
End If

If CFlag = True Then

   Set oXL = New Excel.Application
   Dim Pat As String
   fn = GetFileFromPath(strFilename)
   Pat = App.Path
   Set oWB = oXL.Workbooks.Open(strFilename) '(Pat & "\excel\" & fn & ".xls", , , , "")
   Set oWS = oWB.Worksheets(1)
   If oWS.Name <> "Sheet1" Then
      CFlag = False
   Else
      If CheckFile = False Then
         MsgBox "Please Select proper file for uploading!!!"
         CmdImport.Enabled = False
         CFlag = False
      End If
   End If
   Check_File = CFlag
End If

If CFlag = True Then
   Dim i As Variant
    ' ---------------- GSL Master Import
   Call TableMst_GSL(" Where CommodityID in (Select CommodityID from Mst_Commodity Where CommTrackFlag = 0) ")
   Set oWS = oWB.Worksheets(1)
   i = 2
   x = 0
   ProgressBar1.Value = 0
   ProgressBar1.Max = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row
   Do
       If oWS.Cells(i, 1) = "" Then
          Exit Do
       End If
       If RsMst_GSL.RecordCount > 0 Then
          RsMst_GSL.MoveFirst
          RsMst_GSL.Filter = "SM_Prefix =  '" & oWS.Cells(i, 1) & "' And GSLID = " & oWS.Cells(i, 2) & " "
          If RsMst_GSL.RecordCount > 0 Then
             If oWS.Cells(i, 18) <> "" Then
                If oWS.Cells(i, 18) <> 0 Then
                   RsMst_GSL!KRNo = oWS.Cells(i, 18)
                   If oWS.Cells(i, 19) = "" Then
                      RsMst_GSL!KRDate = Date
                   Else
                      RsMst_GSL!KRDate = oWS.Cells(i, 19)
                   End If
                   RsMst_GSL.Update
                   x = x + 1
                End If
             End If
          End If
          RsMst_GSL.Filter = ""
          i = i + 1
          ProgressBar1.Value = ProgressBar1.Value + 1
       End If
   Loop Until False
'   lblProgress.Caption = "Completed : " & Int(ProgressBar1.Value * 100 / ProgressBar1.Max) & "%"
   lblProcess.Caption = "Buffering " & oWS.Name & " Records"
   lblProgress.Refresh
   lblProcess.Refresh
   Label1.Caption = x & " Records updated!!!"
End If

End Function

Private Sub CmdCancel_Click()
CmdImport.Enabled = False
CmdGetFile.Enabled = True
lblFile.Refresh
lblFile.Caption = ""
Label1.Caption = ""
ProgressBar1.Value = 0
End Sub

Private Sub CmdGetFile_Click()
With cdialog
   .InitDir = App.Path
   .DialogTitle = "Select File To Upload"
   .ShowOpen
End With
If cdialog.CancelError Then
   lblFile.Refresh
   lblFile.Caption = ""
   CmdImport.Enabled = False
Else
   strFilename = cdialog.FileName
   lblFile.Refresh
   lblFile.Caption = strFilename 'GetFileFromPath(strFilename)
   If strFilename <> "" Then
      CmdImport.Enabled = True
   End If
End If
'If Left(GetFileFromPath(strFilename), 9) = "BASE FILE" Then
'    CmdGetFile.Enabled = False

'Else
'    lblFile.Caption = "Please Select the correct file "
'    CmdGetFile.Enabled = True
'    CmdImport.Enabled = False
'End If
End Sub

Private Sub CmdImport_Click()
Call Import_Data

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True


lblProcess.Caption = ""
lblProgress.Caption = ""
Label1.Caption = ""
End Sub
Function GetFileFromPath(vPath As String)
    Dim Items() As String
    Items = Split(vPath, "\")
    If UBound(Items) = -1 Then Exit Function
    GetFileFromPath = Items(UBound(Items))
End Function
Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True
If oWS.Cells(1, 1) <> "Location ID" Then
   Retval = False
ElseIf oWS.Cells(1, 2) <> "GSL ID." Then
   Retval = False
ElseIf oWS.Cells(1, 3) <> "CDTF No" Then
   Retval = False
ElseIf oWS.Cells(1, 4) <> "CDTF Date" Then
   Retval = False
ElseIf oWS.Cells(1, 5) <> "CMP" Then
   Retval = False
ElseIf oWS.Cells(1, 6) <> "Commodity" Then
   Retval = False
ElseIf oWS.Cells(1, 7) <> "dp_id" Then
   Retval = False
ElseIf oWS.Cells(1, 8) <> "CM_ClientID" Then
   Retval = False
ElseIf oWS.Cells(1, 9) <> "Client Name" Then
   Retval = False
ElseIf oWS.Cells(1, 10) <> "GSL No." Then
   Retval = False
ElseIf oWS.Cells(1, 11) <> "Godown Address" Then
   Retval = False
ElseIf oWS.Cells(1, 12) <> "No of Bags" Then
   Retval = False
ElseIf oWS.Cells(1, 13) <> "Weight in MT" Then
   Retval = False
ElseIf oWS.Cells(1, 14) <> "Depositor Address" Then
   Retval = False
ElseIf oWS.Cells(1, 15) <> "Depositor Name" Then
   Retval = False
ElseIf oWS.Cells(1, 16) <> "Amount Received" Then
   Retval = False
ElseIf oWS.Cells(1, 17) <> "SC upto Date" Then
   Retval = False
ElseIf oWS.Cells(1, 18) <> "KR No" Then
   Retval = False
ElseIf oWS.Cells(1, 19) <> "KR Date" Then
   Retval = False
ElseIf oWS.Cells(1, 20) <> "CDTF Agent Name" Then
   Retval = False
ElseIf oWS.Cells(1, 21) <> "TRF No" Then
   Retval = False
ElseIf oWS.Cells(1, 22) <> "TRF Date" Then
   Retval = False
ElseIf oWS.Cells(1, 23) <> "NCOM NCDEX CODE" Then
   Retval = False
End If

CheckFile = Retval

End Function

