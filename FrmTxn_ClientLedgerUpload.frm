VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmTxn_ClientLedgerUpload 
   Caption         =   "Client Ledger Upload"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15060
      Begin VB.CommandButton CmdGetfile 
         Caption         =   "Get File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6435
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   16
         Top             =   1320
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   14
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   435
      End
      Begin VB.Label lblLines 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   12
         Top             =   1680
         Width           =   60
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lines :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
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
         Left            =   7560
         TabIndex        =   10
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Entry No Already Exist "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   9
         Top             =   1770
         Width           =   1965
      End
      Begin VB.Label Label8 
         Caption         =   "NAV Code not mapped in Employee Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11205
         TabIndex        =   8
         Top             =   1770
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
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
         Left            =   10560
         TabIndex        =   7
         Top             =   1755
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   4920
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6840
      Left            =   120
      TabIndex        =   17
      Top             =   2295
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   12065
      _Version        =   393216
      WordWrap        =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   9195
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmTxn_ClientLedgerUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstemp As New ADODB.Recordset
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream

Dim objFScp As Scripting.FileSystemObject
Dim objFilecp As File
Dim objTScp As TextStream

Dim sFileLine As String
Dim sFileLineNo As Long

Dim temp As Variant
Dim arecordset As String
Dim strFilename As String
Dim strsql As String
Dim strdata As String
Dim t_line As String
Dim strfile As String
Dim strqry As String
Dim sArray(0 To 17) As String
Dim mEmployeeID As Variant

Private Sub CmdCancel_Click()
lblFile.Caption = ""
lblStatus.Caption = ""
lblLines.Caption = ""
strFilename = ""
CmdGetfile.Enabled = True
CmdUpdate.Enabled = False
CmdClose.Enabled = True
CmdCancel.Enabled = True
CmdSave.Enabled = False
CmdExcel.Enabled = False
Call Grid_Head
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub CmdGetFile_Click()
    
CmdGetfile.Enabled = False
CmdUpdate.Enabled = False
CmdClose.Enabled = False
CmdCancel.Enabled = True

lblStatus.Refresh
lblStatus.Caption = ""
lblFile.Refresh
lblFile.Caption = ""


With cdialog
   .InitDir = App.Path
   .DialogTitle = "Select Client Ledger File"
   .Filter = "TxT files(*.txt)|*.txt"
   .ShowOpen
End With
strFilename = cdialog.FileName
lblFile.Refresh
lblFile.Caption = strFilename
If cdialog.CancelError Then
   lblStatus.Refresh
   lblStatus.Caption = ""
Else
   lblStatus.Refresh
   lblStatus.Caption = "File Dectected..." & GetFileFromPath(strFilename)
'   If Left(GetFileFromPath(strFilename), 10) = "EMP LEDGER" Then
   CmdUpdate.Enabled = True
'   Else
'      lblStatus.Caption = "Please Select the correct file "
'      CmdGetfile.Enabled = True
'   End If
                  
End If
End Sub
Function GetFileFromPath(vPath As String)
    Dim Items() As String
    Items = Split(vPath, "\")
    If UBound(Items) = -1 Then Exit Function
    GetFileFromPath = Items(UBound(Items))
End Function

Private Sub CmdSave_Click()

CmdSave.Enabled = False

If MSHFlexGrid1.Rows < 3 Then
   MsgBox "No data to save!!!"
   Exit Sub
End If

If MSHFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
MsgBox "Wait till Next Message....."
 
Call TableMst_Client
Call TableMst_Depository
Call TableTxn_ClientLedger

ErrMsg = ""
RowCount = MSHFlexGrid1.Rows
ProgressBar1.Value = 0
ProgressBar1.Max = RowCount
lblStatus.Caption = "Please wait saving data..."
For i = 1 To RowCount - 2
    wc = ""
    mStatus = 0
    If MSHFlexGrid1.TextMatrix(i, 0) <> "" Then
       If RsTxn_ClientLedger.RecordCount > 0 Then
          RsTxn_ClientLedger.MoveFirst
       End If
       RsTxn_ClientLedger.Find "EntryNo = '" & MSHFlexGrid1.TextMatrix(i, 2) & "'"
       If RsTxn_ClientLedger.EOF Then
          RsMst_Client.MoveFirst
          RsMst_Client.Find "NAVCode  = '" & MSHFlexGrid1.TextMatrix(i, 0) & "'  "
          If Not RsMst_Client.EOF Then
             If IIf(IsNull(RsMst_Client!NAVClientName), "", RsMst_Client!NAVClientName) <> MSHFlexGrid1.TextMatrix(i, 1) Then
                RsMst_Client!NAVClientName = MSHFlexGrid1.TextMatrix(i, 1)
                RsMst_Client.Update
             End If
          Else
             RsMst_Depository.MoveFirst
             RsMst_Depository.Find "NAVCode  = '" & MSHFlexGrid1.TextMatrix(i, 0) & "'  "
             If Not RsMst_Depository.EOF Then
                If IIf(IsNull(RsMst_Depository!NAVDepositoryName), "", RsMst_Depository!NAVDepositoryName) <> MSHFlexGrid1.TextMatrix(i, 1) Then
                   RsMst_Depository!NAVDepositoryName = MSHFlexGrid1.TextMatrix(i, 1)
                   RsMst_Depository.Update
                End If
             End If
          End If
'             'mEmployeeID = RsMst_Employee!EmployeeID
'          Else
'             'mEmployeeID = Null
'             ErrMsg = ErrMsg & "Client not found in Client Master At Row " & I + 1 & vbCrLf
'             MSHFlexGrid1.Row = I
'             For c = 0 To MSHFlexGrid1.Col - 1
'                 MSHFlexGrid1.Col = c
'                 MSHFlexGrid1.CellBackColor = vbRed
'             Next
'          End If
           
          
          RsTxn_ClientLedger.AddNew
          RsTxn_ClientLedger!CustomerNo = Trim(MSHFlexGrid1.TextMatrix(i, 0))
          RsTxn_ClientLedger!CustomerName = Trim(MSHFlexGrid1.TextMatrix(i, 1))
          RsTxn_ClientLedger!EntryNo = MSHFlexGrid1.TextMatrix(i, 2)
          RsTxn_ClientLedger!EHOPInvoiceNo = Trim(MSHFlexGrid1.TextMatrix(i, 3))
          RsTxn_ClientLedger!DocumentType = Trim(MSHFlexGrid1.TextMatrix(i, 4))
          RsTxn_ClientLedger!DocumentNo = Trim(MSHFlexGrid1.TextMatrix(i, 5))
          RsTxn_ClientLedger!VoucherNo = Trim(MSHFlexGrid1.TextMatrix(i, 6))
          RsTxn_ClientLedger!PostingDate = CDate(MSHFlexGrid1.TextMatrix(i, 7))
          'RsTxn_ClientLedger!InvoiceAmount = Val(MSHFlexGrid1.TextMatrix(I, 8))
          'RsTxn_ClientLedger!BaseAmount = Val(MSHFlexGrid1.TextMatrix(I, 9))
          'RsTxn_ClientLedger!TaxAmount = Val(MSHFlexGrid1.TextMatrix(I, 10))
          'RsTxn_ClientLedger!RemainingAmount = Val(MSHFlexGrid1.TextMatrix(I, 11))
          
          RsTxn_ClientLedger!InvoiceAmount = Val(Replace(MSHFlexGrid1.TextMatrix(i, 8), ",", ""))
          RsTxn_ClientLedger!BaseAmount = Val(Replace(MSHFlexGrid1.TextMatrix(i, 9), ",", ""))
          RsTxn_ClientLedger!TaxAmount = Val(Replace(MSHFlexGrid1.TextMatrix(i, 10), ",", ""))
          RsTxn_ClientLedger!RemainingAmount = Val(Replace(MSHFlexGrid1.TextMatrix(i, 11), ",", ""))
          
          
          
          RsTxn_ClientLedger!Narration = Trim(MSHFlexGrid1.TextMatrix(i, 12))
          RsTxn_ClientLedger!ActualDatePosted = CDate(MSHFlexGrid1.TextMatrix(i, 13))
          RsTxn_ClientLedger!ActualTimePosted = MSHFlexGrid1.TextMatrix(i, 14)
          RsTxn_ClientLedger!RegionCode = Trim(MSHFlexGrid1.TextMatrix(i, 15))
          RsTxn_ClientLedger!DepartmentCode = Trim(MSHFlexGrid1.TextMatrix(i, 16))
          RsTxn_ClientLedger!CustomerPostingGroup = Trim(MSHFlexGrid1.TextMatrix(i, 17))
          RsTxn_ClientLedger!CreatedBy = Gen_UserLoginID
          RsTxn_ClientLedger!CreatedDate = Now 'Trim(MSHFlexGrid1.TextMatrix(I, 0))
          RsTxn_ClientLedger!G_UID = GetGUID
          RsTxn_ClientLedger.Update
       Else
          ErrMsg = ErrMsg & "Entry No already exist Client Ledger At Row " & i + 1 & vbCrLf
          'MSHFlexGrid1.Col = 0
          MSHFlexGrid1.Row = i
          MSHFlexGrid1.CellBackColor = vbGreen
          For c = 0 To MSHFlexGrid1.Col - 1
              MSHFlexGrid1.Col = c
              MSHFlexGrid1.CellBackColor = vbGreen
          Next
       End If
    Else
       ErrMsg = ErrMsg & "Entry No Not Found in Grid " & MSHFlexGrid1.TextMatrix(i, 0) & "  At Row " & i + 1 & vbCrLf
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.Row = i
       MSHFlexGrid1.CellBackColor = vbRed
          'mStatus = 1
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max
lblStatus.Caption = "Done"
MsgBox "Done!!!"
CmdExcel.Enabled = True

End Sub

Private Sub CmdUpdate_Click()
Call Grid_Head
'CmdGetFile.Enabled = False
'cmdUpdate.Enabled = False
'cmdClose.Enabled = False
'CmdCancel.Enabled = False
              
lblStatus.Refresh
Call ReadFiles(strFilename)
lblStatus.Caption = ""
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
'   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Call CmdCancel_Click
End Sub

Public Sub ReadFiles(file_name As String)
   
Dim lngFreefile As Long
'Call TableMst_EmpLedger

Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(file_name)
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
  
If CheckFile(file_name) = False Then
   MsgBox "Please Select proper file!!!"
   Exit Sub
End If
  
sFileLineNo = 0

lblStatus.Refresh
'lblStatus.Caption = "Reading File These will take few mins..Please Wait"
ProgressBar1.Value = 0
ProgressBar1.Max = 10000 'objTS.AtEndOfStream
Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   tmp = Chr(34) + "," & Chr(34)
   sFileLine = Replace(sFileLine, tmp, "~")
   sFileLine = Replace(sFileLine, Chr(34), "")
  
   sFileLineNo = sFileLineNo + 1
    'tmp = Chr(59)
'   If Left(sFileLine, 3) = "EMP" Then
   sFileLine = Replace(sFileLine, "'", " ")
   ps = Split(sFileLine, "~")
   If UBound(ps) = 17 Then
      L = L + 1
      For c = 0 To MSHFlexGrid1.Cols - 1
          If InStr(Trim(ps(c)), "/") > 0 Then
             If IsDate(Trim(ps(c))) Then
                mDate = Split(Trim(ps(c)), "/")
               'dd/nn/yy
                'StrDate = mdate(0) & "/" & mdate(1) & "/" & "20" & mdate(2)
                If Len(Trim(ps(c))) = "8" Then
                   StrDate = "20" & mDate(2) & "/" & mDate(1) & "/" & mDate(0)
                Else
                   StrDate = mDate(2) & "/" & mDate(1) & "/" & mDate(0)
                End If
                MSHFlexGrid1.TextMatrix(L, c) = StrDate 'Format(Trim(ps(c)), "dd/mm/yyyy")
              Else
                MSHFlexGrid1.TextMatrix(L, c) = Trim(ps(c))
              End If
           Else
              MSHFlexGrid1.TextMatrix(L, c) = Trim(ps(c))
           End If
         'MSHFlexGrid1.TextMatrix(L, c) = Trim(ps(c)) 'Trim(oWS.Cells(mRow, c + 1))
      Next
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   End If
      lblLines.Caption = sFileLineNo
'   End If
   If ProgressBar1.Value = ProgressBar1.Max Then
      ProgressBar1.Value = 0
   End If
   ProgressBar1.Value = ProgressBar1.Value + 1
Loop
ProgressBar1.Value = ProgressBar1.Max
lblStatus.Refresh
MsgBox "File Imported Sucessfully...."
ProgressBar1.Value = 0
CmdGetfile.Enabled = False
CmdUpdate.Enabled = False
CmdClose.Enabled = True
CmdCancel.Enabled = True
CmdSave.Enabled = True
CmdExcel.Enabled = True

End Sub

Private Function CheckFile(file_name) As Boolean
Dim Retval As Boolean
Retval = True

'Set objFS = New Scripting.FileSystemObject
'Set objFile = objFS.GetFile(file_name)
'Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
  
sFileLineNo = 0

lblStatus.Refresh
lblStatus.Caption = "Reading File These will take few mins..Please Wait"

sFileLine = objTS.ReadLine
tmp = Chr(34) + "," & Chr(34)
sFileLine = Replace(sFileLine, tmp, "~")
sFileLine = Replace(sFileLine, Chr(34), "")

sFileLine = Replace(sFileLine, "'", " ")
ps = Split(sFileLine, "~")
If UBound(ps) = 17 Then
   For i = 0 To 17
       sArray(i) = Trim(ps(i))
   Next i
Else
   Retval = False
End If     'If RsDistrict.RecordCount > 0 Then

'sFileLineNo = sFileLineNo + 1
  
'If LCase(sArray(0)) <> "vendor code" Then
'   Retval = False
'ElseIf LCase(sArray(1)) <> "vendor name" Then
'   Retval = False
'ElseIf LCase(sArray(2)) <> "document type" Then
'   Retval = False
'ElseIf LCase(sArray(3)) <> "document no." Then
'   Retval = False
'ElseIf LCase(sArray(4)) <> "transaction date" Then
'   Retval = False
'ElseIf LCase(sArray(5)) <> "narration" Then
'   Retval = False
'ElseIf LCase(sArray(6)) <> "transaction amount" Then
'   Retval = False
'ElseIf LCase(sArray(7)) <> "remaining amount" Then
'   Retval = False
'ElseIf LCase(sArray(8)) <> "acutal posting date" Then
'   Retval = False
'ElseIf LCase(sArray(9)) <> "acutal posting time" Then
'   Retval = False
'ElseIf LCase(sArray(10)) <> "entry no." Then
'   Retval = False
'End If

If LCase(sArray(0)) <> "customer no." Then
   Retval = False
ElseIf LCase(sArray(1)) <> "customer name" Then
   Retval = False
ElseIf LCase(sArray(2)) <> "entry no." Then
   Retval = False
ElseIf LCase(sArray(3)) <> "ehop invoice no." Then
   Retval = False
ElseIf LCase(sArray(4)) <> "document type" Then
   Retval = False
ElseIf LCase(sArray(5)) <> "document no." Then
   Retval = False
ElseIf LCase(sArray(6)) <> "voucher no." Then
   Retval = False
ElseIf LCase(sArray(7)) <> "posting date" Then
   Retval = False
ElseIf LCase(sArray(8)) <> "invoice amount" Then
   Retval = False
ElseIf LCase(sArray(9)) <> "base amount" Then
   Retval = False
ElseIf LCase(sArray(10)) <> "tax amount" Then
   Retval = False
ElseIf LCase(sArray(11)) <> "remaining amount" Then
   Retval = False
ElseIf LCase(sArray(12)) <> "narration" Then
   Retval = False
ElseIf LCase(sArray(13)) <> "actual date posted" Then
   Retval = False
ElseIf LCase(sArray(14)) <> "actual time posted" Then
   Retval = False
ElseIf LCase(sArray(15)) <> "region code" Then
   Retval = False
ElseIf LCase(sArray(16)) <> "department code" Then
   Retval = False
ElseIf LCase(sArray(17)) <> "customer posting group" Then
   Retval = False
End If

CheckFile = Retval

End Function

Private Sub Grid_Head()
    
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 18
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.AllowUserResizing = flexResizeColumns

'MSHFlexGrid1.TextMatrix(0, 0) = "Vendor Code"
'MSHFlexGrid1.TextMatrix(0, 1) = "Vendor Name"
'MSHFlexGrid1.TextMatrix(0, 2) = "DocumentType"
'MSHFlexGrid1.TextMatrix(0, 3) = "Document No"
'MSHFlexGrid1.TextMatrix(0, 4) = "Transaction Date"
'MSHFlexGrid1.TextMatrix(0, 5) = "Narration"
'MSHFlexGrid1.TextMatrix(0, 6) = "Transaction Amount"
'MSHFlexGrid1.TextMatrix(0, 7) = "Remaining Amount"
'MSHFlexGrid1.TextMatrix(0, 8) = "AcutalPosting Date"
'MSHFlexGrid1.TextMatrix(0, 9) = "AcutalPosting Time"
'MSHFlexGrid1.TextMatrix(0, 10) = "Entry No"


MSHFlexGrid1.TextMatrix(0, 0) = "Customer No."
MSHFlexGrid1.TextMatrix(0, 1) = "Customer Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Entry No."
MSHFlexGrid1.TextMatrix(0, 3) = "EHOP Invoice No."
MSHFlexGrid1.TextMatrix(0, 4) = "Document Type"
MSHFlexGrid1.TextMatrix(0, 5) = "Document No."
MSHFlexGrid1.TextMatrix(0, 6) = "Voucher No."
MSHFlexGrid1.TextMatrix(0, 7) = "Posting Date"
MSHFlexGrid1.TextMatrix(0, 8) = "Invoice Amount"
MSHFlexGrid1.TextMatrix(0, 9) = "Base Amount"
MSHFlexGrid1.TextMatrix(0, 10) = "Tax Amount"
MSHFlexGrid1.TextMatrix(0, 11) = "Remaining Amount"
MSHFlexGrid1.TextMatrix(0, 12) = "Narration"
MSHFlexGrid1.TextMatrix(0, 13) = "Actual Date Posted"
MSHFlexGrid1.TextMatrix(0, 14) = "Actual Time Posted"
MSHFlexGrid1.TextMatrix(0, 15) = "Region Code"
MSHFlexGrid1.TextMatrix(0, 16) = "Department Code"
MSHFlexGrid1.TextMatrix(0, 17) = "Customer Posting Group"


MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 2000
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1400
MSHFlexGrid1.ColWidth(15) = 1000
MSHFlexGrid1.ColWidth(16) = 1000
MSHFlexGrid1.ColWidth(17) = 1000
   
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ClientLedger.xls")
Call Xls_Detail_Rpt
CmdExcel.Enabled = False
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "ClientLedger.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
        
        MSHFlexGrid1.Row = i
        MSHFlexGrid1.Col = c
        
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbRed
        ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbYellow
        ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbGreen
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

