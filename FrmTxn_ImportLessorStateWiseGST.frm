VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTxn_ImportLessorStateWiseGST 
   Caption         =   "Import Lessor State Wise GST"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   15015
      Begin VB.CommandButton CmdGet 
         Caption         =   "Get Data"
         Height          =   345
         Left            =   7200
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox CmbLessor 
         Height          =   360
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   9435
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
         Height          =   345
         Left            =   8280
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
         Enabled         =   0   'False
         Height          =   345
         Left            =   11565
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   345
         Left            =   13755
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
         Enabled         =   0   'False
         Height          =   345
         Left            =   9375
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   10470
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   12660
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxtNAVCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   12600
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   11
         Top             =   9360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7740
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13653
         _Version        =   393216
         AllowUserResizing=   3
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
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lessor Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label LblTotalRecords 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   14070
         TabIndex        =   15
         Top             =   1150
         Width           =   90
      End
      Begin VB.Label Label2 
         Caption         =   "Total Record :"
         Height          =   255
         Left            =   12600
         TabIndex        =   14
         Top             =   1150
         Width           =   1335
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1150
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NAV Code"
         Height          =   240
         Left            =   11400
         TabIndex        =   12
         Top             =   300
         Width           =   960
      End
   End
End
Attribute VB_Name = "FrmTxn_ImportLessorStateWiseGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet

Dim ErrMsg As Variant
Dim Status As Variant
Dim mLessorID As Variant
Dim ErrMsgFile As String

''-- Ullhas Patil 08 Nov 2017
Dim mAllowUpdateFlag As Boolean
Dim mNewEntry As Boolean

Private Sub CmbLessor_GotFocus()
tmp = CmbLessor.Text
mTmp = CmbLessor.Text
Call TableMst_Lessor
CmbLessor.Clear
If RsMst_Lessor.RecordCount = 0 Then
   MsgBox "No Lessor found"
   CmbLessor.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Lessor.RecordCount
    CmbLessor.AddItem RsMst_Lessor!LessorName & "~" & RsMst_Lessor!LessorID
    RsMst_Lessor.MoveNext
Next
CmbLessor.Text = tmp
End Sub

Private Sub CmbLessor_LostFocus()
Dim mSplit As Variant
If CmbLessor.Text = "" Then
   Exit Sub
End If
mSplit = Split(CmbLessor.Text, "~")
If UBound(mSplit) <> 1 Then
   MsgBox "Invalid selection "
   CmbLessor.SetFocus
   Exit Sub
End If
RsMst_Lessor.MoveFirst
RsMst_Lessor.Find "LessorID = '" & mSplit(1) & "'"
If RsMst_Lessor.EOF Then
   MsgBox "Invalid selection "
   CmbLessor.SetFocus
   Exit Sub
Else
   If LCase(RsMst_Lessor!LessorName) <> LCase(mSplit(0)) Then
    MsgBox "Invalid selection "
    CmbLessor.SetFocus
    Exit Sub
   End If
End If
mLessorID = IIf(IsNull(RsMst_Lessor!LessorID), "", RsMst_Lessor!LessorID)
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_LessorStateWiseGST.xls")
Call Xls_Detail_Rpt
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGet_Click()
Dim wc As Variant
If Trim(TxtNAVCode.Text) = "" Then
    If CmbLessor.Text = "" Then
        MsgBox "Select Lessor"
        Exit Sub
    End If
End If
Call Grid_Head

wc = ""
If Trim(TxtNAVCode.Text) = "" Then
    wc = mLessorID
Else
    wc = " SELECT LessorID FROM Mst_Lessor Where NAVCode = '" & Trim(TxtNAVCode) & "'"
End If

tmp = " Select A.LessorStateWiseGSTID,A.LessorID,A.LessorName,A.NAVCode,A.StateIndiaID,A.StateIndiaName,A.LessorStateOfficeAddress,A.GSTNo,A.GSTNavCode From"
tmp = tmp & " (Select MLSWG.LessorStateWiseGSTID,MLSWG.LessorID,ML.LessorName,ML.NAVCode,MLSWG.StateIndiaID,MSI.StateIndiaName,MLSWG.LessorStateOfficeAddress,MLSWG.GSTNo,MLSWG.GSTNavCode"
tmp = tmp & " From Mst_LessorStateWiseGST MLSWG"
tmp = tmp & " Inner Join Mst_Lessor ML On MLSWG.LessorID = ML.LessorID"
tmp = tmp & " Inner Join Mst_StateIndia MSI On MLSWG.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Where MLSWG.LessorID IN (" & wc & ")"
tmp = tmp & " Union"
tmp = tmp & " SELECT '0'LessorStateWiseGSTID,ML.LessorID,ML.LessorName,ML.NAVCode,MSI.StateIndiaID,MSI.StateIndiaName,''LessorStateOfficeAddress,''GSTNo,''GSTNavCode"
tmp = tmp & " FROM Mst_Lessor ML ,Mst_StateIndia MSI"
tmp = tmp & " Where ML.LessorID IN ( " & wc & ") And Convert(Varchar(18),ML.LessorID)+'~'+Convert(Varchar(18),MSI.StateIndiaID) Not in (SELECT Convert(Varchar(18),LessorID) +'~'+Convert(Varchar(18),StateIndiaID) FROM Mst_LessorStateWiseGST Where LessorID IN (SELECT LessorID FROM Mst_Lessor Where LessorID IN (" & wc & ")))"
tmp = tmp & " ) AS A Order By A.LessorName,A.LessorID,A.StateIndiaName,A.LessorStateWiseGSTID"
Call TmpTable

If TmpRs.EOF = False Then
    
    With MSHFlexGrid1
    If TmpRs.RecordCount > 0 Then
       ProgressBar1.Max = TmpRs.RecordCount
       For mR = 1 To TmpRs.RecordCount
              
           .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!LessorName), "", TmpRs!LessorName)
           .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
           .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
           .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!LessorStateOfficeAddress), "", TmpRs!LessorStateOfficeAddress)
           .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
           .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
           .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!LessorID), "", TmpRs!LessorID)
           .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
           .TextMatrix(mR, 8) = IIf(TmpRs!LessorStateWiseGSTID = 0, "", TmpRs!LessorStateWiseGSTID)

           .Rows = .Rows + 1
           TmpRs.MoveNext
           ProgressBar1.Value = mR
       Next
       CmdExcel.Enabled = True
    End If
    End With
    LblTotalRecords.Caption = TmpRs.RecordCount
End If
End Sub

Private Sub CmdGetFile_Click()
On Error GoTo errortrap

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx"
     .CancelError = False
procReOpen:
     .ShowOpen
End With

Path = CommonDialog1.FileName
LblFile.Caption = Path

If LblFile.Caption <> "" Then
   CmdExit.Enabled = True
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   Dim i As Integer
   Dim runApp As Boolean
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   oXL.Quit
   MsgBox "Select Import !!!"
   CmdImport.SetFocus
End If

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub CmdImport_Click()
Call Grid_Head
MsgBox "Wait till next message!!!"
Status = 0
RowCount = 0
Set oXL = New Excel.Application
Set oWB = oXL.Workbooks.Open(LblFile.Caption)
Set oWS = oWB.Worksheets(1)

If CheckFile = False Then
   MsgBox ErrMsgFile, vbOKOnly + vbCritical
   oXL.Quit
   Call CmdReset_Click
   Exit Sub
End If

Call FillData
''-- Add Two more columns for Errors and Update Log
MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + 1
MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Cols - 1) = "Errors"
MSHFlexGrid1.ColWidth(MSHFlexGrid1.Cols - 1) = 1500
MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + 1
MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Cols - 1) = "Update Log"
MSHFlexGrid1.ColWidth(MSHFlexGrid1.Cols - 1) = 1500
''---------------------------------------
CmdGetFile.Enabled = False
CmdImport.Enabled = False
CmdSave.Enabled = True
CmdReset.Enabled = True
CmdSave.SetFocus
ProgressBar1.Value = 0
oXL.Quit
MsgBox "Import Succesful!!!"
CmdSave.SetFocus
End Sub

Private Sub CmdReset_Click()
    Call Grid_Head
    TxtNAVCode.Text = ""
    CmbLessor.Text = ""
    CmdExit.Enabled = True
    CmdSave.Enabled = False
    CmdGetFile.Enabled = True
    CmdImport.Enabled = False
    LblFile.Caption = ""
    LblTotalRecords.Caption = ""
    ProgressBar1.Value = 0
End Sub

Private Sub CmdSave_Click()
Dim mError As Boolean
Dim mErrorMsg, mStateIndiaNo, mSplitError As Variant
Dim mR, mC As Variant
Dim mExisitEntryFlag As Boolean
Dim mLessorPANNo, mGSTPANNo As Variant
With MSHFlexGrid1
    If .Rows > 2 Then
        ProgressBar1.Max = .Rows - 2
        Call TableMst_LessorStateWiseGST
        For mR = 1 To .Rows - 2
            mErrorMsg = ""
            mLessorPANNo = ""
            mGSTPANNo = ""
            mError = False
            ''--- Lessor Name for Checking
            If Trim(.TextMatrix(mR, 0)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Lessor Name not allowed." & vbNewLine
            Else
                Call TableMst_Lessor(" Where LessorName = '" & Trim(.TextMatrix(mR, 0)) & "'")
                If RsMst_Lessor.EOF = True Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Lessor Name not found." & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 6)) <> RsMst_Lessor!LessorID Then
                       mError = True
                       If mError = True Then mErrorMsg = mErrorMsg & "Lessor Name and Lessor ID is mismatch." & vbNewLine
                    Else
                        mLessorPANNo = IIf(IsNull(RsMst_Lessor!PANNO), "", RsMst_Lessor!PANNO)
                    End If
                End If
            End If
            ''--- NAV Code for Checking
            Call TableMst_Lessor(" Where LessorID = '" & Val(.TextMatrix(mR, 6)) & "'")
            If RsMst_Lessor.EOF = False Then
                If LCase(IIf(IsNull(RsMst_Lessor!NavCode), "", RsMst_Lessor!NavCode)) <> LCase(.TextMatrix(mR, 1)) Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "NAV Code is mismatch with given Lessor." & vbNewLine
                End If
            End If
            ''--- State India Name for Checking
            If Trim(.TextMatrix(mR, 2)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank State India Name not allowed." & vbNewLine
            Else
                Call TableMst_StateIndia(" Where Flag = 'A' And StateIndiaName = '" & Trim(.TextMatrix(mR, 2)) & "'  ")
                If RsMst_StateIndia.EOF = True Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "State India Name not found or given State India Name is deactive" & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 7)) <> RsMst_StateIndia!StateIndiaID Then
                       mError = True
                       If mError = True Then mErrorMsg = mErrorMsg & "State India Name and State India ID is mismatch." & vbNewLine
                    End If
                End If
            End If
            ''--- State Wise Bank Address Checking
            If Trim(.TextMatrix(mR, 3)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank State Wise Lessor Address not allowed." & vbNewLine
            End If
            If Len(.TextMatrix(mR, 3)) > 500 Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "State Wise Lessor Address more then 500 characters not allowed." & vbNewLine
            End If
            ''--- GST NAV Code for Checking
''            If Trim(.TextMatrix(mR, 5)) = "" Then
''                mError = True
''                If mError = True Then mErrorMsg = mErrorMsg & "Blank GST NAV Code not allowed." & vbNewLine
''            End If
            If Len(.TextMatrix(mR, 5)) > 25 Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "GST NAV Code more then 25 characters not allowed." & vbNewLine
            End If
            ''--- Lessor ID for Checking
            If Trim(.TextMatrix(mR, 6)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Lerror ID allowed." & vbNewLine
            Else
                If Val(.TextMatrix(mR, 6)) = 0 Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Given Lerror ID allowed." & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 6)) > 0 Then
                        Call TableMst_Lessor(" Where LessorID = " & Val(.TextMatrix(mR, 6)) & " ")
                        If RsMst_Lessor.EOF = True Then
                            mError = True
                            If mError = True Then mErrorMsg = mErrorMsg & "Lessor ID not found." & vbNewLine
                        End If
                    End If
                End If
            End If
            ''--- StateIndia ID for Checking
            If Trim(.TextMatrix(mR, 7)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank StateIndia ID allowed." & vbNewLine
            Else
                If Val(.TextMatrix(mR, 7)) = 0 Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Given StateIndia ID allowed." & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 7)) > 0 Then
                        Call TableMst_StateIndia(" Where Flag = 'A' And StateIndiaID = " & Val(.TextMatrix(mR, 7)) & " ")
                        If RsMst_StateIndia.EOF = True Then
                            mError = True
                            If mError = True Then mErrorMsg = mErrorMsg & "StateIndia ID not found or given StateIndia ID is deactive" & vbNewLine
                        Else
                            mStateIndiaNo = IIf(IsNull(RsMst_StateIndia!TINNo), 0, RsMst_StateIndia!TINNo)
                        End If
                    End If
                End If
            End If
            ''--- GST No for Checking
            If Trim(.TextMatrix(mR, 4)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank GST No not allowed." & vbNewLine
            Else
                ''-- Lessor PAN No validate with GST PAN NO
                mGSTPANNo = mID(Trim(.TextMatrix(mR, 4)), 3, 10)
'                If LCase(mGSTPANNo) <> LCase(mLessorPANNo) Then
'                   mError = True
'                   If mError = True Then mErrorMsg = mErrorMsg & "Provided GST No PAN No and Lessor PAN No is mismatch." & vbNewLine
'                End If
                ''-- GST No Validation
                ''-- Validation is with specific Error
                mSplitError = Split(ValidateGSTIN(mStateIndiaNo, Trim(.TextMatrix(mR, 4))), "~")
                If CBool(mSplitError(0)) = False Then
                   mError = True
                   If mError = True Then mErrorMsg = mErrorMsg & CStr(mSplitError(1)) & vbNewLine
                Else
''''                    ''-- Check for Duplicate GST No
''''                    If RsMst_LessorStateWiseGST.EOF = False Then
''''                        RsMst_LessorStateWiseGST.MoveFirst
''''                        RsMst_LessorStateWiseGST.Find "GSTNo = '" & Trim(.TextMatrix(mR, 4)) & "'"
''''                    End If
''''                    If RsMst_LessorStateWiseGST.EOF = False Then
''''                        If Val(.TextMatrix(mR, 8)) > 0 Then
''''                            If Val(.TextMatrix(mR, 8)) <> Val(RsMst_LessorStateWiseGST!LessorStateWiseGSTID) Then
''''                                mError = True
''''                                If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number not allowed." & vbNewLine
''''                            End If
''''                        Else
''''                            mError = True
''''                            If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number not allowed." & vbNewLine
''''                        End If
''''                    End If
                End If
            End If
            ''--- GST NAV Code for Checking
''''            If Trim(.TextMatrix(mR, 5)) <> "" Then
''''                ''-- Check for Duplicate GST No
''''                If RsMst_LessorStateWiseGST.EOF = False Then
''''                    RsMst_LessorStateWiseGST.MoveFirst
''''                    RsMst_LessorStateWiseGST.Find "GSTNavCode = '" & Trim(.TextMatrix(mR, 5)) & "'"
''''                End If
''''                If RsMst_LessorStateWiseGST.EOF = False Then
''''                    If Val(.TextMatrix(mR, 8)) > 0 Then
''''                        If Val(.TextMatrix(mR, 8)) <> Val(RsMst_LessorStateWiseGST!LessorStateWiseGSTID) Then
''''                            mError = True
''''                            If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST NAV Code not allowed." & vbNewLine
''''                        End If
''''                    Else
''''                        mError = True
''''                        If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST NAV Code not allowed." & vbNewLine
''''                    End If
''''                End If
''''            End If
            ''--- ID for Checking
            If Trim(.TextMatrix(mR, 8)) <> "" And Val(.TextMatrix(mR, 8)) > 0 Then
                If RsMst_LessorStateWiseGST.RecordCount > 0 Then
                    RsMst_LessorStateWiseGST.MoveFirst
                    RsMst_LessorStateWiseGST.Find "LessorStateWiseGSTID = " & Val(.TextMatrix(mR, 8)) & ""
                    If RsMst_LessorStateWiseGST.EOF = True Then
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Given ID is not available." & vbNewLine
                    Else
                        If Val(.TextMatrix(mR, 6)) > 0 And Val(.TextMatrix(mR, 7)) > 0 Then
                            If Val(.TextMatrix(mR, 6)) <> Val(RsMst_LessorStateWiseGST!LessorID) Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "LessorID is mismatch with Given ID." & vbNewLine
                            End If
                            If Val(.TextMatrix(mR, 7)) <> Val(RsMst_LessorStateWiseGST!StateIndiaID) Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "StateIndiaID is mismatch with Given ID." & vbNewLine
                            End If
                        End If
                    End If
                End If
            End If
            ''-- Duplicate Entry by Lessor ID and StateIndia ID
            If Val(.TextMatrix(mR, 6)) > 0 And Val(.TextMatrix(mR, 7)) > 0 Then
                If RsMst_LessorStateWiseGST.RecordCount > 0 Then
                   RsMst_LessorStateWiseGST.MoveFirst
                End If
                RsMst_LessorStateWiseGST.Filter = " LessorID = " & Val(.TextMatrix(mR, 6)) & " And StateIndiaID = " & Val(.TextMatrix(mR, 7)) & " "
                If RsMst_LessorStateWiseGST.EOF = False Then
                   If .TextMatrix(mR, 8) <> "" And Val(.TextMatrix(mR, 8)) > 0 Then
                        If Val(.TextMatrix(mR, 8)) <> Val(RsMst_LessorStateWiseGST!LessorStateWiseGSTID) Then
                            mError = True
                            If mError = True Then mErrorMsg = mErrorMsg & "Same Lessor and Same State India not allowed." & vbNewLine
                        End If
                   Else
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Same Lessor and Same State India not allowed." & vbNewLine
                   End If
                End If
                RsMst_LessorStateWiseGST.Filter = ""
            End If
            
            mNewEntry = False
            ''-- Create 1 Col Extra for Error
            If mError = True Then
                .row = mR
                .Col = .Cols - 2
                .CellBackColor = vbRed
                .TextMatrix(mR, .Cols - 2) = mErrorMsg
            Else
                If Val(.TextMatrix(mR, 8)) >= 0 Then
                    If RsMst_LessorStateWiseGST.RecordCount > 0 Then
                       RsMst_LessorStateWiseGST.MoveFirst
                    End If
                    RsMst_LessorStateWiseGST.Find "LessorStateWiseGSTID = " & Val(.TextMatrix(mR, 8)) & ""
                    If RsMst_LessorStateWiseGST.EOF = True Then
                       RsMst_LessorStateWiseGST.AddNew
                       RsMst_LessorStateWiseGST!LessorStateWiseGSTID = GetMaxID("LessorStateWiseGSTID", "Mst_LessorStateWiseGST")
                       .TextMatrix(mR, 8) = RsMst_LessorStateWiseGST!LessorStateWiseGSTID
                       RsMst_LessorStateWiseGST!G_UID = GetGUID
                       RsMst_LessorStateWiseGST!LessorID = Val(.TextMatrix(mR, 6))
                       RsMst_LessorStateWiseGST!StateIndiaID = Val(.TextMatrix(mR, 7))
                       mNewEntry = True
                    Else
                       ''''-- Update Log
                       ''mExisitEntryFlag = True
                       ''.row = mR
                       ''.Col = .Cols - 1
                       ''.CellBackColor = vbYellow
                       ''.TextMatrix(mR, .Cols - 1) = "Mapping is update with old data with new data old values are GSTNo = [" & RsMst_LessorStateWiseGST!GSTNo & "],GSTNavCode = [" & RsMst_LessorStateWiseGST!GSTNavCode & "],LessorStateOfficeAddress = [" & RsMst_LessorStateWiseGST!LessorStateOfficeAddress & "] and new values are GSTNo = [" & Trim(.TextMatrix(mR, 4)) & "],GSTNavCode = [" & Trim(.TextMatrix(mR, 5)) & "],LessorStateOfficeAddress = [" & Trim(.TextMatrix(mR, 3)) & "]"
                       
                        ''-- Ullhas Patil 08 Nov 2017
                        ''-- Appply to block the update the GSTNO/ Address/ NAV Code as per instruct by Mayur Sir only power user can allow to upload with update existing entry
                        If mAllowUpdateFlag = True Then
                           ''-- Update Log
                           mExisitEntryFlag = True
                           .row = mR
                           .Col = .Cols - 1
                           .CellBackColor = vbYellow
                           .TextMatrix(mR, .Cols - 1) = "Mapping is update with old data with new data old values are GSTNo = [" & RsMst_LessorStateWiseGST!GSTNo & "],GSTNavCode = [" & RsMst_LessorStateWiseGST!GSTNavCode & "],LessorStateOfficeAddress = [" & RsMst_LessorStateWiseGST!LessorStateOfficeAddress & "] and new values are GSTNo = [" & Trim(.TextMatrix(mR, 4)) & "],GSTNavCode = [" & Trim(.TextMatrix(mR, 5)) & "],LessorStateOfficeAddress = [" & Trim(.TextMatrix(mR, 3)) & "]"
                        Else
                           ''-- Update Log
                           mExisitEntryFlag = True
                           .row = mR
                           .Col = .Cols - 1
                           .CellBackColor = vbRed
                           .TextMatrix(mR, .Cols - 1) = "Not allow to Update Existing GST Data."
                        End If
                        ''------------------------------
                    End If
                    
                    ''RsMst_LessorStateWiseGST!LessorStateOfficeAddress = Trim(.TextMatrix(mR, 3))
                    ''RsMst_LessorStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 4))
                    ''RsMst_LessorStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 5))
                    '''RsMst_LessorStateWiseGST!Flag = ""
                    '''RsMst_LessorStateWiseGST!Remarks = ""
                    ''If IsNull(RsMst_LessorStateWiseGST!CreatedBy) = True Or RsMst_LessorStateWiseGST!CreatedBy = "" Then
                    ''   RsMst_LessorStateWiseGST!CreatedBy = Gen_UserLoginID
                    ''   RsMst_LessorStateWiseGST!CreatedDate = Now
                    ''Else
                    ''   RsMst_LessorStateWiseGST!UpdatedBy = Gen_UserLoginID
                    ''   RsMst_LessorStateWiseGST!UpdatedDate = Now
                    ''End If
                    
                    ''-- Ullhas Patil 08 Nov 2017
                    If mAllowUpdateFlag = True Then
                        RsMst_LessorStateWiseGST!LessorStateOfficeAddress = Trim(.TextMatrix(mR, 3))
                        RsMst_LessorStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 4))
                        RsMst_LessorStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 5))
                        If IsNull(RsMst_LessorStateWiseGST!CreatedBy) = True Or RsMst_LessorStateWiseGST!CreatedBy = "" Then
                           RsMst_LessorStateWiseGST!CreatedBy = Gen_UserLoginID
                           RsMst_LessorStateWiseGST!CreatedDate = Now
                        Else
                           RsMst_LessorStateWiseGST!UpdatedBy = Gen_UserLoginID
                           RsMst_LessorStateWiseGST!UpdatedDate = Now
                        End If
                    Else
                        If mNewEntry = True Then
                            RsMst_LessorStateWiseGST!LessorStateOfficeAddress = Trim(.TextMatrix(mR, 3))
                            RsMst_LessorStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 4))
                            RsMst_LessorStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 5))
                            RsMst_LessorStateWiseGST!CreatedBy = Gen_UserLoginID
                            RsMst_LessorStateWiseGST!CreatedDate = Now
                        End If
                    End If
                    ''------------------------------------
                    RsMst_LessorStateWiseGST.Update
                    RsMst_LessorStateWiseGST.Requery
                    
                 End If
            End If
            ProgressBar1.Value = mR
            mErrorMsg = ""
            .Refresh
        Next mR
    End If
    CmdSave.Enabled = False
    If mExisitEntryFlag = False Then
        MsgBox "Done.", vbOKOnly + vbInformation
    Else
        ''MsgBox "Some existing Lessor-State entries are updated, so please export Excel for more details.", vbOKOnly + vbInformation
        If mAllowUpdateFlag = True Then
            MsgBox "Some existing Lessor-State entries are updated, so please export Excel for more details.", vbOKOnly + vbInformation
        Else
            MsgBox "Existing Lessor-State entries are not allow to update.", vbOKOnly + vbCritical
        End If
    End If
End With
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Call Grid_Head
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

mAllowUpdateFlag = False
mAllowUpdateFlag = Gen_GSTDataUpdateUserRole()
''If LCase(Gen_UserRole) = "power" Then mAllowUpdateFlag = True

Call Grid_Head
LblFile.Caption = ""
ProgressBar1.Value = 0
Frame1.Enabled = True
CmdGet.Enabled = True
CmdGetFile.Enabled = True
CmdImport.Enabled = False
CmdSave.Enabled = False
CmdReset.Enabled = True
CmdExcel.Enabled = True
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 9
     .FixedCols = 0
     .Font.Size = 9
     .AllowUserResizing = flexResizeColumns
     .WordWrap = True
     .TextMatrix(0, 0) = "Lessor Name":                     .ColWidth(0) = 1500
     .TextMatrix(0, 1) = "NAV Code":                        .ColWidth(1) = 1000
     .TextMatrix(0, 2) = "State India Name":                .ColWidth(2) = 1500
     .TextMatrix(0, 3) = "State Wise Lessor Address":       .ColWidth(3) = 2000
     .TextMatrix(0, 4) = "GST No":                          .ColWidth(4) = 1500
     .TextMatrix(0, 5) = "GST NAV Code":                    .ColWidth(5) = 1500
     .TextMatrix(0, 6) = "Lessor ID":                       .ColWidth(6) = 0
     .TextMatrix(0, 7) = "StateIndiaID":                    .ColWidth(7) = 0
     .TextMatrix(0, 8) = "ID":                              .ColWidth(8) = 0
     
     .RowHeight(0) = 500
End With
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_LessorStateWiseGST.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

With MSHFlexGrid1
For i = 0 To .Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To .Cols - 1
        If InStr(.TextMatrix(i, C), "/") > 0 Then
           If IsDate(.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(.TextMatrix(i, C), "dd-mmm-yyyy")
           Else
              oWS.Cells(mRow, C + 1) = .TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = .TextMatrix(i, C)
        End If
    Next
Next
End With

oWB.Save
'oWB.Close
'oXL.Quit
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub


Private Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True
ErrMsgFile = ""
With MSHFlexGrid1
    For mC = 0 To .Cols - 1
        If LCase(Trim(oWS.Cells(1, mC + 1))) <> LCase(.TextMatrix(0, mC)) Then
           Retval = False
           ErrMsgFile = " Please Check Column Heading for Column " & mC + 1 & ""
        End If
    Next mC
End With
CheckFile = Retval
End Function

Private Sub FillData()
'On Error GoTo errortrap
Dim l, C As Variant
Dim RowCount As Variant
oWS.ClearArrows
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
LblTotalRecords.Caption = RowCount - 1
ProgressBar1.Value = 0
ProgressBar1.Max = RowCount
With MSHFlexGrid1
For l = 1 To RowCount
    If Trim(oWS.Cells(l + 1, 1)) = "" Then
       Exit For
    End If
    For C = 0 To .Cols - 1
        .TextMatrix(l, C) = Trim(oWS.Cells(l + 1, C + 1))
    Next
    .Rows = MSHFlexGrid1.Rows + 1
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
End With
'errortrap:   MsgBox Err.Description, vbExclamation, "Error"
oXL.Quit
End Sub



