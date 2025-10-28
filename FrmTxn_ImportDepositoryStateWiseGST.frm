VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTxn_ImportDepositoryStateWiseGST 
   Caption         =   "Import Depository State Wise GST"
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
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9615
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   15015
      Begin VB.TextBox TxtNAVCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   12840
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   12660
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   10470
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
         Enabled         =   0   'False
         Height          =   345
         Left            =   9375
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   345
         Left            =   13755
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
         Enabled         =   0   'False
         Height          =   345
         Left            =   11565
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
         Height          =   345
         Left            =   8280
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox CmbDepository 
         Height          =   360
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   9315
      End
      Begin VB.CommandButton CmdGet 
         Caption         =   "Get Data"
         Height          =   345
         Left            =   7200
         TabIndex        =   2
         Top             =   720
         Width           =   1095
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
         Height          =   7860
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13864
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NAV Code"
         Height          =   240
         Left            =   11520
         TabIndex        =   16
         Top             =   300
         Width           =   960
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1150
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Total Record :"
         Height          =   255
         Left            =   12600
         TabIndex        =   14
         Top             =   1120
         Width           =   1335
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
         TabIndex        =   13
         Top             =   1127
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Depository Name"
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
         TabIndex        =   12
         Top             =   300
         Width           =   1830
      End
   End
End
Attribute VB_Name = "FrmTxn_ImportDepositoryStateWiseGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet

Dim ErrMsg As Variant
Dim Status As Variant
Dim mDepositoryID As Variant
Dim ErrMsgFile As String

''-- Ullhas Patil 08 Nov 2017
Dim mAllowUpdateFlag As Boolean
Dim mNewEntry As Boolean

Private Sub CmbDepository_GotFocus()
tmp = CmbDepository.Text
mTmp = CmbDepository.Text
Call TableMst_Depository
CmbDepository.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No Depository found"
   CmbDepository.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
    CmbDepository.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    RsMst_Depository.MoveNext
Next
CmbDepository.Text = tmp
End Sub

Private Sub CmbDepository_LostFocus()
Dim mSplit As Variant
If CmbDepository.Text = "" Then
   Exit Sub
End If
mSplit = Split(CmbDepository.Text, "~")
If UBound(mSplit) <> 1 Then
   MsgBox "Invalid selection "
   CmbDepository.SetFocus
   Exit Sub
End If
RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID = '" & mSplit(1) & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbDepository.SetFocus
   Exit Sub
Else
   If LCase(RsMst_Depository!DepositoryName) <> LCase(mSplit(0)) Then
    MsgBox "Invalid selection "
    CmbDepository.SetFocus
    Exit Sub
   End If
End If
mDepositoryID = IIf(IsNull(RsMst_Depository!DepositoryID), "", RsMst_Depository!DepositoryID)
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_DepositoryStateWiseGST.xls")
Call Xls_Detail_Rpt
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGet_Click()
Dim wc As Variant
If Trim(TxtNAVCode.Text) = "" Then
    If CmbDepository.Text = "" Then
        MsgBox "Select Depository"
        Exit Sub
    End If
End If
Call Grid_Head

wc = ""
If Trim(TxtNAVCode.Text) = "" Then
    wc = "'" & mDepositoryID & "'"
Else
    wc = " SELECT DepositoryID  FROM Mst_Depository Where NAVCode = '" & Trim(TxtNAVCode) & "'"
End If

tmp = " Select A.DepositoryStateWiseGSTID,A.DepositoryID,A.DepositoryName,A.NavCode,A.ExchangeTypeID,A.ExchangeType,A.StateIndiaID,A.StateIndiaName,A.DepositoryStateOfficeAddress,A.GSTNo,A.GSTNavCode From"
tmp = tmp & " (Select MDSWG.DepositoryStateWiseGSTID,MDSWG.DepositoryID,MD.DepositoryName,MD.NavCode,MDSWG.ExchangeTypeID,MET.ExchangeType,"
tmp = tmp & " MDSWG.StateIndiaID , MSI.StateIndiaName, MDSWG.DepositoryStateOfficeAddress, MDSWG.GSTNo, MDSWG.GSTNavCode"
tmp = tmp & " From Mst_DepositoryStateWiseGST MDSWG"
tmp = tmp & " Inner Join Mst_Depository MD On MDSWG.DepositoryID = MD.DepositoryID"
tmp = tmp & " Inner Join Mst_ExchangeType MET On MDSWG.ExchangeTypeID = MET.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_StateIndia MSI On MDSWG.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Where MDSWG.DepositoryID IN (" & wc & ")"
tmp = tmp & " Union"
tmp = tmp & " SELECT '0'DepositoryStateWiseGSTID,MD.DepositoryID,MD.DepositoryName,MD.NavCode,MD.ExchangeTypeID,MET.ExchangeType,"
tmp = tmp & " MSI.StateIndiaID,MSI.StateIndiaName,''DepositoryStateOfficeAddress,''GSTNo,''GSTNavCode"
tmp = tmp & " FROM Mst_Depository MD Inner Join Mst_ExchangeType MET On MD.ExchangeTypeID = MET.ExchangeTypeID ,Mst_StateIndia MSI"
tmp = tmp & " Where MD.DepositoryID IN (" & wc & ") And Convert(Varchar(50),MD.DepositoryID)+'~'+Convert(Varchar(18),MSI.StateIndiaID) Not in (SELECT Convert(Varchar(50),DepositoryID) +'~'+Convert(Varchar(18),StateIndiaID) FROM Mst_DepositoryStateWiseGST  Where DepositoryID IN (SELECT DepositoryID FROM Mst_Depository Where DepositoryID IN (" & wc & ")))"
tmp = tmp & " ) AS A"
tmp = tmp & " Order By A.DepositoryName,A.DepositoryID,A.StateIndiaName,A.DepositoryStateWiseGSTID"
Call TmpTable
If TmpRs.EOF = False Then
    
    With MSHFlexGrid1
    If TmpRs.RecordCount > 0 Then
       ProgressBar1.Max = TmpRs.RecordCount
       For mR = 1 To TmpRs.RecordCount
              
           .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!DepositoryID), "", TmpRs!DepositoryID)
           .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
           .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
           .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
           .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
           .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!DepositoryStateOfficeAddress), "", TmpRs!DepositoryStateOfficeAddress)
           .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
           .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
           .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
           .TextMatrix(mR, 9) = IIf(TmpRs!DepositoryStateWiseGSTID = 0, "", TmpRs!DepositoryStateWiseGSTID)

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
   Dim I As Integer
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
    TxtNAVCode = ""
    CmbDepository.Text = ""
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
Dim mErrorMsg, mStateIndiaNo, mSplitError, mNAVCode As Variant
Dim mR, mC As Variant
Dim mExisitEntryFlag As Boolean

With MSHFlexGrid1
    If .Rows > 2 Then
        ProgressBar1.Max = .Rows - 2
        Call TableMst_DepositoryStateWiseGST
        For mR = 1 To .Rows - 2
            mErrorMsg = ""
            mError = False
            ''--- Depository ID for Checking
            If Trim(.TextMatrix(mR, 0)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Depository ID not allowed." & vbNewLine
            Else
                Call TableMst_Depository(" Where DepositoryID = '" & .TextMatrix(mR, 0) & "'")
                If RsMst_Depository.EOF = True Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Depository ID not found." & vbNewLine
                Else
                    If LCase(.TextMatrix(mR, 1)) <> LCase(RsMst_Depository!DepositoryName) Then
                       mError = True
                       If mError = True Then mErrorMsg = mErrorMsg & "Depository ID and Depository Name is mismatch." & vbNewLine
                    End If
                    mNAVCode = IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)
                End If
            End If
            
            ''--- Depository Name for Checking
            If Trim(.TextMatrix(mR, 1)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Depository Name not allowed." & vbNewLine
            Else
                Call TableMst_Depository(" Where DepositoryID = '" & .TextMatrix(mR, 0) & "'")
                If RsMst_Depository.EOF = True Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Depository ID not found." & vbNewLine
                Else
                    If LCase(.TextMatrix(mR, 1)) <> LCase(RsMst_Depository!DepositoryName) Then
                       mError = True
                       If mError = True Then mErrorMsg = mErrorMsg & "Depository ID and Depository Name is mismatch." & vbNewLine
                    End If
                End If
            End If
            ''--- Exchange Type for Checking
            If Trim(.TextMatrix(mR, 2)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Exchange Type not allowed." & vbNewLine
            Else
                Call TableMst_Depository(" Where DepositoryID = '" & .TextMatrix(mR, 0) & "'")
                If RsMst_Depository.EOF = False Then
                    If LCase(GetExchangeName(RsMst_Depository!ExchangeTypeID)) <> LCase(.TextMatrix(mR, 2)) Then
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Exchange Type is mismatch with given Depository." & vbNewLine
                    End If
                End If
            End If
            ''--- NAV Code for Checking
            Call TableMst_Depository(" Where DepositoryID = '" & Val(.TextMatrix(mR, 0)) & "'")
            If RsMst_Depository.EOF = False Then
                If LCase(IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)) <> LCase(.TextMatrix(mR, 3)) Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "NAV Code is mismatch with given Depository." & vbNewLine
                End If
            End If
            ''--- State India Name for Checking
            If Trim(.TextMatrix(mR, 4)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank State India Name not allowed." & vbNewLine
            Else
                Call TableMst_StateIndia(" Where Flag = 'A' And StateIndiaName = '" & Trim(.TextMatrix(mR, 4)) & "'  ")
                If RsMst_StateIndia.EOF = True Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "State India Name not found or given State India Name is deactive" & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 8)) <> RsMst_StateIndia!StateIndiaID Then
                       mError = True
                       If mError = True Then mErrorMsg = mErrorMsg & "State India Name and State India ID is mismatch." & vbNewLine
                    End If
                End If
            End If
            ''--- State Wise Bank Address Checking
            If Trim(.TextMatrix(mR, 5)) = "" Then
              mError = True
              If mError = True Then mErrorMsg = mErrorMsg & "Blank State Wise Depository Address not allowed." & vbNewLine
            End If
            If Len(.TextMatrix(mR, 5)) > 500 Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "State Wise Depository Address more then 500 characters not allowed." & vbNewLine
            End If
            ''--- GST NAV Code for Checking
''            If Trim(.TextMatrix(mR, 7)) = "" Then
''                mError = True
''                If mError = True Then mErrorMsg = mErrorMsg & "Blank GST NAV Code not allowed." & vbNewLine
''            End If
            If Len(.TextMatrix(mR, 7)) > 25 Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "GST NAV Code more then 25 characters not allowed." & vbNewLine
            End If
            ''--- StateIndia ID for Checking
            If Trim(.TextMatrix(mR, 8)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank StateIndia ID allowed." & vbNewLine
            Else
                If Val(.TextMatrix(mR, 8)) = 0 Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Given StateIndia ID allowed." & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 8)) > 0 Then
                        Call TableMst_StateIndia(" Where Flag = 'A' And StateIndiaID = " & Val(.TextMatrix(mR, 8)) & " ")
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
            If Trim(.TextMatrix(mR, 6)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank GST No not allowed." & vbNewLine
            Else
                ''-- GST No Validation
                ''-- Validation is with specific Error
                mSplitError = Split(ValidateGSTIN(mStateIndiaNo, Trim(.TextMatrix(mR, 6))), "~")
                If CBool(mSplitError(0)) = False Then
                   mError = True
                   If mError = True Then mErrorMsg = mErrorMsg & CStr(mSplitError(1)) & vbNewLine
                Else
''''                    ''-- Check for Duplicate GST No
''''                    If RsMst_DepositoryStateWiseGST.RecordCount > 0 Then
''''                        RsMst_DepositoryStateWiseGST.MoveFirst
''''                        RsMst_DepositoryStateWiseGST.Find "GSTNo = '" & Trim(.TextMatrix(mR, 6)) & "'"
''''                    End If
''''                    If RsMst_DepositoryStateWiseGST.EOF = False Then
''''                        If Val(.TextMatrix(mR, 9)) > 0 Then
''''                            tmp = " SELECT Distinct MD.NavCode[NavCode] FROM Mst_DepositoryStateWiseGST MDSWG Inner Join Mst_Depository MD On MDSWG.DepositoryID = MD.DepositoryID Where MDSWG.DepositoryID = '" & RsMst_DepositoryStateWiseGST!DepositoryID & "' "
''''                            Call Tmp2Table
''''                            If TmpRs2.RecordCount > 0 Then
''''                               If IsNull(TmpRs2!NavCode) = False Then
''''                                If IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode) <> mNAVCode Then
''''                                    mError = True
''''                                    If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number with differnet NAV Code not allowed." & vbNewLine
''''                                End If
''''                               End If
''''                            End If
''''''                            If Val(.TextMatrix(mR, 9)) <> Val(RsMst_DepositoryStateWiseGST!DepositoryStateWiseGSTID) Then
''''''                                mError = True
''''''                                If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number not allowed." & vbNewLine
''''''                            End If
''''                        Else
''''                            tmp = " SELECT Distinct MD.NavCode[NavCode] FROM Mst_DepositoryStateWiseGST MDSWG Inner Join Mst_Depository MD On MDSWG.DepositoryID = MD.DepositoryID Where MDSWG.DepositoryID = '" & RsMst_DepositoryStateWiseGST!DepositoryID & "' "
''''                            Call Tmp2Table
''''                            If TmpRs2.RecordCount > 0 Then
''''                               If IsNull(TmpRs2!NavCode) = False Then
''''                                If IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode) <> mNAVCode Then
''''                                    mError = True
''''                                    If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number with differnet NAV Code not allowed." & vbNewLine
''''                                End If
''''                               End If
''''                            End If
''''''                            mError = True
''''''                            If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number not allowed." & vbNewLine
''''                        End If
''''                    End If
                End If
            End If
            
            ''--- GST NAV Code Check for Duplicate GST No
''''            If Trim(.TextMatrix(mR, 7)) <> "" Then
''''                If RsMst_DepositoryStateWiseGST.RecordCount > 0 Then
''''                    RsMst_DepositoryStateWiseGST.MoveFirst
''''                    RsMst_DepositoryStateWiseGST.Find "GSTNavCode = '" & Trim(.TextMatrix(mR, 7)) & "'"
''''                End If
''''                If RsMst_DepositoryStateWiseGST.EOF = False Then
''''                    If Val(.TextMatrix(mR, 9)) > 0 Then
''''                        tmp = " SELECT Distinct MD.NavCode[NavCode] FROM Mst_DepositoryStateWiseGST MDSWG Inner Join Mst_Depository MD On MDSWG.DepositoryID = MD.DepositoryID Where MDSWG.DepositoryID = '" & RsMst_DepositoryStateWiseGST!DepositoryID & "' "
''''                        Call Tmp2Table
''''                        If TmpRs2.RecordCount > 0 Then
''''                           If IsNull(TmpRs2!NavCode) = False Then
''''                            If IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode) <> mNAVCode Then
''''                                mError = True
''''                                If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST NAV Code with differnet NAV Code not allowed." & vbNewLine
''''                            End If
''''                           End If
''''                        End If
''''                    Else
''''                        tmp = " SELECT Distinct MD.NavCode[NavCode] FROM Mst_DepositoryStateWiseGST MDSWG Inner Join Mst_Depository MD On MDSWG.DepositoryID = MD.DepositoryID Where MDSWG.DepositoryID = '" & RsMst_DepositoryStateWiseGST!DepositoryID & "' "
''''                        Call Tmp2Table
''''                        If TmpRs2.RecordCount > 0 Then
''''                           If IsNull(TmpRs2!NavCode) = False Then
''''                            If IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode) <> mNAVCode Then
''''                                mError = True
''''                                If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST NAV Code with differnet NAV Code not allowed." & vbNewLine
''''                            End If
''''                           End If
''''                        End If
''''                    End If
''''                End If
''''            End If
            
            ''--- ID for Checking
            If Trim(.TextMatrix(mR, 9)) <> "" And Val(.TextMatrix(mR, 9)) > 0 Then
                If RsMst_DepositoryStateWiseGST.RecordCount > 0 Then
                    RsMst_DepositoryStateWiseGST.MoveFirst
                    RsMst_DepositoryStateWiseGST.Find "DepositoryStateWiseGSTID = " & Val(.TextMatrix(mR, 9)) & ""
                    If RsMst_DepositoryStateWiseGST.EOF = True Then
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Given ID is not available." & vbNewLine
                    Else
                        If .TextMatrix(mR, 0) <> "" And .TextMatrix(mR, 2) <> "" And Val(.TextMatrix(mR, 8)) > 0 Then
                            If .TextMatrix(mR, 0) <> RsMst_DepositoryStateWiseGST!DepositoryID And GetExchangeID(.TextMatrix(mR, 2)) <> RsMst_DepositoryStateWiseGST!ExchangeTypeID Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "Depository ID is mismatch with Given ID." & vbNewLine
                            End If
                            If Val(.TextMatrix(mR, 8)) <> Val(RsMst_DepositoryStateWiseGST!StateIndiaID) Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "StateIndiaID is mismatch with Given ID." & vbNewLine
                            End If
                        End If
                    End If
                End If
            End If
            ''-- Duplicate Entry by Depository ID and StateIndia ID
            If .TextMatrix(mR, 0) <> "" And .TextMatrix(mR, 2) <> "" And Val(.TextMatrix(mR, 8)) > 0 Then
                If RsMst_DepositoryStateWiseGST.RecordCount > 0 Then
                   RsMst_DepositoryStateWiseGST.MoveFirst
                End If
                RsMst_DepositoryStateWiseGST.Filter = " DepositoryID = '" & .TextMatrix(mR, 0) & "' And ExchangeTypeID = " & GetExchangeID(.TextMatrix(mR, 2)) & " And StateIndiaID = " & Val(.TextMatrix(mR, 8)) & " "
                If RsMst_DepositoryStateWiseGST.EOF = False Then
                   If .TextMatrix(mR, 9) <> "" And Val(.TextMatrix(mR, 9)) > 0 Then
                        If Val(.TextMatrix(mR, 9)) <> Val(RsMst_DepositoryStateWiseGST!DepositoryStateWiseGSTID) Then
                            mError = True
                            If mError = True Then mErrorMsg = mErrorMsg & "Same Depository and Same State India not allowed." & vbNewLine
                        End If
                   Else
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Same Depository and Same State India not allowed." & vbNewLine
                   End If
                End If
                RsMst_DepositoryStateWiseGST.Filter = ""
            End If
            
            mNewEntry = False
            ''-- Create 1 Col Extra for Error
            If mError = True Then
                .row = mR
                .Col = .Cols - 2
                .CellBackColor = vbRed
                .TextMatrix(mR, .Cols - 2) = mErrorMsg
            Else
                If Val(.TextMatrix(mR, 7)) >= 0 Then
                    If RsMst_DepositoryStateWiseGST.RecordCount > 0 Then
                       RsMst_DepositoryStateWiseGST.MoveFirst
                    End If
                    RsMst_DepositoryStateWiseGST.Find "DepositoryStateWiseGSTID = " & Val(.TextMatrix(mR, 9)) & ""
                    If RsMst_DepositoryStateWiseGST.EOF = True Then
                       RsMst_DepositoryStateWiseGST.AddNew
                       RsMst_DepositoryStateWiseGST!DepositoryStateWiseGSTID = GetMaxID("DepositoryStateWiseGSTID", "Mst_DepositoryStateWiseGST")
                       .TextMatrix(mR, 9) = RsMst_DepositoryStateWiseGST!DepositoryStateWiseGSTID
                       RsMst_DepositoryStateWiseGST!G_UID = GetGUID
                       RsMst_DepositoryStateWiseGST!DepositoryID = .TextMatrix(mR, 0)
                       RsMst_DepositoryStateWiseGST!ExchangeTypeID = GetExchangeID(.TextMatrix(mR, 2))
                       RsMst_DepositoryStateWiseGST!StateIndiaID = Val(.TextMatrix(mR, 8))
                       mNewEntry = True
                    Else
                       ''''-- Update Log
                       ''mExisitEntryFlag = True
                       ''.row = mR
                       ''.Col = .Cols - 1
                       ''.CellBackColor = vbYellow
                       ''.TextMatrix(mR, .Cols - 1) = "Mapping is update with old data with new data old values are GSTNo = [" & RsMst_DepositoryStateWiseGST!GSTNo & "],GSTNavCode = [" & RsMst_DepositoryStateWiseGST!GSTNavCode & "],DepositoryStateOfficeAddress = [" & RsMst_DepositoryStateWiseGST!DepositoryStateOfficeAddress & "] and new values are GSTNo = [" & Trim(.TextMatrix(mR, 6)) & "],GSTNavCode = [" & Trim(.TextMatrix(mR, 7)) & "],DepositoryStateOfficeAddress = [" & Trim(.TextMatrix(mR, 5)) & "]"
                       
                        ''-- Ullhas Patil 08 Nov 2017
                        ''-- Appply to block the update the GSTNO/ Address/ NAV Code as per instruct by Mayur Sir only power user can allow to upload with update existing entry
                        If mAllowUpdateFlag = True Then
                           ''-- Update Log
                           mExisitEntryFlag = True
                           .row = mR
                           .Col = .Cols - 1
                           .CellBackColor = vbYellow
                           .TextMatrix(mR, .Cols - 1) = "Mapping is update with old data with new data old values are GSTNo = [" & RsMst_DepositoryStateWiseGST!GSTNo & "],GSTNavCode = [" & RsMst_DepositoryStateWiseGST!GSTNavCode & "],DepositoryStateOfficeAddress = [" & RsMst_DepositoryStateWiseGST!DepositoryStateOfficeAddress & "] and new values are GSTNo = [" & Trim(.TextMatrix(mR, 6)) & "],GSTNavCode = [" & Trim(.TextMatrix(mR, 7)) & "],DepositoryStateOfficeAddress = [" & Trim(.TextMatrix(mR, 5)) & "]"
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
                    ''RsMst_DepositoryStateWiseGST!DepositoryStateOfficeAddress = Trim(.TextMatrix(mR, 5))
                    ''RsMst_DepositoryStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 6))
                    ''RsMst_DepositoryStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 7))
                    '''RsMst_DepositoryStateWiseGST!Flag = ""
                    '''RsMst_DepositoryStateWiseGST!Remarks = ""
                    ''If IsNull(RsMst_DepositoryStateWiseGST!CreatedBy) = True Or RsMst_DepositoryStateWiseGST!CreatedBy = "" Then
                    ''   RsMst_DepositoryStateWiseGST!CreatedBy = Gen_UserLoginID
                    ''   RsMst_DepositoryStateWiseGST!CreatedDate = Now
                    ''Else
                    ''   RsMst_DepositoryStateWiseGST!UpdatedBy = Gen_UserLoginID
                    ''   RsMst_DepositoryStateWiseGST!UpdatedDate = Now
                    ''End If
                    
                    ''-- Ullhas Patil 08 Nov 2017
                    If mAllowUpdateFlag = True Then
                        RsMst_DepositoryStateWiseGST!DepositoryStateOfficeAddress = Trim(.TextMatrix(mR, 5))
                        RsMst_DepositoryStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 6))
                        RsMst_DepositoryStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 7))

                        If IsNull(RsMst_DepositoryStateWiseGST!CreatedBy) = True Or RsMst_DepositoryStateWiseGST!CreatedBy = "" Then
                           RsMst_DepositoryStateWiseGST!CreatedBy = Gen_UserLoginID
                           RsMst_DepositoryStateWiseGST!CreatedDate = Now
                        Else
                           RsMst_DepositoryStateWiseGST!UpdatedBy = Gen_UserLoginID
                           RsMst_DepositoryStateWiseGST!UpdatedDate = Now
                        End If
                    Else
                        If mNewEntry = True Then
                           RsMst_DepositoryStateWiseGST!DepositoryStateOfficeAddress = Trim(.TextMatrix(mR, 5))
                           RsMst_DepositoryStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 6))
                           RsMst_DepositoryStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 7))
                           RsMst_DepositoryStateWiseGST!CreatedBy = Gen_UserLoginID
                           RsMst_DepositoryStateWiseGST!CreatedDate = Now
                        End If
                    End If
                    ''------------------------------------
                    
                    RsMst_DepositoryStateWiseGST.Update
                    RsMst_DepositoryStateWiseGST.Requery
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
        ''MsgBox "Some existing Depository-State entries are updated, so please export Excel for more details.", vbOKOnly + vbInformation
        If mAllowUpdateFlag = True Then
            MsgBox "Some existing Depository-State entries are updated, so please export Excel for more details.", vbOKOnly + vbInformation
        Else
            MsgBox "Existing Depository-State entries are not allow to update.", vbOKOnly + vbCritical
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
     .Cols = 10
     .FixedCols = 0
     .Font.Size = 9
     .AllowUserResizing = flexResizeColumns
     .WordWrap = True
     .TextMatrix(0, 0) = "Depository ID":                   .ColWidth(0) = 1000
     .TextMatrix(0, 1) = "Depository Name":                 .ColWidth(1) = 1500
     .TextMatrix(0, 2) = "Exchange Type":                   .ColWidth(2) = 1000
     .TextMatrix(0, 3) = "NAV Code":                        .ColWidth(3) = 1000
     .TextMatrix(0, 4) = "State India Name":                .ColWidth(4) = 1500
     .TextMatrix(0, 5) = "State Wise Depository Address":   .ColWidth(5) = 2000
     .TextMatrix(0, 6) = "GST No":                          .ColWidth(6) = 1600
     .TextMatrix(0, 7) = "GST NAV Code":                    .ColWidth(7) = 1600
     .TextMatrix(0, 8) = "StateIndiaID":                    .ColWidth(8) = 0
     .TextMatrix(0, 9) = "ID":                              .ColWidth(9) = 0
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_DepositoryStateWiseGST.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

With MSHFlexGrid1
For I = 0 To .Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To .Cols - 1
        If InStr(.TextMatrix(I, C), "/") > 0 Then
           If IsDate(.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(.TextMatrix(I, C), "dd-mmm-yyyy")
           Else
              oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
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
oXL.Quit
End Sub


