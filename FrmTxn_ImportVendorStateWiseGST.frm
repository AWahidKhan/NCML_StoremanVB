VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTxn_ImportVendorStateWiseGST 
   Caption         =   "Import Vendor State Wise GST"
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
      TabIndex        =   10
      Top             =   120
      Width           =   15015
      Begin VB.CommandButton CmdGet 
         Caption         =   "Get Data"
         Height          =   345
         Left            =   7200
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox CmbVendor 
         Height          =   360
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   9195
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
         Height          =   345
         Left            =   8280
         TabIndex        =   3
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
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   345
         Left            =   13755
         TabIndex        =   8
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
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   10470
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   12660
         TabIndex        =   7
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
         TabIndex        =   9
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendor Name"
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
         Width           =   1440
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmTxn_ImportVendorStateWiseGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet

Dim ErrMsg As Variant
Dim Status As Variant
Dim mVendorID As Variant
Dim ErrMsgFile As String

''-- Ullhas Patil 08 Nov 2017
Dim mAllowUpdateFlag As Boolean
Dim mNewEntry As Boolean

Private Sub CmbVendor_GotFocus()
tmp = CmbVendor.Text
mTmp = CmbVendor.Text
Call TableMst_Vendor
CmbVendor.Clear
If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Vendor found"
   CmbVendor.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Vendor.RecordCount
    CmbVendor.AddItem RsMst_Vendor!VendorName
    RsMst_Vendor.MoveNext
Next
CmbVendor.Text = tmp
End Sub
Private Sub CmbVendor_LostFocus()
If CmbVendor.Text = "" Then
   Exit Sub
End If
RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbVendor.Text & "'"
If RsMst_Vendor.EOF Then
   MsgBox "Invalid selection "
   CmbVendor.SetFocus
   Exit Sub
End If
mVendorID = IIf(IsNull(RsMst_Vendor!VendorID), "", RsMst_Vendor!VendorID)
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_ImportVendorStateWiseGST.xls")
Call Xls_Detail_Rpt
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGet_Click()
Dim wc As Variant
If Trim(TxtNAVCode.Text) = "" Then
    If CmbVendor.Text = "" Then
        MsgBox "Select Agency"
        Exit Sub
    End If
End If
Call Grid_Head

wc = ""
If Trim(TxtNAVCode.Text) = "" Then
    wc = mVendorID
Else
    wc = " SELECT VendorID FROM Mst_Vendor Where NAVCode = '" & Trim(TxtNAVCode) & "'"
End If

tmp = " SELECT A.VendorStateWiseGSTID,A.VendorID,A.VendorName,A.NAVCode,A.StateIndiaID,A.StateIndiaName,A.VendorStateOfficeAddress,A.GSTNo,A.GSTNavCode"
tmp = tmp & " FROM ("
tmp = tmp & " SELECT MVSWG.VendorStateWiseGSTID,MVSWG.VendorID,MV.VendorName,MV.NAVCode,MVSWG.StateIndiaID,MSI.StateIndiaName,MVSWG.VendorStateOfficeAddress,MVSWG.GSTNo,MVSWG.GSTNavCode"
tmp = tmp & " FROM Mst_VendorStateWiseGST MVSWG"
tmp = tmp & " Inner Join Mst_Vendor MV On MVSWG.VendorID = MV.VendorID"
tmp = tmp & " Inner Join Mst_StateIndia MSI On MVSWG.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Where MVSWG.VendorID IN (" & wc & ")"
tmp = tmp & " Union"
tmp = tmp & " SELECT '0'VendorStateWiseGSTID,MV.VendorID,MV.VendorName,MV.NAVCode,MSI.StateIndiaID,MSI.StateIndiaName,''VendorStateOfficeAddress,''GSTNo,''GSTNavCode"
tmp = tmp & " FROM Mst_Vendor MV,Mst_StateIndia MSI"
tmp = tmp & " Where MV.VendorID IN (" & wc & ") And Convert(Varchar(18),MV.VendorID)+'~'+Convert(Varchar(18),MSI.StateIndiaID) Not in (SELECT Convert(Varchar(18),VendorID) +'~'+Convert(Varchar(18),StateIndiaID) From Mst_VendorStateWiseGST Where VendorID IN (SELECT VendorID FROM Mst_Vendor Where VendorID IN (" & wc & ")))"
tmp = tmp & " ) A Order By A.VendorName,A.VendorID,A.StateIndiaName,A.VendorStateOfficeAddress"
Call TmpTable

If TmpRs.EOF = False Then
    
    With MSHFlexGrid1
    If TmpRs.RecordCount > 0 Then
       ProgressBar1.Max = TmpRs.RecordCount
       For mR = 1 To TmpRs.RecordCount
              
           .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!VendorName), "", TmpRs!VendorName)
           .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
           .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
           .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!VendorStateOfficeAddress), "", TmpRs!VendorStateOfficeAddress)
           .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
           .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
           .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!VendorID), "", TmpRs!VendorID)
           .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
           .TextMatrix(mR, 8) = IIf(TmpRs!VendorStateWiseGSTID = 0, "", TmpRs!VendorStateWiseGSTID)

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
    CmbVendor.Text = ""
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

With MSHFlexGrid1
    If .Rows > 2 Then
        ProgressBar1.Max = .Rows - 2
        Call TableMst_VendorStateWiseGST
        For mR = 1 To .Rows - 2
            mErrorMsg = ""
            mError = False
            ''--- Vendor Name for Checking
            If Trim(.TextMatrix(mR, 0)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Vendor Name not allowed." & vbNewLine
            Else
                Call TableMst_Vendor(" Where VendorName = '" & Trim(.TextMatrix(mR, 0)) & "'")
                If RsMst_Vendor.EOF = True Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Vendor Name not found." & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 6)) <> RsMst_Vendor!VendorID Then
                       mError = True
                       If mError = True Then mErrorMsg = mErrorMsg & "Vendor Name and Vendor ID is mismatch." & vbNewLine
                    End If
                End If
            End If
            
            ''--- NAV Code for Checking
            Call TableMst_Vendor(" Where VendorID = '" & Val(.TextMatrix(mR, 6)) & "'")
            If RsMst_Vendor.EOF = False Then
                If LCase(IIf(IsNull(RsMst_Vendor!NavCode), "", RsMst_Vendor!NavCode)) <> LCase(.TextMatrix(mR, 1)) Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "NAV Code is mismatch with given Vendor ID." & vbNewLine
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
            ''--- Vendor ID for Checking
            If Trim(.TextMatrix(mR, 6)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Vendor ID allowed." & vbNewLine
            Else
                If Val(.TextMatrix(mR, 6)) = 0 Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Given Vendor ID allowed." & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 6)) > 0 Then
                        Call TableMst_Vendor(" Where VendorID = " & Val(.TextMatrix(mR, 6)) & " ")
                        If RsMst_Vendor.EOF = True Then
                            mError = True
                            If mError = True Then mErrorMsg = mErrorMsg & "Vendor ID not found." & vbNewLine
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
                ''-- GST No Validation
                ''-- Validation is with specific Error
                mSplitError = Split(ValidateGSTIN(mStateIndiaNo, Trim(.TextMatrix(mR, 4))), "~")
                If CBool(mSplitError(0)) = False Then
                   mError = True
                   If mError = True Then mErrorMsg = mErrorMsg & CStr(mSplitError(1)) & vbNewLine
                Else
''''                    ''-- Check for Duplicate GST No
''''                    If RsMst_VendorStateWiseGST.EOF = False Then
''''                        RsMst_VendorStateWiseGST.MoveFirst
''''                        RsMst_VendorStateWiseGST.Find "GSTNo = '" & Trim(.TextMatrix(mR, 4)) & "'"
''''                    End If
''''                    If RsMst_VendorStateWiseGST.EOF = False Then
''''                        If Val(.TextMatrix(mR, 8)) > 0 Then
''''                            If Val(.TextMatrix(mR, 8)) <> Val(RsMst_VendorStateWiseGST!VendorStateWiseGSTID) Then
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
''''                If RsMst_VendorStateWiseGST.EOF = False Then
''''                    RsMst_VendorStateWiseGST.MoveFirst
''''                    RsMst_VendorStateWiseGST.Find "GSTNavCode = '" & Trim(.TextMatrix(mR, 5)) & "'"
''''                End If
''''                If RsMst_VendorStateWiseGST.EOF = False Then
''''                    If Val(.TextMatrix(mR, 8)) > 0 Then
''''                        If Val(.TextMatrix(mR, 8)) <> Val(RsMst_VendorStateWiseGST!VendorStateWiseGSTID) Then
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
                If RsMst_VendorStateWiseGST.RecordCount > 0 Then
                    RsMst_VendorStateWiseGST.MoveFirst
                    RsMst_VendorStateWiseGST.Find "VendorStateWiseGSTID = " & Val(.TextMatrix(mR, 8)) & ""
                    If RsMst_VendorStateWiseGST.EOF = True Then
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Given ID is not available." & vbNewLine
                    Else
                        If Val(.TextMatrix(mR, 6)) > 0 And Val(.TextMatrix(mR, 7)) > 0 Then
                            If Val(.TextMatrix(mR, 6)) <> Val(RsMst_VendorStateWiseGST!VendorID) Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "VendorID is mismatch with Given ID." & vbNewLine
                            End If
                            If Val(.TextMatrix(mR, 7)) <> Val(RsMst_VendorStateWiseGST!StateIndiaID) Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "StateIndiaID is mismatch with Given ID." & vbNewLine
                            End If
                        End If
                    End If
                End If
            End If
            
            ''-- Duplicate Entry by Lessor ID and StateIndia ID
            If Val(.TextMatrix(mR, 6)) > 0 And Val(.TextMatrix(mR, 7)) > 0 Then
                If RsMst_VendorStateWiseGST.RecordCount > 0 Then
                   RsMst_VendorStateWiseGST.MoveFirst
                End If
                RsMst_VendorStateWiseGST.Filter = " VendorID = " & Val(.TextMatrix(mR, 6)) & " And StateIndiaID = " & Val(.TextMatrix(mR, 7)) & " "
                If RsMst_VendorStateWiseGST.EOF = False Then
                   If .TextMatrix(mR, 8) <> "" And Val(.TextMatrix(mR, 8)) > 0 Then
                        If Val(.TextMatrix(mR, 8)) <> Val(RsMst_VendorStateWiseGST!VendorStateWiseGSTID) Then
                            mError = True
                            If mError = True Then mErrorMsg = mErrorMsg & "Same Vendor and Same State India not allowed." & vbNewLine
                        End If
                   Else
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Same Vendor and Same State India not allowed." & vbNewLine
                   End If
                End If
                RsMst_VendorStateWiseGST.Filter = ""
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
                    If RsMst_VendorStateWiseGST.RecordCount > 0 Then
                       RsMst_VendorStateWiseGST.MoveFirst
                    End If
                    RsMst_VendorStateWiseGST.Find "VendorStateWiseGSTID = " & Val(.TextMatrix(mR, 8)) & ""
                    If RsMst_VendorStateWiseGST.EOF = True Then
                       RsMst_VendorStateWiseGST.AddNew
                       RsMst_VendorStateWiseGST!VendorStateWiseGSTID = GetMaxID("VendorStateWiseGSTID", "Mst_VendorStateWiseGST")
                       .TextMatrix(mR, 8) = RsMst_VendorStateWiseGST!VendorStateWiseGSTID
                       RsMst_VendorStateWiseGST!G_UID = GetGUID
                       RsMst_VendorStateWiseGST!VendorID = Val(.TextMatrix(mR, 6))
                       RsMst_VendorStateWiseGST!StateIndiaID = Val(.TextMatrix(mR, 7))
                       mNewEntry = True
                    Else
                       ''-- Update Log
                       ''mExisitEntryFlag = True
                       ''.row = mR
                       ''.Col = .Cols - 1
                       ''.CellBackColor = vbYellow
                       ''.TextMatrix(mR, .Cols - 1) = "Mapping is update with old data with new data old values are GSTNo = [" & RsMst_VendorStateWiseGST!GSTNo & "],GSTNavCode = [" & RsMst_VendorStateWiseGST!GSTNavCode & "],VendorStateOfficeAddress = [" & RsMst_VendorStateWiseGST!VendorStateOfficeAddress & "] and new values are GSTNo = [" & Trim(.TextMatrix(mR, 4)) & "],GSTNavCode = [" & Trim(.TextMatrix(mR, 5)) & "],VendorStateOfficeAddress = [" & Trim(.TextMatrix(mR, 3)) & "]"
                       
                        ''-- Ullhas Patil 08 Nov 2017
                        ''-- Appply to block the update the GSTNO/ Address/ NAV Code as per instruct by Mayur Sir only power user can allow to upload with update existing entry
                        If mAllowUpdateFlag = True Then
                           ''-- Update Log
                           mExisitEntryFlag = True
                           .row = mR
                           .Col = .Cols - 1
                           .CellBackColor = vbYellow
                           .TextMatrix(mR, .Cols - 1) = "Mapping is update with old data with new data old values are GSTNo = [" & RsMst_VendorStateWiseGST!GSTNo & "],GSTNavCode = [" & RsMst_VendorStateWiseGST!GSTNavCode & "],VendorStateOfficeAddress = [" & RsMst_VendorStateWiseGST!VendorStateOfficeAddress & "] and new values are GSTNo = [" & Trim(.TextMatrix(mR, 4)) & "],GSTNavCode = [" & Trim(.TextMatrix(mR, 5)) & "],VendorStateOfficeAddress = [" & Trim(.TextMatrix(mR, 3)) & "]"
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
                    
                    ''RsMst_VendorStateWiseGST!VendorStateOfficeAddress = Trim(.TextMatrix(mR, 3))
                    ''RsMst_VendorStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 4))
                    ''RsMst_VendorStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 5))
                    'RsMst_VendorStateWiseGST!Flag = ""
                    'RsMst_VendorStateWiseGST!Remarks = ""
                    ''If IsNull(RsMst_VendorStateWiseGST!CreatedBy) = True Or IsNull(RsMst_VendorStateWiseGST!CreatedDate) = True Then
                    ''   RsMst_VendorStateWiseGST!CreatedBy = Gen_UserLoginID
                    ''   RsMst_VendorStateWiseGST!CreatedDate = Now
                    ''Else
                    ''   RsMst_VendorStateWiseGST!UpdatedBy = Gen_UserLoginID
                    ''   RsMst_VendorStateWiseGST!UpdatedDate = Now
                    ''End If
                    
                    ''-- Ullhas Patil 08 Nov 2017
                    If mAllowUpdateFlag = True Then
                        RsMst_VendorStateWiseGST!VendorStateOfficeAddress = Trim(.TextMatrix(mR, 3))
                        RsMst_VendorStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 4))
                        RsMst_VendorStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 5))
                        If IsNull(RsMst_VendorStateWiseGST!CreatedBy) = True Or IsNull(RsMst_VendorStateWiseGST!CreatedDate) = True Then
                           RsMst_VendorStateWiseGST!CreatedBy = Gen_UserLoginID
                           RsMst_VendorStateWiseGST!CreatedDate = Now
                        Else
                           RsMst_VendorStateWiseGST!UpdatedBy = Gen_UserLoginID
                           RsMst_VendorStateWiseGST!UpdatedDate = Now
                        End If
                    Else
                        If mNewEntry = True Then
                           RsMst_VendorStateWiseGST!VendorStateOfficeAddress = Trim(.TextMatrix(mR, 3))
                           RsMst_VendorStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 4))
                           RsMst_VendorStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 5))
                           RsMst_VendorStateWiseGST!CreatedBy = Gen_UserLoginID
                           RsMst_VendorStateWiseGST!CreatedDate = Now
                        End If
                    End If
                    ''------------------------------------
                    
                    RsMst_VendorStateWiseGST.Update
                    RsMst_VendorStateWiseGST.Requery
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
        ''MsgBox "Some existing Vendor-State entries are updated, so please export Excel for more details.", vbOKOnly + vbInformation
        If mAllowUpdateFlag = True Then
            MsgBox "Some existing Vendor-State entries are updated, so please export Excel for more details.", vbOKOnly + vbInformation
        Else
            MsgBox "Existing Vendor-State entries are not allow to update.", vbOKOnly + vbCritical
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
     .WordWrap = False
     .TextMatrix(0, 0) = "Vendor Name":                .ColWidth(0) = 1500
     .TextMatrix(0, 1) = "NAV Code":                   .ColWidth(1) = 1000
     .TextMatrix(0, 2) = "State India Name":           .ColWidth(2) = 1500
     .TextMatrix(0, 3) = "State Wise Vendor Address":  .ColWidth(3) = 2000
     .TextMatrix(0, 4) = "GST No":                     .ColWidth(4) = 1500
     .TextMatrix(0, 5) = "GST NAV Code":               .ColWidth(5) = 1500
     .TextMatrix(0, 6) = "VendorID":                   .ColWidth(6) = 0
     .TextMatrix(0, 7) = "StateIndiaID":               .ColWidth(7) = 0
     .TextMatrix(0, 8) = "ID":                         .ColWidth(8) = 0
     
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_ImportVendorStateWiseGST.xls")
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

