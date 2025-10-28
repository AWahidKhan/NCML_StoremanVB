VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTxn_ImportClientStateWiseGST 
   Caption         =   "Import Client State Wise GST Mapping"
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
      TabIndex        =   11
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox CmbExchangeType 
         Height          =   360
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox TxtNAVCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   650
         Width           =   1695
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   12660
         TabIndex        =   8
         Top             =   658
         Width           =   1095
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   10470
         TabIndex        =   6
         Top             =   658
         Width           =   1095
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
         Enabled         =   0   'False
         Height          =   345
         Left            =   9375
         TabIndex        =   5
         Top             =   658
         Width           =   1095
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   345
         Left            =   13755
         TabIndex        =   9
         Top             =   658
         Width           =   1095
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
         Enabled         =   0   'False
         Height          =   345
         Left            =   11565
         TabIndex        =   7
         Top             =   658
         Width           =   1095
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
         Height          =   345
         Left            =   8280
         TabIndex        =   4
         Top             =   658
         Width           =   1095
      End
      Begin VB.ComboBox CmbClient 
         Height          =   360
         Left            =   5040
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   9795
      End
      Begin VB.CommandButton CmdGet 
         Caption         =   "Get Data"
         Height          =   345
         Left            =   7200
         TabIndex        =   3
         Top             =   658
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   12
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
         TabIndex        =   10
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Exchange Type"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NAV Code"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   710
         Width           =   960
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1155
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Total Record :"
         Height          =   255
         Left            =   12240
         TabIndex        =   15
         Top             =   1155
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
         Left            =   13590
         TabIndex        =   14
         Top             =   1155
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Client Name"
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
         Left            =   3600
         TabIndex        =   13
         Top             =   300
         Width           =   1275
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
Attribute VB_Name = "FrmTxn_ImportClientStateWiseGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet

Dim ErrMsg As Variant
Dim Status As Variant
Dim mExchangeTypeID As Variant
Dim mClientIDRow, mClientID, mDepositoryID, mCRFID, mNAVCode As Variant
Dim ErrMsgFile As String

''-- Ullhas Patil 08 Nov 2017
Dim mAllowUpdateFlag As Boolean
Dim mNewEntry As Boolean

Private Sub CmbClient_GotFocus()
If CmbExchangeType.Text = "" Then
    MsgBox "Select Exchange Type"
    Exit Sub
End If
tmp = CmbClient.Text
Call TableMst_Client(" Where Status = 'A' And ExchangeTypeID = " & mExchangeTypeID & " Order By ClientName ")
CmbClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Client found"
   CmbClient.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClient.AddItem UCase(RsMst_Client!ClientName) & "#" & RsMst_Client!ClientIDRow & "#" & RsMst_Client!ClientID
    RsMst_Client.MoveNext
Next
CmbClient.Text = tmp
End Sub

Private Sub CmbClient_LostFocus()
Dim mSplit As Variant
If CmbClient.Text = "" Then
   Exit Sub
End If
mSplit = Split(CmbClient.Text, "#")
If UBound(mSplit) <> 2 Then
   MsgBox "Invalid selection "
   CmbClient.SetFocus
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientID = " & mSplit(2) & ""
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbClient.SetFocus
   Exit Sub
Else
   If LCase(RsMst_Client!ClientName) <> LCase(mSplit(0)) Then
    MsgBox "Invalid selection "
    CmbClient.SetFocus
    Exit Sub
   End If
End If
mClientIDRow = IIf(IsNull(RsMst_Client!ClientID), "", RsMst_Client!ClientID)
mClientID = IIf(IsNull(RsMst_Client!ClientIDRow), "", RsMst_Client!ClientIDRow)
mDepositoryID = IIf(IsNull(RsMst_Client!DepositoryID), "", RsMst_Client!DepositoryID)
mNAVCode = IIf(IsNull(RsMst_Client!NavCode), "", RsMst_Client!NavCode)
End Sub

Private Sub CmbExchangeType_GotFocus()
tmp = CmbExchangeType.Text
mTmp = CmbExchangeType.Text
Call TableMst_ExchangeType(" Where Flag = 'A' And ExchangeTypeID In (1,2) ")
CmbExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Exchange Type found"
   CmbExchangeType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeType.Text = tmp
End Sub
Private Sub CmbExchangeType_LostFocus()
If CmbExchangeType.Text = "" Then
   CmbClient.Text = ""
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid Exchange Type selection "
   CmbExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType.Fields(0)
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_ImportClientStateWiseGST.xls")
Call Xls_Detail_Rpt
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGet_Click()
Dim wc As Variant
If Trim(TxtNAVCode.Text) = "" Then
    If CmbExchangeType.Text = "" Then
        MsgBox "Select Exchange Type"
        Exit Sub
    End If
    If CmbClient.Text = "" Then
        MsgBox "Select Client"
        Exit Sub
    End If
End If
Call Grid_Head
wc = ""
If Trim(TxtNAVCode.Text) = "" Then
    wc = mClientIDRow
Else
    wc = " SELECT ClientID  FROM Mst_Client Where NAVCode = '" & Trim(TxtNAVCode) & "'"
End If
tmp = "SELECT A.ClientStateWiseGSTID,A.ClientID,A.ClientIDRow,A.ClientName,A.ExchangeTypeID,A.ExchangeType,A.NAVCode,A.CRFID,A.StateIndiaID,A.StateIndiaName,A.ClientStateOfficeAddress,A.GSTNo,A.GSTNavCode,A.RGSTNavCode From "
tmp = tmp & " ( SELECT MCSWG.ClientStateWiseGSTID,MCSWG.ClientID,MC.ClientIDRow,MC.ClientName,MC.ExchangeTypeID,MET.ExchangeType,MC.NAVCode,MC.CRFID,MCSWG.StateIndiaID,MSI.StateIndiaName,MCSWG.ClientStateOfficeAddress,MCSWG.GSTNo,MCSWG.GSTNavCode,MCSWG.RGSTNavCode"
tmp = tmp & " FROM Mst_ClientStateWiseGST MCSWG"
tmp = tmp & " Inner Join Mst_Client MC On MCSWG.ClientID = MC.ClientID"
tmp = tmp & " Inner Join Mst_ExchangeType MET On MC.ExchangeTypeID = MET.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_StateIndia MSI On MCSWG.StateIndiaID = MSI.StateIndiaID "
tmp = tmp & " Where MC.ClientID IN (" & wc & ")"
tmp = tmp & " Union"
tmp = tmp & " SELECT '0'ClientStateWiseGSTID,MC.ClientID,MC.ClientIDRow,MC.ClientName,MC.ExchangeTypeID,MET.ExchangeType,MC.NAVCode,MC.CRFID,MSI.StateIndiaID,MSI.StateIndiaName,''ClientStateOfficeAddress,''GSTNo,''GSTNavCode,''RGSTNavCode"
tmp = tmp & " FROM Mst_Client MC Inner Join Mst_ExchangeType MET On MC.ExchangeTypeID = MET.ExchangeTypeID , Mst_StateIndia MSI"
tmp = tmp & " Where MC.Status = 'A' And MC.ClientID IN (" & wc & ")"
tmp = tmp & " And Convert(Varchar(18),MC.ClientID)+'~'+Convert(Varchar(18),MSI.StateIndiaID) Not In (SELECT Convert(Varchar(18),ClientID) +'~'+Convert(Varchar(18),StateIndiaID) FROM Mst_ClientStateWiseGST Where ClientID IN (" & wc & "))) AS A "
tmp = tmp & " Order By A.ClientName,A.ClientID,A.StateIndiaName,A.ClientStateWiseGSTID"
Call TmpTable
If TmpRs.EOF = False Then
    
    With MSHFlexGrid1
    If TmpRs.RecordCount > 0 Then
       ProgressBar1.Max = TmpRs.RecordCount
       For mR = 1 To TmpRs.RecordCount
              
           .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
           .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!ClientIDRow), "", "`" + TmpRs!ClientIDRow)
           .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
           .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
           .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!CRFID), "", TmpRs!CRFID)
           .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
           .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!ClientStateOfficeAddress), "", TmpRs!ClientStateOfficeAddress)
           .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
           .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
           ''-- 14 Aug 2017
           .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!RGSTNavCode), "", TmpRs!RGSTNavCode)
           ''---------------
           .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
           .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
           .TextMatrix(mR, 12) = IIf(TmpRs!ClientStateWiseGSTID = 0, "", TmpRs!ClientStateWiseGSTID)
            
           .Rows = .Rows + 1
           TmpRs.MoveNext
           ProgressBar1.Value = mR
       Next
       CmdExcel.Enabled = True
    End If
    End With
    
End If
LblTotalRecords.Caption = TmpRs.RecordCount
MsgBox "Done."
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
MsgBox "Import Succesfull!!!"
CmdSave.SetFocus
End Sub

Private Sub CmdReset_Click()
    Call Grid_Head
    CmbExchangeType.Text = ""
    CmbClient.Text = ""
    TxtNAVCode = ""
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
Dim mErrorMsg, mStateIndiaNo, mCRFID, mTempExchangeTypeID, mCMSPCode, mSplitError As Variant
Dim mR, mC As Variant
Dim mExisitEntryFlag As Boolean
mExisitEntryFlag = False
With MSHFlexGrid1
    If .Rows > 2 Then
        ProgressBar1.Max = .Rows - 2
        Call TableMst_ClientStateWiseGST
        For mR = 1 To .Rows - 2
            mErrorMsg = ""
            mError = False
            ''--- Client Name for Checking
            If Trim(.TextMatrix(mR, 0)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Client Name not allowed." & vbNewLine
            Else
                Call TableMst_Client(" Where Status = 'A' And ClientName = '" & .TextMatrix(mR, 0) & "' And ClientIDRow = '" & Right(.TextMatrix(mR, 1), Len(.TextMatrix(mR, 1)) - 1) & "' And ClientID = " & Val(.TextMatrix(mR, 10)) & " ")
                If RsMst_Client.EOF = True Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Client Name not found or given Client Name is deactive." & vbNewLine
                End If
            End If
            ''--- Client ID for Checking
            If Trim(.TextMatrix(mR, 1)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Client ID not allowed." & vbNewLine
            End If
            ''--- Exchange Type for Checking
            If Trim(.TextMatrix(mR, 2)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank Exchange Type not allowed." & vbNewLine
            Else
                Call TableMst_Client(" Where Status = 'A' And ClientID = " & Val(.TextMatrix(mR, 10)) & "")
                If RsMst_Client.EOF = False Then
                    If LCase(GetExchangeName(RsMst_Client!ExchangeTypeID)) <> LCase(.TextMatrix(mR, 2)) Then
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Exchange Type is mismatch with given Client." & vbNewLine
                    End If
                End If
            End If
            ''--- NAV Code for Checking
            Call TableMst_Client(" Where Status = 'A' And ClientID = " & Val(.TextMatrix(mR, 10)) & "")
            If RsMst_Client.EOF = False Then
                If LCase(IIf(IsNull(RsMst_Client!NavCode), "", Trim(RsMst_Client!NavCode))) <> LCase(.TextMatrix(mR, 3)) Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "NAV Code is mismatch with given Client." & vbNewLine
                End If
            End If
            ''--- CRFID for Checking
            Call TableMst_Client(" Where Status = 'A' And ClientID = " & Val(.TextMatrix(mR, 10)) & "")
            If RsMst_Client.EOF = False Then
                If LCase(IIf(IsNull(RsMst_Client!CRFID), "", RsMst_Client!CRFID)) <> LCase(.TextMatrix(mR, 4)) Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "CRF ID is mismatch with given Client." & vbNewLine
                End If
            End If
            ''--- State India Name for Checking
            If Trim(.TextMatrix(mR, 5)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank State India Name not allowed." & vbNewLine
            Else
                Call TableMst_StateIndia(" Where Flag = 'A' And StateIndiaName = '" & Trim(.TextMatrix(mR, 5)) & "'  ")
                If RsMst_StateIndia.EOF = True Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "State India Name not found or given State India Name is deactive" & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 11)) <> RsMst_StateIndia!StateIndiaID Then
                       mError = True
                       If mError = True Then mErrorMsg = mErrorMsg & "State India Name and State India ID is mismatch." & vbNewLine
                    End If
                End If
            End If
            ''--- State Wise Bank Address Checking
            If Trim(.TextMatrix(mR, 6)) = "" Then
              mError = True
              If mError = True Then mErrorMsg = mErrorMsg & "Blank State Wise Client Address not allowed." & vbNewLine
            End If
            If Len(.TextMatrix(mR, 6)) > 500 Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "State Wise Client Address more then 500 characters not allowed." & vbNewLine
            End If
            ''--- GST No
            
            ''--- GST NAV Code for Checking
''            If Trim(.TextMatrix(mR, 8)) = "" Then
''              mError = True
''              If mError = True Then mErrorMsg = mErrorMsg & "Blank GST NAV Code not allowed." & vbNewLine
''            End If
            If Len(.TextMatrix(mR, 8)) > 25 Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "GST NAV Code more then 25 characters not allowed." & vbNewLine
            End If
            ''-- 14 Aug 2017
            ''--- GST RNAV Code for Checking
''            If Trim(.TextMatrix(mR, 9)) = "" Then
''              mError = True
''              If mError = True Then mErrorMsg = mErrorMsg & "Blank GST RNAV Code not allowed." & vbNewLine
''            End If
            If Len(.TextMatrix(mR, 9)) > 25 Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "GST RNAV Code more then 25 characters not allowed." & vbNewLine
            End If
            ''----------
            
            ''--- ClientIDRow for Checking
            mCRFID = ""
            mTempExchangeTypeID = ""
            mCMSPCode = ""
            If Trim(.TextMatrix(mR, 10)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank ClientIDRow allowed." & vbNewLine
            Else
                If Val(.TextMatrix(mR, 10)) = 0 Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Given ClientIDRow allowed." & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 10)) > 0 Then
                        Call TableMst_Client(" Where Status = 'A' And ClientID = " & Val(.TextMatrix(mR, 10)) & " ")
                        If RsMst_Client.EOF = True Then
                            mError = True
                            If mError = True Then mErrorMsg = mErrorMsg & "ClientIDRow not found or given ClientIDRow is deactive." & vbNewLine
                        Else
                            If LCase(.TextMatrix(mR, 0)) <> LCase(RsMst_Client!ClientName) Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "Client Name is mismatch with given ClientIDRow." & vbNewLine
                            End If
                            If LCase(Right(.TextMatrix(mR, 1), Len(.TextMatrix(mR, 1)) - 1)) <> LCase(RsMst_Client!ClientIDRow) Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "Client ID is mismatch with given ClientIDRow." & vbNewLine
                            End If
                            mCRFID = IIf(IsNull(RsMst_Client!CRFID), "", RsMst_Client!CRFID)
                            mTempExchangeTypeID = IIf(IsNull(RsMst_Client!ExchangeTypeID), "", RsMst_Client!ExchangeTypeID)
                            mCMSPCode = IIf(IsNull(RsMst_Client!CMSPCode), "", RsMst_Client!CMSPCode)
                        End If
                    End If
                End If
            End If
            ''--- StateIndia ID for Checking
            If Trim(.TextMatrix(mR, 11)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank StateIndiaID allowed." & vbNewLine
            Else
                If Val(.TextMatrix(mR, 11)) = 0 Then
                    mError = True
                    If mError = True Then mErrorMsg = mErrorMsg & "Given StateIndiaID allowed." & vbNewLine
                Else
                    If Val(.TextMatrix(mR, 11)) > 0 Then
                        Call TableMst_StateIndia(" Where Flag = 'A' And StateIndiaID = " & Val(.TextMatrix(mR, 11)) & " ")
                        If RsMst_StateIndia.EOF = True Then
                            mError = True
                            If mError = True Then mErrorMsg = mErrorMsg & "StateIndiaID not found or given StateIndiaID is deactive" & vbNewLine
                        Else
                            mStateIndiaNo = IIf(IsNull(RsMst_StateIndia!TINNo), 0, RsMst_StateIndia!TINNo)
                        End If
                    End If
                End If
            End If
            
            ''--- GST No for Checking
            If Trim(.TextMatrix(mR, 7)) = "" Then
                mError = True
                If mError = True Then mErrorMsg = mErrorMsg & "Blank GST No not allowed." & vbNewLine
            Else
                ''-- GST No Validation
                ''-- Validation is with specific Error
                mSplitError = Split(ValidateGSTIN(mStateIndiaNo, Trim(.TextMatrix(mR, 7))), "~")
                If CBool(mSplitError(0)) = False Then
                   mError = True
                   If mError = True Then mErrorMsg = mErrorMsg & CStr(mSplitError(1)) & vbNewLine
                Else
'''''                    ''-- Check for Duplicate GST No
'''''                    If RsMst_ClientStateWiseGST.EOF = False Then
'''''                        RsMst_ClientStateWiseGST.MoveFirst
'''''                        RsMst_ClientStateWiseGST.Find "GSTNo = '" & Trim(.TextMatrix(mR, 7)) & "'"
'''''                    End If
'''''                    If RsMst_ClientStateWiseGST.EOF = False Then
'''''                        If Val(.TextMatrix(mR, 11)) > 0 Then
'''''                            If mTempExchangeTypeID = 1 Then ''-- Non-NCDEX
'''''                                tmp = " SELECT Distinct MC.CRFID[CRFID] FROM Mst_ClientStateWiseGST MCSWG Inner Join Mst_Client MC On MCSWG.ClientIDRow = MC.ClientIDRow Where MCSWG.ClientIDRow = " & Val(RsMst_ClientStateWiseGST!ClientIDRow) & " "
'''''                                Call Tmp2Table
'''''                                If TmpRs2.RecordCount > 0 Then
'''''                                   If IsNull(TmpRs2!CRFID) = False Then
'''''                                    If Val(IIf(IsNull(TmpRs2!CRFID), "", TmpRs2!CRFID)) <> mCRFID Then
'''''                                        mError = True
'''''                                        If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number with differnet CRFID not allowed." & vbNewLine
'''''                                    End If
'''''                                   End If
'''''                                End If
'''''                            Else                            ''-- NCDEX
'''''                                tmp = " SELECT Distinct MC.CMSPCode[CMSPCode] FROM Mst_ClientStateWiseGST MCSWG Inner Join Mst_Client MC On MCSWG.ClientIDRow = MC.ClientIDRow Where MCSWG.ClientIDRow = " & Val(RsMst_ClientStateWiseGST!ClientIDRow) & " "
'''''                                Call Tmp2Table
'''''                                If TmpRs2.RecordCount > 0 Then
'''''                                   If IsNull(TmpRs2!CMSPCode) = False Then
'''''                                    If Val(IIf(IsNull(TmpRs2!CMSPCode), "", TmpRs2!CMSPCode)) <> mCMSPCode Then
'''''                                        mError = True
'''''                                        If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number with different NCDEX Client not allowed." & vbNewLine
'''''                                    End If
'''''                                   End If
'''''                                End If
'''''                            End If
'''''                        Else
'''''                            If mTempExchangeTypeID = 1 Then ''-- Non-NCDEX
'''''                                tmp = " SELECT Distinct MC.CRFID[CRFID] FROM Mst_ClientStateWiseGST MCSWG Inner Join Mst_Client MC On MCSWG.ClientIDRow = MC.ClientIDRow Where MCSWG.ClientIDRow = " & Val(RsMst_ClientStateWiseGST!ClientIDRow) & " "
'''''                                Call Tmp2Table
'''''                                If TmpRs2.RecordCount > 0 Then
'''''                                   If IsNull(TmpRs2!CRFID) = False Then
'''''                                    If Val(IIf(IsNull(TmpRs2!CRFID), "", TmpRs2!CRFID)) <> mCRFID Then
'''''                                        mError = True
'''''                                        If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number with differnet CRFID not allowed." & vbNewLine
'''''                                    End If
'''''                                   End If
'''''                                End If
'''''                            Else                            ''-- NCDEX
'''''                                tmp = " SELECT Distinct MC.CMSPCode[CMSPCode] FROM Mst_ClientStateWiseGST MCSWG Inner Join Mst_Client MC On MCSWG.ClientIDRow = MC.ClientIDRow Where MCSWG.ClientIDRow = " & Val(RsMst_ClientStateWiseGST!ClientIDRow) & " "
'''''                                Call Tmp2Table
'''''                                If TmpRs2.RecordCount > 0 Then
'''''                                   If IsNull(TmpRs2!CMSPCode) = False Then
'''''                                    If Val(IIf(IsNull(TmpRs2!CMSPCode), "", TmpRs2!CMSPCode)) <> mCMSPCode Then
'''''                                        mError = True
'''''                                        If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST Number with different NCDEX Client not allowed." & vbNewLine
'''''                                    End If
'''''                                   End If
'''''                                End If
'''''                            End If
'''''                        End If
'''''                    End If
                End If
            End If
            
            ''--- GST NAV Code Check for Duplicate GST No
'''''            If Trim(.TextMatrix(mR, 8)) <> "" Then
'''''                If RsMst_ClientStateWiseGST.EOF = False Then
'''''                    RsMst_ClientStateWiseGST.MoveFirst
'''''                    RsMst_ClientStateWiseGST.Find "GSTNavCode = '" & Trim(.TextMatrix(mR, 8)) & "'"
'''''                End If
'''''                If RsMst_ClientStateWiseGST.EOF = False Then
'''''                    If Val(.TextMatrix(mR, 11)) > 0 Then
'''''                        If mTempExchangeTypeID = 1 Then ''-- Non-NCDEX
'''''                            tmp = " SELECT Distinct MC.CRFID[CRFID] FROM Mst_ClientStateWiseGST MCSWG Inner Join Mst_Client MC On MCSWG.ClientIDRow = MC.ClientIDRow Where MCSWG.ClientIDRow = " & Val(RsMst_ClientStateWiseGST!ClientIDRow) & " "
'''''                            Call Tmp2Table
'''''                            If TmpRs2.RecordCount > 0 Then
'''''                               If IsNull(TmpRs2!CRFID) = False Then
'''''                                If Val(IIf(IsNull(TmpRs2!CRFID), "", TmpRs2!CRFID)) <> mCRFID Then
'''''                                    mError = True
'''''                                    If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST NAV Code with differnet CRFID not allowed." & vbNewLine
'''''                                End If
'''''                               End If
'''''                            End If
'''''                        Else                            ''-- NCDEX
'''''                            tmp = " SELECT Distinct MC.CMSPCode[CMSPCode] FROM Mst_ClientStateWiseGST MCSWG Inner Join Mst_Client MC On MCSWG.ClientIDRow = MC.ClientIDRow Where MCSWG.ClientIDRow = " & Val(RsMst_ClientStateWiseGST!ClientIDRow) & " "
'''''                            Call Tmp2Table
'''''                            If TmpRs2.RecordCount > 0 Then
'''''                               If IsNull(TmpRs2!CMSPCode) = False Then
'''''                                If Val(IIf(IsNull(TmpRs2!CMSPCode), "", TmpRs2!CMSPCode)) <> mCMSPCode Then
'''''                                    mError = True
'''''                                    If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST NAV Code for differnet NCDEX Client not allowed." & vbNewLine
'''''                                End If
'''''                               End If
'''''                            End If
'''''                        End If
'''''                    Else
'''''                        If mTempExchangeTypeID = 1 Then ''-- Non-NCDEX
'''''                            tmp = " SELECT Distinct MC.CRFID[CRFID] FROM Mst_ClientStateWiseGST MCSWG Inner Join Mst_Client MC On MCSWG.ClientIDRow = MC.ClientIDRow Where MCSWG.ClientIDRow = " & Val(RsMst_ClientStateWiseGST!ClientIDRow) & " "
'''''                            Call Tmp2Table
'''''                            If TmpRs2.RecordCount > 0 Then
'''''                               If IsNull(TmpRs2!CRFID) = False Then
'''''                                If Val(IIf(IsNull(TmpRs2!CRFID), "", TmpRs2!CRFID)) <> mCRFID Then
'''''                                    mError = True
'''''                                    If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST NAV Code with differnet CRFID not allowed." & vbNewLine
'''''                                End If
'''''                               End If
'''''                            End If
'''''                        Else                            ''-- NCDEX
'''''                            tmp = " SELECT Distinct MC.CMSPCode[CMSPCode] FROM Mst_ClientStateWiseGST MCSWG Inner Join Mst_Client MC On MCSWG.ClientIDRow = MC.ClientIDRow Where MCSWG.ClientIDRow = " & Val(RsMst_ClientStateWiseGST!ClientIDRow) & " "
'''''                            Call Tmp2Table
'''''                            If TmpRs2.RecordCount > 0 Then
'''''                               If IsNull(TmpRs2!CMSPCode) = False Then
'''''                                If Val(IIf(IsNull(TmpRs2!CMSPCode), "", TmpRs2!CMSPCode)) <> mCMSPCode Then
'''''                                    mError = True
'''''                                    If mError = True Then mErrorMsg = mErrorMsg & "Duplicate GST NAV Code for differnet NCDEX Client not allowed." & vbNewLine
'''''                                End If
'''''                               End If
'''''                            End If
'''''                        End If
'''''                    End If
'''''                End If
'''''            End If
            
            ''--- ID for Checking
            If Trim(.TextMatrix(mR, 12)) <> "" And Val(.TextMatrix(mR, 12)) > 0 Then
                If RsMst_ClientStateWiseGST.RecordCount > 0 Then
                    RsMst_ClientStateWiseGST.MoveFirst
                    RsMst_ClientStateWiseGST.Find "ClientStateWiseGSTID = " & Val(.TextMatrix(mR, 12)) & ""
                    If RsMst_ClientStateWiseGST.EOF = True Then
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Given ID is not available." & vbNewLine
                    Else
                        If Val(.TextMatrix(mR, 10)) > 0 And Val(.TextMatrix(mR, 11)) > 0 Then
                            If Val(.TextMatrix(mR, 10)) <> Val(RsMst_ClientStateWiseGST!ClientIDRow) Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "ClientIDRow is mismatch with Given ID." & vbNewLine
                            End If
                            If Val(.TextMatrix(mR, 11)) <> Val(RsMst_ClientStateWiseGST!StateIndiaID) Then
                                mError = True
                                If mError = True Then mErrorMsg = mErrorMsg & "StateIndiaID is mismatch with Given ID." & vbNewLine
                            End If
                        End If
                    End If
                End If
            End If
            
            ''-- Duplicate Entry by ClientID Row and StateIndia ID Checking
            If Val(.TextMatrix(mR, 10)) > 0 And Val(.TextMatrix(mR, 11)) > 0 Then
                If RsMst_ClientStateWiseGST.RecordCount > 0 Then
                   RsMst_ClientStateWiseGST.MoveFirst
                End If
                RsMst_ClientStateWiseGST.Filter = " ClientID = " & Val(.TextMatrix(mR, 10)) & " And StateIndiaID = " & Val(.TextMatrix(mR, 11)) & " "
                If RsMst_ClientStateWiseGST.EOF = False Then
                   If .TextMatrix(mR, 12) <> "" And Val(.TextMatrix(mR, 12)) > 0 Then
                        If Val(.TextMatrix(mR, 12)) <> Val(RsMst_ClientStateWiseGST!ClientStateWiseGSTID) Then
                            mError = True
                            If mError = True Then mErrorMsg = mErrorMsg & "Same ClientIDRow with Same State India not allowed." & vbNewLine
                        End If
                   Else
                        mError = True
                        If mError = True Then mErrorMsg = mErrorMsg & "Same ClientIDRow with Same State India not allowed." & vbNewLine
                   End If
                End If
                RsMst_ClientStateWiseGST.Filter = ""
            End If
            mNewEntry = False
            ''-- Error Log
            If mError = True Then
                .row = mR
                .Col = .Cols - 2
                .CellBackColor = vbRed
                .TextMatrix(mR, .Cols - 2) = mErrorMsg
            Else
                If Val(.TextMatrix(mR, 12)) >= 0 Then
                    If RsMst_ClientStateWiseGST.RecordCount > 0 Then
                       RsMst_ClientStateWiseGST.MoveFirst
                    End If
                    RsMst_ClientStateWiseGST.Find "ClientStateWiseGSTID = " & Val(.TextMatrix(mR, 12)) & ""
                    If RsMst_ClientStateWiseGST.EOF = True Then
                       RsMst_ClientStateWiseGST.AddNew
                       'RsMst_ClientStateWiseGST!ClientStateWiseGSTID = GetMaxID("ClientStateWiseGSTID", "Mst_ClientStateWiseGST")
                       '.TextMatrix(mR, 12) = RsMst_ClientStateWiseGST!ClientStateWiseGSTID
                       RsMst_ClientStateWiseGST!G_UID = GetGUID
                       RsMst_ClientStateWiseGST!ClientID = Val(.TextMatrix(mR, 10))
                       RsMst_ClientStateWiseGST!StateIndiaID = Val(.TextMatrix(mR, 11))
                       mNewEntry = True
                    Else
                       ''''-- Update Log
                       ''mExisitEntryFlag = True
                       ''.row = mR
                       ''.Col = .Cols - 1
                       ''.CellBackColor = vbYellow
                       ''.TextMatrix(mR, .Cols - 1) = "Mapping is update with old data with new data old values are GSTNo = [" & RsMst_ClientStateWiseGST!GSTNo & "],GSTNavCode = [" & RsMst_ClientStateWiseGST!GSTNavCode & "],ClientStateOfficeAddress = [" & RsMst_ClientStateWiseGST!ClientStateOfficeAddress & "] and new values are GSTNo = [" & Trim(.TextMatrix(mR, 7)) & "],GSTNavCode = [" & Trim(.TextMatrix(mR, 8)) & "],ClientStateOfficeAddress = [" & Trim(.TextMatrix(mR, 6)) & "]"
                       
                        ''-- Ullhas Patil 08 Nov 2017
                        ''-- Appply to block the update the GSTNO/ Address/ NAV Code as per instruct by Mayur Sir only power user can allow to upload with update existing entry
                        If mAllowUpdateFlag = True Then
                           ''-- Update Log
                           mExisitEntryFlag = True
                           .row = mR
                           .Col = .Cols - 1
                           .CellBackColor = vbYellow
                           .TextMatrix(mR, .Cols - 1) = "Mapping is update with old data with new data old values are GSTNo = [" & RsMst_ClientStateWiseGST!GSTNo & "],GSTNavCode = [" & RsMst_ClientStateWiseGST!GSTNavCode & "],ClientStateOfficeAddress = [" & RsMst_ClientStateWiseGST!ClientStateOfficeAddress & "] and new values are GSTNo = [" & Trim(.TextMatrix(mR, 7)) & "],GSTNavCode = [" & Trim(.TextMatrix(mR, 8)) & "],ClientStateOfficeAddress = [" & Trim(.TextMatrix(mR, 6)) & "]"
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
                    
                    ''RsMst_ClientStateWiseGST!ClientStateOfficeAddress = Trim(.TextMatrix(mR, 6))
                    ''RsMst_ClientStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 7))
                    ''RsMst_ClientStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 8))
                    ''RsMst_ClientStateWiseGST!RGSTNavCode = Trim(.TextMatrix(mR, 9)) ''-- 14 Aug 2017
                    ''If IsNull(RsMst_ClientStateWiseGST!CreatedBy) = True Or RsMst_ClientStateWiseGST!CreatedBy = "" Then
                    ''   RsMst_ClientStateWiseGST!CreatedBy = Gen_UserLoginID
                    ''   RsMst_ClientStateWiseGST!CreatedDate = Now
                    ''Else
                    ''    RsMst_ClientStateWiseGST!UpdatedBy = Gen_UserLoginID
                    ''    RsMst_ClientStateWiseGST!UpdatedDate = Now
                    ''End If
                    
                    ''-- Ullhas Patil 08 Nov 2017
                    If mAllowUpdateFlag = True Then
                        RsMst_ClientStateWiseGST!ClientStateOfficeAddress = Trim(.TextMatrix(mR, 6))
                        RsMst_ClientStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 7))
                        RsMst_ClientStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 8))
                        RsMst_ClientStateWiseGST!RGSTNavCode = Trim(.TextMatrix(mR, 9)) ''-- 14 Aug 2017
                        
                        If IsNull(RsMst_ClientStateWiseGST!CreatedBy) = True Or RsMst_ClientStateWiseGST!CreatedBy = "" Then
                           RsMst_ClientStateWiseGST!CreatedBy = Gen_UserLoginID
                           RsMst_ClientStateWiseGST!CreatedDate = Now
                        Else
                            RsMst_ClientStateWiseGST!UpdatedBy = Gen_UserLoginID
                            RsMst_ClientStateWiseGST!UpdatedDate = Now
                        End If
                    Else
                        If mNewEntry = True Then
                           RsMst_ClientStateWiseGST!ClientStateOfficeAddress = Trim(.TextMatrix(mR, 6))
                           RsMst_ClientStateWiseGST!GSTNo = Trim(.TextMatrix(mR, 7))
                           RsMst_ClientStateWiseGST!GSTNavCode = Trim(.TextMatrix(mR, 8))
                           RsMst_ClientStateWiseGST!RGSTNavCode = Trim(.TextMatrix(mR, 9)) ''-- 14 Aug 2017
                           RsMst_ClientStateWiseGST!CreatedBy = Gen_UserLoginID
                           RsMst_ClientStateWiseGST!CreatedDate = Now
                        End If
                    End If
                    ''------------------------------------
                        
                    RsMst_ClientStateWiseGST.Update
                    RsMst_ClientStateWiseGST.Requery
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
        If mAllowUpdateFlag = True Then
            MsgBox "Some existing Client-State entries are updated, so please export Excel for more details.", vbOKOnly + vbInformation
        Else
            MsgBox "Existing Client-State entries are not allow to update.", vbOKOnly + vbCritical
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
CmbExchangeType = ""
CmbClient = ""
TxtNAVCode = ""
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
     .Cols = 13
     .FixedCols = 0
     .Font.Size = 9
     .RowHeight(0) = 600
     .AllowUserResizing = flexResizeColumns
     .WordWrap = True
     .TextMatrix(0, 0) = "Client Name":                 .ColWidth(0) = 1500
     .TextMatrix(0, 1) = "Client ID":                   .ColWidth(1) = 1000
     .TextMatrix(0, 2) = "Exchange Type":               .ColWidth(2) = 1000
     .TextMatrix(0, 3) = "NAV Code":                    .ColWidth(3) = 1000
     .TextMatrix(0, 4) = "CRFID":                       .ColWidth(4) = 800
     .TextMatrix(0, 5) = "State India Name":            .ColWidth(5) = 1500
     .TextMatrix(0, 6) = "State Wise Client Address":   .ColWidth(6) = 2000
     .TextMatrix(0, 7) = "GST No":                      .ColWidth(7) = 1600
     .TextMatrix(0, 8) = "GST NAV Code":                .ColWidth(8) = 1600
     ''--- 14 Aug 2017
     .TextMatrix(0, 9) = "GST RNAV Code":                .ColWidth(9) = 1600
     ''-----------------
     .TextMatrix(0, 10) = "ClientIDRow":                .ColWidth(10) = 0
     .TextMatrix(0, 11) = "StateIndiaID":               .ColWidth(11) = 0
     .TextMatrix(0, 12) = "ID":                         .ColWidth(12) = 0
     
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_ImportClientStateWiseGST.xls")
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
Dim mR, mC As Variant
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
