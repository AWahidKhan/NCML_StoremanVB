VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_LocationInvoiceMapping 
   Caption         =   "Location Mapping to Tax Invoices"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Invoice Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7005
      Left            =   120
      TabIndex        =   11
      Top             =   1455
      Width           =   14970
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   315
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6285
         Left            =   120
         TabIndex        =   4
         Top             =   585
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   11086
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   370
         AllowBigSelection=   0   'False
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
      Begin VB.Label Label3 
         Caption         =   "Total Invoices Found :-"
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
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "_"
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
         Left            =   4920
         TabIndex        =   14
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1425
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   14985
      Begin VB.CheckBox ChkDP 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   0
         Top             =   195
         Width           =   210
      End
      Begin VB.ComboBox CmbCNonNcdexClient 
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   12015
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   1635
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   975
         Width           =   12030
      End
      Begin VB.Label Label1 
         Caption         =   "Show DP"
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
         Left            =   105
         TabIndex        =   13
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "Client Name "
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
         Left            =   105
         TabIndex        =   12
         Top             =   533
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   10
         Top             =   1005
         Width           =   1560
      End
   End
   Begin VB.Frame Frame3 
      Height          =   630
      Left            =   120
      TabIndex        =   7
      Top             =   8445
      Width           =   14985
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   11220
         TabIndex        =   5
         Top             =   165
         Width           =   1740
      End
      Begin VB.CommandButton ExitCmd 
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
         Height          =   400
         Left            =   12990
         TabIndex        =   6
         Top             =   165
         Width           =   1740
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   120
      TabIndex        =   16
      Top             =   9120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmTxn_LocationInvoiceMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mClientIDRow, mLocationID, mDPID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mClientNAVCode, mClientRNAVCode, mDPNAVCode As Variant

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.FixedCols = 1
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = " "
MSHFlexGrid1.TextMatrix(0, 1) = "Selection"
MSHFlexGrid1.TextMatrix(0, 2) = "Invoice No"
MSHFlexGrid1.TextMatrix(0, 3) = "Narration"
MSHFlexGrid1.TextMatrix(0, 4) = "EntryNo" '"Invoice Amount"

MSHFlexGrid1.ColWidth(0) = 200
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 12000
MSHFlexGrid1.ColWidth(4) = 0

MSHFlexGrid1.RowHeight(0) = 600

End Sub

Private Sub ChkDP_Click()
CmbCNonNcdexClient.Text = ""
CmbCNonNcdexClient.Clear
Call Grid_Head
End Sub

Private Sub CmbCNonNcdexClient_GotFocus()

tmp1 = CmbCNonNcdexClient.Text

If ChkDP.Value = 0 Then
    tmp = "Select * from Mst_Client Where NAVCode in  "
    tmp = tmp & " (Select TT.NAVCode from Txn_TIDetails TT "
    tmp = tmp & " inner join Txn_ClientLedger TC on TT.EntryNo=tc.EntryNo  "
    tmp = tmp & " Where TT.LocationID is null) "
    
    Call TmpTable
    
    CmbCNonNcdexClient.Clear
    
    If TmpRs.RecordCount = 0 Then
       MsgBox "No record found!!! "
       Exit Sub
    End If
    
    For I = 1 To TmpRs.RecordCount
        CmbCNonNcdexClient.AddItem TmpRs!ClientName & "~" & TmpRs!ClientID & "~" & TmpRs!ClientIDRow
        TmpRs.MoveNext
    Next
Else
    tmp = "Select * from Mst_Depository Where NAVCode in  "
    tmp = tmp & " (Select TT.NAVCode from Txn_TIDetails TT "
    tmp = tmp & " inner join Txn_ClientLedger TC on TT.EntryNo=tc.EntryNo  "
    tmp = tmp & " Where TT.LocationID is null) "
    
    Call TmpTable
    
    CmbCNonNcdexClient.Clear
    
    If TmpRs.RecordCount = 0 Then
       MsgBox "No record found!!! "
       Exit Sub
    End If
    
    For I = 1 To TmpRs.RecordCount
        CmbCNonNcdexClient.AddItem TmpRs!DepositoryName & "~" & TmpRs!DepositoryID
        TmpRs.MoveNext
    Next
End If
CmbCNonNcdexClient.Text = tmp1

End Sub

Private Sub CmbCNonNcdexClient_LostFocus()

Dim strSplitString() As String
Dim ClientNo_, ClientName_, ClientIDRow_ As String
Dim DPID_, DPName_ As String

If CmbCNonNcdexClient.Text = "" Then
   mClientRowID = Null
   Exit Sub
End If

If tmp1 <> CmbCNonNcdexClient.Text Then
   Call Grid_Head
End If

strSplitString = Split(CmbCNonNcdexClient, "~")

If ChkDP.Value = 0 Then
    If UBound(strSplitString) <> 2 Then
       MsgBox "Invalid Selection!!!"
       CmbCNonNcdexClient.SetFocus
       Exit Sub
    End If
    
    
    ClientName_ = Trim(strSplitString(0))
    ClientNo_ = Trim(strSplitString(1))
    ClientIDRow_ = Trim(strSplitString(2))
    
    TmpRs.MoveFirst
    TmpRs.Find "ClientIDRow = " & ClientIDRow_ & ""
    
    If TmpRs.EOF Then
       MsgBox "Invalid Selection "
       CmbCNonNcdexClient.SetFocus
       Exit Sub
    End If
    
    mClientIDRow = TmpRs!ClientIDRow
    mClientNAVCode = TmpRs!NAVCode
    mClientRNAVCode = TmpRs!RNAVCode
Else
    If UBound(strSplitString) <> 1 Then
       MsgBox "Invalid Selection!!!"
       CmbCNonNcdexClient.SetFocus
       Exit Sub
    End If
    
    DPName_ = Trim(strSplitString(0))
    DPID_ = Trim(strSplitString(1))
    
    TmpRs.MoveFirst
    TmpRs.Find "DepositoryID = '" & DPID_ & "'"
    
    If TmpRs.EOF Then
       MsgBox "Invalid Selection "
       CmbCNonNcdexClient.SetFocus
       Exit Sub
    End If
    
    mDPID = TmpRs!DepositoryID
    mDPNAVCode = TmpRs!NAVCode
End If

If tmp1 <> CmbCNonNcdexClient.Text Then
   Call GetGridData
End If

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True

Call ClearData

End Sub

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
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = Null
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

Private Sub GetGridData()

ProgressBar1.Value = 0

Call Grid_Head

tmp = " Select TT.TINo,TC.Narration,TT.EntryNo from Txn_TIDetails TT "
tmp = tmp & " inner join Txn_ClientLedger TC on TT.EntryNo=tc.EntryNo "
tmp = tmp & " Where TT.LocationID Is Null "
If ChkDP.Value = 0 Then
   tmp = tmp & " And (TT.NAVCode = '" & mClientNAVCode & "' Or TT.NAVCode = '" & mClientRNAVCode & "')"
Else
   tmp = tmp & " And (TT.NAVCode = '" & mDPNAVCode & "')"
End If

Call Tmp1Table

ProgressBar1.Max = TmpRs1.RecordCount + 1

For I = 1 To TmpRs1.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 2 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs1.Fields(c - 2)), "", TmpRs1.Fields(c - 2))
    Next
    MSHFlexGrid1.Row = I
    MSHFlexGrid1.Col = 1
    MSHFlexGrid1.CellFontName = "Wingdings"
    MSHFlexGrid1.CellFontSize = 12
    MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSHFlexGrid1.Text = strUnChecked

    TmpRs1.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
Label2.Caption = TmpRs1.RecordCount

End Sub

Private Sub ChkCheckAll_Click()

If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.Row = I
       MSHFlexGrid1.Col = 1
       If MSHFlexGrid1.CellBackColor <> vbRed And MSHFlexGrid1.CellBackColor <> vbGreen Then
          MSHFlexGrid1.Row = I
          MSHFlexGrid1.TextMatrix(I, 1) = strChecked
       End If
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.Row = I
       MSHFlexGrid1.TextMatrix(I, 1) = strUnChecked
   Next
End If

End Sub

Private Sub MSHFlexGrid1_Click()

With MSHFlexGrid1
     If .Col = 1 Then
        If .TextMatrix(.Row, .Col) = "" Then Exit Sub
           .Col = 1
        If (.CellBackColor) <> vbRed And (.CellBackColor) <> vbGreen Then
           If .TextMatrix(.Row, .Col) = strUnChecked Then
              .TextMatrix(.Row, .Col) = strChecked
           Else
              .TextMatrix(.Row, .Col) = strUnChecked
           End If
        End If
     End If
End With

End Sub
Private Sub ClearData()

ChkDP.Value = 0
CmbCNonNcdexClient.Text = ""
CmbLocationID.Text = ""
CmbCNonNcdexClient.Clear
Call Grid_Head
SaveCmd.Enabled = True
Label2.Caption = "_"
ProgressBar1.Value = 0
ChkCheckAll.Value = 0

End Sub
Private Sub SaveCmd_Click()

If CmbCNonNcdexClient.Text = "" Then
   MsgBox "Blank Client not allowed!!!"
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

SaveCmd.Enabled = False
ProgressBar1.Value = 0
Call TableTxn_TIDetails(" Where LocationID is Null ")
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For I = 1 To MSHFlexGrid1.Rows
    ProgressBar1.Value = ProgressBar1.Value + 1
    If MSHFlexGrid1.TextMatrix(I, 2) = "" Then Exit For
    If MSHFlexGrid1.TextMatrix(I, 1) = strChecked Then
       If RsTxn_TIDetails.RecordCount > 0 Then
          RsTxn_TIDetails.MoveFirst
       End If
       RsTxn_TIDetails.Find "EntryNo = " & Val(MSHFlexGrid1.TextMatrix(I, 4)) & "  "
       If RsTxn_TIDetails.EOF = False Then
          RsTxn_TIDetails!LocationID = mLocationID
          RsTxn_TIDetails!LocationUpdatedBy = Gen_UserLoginID
          RsTxn_TIDetails!LocationUpdatedDate = Now
          RsTxn_TIDetails.Update
          RsTxn_TIDetails.Requery
       End If
    End If
Next
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!!"
Call ClearData


End Sub
