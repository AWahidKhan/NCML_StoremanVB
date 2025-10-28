VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_PendingAdvanceCashReceipt 
   Caption         =   "Pending Advance Cash Receipt"
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
      Height          =   1725
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   15000
      Begin VB.ComboBox CmbClient 
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
         Left            =   150
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   1282
         Width           =   10575
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
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
         Left            =   11265
         TabIndex        =   21
         Top             =   1275
         Width           =   1215
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         Left            =   9720
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   3720
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show DP "
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
         Left            =   13680
         TabIndex        =   5
         Top             =   533
         Width           =   1140
      End
      Begin VB.ComboBox CmbStateID 
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
         Left            =   3340
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
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
         Left            =   12480
         TabIndex        =   6
         Top             =   1275
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
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
         Left            =   13695
         TabIndex        =   7
         Top             =   1275
         Width           =   1215
      End
      Begin VB.ComboBox CmbLocation 
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
         Left            =   6540
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   135
         TabIndex        =   0
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   107806721
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   1695
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   107806721
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Client Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4890
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label LblExchangeType 
         Caption         =   "Exchange Type"
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
         Left            =   10845
         TabIndex        =   20
         Top             =   195
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4320
         TabIndex        =   14
         Top             =   202
         Width           =   1065
      End
      Begin VB.Label Label5 
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
         Left            =   13680
         TabIndex        =   13
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Label4 
         Caption         =   "Total Record Found :-"
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
         Left            =   11280
         TabIndex        =   12
         Top             =   953
         Width           =   2325
      End
      Begin VB.Label Label3 
         Caption         =   "Location"
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
         Left            =   7620
         TabIndex        =   11
         Top             =   195
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "From Date"
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
         Left            =   345
         TabIndex        =   10
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "To Date"
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
         Left            =   2025
         TabIndex        =   9
         Top             =   195
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pending Advance Cash Receipt"
      TabPicture(0)   =   "FrmRpt_PendingAdvanceCashReceipt.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "TI Details"
      TabPicture(1)   =   "FrmRpt_PendingAdvanceCashReceipt.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   7275
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   6855
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12091
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
      End
      Begin VB.Frame Frame2 
         Height          =   7275
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   6855
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12091
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
      End
   End
End
Attribute VB_Name = "FrmRpt_PendingAdvanceCashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mStateID, mExchangeTypeID, mDepositoryID, mClientIDRow As Variant 'mClientIDRow, mDepositoryID
Dim mClientRNAVCode, mClientNAVCode, mDPNAVCode As Variant

Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
   CmbClient.Clear
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
   Check3.Enabled = False
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   Check3.Value = 0
   Check3.Enabled = False
Else
   Check3.Enabled = True
End If
End Sub

Private Sub CmbClient_GotFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please Select Exchage Type !!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
If Check3.Value = 0 Then
    tmp1 = CmbClient.Text
    'Call TableMst_Client("Where ExchangeTypeID = " & mExchangeTypeID)   ' 'And ClientGroupID=" & mClientGroupID
    tmp = "Select MC.ClientName,MC.ClientID,MD.DepositoryName,MC.ClientIDRow,MC.NAVCode,MC.RNAVCode From Mst_Client MC"
    tmp = tmp & " inner join Mst_Depository MD on MD.DepositoryID=MC.DepositoryID"
    tmp = tmp & " Where MC.ExchangeTypeId=" & mExchangeTypeID

    Call TmpTable

    CmbClient.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    'For I = 1 To RsMst_Client.RecordCount
    '    'CmbCNonNcdexClient.AddItem TmpRs!ClientName & "~" & TmpRs!ClientID & "~" & TmpRs!DepositoryName & "~" & TmpRs!ClientIDRow
    '    CmbClient.AddItem RsMst_Client!ClientName & "~" & TmpRs!ClientID & "~" & TmpRs!DepositoryName & "~" & TmpRs!ClientIDRow
    '    RsMst_Client.MoveNext
    'Next
    For I = 1 To TmpRs.RecordCount
        CmbClient.AddItem TmpRs!ClientName & "~" & TmpRs!ClientIDRow & "~" & TmpRs!DepositoryName & "~" & TmpRs!ClientID
        TmpRs.MoveNext
    Next
    CmbClient.Text = tmp1
ElseIf Check3.Value = 1 Then
    tmp = CmbClient.Text
    Call TableMst_Depository("Where ExchangeTypeID = " & mExchangeTypeID)
    CmbClient.Clear
    If RsMst_Depository.RecordCount = 0 Then
        MsgBox "No record found !!!! "
        Exit Sub
    End If
    For I = 1 To RsMst_Depository.RecordCount
        CmbClient.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
        RsMst_Depository.MoveNext
    Next
    CmbClient.Text = tmp
End If
End Sub

Private Sub CmbClient_LostFocus()

If CmbClient.Text = "" Then
   Exit Sub
End If

If Check3.Value = 0 Then
   strSplitString = Split(CmbClient, "~")
   If UBound(strSplitString) <> 3 Then
      MsgBox "Invalid Selection!!!"
      CmbClient.SetFocus
      Exit Sub
   End If
    
   ClientName_ = Trim(strSplitString(0))
   ClientNo_ = Trim(strSplitString(1))
   DPName_ = Trim(strSplitString(2))
   ClientIDRow_ = Trim(strSplitString(3))
   TmpRs.MoveFirst
   TmpRs.Find "ClientID = " & ClientIDRow_ & ""
   
   If TmpRs.EOF Then
       MsgBox "Invalid Selection "
       CmbClient.SetFocus
       Exit Sub
   End If
   mClientIDRow = TmpRs!ClientID
   mClientNAVCode = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
   mClientRNAVCode = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)

ElseIf Check3.Value = 1 Then
    
   strSplitString = Split(CmbClient, "~")
   'DPName_ = Trim(strSplitString(0))
   DPID_ = Trim(strSplitString(1))
    
   RsMst_Depository.MoveFirst
   RsMst_Depository.Find "DepositoryID = '" & DPID_ & "'"
   If RsMst_Depository.EOF Then
      MsgBox "Invalid selection "
      CmbClient.SetFocus
      Exit Sub
   End If
   mDepositoryID = RsMst_Depository!DepositoryID
   mDPNAVCode = IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)
End If
End Sub


Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" where " & Gen_WcState & "")
Else
   Call TableMst_State
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbLocation.Clear
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.Clear
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetReport.Enabled = False
   CmdExcel.Enabled = False
   Exit Sub
End If

Call CmdClear_Click
'Call Grid_Head
'CmdExcel.Enabled = False
End Sub
Private Sub CmdClear_Click()
TxtFromDate = Date - 30
TxtToDate = Date
CmbExchangeTypeID.Text = ""
CmbLocation.Text = ""
CmbClient.Text = ""
CmbStateID.Text = ""
Call Grid_Head
Call Grid_Head2
Label5.Caption = ""
CmdExcel.Enabled = False
End Sub
Private Sub CmdGetReport_Click()

Dim wc As Variant
wc = ""
Call Grid_Head
Call Grid_Head2

If Gen_DBType = "MDB" Then
    wc = wc & " And (CR.CRDate between #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#)"
Else
    wc = wc & " And CR.CRDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' and '" & Format(TxtToDate.Value, Gen_DatFormat) & "'" '''   tmp = tmp & " And (CDTFDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'  And '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
End If

If CmbStateID.Text <> "" Then
   wc = wc & " And MS.StateID = " & mStateID
Else
   If Gen_WcState <> "" Then
      wc = wc & " And MS." & Gen_WcState
   End If
End If

If CmbLocation.Text <> "" Then
   wc = wc & " And ML.LocationID = " & mLocationID
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And ML." & Gen_WcLocation
   End If
End If

'If CmbClient.Text <> "" Then
'   If Check3.Value = 0 Then
'        wc = wc & " And MC.ClientName = '" & CmbClient.Text & "'"
'   ElseIf Check3.Value = 1 Then
'        wc = wc & " And MD.DepositoryName = '" & CmbClient.Text & "'"
'   End If
'End If
If CmbExchangeTypeID.Text <> "" Then
    'wc = wc & " And (MC.ExchangeTypeID = " & mExchangeTypeID & " Or MD.ExchangeTypeID = " & mExchangeTypeID & " )"
    If CmbClient.Text <> "" Then
        If Check3.Value = 0 Then
           If mClientNAVCode <> "" Then
              wc = wc & " And (MC.NAVCode = '" & mClientNAVCode & "' ) " 'or MC.RNAVCode = '" & mClientRNAVCode & "')"
           End If
        ElseIf Check3.Value = 1 Then
           If mDPNAVCode <> "" Then
              wc = wc & " And MD.NAVCode='" & mDPNAVCode & "'"
           End If
        End If
    Else
        wc = wc & " And (MC.ExchangeTypeID = " & mExchangeTypeID & " Or MD.ExchangeTypeID = " & mExchangeTypeID & " )"
    End If
End If

tmp = " Select"
tmp = tmp & " CR.NewCRID, CR.CRDate, CR.SM_Prefix, MS.StateName, ML.LocationName, TCRCD.AdvNAVCode NAVCode1, MC.ClientName ,MC.NAVCode NAVCode2,"
tmp = tmp & " MD.DepositoryName, MD.NAVCode NAVCode3,"
'tmp = tmp & " TCRCD.Amount - (select isnull(sum(RecAmount),0)+isnull(sum(RecTDS),0)+isnull(sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail Where CRID= CR.CRID And SM_Prefix = CR.SM_Prefix And (ClientIDrow = TCRCD.ClientIDRow Or DPID = TCRCD.DPID) And CRAAID>0) as AdvanceAmount   "
tmp = tmp & " TCRCD.Amount  - (select isnull(sum(a.RecAmount),0)+isnull(sum(a.RecTDS),0)+isnull(sum(a.RecOtherDed),0) From Txn_NewCashReceiptTIDetail a inner join txn_tidetails b on b.tiid = a.tiid Where NewCRID= CR.NewCRID And SM_Prefix = CR.SM_Prefix And  b.NAVCode = TCRCD.AdvNAVCode  And CRAAID>0 ) as AdvanceAmount,TCRCD.DocNo,TCRCD.EntryNo"
tmp = tmp & " From Txn_NewCashReceipt CR"
tmp = tmp & " Inner Join Mst_Location ML on CR.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Txn_NewCashReceiptClientDetail TCRCD on CR.NewCRID = TCRCD.NewCRID And CR.SM_Prefix = TCRCD.SM_Prefix"
tmp = tmp & " Left Join Mst_Client MC on TCRCD.ClientID = MC.ClientID"
tmp = tmp & " Left Join Mst_Depository MD on TCRCD.DPID = MD.DepositoryID"
tmp = tmp & " where CR.AdvanceFlag = 1 And TCRCD.Amount > 0 And "
'tmp = tmp & " (TCRCD.Amount  - (select isnull(sum(RecAmount),0)+isnull(sum(RecTDS),0)+isnull(sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail Where CRID= CR.CRID And SM_Prefix = CR.SM_Prefix And (ClientIDrow = TCRCD.ClientIDRow Or DPID = TCRCD.DPID) And CRAAID>0 )) >0"
tmp = tmp & " (TCRCD.Amount  - (select isnull(sum(a.RecAmount),0)+isnull(sum(a.RecTDS),0)+isnull(sum(a.RecOtherDed),0) From Txn_NewCashReceiptTIDetail a inner join txn_tidetails b on b.tiid = a.tiid Where NewCRID= CR.NewCRID And SM_Prefix = CR.SM_Prefix And  b.NAVCode = TCRCD.AdvNAVCode  And CRAAID>0 )) > 0"
'- (Select isnull(sum(AdjAmountFAG),0) from Txn_TIDetails Where TIID in  (Select TIID from Txn_NewCashReceiptTIDetail Where CRID= CR.CRID And SM_Prefix = CR.SM_Prefix  And ClientIDrow = TCRCD.ClientIDRow And CRAAID > 0))
tmp = tmp & " And (CR.RejectFlag is Null or CR.RejectFlag = 'N') And (isnull(CR.ReturnFlag,'') <> 'Y') "
'tmp1 = " Order by MS.StateName, ML.LocationName, MC.ClientName, MD.DepositoryName, CR.CRID, CR.CRDate "
tmp1 = " Order by MS.StateName, ML.LocationName, MC.ClientName, MD.DepositoryName, CR.NewCRID, CR.CRDate "
tmp = tmp & wc & tmp1

'tmp = " Select CR.CRID, CR.CRDate, MS.StateName, ML.LocationName, TCRCD.AdvNAVCode NAVCode1, MC.ClientName ,MC.NAVCode NAVCode2, MD.DepositoryName,"
'tmp = tmp & " MD.NAVCode NAVCode3, max(TCRCD.Amount) AdvanceAmount"
''tmp = tmp & " sum(CRTD.RecAmount) RecAmount, CR.InstrumentAmount, sum(TTI.TIAmount) TIAmount, "
'tmp = tmp & " From Txn_NewCashReceipt CR"
'tmp = tmp & " Inner Join Mst_Location ML on CR.LocationID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
'tmp = tmp & " Inner Join Txn_NewCashReceiptClientDetail TCRCD on CR.CRID = TCRCD.CRID And CR.SM_Prefix = TCRCD.SM_Prefix"
'tmp = tmp & " Left Join Txn_NewCashReceiptTIDetail CRTD on CR.CRID = CRTD.CRID And CR.SM_Prefix = CRTD.SM_Prefix"
'tmp = tmp & " Left Join Txn_TIDetails TTI on CRTD.TIID = TTI.TIID "
'tmp = tmp & " Left Join Mst_Client MC on TCRCD.ClientIDRow = MC.ClientIDRow"
'tmp = tmp & " Left Join Mst_Depository MD on TCRCD.DPID = MD.DepositoryID"

'tmp1 = " group by CR.CRID, CR.InstrumentAmount, CR.CRDate, MS.StateName, ML.LocationName, TCRCD.AdvNAVCode, MC.ClientName, MC.NAVCode, MD.DepositoryName, MD.NAVCode "
'tmp1 = tmp1 & " Order by MS.StateName, ML.LocationName, MC.ClientName, MD.DepositoryName, CR.CRID, CR.CRDate "
'tmp = tmp & wc & tmp1

Call TmpTable

Call dispdata

End Sub

Private Sub GetTIDetail(x_ As Variant)

tmp = " Select TINo, TIDate, TIAmount, TI.TIID, PostInvoiceNo, TI.EntryNo, TI.NAVCode, "
tmp = tmp & " TI.RecAmount+isnull(TI.AdjAmountFAG,0) RecAmount, TI.RecTDSAmount RecTDS,  TI.RecOtherDed"
tmp = tmp & " From Txn_TIDetails TI"
'tmp = tmp & " Left JOin Txn_NewCashReceiptTIDetail CRTI On CRTI.TIID = TI.TIID"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID = ML.LocationID"
tmp = tmp & " left Join Mst_State MS on ML.StateID = MS.StateID"
'tmp = tmp & " Left Join Mst_Client MC on TI.NAVCode =  MC.NAVCode "
'tmp = tmp & " Left Join Mst_Depository MD on TI.NAVCode =  MD.NAVCode "
tmp = tmp & " Where  TINo <> '0' "

tmp = tmp & " And ( MS.StateName is null or MS.StateName = '" & MSHFlexGrid1.TextMatrix(x_, 2) & "')"
tmp = tmp & " And ( ML.LocationName is null or ML.LocationName = '" & MSHFlexGrid1.TextMatrix(x_, 3) & "')"
tmp = tmp & " And TI.NAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 5) & "'"

'tmp = tmp & " And MC.Commodity = '" & MSHFlexGrid1.TextMatrix(x_, 3) & "'"
'tmp = tmp & " Group By MCl.ClientIDRow,MCl.ClientName"

'tmp = tmp & " group by TINo, TIDate, TIAmount, TI.TIID, PostInvoiceNo, TI.EntryNo, TI.NAVCode, TI.AdjAmountFAG"
tmp = tmp & " And TIAmount > TI.RecAmount + TI.RecTDSAmount + TI.RecOtherDed + isnull(TI.AdjAmountFAG,0)"

Call Tmp1Table
MSHFlexGrid1.TextMatrix(x_, 7) = TmpRs1.RecordCount
If TmpRs1.RecordCount > 0 Then
    k = MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs1.RecordCount
    For l = 1 To TmpRs1.RecordCount
         MSHFlexGrid2.TextMatrix(k, 0) = IIf(IsNull(TmpRs1!TINo), "", TmpRs1!TINo)
        MSHFlexGrid2.TextMatrix(k, 1) = IIf(IsNull(TmpRs1!TIDate), "", Format(TmpRs1!TIDate, "dd-MMM-yyyy"))
        MSHFlexGrid2.TextMatrix(k, 2) = IIf(IsNull(TmpRs1!PostInvoiceNo), "", TmpRs1!PostInvoiceNo)
        MSHFlexGrid2.TextMatrix(k, 3) = IIf(IsNull(TmpRs1!EntryNo), "", TmpRs1!EntryNo)
        
        MSHFlexGrid2.TextMatrix(k, 4) = MSHFlexGrid1.TextMatrix(x_, 4)
        
        MSHFlexGrid2.TextMatrix(k, 5) = IIf(IsNull(TmpRs1!NavCode), "", TmpRs1!NavCode)
        MSHFlexGrid2.TextMatrix(k, 6) = IIf(IsNull(TmpRs1!TIAmount), "", TmpRs1!TIAmount)
        MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(k, 6))
        MSHFlexGrid2.TextMatrix(k, 7) = IIf(IsNull(TmpRs1!RecAmount), "", TmpRs1!RecAmount)
        MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(k, 7))
        MSHFlexGrid2.TextMatrix(k, 8) = IIf(IsNull(TmpRs1!RecTDS), "", TmpRs1!RecTDS)
        MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(k, 8))
        MSHFlexGrid2.TextMatrix(k, 9) = IIf(IsNull(TmpRs1!RecOtherDed), "", TmpRs1!RecOtherDed)
        MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(MSHFlexGrid2.TextMatrix(k, 9))
        MSHFlexGrid1.TextMatrix(x_, 8) = Val(MSHFlexGrid1.TextMatrix(x_, 8)) + (Val(MSHFlexGrid2.TextMatrix(k, 6)) - Val(MSHFlexGrid2.TextMatrix(k, 7)) - Val(MSHFlexGrid2.TextMatrix(k, 8)) - Val(MSHFlexGrid2.TextMatrix(k, 9)))
        MSHFlexGrid2.TextMatrix(k, 6) = Format(MSHFlexGrid2.TextMatrix(k, 6), "#########0.00")
        MSHFlexGrid2.TextMatrix(k, 7) = Format(MSHFlexGrid2.TextMatrix(k, 7), "#########0.00")
        MSHFlexGrid2.TextMatrix(k, 8) = Format(MSHFlexGrid2.TextMatrix(k, 8), "#########0.00")
        MSHFlexGrid2.TextMatrix(k, 9) = Format(MSHFlexGrid2.TextMatrix(k, 9), "#########0.00")
        TmpRs1.MoveNext
        k = k + 1
    Next l
    'MSHFlexGrid1.TextMatrix(x_, 5) = IIf(IsNull(TmpRs1!BOEAmount), 0, TmpRs1!BOEAmount)
End If
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_PendingAdvanceCashReceipt.xls")
Call Xls_Detail_Rpt
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_PendingAdvanceCashReceipt.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Pending Advance Cash Receipt"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    If I = 1 Then
        oWS.Cells(mRow, C + 1).RowHeight = 30
        For C = 0 To MSHFlexGrid1.Cols
            oWS.Cells(mRow, C + 1).WrapText = True
        Next
        oWS.Cells(mRow, C + 1).WrapText = True
    End If
    For C = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
    Next
Next
Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "TI Details"
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    If I = 1 Then
        oWS.Cells(mRow, C + 1).RowHeight = 30
        For C = 0 To MSHFlexGrid1.Cols
            oWS.Cells(mRow, C + 1).WrapText = True
        Next
    End If
    For C = 0 To MSHFlexGrid2.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.AllowUserResizing = flexResizeBoth

MSHFlexGrid1.TextMatrix(1, 0) = "Cash Receipt No"
MSHFlexGrid1.TextMatrix(1, 1) = "Cash Receipt Date"
MSHFlexGrid1.TextMatrix(1, 2) = "State Name"
MSHFlexGrid1.TextMatrix(1, 3) = "Location Name"
'If Check3.Value = 0 Then
'    MSHFlexGrid1.TextMatrix(1, 4) = "Client Name"
'ElseIf Check3.Value = 1 Then
'    MSHFlexGrid1.TextMatrix(1, 4) = "DP Name"
'End If
MSHFlexGrid1.TextMatrix(1, 4) = "Client Name"
MSHFlexGrid1.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid1.TextMatrix(1, 6) = "Advance Amount"
MSHFlexGrid1.TextMatrix(1, 7) = "Total TI"
MSHFlexGrid1.TextMatrix(1, 8) = "TI Amount"
MSHFlexGrid1.TextMatrix(1, 9) = "Payment Doc No"
MSHFlexGrid1.TextMatrix(1, 10) = "Payment Entry No"

MSHFlexGrid1.ColWidth(0) = 1300
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1700
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 3500
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1400
'MSHFlexGrid1.ColWidth(8) = 1400
'MSHFlexGrid1.ColWidth(9) = 1400

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(1) = 600


End Sub

Private Sub Grid_Head2()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 10
MSHFlexGrid2.AllowUserResizing = flexResizeBoth

MSHFlexGrid2.TextMatrix(1, 0) = "TI No"
MSHFlexGrid2.TextMatrix(1, 1) = "TI Date"
MSHFlexGrid2.TextMatrix(1, 2) = "Post Invoice No"
MSHFlexGrid2.TextMatrix(1, 3) = "Entry No"
MSHFlexGrid2.TextMatrix(1, 4) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 5) = "NAVCode"
'MSHFlexGrid2.TextMatrix(1, 6) = "Instrument Amount"
'MSHFlexGrid2.TextMatrix(1, 7) = "CashReceipt Amount"
MSHFlexGrid2.TextMatrix(1, 6) = "TI Amount"
MSHFlexGrid2.TextMatrix(1, 7) = "Received Amount"
MSHFlexGrid2.TextMatrix(1, 8) = "Received TDS"
MSHFlexGrid2.TextMatrix(1, 9) = "Received Other Ded."

MSHFlexGrid2.ColWidth(0) = 1500
MSHFlexGrid2.ColWidth(1) = 1300
MSHFlexGrid2.ColWidth(2) = 2200
MSHFlexGrid2.ColWidth(3) = 1300
MSHFlexGrid2.ColWidth(4) = 2500
MSHFlexGrid2.ColWidth(5) = 1500
MSHFlexGrid2.ColWidth(6) = 1400
MSHFlexGrid2.ColWidth(7) = 1500
MSHFlexGrid2.ColWidth(8) = 1500
MSHFlexGrid2.ColWidth(9) = 1500
'MSHFlexGrid2.ColWidth(9) = 1400

MSHFlexGrid2.WordWrap = True
MSHFlexGrid2.RowHeight(1) = 600

End Sub

Public Sub dispdata()

If TmpRs.RecordCount > 0 Then
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!NewCRID), "", TmpRs!NewCRID)
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!CRDate), "", Format(TmpRs!CRDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       If IsNull(TmpRs!ClientName) = True Then
           MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!DepositoryName
       Else
           MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!ClientName
       End If
       If IsNull(TmpRs!NAVCode1) = True Then
          If IsNull(TmpRs!navcode2) = True Then
             MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!NAVCode3
          Else
             MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!navcode2
          End If
       Else
          MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!NAVCode1
       End If
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!AdvanceAmount), "", TmpRs!AdvanceAmount)
       MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!DocNo), "", TmpRs!DocNo)
       MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
       
       MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
       mStatus = 0
       If I > 2 Then
           For j = 2 To I - 1
                If MSHFlexGrid1.TextMatrix(j, 5) = MSHFlexGrid1.TextMatrix(I, 5) Then
                    MSHFlexGrid1.TextMatrix(I, 8) = MSHFlexGrid1.TextMatrix(j, 8)
                    MSHFlexGrid1.TextMatrix(I, 7) = MSHFlexGrid1.TextMatrix(j, 7)
                    mStatus = 1
                    Exit For
                End If
           Next j
       End If
       If mStatus = 0 Then
            Call GetTIDetail(I)
       End If
       MSHFlexGrid1.TextMatrix(I, 6) = Format(MSHFlexGrid1.TextMatrix(I, 6), "#########0.00")
       MSHFlexGrid1.TextMatrix(I, 8) = Format(MSHFlexGrid1.TextMatrix(I, 8), "#########0.00")
       TmpRs.MoveNext
   Next I
'   If TmpRs1.RecordCount > 0 Then
'       For I = 2 To TmpRs1.RecordCount + 1
'       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
'       MSHFlexGrid2.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!TINo), "", TmpRs1!TINo)
'       MSHFlexGrid2.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!TIDate), "", Format(TmpRs1!TIDate, "dd-MMM-yyyy"))
'       MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!PostInvoiceNo), "", TmpRs1!PostInvoiceNo)
'       MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!EntryNo), "", TmpRs1!EntryNo)
'       MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!NAVCode), "", TmpRs1!NAVCode)
'       MSHFlexGrid2.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!TIAmount), "", TmpRs1!TIAmount)
'       MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(I, 5))
'       MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!RecAmount), "", TmpRs1!RecAmount)
'       MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(I, 6))
'       MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!RecTDS), "", TmpRs1!RecTDS)
'       MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(I, 7))
'       MSHFlexGrid2.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!RecOtherDed), "", TmpRs1!RecOtherDed)
'       MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(I, 8))
'       TmpRs1.MoveNext
'   Next
'   End If
   CmdExcel.Enabled = True
   
Else
   Call Grid_Head
   Call Grid_Head2
End If

Label5.Refresh
Label5.Caption = TmpRs.RecordCount

End Sub


Private Sub CmbLocation_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State !!"
   CmbStateID.SetFocus
   Exit Sub
End If
tmp = CmbLocation.Text

If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where StateID= " & mStateID & " And " & Gen_WcLocation)
Else
   Call TableMst_Location(" Where StateID= " & mStateID & "")
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_LostFocus()

If CmbLocation.Text = "" Then
   mLocationID = 0
   TxtPreFix = ""
   Exit Sub
End If
' Code for clearing Godown combox in case of change
If CmbLocation.Text <> tmp Then
   TxtPreFix = ""
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub OptClient_Click()
CmbClient.Text = ""
End Sub

Private Sub OptDepository_Click()
CmbClient.Text = ""
End Sub
