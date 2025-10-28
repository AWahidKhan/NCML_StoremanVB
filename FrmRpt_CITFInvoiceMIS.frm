VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CITFInvoiceMIS 
   Caption         =   "CITF Invoice MIS"
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
      Height          =   7935
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7575
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13361
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1590
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbNAVCode 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   1125
         Width           =   2760
      End
      Begin VB.ComboBox CmbClientID 
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
         Left            =   8280
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   6255
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
         Left            =   12615
         TabIndex        =   7
         Top             =   1110
         Width           =   1095
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Go"
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
         Left            =   11520
         TabIndex        =   6
         Top             =   1110
         Width           =   1095
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
         Left            =   13710
         TabIndex        =   8
         Top             =   1110
         Width           =   1095
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
         Left            =   2952
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   2595
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
         Left            =   5598
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   2640
      End
      Begin VB.ComboBox CmbCommodityID 
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
         Left            =   2952
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1125
         Width           =   3480
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   75563009
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   1551
         TabIndex        =   1
         Top             =   420
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   75563009
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label35 
         Caption         =   "NAV Code"
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
         Left            =   1013
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.Label LblClient 
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
         Height          =   300
         Left            =   10807
         TabIndex        =   19
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lblRec 
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
         Left            =   7680
         TabIndex        =   16
         Top             =   1170
         Width           =   2325
      End
      Begin VB.Label lblRecordCount 
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
         Left            =   10080
         TabIndex        =   15
         Top             =   1170
         Width           =   210
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   3634
         TabIndex        =   14
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label LblLocationID 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   6468
         TabIndex        =   13
         Top             =   120
         Width           =   900
      End
      Begin VB.Label LblCommodityID 
         Caption         =   "Commodity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4152
         TabIndex        =   12
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "To  Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1791
         TabIndex        =   11
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "From  Date"
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
         Left            =   278
         TabIndex        =   10
         Top             =   120
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmRpt_CITFInvoiceMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mSM_Prefix, mCMPID, mGodownID, mCommodityID, mExchangeTypeID As Variant
Dim mClientNAVCode, mClientRNAVCode As Variant

Private Sub CmbNAVCode_LostFocus()
CmbNAVCode = Replace(CmbNAVCode, Chr(32), "")
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text
CmbLocationID.Clear

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection!!! "
   CmbCommodityID.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID

End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 3
     .FixedRows = 2
     .Cols = 13
     .Font.Size = 10
     .AllowUserResizing = flexResizeBoth
     
     '.MergeCells = flexMergeRestrictRows
     
     .TextMatrix(1, 0) = "State"
     .TextMatrix(1, 1) = "Location"
     .TextMatrix(1, 2) = "Client Name"
     .TextMatrix(1, 3) = "Commodity"
     .TextMatrix(1, 4) = "Qty withdrawn during the given period (in MT)"
     .TextMatrix(1, 5) = "BoE amount of the withdrawn quantity"
     .TextMatrix(1, 6) = "Amount recovered during given period"
     .TextMatrix(1, 7) = "Balance to be recovered for the withdrawn quantity"
     .TextMatrix(1, 8) = "Total outstanding invoice amount of this client in this location (as per store-man)"
     .TextMatrix(1, 9) = "Total value of all commodities of this client  stored in this location"
     .TextMatrix(1, 10) = "Invoice Risk"
     .TextMatrix(1, 11) = "Client NAV Code"
     .TextMatrix(1, 12) = "RClient NAV Code"
     
     
     .ColWidth(0) = 1800
     .ColWidth(1) = 1800
     .ColWidth(2) = 1600
     .ColWidth(3) = 1600
     .ColWidth(4) = 1600
     .ColWidth(5) = 1600
     .ColWidth(6) = 1800
     .ColWidth(7) = 2200
     .ColWidth(8) = 2600
     .ColWidth(9) = 2600
     .ColWidth(10) = 2600
     .ColWidth(11) = 0
     .ColWidth(12) = 0
     
     .WordWrap = True
     .RowHeight(1) = 1200
End With
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbClientID.Text = ""
CmbCommodityID.Text = ""
lblRecordCount.Caption = "-"
End Sub

Private Sub CmbClientId_GotFocus()
tmp = CmbClientID.Text
Call TableMst_Client("Where ExchangeTypeID = 1") 'And ClientGroupID=" & mClientGroupID
CmbClientID.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
    CmbClientID.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClientID.Text = tmp
End Sub
'----
Private Sub CmbClientId_LostFocus()
If CmbClientID.Text = "" Then
   CmbNAVCode.Text = ""
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClientID.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbClientID.SetFocus
   Exit Sub
End If
CmbNAVCode.Text = ""
mClientIDRow = RsMst_Client!ClientIDRow
mClientNAVCode = IIf(IsNull(RsMst_Client!NAVCode), "", RsMst_Client!NAVCode)
mClientRNAVCode = IIf(IsNull(RsMst_Client!RNAVCode), "", RsMst_Client!RNAVCode)
CmbNAVCode.Clear
If mClientRNAVCode <> "" And mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.AddItem mClientRNAVCode
ElseIf mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.Text = mClientNAVCode
ElseIf mClientRNAVCode <> "" Then
   CmbNAVCode.AddItem mClientRNAVCode
   CmbNAVCode.Text = mClientRNAVCode
End If

End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant

''On Error GoTo ErrorMsg


Call Grid_Head
lblRecordCount.Caption = ""

MsgBox "Wait till next message!!!"

wc = ""

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "'From Date' must be less than 'To Date' always!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If TxtFromDate.Value > Date Then
   MsgBox "Future date not allowed!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If
'
Dim ans As Variant
If CmbStateID.Text = "" Then
   If LCase(Gen_UserRole) <> "power" Then
      MsgBox "Blank State Name not allowed!!!"
      CmbStateID.SetFocus
      Exit Sub
   Else
      ans = MsgBox("Do you want to run report for all states??", vbYesNo)
      If ans = vbNo Then
         CmbStateID.SetFocus
         Exit Sub
      End If
   End If
End If

If CmbLocationID.Text = "" Then
   If LCase(Gen_UserRole) <> "power" Then
      MsgBox "Blank Location Name not allowed!!!"
      CmbLocationID.SetFocus
      Exit Sub
   Else
      ans = MsgBox("Do you want to run report for all location??", vbYesNo)
      If ans = vbNo Then
         CmbLocationID.SetFocus
         Exit Sub
      End If
   End If
End If


tmp = " Select MS.StateName,ML.LocationName,MCl.ClientName,MCL.NAVCode,MCL.RNAVCode,"
tmp = tmp & " MC.Commodity , Sum(TCG.DematWeight) 'QtyWithdrawn' " '-- --Sum(BOEAmount) 'BOEAmount'
tmp = tmp & " from Txn_CxTF_GSL TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On  TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Inner Join Mst_Location ML On TCG.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
'tmp = tmp & " Inner Join Mst_GSL MG On TCG.SM_Prefix = MG.SM_Prefix And TCG.GSLID = MG.GSLID And MG.ExchangeTypeID = 1"
tmp = tmp & " Inner Join Mst_Client MCl On TCG.ClientIDRow = MCl.ClientIDRow"
tmp = tmp & " Inner Join Mst_Commodity MC On TCD.CommodityID = MC.CommodityID"
'tmp = tmp & " Left Join  Txn_CxTF_BOE TCB On TCD.CxTFNo = TCB.CxTFNo And TCD.SM_Prefix = TCB.SM_Prefix"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TCD.CxTFDate Between #" & TxtFromDate.Value & "# And #" & TxtToDate.Value & "# And TCD.ExchangeTypeID = 1 And TCG.TxnType = 'W' "
Else
   tmp = tmp & " Where TCD.CxTFDate Between '" & TxtFromDate.Value & "' And '" & TxtToDate.Value & "' And TCD.ExchangeTypeID = 1 And TCG.TxnType = 'W' "
End If
tmp1 = " Group By MS.StateName,ML.LocationName,MC.Commodity,MCl.ClientName,MCL.NAVCode,MCL.RNAVCode"

If CmbStateID.Text <> "" Then
   wc = wc & " And MS.StateName = '" & CmbStateID.Text & "'"
End If

If CmbLocationID.Text <> "" Then
   wc = wc & " And ML.LocationName = '" & CmbLocationID.Text & "'"
End If

If CmbClientID.Text <> "" Then
   wc = wc & " And MCl.ClientName = '" & CmbClientID.Text & "'"
End If

If CmbNAVCode.Text <> "" Then
   wc = wc & " And (MCL.NAVCode = '" & mClientNAVCode & "' Or MCL.RNAVCode = '" & mClientRNAVCode & "') "
End If

If CmbCommodityID.Text <> "" Then
   wc = wc & " And MC.Commodity = '" & CmbCommodityID.Text & "'"
End If

tmp = tmp & wc & tmp1

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
X = 2
For i = 1 To TmpRs.RecordCount
'    For c = 0 To 5
'        MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
'    Next
'    If MSHFlexGrid1.TextMatrix(I, 5) = "" Then
'       MSHFlexGrid1.TextMatrix(I, 5) = 0
'    End If
    MSHFlexGrid1.TextMatrix(X, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
    MSHFlexGrid1.TextMatrix(X, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
    MSHFlexGrid1.TextMatrix(X, 2) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
    MSHFlexGrid1.TextMatrix(X, 3) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
    MSHFlexGrid1.TextMatrix(X, 4) = IIf(IsNull(TmpRs!QtyWithdrawn), 0, TmpRs!QtyWithdrawn)
    GetBOEAmount (X)
    'MSHFlexGrid1.TextMatrix(x, 5) = IIf(IsNull(TmpRs!BOEAmount), 0, TmpRs!BOEAmount)
    MSHFlexGrid1.TextMatrix(X, 10) = IIf(IsNull(TmpRs!NAVCode), "", TmpRs!NAVCode)
    MSHFlexGrid1.TextMatrix(X, 11) = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)
    GetAmountRecoveredDetails (X)
    GetTotalOutStandingAmount (X)
    GetTotalValueOfCommodity (X)
    
    MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(X, 4))
    MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(X, 5))
    MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(X, 6))
    MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(X, 7))
    MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(X, 8))
    MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + Val(MSHFlexGrid1.TextMatrix(X, 9))
    
    TmpRs.MoveNext
    X = X + 1
Next

lblRecordCount.Caption = TmpRs.RecordCount
MsgBox "Done!!!"

Exit Sub

ErrorMsg:
     
   MsgBox "Data fetching is taking time from database. Please try after some time!!!"
   Exit Sub

End Sub
Private Sub GetBOEAmount(x_ As Variant)

tmp = " Select MCl.ClientIDRow,MCl.ClientName,Sum(BOEAmount) 'BOEAmount'"
tmp = tmp & " From Txn_CxTF_BOE TCB"
tmp = tmp & " Left Join  Txn_CxTF_GSL TCG  On TCG.CxTFNo = TCB.CxTFNo And TCG.SM_Prefix = TCB.SM_Prefix"
tmp = tmp & " Left Join Mst_Location ML On TCB.LocationID = ML.LOcationID"
tmp = tmp & " Left Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Left Join Mst_Client MCl On TCB.ClientIDRow = MCl.ClientIDRow"
tmp = tmp & " Left Join Mst_Commodity MC On TCB.CommodityID = MC.CommodityID"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TCB.CxTFDate Between #" & TxtFromDate.Value & "# And #" & TxtToDate.Value & "# And TCB.ExchangeTypeID = 1 And TCG.TxnType = 'W'"
Else
   tmp = tmp & " Where TCB.CxTFDate Between '" & TxtFromDate.Value & "' And '" & TxtToDate.Value & "' And TCB.ExchangeTypeID = 1 And TCG.TxnType = 'W'"
End If
tmp = tmp & " And MS.StateName = '" & MSHFlexGrid1.TextMatrix(x_, 0) & "'"
tmp = tmp & " And ML.LocationName = '" & MSHFlexGrid1.TextMatrix(x_, 1) & "'"
tmp = tmp & " And MCl.ClientName = '" & MSHFlexGrid1.TextMatrix(x_, 2) & "'"
tmp = tmp & " And MC.Commodity = '" & MSHFlexGrid1.TextMatrix(x_, 3) & "'"
tmp = tmp & " Group By MCl.ClientIDRow,MCl.ClientName"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x_, 5) = IIf(IsNull(TmpRs1!BOEAmount), 0, TmpRs1!BOEAmount)
Else
   MSHFlexGrid1.TextMatrix(x_, 5) = 0
End If
End Sub

Private Sub GetAmountRecoveredDetails(x_ As Variant)

tmp = " Select Sum(TI.RecAmount)+ Sum(isnull(TTI.AdjAmountFAG,0)) 'RecAmt ',Sum(TI.RecTDS) 'RecTDS',Sum(TI.RecOtherDed) 'RecOthr',Sum(TI.RecAmount)+Sum(isnull(TTI.AdjAmountFAG,0)) + Sum(TI.RecTDS) + Sum(TI.RecOtherDed) as 'TotalRec',MC.ClientName"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TI"
tmp = tmp & " inner Join  Txn_TIDetails TTI On TI.TIID=TTI.TIID"
tmp = tmp & " Inner Join Txn_NewCashReceipt TNCR On TI.CRID = TNCR.CRID And TI.SM_Prefix = TNCR.SM_Prefix"
tmp = tmp & " inner Join Mst_Location ML On TNCR.SM_PRefix = ML.SM_PRefix "
tmp = tmp & " Left Join Mst_Client MC On TI.ClientIDRow=MC.ClientIDRow"
'--tmp = tmp & " Left Join Mst_Depository MD On TI.DPID=MD.DepositoryID"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Where TNCR.CRDate >=  #" & TxtFromDate.Value & "# And TNCR.CRDate <=  #" & TxtToDate.Value & "# "
'Else
'   tmp = tmp & " Where TNCR.CRDate >=  '" & TxtFromDate.Value & "' And TNCR.CRDate <=  '" & TxtToDate.Value & "' "
'End If

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TNCR.CRDate <=  #" & TxtToDate.Value & "# "
Else
   tmp = tmp & " Where TNCR.CRDate <=  '" & TxtToDate.Value & "' "
End If

tmp = tmp & " And MC.ClientName = '" & MSHFlexGrid1.TextMatrix(x_, 2) & "'"
tmp = tmp & " And ML.LocationName = '" & MSHFlexGrid1.TextMatrix(x_, 1) & "'"
tmp = tmp & " Group By MC.ClientName "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x_, 6) = TmpRs1!TotalRec
   MSHFlexGrid1.TextMatrix(x_, 7) = Val(MSHFlexGrid1.TextMatrix(x_, 5)) - Val(MSHFlexGrid1.TextMatrix(x_, 6))

Else
   MSHFlexGrid1.TextMatrix(x_, 6) = 0
   MSHFlexGrid1.TextMatrix(x_, 7) = Val(MSHFlexGrid1.TextMatrix(x_, 5)) - Val(MSHFlexGrid1.TextMatrix(x_, 6))
End If
If Val(MSHFlexGrid1.TextMatrix(x_, 7)) < 0 Then
   MSHFlexGrid1.TextMatrix(x_, 7) = 0
End If

MSHFlexGrid1.TextMatrix(x_, 7) = Format(MSHFlexGrid1.TextMatrix(x_, 7), "#######0.00")


End Sub
Private Sub GetTotalOutStandingAmount(x_ As Variant)
Dim mTIAmount, mRecAmt, mTDSAmt, mOthrAmt, mBalAmt As Variant
mTIAmount = 0
mRecAmt = 0
mTDSAmt = 0
mOthrAmt = 0
mBalAmt = 0

'''tmp = " Select Sum(TIAmount) 'TIAmount',(Select Sum(RecAmount) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecAmount,"
'''tmp = tmp & " (Select Sum(RecTDS) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecTDSAmount ,"
'''tmp = tmp & " (Select Sum(RecOtherDed) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecOtherDed ,"

tmp = " Select Sum(TIAmount) 'TIAmount',"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ sum(isnull(TI.AdjAmountFAG,0)) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDSAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecOtherDed,"
tmp = tmp & " TI.NAVCode , ML.LocationName"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Where  TINo <> '' And (ML.LocationName = '" & MSHFlexGrid1.TextMatrix(x_, 1) & "' ) "

If CmbNAVCode.Text <> "" Then
   tmp = tmp & " And (TI.NAVCode = '" & CmbNAVCode.Text & "')"  ' Or CustomerNo = '" & mClientRNAVCode & "') "
Else
   tmp = tmp & " And (TI.NAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 10) & "' Or TI.NAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 11) & "') "
End If

tmp = tmp & " Group By TI.TIID,TI.NAVCode,ML.LocationName"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For y = 1 To TmpRs1.RecordCount
       mTIAmount = Val(mTIAmount) + TmpRs1!TIAmount
       If IsNull(TmpRs1!RecAmount) = False Then
          mRecAmt = Val(mRecAmt) + TmpRs1!RecAmount
       End If
       If IsNull(TmpRs1!RecTDSAmount) = False Then
          mTDSAmt = Val(mTDSAmt) + TmpRs1!RecTDSAmount
       End If
       If IsNull(TmpRs1!RecOtherDed) = False Then
          mOthrAmt = Val(mOthrAmt) + TmpRs1!RecOtherDed
       End If
       TmpRs1.MoveNext
   Next
   mBalAmt = Val(mTIAmount) - (Val(mRecAmt) + Val(mTDSAmt) + Val(mOthrAmt))
   MSHFlexGrid1.TextMatrix(x_, 8) = Val(mBalAmt)
   MSHFlexGrid1.TextMatrix(x_, 8) = Format(MSHFlexGrid1.TextMatrix(x_, 8), "#######0.00")
Else
   MSHFlexGrid1.TextMatrix(x_, 8) = 0
End If

End Sub
Private Sub GetTotalValueOfCommodity(x_ As Variant)

'---- For Balance
'--Deposit
tmp = " Select sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid"
tmp = tmp & " inner Join mst_location l on c.sm_prefix = l.sm_prefix "
tmp = tmp & " inner Join mst_client cl On c.clientIDRow = cl.clientidrow"
tmp = tmp & " inner Join mst_Commodity mc On c.commodityID = mc.commodityid "

If Gen_DBType = "MDB" Then
   tmp = tmp & " where a.exchangetypeid = 1 And a.TxnType='D' And a.CxTFDate <= #" & TxtToDate.Value & "#"
Else
   tmp = tmp & " where a.exchangetypeid = 1 And a.TxnType='D' And a.CxTFDate <= '" & TxtToDate.Value & "'"
End If

tmp = tmp & " And cl.ClientName = '" & MSHFlexGrid1.TextMatrix(x_, 2) & "' And l.LocationName= '" & MSHFlexGrid1.TextMatrix(x_, 1) & "' "
tmp = tmp & " And mc.commodity = '" & MSHFlexGrid1.TextMatrix(x_, 3) & "' "

Call Tmp1Table
mOpDepo = 0
If TmpRs1.RecordCount > 0 Then
   mOpDepo = TmpRs1!GrossWeight
End If


'--Withdrawal
tmp = " Select sum(b.NoOfBags) as NoOfBags ,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid"
tmp = tmp & " inner Join mst_location l on c.sm_prefix = l.sm_prefix "
tmp = tmp & " inner Join mst_client cl On c.clientIDRow = cl.clientidrow"
tmp = tmp & " inner Join mst_Commodity mc On c.commodityID = mc.commodityid "
'tmp = tmp & " where a.exchangetypeid = 1 And a.TxnType='W' And a.CxTFDate <'2011/11/09'"
'tmp = tmp & " And a.SM_Prefix='ABHR' And c.ClientIDRow in (2791,7580,7579,5431)"

If Gen_DBType = "MDB" Then
   tmp = tmp & " where a.exchangetypeid = 1 And a.TxnType='W' And a.CxTFDate <= #" & TxtToDate.Value & "#"
Else
   tmp = tmp & " where a.exchangetypeid = 1 And a.TxnType='W' And a.CxTFDate <= '" & TxtToDate.Value & "'"
End If

tmp = tmp & " And cl.ClientName = '" & MSHFlexGrid1.TextMatrix(x_, 2) & "' And l.LocationName= '" & MSHFlexGrid1.TextMatrix(x_, 1) & "' "
tmp = tmp & " And mc.commodity = '" & MSHFlexGrid1.TextMatrix(x_, 3) & "' "

Call Tmp2Table

mOpWith = 0

If TmpRs2.RecordCount > 0 Then
   mOpWith = TmpRs2!GrossWeight
End If

mOpBal = Val(mOpDepo) - Val(mOpWith)

'MSHFlexGrid1.TextMatrix(x_, 9) = mOpBal

'tmp = " Select MS.StateName,MC.Commodity,MCP.MonthYear,MCP.Day28,MCP.Day29,MCP.Day30,MCP.Day31 from Mst_CommodityPrice MCP"
'tmp = tmp & " Inner Join Mst_Commodity MC on MCP.CommodityID = MC.CommodityID"
'tmp = tmp & " Inner Join Mst_State MS On MCP.StateID = MS.StateID"
'tmp = tmp & " Where MonthYear =  '" & Left(MonthName(Month((TxtToDate.Value))), 3) & "-" & Year(TxtToDate.Value) & "'"
'tmp = tmp & " And MS.StateName = '" & MSHFlexGrid1.TextMatrix(x_, 0) & "' And MC.Commodity = '" & MSHFlexGrid1.TextMatrix(x_, 3) & "' "

tmp = " Select MS.StateName,MC.Commodity,MCP.MonthYear,MCP.YearMonth,MCP.Day28,MCP.Day29,MCP.Day30,MCP.Day31 from Mst_CommodityPrice MCP"
tmp = tmp & " Inner Join Mst_Commodity MC on MCP.CommodityID = MC.CommodityID"
tmp = tmp & " Inner Join Mst_State MS On MCP.StateID = MS.StateID"
'tmp = tmp & " Where YearMonth =  '" & Left(MonthName(Month((TxtToDate.Value))), 3) & "-" & Year(TxtToDate.Value) & "'"
tmp = tmp & " And MS.StateName = '" & MSHFlexGrid1.TextMatrix(x_, 0) & "' And MC.Commodity = '" & MSHFlexGrid1.TextMatrix(x_, 3) & "' "
tmp = tmp & " Order By MCP.YearMonth "

Call Tmp1Table

mRate = 0

If TmpRs1.RecordCount > 0 Then
   TmpRs1.MoveLast
   'If Month(TxtToDate.Value) = 1 Or Month(TxtToDate.Value) = 3 Or Month(TxtToDate.Value) = 5 Or Month(TxtToDate.Value) = 7 Or Month(TxtToDate.Value) = 8 Or Month(TxtToDate.Value) = 10 Or Month(TxtToDate.Value) = 12 Then
   If Right(TmpRs1!YearMonth, 2) = "01" Or Right(TmpRs1!YearMonth, 2) = "03" Or Right(TmpRs1!YearMonth, 2) = "05" Or Right(TmpRs1!YearMonth, 2) = "07" Or Right(TmpRs1!YearMonth, 2) = "08" Or Right(TmpRs1!YearMonth, 2) = "10" Or Right(TmpRs1!YearMonth, 2) = "12" Then
      mRate = TmpRs1!Day31
   ElseIf Right(TmpRs1!YearMonth, 2) = "04" Or Right(TmpRs1!YearMonth, 2) = "06" Or Right(TmpRs1!YearMonth, 2) = "09" Or Right(TmpRs1!YearMonth, 2) = "11" Then
      mRate = TmpRs1!Day30
   Else
      If IsNull(TmpRs1!Day29) = False Then
         mRate = TmpRs1!Day29
      Else
         mRate = TmpRs1!Day28
      End If
   End If
Else
   tmp = " Select MinRate from Mst_Commodity Where Commodity = '" & MSHFlexGrid1.TextMatrix(x_, 3) & "'"

   Call Tmp1Table
   
   If TmpRs1.RecordCount > 0 Then
      mRate = IIf(IsNull(TmpRs1!MinRate), 0, TmpRs1!MinRate)
   End If
End If

MSHFlexGrid1.TextMatrix(x_, 9) = Val(mOpBal) * Val(mRate)
MSHFlexGrid1.TextMatrix(x_, 9) = Format(MSHFlexGrid1.TextMatrix(x_, 9), "#######0.00")
MSHFlexGrid1.TextMatrix(x_, 10) = Val(MSHFlexGrid1.TextMatrix(x_, 8)) - Val(MSHFlexGrid1.TextMatrix(x_, 9))

If Val(MSHFlexGrid1.TextMatrix(x_, 10)) < 0 Then
   MSHFlexGrid1.TextMatrix(x_, 10) = 0
End If

MSHFlexGrid1.TextMatrix(x_, 10) = Format(MSHFlexGrid1.TextMatrix(x_, 10), "#######0.00")

End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   lblRecordCount.Caption = ""
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
lblRecordCount.Caption = ""
Call Grid_Head
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CITFInvoiceMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CITFInvoiceMIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "mm/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

