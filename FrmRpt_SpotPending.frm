VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_SpotPending 
   Caption         =   "SPOT Pending Report"
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
      TabIndex        =   18
      Top             =   1260
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7575
         Left            =   120
         TabIndex        =   19
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
      Height          =   1245
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbType 
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
         ItemData        =   "FrmRpt_SpotPending.frx":0000
         Left            =   6480
         List            =   "FrmRpt_SpotPending.frx":000A
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox CmbCDC 
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
         ItemData        =   "FrmRpt_SpotPending.frx":002A
         Left            =   13320
         List            =   "FrmRpt_SpotPending.frx":0034
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox TxtSClientName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1440
         TabIndex        =   7
         Top             =   795
         Width           =   4260
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
         Left            =   13785
         TabIndex        =   10
         Top             =   780
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
         Left            =   11595
         TabIndex        =   8
         Top             =   780
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
         Left            =   12690
         TabIndex        =   9
         Top             =   780
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
         Left            =   75
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
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
         Left            =   2715
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   2880
      End
      Begin VB.ComboBox CmbCMPName 
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
         Left            =   5640
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox CmbSGodown 
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
         Left            =   8460
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1485
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
         Left            =   9975
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   3240
      End
      Begin VB.Label Label2 
         Caption         =   "Type"
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
         Left            =   5760
         TabIndex        =   23
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "CDC Generated"
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
         Left            =   13335
         TabIndex        =   21
         Top             =   105
         Width           =   1485
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1320
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
         Left            =   8520
         TabIndex        =   17
         Top             =   810
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
         Left            =   10920
         TabIndex        =   16
         Top             =   810
         Width           =   150
      End
      Begin VB.Label LblStateID 
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
         Left            =   825
         TabIndex        =   15
         Top             =   105
         Width           =   1065
      End
      Begin VB.Label LblLocationID 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3780
         TabIndex        =   14
         Top             =   105
         Width           =   765
      End
      Begin VB.Label lblCMP 
         AutoSize        =   -1  'True
         Caption         =   "CMP Name"
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
         Left            =   6510
         TabIndex        =   13
         Top             =   105
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Godown"
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
         Left            =   8782
         TabIndex        =   12
         Top             =   105
         Width           =   840
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
         Left            =   11055
         TabIndex        =   11
         Top             =   105
         Width           =   1320
      End
   End
End
Attribute VB_Name = "FrmRpt_SpotPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mSM_Prefix, mCMPID, mGodownID, mCommodityID As Variant


Private Sub CmbCDC_LostFocus()
CmbType.Enabled = False
CmbType.Text = ""
If CmbCDC.Text = "" Then Exit Sub

If LCase(CmbCDC.Text) <> "yes" And LCase(CmbCDC.Text) <> "no" Then
   MsgBox "Invalid selection!!!"
   CmbCDC.SetFocus
   Exit Sub
End If
If LCase(CmbCDC.Text) = "no" Then
   CmbType.Enabled = True
End If
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
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
   CmbSGodown.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

If tmp <> RsMst_State!StateName Then
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
   CmbSGodown.Text = ""
End If

mStateID = RsMst_State!StateID
End Sub
Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!"
   CmbCMPName.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" where LocationID=" & mLocationID & " ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "First Select Location!!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub

Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   CmbSGodown.Text = ""
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If

If tmp <> RsMst_CMP!CMPName Then
   CmbSGodown.Text = ""
End If

mCMPID = RsMst_CMP!CMPID

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
   CmbCMPName.Text = ""
   CmbSGodown.Text = ""
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

If tmp <> RsMst_Location!LocationName Then
   CmbCMPName.Text = ""
   CmbSGodown.Text = ""
End If

mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub
Private Sub CmbSGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Please Select CMP  !!! "
   CmbSGodown.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPID)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbSGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
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

Private Sub ClearData()

Call Grid_Head
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbSGodown.Text = ""
TxtSClientName = ""
CmbCommodityID.Text = ""
lblRecordCount.Caption = ""
CmbCDC.Text = ""
CmbType.Enabled = False
CmbType.Text = ""

End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.Cols = 14
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.AllowUserResizing = flexResizeColumns

MSHFlexGrid1.TextMatrix(0, 0) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 1) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Godown Address"
MSHFlexGrid1.TextMatrix(0, 3) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 5) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 6) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 7) = "Stack No"
MSHFlexGrid1.TextMatrix(0, 8) = "Lot No"
MSHFlexGrid1.TextMatrix(0, 9) = "CDTF Date"
MSHFlexGrid1.TextMatrix(0, 10) = "Deposite Bags"
MSHFlexGrid1.TextMatrix(0, 11) = "Gross Wt"
MSHFlexGrid1.TextMatrix(0, 12) = "CDTF No"
MSHFlexGrid1.TextMatrix(0, 13) = "Truck No"

MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 1600
MSHFlexGrid1.ColWidth(2) = 2400
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1200

MSHFlexGrid1.RowHeight(0) = 800
MSHFlexGrid1.Font.Size = 10

End Sub





Private Sub CmbType_LostFocus()
If CmbType.Text = "" Then Exit Sub

If LCase(CmbType.Text) <> "marketable" And LCase(CmbType.Text) <> "non marketable" Then
   MsgBox "Invalid selection!!!"
   CmbType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdClear_Click()
Call ClearData
End Sub

Private Sub CmdGetReport_Click()

Dim wc As String
Dim tmp1 As String

If LCase(CmbCDC.Text) = "no" Then
   If CmbType.Text = "" Then
      MsgBox " please select Type!!"
      CmbType.SetFocus
      Exit Sub
   End If
End If

lblRecordCount.Caption = ""
Call Grid_Head

MsgBox "Wait till next message!!!"

wc = " Where MGSL.ExchangeTypeID = 3  "


If CmbStateID.Text <> "" Then
   wc = GenWc(wc, CmbStateID.Text, "MS.StateName", "S", "N")
Else
   If Gen_WcState <> "" Then
      tmp2 = Replace(Gen_WcState, "StateID", "MS.StateID")
      wc = GenWc(wc, tmp2, "", "S", "I")
   End If
End If

If CmbLocationID.Text <> "" Then
   wc = GenWc(wc, CmbLocationID.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp2 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, tmp2, "", "S", "I")
   End If
End If

If CmbCMPName.Text <> "" Then
   wc = GenWc(wc, CmbCMPName.Text, "MC.CMPName", "S", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp2 = Replace(Gen_WcCMP, "CMPID", "MC.CMPID")
      wc = GenWc(wc, tmp2, "", "S", "I")
   End If
End If

If CmbSGodown.Text <> "" Then
   wc = wc & " And MG.GodownNo = '" & CmbSGodown.Text & "' "
End If

If CmbCommodityID.Text <> "" Then
   wc = GenWc(wc, CmbCommodityID.Text, "Commodity", "S", "N")
End If

If TxtSClientName.Text <> "" Then
   wc = GenWc(wc, TxtSClientName.Text, "MCl.ClientName", "S", "L")
End If

If LCase(CmbCDC.Text) = "yes" Then
   wc = wc & " And MGSL.ISINID <> '' And MGSL.ISINID is not null And MGSL.ISINID <> 0 "
ElseIf LCase(CmbCDC.Text) = "no" Then
   wc = wc & " And (MGSL.ISINID = '' Or MGSL.ISINID is null Or MGSL.ISINID = 0) "
End If

If LCase(CmbCDC.Text) = "no" Then
   If LCase(CmbType.Text) = "marketable" Then
      wc = wc & "  AND (MGSL.DepositeWeight>=CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
   ElseIf LCase(CmbType.Text) = "non marketable" Then
      wc = wc & "  AND (MGSL.DepositeWeight<CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
   End If
End If

tmp = " Select MS.StateName,ML.LocationName,MC.CMPName,MG.Address,MCl.ClientName,MCl.ClientID,MCom.Commodity,"
tmp = tmp & " MG.GodownNo , MGSL.StackNo, MGSL.LotNo,MGSL.CDTFDate, MGSL.DepositeBags, MGSL.GrossWt,Min(TCG.CxTFNo)CxTFNo,Min(TCG.TruckNo)TruckNo"
tmp = tmp & " From Mst_GSL MGSL"
tmp = tmp & " Inner Join Mst_Godown MG On MGSL.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MGSL.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_Client MCl On MGSL.ClientIDRow = MCl.ClientIDRow"
tmp = tmp & " Inner Join Mst_Commodity MCom On MGSL.CommodityID = MCom.CommodityID"
tmp = tmp & " Inner join Mst_CommodityDetail CMD on CMD.CommodityID=MGSL.CommodityID And CMD.ExchangeTypeID=MGSL.ExchangeTypeID"
tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On MGSL.GSLID = TCG.GSLID And MGSL.SM_Prefix = TCG.SM_Prefix And TCG.TxnType='D'"

'tmp = tmp & " Where MGSL.ExchangeTypeID = 3 And (ISINID = '' Or ISINID is null Or ISINID = 0)"
tmp1 = " Group by  MS.StateName,ML.LocationName,MC.CMPName,MG.Address,MCl.ClientName,MCl.ClientID,MCom.Commodity,"
tmp1 = tmp1 & " MG.GodownNo , MGSL.StackNo,MGSL.CDTFDate, MGSL.LotNo,MGSL.DepositeBags, MGSL.GrossWt"
tmp1 = tmp1 & " Order By ML.LocationName,MC.CMPName,MCl.ClientName "

tmp = tmp & " " & wc & " " & tmp1

Call TmpTable
   
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For i = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(i, 0) = TmpRs!LocationName
       MSHFlexGrid1.TextMatrix(i, 1) = TmpRs!CMPName
       MSHFlexGrid1.TextMatrix(i, 2) = TmpRs!Address
       MSHFlexGrid1.TextMatrix(i, 3) = TmpRs!ClientName
       MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
       MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
       MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo)
       MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(TmpRs!LotNo), "", TmpRs!LotNo)
       MSHFlexGrid1.TextMatrix(i, 9) = IIf(IsNull(TmpRs!CDTFDate), "", Format(TmpRs!CDTFDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(i, 10) = IIf(IsNull(TmpRs!DepositeBags), "", TmpRs!DepositeBags)
       MSHFlexGrid1.TextMatrix(i, 11) = IIf(IsNull(TmpRs!GrossWt), "", TmpRs!GrossWt)
       MSHFlexGrid1.TextMatrix(i, 12) = IIf(IsNull(TmpRs!CxTFNo), "", TmpRs!CxTFNo)
       MSHFlexGrid1.TextMatrix(i, 13) = IIf(IsNull(TmpRs!TruckNO), "", TmpRs!TruckNO)

       TmpRs.MoveNext
   Next
End If

lblRecordCount.Caption = TmpRs.RecordCount
MsgBox "Done!!!"

End Sub

Private Sub Form_Load()

Call Grid_Head
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call CmdClear_Click

End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_SpotPending.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_SpotPending.xls")
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
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

