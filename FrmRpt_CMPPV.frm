VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_CMPPV 
   Caption         =   "CMP PV"
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
      Height          =   7695
      Left            =   120
      TabIndex        =   27
      Top             =   2160
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   34
         Top             =   7485
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7185
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12674
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2085
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14985
      Begin VB.ComboBox CmbMadeupFlag 
         Height          =   315
         ItemData        =   "FrmRpt_CMPPV.frx":0000
         Left            =   10800
         List            =   "FrmRpt_CMPPV.frx":000A
         TabIndex        =   32
         Top             =   1200
         Width           =   1605
      End
      Begin VB.ComboBox CmbStateID 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   465
         Width           =   2205
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
         Height          =   360
         Left            =   13560
         TabIndex        =   15
         Top             =   1605
         Width           =   1215
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
         Height          =   360
         Left            =   11130
         TabIndex        =   13
         Top             =   1605
         Width           =   1215
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   465
         Width           =   3165
      End
      Begin VB.ComboBox CmbSGodown 
         Height          =   315
         Left            =   7500
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   465
         Width           =   1605
      End
      Begin VB.ComboBox CmbSExchangeType 
         Height          =   315
         Left            =   13200
         TabIndex        =   7
         Top             =   465
         Width           =   1605
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   10845
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   465
         Width           =   2325
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
         Height          =   360
         Left            =   12345
         TabIndex        =   14
         Top             =   1605
         Width           =   1215
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   2325
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   465
         Width           =   1965
      End
      Begin VB.ComboBox CmbStack 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmRpt_CMPPV.frx":0027
         Left            =   9135
         List            =   "FrmRpt_CMPPV.frx":0029
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   465
         Width           =   1575
      End
      Begin VB.TextBox TxtSClientName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1200
         Width           =   3405
      End
      Begin VB.TextBox TxtSISINNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3630
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1200
         Width           =   2085
      End
      Begin VB.TextBox TxtSBalBags 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7590
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1200
         Width           =   1485
      End
      Begin VB.TextBox TxtSBalWt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9210
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1200
         Width           =   1365
      End
      Begin VB.ComboBox CmbSSplilageFlag 
         Height          =   315
         ItemData        =   "FrmRpt_CMPPV.frx":002B
         Left            =   5850
         List            =   "FrmRpt_CMPPV.frx":0035
         TabIndex        =   10
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Madeup Flag"
         Height          =   195
         Left            =   11145
         TabIndex        =   33
         Top             =   900
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "State"
         Height          =   195
         Left            =   1035
         TabIndex        =   31
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   7530
         TabIndex        =   30
         Top             =   1665
         Width           =   2280
      End
      Begin VB.Label Label1 
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
         Left            =   9960
         TabIndex        =   29
         Top             =   1665
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CMP"
         Height          =   195
         Left            =   5730
         TabIndex        =   26
         Top             =   195
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Commodity"
         Height          =   195
         Left            =   11625
         TabIndex        =   25
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Godown"
         Height          =   195
         Left            =   8002
         TabIndex        =   24
         Top             =   195
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   195
         Left            =   3000
         TabIndex        =   23
         Top             =   195
         Width           =   615
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Stack No"
         Height          =   195
         Left            =   9585
         TabIndex        =   22
         Top             =   195
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Client Name"
         Height          =   195
         Left            =   1395
         TabIndex        =   21
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Exchange Type"
         Height          =   195
         Left            =   13440
         TabIndex        =   20
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CMS E Lot No"
         Height          =   195
         Left            =   4162
         TabIndex        =   19
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Balance No of Bags "
         Height          =   195
         Left            =   7597
         TabIndex        =   18
         Top             =   900
         Width           =   1470
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Balance Weight"
         Height          =   195
         Left            =   9322
         TabIndex        =   17
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Spillage Flag"
         Height          =   195
         Left            =   6202
         TabIndex        =   16
         Top             =   900
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmRpt_CMPPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCMPName, mCommodityID, mGodownID, mExchangeTypeID, mClientID, mISINID, mClientIDRow As Double
Dim msSM_prefix, mSLocationID, mStateID, mSFlag, mSSplilageFlag, mMadeupFlag As Variant
Dim mInvalidFlag As Boolean
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 31 '28 '26 '25 '24 '18

MSHFlexGrid1.TextMatrix(0, 0) = "State Name"
MSHFlexGrid1.TextMatrix(0, 1) = "Location"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 4) = "Stack No"
MSHFlexGrid1.TextMatrix(0, 5) = "Lot No"
MSHFlexGrid1.TextMatrix(0, 6) = "Lot Type"
MSHFlexGrid1.TextMatrix(0, 7) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 8) = "No Of Bags"
MSHFlexGrid1.TextMatrix(0, 9) = "Quantity"
MSHFlexGrid1.TextMatrix(0, 10) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 11) = "Depositor/Client Name"
MSHFlexGrid1.TextMatrix(0, 12) = "CMSE Lot number" '"ISIN No"
MSHFlexGrid1.TextMatrix(0, 13) = "Pledge No"
MSHFlexGrid1.TextMatrix(0, 14) = "Pledge Bags"
MSHFlexGrid1.TextMatrix(0, 15) = "Pledge Qty"
MSHFlexGrid1.TextMatrix(0, 16) = "Flag"
MSHFlexGrid1.TextMatrix(0, 17) = "Madeup Flag"
MSHFlexGrid1.TextMatrix(0, 18) = "CMSE Lot NED" '"ISIN NED"
MSHFlexGrid1.TextMatrix(0, 19) = "CMSE Lot FED" '"ISIN FED"
MSHFlexGrid1.TextMatrix(0, 20) = "RRN No / WHR No"
MSHFlexGrid1.TextMatrix(0, 21) = "RRN Date / WHR Date"
MSHFlexGrid1.TextMatrix(0, 22) = "RRN Client"
MSHFlexGrid1.TextMatrix(0, 23) = "KR No"
MSHFlexGrid1.TextMatrix(0, 24) = "Agent Name"
MSHFlexGrid1.TextMatrix(0, 25) = "Grade"
MSHFlexGrid1.TextMatrix(0, 26) = "CDTF No"
MSHFlexGrid1.TextMatrix(0, 27) = "CDTF Date"
MSHFlexGrid1.TextMatrix(0, 28) = "Contract ID"
MSHFlexGrid1.TextMatrix(0, 29) = "NContract ID"
MSHFlexGrid1.TextMatrix(0, 30) = "GSL ID"

MSHFlexGrid1.ColWidth(0) = 1800
MSHFlexGrid1.ColWidth(1) = 1800
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 1600
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1600
MSHFlexGrid1.ColWidth(11) = 2000
MSHFlexGrid1.ColWidth(12) = 1400
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1200
MSHFlexGrid1.ColWidth(15) = 1200
MSHFlexGrid1.ColWidth(16) = 1200
MSHFlexGrid1.ColWidth(24) = 1400
'a.  Location
'b.CMP Name
'c.Godown ID
'd.Stack no
'e.Lot no
'f.  Lot type ( Regular / Spillage)
'g.Commodity
'h.no.of bags
'i.Qty
'j.  Exchange Type
'k.  Depositor / Client Name
'L.ISIN no
'm.Pledge ID
'n.  Pledged no. of bags
'o.Pledged Qty

MSHFlexGrid1.RowHeight(0) = 500

End Sub

Private Sub CmbMadeupFlag_LostFocus()
mMadeupFlag = ""
If CmbMadeupFlag.Text = "" Then Exit Sub

If CmbMadeupFlag.Text = "0 ~ Regular" Or CmbMadeupFlag.Text = "1 ~ Madeup" Then
   mMadeupFlag = mID(CmbMadeupFlag.Text, 1, 1)
Else
   MsgBox "Invalid Selection!!!"
   CmbMadeupFlag.SetFocus
End If
End Sub

Private Sub CmbPreFix_GotFocus()
If CmbStateID.Text = "" Then
    MsgBox "Select State First!!!"
    CmbStateID.SetFocus
    Exit Sub
End If

tmp = CmbPreFix.Text

Call TableMst_Location("Where StateID = " & mStateID)
CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbPreFix.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!Locationname ''
    RsMst_Location.MoveNext
Next
CmbPreFix.Text = tmp

'Call TableMst_StoreManLocations("where " & Gen_WcSMLocation)


'If Gen_WcLocation = "" Then
'   Call TableMst_StoreManLocations("")
'Else
'   Call TableMst_StoreManLocations("where " & Gen_WcSMLocation)
'End If


'CmbPreFix.Clear
'If RsMst_StoreManLocations.RecordCount = 0 Then
'   MsgBox "No Bank found"
'   CmbPreFix.SetFocus
'   Exit Sub
'End If
'For I = 1 To RsMst_StoreManLocations.RecordCount
'    CmbPreFix.AddItem RsMst_StoreManLocations!StoreManLocations''

'    'RsMst_StoreManLocations.MoveNext
'Next
'CmbPreFix.Text = tmp

End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
   CmbStack.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location Selection !!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
If tmp <> CmbPreFix.Text Then
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
   CmbStack.Text = ""
End If
msSM_prefix = RsMst_Location!SM_Prefix
mSLocationID = RsMst_Location!LocationID

'RsMst_StoreManLocations.MoveFirst
'RsMst_StoreManLocations.Find "StoreManLocations = '" & CmbPreFix.Text & "'"
'If RsMst_StoreManLocations.EOF Then
'   MsgBox "Invalid PreFix Selection !!!!"
'   CmbPreFix.SetFocus
'   Exit Sub
'End If
'msSM_prefix = RsMst_StoreManLocations!SM_Prefix
'mSLocationId = RsMst_StoreManLocations!StoreManLocationsID
End Sub

Private Sub CmbSCmp_GotFocus()
If CmbPreFix.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If

tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "" & tmp1)
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If

For I = 1 To RsMst_CMP.RecordCount
    CmbSCmp.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp

End Sub

Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   CmbSGodown.Text = ""
   CmbStack.Text = ""
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   mInvalidFlag = True
   CmbSCmp.SetFocus
   Exit Sub
End If
mInvalidFlag = False
If tmp <> CmbSCmp.Text Then
   CmbSGodown.Text = ""
   CmbStack.Text = ""
End If

mCMPName = RsMst_CMP!CMPID
End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbSCommodity.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbSCommodity.Text = tmp
End Sub

Private Sub CmbSCommodity_LostFocus()
If CmbSCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodity.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID
End Sub

Private Sub CmbSExchangeType_GotFocus()
tmp = CmbSExchangeType.Text
Call TableMst_ExchangeType
CmbSExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeType.Text = tmp
End Sub

Private Sub CmbSExchangeType_LostFocus()
If CmbSExchangeType.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

Private Sub CmbSGodown_GotFocus()
If CmbSCmp.Text = "" Or mInvalidFlag = True Then
   MsgBox "Select CMP first !!! "
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPName)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   CmbStack.Text = ""
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbSGodown.SetFocus
   Exit Sub
End If
If tmp <> CmbSGodown.Text Then
   CmbStack.Text = ""
End If

mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbSSplilageFlag_LostFocus()
mSSplilageFlag = ""
If CmbSSplilageFlag.Text = "" Then Exit Sub

If CmbSSplilageFlag.Text = "0 ~ Regular" Or CmbSSplilageFlag.Text = "1 ~ Splillage" Then
   mSSplilageFlag = mID(CmbSSplilageFlag.Text, 1, 1)
Else
   MsgBox "Invalid Selection !!!"
   CmbSSplilageFlag.SetFocus
End If
End Sub

Private Sub CmbStack_GotFocus()

If CmbSGodown.Text = "" Or mGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  CmbSGodown.SetFocus
  Exit Sub
End If
tmp = CmbStack.Text
Call TableMst_GSL(" Where CMPID = " & mCMPName & " And GodownID = " & mGodownID, "D")
CmbStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
    CmbStack.AddItem RsMst_GSL!StackNo
    RsMst_GSL.MoveNext
Next

CmbStack.Text = tmp
End Sub

Private Sub CmbStack_LostFocus()
If CmbStack.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "StackNo = '" & CmbStack.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbStack.SetFocus
   Exit Sub
End If

mCmbStack = CmbStack.Text
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text

If Gen_WcState <> "" Then
   Call TableMst_State("Where " & Gen_WcState)
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
'----
Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbPreFix.Text = ""
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
   CmbStack.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
If CmbStateID.Text <> tmp Then
   CmbPreFix.Text = ""
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
   CmbStack.Text = ""
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Label1.Caption = ""
CmdExcel.Enabled = False
CmbStateID.Text = ""
CmbPreFix.Text = ""
CmbSCmp.Text = ""
CmbSCommodity.Text = ""
CmbSExchangeType.Text = ""
CmbSGodown.Text = ""
CmbSSplilageFlag.Text = ""
CmbStack.Text = ""
TxtSBalBags = ""
TxtSBalWt = ""
TxtSClientName = ""
TxtSISINNo = ""
mClientID = ""
mClientIDRow = 0
mCMPName = ""
mCommodityID = ""
mGodownID = ""
mExchangeTypeID = ""
msSM_prefix = ""
mSSplilageFlag = ""
ProgressBar1.Value = 0

End Sub

Private Sub CmdExcel_Click()

Gen_mTimeStamp = GetTimeStamp


tmp = ""
ans = MsgBox("Do you want to Genearate CSV File", vbYesNo)
If ans = vbNo Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GSLPV.xls")
   Call Xls_Detail_Rpt
Else
   Path = App.Path
   strFilName = "\Excel\" & Gen_mTimeStamp & "Rpt_GSLPV" & Gen_UserName & ".csv"
   strFilLocation = Path + strFilName
   If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
   Open strFilLocation For Output As #1
   For I = 0 To MSHFlexGrid1.Rows - 1
       tmp = ""
       For C = 0 To MSHFlexGrid1.Cols - 1
          If C = 0 Then
             tmp = tmp & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
'          ElseIf c = 6 Then
'             tmp = tmp & "," & "'" & Replace(Replace(MSHFlexGrid1.TextMatrix(I, c), vbNewLine, ""), ",", "~")
          Else
             tmp = tmp & "," & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          End If
       Next
       Print #1, tmp
   Next
   Close #1
   MsgBox "CSV Generated!!!"
End If



End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GSLPV.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String

MsgBox "Please wait till next Message."

Call Grid_Head
ProgressBar1.Value = 0
Label1.Caption = ""
If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, mCMPName, "Mst_CMP.CMPID", "N", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp = Replace(Gen_WcCMP, "CMPID", "mst_CMP.CMPID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If
If CmbSGodown.Text <> "" Then
   wc = GenWc(wc, mGodownID, "Mst_Godown.GodownID", "N", "N")
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "Mst_Commodity.CommodityID", "N", "N")
End If

If CmbSExchangeType.Text <> "" Then
   wc = GenWc(wc, mExchangeTypeID, "Mst_ExchangeType.ExchangeTypeID", "N", "N")
End If
If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, msSM_prefix, "Mst_GSL.sm_prefix", "S", "N")
End If
If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "MS.StateID", "N", "N")
End If

If CmbStack.Text <> "" Then
   wc = GenWc(wc, CmbStack.Text, "Mst_GSL.StackNo", "S", "N")
End If

If TxtSClientName.Text <> "" Then
   wc = GenWc(wc, TxtSClientName, "Mst_Client.ClientName", "S", "L")
End If
If TxtSISINNo.Text <> "" Then
   wc = GenWc(wc, TxtSISINNo, "Mst_GSL.ISINID", "S", "L")
End If

If CmbSSplilageFlag.Text <> "" Then
   wc = GenWc(wc, mSSplilageFlag, "Mst_GSL.SpillageFlag", "S", "N")
End If

If CmbMadeupFlag.Text <> "" Then
   wc = GenWc(wc, mMadeupFlag, "Mst_GSL.MadeupFlag", "S", "N")
End If


If TxtSBalBags <> "" Then
    If Not IsOprater(TxtSBalBags) Then
       wc = GenWc(wc, TxtSBalBags, "Mst_GSL.BalanceBags", "N", "N")
    Else
       wc = GenWc(wc, TxtSBalBags, "Mst_GSL.BalanceBags", "N", "I")
    End If
End If

If TxtSBalWt <> "" Then
    If Not IsOprater(TxtSBalWt) Then
       wc = GenWc(wc, TxtSBalWt, "Mst_GSL.BalanceWeight", "N", "N")
    Else
       wc = GenWc(wc, TxtSBalWt, "Mst_GSL.BalanceWeight", "N", "I")
    End If
End If



'tmp = "SELECT Mst_GSL.GSLID, Mst_GSL.sm_prefix , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo,"
'tmp = tmp & " Mst_GSL.StackNo, Mst_GSL.LotNo , Mst_GSL.DepositeBags, Mst_GSL.CDTFBags, Mst_GSL.BalanceBags, Mst_GSL.DepositeWeight,"
'tmp = tmp & " Mst_GSL.CDTFWeight , Mst_GSL.BalanceWeight, Mst_ExchangeType.ExchangeType, Mst_Client.ClientName, Mst_GSL.Flag ,"
'tmp = tmp & " Mst_GSL.ISINID, Mst_GSL.CDTFNo, Mst_GSL.CDTFDate, Mst_GSL.PledgeNo, Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight ,"
'tmp = tmp & " Mst_GSL.SpillageFlag , Mst_GSL.DepositAdjWt,Mst_GSL.WithdrawalAdjWt,mst_gsl.g_uid,  Mst_GSL.KRNo, Mst_GSL.GrossWt,"
'tmp = tmp & " Mst_GSL.BGSLID , Mst_GSL.FGSLID, mst_gsl.WHRBags,mst_gsl.WHRWeight,Mst_StorageContract.ContractName, "
'tmp = tmp & " Mst_GSL.BagSize,Mst_GSL.BagEquivalent,Mst_GSL.WtQty   FROM "
'tmp = tmp & " Mst_GSL"
'tmp = tmp & " left JOIN Mst_CMP on Mst_CMP.CMPID = Mst_GSL.CMPID"
'tmp = tmp & " left JOIN Mst_Godown on Mst_Godown.GodownID = Mst_GSL.GodownID"
'tmp = tmp & " left JOIN Mst_Commodity on Mst_Commodity.CommodityID = Mst_GSL.CommodityID"
'tmp = tmp & " LEFT JOIN Mst_ExchangeType ON Mst_GSL.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
'tmp = tmp & " left JOIN Mst_Client on Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow"
'tmp = tmp & " LEFT JOIN Mst_ISIN ON Mst_GSL.ISINID = Mst_ISIN.ISINID"
'tmp = tmp & " LEFT JOIN Mst_StorageContract ON Mst_GSL.ContractID = Mst_StorageContract.ContractID"

tmp = "SELECT  MS.StateName, ML.LocationName, Mst_CMP.CMPName,Mst_Godown.GodownNo,"
tmp = tmp & " Mst_GSL.StackNo, Mst_GSL.LotNo , Mst_GSL.SpillageFlag, Mst_Commodity.Commodity,"
tmp = tmp & " Mst_GSL.BalanceBags,Mst_GSL.BalanceWeight, Mst_ExchangeType.ExchangeType, Mst_Client.ClientName,"
tmp = tmp & " Mst_GSL.ISINID , Mst_GSL.PledgeNo, Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight,Mst_GSL.State,Mst_GSL.MadeupFlag,Mst_ISIN.NE_Date,Mst_ISIN.FED,Mst_RRN.RRNNo,Mst_RRN.RRNDate,Mst_Client1.ClientName 'RRNClient',Mst_GSL.KRNo, MA.AgentName,Mst_ISIN.Grade,Mst_GSL.CxTF_DetailsID,Mst_GSL.CDTFDate,Mst_GSL.ContractID,Mst_GSL.NContractID,Mst_GSL.Mst_GSL_ID,TWD.WHRNo,TWD.WHRDate,TWD.CMPID "
tmp = tmp & " From Mst_GSL"
tmp = tmp & " left JOIN Mst_CMP on Mst_CMP.CMPID = Mst_GSL.CMPID"
tmp = tmp & " left join Mst_Location ML on Mst_CMP.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_State MS on MS.StateID=ML.StateID"
tmp = tmp & " left JOIN Mst_Godown on Mst_Godown.GodownID = Mst_GSL.GodownID"
tmp = tmp & " left JOIN Mst_Commodity on Mst_Commodity.CommodityID = Mst_GSL.CommodityID"
tmp = tmp & " LEFT JOIN Mst_ExchangeType ON Mst_GSL.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
'Parag
'tmp = tmp & " left JOIN Mst_Client on Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow"
tmp = tmp & " left JOIN Mst_Client on Mst_Client.ClientID = Mst_GSL.ClientID"
tmp = tmp & " LEFT JOIN Mst_ISIN ON Mst_GSL.ISINID = Mst_ISIN.ISINID"
tmp = tmp & " LEFT JOIN Mst_RRN_LotDetail ON Mst_RRN_LotDetail.ISINID = Mst_ISIN.ISINID"
tmp = tmp & " LEFT JOIN Mst_RRN ON Mst_RRN_LotDetail.RRNID = Mst_RRN.RRNID"
'Parag
'tmp = tmp & " left JOIN Mst_Client Mst_Client1 on Mst_Client1.ClientIDRow = Mst_RRN.ClientIDRow"
tmp = tmp & " left JOIN Mst_Client Mst_Client1 on Mst_Client1.ClientID = Mst_RRN.ClientID"
'Parag
'tmp = tmp & " Left Join Txn_CxTF_Details TCD On Mst_GSL.CDTFNo = TCD.CxTFNo And Mst_GSL.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D' "
tmp = tmp & " Left Join Txn_CxTF_Details TCD On Mst_GSL.CxTF_DetailsID = TCD.CxTF_DetailsID And Mst_GSL.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D' "
tmp = tmp & " Left Join Mst_Agent MA On TCD.AgentID = MA.AgentID "
'Parag
'tmp = tmp & " LEFT JOIN Txn_WHR_Detail TWD ON Mst_GSL.WHRNo = TWD.WHRNo"
tmp = tmp & " LEFT JOIN Txn_WHR_Detail TWD ON Mst_GSL.Txn_WHR_Detail_ID = TWD.Txn_WHR_Detail_ID"


If wc = "" Then
   tmp = tmp & " Where Mst_GSL.Status<>'E' And Mst_GSL.Status<>'B'And Mst_GSL.Status<>'Z'"
Else
   tmp = tmp & " " & wc & " And Mst_GSL.Status<>'E' And Mst_GSL.Status<>'B'And Mst_GSL.Status<>'Z'"
End If
tmp = tmp & " ORDER BY MS.StateName, ML.LocationName,Mst_CMP.CMPName,Mst_GSL.GSLID;"
Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    
    For C = 0 To TmpRs.Fields.Count - 4
        MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
     
    ''-- Ullhas Patil 27 Sep 2017
    MSHFlexGrid1.TextMatrix(I, 3) = "'" & MSHFlexGrid1.TextMatrix(I, 3)
    ''------------
     
    MSHFlexGrid1.TextMatrix(I, 8) = Format(MSHFlexGrid1.TextMatrix(I, 8), "########0")
    MSHFlexGrid1.TextMatrix(I, 9) = Format(MSHFlexGrid1.TextMatrix(I, 9), "########0.000")
    MSHFlexGrid1.TextMatrix(I, 14) = Format(MSHFlexGrid1.TextMatrix(I, 14), "########0")
    MSHFlexGrid1.TextMatrix(I, 15) = Format(MSHFlexGrid1.TextMatrix(I, 15), "########0.000")
    
    If MSHFlexGrid1.TextMatrix(I, 18) <> "" Then
       MSHFlexGrid1.TextMatrix(I, 18) = Format(MSHFlexGrid1.TextMatrix(I, 18), "dd-mmm-yyyy")
    End If
    
    
    If MSHFlexGrid1.TextMatrix(I, 19) <> "" Then
       MSHFlexGrid1.TextMatrix(I, 19) = Format(MSHFlexGrid1.TextMatrix(I, 19), "dd-mmm-yyyy")
    End If
    
    
    If MSHFlexGrid1.TextMatrix(I, 21) <> "" Then
       MSHFlexGrid1.TextMatrix(I, 21) = Format(MSHFlexGrid1.TextMatrix(I, 21), "dd-mmm-yyyy")
    End If
    
    If MSHFlexGrid1.TextMatrix(I, 27) <> "" Then
       MSHFlexGrid1.TextMatrix(I, 27) = Format(MSHFlexGrid1.TextMatrix(I, 27), "dd-mmm-yyyy")
    End If
    
    If MSHFlexGrid1.TextMatrix(I, 6) = "False" Then
       MSHFlexGrid1.TextMatrix(I, 6) = "Regular"
    Else
       MSHFlexGrid1.TextMatrix(I, 6) = "Spillage"
    End If
'    If MSHFlexGrid1.TextMatrix(I, 16) = "O" Then
'       MSHFlexGrid1.TextMatrix(I, 16) = "O~Original"
'    ElseIf MSHFlexGrid1.TextMatrix(I, 16) = "R" Then
'       MSHFlexGrid1.TextMatrix(I, 16) = "R~Revalidation"
'    ElseIf MSHFlexGrid1.TextMatrix(I, 16) = "J" Then
'       MSHFlexGrid1.TextMatrix(I, 16) = "J~Rejected"
'    End If
    
    
    If MSHFlexGrid1.TextMatrix(I, 16) = "A" Then
       MSHFlexGrid1.TextMatrix(I, 16) = "O~Original"
    ElseIf MSHFlexGrid1.TextMatrix(I, 16) = "R" Then
       MSHFlexGrid1.TextMatrix(I, 16) = "R~Revalidation"
    ElseIf MSHFlexGrid1.TextMatrix(I, 16) = "J" Then
       MSHFlexGrid1.TextMatrix(I, 16) = "J~Rejected"
    End If
    
    If MSHFlexGrid1.TextMatrix(I, 17) = "False" Then
       MSHFlexGrid1.TextMatrix(I, 17) = "Regular"
    Else
       MSHFlexGrid1.TextMatrix(I, 17) = "Madeup"
    End If
    If LCase(TmpRs!ExchangeType) = "non ncdex" Then
       'MSHFlexGrid1.TextMatrix(I, 20) = TmpRs!WHRNo
       If IsNull(TmpRs!WHRNo) = False Then
          MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) & "-" & Val(TmpRs!WHRNo) - (IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) * 1000000)
       End If
       MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs!WHRDate), "", Format(TmpRs!WHRDate, "DD-MMM-YYYY"))
    End If
    
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max
Label1.Caption = TmpRs.RecordCount
Label1.Refresh
CmdExcel.Enabled = True
MsgBox "Done!!"

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Grid_Head
CmdExcel.Enabled = False
ProgressBar1.Value = 0
mInvalidFlag = False
End Sub

