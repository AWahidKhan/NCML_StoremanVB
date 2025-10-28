VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_ClientWiseCxTFDetails 
   Caption         =   "Client wise CxTF Details"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11340
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   10560
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9495
      Left            =   75
      TabIndex        =   13
      Top             =   60
      Width           =   14895
      Begin TabDlg.SSTab SSTab1 
         Height          =   7095
         Left            =   105
         TabIndex        =   28
         Top             =   2280
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   12515
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "CxTF Detail"
         TabPicture(0)   =   "FrmRpt_ClientWiseCxTFDetails.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Client Detail"
         TabPicture(1)   =   "FrmRpt_ClientWiseCxTFDetails.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame3 
            Height          =   6615
            Left            =   -74880
            TabIndex        =   31
            Top             =   360
            Width           =   14415
            Begin VB.CheckBox ChkCheckAll 
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
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   255
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   5895
               Left            =   120
               TabIndex        =   32
               Top             =   600
               Width           =   14175
               _ExtentX        =   25003
               _ExtentY        =   10398
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
               Caption         =   "Select / Deselect"
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
               Left            =   435
               TabIndex        =   34
               Top             =   240
               Width           =   1530
            End
         End
         Begin VB.Frame Frame2 
            Height          =   6615
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   14415
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
               Height          =   6255
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   14175
               _ExtentX        =   25003
               _ExtentY        =   11033
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
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   14640
         Begin VB.CheckBox Check1 
            Caption         =   "Opng Bal From Date"
            Height          =   255
            Left            =   165
            TabIndex        =   36
            Top             =   150
            Width           =   1815
         End
         Begin VB.ComboBox CmbClientGroup 
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
            Left            =   5055
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   420
            Width           =   3495
         End
         Begin VB.ComboBox CmbGodownID 
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
            Left            =   8085
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   1140
            Width           =   2400
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
            Left            =   3120
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1140
            Width           =   4935
         End
         Begin VB.OptionButton OptAll 
            Caption         =   "All"
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
            Left            =   3915
            TabIndex        =   24
            Top             =   1560
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton OptNonReser 
            Caption         =   "Non-Reservation"
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
            Left            =   1875
            TabIndex        =   23
            Top             =   1560
            Width           =   2055
         End
         Begin VB.OptionButton OptReser 
            Caption         =   "Reservation"
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
            Left            =   195
            TabIndex        =   22
            Top             =   1560
            Width           =   1575
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
            Left            =   10530
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   1140
            Width           =   3960
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
            Left            =   8573
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   420
            Width           =   5895
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
            Left            =   195
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1140
            Width           =   2880
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
            Height          =   420
            Left            =   11565
            TabIndex        =   9
            Top             =   1590
            Width           =   1455
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
            Height          =   420
            Left            =   10110
            TabIndex        =   8
            Top             =   1590
            Width           =   1455
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
            Height          =   420
            Left            =   13020
            TabIndex        =   10
            Top             =   1590
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   375
            Left            =   2040
            TabIndex        =   0
            Top             =   420
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   661
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
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   106364931
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   375
            Left            =   3540
            TabIndex        =   1
            Top             =   420
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   661
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
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   106364931
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin MSComCtl2.DTPicker TxtOpBalDate 
            Height          =   375
            Left            =   360
            TabIndex        =   35
            Top             =   420
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   661
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
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   106364931
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin VB.Label Label3 
            Caption         =   "Client Group"
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
            Left            =   6195
            TabIndex        =   27
            Top             =   150
            Width           =   1200
         End
         Begin VB.Label LblGodownID 
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
            Left            =   8760
            TabIndex        =   26
            Top             =   825
            Width           =   840
         End
         Begin VB.Label lblCMP 
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
            Height          =   360
            Left            =   5280
            TabIndex        =   25
            Top             =   825
            Width           =   1275
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
            Left            =   12105
            TabIndex        =   21
            Top             =   825
            Width           =   1320
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
            Left            =   10920
            TabIndex        =   20
            Top             =   120
            Width           =   1200
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
            Left            =   1320
            TabIndex        =   19
            Top             =   825
            Width           =   765
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
            Left            =   2220
            TabIndex        =   18
            Top             =   150
            Width           =   1005
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
            Left            =   3840
            TabIndex        =   17
            Top             =   150
            Width           =   780
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
            Left            =   7875
            TabIndex        =   16
            Top             =   1635
            Width           =   210
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
            Left            =   5475
            TabIndex        =   15
            Top             =   1635
            Width           =   2325
         End
      End
   End
End
Attribute VB_Name = "FrmRpt_ClientWiseCxTFDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mClientIDRow, mLocationID, mSM_Prefix, mCommodityID, mCMPID, mGodownID, mClientGroupID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim wc As String
Dim mStock, mStockBag As Variant
Dim mOpTransDepBag, mOpTransDepWt, mOpTransWidBag, mOpTransWidWt As Variant
Dim mBGSLIDs As String
Private Sub Check1_Click()
If Check1.Value = 1 Then
    TxtOpBalDate.Visible = True
Else
    TxtOpBalDate.Visible = False
End If
End Sub

Private Sub ChkCheckAll_Click()
If MSHFlexGrid2.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.row = I
       MSHFlexGrid2.TextMatrix(I, 0) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.row = I
       MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
   Next
End If

End Sub

Private Sub CmbClientId_GotFocus()
'If CmbClientGroup.Text = "" Then
'   MsgBox "Please Select Client Group !!!"
'   CmbClientGroup.SetFocus
'   Exit Sub
'End If

tmp = CmbClientID.Text
Call TableMst_Client("Where ExchangeTypeID = 1") 'And ClientGroupID=" & mClientGroupID
CmbClientID.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClientID.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClientID.Text = tmp
End Sub
'----
Private Sub CmbClientId_LostFocus()
If CmbClientID.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClientID.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbClientID.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientID
CmbClientGroup.Text = ""
CmbClientGroup.Enabled = False
Call Grid_Head1
SSTab1.Tab = 0
ChkCheckAll.Value = 0
End Sub
Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("") 'Where StateID=" & mStateID & "
Else
   Call TableMst_Location("Where " & Gen_WcLocation) 'StateID=" & mStateID & " And
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   mSM_Prefix = ""
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

Private Sub CmdClear_Click()
TxtOpBalDate.Value = Date - 60
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmbLocationID.Text = ""
CmbClientID.Text = ""
CmbClientGroup.Text = ""
CmbGodownID.Text = ""
CmbCMPName.Text = ""
CmbClientGroup.Enabled = True
CmbClientID.Enabled = True
OptAll.Value = True
CmbCommodityID.Text = ""
Call Grid_Head
Call Grid_Head1
LblRecordCount.Caption = "-"
SSTab1.Tab = 0
ChkCheckAll.Value = 0
End Sub


Private Sub CmdGetReport_Click()
Dim mTotDays As Double
Dim mDate As Date
Dim mClientIDs As String
Dim mOpDepBag, mOpWidBag, mOpDepWt, mOpWidWt As Variant
Dim mTotDepBag, mTotDepWt, mTotWitBag, mTotWitWt As Variant
Dim mMaxDepWt, mMaxDepBag, mMaxWidWt, mMaxWidBag, mMaxOpBalBag, mMaxOpBalWt, mMaxBalBag, mMaxBalWt, mMaxTransRBag, mMaxTransRWt, mMaxTransIBag, mMaxTransIWt As Variant
Call Grid_Head
SSTab1.Tab = 0

MsgBox "Wait till next message"

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If
If Check1.Value = 1 Then
    If TxtOpBalDate.Value > TxtFromDate.Value Then
        MsgBox "'Opening Balance From Date' must be less than 'From date'!!!"
        TxtOpBalDate.SetFocus
    Exit Sub
    End If
End If
If CmbLocationID.Text = "" Then
   MsgBox "Please select Location!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbClientID.Text = "" And CmbClientGroup.Text = "" Then
   MsgBox "Please select 'Client Group' or 'Client Name'!!"
   CmbClientGroup.SetFocus
   Exit Sub
End If
If CmbClientGroup.Text <> "" Then
   If CheckClientSelect = False Then
      MsgBox " Select atleat one Client!!"
      SSTab1.Tab = 1
      Exit Sub
   End If
   mClientIDs = "(" & GetSelectedClients & ")"
End If

'--For Deposit Op Bal
tmp = " select sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTF_DetailsID  = b.CxTF_DetailsID and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.Mst_GSL_ID = b.Mst_GSL_ID " 'and SpillageFlag = 0"
tmp = tmp & " where a.exchangetypeid = 1 And a.TxnType='D' And a.CxTFDate<'" & Format(TxtFromDate, Gen_DatFormat) & "' "
If Check1.Value = 1 Then
    tmp = tmp & " And a.CxTFDate>='" & Format(TxtOpBalDate, Gen_DatFormat) & "'"
End If
If CmbClientGroup.Text <> "" Then
   tmp = tmp & " And b.ClientID in " & mClientIDs
Else
   tmp = tmp & " And b.ClientID = " & mClientIDRow & ""
End If
tmp = tmp & " And a.SM_Prefix='" & mSM_Prefix & "'"
If OptReser.Value = True Then
   tmp = tmp & " and isnull(C.NContractID,0)>0"
ElseIf OptNonReser.Value = True Then
   tmp = tmp & " and isnull(C.NContractID,0)=0"
End If
If CmbCommodityID.Text <> "" Then
   tmp = tmp & " And a.CommodityID=" & mCommodityID
End If
If CmbCMPName.Text <> "" Then
   tmp = tmp & " And c.CMPID=" & mCMPID
End If
If CmbGodownID.Text <> "" Then
   tmp = tmp & " And c.GodownID=" & mGodownID
End If

Call TmpTable

'--For Withdrawn Op Bal

tmp = "select sum(b.NoOfBags) as NoOfBags ,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTF_DetailsID  = b.CxTF_DetailsID and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.Mst_GSL_ID = b.Mst_GSL_ID" ' and SpillageFlag = 0"
tmp = tmp & " where a.exchangetypeid = 1 And a.TxnType='W' And a.CxTFDate <'" & Format(TxtFromDate, Gen_DatFormat) & "' And a.SM_Prefix='" & mSM_Prefix & "'"
If Check1.Value = 1 Then
    tmp = tmp & " And a.CxTFDate>='" & Format(TxtOpBalDate, Gen_DatFormat) & "'"
End If

If CmbClientGroup.Text <> "" Then
   tmp = tmp & " And b.ClientID in " & mClientIDs
Else
   tmp = tmp & " And b.ClientID = " & mClientIDRow & ""
End If

If OptReser.Value = True Then
   tmp = tmp & " and isnull(C.NContractID,0)>0"
ElseIf OptNonReser.Value = True Then
   tmp = tmp & " and isnull(C.NContractID,0)=0"
End If

If CmbCommodityID.Text <> "" Then
   tmp = tmp & " And a.CommodityID=" & mCommodityID
End If
If CmbCMPName.Text <> "" Then
   tmp = tmp & " And c.CMPID=" & mCMPID
End If
If CmbGodownID.Text <> "" Then
   tmp = tmp & " And c.GodownID=" & mGodownID
End If

Call Tmp1Table

'--For Deposit
'''tmp = " select a.CxTFDate,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
'''tmp = tmp & " from txn_cxtf_details a"
'''tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
'''tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid and SpillageFlag = 0"
'''tmp = tmp & " where a.exchangetypeid = 1 And a.TxnType='D' And a.CxTFDate>='" & Format(TxtFromDate, Gen_DatFormat) & "' And  a.CxTFDate<='" & Format(TxtToDate, Gen_DatFormat) & "' And c.ClientIDRow=" & mClientIDRow & " And a.SM_Prefix='" & mSM_Prefix & "'"
'''If OptReser.Value = True Then
'''   tmp = tmp & " and isnull(C.NContractID,0)>0"
'''ElseIf OptNonReser.Value = True Then
'''   tmp = tmp & " and isnull(C.NContractID,0)=0"
'''End If
'''If CmbCommodityID.Text <> "" Then
'''   tmp = tmp & " And a.CommodityID=" & mCommodityID
'''End If
'''tmp = tmp & " group by  a.CxTFDate"

tmp = "select a.CxTF_DetailsID,a.TxnType,a.CxTFDate,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTF_DetailsID  = b.CxTF_DetailsID and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.Mst_GSL_ID = b.Mst_GSL_ID" ' and SpillageFlag = 0"
tmp = tmp & " where a.exchangetypeid = 1 And a.CxTFDate>='" & Format(TxtFromDate, Gen_DatFormat) & "' And  a.CxTFDate<='" & Format(TxtToDate, Gen_DatFormat) & "' And a.SM_Prefix='" & mSM_Prefix & "'"
wc = ""
If CmbClientGroup.Text <> "" Then
   tmp = tmp & " And b.ClientID in " & mClientIDs
   wc = wc & " And C.ClientID in " & mClientIDs
Else
   tmp = tmp & " And b.ClientID = " & mClientIDRow & ""
   wc = wc & " And C.ClientID = " & mClientIDRow & ""
End If

If OptReser.Value = True Then
   tmp = tmp & " and isnull(C.NContractID,0)>0"
   wc = wc & " and isnull(C.NContractID,0)>0"
ElseIf OptNonReser.Value = True Then
   tmp = tmp & " and isnull(C.NContractID,0)=0"
   wc = wc & " and isnull(C.NContractID,0)=0"
End If

If CmbCommodityID.Text <> "" Then
   tmp = tmp & " And a.CommodityID=" & mCommodityID
   wc = wc & " And C.CommodityID=" & mCommodityID
End If

If CmbCMPName.Text <> "" Then
   tmp = tmp & " And c.CMPID=" & mCMPID
End If
If CmbGodownID.Text <> "" Then
   tmp = tmp & " And c.GodownID=" & mGodownID
End If

tmp = tmp & " group by a.CxTF_DetailsID,a.CxTFDate,a.TxnType"
tmp = tmp & " order by  a.CxTFDate,a.TxnType,a.CxTF_DetailsID"

Call Tmp2Table

'Commented on 04 Dec 2011 135 PM Parag because TmpRs6 is no where used in below code except commented code in loop

'tmp = "Select max(a.CxTFDate) as CxTFDate,a.CxTFNo,C.gslid,max(C.BalanceWeight) as GrossWeight"
'tmp = tmp & " From txn_cxtf_details a"
'tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
'tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid and SpillageFlag = 0"
'tmp = tmp & " Where c.GSLid in (Select distinct b.GSLID from txn_cxtf_details a"
'tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
'tmp = tmp & " where a.exchangetypeid = 1 And a.TxnType='W' And a.CxTFDate>='" & Format(TxtFromDate, Gen_DatFormat) & "' And  a.CxTFDate<='" & Format(TxtToDate, Gen_DatFormat) & "' And a.SM_Prefix='" & mSM_Prefix & "')"
'tmp = tmp & " And c.Flag='E' And c.BalanceBags=0 And a.SM_Prefix='" & mSM_Prefix & "' And  a.exchangetypeid = 1 And a.TxnType='W'" 'and c.SpillageFlag = 0
'If CmbClientGroup.Text <> "" Then
'   tmp = tmp & " And b.ClientIDRow in " & mClientIDs
'Else
'   tmp = tmp & " And b.ClientIDRow = " & mClientIDRow & ""
'End If
'
'If OptReser.Value = True Then
'   tmp = tmp & " and isnull(C.NContractID,0)>0"
'ElseIf OptNonReser.Value = True Then
'   tmp = tmp & " and isnull(C.NContractID,0)=0"
'End If
'
'If CmbCommodityID.Text <> "" Then
'   tmp = tmp & " And a.CommodityID=" & mCommodityID
'End If
'
'If CmbCMPName.Text <> "" Then
'   tmp = tmp & " And c.CMPID=" & mCMPID
'End If
'If CmbGodownID.Text <> "" Then
'   tmp = tmp & " And c.GodownID=" & mGodownID
'End If
'tmp = tmp & " group by c.gslid,a.CxTFNo"
'tmp = tmp & " order by a.CxTFDate,a.CxTFNo,C.gslid"
'
'Call Tmp6Table

mTotDays = DateDiff("d", TxtFromDate.Value, TxtToDate.Value)
mDate = TxtFromDate.Value
Call Grid_Head

mOpDepBag = 0
mOpDepWt = 0
mOpWidBag = 0
mOpWidWt = 0

If TmpRs.RecordCount > 0 Then
   mOpDepBag = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
   mOpDepWt = IIf(IsNull(TmpRs.Fields(1)), 0, TmpRs.Fields(1))
End If
If TmpRs1.RecordCount > 0 Then
   mOpWidBag = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
   mOpWidWt = IIf(IsNull(TmpRs1.Fields(1)), 0, TmpRs1.Fields(1))
End If

'If TmpRs5.RecordCount > 0 Then
   'mOpWidWt = mOpWidWt + IIf(IsNull(TmpRs5.Fields(0)), 0, TmpRs5.Fields(0))
'End If


For I = 0 To mTotDays
    r = MSHFlexGrid1.Rows - 1
    
    Call TransferReceived(mDate)
    MSHFlexGrid1.TextMatrix(r, 8) = mStockBag
    MSHFlexGrid1.TextMatrix(r, 9) = mStock
    
    Call TransferIssued(mDate)
    MSHFlexGrid1.TextMatrix(r, 10) = mStockBag
    MSHFlexGrid1.TextMatrix(r, 11) = mStock
           
    If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
    TmpRs2.Filter = "CxTFDate='" & mDate & "'"
    If TmpRs2.RecordCount > 0 Then
       For rn = 1 To TmpRs2.RecordCount
           r = MSHFlexGrid1.Rows - 1
           MSHFlexGrid1.TextMatrix(r, 0) = TmpRs2!CxTF_DetailsID
           MSHFlexGrid1.TextMatrix(r, 1) = Format(mDate, "dd-MMM-yyyy")
           If TmpRs2!TxnType = "D" Then
              MSHFlexGrid1.TextMatrix(r, 4) = Val(MSHFlexGrid1.TextMatrix(r, 4)) + IIf(IsNull(TmpRs2!NoofBags), 0, TmpRs2!NoofBags)
              'mTotDepBag = mTotDepBag + IIf(IsNull(TmpRs2!NoofBags), 0, TmpRs2!NoofBags)
              MSHFlexGrid1.TextMatrix(r, 5) = Val(MSHFlexGrid1.TextMatrix(r, 5)) + IIf(IsNull(TmpRs2!GrossWeight), 0, TmpRs2!GrossWeight)
              'mTotDepWt = mTotDepWt + IIf(IsNull(TmpRs2!GrossWeight), 0, TmpRs2!GrossWeight)
           Else
              MSHFlexGrid1.TextMatrix(r, 6) = Val(MSHFlexGrid1.TextMatrix(r, 6)) + IIf(IsNull(TmpRs2!NoofBags), 0, TmpRs2!NoofBags)
              'mTotWitBag = mTotWitBag + IIf(IsNull(TmpRs2!NoofBags), 0, TmpRs2!NoofBags)
              MSHFlexGrid1.TextMatrix(r, 7) = Val(MSHFlexGrid1.TextMatrix(r, 7)) + IIf(IsNull(TmpRs2!GrossWeight), 0, TmpRs2!GrossWeight)
              'mTotWitWt = mTotWitWt + IIf(IsNull(TmpRs2!NoofBags), 0, TmpRs2!NoofBags)
           End If
           
           If r = 3 Then
              MSHFlexGrid1.TextMatrix(r, 2) = Format((mOpDepBag + mOpTransDepBag) - (mOpWidBag + mOpTransWidBag), "#####0.000")   ' mOpDepBag - mOpWidBag
              MSHFlexGrid1.TextMatrix(r, 3) = Format((mOpDepWt + mOpTransDepWt) - (mOpWidWt + mOpTransWidWt), "#####0.000") 'mOpDepWt - mOpWidWt
           Else
              MSHFlexGrid1.TextMatrix(r, 2) = MSHFlexGrid1.TextMatrix(r - 1, 12)
              MSHFlexGrid1.TextMatrix(r, 3) = MSHFlexGrid1.TextMatrix(r - 1, 13)
           End If
           MSHFlexGrid1.TextMatrix(r, 12) = Val(MSHFlexGrid1.TextMatrix(r, 2)) + Val(MSHFlexGrid1.TextMatrix(r, 4)) - Val(MSHFlexGrid1.TextMatrix(r, 6))
           MSHFlexGrid1.TextMatrix(r, 13) = Val(MSHFlexGrid1.TextMatrix(r, 3)) + Val(MSHFlexGrid1.TextMatrix(r, 5)) - Val(MSHFlexGrid1.TextMatrix(r, 7))
    
           MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
           TmpRs2.MoveNext
       Next
       
    Else
       r = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(r, 1) = Format(mDate, "dd-MMM-yyyy")
       If r = 3 Then
          MSHFlexGrid1.TextMatrix(r, 2) = (mOpDepBag + mOpTransDepBag) - (mOpWidBag + mOpTransWidBag)   ' mOpDepBag - mOpWidBag
          MSHFlexGrid1.TextMatrix(r, 3) = Format((mOpDepWt + mOpTransDepWt) - (mOpWidWt + mOpTransWidWt), "#####0.000")   ' mOpDepBag - mOpWidBag
       Else
          MSHFlexGrid1.TextMatrix(r, 2) = MSHFlexGrid1.TextMatrix(r - 1, 12)
          MSHFlexGrid1.TextMatrix(r, 3) = MSHFlexGrid1.TextMatrix(r - 1, 13)
       End If
       MSHFlexGrid1.TextMatrix(r, 12) = Val(MSHFlexGrid1.TextMatrix(r, 2)) + Val(MSHFlexGrid1.TextMatrix(r, 4)) + Val(MSHFlexGrid1.TextMatrix(r, 8)) - (Val(MSHFlexGrid1.TextMatrix(r, 6)) + Val(MSHFlexGrid1.TextMatrix(r, 10)))
       MSHFlexGrid1.TextMatrix(r, 13) = Val(MSHFlexGrid1.TextMatrix(r, 3)) + Val(MSHFlexGrid1.TextMatrix(r, 5)) + Val(MSHFlexGrid1.TextMatrix(r, 9)) - (Val(MSHFlexGrid1.TextMatrix(r, 7)) + Val(MSHFlexGrid1.TextMatrix(r, 11)))
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    End If
    
    
    TmpRs2.Filter = ""
    mDate = mDate + 1
Next
mMaxDepBag = 0
mMaxDepWt = 0
mMaxOpBalBag = 0
mMaxOpBalWt = 0
mMaxWidBag = 0
mMaxWidWt = 0
mOpBalBag = 0
mOpBalWt = 0
mMaxBalBag = 0
mMaxBalWt = 0
mMaxTransRBag = 0
mMaxTransRWt = 0
mMaxTransIBag = 0
mMaxTransIWt = 0

For I = 3 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.TextMatrix(1, 4) = Val(MSHFlexGrid1.TextMatrix(1, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
    MSHFlexGrid1.TextMatrix(1, 5) = Val(MSHFlexGrid1.TextMatrix(1, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
    MSHFlexGrid1.TextMatrix(1, 6) = Val(MSHFlexGrid1.TextMatrix(1, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
    MSHFlexGrid1.TextMatrix(1, 7) = Val(MSHFlexGrid1.TextMatrix(1, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
    MSHFlexGrid1.TextMatrix(1, 8) = Val(MSHFlexGrid1.TextMatrix(1, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 8))
    MSHFlexGrid1.TextMatrix(1, 9) = Val(MSHFlexGrid1.TextMatrix(1, 9)) + Val(MSHFlexGrid1.TextMatrix(I, 9))
    MSHFlexGrid1.TextMatrix(1, 10) = Val(MSHFlexGrid1.TextMatrix(1, 10)) + Val(MSHFlexGrid1.TextMatrix(I, 10))
    MSHFlexGrid1.TextMatrix(1, 11) = Val(MSHFlexGrid1.TextMatrix(1, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 11))

    If Val(MSHFlexGrid1.TextMatrix(I, 2)) > mMaxOpBalBag Then
       mMaxOpBalBag = Val(MSHFlexGrid1.TextMatrix(I, 2))
    End If
    If Val(MSHFlexGrid1.TextMatrix(I, 3)) > mMaxOpBalWt Then
       mMaxOpBalWt = Val(MSHFlexGrid1.TextMatrix(I, 3))
    End If
    If Val(MSHFlexGrid1.TextMatrix(I, 4)) > mMaxDepBag Then
       mMaxDepBag = Val(MSHFlexGrid1.TextMatrix(I, 4))
    End If
    If Val(MSHFlexGrid1.TextMatrix(I, 5)) > mMaxDepWt Then
       mMaxDepWt = Val(MSHFlexGrid1.TextMatrix(I, 5))
    End If
    If Val(MSHFlexGrid1.TextMatrix(I, 6)) > mMaxWidBag Then
       mMaxWidBag = Val(MSHFlexGrid1.TextMatrix(I, 6))
    End If
    If Val(MSHFlexGrid1.TextMatrix(I, 7)) > mMaxWidWt Then
       mMaxWidWt = Val(MSHFlexGrid1.TextMatrix(I, 7))
    End If
    
    If Val(MSHFlexGrid1.TextMatrix(I, 8)) > mMaxTransRBag Then
       mMaxTransRBag = Val(MSHFlexGrid1.TextMatrix(I, 8))
    End If
    If Val(MSHFlexGrid1.TextMatrix(I, 9)) > mMaxTransRWt Then
       mMaxTransRWt = Val(MSHFlexGrid1.TextMatrix(I, 9))
    End If
    
    If Val(MSHFlexGrid1.TextMatrix(I, 10)) > mMaxTransIBag Then
       mMaxTransIBag = Val(MSHFlexGrid1.TextMatrix(I, 10))
    End If
    If Val(MSHFlexGrid1.TextMatrix(I, 11)) > mMaxTransIWt Then
       mMaxTransIWt = Val(MSHFlexGrid1.TextMatrix(I, 11))
    End If
    
    If Val(MSHFlexGrid1.TextMatrix(I, 12)) > mMaxBalBag Then
       mMaxBalBag = Val(MSHFlexGrid1.TextMatrix(I, 12))
    End If
    If Val(MSHFlexGrid1.TextMatrix(I, 13)) > mMaxBalWt Then
       mMaxBalWt = Val(MSHFlexGrid1.TextMatrix(I, 13))
    End If
Next
r = MSHFlexGrid1.Rows - 1
MSHFlexGrid1.TextMatrix(r, 1) = "Peak Values"
MSHFlexGrid1.TextMatrix(r, 2) = mMaxOpBalBag
MSHFlexGrid1.TextMatrix(r, 3) = mMaxOpBalWt
MSHFlexGrid1.TextMatrix(r, 4) = mMaxDepBag
MSHFlexGrid1.TextMatrix(r, 5) = mMaxDepWt
MSHFlexGrid1.TextMatrix(r, 6) = mMaxWidBag
MSHFlexGrid1.TextMatrix(r, 7) = mMaxWidWt
MSHFlexGrid1.TextMatrix(r, 8) = mMaxTransRBag
MSHFlexGrid1.TextMatrix(r, 9) = mMaxTransRWt
MSHFlexGrid1.TextMatrix(r, 10) = mMaxTransIBag
MSHFlexGrid1.TextMatrix(r, 11) = mMaxTransIWt
MSHFlexGrid1.TextMatrix(r, 12) = mMaxBalBag
MSHFlexGrid1.TextMatrix(r, 13) = mMaxBalWt

MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1

CmdExcel.Enabled = True
MsgBox "Done!!"
End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 4
MSHFlexGrid1.Cols = 14
MSHFlexGrid1.FixedRows = 3
MSHFlexGrid1.FixedCols = 2
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.MergeCells = flexMergeRestrictRows

MSHFlexGrid1.TextMatrix(2, 0) = "CxTF No"
MSHFlexGrid1.TextMatrix(2, 1) = "Date"

MSHFlexGrid1.TextMatrix(0, 2) = "Opening Balance"
MSHFlexGrid1.TextMatrix(0, 3) = "Opening Balance"
MSHFlexGrid1.TextMatrix(0, 4) = "Deposits"
MSHFlexGrid1.TextMatrix(0, 5) = "Deposits"
MSHFlexGrid1.TextMatrix(0, 6) = "Withdrawal"
MSHFlexGrid1.TextMatrix(0, 7) = "Withdrawal"

MSHFlexGrid1.TextMatrix(0, 8) = "Transfer Received"
MSHFlexGrid1.TextMatrix(0, 9) = "Transfer Received"
MSHFlexGrid1.TextMatrix(0, 10) = "Transfer Issued"
MSHFlexGrid1.TextMatrix(0, 11) = "Transfer Issued"

MSHFlexGrid1.TextMatrix(0, 12) = "Closing Stock"
MSHFlexGrid1.TextMatrix(0, 13) = "Closing Stock"

MSHFlexGrid1.TextMatrix(2, 2) = "Bags"
MSHFlexGrid1.TextMatrix(2, 3) = "Gross Weight(MT)"
MSHFlexGrid1.TextMatrix(2, 4) = "Bags"
MSHFlexGrid1.TextMatrix(2, 5) = "Gross Weight(MT)"
MSHFlexGrid1.TextMatrix(2, 6) = "Bags"
MSHFlexGrid1.TextMatrix(2, 7) = "Gross Weight(MT)"
MSHFlexGrid1.TextMatrix(2, 8) = "Bags"
MSHFlexGrid1.TextMatrix(2, 9) = "Gross Weight(MT)"
MSHFlexGrid1.TextMatrix(2, 10) = "Bags"
MSHFlexGrid1.TextMatrix(2, 11) = "Gross Weight(MT)"
MSHFlexGrid1.TextMatrix(2, 12) = "Bags"
MSHFlexGrid1.TextMatrix(2, 13) = "Gross Weight(MT)"

MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1200

MSHFlexGrid1.MergeRow(0) = True
MSHFlexGrid1.MergeCol(1) = True
MSHFlexGrid1.MergeCol(2) = True
MSHFlexGrid1.MergeCol(3) = True
MSHFlexGrid1.MergeCol(4) = True
MSHFlexGrid1.MergeCol(5) = True
MSHFlexGrid1.MergeCol(6) = True
MSHFlexGrid1.MergeCol(7) = True
MSHFlexGrid1.MergeCol(8) = True
MSHFlexGrid1.MergeCol(9) = True
MSHFlexGrid1.MergeCol(10) = True
MSHFlexGrid1.MergeCol(11) = True
MSHFlexGrid1.MergeCol(12) = True
MSHFlexGrid1.MergeCol(13) = True

MSHFlexGrid1.RowHeight(2) = 800

End Sub

Private Sub Form_Load()
Call Grid_Head
Dim FrmLoadAccess() As Boolean

''FrmLoadAccess = GetFrmLoadAccess(Me.Name)
''If FrmLoadAccess(0) = False Then
''   MsgBox "Access Denied !!!!!!!!!"
''   Frame1.Enabled = False
''   Frame2.Enabled = False
''   Exit Sub
''End If

Frame1.Enabled = True
Frame2.Enabled = True

Call CmdClear_Click
CmdExcel.Enabled = False
End Sub
Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
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

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub


Private Sub MSHFlexGrid2_Click()
With MSHFlexGrid2
    If .TextMatrix(.row, 2) = "" Then Exit Sub
    If .Col = 0 Then
        If .TextMatrix(.row, 0) = strUnChecked Then
           .TextMatrix(.row, 0) = strChecked
        Else
           .TextMatrix(.row, 0) = strUnChecked
          
        End If
    End If
End With

End Sub

Private Sub VScroll1_Change()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l
'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l
End Sub

Private Sub HScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub SetScrollBars()
'    Frame1.Width = 10000
'    Frame1.Height = 20000
VScroll1.Top = 0
HScroll1.Left = 0
VScroll1.Left = Me.Width - (VScroll1.Width + 120)
HScroll1.Top = Me.Height - ((HScroll1.Height * 2) + 250)
VScroll1.Height = Me.Height - 500
HScroll1.Width = Me.Width - 500
VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = Me.ScaleHeight - Frame4.Height
'VScroll1.LargeChange = Me.ScaleHeight
'VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True
If Frame4.Width < Me.Width Then HScroll1.Enabled = False
'HScroll1.Enabled = True
'VScroll1.Enabled = True
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CxTFDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CxTFDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 2
tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge

If CmbClientGroup.Text <> "" Then
   oWS.Cells(mRow, C + 1) = "Client Group :" & CmbClientGroup.Text
Else
   oWS.Cells(mRow, C + 1) = "Client Name :" & CmbClientID.Text
End If

mRow = mRow + 1
tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "From Date :" & Format(TxtFromDate.Value, "dd-MMM-yyyy") & " To Date : " & Format(TxtToDate.Value, "dd-MMM-yyyy")

mRow = mRow + 1
tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
mstr = "Location :" & CmbLocationID.Text & ""
If CmbCMPName.Text <> "" Then mstr = mstr & ", CMP Name : " & CmbCMPName.Text
If CmbGodownID.Text <> "" Then mstr = mstr & ", Godown No : " & CmbGodownID.Text
oWS.Cells(mRow, C + 1) = mstr

mRow = mRow + 1
tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Commodity :" & CmbCommodityID.Text & ""

mRow = mRow + 2

For I = 0 To MSHFlexGrid1.Rows - 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
        
        
    Next
    mRow = mRow + 1
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub

Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text

Call TableMst_CMP(" where LocationID=" & mLocationID & " ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "First Select Location!!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub

Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodownID_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownID.Text
'Call TableMst_Godown("Where CloseDate is null And CMPID = " & mCMPID)
Call TableMst_Godown("Where CMPID = " & mCMPID)

CmbGodownID.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownID.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub
'----
Private Sub CmbGodownID_LostFocus()
If CmbGodownID.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownID.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodownID.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID

End Sub

Private Sub CmbClientGroup_GotFocus()
tmp = CmbClientGroup.Text
Call TableMst_ClientGroup("")
CmbClientGroup.Clear
If RsMst_ClientGroup.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ClientGroup.RecordCount
    CmbClientGroup.AddItem RsMst_ClientGroup!ClientGroupName
    RsMst_ClientGroup.MoveNext
Next
CmbClientGroup.Text = tmp
End Sub
'----
Private Sub CmbClientGroup_LostFocus()
If CmbClientGroup.Text = "" Then
   Exit Sub
End If
RsMst_ClientGroup.MoveFirst
RsMst_ClientGroup.Find "ClientGroupName = '" & CmbClientGroup.Text & "'"
If RsMst_ClientGroup.EOF Then
   MsgBox "Invalid selection "
   CmbClientGroup.SetFocus
   Exit Sub
End If
mClientGroupID = RsMst_ClientGroup!ClientGroupID
CmbClientID.Text = ""
CmbClientID.Enabled = False
Call GetGridData
SSTab1.Tab = 1
End Sub

Private Sub GetGridData()

Call Grid_Head1
tmp = " Select ClientID,ClientIDRow,ClientName From Mst_Client Where ClientGroupID=" & mClientGroupID & "Order By ClientName"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   
   For I = 1 To TmpRs.RecordCount
       mRow = MSHFlexGrid2.Rows - 1
       MSHFlexGrid2.TextMatrix(mRow, 1) = TmpRs!ClientID
       MSHFlexGrid2.TextMatrix(mRow, 2) = TmpRs!ClientIDRow
       MSHFlexGrid2.TextMatrix(mRow, 3) = TmpRs!ClientName
       MSHFlexGrid2.row = mRow
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.CellFontName = "Wingdings"
       MSHFlexGrid2.CellFontSize = 12
       MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid2.Text = strChecked
       TmpRs.MoveNext
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   Next
   ChkCheckAll.Value = 1
Else
  MsgBox "No Record Found For Grid!!!"
  Exit Sub
End If
End Sub


Public Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .Cols = 4
    .Font.Size = 10
    .FixedCols = 0
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Select"
    .TextMatrix(0, 1) = "ClientID"
    .TextMatrix(0, 2) = "Client ID"
    .TextMatrix(0, 3) = "Client Name"
    .ColAlignment(2) = vbLeftJustify
    .ColAlignment(3) = vbLeftJustify
    .ColWidth(0) = 800
    .ColWidth(1) = 0
    .ColWidth(2) = 1200
    .ColWidth(3) = 4000
End With
End Sub

Private Function GetSelectedClients() As String
Dim Retval As String
Retval = ""
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
       If Retval = "" Then
          Retval = MSHFlexGrid2.TextMatrix(I, 1)
       Else
          Retval = Retval & "," & MSHFlexGrid2.TextMatrix(I, 1)
       End If
    End If
Next
GetSelectedClients = Retval
End Function

Private Function CheckClientSelect() As Boolean
Dim Retval As Boolean
Retval = False
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
CheckClientSelect = Retval
End Function

Private Function OpeningTransferReceived() As Variant
Dim mDepStock, mWidStock As Variant
Dim mDepStockBag, mWidStockBag As Variant
Dim mGSLIDs As String
mGSLIDs = GetBGSLIDs()

tmp = "Select GSLTransferDate,a.TxnType,sum(b.NoOfBags) as NoOfBags, sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTF_DetailsID  = b.CxTF_DetailsID and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.Mst_GSL_ID = b.Mst_GSL_ID"
tmp = tmp & " inner join Txn_GSL_Transfer d on  c.SM_Prefix =d.SM_Prefix and c.Mst_GSL_ID = d.Old_Mst_GSL_ID"
tmp = tmp & " where a.exchangetypeid =1 And d.GSLTransferDate<'" & TxtFromDate.Value & "' "
tmp = tmp & " And c.Mst_GSL_ID in (" & mGSLIDs & ")"
tmp = tmp & " And c.SM_Prefix='" & mSM_Prefix & "'"

If CmbGodownID.Text <> "" Then
   tmp = tmp & " And c.GodownID<>" & mGodownID
Else
   If CmbCMPName.Text <> "" Then
      tmp = tmp & " And c.CMPID<>" & mCMPID
   End If
End If


''If CmbCMPName.Text <> "" Then
''   tmp = tmp & " And c.CMPID<>" & mCMPID
''End If
''If CmbGodownID.Text <> "" Then
''   tmp = tmp & " And c.GodownID<>" & mGodownID
''End If

tmp = tmp & " " & wc
tmp = tmp & " group by d.GSLTransferDate,a.TxnType"

Call Tmp5Table

mStock = 0
mWidStock = 0
mDepStock = 0

mStockBag = 0
mWidStockBag = 0
mDepStockBag = 0

If TmpRs5.RecordCount > 0 Then
   For I = 1 To TmpRs5.RecordCount
       If TmpRs5!TxnType = "D" Then
          mDepStock = TmpRs5!GrossWeight
          mDepStockBag = TmpRs5!NoofBags
       Else
          mWidStock = TmpRs5!GrossWeight
          mWidStockBag = TmpRs5!NoofBags
       End If
       TmpRs5.MoveNext
   Next
End If
mOpTransDepWt = Format(mDepStock - mWidStock, "#####0.000")
mOpTransDepBag = Format(mDepStockBag - mWidStockBag, "#####0")
End Function

Private Function OpeningTransferIssued() As Variant
Dim mDepStock, mWidStock As Variant
Dim mDepStockBag, mWidStockBag As Variant

tmp = "Select GSLTransferDate,a.TxnType,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTF_DetailsID  = b.CxTF_DetailsID and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.Mst_GSL_ID = b.Mst_GSL_ID"
tmp = tmp & " inner join Txn_GSL_Transfer d on  c.SM_Prefix =d.SM_Prefix and c.Mst_GSL_ID = d.Old_Mst_GSL_ID"
tmp = tmp & " where a.exchangetypeid =1 And d.GSLTransferDate<'" & TxtFromDate.Value & "' "
tmp = tmp & " And c.Mst_GSL_ID in ( select Mst_GSL_ID from  mst_gsl C where c.exchangetypeid =1 "

If CmbGodownID.Text <> "" Then
   tmp = tmp & " And c.GodownID=" & mGodownID
Else
   If CmbCMPName.Text <> "" Then
      tmp = tmp & " And c.CMPID=" & mCMPID
   End If
End If

tmp = tmp & " And C.FGSLID_Mst_GSL_ID is not null c.SM_Prefix='" & mSM_Prefix & "'" & wc & ")"

If CmbGodownID.Text <> "" Then
   tmp = tmp & " And c.GodownID<>" & mGodownID
Else
   If CmbCMPName.Text <> "" Then
      tmp = tmp & " And c.CMPID<>" & mCMPID
   End If
End If

tmp = tmp & " And c.SM_Prefix='" & mSM_Prefix & "'"
tmp = tmp & " " & wc
tmp = tmp & " group by d.GSLTransferDate,a.TxnType"

Call Tmp5Table

mStock = 0
mWidStock = 0
mDepStock = 0

mStockBag = 0
mWidStockBag = 0
mDepStockBag = 0

If TmpRs5.RecordCount > 0 Then
   For I = 1 To TmpRs5.RecordCount
       If TmpRs5!TxnType = "D" Then
          mDepStock = TmpRs5!GrossWeight
          mDepStockBag = TmpRs5!NoofBags
       Else
          mWidStock = TmpRs5!GrossWeight
          mWidStockBag = TmpRs5!NoofBags
       End If
       TmpRs5.MoveNext
   Next
End If
mOpTransWidWt = Format(mDepStock - mWidStock, "#####0.000")
mOpTransWidBag = Format(mDepStockBag - mWidStockBag, "#####0")
End Function

Private Function TransferReceived(mDate_ As Variant) As Variant
Dim mDepStock, mWidStock As Variant
Dim mGSLIDs As String
Dim mDepStockBag, mWidStockBag As Variant

mGSLIDs = GetBGSLIDs()

tmp = "Select GSLTransferDate,a.TxnType,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTF_DetailsID  = b.CxTF_DetailsID and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.Mst_GSL_ID = b.Mst_GSL_ID"
tmp = tmp & " inner join Txn_GSL_Transfer d on  c.SM_Prefix =d.SM_Prefix and c.Mst_GSL_ID = d.Old_Mst_GSL_ID"
tmp = tmp & " where a.exchangetypeid =1 And d.GSLTransferDate>='" & mDate_ & "' And d.GSLTransferDate<='" & mDate_ & "'"
tmp = tmp & " And c.Mst_GSL_ID in (" & mGSLIDs & ")"

If CmbGodownID.Text <> "" Then
   tmp = tmp & " And c.GodownID<>" & mGodownID
Else
   If CmbCMPName.Text <> "" Then
      tmp = tmp & " And c.CMPID<>" & mCMPID
   End If
End If

tmp = tmp & " And c.SM_Prefix='" & mSM_Prefix & "'"
tmp = tmp & " " & wc
tmp = tmp & " group by d.GSLTransferDate,a.TxnType"

Call Tmp5Table

mStock = 0
mWidStock = 0
mDepStock = 0
mStockBag = 0
mWidStockBag = 0
mDepStockBag = 0

If TmpRs5.RecordCount > 0 Then
   For I = 1 To TmpRs5.RecordCount
       If TmpRs5!TxnType = "D" Then
          mDepStock = TmpRs5!GrossWeight
          mDepStockBag = TmpRs5!NoofBags
       Else
          mWidStock = TmpRs5!GrossWeight
          mWidStockBag = TmpRs5!NoofBags
       End If
       TmpRs5.MoveNext
   Next
End If
mStock = Format(mDepStock - mWidStock, "#####0.000")
mStockBag = Format(mDepStockBag - mWidStockBag, "#####0")

End Function

Private Function TransferIssued(mDate_ As Variant) As Variant

Dim mDepStock, mWidStock As Variant
Dim mDepStockBag, mWidStockBag As Variant

tmp = "Select GSLTransferDate,a.TxnType,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTF_DetailsID  = b.CxTF_DetailsID and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.Mst_GSL_ID = b.Mst_GSL_ID"
tmp = tmp & " inner join Txn_GSL_Transfer d on  c.SM_Prefix =d.SM_Prefix and c.Mst_GSL_ID = d.Old_Mst_GSL_ID"
tmp = tmp & " inner join mst_gsl c1 on  c1.SM_Prefix = d.SM_Prefix and c1.Mst_GSL_ID = d.newgslid"
tmp = tmp & " where a.exchangetypeid =1 And d.GSLTransferDate>='" & mDate_ & "' And d.GSLTransferDate<='" & mDate_ & "'"
tmp = tmp & " And c.Mst_GSL_ID in ( select Mst_GSL_ID from  mst_gsl C where c.exchangetypeid =1 "

If CmbGodownID.Text <> "" Then
   tmp = tmp & " And c.GodownID=" & mGodownID
Else
   If CmbCMPName.Text <> "" Then
      tmp = tmp & " And c.CMPID=" & mCMPID
   End If
End If

tmp = tmp & " And C.FGSLID_Mst_GSL_ID is not null And c.SM_Prefix='" & mSM_Prefix & "'" & wc & ")"

If CmbGodownID.Text <> "" Then
   tmp = tmp & " And c1.GodownID<>" & mGodownID
Else
   If CmbCMPName.Text <> "" Then
      tmp = tmp & " And c1.CMPID<>" & mCMPID
   End If
End If

tmp = tmp & " And c.SM_Prefix='" & mSM_Prefix & "'"
tmp = tmp & " " & wc

tmp = tmp & " group by d.GSLTransferDate,a.TxnType"

Call Tmp5Table

mStock = 0
mWidStock = 0
mDepStock = 0

mStockBag = 0
mWidStockBag = 0
mDepStockBag = 0

If TmpRs5.RecordCount > 0 Then
   For I = 1 To TmpRs5.RecordCount
       If TmpRs5!TxnType = "D" Then
          mDepStock = TmpRs5!GrossWeight
          mDepStockBag = TmpRs5!NoofBags
       Else
          mWidStock = TmpRs5!GrossWeight
          mWidStockBag = TmpRs5!NoofBags
       End If
       TmpRs5.MoveNext
   Next
End If
mStock = Format(mDepStock - mWidStock, "#####0.000")
mStockBag = Format(mDepStockBag - mWidStockBag, "#####0")
End Function
Private Function GetBGSLIDs() As String

mBGSLIDs = ""

tmp1 = "select distinct Mst_GSL_ID from  mst_gsl C where c.exchangetypeid =1 And SM_Prefix='" & mSM_Prefix & "' "
If CmbCMPName.Text <> "" Then
   tmp1 = tmp1 & " And c.CMPID=" & mCMPID
End If
If CmbGodownID.Text <> "" Then
   tmp1 = tmp1 & " And c.GodownID=" & mGodownID
End If
tmp1 = tmp1 & " " & wc

Call GetBGSLID(tmp1)

If mBGSLIDs = "" Then mBGSLIDs = 0
GetBGSLIDs = mBGSLIDs

End Function

Private Sub GetBGSLID(mGSLID_ As Variant)
Dim mTempBGSLIDs As String

tmp = "select distinct C.BGSLID_Mst_GSL_ID from  mst_gsl C where C.Mst_GSL_ID in (" & mGSLID_ & ") And SM_Prefix='" & mSM_Prefix & "'"

Call Tmp6Table

mTempBGSLIDs = ""
If TmpRs6.RecordCount > 0 Then
   For I = 1 To TmpRs6.RecordCount
       If IsNull(TmpRs6!BGSLID_Mst_GSL_ID) = False Then
          If mTempBGSLIDs = "" Then
             mTempBGSLIDs = TmpRs6!BGSLID_Mst_GSL_ID
          Else
             mTempBGSLIDs = mTempBGSLIDs & "," & TmpRs6!BGSLID_Mst_GSL_ID
          End If
       End If
       TmpRs6.MoveNext
   Next
   If mTempBGSLIDs <> "" Then
      If mBGSLIDs = "" Then
         mBGSLIDs = mTempBGSLIDs
      Else
         mBGSLIDs = mBGSLIDs & "," & mTempBGSLIDs
      End If
      Call GetBGSLID(mTempBGSLIDs)
   End If
End If
End Sub

