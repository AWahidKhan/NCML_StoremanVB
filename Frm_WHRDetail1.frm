VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_WHRDetail 
   Caption         =   "WHR Detail"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   14460
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   765
      Left            =   120
      TabIndex        =   20
      Top             =   10080
      Visible         =   0   'False
      Width           =   1005
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   21
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   22
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   29
      Top             =   8280
      Width           =   14205
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11145
         TabIndex        =   18
         Top             =   600
         Width           =   2580
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11145
         TabIndex        =   17
         Top             =   195
         Width           =   2580
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   8445
         TabIndex        =   35
         Top             =   600
         Width           =   2700
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   8445
         TabIndex        =   34
         Top             =   195
         Width           =   2700
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   5745
         TabIndex        =   33
         Top             =   600
         Width           =   2700
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   5745
         TabIndex        =   32
         Top             =   195
         Width           =   2700
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2925
         TabIndex        =   31
         Top             =   600
         Width           =   2820
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2925
         TabIndex        =   16
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   30
         Top             =   600
         Width           =   2820
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   0
         Top             =   195
         Width           =   2820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8265
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   14205
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1290
         Left            =   105
         TabIndex        =   51
         Top             =   3480
         Width           =   13965
         Begin VB.TextBox TxtWHRNoOfBags 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   10440
            MaxLength       =   50
            TabIndex        =   13
            Top             =   405
            Width           =   1710
         End
         Begin VB.CommandButton CmdUpdate 
            Caption         =   "Update List"
            Height          =   360
            Left            =   12150
            TabIndex        =   15
            Top             =   840
            Width           =   1680
         End
         Begin VB.TextBox TxtWHRQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   12165
            MaxLength       =   50
            TabIndex        =   14
            Top             =   405
            Width           =   1665
         End
         Begin VB.TextBox TxtEleWeight 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   8745
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   405
            Width           =   1680
         End
         Begin VB.TextBox TxtLicenseDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   405
            Width           =   5280
         End
         Begin VB.TextBox TxtLicenseNo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   60
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   405
            Width           =   1515
         End
         Begin VB.TextBox TxtNoOfBags 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   6915
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   405
            Width           =   1800
         End
         Begin VB.Label Label16 
            Caption         =   "WHR No Of Bags"
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
            Left            =   10560
            TabIndex        =   61
            Top             =   120
            Width           =   1620
         End
         Begin VB.Label Label14 
            Caption         =   "WHR Qty"
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
            Left            =   12622
            TabIndex        =   60
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label13 
            Caption         =   "Eletronic Weight"
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
            Left            =   8835
            TabIndex        =   59
            Top             =   105
            Width           =   1500
         End
         Begin VB.Label Label12 
            Caption         =   "License Description"
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
            Left            =   3315
            TabIndex        =   58
            Top             =   135
            Width           =   1860
         End
         Begin VB.Label Label11 
            Caption         =   "License No"
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
            Left            =   307
            TabIndex        =   57
            Top             =   150
            Width           =   1020
         End
         Begin VB.Label Label15 
            Caption         =   "No Of Bags"
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
            Left            =   7275
            TabIndex        =   56
            Top             =   120
            Width           =   1080
         End
      End
      Begin VB.TextBox TxtAssayerCerti 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2640
         Width           =   6105
      End
      Begin VB.TextBox TxtStorageAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3045
         Width           =   1680
      End
      Begin VB.TextBox TxtStorageCharges 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   12
         Top             =   3045
         Width           =   1800
      End
      Begin VB.TextBox TxtAssayer 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2235
         Width           =   9240
      End
      Begin VB.TextBox TxtGrade 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   9720
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
         Width           =   3120
      End
      Begin VB.ComboBox CmbCommodityCondition 
         Height          =   315
         ItemData        =   "Frm_WHRDetail.frx":0000
         Left            =   3600
         List            =   "Frm_WHRDetail.frx":000D
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtWeigher 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1815
         Width           =   9240
      End
      Begin VB.ComboBox CmbClientId 
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Top             =   1080
         Width           =   9255
      End
      Begin VB.TextBox TxtTotalWHRQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3045
         Width           =   1425
      End
      Begin VB.ComboBox CmbWHRCommodity 
         Height          =   315
         Left            =   3600
         TabIndex        =   7
         Top             =   1440
         Width           =   9255
      End
      Begin VB.TextBox TxtWHRNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker WHRDate 
         Height          =   330
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         _Version        =   393216
         Format          =   57081857
         CurrentDate     =   39884
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3345
         Left            =   105
         TabIndex        =   19
         Top             =   4800
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   5900
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker GradeValidUpto 
         Height          =   375
         Left            =   11400
         TabIndex        =   11
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   57081857
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker StorageFrom 
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57081857
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker StorageTo 
         Height          =   360
         Left            =   7680
         TabIndex        =   4
         Top             =   600
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   635
         _Version        =   393216
         Format          =   57081857
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Top             =   1080
         Width           =   9240
      End
      Begin VB.TextBox TxtWHRCommodityId 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         Top             =   1440
         Width           =   9240
      End
      Begin VB.Label Label5 
         Caption         =   "Grade as per Assayer’s Certificate"
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
         Left            =   360
         TabIndex        =   48
         Top             =   2670
         Width           =   3060
      End
      Begin VB.Label LblWHRCommodity 
         Alignment       =   1  'Right Justify
         Caption         =   "Commodity"
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
         Left            =   120
         TabIndex        =   47
         Top             =   1440
         Width           =   3300
      End
      Begin VB.Label LblClientName 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   3300
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name and License No. of Weigher"
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
         TabIndex        =   45
         Top             =   1845
         Width           =   3420
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Name and License No. of Assayer"
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
         TabIndex        =   44
         Top             =   2265
         Width           =   3300
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Storage Charge (Rate)"
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
         Left            =   120
         TabIndex        =   43
         Top             =   3075
         Width           =   3300
      End
      Begin VB.Label Label9 
         Caption         =   "Storage From"
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
         Left            =   6120
         TabIndex        =   42
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label10 
         Caption         =   "Storage To"
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
         Left            =   7830
         TabIndex        =   41
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label8 
         Caption         =   "Storage Charge Amount"
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
         Left            =   5760
         TabIndex        =   40
         Top             =   3075
         Width           =   2160
      End
      Begin VB.Label Label6 
         Caption         =   "Grade Valid Till"
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
         Left            =   9825
         TabIndex        =   38
         Top             =   2655
         Width           =   1470
      End
      Begin VB.Label Label3 
         Caption         =   "Grade"
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
         Left            =   10920
         TabIndex        =   37
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Condition of Commodities"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3630
         TabIndex        =   36
         Top             =   255
         Width           =   2220
      End
      Begin VB.Label LblWHRDate 
         Caption         =   "WHR Date"
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
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label LblTotalWHRQty 
         Caption         =   "Total WHR Qty"
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
         Left            =   9870
         TabIndex        =   26
         Top             =   3082
         Width           =   1500
      End
      Begin VB.Label LblWHRNo 
         Caption         =   "WHR No"
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
         Left            =   600
         TabIndex        =   25
         Top             =   240
         Width           =   900
      End
   End
End
Attribute VB_Name = "Frm_WHRDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID, mCommodityCondition, mWHRNoOfBags, mWHRQty As Double
Dim pq, temp, GSLQty, ChkSelection As Double
Dim mClientIDRow As Double
Dim mCommodityID As Double
Dim Edit_Flg, mLicenseNo  As Variant
Private Sub AddCmd_Click()
  
  Edit_Flg = "A"
  WHRDate.Enabled = True
  CmbCommodityCondition.Locked = False
  StorageFrom.Enabled = True
  StorageTo.Enabled = True
  TxtGrade.Locked = False
  TxtClientID.Visible = False
  CmbClientId.Visible = True
  CmbClientId.Locked = False
  TxtWHRCommodityId.Visible = False
  CmbWHRCommodity.Visible = True
  CmbWHRCommodity.Locked = False
  TxtWeigher.Locked = False
  TxtAssayer.Locked = False
  TxtAssayerCerti.Locked = False
  GradeValidUpto.Enabled = True
  TxtStorageCharges.Locked = False
  TxtWHRNoOfBags.Locked = True
  TxtWHRQty.Locked = True
  TxtWHRNo = ""
  CmbCommodityCondition.Text = ""
  mCommodityID = 0
  mClientIDRow = 0
  WHRDate.Value = Date
  StorageFrom.Value = Date
  StorageTo.Value = Date
  TxtGrade = ""
  CmbClientId.Text = ""
  CmbWHRCommodity.Text = ""
  TxtWeigher = ""
  TxtAssayer = ""
  TxtAssayerCerti = ""
  GradeValidUpto = Date
  TxtStorageCharges = ""
  TxtStorageAmount = ""
  TxtTotalWHRQty = ""
  ChkSelection = 0
  Call Grid_Head
  CmdUpdate.Enabled = False
  Call GButtonLock(Me, False)
  
End Sub

Private Sub CmbClientID_GotFocus()
tmp = CmbClientId.Text
mExchangeTypeID = GetExchangeID("NON NCDEX")
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
CmbClientId.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
    CmbClientId.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClientId.Text = tmp
End Sub
Private Sub CmbClientID_LostFocus()
If CmbClientId.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClientId.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid  selection "
   CmbClientId.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow

Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow)


End Sub

Private Sub CmbCommodityCondition_LostFocus()
mCommodityCondition = CmbCommodityCondition.ListIndex
End Sub

Private Sub CmbWHRCommodity_GotFocus()
tmp = CmbWHRCommodity.Text
Call TableMst_Commodity
CmbWHRCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
    CmbWHRCommodity.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbWHRCommodity.Text = tmp
End Sub
'----
Private Sub CmbWHRCommodity_LostFocus()
If CmbWHRCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbWHRCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid  selection "
   CmbWHRCommodity.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID

Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow)

End Sub
Private Sub GetGridData(mExchangeTypeID_, mCommodityID_, mClientIDRow_)
Call Grid_Head
If mExchangeTypeID_ = "" Then Exit Sub
If mCommodityID_ = "" Or mCommodityID_ = 0 Then Exit Sub
If mClientIDRow_ = "" Then Exit Sub

tmp = "SELECT Sum(Mst_GSL.BalanceBags) AS BalanceBags, Sum(Mst_GSL.BalanceWeight) AS BalanceWeight, " & _
      " Mst_License.LicenseNo , Mst_License.LicenseDesc, Mst_Godown.CMPID From Mst_License " & _
      " INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) " & _
      " ON Mst_License.LicenseID = Mst_Godown.LicenseID Where ((Mst_GSL.ClientIDRow = " & mClientIDRow & ") And " & _
      " (Mst_GSL.ExchangeTypeID = " & mExchangeTypeID_ & ") And (Mst_GSL.CommodityID = " & mCommodityID_ & ")) " & _
      " GROUP BY Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_Godown.CMPID"
Call TmpTable

If TmpRs.RecordCount = 0 Then
  MsgBox "No Record For Grid"
  Exit Sub
End If

k = 1
For i = 0 To TmpRs.RecordCount - 1
    
  If TmpRs!BalanceBags <> 0 And TmpRs!BalanceWeight <> 0 Then
     k = MSHFlexGrid2.Rows - 1
     MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
     MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!LicenseNo
     MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!LicenseDesc
     MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!BalanceBags
     MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!BalanceWeight
     MSHFlexGrid2.TextMatrix(k, 5) = ""
  End If
  TmpRs.MoveNext
Next

End Sub
Private Sub CmdUpdate_Click()

If TxtWHRNoOfBags.Text = 0 And TxtWHRQty = 0 Then
    ChkSelection = 0
    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
Else
    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
    ChkSelection = 1
End If

MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = TxtLicenseNo
mLicenseNo = TxtLicenseNo
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) = TxtLicenseDescription
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3) = Val(TxtNoOfBags) - Val(TxtWHRNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4) = Val(TxtEleWeight) - Val(TxtWHRQty)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5) = Val(TxtWHRNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6) = Val(TxtWHRQty)
TxtLicenseNo = ""
TxtLicenseDescription = ""
TxtNoOfBags = ""
TxtEleWeight = ""
TxtWHRNoOfBags = ""
TxtWHRQty = ""



CmdUpdate.Enabled = False
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
mLicenseNo = GetLicenseNo(RsTxn_WHRDetail!LicenseID)
ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If ans = vbYes Then
    Call TableMst_GSL
     tmp = "SELECT Mst_GSL.GSLID,Mst_GSL.BalanceBags, Mst_GSL.BalanceWeight,Mst_GSL.DepositeBags,Mst_GSL.DepositeWeight,Mst_License.LicenseNo , " & _
           "  Mst_GSL!WHRBags,Mst_GSL!WHRWeight,Mst_License.LicenseDesc, Mst_Godown.CMPID,Mst_GSL.WHRBags,Mst_GSL.WHRWeight From Mst_License " & _
           " INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) " & _
           " ON Mst_License.LicenseID = Mst_Godown.LicenseID Where ((Mst_GSL.ClientIDRow = " & mClientIDRow & ") And " & _
           " (Mst_GSL.ExchangeTypeID = " & mExchangeTypeID & ") And " & _
           " (Mst_GSL.CommodityID = " & mCommodityID & ") " & _
           " AND Mst_License.LicenseNo = '" & mLicenseNo & "') AND ((Mst_GSL.BalanceBags < Mst_GSL.DepositeBags) Or (Mst_GSL.BalanceWeight < Mst_GSL.DepositeWeight))"
     Call Tmp1Table

     TBags = RsTxn_WHRDetail!WHRNoOfBags
     TQty = RsTxn_WHRDetail!WHRQuantity
     For i = 1 To TmpRs1.RecordCount
         If TBags >= TmpRs1!WHRBags Then
            PBags = TmpRs1!WHRBags
            TBags = TBags - PBags
         Else
            PBags = TBags
            TBags = 0
         End If
         TmpRs1!BalanceBags = IIf(IsNull(TmpRs1!BalanceBags), 0, TmpRs1!BalanceBags) + PBags
         TmpRs1!WHRBags = TmpRs1!WHRBags - PBags
         If TBags = 0 Then Exit For
         TmpRs1.MoveNext
     Next
     TmpRs1.MoveFirst
     'TQty = WHRQty_
     For i = 1 To TmpRs1.RecordCount
         If TQty >= TmpRs1!WHRWeight Then
            PQty = TmpRs1!WHRWeight
            TQty = TQty - PQty
         Else
            PQty = TQty
            TQty = 0
         End If
         TmpRs1!BalanceWeight = IIf(IsNull(TmpRs1!BalanceWeight), 0, TmpRs1!BalanceWeight) + PQty
         TmpRs1!WHRWeight = TmpRs1!WHRWeight - PQty
         If TQty = 0 Then Exit For
         TmpRs1.MoveNext
     Next

     TmpRs1.Update
     RsMst_GSL.Update

    RsTxn_WHRDetail.Delete
    RsTxn_WHRDetail.Update
    
    If RsTxn_WHRDetail.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
    End If
    
    RsTxn_WHRDetail.MoveNext
    If RsTxn_WHRDetail.EOF() Then
      RsTxn_WHRDetail.MoveLast
    End If
    Call InData

 
End If
Exit Sub
msgError:
MsgBox "Child Record Present "

End Sub

Private Sub EditCmd_Click()
  Edit_Flg = "E"
  
  CmbCommodityCondition.Locked = False
  StorageFrom.Enabled = True
  StorageTo.Enabled = True
  TxtGrade.Locked = False
  TxtClientID.Visible = False
  CmbClientId.Visible = True
  CmbClientId.Locked = False
  TxtWHRCommodityId.Visible = False
  CmbWHRCommodity.Visible = True
  CmbWHRCommodity.Locked = False
  TxtWeigher.Locked = False
  TxtAssayer.Locked = False
  TxtAssayerCerti.Locked = False
  GradeValidUpto.Enabled = True
  TxtStorageCharges.Locked = False
  TxtWHRNoOfBags.Locked = True
  TxtWHRQty.Locked = True
  
  Call GButtonLock(Me, False)
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Grid_Head
Call TableTxn_WHRDetail("")
If RsTxn_WHRDetail.RecordCount = 0 Then
  If UserType = "Auditor" Then
     Call GButtonLock(Me, False)
     SaveCmd.Enabled = False
     SearchCmd.Enabled = False
     Exit Sub
  End If
  Call AddCmd_Click
  Exit Sub
End If

Call InData
End Sub
Private Sub InData()
  WHRDate.Enabled = False
  CmbCommodityCondition.Locked = True
  StorageFrom.Enabled = False
  StorageTo.Enabled = False
  TxtGrade.Locked = True
  TxtClientID.Visible = True
  CmbClientId.Visible = False
  CmbClientId.Locked = True
  TxtWHRCommodityId.Visible = True
  CmbWHRCommodity.Visible = False
  CmbWHRCommodity.Locked = True
  TxtWeigher.Locked = True
  TxtAssayer.Locked = True
  TxtAssayerCerti.Locked = True
  GradeValidUpto.Enabled = False
  TxtStorageCharges.Locked = True
  TxtWHRNoOfBags.Locked = True
  TxtWHRQty.Locked = True
  CmdUpdate.Enabled = False
  Call GButtonLock(Me, True)
  mExchangeTypeID = GetExchangeID("NON NCDEX")
  Call Grid_Head
  TxtWHRNo = IIf(IsNull(RsTxn_WHRDetail!WHRNo), "", RsTxn_WHRDetail!WHRNo)
  WHRDate.Value = IIf(IsNull(RsTxn_WHRDetail!WHRDate), "", RsTxn_WHRDetail!WHRDate)
  mCommodityCondition = IIf(IsNull(RsTxn_WHRDetail!ConditionId), "", RsTxn_WHRDetail!ConditionId)
  If mCommodityCondition = 0 Then
     CmbCommodityCondition.Text = "Average"
  ElseIf mCommodityCondition = 1 Then
     CmbCommodityCondition.Text = "Fair"
  ElseIf mCommodityCondition = 2 Then
     CmbCommodityCondition.Text = "Good"
  End If
  StorageFrom.Value = IIf(IsNull(RsTxn_WHRDetail!StorageFrom), "", RsTxn_WHRDetail!StorageFrom)
  StorageTo.Value = IIf(IsNull(RsTxn_WHRDetail!StorageTo), "", RsTxn_WHRDetail!StorageTo)
  TxtGrade = IIf(IsNull(RsTxn_WHRDetail!Grade), "", RsTxn_WHRDetail!Grade)
  mClientIDRow = IIf(IsNull(RsTxn_WHRDetail!ClientIDRow), "", RsTxn_WHRDetail!ClientIDRow)
  ans = GetClientNameByRow(mClientIDRow)
  
  If ans(0) = "" Then
     MsgBox "Not Found!!!"
     Exit Sub
  End If
  TxtClientID.Text = ans(1)
  mCommodityID = IIf(IsNull(RsTxn_WHRDetail!CommodityID), "", RsTxn_WHRDetail!CommodityID)
  TxtWHRCommodityId.Text = GetCommodityName(mCommodityID)
  TxtWeigher = IIf(IsNull(RsTxn_WHRDetail!Weigher), "", RsTxn_WHRDetail!Weigher)
  TxtAssayer = IIf(IsNull(RsTxn_WHRDetail!Assayer), "", RsTxn_WHRDetail!Assayer)
  TxtAssayerCerti = IIf(IsNull(RsTxn_WHRDetail!GradeByAssayer), "", RsTxn_WHRDetail!GradeByAssayer)
  GradeValidUpto.Value = IIf(IsNull(RsTxn_WHRDetail!GradeValidTill), "", RsTxn_WHRDetail!GradeValidTill)
  TxtStorageCharges = IIf(IsNull(RsTxn_WHRDetail!StorageCharges), "", RsTxn_WHRDetail!StorageCharges)
  TxtStorageAmount = IIf(IsNull(RsTxn_WHRDetail!StorageAmount), "", RsTxn_WHRDetail!StorageAmount)
  
  Call GetGridInData
  
  'Call GridInData
End Sub
Private Sub GetGridInData()
   
   'Call TableTxn_WHRLicenseDetail("Where WHRNo = " & TxtWHRNo)
  mLicenseID = IIf(IsNull(RsTxn_WHRDetail!LicenseID), "", RsTxn_WHRDetail!LicenseID)
  
   tmp = "SELECT sum(Mst_GSL.BalanceBags)  as DepositeBags, Sum(Mst_GSL.BalanceWeight) as DepositeWeight, " & _
         " Sum(Mst_GSL.WHRBags) as WHRBags ,Sum(Mst_GSL.WHRWeight) as WHRWeight, Sum(Mst_GSL.BalanceBags) as BalanceBags, " & _
         " Sum(Mst_GSL.BalanceWeight) As BalanceWeight ,Mst_License.LicenseNo , Mst_License.LicenseDesc, Mst_Godown.CMPID " & _
         " From Mst_License  INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) " & _
         " ON Mst_License.LicenseID = Mst_Godown.LicenseID " & _
         " Where ((Mst_GSL.ClientIDRow = " & mClientIDRow & ") And  (Mst_GSL.ISINID Is Null) And (Mst_GSL.ExchangeTypeID = " & mExchangeTypeID & ") And (Mst_GSL.CommodityID = " & mCommodityID & ") " & _
         " AND Mst_License.LicenseID = " & mLicenseID & " Or (Mst_GSL.ISINID = '')) " & _
         " GROUP BY Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_Godown.CMPID "
   Call Tmp1Table
   k = 1
   For i = 0 To TmpRs1.RecordCount - 1
     k = k + 1
     MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
     MSHFlexGrid2.TextMatrix(k, 1) = TmpRs1!LicenseNo
     MSHFlexGrid2.TextMatrix(k, 2) = TmpRs1!LicenseDesc
     MSHFlexGrid2.TextMatrix(k, 3) = TmpRs1!DepositeBags
     MSHFlexGrid2.TextMatrix(k, 4) = TmpRs1!DepositeWeight
     MSHFlexGrid2.TextMatrix(k, 5) = RsTxn_WHRDetail!WHRNoOfBags
     MSHFlexGrid2.TextMatrix(k, 6) = RsTxn_WHRDetail!WHRQuantity
   Next
   TmpRs1.Close
End Sub
Public Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 8
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "   "
MSHFlexGrid2.TextMatrix(1, 1) = "Lincese No"
MSHFlexGrid2.TextMatrix(1, 2) = "Lincese Description "
MSHFlexGrid2.TextMatrix(1, 3) = "No of Bags"
MSHFlexGrid2.TextMatrix(1, 4) = "Electronic Weight"
MSHFlexGrid2.TextMatrix(1, 5) = "WHR No Of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "WHR Quantity"
MSHFlexGrid2.TextMatrix(1, 7) = "WHR LicenseDetailID"

MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 4500
MSHFlexGrid2.ColWidth(3) = 1400
MSHFlexGrid2.ColWidth(4) = 1800
MSHFlexGrid2.ColWidth(5) = 1600
MSHFlexGrid2.ColWidth(6) = 1800
MSHFlexGrid2.ColWidth(7) = 0
'MSHFlexGrid2.RowHeight(1) = 400
End Sub
Private Sub MSHFlexGrid2_Click()

If SaveCmd.Enabled = False Then Exit Sub

If MSHFlexGrid2.Rows > 3 Then
    If MSHFlexGrid2.RowSel <> "0" Then
       If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then Exit Sub
       
       If ChkSelection = 1 Then
          If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) <> "*" Then
             MsgBox "Invalid Selection !!!!"
             Exit Sub
          End If
       End If
      TxtWHRNoOfBags.Locked = False
      TxtWHRQty.Locked = False
      TxtLicenseNo = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1)
      TxtLicenseDescription = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2)
      TxtNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5))
      TxtEleWeight = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6))
      TxtWHRNoOfBags = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5)
      TxtWHRQty = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6)
      TxtWHRNoOfBags.SetFocus
    End If
End If
End Sub

Private Sub SaveCmd_Click()

If CmbCommodityCondition.Text = "" Then
   MsgBox "Blank Commodity Condition Not Allowed!!!"
   CmbCommodityCondition.SetFocus
   Exit Sub
End If

If WHRDate.Value < StorageFrom.Value Then
   MsgBox "WHR Date must Be Greater Than Storage From Date!!!"
   WHRDate.SetFocus
   Exit Sub
End If


If StorageFrom.Value > StorageTo.Value Then
   MsgBox "Storage To Date must Be Greater Than Storage From Date!!!"
   StorageTo.SetFocus
   Exit Sub
End If

If TxtGrade = "" Then
   MsgBox "Blank Grade Not Allowed!!!"
   TxtGrade.SetFocus
   Exit Sub
End If

If CmbClientId.Text = "" Then
   MsgBox "Blank ClientName Not Allowed!!!"
   CmbClientId.SetFocus
   Exit Sub
End If

If CmbWHRCommodity.Text = "" Then
   MsgBox "Blank WHR Commodity Not Allowed!!!"
   CmbWHRCommodity.SetFocus
   Exit Sub
End If

If TxtStorageCharges.Text = "" Then
   MsgBox "Blank  Storage Charges Not Allowed!!!"
   TxtStorageCharges.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "Atleast One License Detail Should Be In Grid"
   MSHFlexGrid2.SetFocus
   Exit Sub
End If
If ChkSelection = 0 Then
   MsgBox "Select Atleast One License Detail from Grid"
   MSHFlexGrid2.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
'        If RsTxn_WHRDetail.RecordCount > 0 Then
'            RsTxn_WHRDetail.MoveFirst
'            RsTxn_WHRDetail.Find "WHRNo= '" & TxtWHRNo.Text & "'"
'            If Not RsTxn_WHRDetail.EOF Then
'                MsgBox "Duplicate WHR No Not Allowed !!! "
'                Exit Sub
'            End If
'        End If
    RsTxn_WHRDetail.AddNew
    RsTxn_WHRDetail!WHRNo = GetWHRNo
    TxtWHRNo.Text = RsTxn_WHRDetail!WHRNo
    Else
        RsTxn_WHRDetail.MoveFirst
        RsTxn_WHRDetail.Find "WHRNo= '" & TxtWHRNo.Text & "'"
        If Not RsTxn_WHRDetail.EOF Then
            If RsTxn_WHRDetail!WHRNo <> TxtWHRNo.Text Then
                MsgBox "Duplicate WHR No Not Allowed !!! "
                'TxtISINRevalReqNo.SetFocus
            Exit Sub
            End If
        End If
    RsTxn_WHRDetail.MoveFirst
    RsTxn_WHRDetail.Find " WHRNo= " & TxtWHRNo.Text
    End If
    RsTxn_WHRDetail!WHRDate = WHRDate.Value
    RsTxn_WHRDetail!ConditionId = mCommodityCondition
    RsTxn_WHRDetail!StorageFrom = StorageFrom.Value
    RsTxn_WHRDetail!StorageTo = StorageTo.Value
    RsTxn_WHRDetail!Grade = TxtGrade
    RsTxn_WHRDetail!Weigher = TxtWeigher
    RsTxn_WHRDetail!Assayer = TxtAssayer
    RsTxn_WHRDetail!GradeByAssayer = TxtAssayerCerti
    RsTxn_WHRDetail!GradeValidTill = GradeValidUpto.Value
    RsTxn_WHRDetail!StorageCharges = TxtStorageCharges
    RsTxn_WHRDetail!StorageAmount = "0" 'TxtStorageAmount
    RsTxn_WHRDetail!ClientIDRow = mClientIDRow
    RsTxn_WHRDetail!CommodityID = mCommodityID
    
    
    RsTxn_WHRDetail.Update
    'Call InData
    
For i = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 0) = "*" Then
       RsTxn_WHRDetail!LicenseID = GetLicenseID(MSHFlexGrid2.TextMatrix(i, 1))
       RsTxn_WHRDetail!WHRNoOfBags = MSHFlexGrid2.TextMatrix(i, 5)
       RsTxn_WHRDetail!WHRQuantity = MSHFlexGrid2.TextMatrix(i, 6)
       RsTxn_WHRDetail.Update
       Call PostGSLData(RsTxn_WHRDetail!LicenseID, RsTxn_WHRDetail!WHRNoOfBags, RsTxn_WHRDetail!WHRQuantity)
       Exit For
    End If
Next
Call InData
End Sub

Function PostGSLData(LicenseID_ As Double, WHRNoBags_ As Double, WHRQty_ As Double)
Dim TBags, TQty, PBags, PQty As Variant

Call TableMst_GSL
tmp = "SELECT Mst_GSL.GSLID,Mst_GSL.BalanceBags, Mst_GSL.BalanceWeight,Mst_License.LicenseNo , " & _
      "  Mst_License.LicenseDesc, Mst_Godown.CMPID,Mst_GSL.WHRBags,Mst_GSL.WHRWeight From Mst_License " & _
      " INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) " & _
      " ON Mst_License.LicenseID = Mst_Godown.LicenseID Where ((Mst_GSL.ClientIDRow = " & mClientIDRow & ") And " & _
      " (Mst_GSL.ExchangeTypeID = " & mExchangeTypeID & ") And " & _
      " (Mst_GSL.CommodityID = " & mCommodityID & ") " & _
      " AND Mst_License.LicenseNo = '" & mLicenseNo & "') And (Mst_GSL.BalanceBags > 0 OR Mst_GSL.BalanceWeight > 0)  "
Call Tmp1Table

TBags = WHRNoBags_
TQty = WHRQty_
For i = 1 To TmpRs1.RecordCount
    If TBags >= TmpRs1!BalanceBags Then
       PBags = TmpRs1!BalanceBags
       TBags = TBags - PBags
    Else
       PBags = TBags
       TBags = 0
    End If
    TmpRs1!WHRBags = IIf(IsNull(TmpRs1!WHRBags), 0, TmpRs1!WHRBags) + PBags
    TmpRs1!BalanceBags = TmpRs1!BalanceBags - PBags
    If TBags = 0 Then Exit For
    TmpRs1.MoveNext
Next
TmpRs1.MoveFirst
TQty = WHRQty_
For i = 1 To TmpRs1.RecordCount
    If TQty >= TmpRs1!BalanceWeight Then
       PQty = TmpRs1!BalanceWeight
       TQty = TQty - PQty
    Else
       PQty = TQty
       TQty = 0
    End If
    TmpRs1!WHRWeight = IIf(IsNull(TmpRs1!WHRWeight), 0, TmpRs1!WHRWeight) + PQty
    TmpRs1!BalanceWeight = TmpRs1!BalanceWeight - PQty
    If PQty = 0 Then Exit For
    TmpRs1.MoveNext
Next

TmpRs1.Close


End Function
Private Function GetWHRNo() As Double
Dim RetVal As Double
tmp = "Select max(WHRNo) from Txn_WHR_Detail"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetWHRNo = RetVal
End Function
Private Function GetWHRLicenseDetailID() As Double
Dim RetVal As Double
tmp = "Select max(WHRLicenseDetailID) from Txn_WHRLicense_Detail"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetWHRLicenseDetailID = RetVal
End Function

Private Sub TxtStorageCharges_LostFocus()
If TxtStorageCharges = "" Then Exit Sub
If IsNumeric(TxtStorageCharges) = False Then
  MsgBox "Not Proper Quantity Entered!!!"
  TxtStorageCharges.SetFocus
  Exit Sub
End If
End Sub

Private Sub TxtWHRNoOfBags_LostFocus()
If TxtWHRNoOfBags = "" Then Exit Sub
If IsNumeric(TxtWHRNoOfBags) = False Then
  MsgBox "Not Proper Quantity Entered!!!"
  TxtWHRNoOfBags.SetFocus
  Exit Sub
End If
If Val(TxtWHRNoOfBags) > Val(TxtNoOfBags) Then
   MsgBox "WHR No Of Bags Should Be Less Than Electronic No Of Bags!!!"
   TxtWHRNoOfBags.SetFocus
   Exit Sub
End If
TxtWHRQty.SetFocus
End Sub

Private Sub TxtWHRQty_LostFocus()
If TxtWHRQty = "" Then Exit Sub
If IsNumeric(TxtWHRQty) = False Then
  MsgBox "Not Proper Quantity Entered!!!"
  TxtWHRQty.SetFocus
  Exit Sub
End If
If Val(TxtWHRQty) > Val(TxtEleWeight) Then
   MsgBox "WHR Weight Should Be Less Than Electronic Weight"
   TxtWHRQty.SetFocus
   Exit Sub
End If
CmdUpdate.Enabled = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_WHRDetail.MoveNext
If RsTxn_WHRDetail.EOF Then
   RsTxn_WHRDetail.MoveLast
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    NextCmd.Enabled = False
    LastCmd.Enabled = False
End If

PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_WHRDetail.MovePrevious
If RsTxn_WHRDetail.BOF Then
   RsTxn_WHRDetail.MoveFirst
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If

NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsTxn_WHRDetail.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_WHRDetail.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_WHRDetail.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
End Sub


