VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_SpecialRateCard 
   Caption         =   "Special Rate Card"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   14880
      Begin VB.ComboBox CmbSConClient 
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
         ItemData        =   "FrmMst_SpecialRateCard.frx":0000
         Left            =   3960
         List            =   "FrmMst_SpecialRateCard.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   113
         Top             =   1125
         Width           =   1410
      End
      Begin VB.CheckBox ChkUpdated 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9840
         TabIndex        =   106
         Top             =   810
         Width           =   210
      End
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmMst_SpecialRateCard.frx":0017
         Left            =   1020
         List            =   "FrmMst_SpecialRateCard.frx":0021
         Sorted          =   -1  'True
         TabIndex        =   103
         Top             =   1125
         Width           =   1410
      End
      Begin VB.TextBox TxtSearchClientName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   84
         Top             =   750
         Width           =   7920
      End
      Begin VB.ComboBox CmbSearchBaseRate 
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
         ItemData        =   "FrmMst_SpecialRateCard.frx":0038
         Left            =   11280
         List            =   "FrmMst_SpecialRateCard.frx":0042
         Sorted          =   -1  'True
         TabIndex        =   82
         Top             =   345
         Width           =   2280
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   350
         Left            =   13290
         TabIndex        =   77
         Top             =   1200
         Width           =   1000
      End
      Begin VB.ComboBox CmbSLocationID 
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
         TabIndex        =   42
         Top             =   345
         Width           =   3210
      End
      Begin VB.ComboBox CmbSCommodityID 
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
         Height          =   360
         Left            =   3345
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   345
         Width           =   3000
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   350
         Left            =   11280
         TabIndex        =   40
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12285
         TabIndex        =   39
         Top             =   1200
         Width           =   1000
      End
      Begin VB.ComboBox CmbSBillingUnit 
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
         ItemData        =   "FrmMst_SpecialRateCard.frx":004F
         Left            =   8970
         List            =   "FrmMst_SpecialRateCard.frx":005C
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   345
         Width           =   2265
      End
      Begin VB.ComboBox CmbSBillingCycle 
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
         Left            =   6405
         Sorted          =   -1  'True
         TabIndex        =   37
         Top             =   345
         Width           =   2520
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   855
         Left            =   120
         TabIndex        =   43
         Top             =   1590
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   1508
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker TxtUpdatedFrom 
         Height          =   375
         Left            =   11040
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   750
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47775745
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtUpdatedTo 
         Height          =   375
         Left            =   13425
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47775745
         CurrentDate     =   39884
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Converted Client"
         Height          =   195
         Left            =   2640
         TabIndex        =   114
         Top             =   1208
         Width           =   1170
      End
      Begin VB.Label Label29 
         Caption         =   "Updated From"
         Height          =   405
         Left            =   10320
         TabIndex        =   111
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "Updated To"
         Height          =   240
         Left            =   12540
         TabIndex        =   110
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Only Updated"
         Height          =   450
         Left            =   9120
         TabIndex        =   109
         Top             =   750
         Width           =   630
      End
      Begin VB.Label Label25 
         Caption         =   "Status"
         Height          =   225
         Left            =   120
         TabIndex        =   104
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label18 
         Caption         =   "Base Rate"
         Height          =   210
         Left            =   11970
         TabIndex        =   83
         Top             =   135
         Width           =   930
      End
      Begin VB.Label Label24 
         Caption         =   "Client Name "
         Height          =   255
         Left            =   75
         TabIndex        =   78
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Location"
         Height          =   210
         Left            =   1050
         TabIndex        =   49
         Top             =   165
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Commodity"
         Height          =   240
         Left            =   3525
         TabIndex        =   48
         Top             =   165
         Width           =   840
      End
      Begin VB.Label Label23 
         Caption         =   "Total Records Found :"
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
         Left            =   6480
         TabIndex        =   47
         Top             =   1290
         Width           =   2325
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
         Left            =   8880
         TabIndex        =   46
         Top             =   1290
         Width           =   90
      End
      Begin VB.Label Label15 
         Caption         =   "Billing Unit"
         Height          =   240
         Left            =   9720
         TabIndex        =   45
         Top             =   135
         Width           =   765
      End
      Begin VB.Label Label16 
         Caption         =   "Billing Cycle"
         Height          =   225
         Left            =   7305
         TabIndex        =   44
         Top             =   135
         Width           =   960
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1110
      Left            =   120
      TabIndex        =   70
      Top             =   8220
      Width           =   7575
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5970
         TabIndex        =   31
         Top             =   615
         Width           =   1260
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5970
         TabIndex        =   30
         Top             =   195
         Width           =   1260
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4590
         TabIndex        =   35
         Top             =   615
         Width           =   1380
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4590
         TabIndex        =   29
         Top             =   195
         Width           =   1380
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   3210
         TabIndex        =   34
         Top             =   615
         Width           =   1380
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   3210
         TabIndex        =   28
         Top             =   195
         Width           =   1380
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1830
         TabIndex        =   33
         Top             =   615
         Width           =   1380
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   450
         TabIndex        =   32
         Top             =   615
         Width           =   1380
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   450
         TabIndex        =   27
         Top             =   195
         Width           =   1380
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1830
         TabIndex        =   26
         Top             =   195
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4305
      Left            =   120
      TabIndex        =   50
      Top             =   0
      Width           =   7605
      Begin VB.CheckBox Check1 
         Caption         =   "Converted Client"
         Height          =   255
         Left            =   5040
         TabIndex        =   112
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtRemarks 
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
         Height          =   495
         Left            =   3600
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   3720
         Width           =   3855
      End
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmMst_SpecialRateCard.frx":0077
         Left            =   1230
         List            =   "FrmMst_SpecialRateCard.frx":0081
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   3720
         Width           =   1410
      End
      Begin VB.ComboBox CmbBillingUnit 
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
         ItemData        =   "FrmMst_SpecialRateCard.frx":0098
         Left            =   1200
         List            =   "FrmMst_SpecialRateCard.frx":00A5
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2295
         Width           =   3345
      End
      Begin VB.ComboBox CmbDiscount 
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
         ItemData        =   "FrmMst_SpecialRateCard.frx":00C0
         Left            =   3135
         List            =   "FrmMst_SpecialRateCard.frx":00CA
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   2745
         Width           =   1410
      End
      Begin VB.ComboBox CmbDepoBillingCycle 
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
         Left            =   1230
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   3270
         Width           =   3255
      End
      Begin VB.TextBox TxtDepoTotDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   5985
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   3270
         Width           =   1110
      End
      Begin VB.CommandButton CmdDepoBillingCycle 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4590
         TabIndex        =   97
         Top             =   3300
         Width           =   375
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   5880
      End
      Begin VB.TextBox TxtClientID 
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
         Height          =   360
         Left            =   1215
         TabIndex        =   80
         Top             =   630
         Width           =   5865
      End
      Begin VB.CommandButton CmdClientID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7140
         TabIndex        =   79
         Top             =   630
         Width           =   375
      End
      Begin VB.ComboBox CmbBaseRate 
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
         ItemData        =   "FrmMst_SpecialRateCard.frx":00E2
         Left            =   1200
         List            =   "FrmMst_SpecialRateCard.frx":00EC
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2745
         Width           =   975
      End
      Begin VB.TextBox TxtAmount 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   5970
         TabIndex        =   9
         Top             =   2745
         Width           =   1095
      End
      Begin VB.ComboBox CmbCommodityGroupID 
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
         Height          =   360
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   5880
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1035
         Width           =   5880
      End
      Begin VB.TextBox TxtSPRCID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   195
         Width           =   1215
      End
      Begin VB.CommandButton CmdLocationID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7125
         TabIndex        =   56
         Top             =   1065
         Width           =   375
      End
      Begin VB.CommandButton CmdCommodityGroupID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7125
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1455
         Width           =   375
      End
      Begin VB.TextBox TxtPreFix 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   5865
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   195
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdBillingCycle 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4575
         TabIndex        =   53
         Top             =   1890
         Width           =   375
      End
      Begin VB.ComboBox CmbBillingCycle 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1860
         Width           =   3345
      End
      Begin VB.TextBox TxtRateOnQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   5970
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1110
      End
      Begin VB.TextBox TxtTotDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   5970
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1110
      End
      Begin MSComCtl2.DTPicker TxtSPRCDate 
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47775745
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtLocationID 
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
         Height          =   360
         Left            =   1200
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1035
         Width           =   5880
      End
      Begin VB.TextBox TxtCommodityGroupID 
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
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1440
         Width           =   5880
      End
      Begin VB.TextBox TxtBillingCycle 
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
         Height          =   360
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1860
         Width           =   3330
      End
      Begin VB.TextBox TxtBillingUnit 
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
         Height          =   360
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2295
         Width           =   3330
      End
      Begin VB.TextBox TxtBaseRate 
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
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   73
         Top             =   2745
         Width           =   960
      End
      Begin VB.TextBox TxtDiscount 
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
         Height          =   375
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   2745
         Width           =   1410
      End
      Begin VB.TextBox TxtDepoBillingCycle 
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
         Height          =   360
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   3270
         Width           =   3210
      End
      Begin VB.TextBox TxtStatus 
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
         Height          =   375
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3720
         Width           =   1410
      End
      Begin VB.Label Label26 
         Caption         =   "Remarks"
         Height          =   210
         Left            =   2880
         TabIndex        =   105
         Top             =   3855
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   102
         Top             =   3795
         Width           =   870
      End
      Begin VB.Label Label21 
         Caption         =   "Total Days"
         Height          =   210
         Left            =   5145
         TabIndex        =   101
         Top             =   3345
         Width           =   780
      End
      Begin VB.Label Label10 
         Caption         =   "Deposit Billing Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   100
         Top             =   3270
         Width           =   1080
      End
      Begin VB.Label LblClientID 
         Caption         =   "Client Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   81
         Top             =   705
         Width           =   1185
      End
      Begin VB.Label Label20 
         Caption         =   "Actual Connected to BaseRate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   76
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Absolute Discount"
         Height          =   465
         Left            =   2385
         TabIndex        =   75
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label Label17 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5130
         TabIndex        =   74
         Top             =   2820
         Width           =   735
      End
      Begin VB.Label Label41 
         Caption         =   "Sp Rt Date"
         Height          =   225
         Left            =   2595
         TabIndex        =   69
         Top             =   270
         Width           =   855
      End
      Begin VB.Label LblWHRNo 
         Caption         =   "Sp Rate ID"
         Height          =   210
         Left            =   90
         TabIndex        =   68
         Top             =   277
         Width           =   900
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   67
         Top             =   1110
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Commodity Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   66
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label LblRatePerBag 
         Caption         =   "Rate on Units"
         Height          =   360
         Left            =   5130
         TabIndex        =   65
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label Label6 
         Caption         =   "SM_Prefix"
         Height          =   180
         Left            =   5130
         TabIndex        =   64
         Top             =   240
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label8 
         Caption         =   "Billing Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   105
         TabIndex        =   63
         Top             =   1830
         Width           =   960
      End
      Begin VB.Label Label11 
         Caption         =   "Billing Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   2355
         Width           =   1005
      End
      Begin VB.Label Label14 
         Caption         =   "Total Days"
         Height          =   210
         Left            =   5130
         TabIndex        =   61
         Top             =   1935
         Width           =   780
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3900
      Left            =   120
      TabIndex        =   71
      Top             =   4320
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6879
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Rate Details As per Bags"
      TabPicture(0)   =   "FrmMst_SpecialRateCard.frx":00F9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Deposit Rate Details As per Bags"
      TabPicture(1)   =   "FrmMst_SpecialRateCard.frx":0115
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "MSHFlexGrid3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Log History"
      TabPicture(2)   =   "FrmMst_SpecialRateCard.frx":0131
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtCreated"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "TxtUpdated"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label13"
      Tab(2).Control(3)=   "Label12"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   -74880
         TabIndex        =   93
         Top             =   330
         Width           =   7335
         Begin VB.CommandButton CmdDepoDeleteList 
            Caption         =   "Delete from List"
            Height          =   375
            Left            =   5670
            TabIndex        =   24
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton CmdDepoAdd2List 
            Caption         =   "Add to List"
            Height          =   375
            Left            =   4440
            TabIndex        =   23
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TxtDepoToBagSize 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   21
            Top             =   495
            Width           =   1215
         End
         Begin VB.TextBox TxtDepoFromBagSize 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   75
            MaxLength       =   100
            TabIndex        =   20
            Top             =   495
            Width           =   1215
         End
         Begin VB.TextBox TxtDepoRatePerBag 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   2565
            TabIndex        =   22
            Top             =   495
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "Bag Size To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1387
            TabIndex        =   96
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label7 
            Caption         =   "Bag Size From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   165
            TabIndex        =   95
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label1 
            Caption         =   "Rate Per Bag per Cycle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2580
            TabIndex        =   94
            Top             =   120
            Width           =   1860
         End
      End
      Begin VB.TextBox TxtCreated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   -73815
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   720
         Width           =   6255
      End
      Begin VB.TextBox TxtUpdated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   -73815
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1185
         Width           =   6255
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   85
         Top             =   360
         Width           =   7335
         Begin VB.TextBox TxtRatePerBag 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   2565
            TabIndex        =   16
            Top             =   495
            Width           =   1815
         End
         Begin VB.TextBox TxtFromBagSize 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   75
            MaxLength       =   100
            TabIndex        =   14
            Top             =   495
            Width           =   1215
         End
         Begin VB.TextBox TxtToBagSize 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   15
            Top             =   495
            Width           =   1215
         End
         Begin VB.CommandButton CmdAdd2List 
            Caption         =   "Add to List"
            Height          =   375
            Left            =   4440
            TabIndex        =   17
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton CmdDeleteList 
            Caption         =   "Delete from List"
            Height          =   375
            Left            =   5670
            TabIndex        =   18
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Rate Per Bag per Cycle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2655
            TabIndex        =   88
            Top             =   120
            Width           =   1860
         End
         Begin VB.Label LblBagSizeFrom 
            Caption         =   "Bag Size From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   165
            TabIndex        =   87
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label LblBagsSizeTo 
            Caption         =   "Bag Size To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1387
            TabIndex        =   86
            Top             =   120
            Width           =   1020
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2415
         Left            =   120
         TabIndex        =   19
         Top             =   1365
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4260
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   25
         Top             =   1365
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4260
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Created  :"
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
         Left            =   -74880
         TabIndex        =   92
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Updated :"
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
         Left            =   -74880
         TabIndex        =   91
         Top             =   1245
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmMst_SpecialRateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mFlag, mCommodityGroupId, mLocationID, mGodownID, mCMPID, mSRCID As Variant
Dim mSCommodityID, mSLocationID, mSExchangeTypeID, mBillingCycleID, mSBillingCycleID, mSFlag, mSearchBase, mBaseRate As Variant
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mDepoBillingCycleID As Variant
Dim mStatus, mSStatus As Variant


Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 18

MSHFlexGrid1.TextMatrix(0, 0) = "Sp Rate ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Sp Rate Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Location "
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity Group"
MSHFlexGrid1.TextMatrix(0, 4) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 5) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 6) = "Billing Cycle"
MSHFlexGrid1.TextMatrix(0, 7) = "Billing Unit"
MSHFlexGrid1.TextMatrix(0, 8) = "Rate on Units"
MSHFlexGrid1.TextMatrix(0, 9) = "Base Rate"
MSHFlexGrid1.TextMatrix(0, 10) = "Absolute Discount"
MSHFlexGrid1.TextMatrix(0, 11) = "Amount"
MSHFlexGrid1.TextMatrix(0, 12) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 13) = "Flag"
MSHFlexGrid1.TextMatrix(0, 14) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 15) = "Updated By"
MSHFlexGrid1.TextMatrix(0, 16) = "Updated Date"
MSHFlexGrid1.TextMatrix(0, 17) = "Converted Client"
MSHFlexGrid1.RowHeight(0) = 800

MSHFlexGrid1.ColWidth(0) = 600
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 2800
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 800
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 0
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.ColWidth(14) = 2000
MSHFlexGrid1.ColWidth(15) = 2000
MSHFlexGrid1.ColWidth(16) = 2000

MSHFlexGrid1.RowHeight(0) = 400
End Sub

Private Sub Check1_Click()
If Edit_Flg = "A" Then
   CmbClientID.Text = ""
End If
End Sub

Private Sub CmbBillingCycle_GotFocus()
tmp = CmbBillingCycle.Text
Call TableMst_BillingCycle
CmbBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_BillingCycle.RecordCount
   CmbBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
   RsMst_BillingCycle.MoveNext
Next
CmbBillingCycle.Text = tmp
End Sub

Private Sub CmbBillingCycle_LostFocus()
If CmbBillingCycle.Text = "" Then
   mBillingCycleID = Null
   Exit Sub
End If
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle = '" & CmbBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF Then
   MsgBox "Invalid selection "
   CmbBillingCycle.SetFocus
   Exit Sub
End If
mBillingCycleID = RsMst_BillingCycle!BillingCycleID
TxtTotDays = RsMst_BillingCycle!NoOfDays
Call GetStorageRateFromMaster
End Sub

Private Sub CmbSBillingCycle_GotFocus()
tmp = CmbSBillingCycle.Text
Call TableMst_BillingCycle
CmbSBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_BillingCycle.RecordCount
   CmbSBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
   RsMst_BillingCycle.MoveNext
Next
CmbSBillingCycle.Text = tmp
End Sub

Private Sub CmbSBillingCycle_LostFocus()
If CmbSBillingCycle.Text = "" Then
   mSBillingCycleID = Null
   Exit Sub
End If
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle = '" & CmbSBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF Then
   MsgBox "Invalid selection "
   CmbSBillingCycle.SetFocus
   Exit Sub
End If
mSBillingCycleID = RsMst_BillingCycle!BillingCycleID
End Sub
Private Sub CmbBillingUnit_LostFocus()
If CmbBillingUnit.Text = "" Then
   Exit Sub
End If

If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBillingUnit.Text) <> "quantity" And LCase(CmbBillingUnit.Text) <> "sq ft" Then
   MsgBox "Invalid Selection!!!"
   CmbBillingUnit.SetFocus
   Exit Sub
End If

If LCase(CmbBillingUnit.Text) = "bags" Then
   TxtRateOnQty = ""
   TxtRateOnQty.Enabled = False
   If MSHFlexGrid2.Rows <= 2 Then
      Call GetStorageRateFromMaster
   End If
   If MSHFlexGrid2.Rows <= 3 Then
      Call GetDepoStorageRateFromMaster
   End If
Else
   If MSHFlexGrid2.Rows > 2 Then
      If MsgBox("Data in Bag Size Details will be lost. Are you want to Continue.", vbYesNo) = vbNo Then
         CmbBillingUnit.SetFocus
         Exit Sub
      End If
   End If
   Call GetStorageRateFromMaster
   
   If MSHFlexGrid2.Rows > 3 Then
      If MsgBox("Data in Deposit Bag Size Details will be lost. Are you want to Continue.", vbYesNo) = vbNo Then
         CmbBillingUnit.SetFocus
         Exit Sub
      End If
   End If
   Call GetDepoStorageRateFromMaster
End If
End Sub
Private Sub CmbSBillingUnit_LostFocus()
If CmbSBillingUnit.Text = "" Then
   mSFlag = Null
   Exit Sub
End If
If LCase(CmbSBillingUnit.Text) <> "bags" And LCase(CmbSBillingUnit.Text) <> "quantity" And LCase(CmbSBillingUnit.Text) <> "sq ft" Then
   MsgBox "Invalid Selection!!!"
   CmbSBillingUnit.SetFocus
   Exit Sub
End If

mSFlag = Left(CmbSBillingUnit.Text, 1)
End Sub

Private Sub CmbSConClient_LostFocus()
If Trim(CmbSConClient.Text) = "" Then Exit Sub

If LCase(CmbSConClient.Text) <> "yes" And LCase(CmbSConClient.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSConClient.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSearchBaseRate_LostFocus()
If CmbSearchBaseRate.Text = "" Then
   mSearchDefault = Null
   Exit Sub
End If
If LCase(CmbSearchBaseRate.Text) <> "yes" And LCase(CmbSearchBaseRate.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSearchBaseRate.SetFocus
   Exit Sub
End If

If LCase(CmbSearchBaseRate.Text) = "yes" Then
   mSearchBase = "Y"
ElseIf LCase(CmbSearchBaseRate.Text) = "no" Then
   mSearchBase = "N"
End If
End Sub

Private Sub CmbSStatus_LostFocus()

If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "de-active" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If

mSStatus = Left(CmbSStatus.Text, 1)

End Sub

Private Sub CmbStatus_LostFocus()

If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "de-active" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

mStatus = Left(CmbStatus.Text, 1)
End Sub

Private Sub CmdAdd2List_Click()

If TxtFromBagSize.Text = "" Then
   MsgBox "Blank From Bag Size not Allowed!!!"
   TxtFromBagSize.SetFocus
   Exit Sub
End If

If TxtToBagSize.Text = "" Then
   MsgBox "Blank To Bag Size not Allowed!!!"
   TxtToBagSize.SetFocus
   Exit Sub
End If

If Val(TxtToBagSize.Text) < Val(TxtFromBagSize.Text) Then
   MsgBox "'To Bag Size' must be more than 'From Bag Size'"
   TxtToBagSize.SetFocus
   Exit Sub
End If

If TxtRatePerBag.Text = "" Then
   MsgBox "Blank Rate Per Bag not Allowed!!!"
   TxtRatePerBag.SetFocus
   Exit Sub
End If

If CheckDuplicateBags = True Then
   MsgBox "'Bag Size From' or 'Bag Size To' in already used in range. Please change range!!! "
   TxtFromBagSize.SetFocus
   Exit Sub
End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If

MSHFlexGrid2.TextMatrix(r, 0) = Val(TxtFromBagSize)
MSHFlexGrid2.TextMatrix(r, 1) = Val(TxtToBagSize)
MSHFlexGrid2.TextMatrix(r, 2) = Format(Val(TxtRatePerBag.Text), "#####0.0000")
MSHFlexGrid2.TextMatrix(r, 3) = "" ' Val(TxtFromBagSize)
MSHFlexGrid2.TextMatrix(r, 4) = "" 'Val(TxtFromBagSize)
Call ClearDetail
TxtFromBagSize.SetFocus
End Sub

Private Sub CmdDepoAdd2List_Click()

If TxtDepoFromBagSize.Text = "" Then
   MsgBox "Blank From Bag Size not Allowed!!!"
   TxtDepoFromBagSize.SetFocus
   Exit Sub
End If

If TxtDepoToBagSize.Text = "" Then
   MsgBox "Blank To Bag Size not Allowed!!!"
   TxtDepoToBagSize.SetFocus
   Exit Sub
End If

If Val(TxtDepoToBagSize.Text) < Val(TxtDepoFromBagSize.Text) Then
   MsgBox "'To Bag Size' must be more than 'From Bag Size'"
   TxtDepoToBagSize.SetFocus
   Exit Sub
End If

If TxtDepoRatePerBag.Text = "" Then
   MsgBox "Blank Rate Per Bag not Allowed!!!"
   TxtDepoRatePerBag.SetFocus
   Exit Sub
End If

If CheckDepoDuplicateBags = True Then
   MsgBox "'Bag Size From' or 'Bag Size To' in already used in range. Please change range!!! "
   TxtFromBagSize.SetFocus
   Exit Sub
End If

If CmdDepoAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid3.RowSel
Else
   r = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If

MSHFlexGrid3.TextMatrix(r, 0) = Val(TxtDepoFromBagSize)
MSHFlexGrid3.TextMatrix(r, 1) = Val(TxtDepoToBagSize)
MSHFlexGrid3.TextMatrix(r, 2) = Format(Val(TxtDepoRatePerBag.Text), "#####0.0000")
MSHFlexGrid3.TextMatrix(r, 3) = "" ' Val(TxtFromBagSize)
MSHFlexGrid3.TextMatrix(r, 4) = "" 'Val(TxtFromBagSize)
Call ClearDepoDetail
TxtFromBagSize.SetFocus

End Sub
Private Sub ClearDepoDetail()
TxtDepoFromBagSize = ""
TxtDepoToBagSize = ""
TxtDepoRatePerBag = ""
CmdDepoAdd2List.Caption = "Add to List"
End Sub

Private Sub ClearDetail()
TxtFromBagSize = ""
TxtToBagSize = ""
TxtRatePerBag = ""
CmdAdd2List.Caption = "Add to List"
End Sub
Private Sub CmdBillingCycle_Click()
FrmMst_BillingCycle.Show
End Sub

Function CheckDuplicateBags() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If (Val(TxtFromBagSize) >= Val(MSHFlexGrid2.TextMatrix(u, 0)) And Val(TxtFromBagSize) <= Val(MSHFlexGrid2.TextMatrix(u, 1))) Or Val(TxtToBagSize) >= Val(MSHFlexGrid2.TextMatrix(u, 0)) And Val(TxtToBagSize) <= Val(MSHFlexGrid2.TextMatrix(u, 1)) Or (Val(TxtFromBagSize) <= Val(MSHFlexGrid2.TextMatrix(u, 0)) And Val(TxtToBagSize) >= Val(MSHFlexGrid2.TextMatrix(u, 1))) Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateBags = Retval

End Function
Function CheckDepoDuplicateBags() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid3.Rows > 1 Then
   For u = 1 To MSHFlexGrid3.Rows - 1
       If (Val(TxtDepoFromBagSize) >= Val(MSHFlexGrid3.TextMatrix(u, 0)) And Val(TxtDepoFromBagSize) <= Val(MSHFlexGrid3.TextMatrix(u, 1))) Or Val(TxtDepoToBagSize) >= Val(MSHFlexGrid3.TextMatrix(u, 0)) And Val(TxtDepoToBagSize) <= Val(MSHFlexGrid3.TextMatrix(u, 1)) Or (Val(TxtDepoFromBagSize) <= Val(MSHFlexGrid3.TextMatrix(u, 0)) And Val(TxtDepoToBagSize) >= Val(MSHFlexGrid3.TextMatrix(u, 1))) Then
          If LCase(CmdDepoAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid3.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If

CheckDepoDuplicateBags = Retval

End Function


Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.FixedCols = 0
MSHFlexGrid2.Cols = 5

MSHFlexGrid2.TextMatrix(0, 0) = "Bag Size From"
MSHFlexGrid2.TextMatrix(0, 1) = "Bag Size To"
MSHFlexGrid2.TextMatrix(0, 2) = "Rate Per Bag"
MSHFlexGrid2.TextMatrix(0, 3) = "Discount Rate"
MSHFlexGrid2.TextMatrix(0, 4) = "SRCDID"

MSHFlexGrid2.ColWidth(0) = 1400
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 1800
MSHFlexGrid2.ColWidth(3) = 1800
MSHFlexGrid2.ColWidth(4) = 0

End Sub
Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.FixedRows = 1
MSHFlexGrid3.FixedCols = 0
MSHFlexGrid3.Cols = 5

MSHFlexGrid3.TextMatrix(0, 0) = "Bag Size From"
MSHFlexGrid3.TextMatrix(0, 1) = "Bag Size To"
MSHFlexGrid3.TextMatrix(0, 2) = "Rate Per Bag"
MSHFlexGrid3.TextMatrix(0, 3) = "Discount Rate"
MSHFlexGrid3.TextMatrix(0, 4) = "SRCDID"

MSHFlexGrid3.ColWidth(0) = 1400
MSHFlexGrid3.ColWidth(1) = 1400
MSHFlexGrid3.ColWidth(2) = 1800
MSHFlexGrid3.ColWidth(3) = 1800
MSHFlexGrid3.ColWidth(4) = 0

End Sub

Private Sub Grid_Head3()
MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 2
MSHFlexGrid4.Cols = 11


MSHFlexGrid4.TextMatrix(0, 0) = "SRC ID"
MSHFlexGrid4.TextMatrix(0, 1) = "SRC Date"
MSHFlexGrid4.TextMatrix(0, 2) = "Location "
MSHFlexGrid4.TextMatrix(0, 3) = "Exchange Type"
MSHFlexGrid4.TextMatrix(0, 4) = "Commodity Group"
MSHFlexGrid4.TextMatrix(0, 5) = "Billing Cycle"
MSHFlexGrid4.TextMatrix(0, 6) = "Billing Unit"
MSHFlexGrid4.TextMatrix(0, 7) = "Rate on Units"
MSHFlexGrid4.TextMatrix(0, 8) = "Deposit Billing Cycle"
MSHFlexGrid4.TextMatrix(0, 9) = "Status"
MSHFlexGrid4.TextMatrix(0, 10) = "G_UID"

MSHFlexGrid4.ColWidth(0) = 500
MSHFlexGrid4.ColWidth(1) = 0
MSHFlexGrid4.ColWidth(2) = 1000
MSHFlexGrid4.ColWidth(3) = 1000
MSHFlexGrid4.ColWidth(4) = 800
MSHFlexGrid4.ColWidth(5) = 1000
MSHFlexGrid4.ColWidth(6) = 1000
MSHFlexGrid4.ColWidth(7) = 900
MSHFlexGrid4.ColWidth(8) = 1000
MSHFlexGrid4.ColWidth(9) = 600
MSHFlexGrid4.ColWidth(10) = 0

MSHFlexGrid4.RowHeight(0) = 600


End Sub

'Private Sub GetSPRCGridData()
'
'Call Grid_Head3
'If CmbLocationID.Text = "" Then Exit Sub
'If CmbExchangeTypeID.Text = "" Then Exit Sub
'If CmbCommodityGroupID.Text = "" Then Exit Sub
'
'
'
'tmp = "Select MSRC.SRCID,MSRC.SRCDate,ML.LocationName,ME.ExchangeType,MCG.CommodityGroup,"
'tmp = tmp & " MB.BillingCycle ,MSRC.Flag, MSRC.RateonQty,MB1.BillingCycle ,MSRC.DefaultRate,MSRC.G_UID"
'tmp = tmp & " From Mst_StorageRateCard MSRC"
'tmp = tmp & " Inner Join Mst_Location ML On MSRC.LocationID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_ExchangeType ME ON MSRC.ExchangeTypeID = ME.ExchangeTypeID"
'tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MSRC.CommodityGroupID = MCG.CommodityGroupID"
'tmp = tmp & " Left Join Mst_BillingCycle MB On MSRC.BillingCycleID = MB.BillingCycleID "
'tmp = tmp & " Left Join Mst_BillingCycle MB1 On MSRC.DepoBillingCycleID = MB1.BillingCycleID "
'
'tmp = tmp & "Where ML.LocationName = '" & CmbLocationID.Text & "' And ME.ExchangeType = '" & CmbExchangeTypeID.Text & "' And MCG.CommodityGroup = '" & CmbCommodityGroupID.Text & "'"
'
'
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'   MSHFlexGrid4.Rows = TmpRs.RecordCount + 2
'   For I = 1 To TmpRs.RecordCount
'       For c = 0 To TmpRs.Fields.Count - 1
'           MSHFlexGrid4.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
'       Next
'       MSHFlexGrid4.TextMatrix(I, 1) = Format(MSHFlexGrid4.TextMatrix(I, 1), "dd-mmm-yyyy")
'       If LCase(MSHFlexGrid4.TextMatrix(I, 6)) = "b" Then
'          MSHFlexGrid4.TextMatrix(I, 6) = "Bags"
'       ElseIf LCase(MSHFlexGrid4.TextMatrix(I, 6)) = "q" Then
'          MSHFlexGrid4.TextMatrix(I, 6) = "Quantity"
'       ElseIf LCase(MSHFlexGrid4.TextMatrix(I, 6)) = "s" Then
'          MSHFlexGrid4.TextMatrix(I, 6) = "Sq ft"
'       Else
'          MSHFlexGrid4.TextMatrix(I, 6) = "" '"Sq ft"
'       End If
'       If MSHFlexGrid4.TextMatrix(I, 9) = 0 Then
'          MSHFlexGrid4.TextMatrix(I, 9) = "No"
'       Else
'          MSHFlexGrid4.TextMatrix(I, 9) = "Yes"
'          If TxtSRCID.Text <> MSHFlexGrid4.TextMatrix(I, 0) Then
'             ChkDefault.Value = 0
'             ChkDefault.Enabled = False
'          End If
'       End If
'       TmpRs.MoveNext
'   Next
'End If
'
'
'End Sub

Private Sub CmdClear_Click()

Call Grid_Head
ChkUpdated.Value = 0
TxtUpdatedFrom.Value = "01/01/2008"
TxtUpdatedTo.Value = Date
CmbSCommodityID.Text = ""
CmbSLocationID.Text = ""
LblTotalRecords.Caption = ""
CmbSBillingCycle.Text = ""
CmbSBillingUnit.Text = ""
CmbSearchBaseRate.Text = ""
TxtSearchClientName = ""
CmbSStatus.Text = ""

End Sub

Private Sub CmdClientID_Click()
FrmMst_Client.Show
End Sub

Private Sub CmdDeleteList_Click()
I = MSHFlexGrid2.RowSel

If TxtFromBagSize = "" Then Exit Sub


If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   End If
End If
Call ClearDetail
End Sub

Private Sub CmdDepoBillingCycle_Click()
FrmMst_BillingCycle.Show
End Sub

Private Sub CmdDepoDeleteList_Click()
I = MSHFlexGrid3.RowSel

If TxtDepoFromBagSize = "" Then Exit Sub

If I > 0 Then
   If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
      For RO = I To MSHFlexGrid3.Rows - 2
          For C = 0 To MSHFlexGrid3.Cols - 1
              MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
   End If
End If
Call ClearDepoDetail
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_SpecialRateCard.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_SpecialRateCard.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
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
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim wc As String
Call Grid_Head
wc = ""


If ChkUpdated.Value = 1 Then
   wc = " Where MSPRC.UpdatedDate between '" & Format(TxtUpdatedFrom.Value, Gen_DatFormat) & "' and '" & Format(TxtUpdatedTo.Value + 1, Gen_DatFormat) & "'   "
Else
   wc = " Where (MSPRC.UpdatedDate between '" & Format(TxtUpdatedFrom.Value, Gen_DatFormat) & "' and '" & Format(TxtUpdatedTo.Value + 1, Gen_DatFormat) & "' Or (MSPRC.UpdatedDate is null))  "
End If

If CmbSLocationID.Text <> "" Then
   wc = GenWc(wc, CmbSLocationID.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, tmp, "", "S", "I")
   End If
End If
If CmbSCommodityID.Text <> "" Then
   wc = GenWc(wc, CmbSCommodityID.Text, "CommodityGroup", "S", "N")
End If
If CmbSBillingCycle.Text <> "" Then
   wc = GenWc(wc, CmbSBillingCycle.Text, "BillingCycle", "S", "N")
End If

If CmbSBillingUnit.Text <> "" Then
   wc = GenWc(wc, Left(CmbSBillingUnit.Text, 1), "MSPRC.BillingUnit", "S", "N")
End If

If CmbSearchBaseRate <> "" Then
   wc = GenWc(wc, mSearchBase, "MSPRC.BaseRate", "S", "N")
End If

If TxtSearchClientName <> "" Then
   wc = GenWc(wc, TxtSearchClientName, "MClient.ClientName", "S", "L")
End If

If CmbSStatus.Text <> "" Then
   wc = GenWc(wc, mSStatus, "MSPRC.Flag", "S", "N")
End If

If Trim(CmbSConClient.Text) <> "" Then
   If LCase(CmbSConClient.Text) = "yes" Then
      wc = GenWc(wc, 1, "MSPRC.ConvertedClient", "S", "N")
   ElseIf LCase(CmbSConClient.Text) = "no" Then
      wc = GenWc(wc, 0, "MSPRC.ConvertedClient", "S", "N")
   End If
End If

tmp = "Select MSPRC.SPRCID,MSPRC.SPRCDate,ML.LocationName,MCG.CommodityGroup,MClient.ClientID,MClient.ClientName,"
tmp = tmp & " MB.BillingCycle ,MSPRC.BillingUnit, MSPRC.RateonQty,MSPRC.BaseRate,MSPRC.AbsoluteDiscount,MSPRC.Amount,MSPRC.G_UID,MSPRC.Flag,MSPRC.Remarks,MSPRC.UpdatedBy,MSPRC.UpdatedDate ,MSPRC.ConvertedClient"
tmp = tmp & " From Mst_SpecialRateCard MSPRC"
tmp = tmp & " Left Join Mst_Location ML On MSPRC.LocationID = ML.LocationID"
tmp = tmp & " Left Join Mst_CommodityGroup MCG On MSPRC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_Client MClient On MSPRC.ClientIDRow = MClient.ClientIDRow"
tmp = tmp & " Left Join Mst_BillingCycle MB On MSPRC.BillingCycleID = MB.BillingCycleID "

tmp1 = "Order By MSPRC.SPRCID"

tmp = tmp & wc & tmp1

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-mmm-yyyy")
       If LCase(MSHFlexGrid1.TextMatrix(I, 7)) = "b" Then
          MSHFlexGrid1.TextMatrix(I, 7) = "Bags"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 7)) = "q" Then
          MSHFlexGrid1.TextMatrix(I, 7) = "Quantity"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 7)) = "s" Then
          MSHFlexGrid1.TextMatrix(I, 7) = "Sq Ft"
       Else
          MSHFlexGrid1.TextMatrix(I, 7) = ""
       End If
       If LCase(MSHFlexGrid1.TextMatrix(I, 9)) = "n" Then
          MSHFlexGrid1.TextMatrix(I, 9) = "No"
       Else
          MSHFlexGrid1.TextMatrix(I, 9) = "Yes"
       End If
       If LCase(MSHFlexGrid1.TextMatrix(I, 13)) = "a" Then
          MSHFlexGrid1.TextMatrix(I, 13) = "Active"
       Else
          MSHFlexGrid1.TextMatrix(I, 13) = "De-Active"
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 16) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 16) = Format(MSHFlexGrid1.TextMatrix(I, 16), "dd-mmm-yyyy hh:mm:ss")
       End If
       If MSHFlexGrid1.TextMatrix(I, 17) = "1" Then
          MSHFlexGrid1.TextMatrix(I, 17) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 17) = "No"
       End If
       TmpRs.MoveNext
       
   Next
Else
   MsgBox "No Record Found"
End If
LblTotalRecords.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub




Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_SpecialRateCard.MoveFirst
RsMst_SpecialRateCard.Find "SPRCID = " & MSHFlexGrid1.TextMatrix(I, 0) & ""
If Not RsMst_SpecialRateCard.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBaseRate.Text) <> "no" Then Exit Sub
Dim r As Variant
r = MSHFlexGrid2.RowSel
If r < 1 Or r = MSHFlexGrid2.Rows - 1 Then
   Call ClearDetail
   Exit Sub
End If
If MSHFlexGrid2.TextMatrix(r, 0) = "" Then
   Exit Sub
End If

TxtFromBagSize = MSHFlexGrid2.TextMatrix(r, 0)
TxtToBagSize = MSHFlexGrid2.TextMatrix(r, 1)
TxtRatePerBag = MSHFlexGrid2.TextMatrix(r, 2)
CmdAdd2List.Caption = "Update List"
TxtFromBagSize.SetFocus
End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBaseRate.Text) <> "no" Then Exit Sub
Dim r As Variant
r = MSHFlexGrid3.RowSel
If r < 1 Or r = MSHFlexGrid3.Rows - 1 Then
   Call ClearDetail
   Exit Sub
End If
If MSHFlexGrid3.TextMatrix(r, 0) = "" Then
   Exit Sub
End If

TxtDepoFromBagSize = MSHFlexGrid3.TextMatrix(r, 0)
TxtDepoToBagSize = MSHFlexGrid3.TextMatrix(r, 1)
TxtDepoRatePerBag = MSHFlexGrid3.TextMatrix(r, 2)
CmdDepoAdd2List.Caption = "Update List"
TxtDepoFromBagSize.SetFocus
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_SpecialRateCard.MoveNext
If RsMst_SpecialRateCard.EOF Then
   RsMst_SpecialRateCard.MoveLast
   LF_Flag = "Y"
End If
Call Indata
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
RsMst_SpecialRateCard.MovePrevious
If RsMst_SpecialRateCard.BOF Then
   RsMst_SpecialRateCard.MoveFirst
   LF_Flag = "Y"
End If
Call Indata
If LF_Flag = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_SpecialRateCard.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_SpecialRateCard.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True

CmbLocationID.Text = TxtLocationID.Text
CmbCommodityGroupID.Text = TxtCommodityGroupID.Text

CmbBillingCycle.Text = TxtBillingCycle.Text
CmbBillingUnit.Text = TxtBillingUnit.Text

CmbBaseRate.Visible = True
TxtBaseRate.Visible = False
CmbBaseRate.Text = TxtBaseRate

CmdDepoBillingCycle.Visible = True
CmbDepoBillingCycle.Visible = True
TxtDepoBillingCycle.Visible = False
CmbDepoBillingCycle.Text = TxtDepoBillingCycle

CmbDiscount.Visible = True
TxtDiscount.Visible = False
CmbDiscount.Text = TxtDiscount

CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = TxtStatus

TxtRemarks.Locked = False

If LCase(CmbBaseRate.Text) = "yes" Then
   CmbDiscount.Enabled = True
Else
   CmbDiscount.Enabled = False
End If
If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBaseRate.Text) <> "no" Then
   Frame2.Enabled = False
   Frame3.Enabled = False
Else
   Frame2.Enabled = True
   Frame3.Enabled = True
End If
Call GButtonLock(Me, False)

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   Call DeleteChildTableDetail
   Call DeleteChildTableDepoDetail
   RsMst_SpecialRateCard.Delete
   RsMst_SpecialRateCard.Update
   RsMst_SpecialRateCard.Requery
   If RsMst_SpecialRateCard.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_SpecialRateCard.MoveNext
   If RsMst_SpecialRateCard.EOF() Then
      RsMst_SpecialRateCard.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_SpecialRateCard.CancelUpdate
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If Not RsMst_SpecialRateCard.EOF Then
'      RsMst_SpecialRateCard.MoveFirst
'      RsMst_SpecialRateCard.Find "SPRCID = " & TxtSPRCID & ""
      Call Indata
   End If
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 300
Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1600

''===========comment on 7 july 2012 for not clearing search grid=============
'Call CmdClear_Click
''===========comment on 7 july 2012 for not clearing search grid=============


End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Call Grid_Head1
Call Grid_Head2
mExchangeTypeID = GetExchangeID("Non Ncdex")
Frame1.Enabled = True

TxtSPRCDate.Enabled = True

TxtLocationID.Visible = False
TxtCommodityGroupID.Visible = False
TxtClientID.Visible = False
Check1.Enabled = True
TxtBillingCycle.Visible = False
TxtBillingUnit.Visible = False

CmbLocationID.Visible = True
CmbCommodityGroupID.Visible = True
CmbBillingCycle.Visible = True
CmbBillingUnit.Visible = True
CmbClientID.Visible = True

CmdLocationID.Visible = True
CmdCommodityGroupID.Visible = True

CmdBillingCycle.Visible = True
CmdClientID.Visible = True

Check1.Value = vbUnchecked
CmbLocationID.Text = ""
CmbCommodityGroupID.Text = ""
CmbBillingCycle.Text = ""
CmbBillingUnit.Text = ""
CmbClientID.Text = ""

CmbDiscount.Visible = True
TxtDiscount.Visible = False
CmbDiscount.Text = ""

CmbBaseRate.Visible = True
TxtBaseRate.Visible = False
CmbBaseRate.Text = ""

CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = ""

CmdDepoBillingCycle.Visible = True
CmbDepoBillingCycle.Visible = True
TxtDepoBillingCycle.Visible = False
CmbDepoBillingCycle.Text = ""

TxtSPRCID.Text = ""

TxtSPRCDate.Value = Date

TxtPreFix = ""
TxtRateOnQty = ""
TxtAmount = ""
TxtTotDays = ""
TxtCreated = ""
TxtUpdated = ""
TxtRemarks.Locked = False
TxtRemarks = ""
'Commented by Amit 15-Mar-2011. Its already called on top of AddCmd_Click
'Call Grid_Head1

If RsMst_SpecialRateCard.RecordCount > 0 Then
   TxtSPRCDate.SetFocus
End If

End Sub
Private Sub CmbSCommodityID_GotFocus()

tmp = CmbSCommodityID.Text
Call TableMst_CommodityGroup
CmbSCommodityID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_CommodityGroup.RecordCount
   CmbSCommodityID.AddItem RsMst_CommodityGroup!COmmodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbSCommodityID.Text = tmp
    
End Sub

'----
Private Sub CmbSCommodityID_LostFocus()
If CmbSCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbSCommodityID.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodityID.SetFocus
   Exit Sub
End If
mSCommodityID = RsMst_CommodityGroup!CommodityGroupID
End Sub
Private Sub CmbCommodityGroupID_GotFocus()

tmp = CmbCommodityGroupID.Text
Call TableMst_CommodityGroup
CmbCommodityGroupID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_CommodityGroup.RecordCount
   CmbCommodityGroupID.AddItem RsMst_CommodityGroup!COmmodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbCommodityGroupID.Text = tmp
    
End Sub

Private Sub CmbCommodityGroupID_LostFocus()
If CmbCommodityGroupID.Text = "" Then
   mCommodityGroupId = Null
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbCommodityGroupID.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If
mCommodityGroupId = RsMst_CommodityGroup!CommodityGroupID
Call GetStorageRateFromMaster
Call GetDepoStorageRateFromMaster
End Sub
Private Sub CmbsExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbSExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbsExchangeTypeID_LostFocus()
If CmbSExchangeTypeID.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSExchangeTypeID.SetFocus
   Exit Sub
End If

mSExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
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
   TxtPreFix = ""
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
If CmbLocationID.Text <> RsMst_Location!LocationName Then
   Call Grid_Head1
   Call Grid_Head2
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = RsMst_Location!SM_Prefix
Call GetStorageRateFromMaster
Call GetDepoStorageRateFromMaster
End Sub
Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If

CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbSLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
    
End Sub

'----
Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmdCommodityGroupID_Click()
FrmMst_CommodityGroup.Show
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call CmdClear_Click

Frame1.Enabled = True
If Gen_WcLocation = "" Then
   Call TableMst_SpecialRateCard
Else
   Call TableMst_SpecialRateCard(" Where " & Gen_WcLocation)
End If
If RsMst_SpecialRateCard.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call GButtonLock(Me, True)

Call Indata

End Sub
Private Sub Indata()

Call Grid_Head1
Call Grid_Head2
Frame1.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False

Check1.Enabled = False

CmbLocationID.Visible = False
CmbCommodityGroupID.Visible = False
CmbClientID.Visible = False

TxtLocationID.Visible = True
TxtClientID.Visible = True
TxtCommodityGroupID.Visible = True
TxtBillingCycle.Visible = True
TxtBillingUnit.Visible = True
CmdLocationID.Visible = False
CmdClientID.Visible = False
CmdCommodityGroupID.Visible = False
CmbBillingCycle.Visible = False
CmbBillingUnit.Visible = False
CmdBillingCycle.Visible = False
mExchangeTypeID = GetExchangeID("Non NCDEX")  'RsMst_SpecialRateCard!ExchangeTypeId

TxtSPRCID.Text = RsMst_SpecialRateCard!SPRCID
TxtSPRCDate.Value = RsMst_SpecialRateCard!SPRCDate

mClientIDRow = RsMst_SpecialRateCard!ClientIDRow
TxtClientID = GetClientName(RsMst_SpecialRateCard!ClientIDRow)

CmbClientID.Text = TxtClientID
Check1.Value = IIf(IsNull(RsMst_SpecialRateCard!ConvertedClient), vbUnchecked, RsMst_SpecialRateCard!ConvertedClient)
CmbDiscount.Visible = False
TxtDiscount.Visible = True

CmdDepoBillingCycle.Visible = False
CmbDepoBillingCycle.Visible = False
TxtDepoBillingCycle.Visible = True

mFlag = RsMst_SpecialRateCard!BillingUnit


If LCase(mFlag) = "b" Then
   TxtBillingUnit = "Bags"
ElseIf LCase(mFlag) = "q" Then
   TxtBillingUnit = "Quantity"
ElseIf LCase(mFlag) = "s" Then
   TxtBillingUnit = "Sq Ft"
Else
   TxtBillingUnit = ""
End If

CmbBaseRate.Visible = False
TxtBaseRate.Visible = True

CmbStatus.Visible = False
TxtStatus.Visible = True


mBaseRate = RsMst_SpecialRateCard!BaseRate

If LCase(mBaseRate) = "y" Then
   TxtBaseRate = "Yes"
ElseIf LCase(mBaseRate) = "n" Then
   TxtBaseRate = "No"
End If
mStatus = IIf(IsNull(RsMst_SpecialRateCard!Flag), "", RsMst_SpecialRateCard!Flag)
TxtDiscount = IIf(IsNull(RsMst_SpecialRateCard!AbsoluteDiscount), "", RsMst_SpecialRateCard!AbsoluteDiscount)
TxtAmount = IIf(IsNull(RsMst_SpecialRateCard!Amount), 0, RsMst_SpecialRateCard!Amount)

CmbBaseRate.Text = TxtBaseRate
CmbDiscount.Text = TxtDiscount

If mStatus = "A" Then
   TxtStatus = "Active"
ElseIf mStatus = "D" Then
   TxtStatus = "De-Active"
Else
   TxtStatus = ""
End If


'If LCase(CmbBaseRate.Text) <> "no" Then
mSRCID = RsMst_SpecialRateCard!SRCID
'End If

mLocationID = RsMst_SpecialRateCard!LocationID
mCommodityGroupId = RsMst_SpecialRateCard!CommodityGroupID
mBillingCycleID = RsMst_SpecialRateCard!BillingCycleID
mDepoBillingCycleID = RsMst_SpecialRateCard!DepoBillingCycleID
  
TxtLocationID.Text = GetLocationName(RsMst_SpecialRateCard!LocationID)
TxtCommodityGroupID.Text = GetCommodityGroupName(RsMst_SpecialRateCard!CommodityGroupID)
If IsNull(RsMst_SpecialRateCard!BillingCycleID) = False Then
   TxtBillingCycle.Text = GetBillingCycle(RsMst_SpecialRateCard!BillingCycleID, "")
   TxtTotDays = GetBillingCycle(RsMst_SpecialRateCard!BillingCycleID, "D")
Else
   TxtBillingCycle.Text = "" 'GetBillingCycle(RsMst_SpecialRateCard!BillingCycleID, "")
   TxtTotDays = "" 'GetBillingCycle(RsMst_SpecialRateCard!BillingCycleID, "D")
End If

TxtRateOnQty.Text = IIf(IsNull(RsMst_SpecialRateCard!RateOnQty), 0, RsMst_SpecialRateCard!RateOnQty) 'IIf(IsNull(RsMst_StorageRateCard!RateOnQty), 0, RsMst_StorageRateCard!RateOnQty)
CmbLocationID.Text = TxtLocationID.Text
CmbCommodityGroupID.Text = TxtCommodityGroupID.Text
CmbBillingCycle.Text = TxtBillingCycle

If IsNull(RsMst_SpecialRateCard!DepoBillingCycleID) = False Then
   TxtDepoBillingCycle.Text = GetBillingCycle(RsMst_SpecialRateCard!DepoBillingCycleID, "")
Else
   TxtDepoBillingCycle.Text = ""
End If

CmbDepoBillingCycle.Text = TxtDepoBillingCycle

TxtRemarks.Locked = True
TxtRemarks = IIf(IsNull(RsMst_SpecialRateCard!Remarks), "", RsMst_SpecialRateCard!Remarks)

TxtCreated = IIf(IsNull(RsMst_SpecialRateCard!CreatedBy), "", RsMst_SpecialRateCard!CreatedBy) & " :- " & IIf(IsNull(RsMst_SpecialRateCard!CreatedDate), "", RsMst_SpecialRateCard!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_SpecialRateCard!UpdatedBy), "", RsMst_SpecialRateCard!UpdatedBy) & " :- " & IIf(IsNull(RsMst_SpecialRateCard!UpdatedDate), "", RsMst_SpecialRateCard!UpdatedDate)

If LCase(TxtBillingUnit) = "bags" Then
   Call InDataBagSize
   Call IndataDepoBagSize
End If
Call GButtonLock(Me, True)

End Sub
Private Sub InDataBagSize()
Call TableMst_SpecialRateCardDetail(" where SPRCID=" & Val(TxtSPRCID.Text) & "")
Call TableMst_StorageCardRateDetail
MSHFlexGrid2.Rows = RsMst_SpecialRateCardDetail.RecordCount + 2
For I = 1 To MSHFlexGrid2.Rows - 2
    If LCase(CmbBaseRate.Text) = "yes" Then
       RsMst_StorageCardRateDetail.Find "SRCDID = " & RsMst_SpecialRateCardDetail!SRCDID & ""
       If Not RsMst_StorageCardRateDetail.EOF Then
          MSHFlexGrid2.TextMatrix(I, 0) = RsMst_StorageCardRateDetail!FromBagSize
          MSHFlexGrid2.TextMatrix(I, 1) = RsMst_StorageCardRateDetail!ToBagSize
          MSHFlexGrid2.TextMatrix(I, 2) = Format(RsMst_StorageCardRateDetail!RatePerBag, "#####0.0000")
          MSHFlexGrid2.TextMatrix(I, 4) = RsMst_SpecialRateCardDetail!SRCDID
       End If
    Else
       MSHFlexGrid2.TextMatrix(I, 0) = RsMst_SpecialRateCardDetail!FromBagSize
       MSHFlexGrid2.TextMatrix(I, 1) = RsMst_SpecialRateCardDetail!ToBagSize
       MSHFlexGrid2.TextMatrix(I, 2) = Format(RsMst_SpecialRateCardDetail!RatePerBag, "#####0.0000")
    End If
    MSHFlexGrid2.TextMatrix(I, 3) = "" 'Format(RsMst_SpecialRateCardDetail!DiscountRate, "#####0.0000")
    RsMst_SpecialRateCardDetail.MoveNext
Next
Call IndataDepoBagSize
Call TxtAmount_LostFocus

End Sub
Private Sub IndataDepoBagSize()
Call TableMst_SpecialCardDepoRateDetail(" Where SPRCID=" & Val(TxtSPRCID.Text) & "")
Call TableMst_StorageCardRateDetail
MSHFlexGrid3.Rows = RsMst_SpecialCardDepoRateDetail.RecordCount + 2
For I = 1 To MSHFlexGrid3.Rows - 2
    If LCase(CmbBaseRate.Text) = "yes" Then
       RsMst_StorageCardRateDetail.Find "SRCDID = " & RsMst_SpecialCardDepoRateDetail!SRCDID & ""
       If Not RsMst_StorageCardRateDetail.EOF Then
          MSHFlexGrid3.TextMatrix(I, 0) = RsMst_StorageCardRateDetail!FromBagSize
          MSHFlexGrid3.TextMatrix(I, 1) = RsMst_StorageCardRateDetail!ToBagSize
          MSHFlexGrid3.TextMatrix(I, 2) = Format(RsMst_StorageCardRateDetail!RatePerBag, "#####0.0000")
          MSHFlexGrid3.TextMatrix(I, 4) = RsMst_StorageCardDepoRateDetail!SRCDID
       End If
    Else
       MSHFlexGrid3.TextMatrix(I, 0) = RsMst_SpecialCardDepoRateDetail!FromBagSize
       MSHFlexGrid3.TextMatrix(I, 1) = RsMst_SpecialCardDepoRateDetail!ToBagSize
       MSHFlexGrid3.TextMatrix(I, 2) = Format(RsMst_SpecialCardDepoRateDetail!RatePerBag, "#####0.0000")
    End If
    MSHFlexGrid3.TextMatrix(I, 3) = "" 'Format(RsMst_StorageCardDepoRateDetail!DiscountRate, "#####0.0000")
    RsMst_SpecialCardDepoRateDetail.MoveNext
Next
End Sub
Private Sub SaveCmd_Click()

'If TxtSPRCDate.Value < Date Then
'   MsgBox "SRC Date Should Not be Less than Current Date !!!"
'   TxtSPRCDate.SetFocus
'   Exit Sub
'End If

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Name not allowed !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCommodityGroupID.Text = "" Then
   MsgBox "Blank Commodity Name not allowed !!!"
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If

If CmbClientID = "" Then
   MsgBox "Blank Client ID not allowed !!! "
   TxtClientID.SetFocus
   Exit Sub
End If

If CmbBillingCycle.Text = "" Then
   MsgBox "Blank Billing Cycle not allowed!!!"
   CmbBillingCycle.SetFocus
   Exit Sub
End If

If CmbBillingUnit.Text = "" Then
   MsgBox "Blank Billing Unit not allowed!!!"
   CmbBillingUnit.SetFocus
   Exit Sub
End If

If LCase(CmbBillingUnit.Text) = "bags" Then
   If MSHFlexGrid2.Rows < 3 Then
      MsgBox "Bag Size details missing."
      TxtFromBagSize.SetFocus
      Exit Sub
   End If
   If MSHFlexGrid3.Rows < 3 Then
      MsgBox "Deposit Bag Size details missing."
      TxtDepoFromBagSize.SetFocus
      Exit Sub
   End If
   mFlag = "B"
ElseIf LCase(CmbBillingUnit.Text) = "quantity" Then
   mFlag = "Q"
ElseIf LCase(CmbBillingUnit.Text) = "sq ft" Then
   mFlag = "S"
Else
   mFlag = Null
End If

If CmbBaseRate.Text = "" Then
   MsgBox "Blank Base Rate Not Allowed !!!"
   CmbBaseRate.SetFocus
   Exit Sub
End If

If LCase(CmbBaseRate.Text) = "yes" Then
   If CmbDiscount.Text = "" Then
      MsgBox "Blank Discount in not allowed !!!"
      CmbDiscount.SetFocus
      Exit Sub
   End If
End If

'If LCase(CmbBillingUnit.Text) <> "bags" Then
'   If TxtRateOnQty = "" Then
'      MsgBox "Blank Rate on Units not allowed !!!"
'      TxtRateOnQty.SetFocus
'      Exit Sub
'   End If
'End If

If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBaseRate.Text) <> "no" Then
   If TxtAmount.Text = "" Then
      MsgBox "Blank Amount Not Allowed !!!"
      TxtAmount.SetFocus
      Exit Sub
   End If
End If


'If TxtFromDate.Value > TxtToDate.Value Then
'   MsgBox "'From Date' must be less than 'To Date' always!!!"
'   TxtToDate.SetFocus
'   Exit Sub
'End If

If ChkDateRange = False Then
   MsgBox "Rates already exist for selected date range!!!"
   TxtToDate.SetFocus
   Exit Sub
End If


'RsMst_SpecialRateCard
If Edit_Flg = "A" Then
   If RsMst_SpecialRateCard.RecordCount > 0 Then
      RsMst_SpecialRateCard.MoveFirst
   End If
   'RsMst_SpecialRateCard.Filter = "LocationID = " & mLocationID & " And ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & " And FromDate = '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And ToDate = '" & Format(TxtToDate.Value, Gen_DatFormat) & "' And Flag = '" & mFlag & "' And BillingCycleID = " & mBillingCycleID & " And DefaultRate = " & ChkDefault.Value & " "
   'RsMst_SpecialRateCard.Filter = "LocationID = " & mLocationID & " And ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & "  And BillingCycleID = " & mBillingCycleID & " And Flag = '" & mFlag & "' "
   If mStatus <> "D" Then
      RsMst_SpecialRateCard.Filter = "ClientIDRow = " & mClientIDRow & "  And LocationID = " & mLocationID & " And ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & " And Flag = '" & mStatus & "'" '"And BillingCycleID = " & mBillingCycleID & " And BillingUnit = '" & mFlag & "'"
      If RsMst_SpecialRateCard.RecordCount > 0 Then
         MsgBox "Dulicate Special Rate Card Not allowed !!!"
         RsMst_SpecialRateCard.Filter = ""
         Exit Sub
      End If
   End If
   RsMst_SpecialRateCard.Filter = ""
   RsMst_SpecialRateCard.AddNew
   RsMst_SpecialRateCard!SPRCID = GetSPRCID
   RsMst_SpecialRateCard!G_UID = GetGUID
   TxtSPRCID.Text = RsMst_SpecialRateCard!SPRCID
Else
   RsMst_SpecialRateCard.MoveFirst
   'RsMst_SpecialRateCard.Filter = "LocationID = " & mLocationID & " And  ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & " And FromDate = '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And ToDate = '" & Format(TxtToDate.Value, Gen_DatFormat) & "'  And Flag = '" & mFlag & "' And BillingCycleID = " & mBillingCycleID & " And DefaultRate = " & ChkDefault.Value & " "
   If mStatus <> "D" Then
      RsMst_SpecialRateCard.Filter = "ClientIDRow = " & mClientIDRow & "  And LocationID = " & mLocationID & " And ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & " And Flag = '" & mStatus & "'" '" And BillingCycleID = " & mBillingCycleID & " And BillingUnit = '" & mFlag & "' "
      If RsMst_SpecialRateCard.RecordCount > 0 Then
         If RsMst_SpecialRateCard!SPRCID <> TxtSPRCID Then
            MsgBox "Dulicate Special Rate Card Not allowed !!!"
            RsMst_SpecialRateCard.Filter = ""
            Exit Sub
         End If
      End If
   End If
   RsMst_SpecialRateCard.Filter = ""
   RsMst_SpecialRateCard.MoveFirst
   RsMst_SpecialRateCard.Find " SPRCID= " & TxtSPRCID.Text
   Call DeleteChildTableDetail
   Call DeleteChildTableDepoDetail
End If
RsMst_SpecialRateCard!SPRCID = TxtSPRCID.Text
RsMst_SpecialRateCard!SPRCDate = TxtSPRCDate.Value
If LCase(CmbBaseRate.Text) <> "no" Then
   RsMst_SpecialRateCard!SRCID = mSRCID
End If
RsMst_SpecialRateCard!LocationID = mLocationID
RsMst_SpecialRateCard!CommodityGroupID = mCommodityGroupId
RsMst_SpecialRateCard!ExchangeTypeID = mExchangeTypeID
RsMst_SpecialRateCard!BillingCycleID = mBillingCycleID
RsMst_SpecialRateCard!RateOnQty = Val(TxtRateOnQty.Text)

RsMst_SpecialRateCard!BillingUnit = mFlag
RsMst_SpecialRateCard!ClientIDRow = mClientIDRow
RsMst_SpecialRateCard!BaseRate = mBaseRate
RsMst_SpecialRateCard!AbsoluteDiscount = CmbDiscount.Text
RsMst_SpecialRateCard!Amount = Val(TxtAmount.Text)
RsMst_SpecialRateCard!DepoBillingCycleID = mDepoBillingCycleID
RsMst_SpecialRateCard!Remarks = Trim(TxtRemarks)
RsMst_SpecialRateCard!Flag = mStatus
RsMst_SpecialRateCard!ConvertedClient = Check1.Value

If IsNull(RsMst_SpecialRateCard!CreatedBy) = True Or RsMst_SpecialRateCard!CreatedBy = "" Then
   RsMst_SpecialRateCard!CreatedBy = Gen_UserLoginID
   RsMst_SpecialRateCard!CreatedDate = Now
Else
   RsMst_SpecialRateCard!UpdatedBy = Gen_UserLoginID
   RsMst_SpecialRateCard!UpdatedDate = Now
End If

RsMst_SpecialRateCard.Update

If CmbBillingUnit.Text = "Bags" Then
   Call SaveSpecialRateCardDetails
   Call SaveSpecialRateCardDepoDetails
End If

Call GButtonLock(Me, True)

RsMst_SpecialRateCard.MoveFirst
RsMst_SpecialRateCard.Find "SPRCID = " & TxtSPRCID & ""

Call Indata

End Sub
Private Sub SaveSpecialRateCardDetails()

Call TableMst_SpecialRateCardDetail(" Where SPRCID=" & Val(TxtSPRCID.Text) & "")

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsMst_SpecialRateCardDetail.AddNew
       RsMst_SpecialRateCardDetail!SPRCDID = GetMaxSPRCDID
       RsMst_SpecialRateCardDetail!SPRCID = Val(TxtSPRCID.Text)
       If LCase(CmbBaseRate.Text) = "no" Then
          RsMst_SpecialRateCardDetail!FromBagSize = Val(MSHFlexGrid2.TextMatrix(I, 0))
          RsMst_SpecialRateCardDetail!ToBagSize = Val(MSHFlexGrid2.TextMatrix(I, 1))
          RsMst_SpecialRateCardDetail!RatePerBag = Val(MSHFlexGrid2.TextMatrix(I, 2))
       Else
          RsMst_SpecialRateCardDetail!SRCDID = Val(MSHFlexGrid2.TextMatrix(I, 4))
       End If
       RsMst_SpecialRateCardDetail.Update
       RsMst_SpecialRateCardDetail.Requery
    End If
Next

End Sub
Private Sub SaveSpecialRateCardDepoDetails()

Call TableMst_SpecialCardDepoRateDetail(" Where SPRCID=" & Val(TxtSPRCID.Text) & "")

For I = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
       RsMst_SpecialCardDepoRateDetail.AddNew
       RsMst_SpecialCardDepoRateDetail!SPRCDDID = GetMaxDepoSPRCDID
       RsMst_SpecialCardDepoRateDetail!SPRCID = Val(TxtSPRCID.Text)
       If LCase(CmbBaseRate.Text) = "no" Then
          RsMst_SpecialCardDepoRateDetail!FromBagSize = Val(MSHFlexGrid3.TextMatrix(I, 0))
          RsMst_SpecialCardDepoRateDetail!ToBagSize = Val(MSHFlexGrid3.TextMatrix(I, 1))
          RsMst_SpecialCardDepoRateDetail!RatePerBag = Val(MSHFlexGrid3.TextMatrix(I, 2))
       Else
          RsMst_SpecialCardDepoRateDetail!SRCDID = Val(MSHFlexGrid3.TextMatrix(I, 4))
       End If
       RsMst_SpecialCardDepoRateDetail.Update
       RsMst_SpecialCardDepoRateDetail.Requery
    End If
Next
End Sub
Private Function ChkDateRange() As Boolean
Dim Retval As Boolean
Retval = True


ChkDateRange = Retval
End Function

Sub DeleteChildTableDetail()
tmp = "Delete from Mst_SpecialRateCardDetail where SPRCID = " & Val(TxtSPRCID.Text) & ""
Call TmpTable
End Sub

Sub DeleteChildTableDepoDetail()
tmp = "Delete from Mst_SpecialRateCardDepoDetail where SPRCID = " & Val(TxtSPRCID.Text) & ""
Call TmpTable
End Sub

Private Function GetSPRCID()
Dim Retval As Double
tmp = "Select max(SPRCID) from Mst_SpecialRateCard "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetSPRCID = Retval
End Function
Private Function GetMaxSPRCDID()
Dim Retval As Double
tmp = "Select max(SPRCDID) from Mst_SpecialRateCardDetail "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxSPRCDID = Retval
End Function

Private Function GetMaxDepoSPRCDID()
Dim Retval As Double
tmp = "Select max(SPRCDDID) from Mst_SpecialRateCardDepoDetail "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxDepoSPRCDID = Retval
End Function

Private Sub TxtAmount_LostFocus()

If TxtAmount = "" Then Exit Sub
If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBaseRate.Text) <> "no" Then
   If IsNumeric(TxtAmount) = False Or Val(TxtAmount) <= 0 Then
      MsgBox "Invalid numeric format!!!"
      TxtAmount.SetFocus
      Exit Sub
   End If
End If

Call GetDiscountedRate
Call GetDiscountedDepoRate
End Sub
Private Sub GetDiscountedRate()
If CmbBaseRate.Text = "" Then Exit Sub

If LCase(CmbBaseRate.Text) = "yes" Then
   If CmbDiscount.Text = "" Then Exit Sub
End If

If LCase(CmbBaseRate.Text) = "yes" Then
   If LCase(CmbDiscount.Text) = "actual" Then
      For I = 1 To MSHFlexGrid2.Rows - 1
          If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit For
          MSHFlexGrid2.TextMatrix(I, 3) = Val(MSHFlexGrid2.TextMatrix(I, 2)) - Val(TxtAmount)
      Next
   ElseIf LCase(CmbDiscount.Text) = "percentage" Then
      For I = 1 To MSHFlexGrid2.Rows - 1
          If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit For
          MSHFlexGrid2.TextMatrix(I, 3) = Val(MSHFlexGrid2.TextMatrix(I, 2)) - Val(MSHFlexGrid2.TextMatrix(I, 2)) * (Val(TxtAmount) / 100)
      Next
   End If
End If

End Sub

Private Sub GetDiscountedDepoRate()
If CmbBaseRate.Text = "" Then Exit Sub

If LCase(CmbBaseRate.Text) = "yes" Then
   If CmbDiscount.Text = "" Then Exit Sub
End If

If LCase(CmbBaseRate.Text) = "yes" Then
   If LCase(CmbDiscount.Text) = "actual" Then
      For I = 1 To MSHFlexGrid3.Rows - 1
          If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit For
          MSHFlexGrid3.TextMatrix(I, 3) = Val(MSHFlexGrid3.TextMatrix(I, 2)) - Val(TxtAmount)
      Next
   ElseIf LCase(CmbDiscount.Text) = "percentage" Then
      For I = 1 To MSHFlexGrid3.Rows - 1
          If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit For
          MSHFlexGrid3.TextMatrix(I, 3) = Val(MSHFlexGrid3.TextMatrix(I, 2)) - Val(MSHFlexGrid3.TextMatrix(I, 2)) * (Val(TxtAmount) / 100)
      Next
   End If
End If
End Sub

Private Sub TxtDepoFromBagSize_LostFocus()
If TxtDepoFromBagSize = "" Then Exit Sub
If IsNumeric(TxtDepoFromBagSize) = False Or Val(TxtDepoFromBagSize.Text) <= 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtDepoFromBagSize.SetFocus
   Exit Sub
End If

If InStr(TxtDepoFromBagSize, ".") Then
   MsgBox "Value will get rounded !!!"
End If


TxtDepoFromBagSize = Round(TxtDepoFromBagSize, 0)

End Sub

Private Sub TxtDepoRatePerBag_LostFocus()

If TxtDepoRatePerBag = "" Then Exit Sub
If IsNumeric(TxtDepoRatePerBag) = False Or Val(TxtDepoRatePerBag.Text) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtDepoRatePerBag.SetFocus
   Exit Sub
End If

TxtDepoRatePerBag = Round(TxtDepoRatePerBag, 4)

End Sub

Private Sub TxtDepoToBagSize_LostFocus()

If TxtDepoToBagSize = "" Then Exit Sub
If IsNumeric(TxtDepoToBagSize) = False Or Val(TxtDepoToBagSize.Text) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtDepoToBagSize.SetFocus
   Exit Sub
End If

If InStr(TxtDepoToBagSize, ".") Then
   MsgBox "Value will get rounded !!!"
End If


TxtDepoToBagSize = Round(TxtDepoToBagSize, 0)

End Sub

Private Sub TxtFromBagSize_LostFocus()
If TxtFromBagSize = "" Then Exit Sub
If IsNumeric(TxtFromBagSize) = False Or Val(TxtFromBagSize.Text) <= 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtFromBagSize.SetFocus
   Exit Sub
End If

If InStr(TxtFromBagSize, ".") Then
   MsgBox "Value will get rounded !!!"
End If


TxtFromBagSize = Round(TxtFromBagSize, 0)

End Sub

Private Sub TxtRateOnQty_LostFocus()
If TxtRateOnQty = "" Then Exit Sub
If IsNumeric(TxtRateOnQty) = False Or Val(TxtRateOnQty) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtRateOnQty.SetFocus
   Exit Sub
End If

TxtRateOnQty = Round(TxtRateOnQty, 2)

End Sub

Private Sub TxtRatePerBag_LostFocus()
If TxtRatePerBag = "" Then Exit Sub
If IsNumeric(TxtRatePerBag) = False Or Val(TxtRatePerBag.Text) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtRatePerBag.SetFocus
   Exit Sub
End If

TxtRatePerBag = Round(TxtRatePerBag, 4)

End Sub

Private Sub TxtToBagSize_LostFocus()
If TxtToBagSize = "" Then Exit Sub
If IsNumeric(TxtToBagSize) = False Or Val(TxtToBagSize.Text) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtToBagSize.SetFocus
   Exit Sub
End If

If InStr(TxtToBagSize, ".") Then
   MsgBox "Value will get rounded !!!"
End If


TxtToBagSize = Round(TxtToBagSize, 0)
End Sub

Private Sub CmbBaseRate_LostFocus()
If CmbBaseRate.Text = "" Then Exit Sub
If LCase(CmbBaseRate.Text) <> "yes" And LCase(CmbBaseRate.Text) <> "no" Then
   MsgBox "Invalid selection."
   CmbBaseRate.SetFocus
   Exit Sub
End If
If LCase(CmbBaseRate.Text) = "yes" Then
   mBaseRate = "y"
   CmbDiscount.Enabled = True
   CmbDiscount.SetFocus
   Frame2.Enabled = False
   Frame3.Enabled = False
   TxtAmount = ""
   TxtAmount.Locked = False
Else
   mBaseRate = "n"
   CmbDiscount.Text = ""
   CmbDiscount.Enabled = False
   If LCase(CmbBillingUnit) = "bags" Then
      TxtAmount = ""
      TxtAmount.Locked = True
      Frame2.Enabled = True
      Frame3.Enabled = True
   Else
      TxtAmount = ""
      TxtAmount.Locked = False
      Frame2.Enabled = False
      Frame3.Enabled = False
   End If
End If

Call GetStorageRateFromMaster
Call GetDepoStorageRateFromMaster
End Sub

Private Sub CmbDiscount_LostFocus()
If CmbDiscount.Text = "" Then Exit Sub

If LCase(CmbDiscount.Text) <> "actual" And LCase(CmbDiscount.Text) <> "percentage" Then
   MsgBox "Invalid Selection !!!"
   CmbDiscount.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbClientId_GotFocus()
tmp = CmbClientID.Text
If Check1.Value = vbUnchecked Then
   Call TableMst_Client("Where Converted = 0 And ExchangeTypeID = " & GetExchangeID("Non NCDEX"))
Else
   Call TableMst_Client("Where Converted = 1 And ExchangeTypeID = " & GetExchangeID("Non NCDEX"))
End If
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
mClientIDRow = RsMst_Client!ClientIDRow
End Sub

Private Sub GetStorageRateFromMaster()

Call Grid_Head1

If CmbCommodityGroupID.Text = "" Then Exit Sub
If CmbLocationID.Text = "" Then Exit Sub
If CmbBillingCycle.Text = "" Then Exit Sub
If CmbBillingUnit.Text = "" Then Exit Sub
If CmbBaseRate.Text = "" Then Exit Sub


If LCase(CmbBaseRate.Text) = "yes" Then
   If LCase(CmbBillingUnit.Text) = "bags" Then
      mFlag = "B"
   ElseIf LCase(CmbBillingUnit.Text) = "quantity" Then
      mFlag = "Q"
   ElseIf LCase(CmbBillingUnit.Text) = "sq ft" Then
      mFlag = "S"
   End If
   
   tmp = "Select * From Mst_StorageRateCard "
   tmp = tmp & " Where LocationID = " & mLocationID & " And ExchangeTypeID = 1  "
   tmp = tmp & " And  CommodityGroupID = " & mCommodityGroupId & " And Flag = '" & mFlag & "' "
   tmp = tmp & " And BillingCycleID = " & mBillingCycleID & " And DefaultRate = 1 "
   
   Call TmpTable
   
   If TmpRs.RecordCount = 0 Then
      MsgBox "No record found !!!"
      CmbBaseRate.SetFocus
      mSRCID = Null
      Exit Sub
   End If
    
   TmpRs.MoveFirst
   TxtRateOnQty = Val(TmpRs!RateOnQty)
   mSRCID = TmpRs!SRCID
   If LCase(TmpRs!Flag) = "b" Then
      tmp = "Select * from Mst_StorageRateCardDetail Where SRCID = " & mSRCID & ""
      Call Tmp1Table
      If TmpRs1.RecordCount > 0 Then
         TmpRs1.MoveFirst
         MSHFlexGrid2.Rows = TmpRs1.RecordCount + 2
         For I = 1 To MSHFlexGrid2.Rows - 2
             MSHFlexGrid2.TextMatrix(I, 0) = TmpRs1!FromBagSize
             MSHFlexGrid2.TextMatrix(I, 1) = TmpRs1!ToBagSize
             MSHFlexGrid2.TextMatrix(I, 2) = Format(TmpRs1!RatePerBag, "#####0.0000")
             MSHFlexGrid2.TextMatrix(I, 3) = ""
             MSHFlexGrid2.TextMatrix(I, 4) = TmpRs1!SRCDID
             TmpRs1.MoveNext
         Next
      End If
   End If
Else
   TxtRateOnQty = "" 'Val(TmpRs!RateOnQty)
End If

End Sub

Private Sub CmbDepoBillingCycle_GotFocus()
tmp = CmbDepoBillingCycle.Text
Call TableMst_BillingCycle
CmbDepoBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_BillingCycle.RecordCount
   CmbDepoBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
   RsMst_BillingCycle.MoveNext
Next
CmbDepoBillingCycle.Text = tmp
End Sub

Private Sub CmbDepoBillingCycle_LostFocus()
If CmbDepoBillingCycle.Text = "" Then
   mDepoBillingCycleID = Null
   Exit Sub
End If
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle = '" & CmbDepoBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF Then
   MsgBox "Invalid selection "
   CmbDepoBillingCycle.SetFocus
   Exit Sub
End If
mDepoBillingCycleID = RsMst_BillingCycle!BillingCycleID
TxtDepoTotDays = RsMst_BillingCycle!NoOfDays
Call GetDepoStorageRateFromMaster
End Sub

Private Sub GetDepoStorageRateFromMaster()

Call Grid_Head2

If CmbCommodityGroupID.Text = "" Then Exit Sub
If CmbLocationID.Text = "" Then Exit Sub
If CmbDepoBillingCycle.Text = "" Then Exit Sub
If CmbBillingUnit.Text = "" Then Exit Sub
If CmbBaseRate.Text = "" Then Exit Sub
'If CmbDepoBillingCycle.Text = "" Then Exit Sub


If LCase(CmbBaseRate.Text) = "yes" Then
   If LCase(CmbBillingUnit.Text) = "bags" Then
      mFlag = "B"
   ElseIf LCase(CmbBillingUnit.Text) = "quantity" Then
      mFlag = "Q"
   ElseIf LCase(CmbBillingUnit.Text) = "sq ft" Then
      mFlag = "S"
   End If
   
   tmp = "Select * From Mst_StorageRateCard "
   tmp = tmp & " Where LocationID = " & mLocationID & " And ExchangeTypeID = 1  "
   tmp = tmp & " And  CommodityGroupID = " & mCommodityGroupId & " And Flag = '" & mFlag & "' "
   tmp = tmp & " And DepoBillingCycleID = " & mDepoBillingCycleID & " And DefaultRate = 1 "
   
   Call TmpTable
   
   If TmpRs.RecordCount = 0 Then
      MsgBox "No record found !!!"
      CmbBaseRate.SetFocus
      mSRCID = Null
      Exit Sub
   End If
    
   TmpRs.MoveFirst
   TxtRateOnQty = Val(TmpRs!RateOnQty)
   mSRCID = TmpRs!SRCID
   If LCase(TmpRs!Flag) = "b" Then
      tmp = "Select * from Mst_StorageRateCardDepoDetail Where SRCID = " & mSRCID & ""
      Call Tmp1Table
      If TmpRs1.RecordCount > 0 Then
         TmpRs1.MoveFirst
         MSHFlexGrid3.Rows = TmpRs1.RecordCount + 2
         For I = 1 To MSHFlexGrid3.Rows - 2
             MSHFlexGrid3.TextMatrix(I, 0) = TmpRs1!FromBagSize
             MSHFlexGrid3.TextMatrix(I, 1) = TmpRs1!ToBagSize
             MSHFlexGrid3.TextMatrix(I, 2) = Format(TmpRs1!RatePerBag, "#####0.0000")
             MSHFlexGrid3.TextMatrix(I, 3) = ""
             MSHFlexGrid3.TextMatrix(I, 4) = TmpRs1!SRCDID
             TmpRs1.MoveNext
         Next
      End If
   End If
Else
   TxtRateOnQty = "" 'Val(TmpRs!RateOnQty)
End If

End Sub
