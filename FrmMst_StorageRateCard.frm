VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_StorageRateCard 
   Caption         =   "Storage Rate Card"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   5055
      Left            =   360
      TabIndex        =   46
      Top             =   3840
      Visible         =   0   'False
      Width           =   12480
      Begin VB.CheckBox ChkUpdated 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9750
         TabIndex        =   104
         Top             =   780
         Width           =   210
      End
      Begin VB.ComboBox CmbSearchDefault 
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
         ItemData        =   "FrmMst_StorageRateCard.frx":0000
         Left            =   13725
         List            =   "FrmMst_StorageRateCard.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   77
         Top             =   240
         Width           =   960
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   72
         Top             =   765
         Width           =   2880
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
         ItemData        =   "FrmMst_StorageRateCard.frx":0017
         Left            =   5670
         List            =   "FrmMst_StorageRateCard.frx":0024
         Sorted          =   -1  'True
         TabIndex        =   71
         Top             =   735
         Width           =   2880
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13485
         TabIndex        =   52
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   350
         Left            =   12480
         TabIndex        =   51
         Top             =   1200
         Width           =   1000
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
         Left            =   9750
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   240
         Width           =   2880
      End
      Begin VB.ComboBox CmbSExchangeTypeID 
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
         ItemData        =   "FrmMst_StorageRateCard.frx":003F
         Left            =   5670
         List            =   "FrmMst_StorageRateCard.frx":0041
         Sorted          =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   2880
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   240
         Width           =   2865
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   855
         Left            =   120
         TabIndex        =   47
         Top             =   1680
         Width           =   14415
         _ExtentX        =   25426
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
         Left            =   10965
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   22544385
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtUpdatedTo 
         Height          =   375
         Left            =   13365
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   22544385
         CurrentDate     =   39884
      End
      Begin VB.Label Label27 
         Caption         =   "Only Updated"
         Height          =   450
         Left            =   8760
         TabIndex        =   103
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label26 
         Caption         =   "Updated To"
         Height          =   240
         Left            =   12480
         TabIndex        =   102
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Updated From"
         Height          =   405
         Left            =   10200
         TabIndex        =   100
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Default"
         Height          =   210
         Left            =   12975
         TabIndex        =   76
         Top             =   315
         Width           =   570
      End
      Begin VB.Label Label16 
         Caption         =   "Billing Cycle"
         Height          =   225
         Left            =   120
         TabIndex        =   74
         Top             =   833
         Width           =   960
      End
      Begin VB.Label Label15 
         Caption         =   "Billing Unit"
         Height          =   240
         Left            =   4080
         TabIndex        =   73
         Top             =   795
         Width           =   1605
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
         Left            =   8040
         TabIndex        =   57
         Top             =   1200
         Width           =   90
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
         Left            =   5640
         TabIndex        =   56
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Commodity"
         Height          =   240
         Left            =   8760
         TabIndex        =   55
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Location"
         Height          =   210
         Left            =   120
         TabIndex        =   54
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Exchange Type"
         Height          =   240
         Left            =   4080
         TabIndex        =   53
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7860
      Begin VB.CommandButton CmdDepositBillingCycle 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4470
         TabIndex        =   82
         Top             =   2820
         Width           =   375
      End
      Begin VB.ComboBox CmbDepositBillingCycle 
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
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   2790
         Width           =   3225
      End
      Begin VB.TextBox TxtDepositTotDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   5985
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   2790
         Width           =   990
      End
      Begin VB.CheckBox ChkDefault 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   2400
         Width           =   225
      End
      Begin VB.TextBox TxtTotDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   5970
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1935
         Width           =   990
      End
      Begin VB.TextBox TxtRateOnQty 
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
         TabIndex        =   8
         Top             =   2355
         Width           =   990
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
         ItemData        =   "FrmMst_StorageRateCard.frx":0043
         Left            =   2760
         List            =   "FrmMst_StorageRateCard.frx":0050
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2370
         Width           =   1650
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
         Top             =   1935
         Width           =   3225
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
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1935
         Width           =   3210
      End
      Begin VB.CommandButton CmdBillingCycle 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4455
         TabIndex        =   59
         Top             =   1965
         Width           =   375
      End
      Begin VB.TextBox TxtPreFix 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   5865
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton CmdCommodityGroupID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7005
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1500
         Width           =   375
      End
      Begin VB.CommandButton CmdLocationID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7005
         TabIndex        =   34
         Top             =   630
         Width           =   375
      End
      Begin VB.TextBox TxtSRCID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   195
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   1200
         TabIndex        =   16
         Top             =   3705
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Format          =   22544385
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   3975
         TabIndex        =   17
         Top             =   3705
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Format          =   22544385
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtTillDate 
         Height          =   330
         Left            =   6945
         TabIndex        =   18
         Top             =   3705
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Format          =   22544385
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtSRCDate 
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   22544385
         CurrentDate     =   39884
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
         TabIndex        =   2
         Top             =   600
         Width           =   5760
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         ItemData        =   "FrmMst_StorageRateCard.frx":006B
         Left            =   1200
         List            =   "FrmMst_StorageRateCard.frx":006D
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1050
         Width           =   5760
      End
      Begin VB.TextBox TxtExchangeTypeID 
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
         TabIndex        =   39
         Top             =   1050
         Width           =   5760
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   600
         Width           =   5760
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
         Top             =   1485
         Width           =   5760
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1485
         Width           =   5760
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2370
         Width           =   1650
      End
      Begin VB.TextBox TxtDepositBillingCycle 
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
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   2790
         Width           =   3210
      End
      Begin VB.Label Label20 
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
         Height          =   435
         Left            =   120
         TabIndex        =   84
         Top             =   2760
         Width           =   1050
      End
      Begin VB.Label Label19 
         Caption         =   "Total Days"
         Height          =   210
         Left            =   4905
         TabIndex        =   83
         Top             =   2865
         Width           =   780
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Default "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   75
         Top             =   2430
         Width           =   675
      End
      Begin VB.Label Label14 
         Caption         =   "Total Days"
         Height          =   210
         Left            =   4890
         TabIndex        =   65
         Top             =   2010
         Width           =   780
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
         Left            =   1800
         TabIndex        =   63
         Top             =   2430
         Width           =   1005
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
         Height          =   435
         Left            =   105
         TabIndex        =   61
         Top             =   1905
         Width           =   960
      End
      Begin VB.Label Label6 
         Caption         =   "SM_Prefix"
         Height          =   180
         Left            =   5130
         TabIndex        =   45
         Top             =   240
         Width           =   780
      End
      Begin VB.Label LblRatePerBag 
         Caption         =   "Rate on Units"
         Height          =   240
         Left            =   4890
         TabIndex        =   42
         Top             =   2415
         Width           =   1080
      End
      Begin VB.Label Label7 
         Caption         =   "Exchange Type"
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
         Left            =   90
         TabIndex        =   41
         Top             =   1020
         Width           =   1215
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
         TabIndex        =   40
         Top             =   1470
         Width           =   1215
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
         TabIndex        =   36
         Top             =   668
         Width           =   960
      End
      Begin VB.Label LblWHRNo 
         Caption         =   "SRC ID"
         Height          =   210
         Left            =   90
         TabIndex        =   33
         Top             =   277
         Width           =   900
      End
      Begin VB.Label LblWHRDate 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   32
         Top             =   3780
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Till Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5970
         TabIndex        =   31
         Top             =   3780
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2970
         TabIndex        =   30
         Top             =   3780
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label41 
         Caption         =   "SRC Date"
         Height          =   225
         Left            =   2715
         TabIndex        =   29
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1110
      Left            =   120
      TabIndex        =   43
      Top             =   7980
      Width           =   7890
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1710
         TabIndex        =   15
         Top             =   195
         Width           =   1500
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   210
         TabIndex        =   58
         Top             =   195
         Width           =   1500
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   210
         TabIndex        =   24
         Top             =   615
         Width           =   1500
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1710
         TabIndex        =   25
         Top             =   615
         Width           =   1500
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   3210
         TabIndex        =   22
         Top             =   195
         Width           =   1500
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   3210
         TabIndex        =   26
         Top             =   615
         Width           =   1500
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4710
         TabIndex        =   23
         Top             =   195
         Width           =   1500
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4710
         TabIndex        =   27
         Top             =   615
         Width           =   1500
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   6210
         TabIndex        =   20
         Top             =   195
         Width           =   1380
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   6210
         TabIndex        =   21
         Top             =   615
         Width           =   1380
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4620
      Left            =   120
      TabIndex        =   66
      Top             =   3360
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   8149
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
      TabPicture(0)   =   "FrmMst_StorageRateCard.frx":006F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Deposit Rate Details As per Bags"
      TabPicture(1)   =   "FrmMst_StorageRateCard.frx":008B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "MSHFlexGrid4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Log History"
      TabPicture(2)   =   "FrmMst_StorageRateCard.frx":00A7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtUpdated"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "TxtCreated"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label12"
      Tab(2).Control(3)=   "Label13"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   -74880
         TabIndex        =   89
         Top             =   360
         Width           =   7335
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
            TabIndex        =   94
            Top             =   495
            Width           =   1815
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
            TabIndex        =   93
            Top             =   495
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
            TabIndex        =   92
            Top             =   495
            Width           =   1215
         End
         Begin VB.CommandButton CmdDepoAdd2List 
            Caption         =   "Add to List"
            Height          =   375
            Left            =   4440
            TabIndex        =   91
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton CmdDepoDeleteList 
            Caption         =   "Delete from List"
            Height          =   375
            Left            =   5670
            TabIndex        =   90
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label24 
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
            TabIndex        =   97
            Top             =   120
            Width           =   1860
         End
         Begin VB.Label Label22 
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
            TabIndex        =   96
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label21 
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
            TabIndex        =   95
            Top             =   120
            Width           =   1020
         End
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
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1185
         Width           =   6255
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
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   720
         Width           =   6255
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   67
         Top             =   330
         Width           =   7710
         Begin VB.CommandButton CmdDeleteList 
            Caption         =   "Delete from List"
            Height          =   375
            Left            =   5670
            TabIndex        =   14
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton CmdAdd2List 
            Caption         =   "Add to List"
            Height          =   375
            Left            =   4440
            TabIndex        =   13
            Top             =   480
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
            TabIndex        =   11
            Top             =   495
            Width           =   1215
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
            TabIndex        =   10
            Top             =   495
            Width           =   1215
         End
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
            TabIndex        =   12
            Top             =   495
            Width           =   1815
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
            TabIndex        =   70
            Top             =   120
            Width           =   1020
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
            TabIndex        =   69
            Top             =   120
            Width           =   1020
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
            TabIndex        =   68
            Top             =   120
            Width           =   1860
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3135
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   5530
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   98
         Top             =   1350
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5530
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
         TabIndex        =   88
         Top             =   1245
         Width           =   885
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
         TabIndex        =   87
         Top             =   780
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Storage Rate Card Details"
      Height          =   9015
      Left            =   8040
      TabIndex        =   78
      Top             =   0
      Width           =   7095
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   8655
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   15266
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
   End
End
Attribute VB_Name = "FrmMst_StorageRateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mExchangeTypeID, mFlag, mCommodityGroupId, mLocationID, mGodownID, mCMPID As Variant

Dim mSCommodityID, mSLocationID, mSExchangeTypeID, mBillingCycleID, mSBillingCycleID, mSFlag, mSearchDefault As Variant
Dim mDepositBillingCycleID As Variant


Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 17


MSHFlexGrid1.TextMatrix(0, 0) = "SRC ID"
MSHFlexGrid1.TextMatrix(0, 1) = "SRC Date"
MSHFlexGrid1.TextMatrix(0, 2) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 3) = "Location "
MSHFlexGrid1.TextMatrix(0, 4) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 5) = "Commodity Group"
MSHFlexGrid1.TextMatrix(0, 6) = "Billing Cycle"
MSHFlexGrid1.TextMatrix(0, 7) = "Billing Unit"
MSHFlexGrid1.TextMatrix(0, 8) = "Rate on Units"
MSHFlexGrid1.TextMatrix(0, 9) = "Deposit Billing Cycle"
MSHFlexGrid1.TextMatrix(0, 10) = "From Date"
MSHFlexGrid1.TextMatrix(0, 11) = "To Date"
MSHFlexGrid1.TextMatrix(0, 12) = "Till Date"
MSHFlexGrid1.TextMatrix(0, 13) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 14) = "Default"
MSHFlexGrid1.TextMatrix(0, 15) = "Updated By"
MSHFlexGrid1.TextMatrix(0, 16) = "Updated Date"

MSHFlexGrid1.RowHeight(0) = 800

MSHFlexGrid1.ColWidth(0) = 600
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 0
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 1600
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1600
MSHFlexGrid1.ColWidth(10) = 0 '1000
MSHFlexGrid1.ColWidth(11) = 0 '1000
MSHFlexGrid1.ColWidth(12) = 0 '1000
MSHFlexGrid1.ColWidth(13) = 0
MSHFlexGrid1.ColWidth(14) = 800
MSHFlexGrid1.ColWidth(15) = 1000
MSHFlexGrid1.ColWidth(16) = 1500

MSHFlexGrid1.RowHeight(0) = 400


End Sub

Private Sub CmbBillingCycle_GotFocus()
tmp = CmbBillingCycle.Text
Call TableMst_BillingCycle
CmbBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_BillingCycle.RecordCount
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
End Sub

Private Sub CmbDepositBillingCycle_GotFocus()

tmp = CmbDepositBillingCycle.Text
Call TableMst_BillingCycle
CmbDepositBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_BillingCycle.RecordCount
   CmbDepositBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
   RsMst_BillingCycle.MoveNext
Next
CmbDepositBillingCycle.Text = tmp


End Sub

Private Sub CmbDepositBillingCycle_LostFocus()
If CmbDepositBillingCycle.Text = "" Then
   mDepositBillingCycleID = Null
   Exit Sub
End If
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle = '" & CmbDepositBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF Then
   MsgBox "Invalid selection "
   CmbDepositBillingCycle.SetFocus
   Exit Sub
End If
mDepositBillingCycleID = RsMst_BillingCycle!BillingCycleID
TxtDepositTotDays = RsMst_BillingCycle!NoOfDays
End Sub

Private Sub CmbSBillingCycle_GotFocus()
tmp = CmbSBillingCycle.Text
Call TableMst_BillingCycle
CmbSBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_BillingCycle.RecordCount
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
   Frame2.Enabled = True
   Frame4.Enabled = True
   TxtRateOnQty = ""
   TxtRateOnQty.Enabled = False
   If MSHFlexGrid2.Rows <= 2 Then
      Call Grid_Head1
      Call Grid_HeadDepoBillingCycle
      Call InDataBagSize
      Call IndataDepoBagSize
   End If

Else
   If MSHFlexGrid2.Rows > 2 Then
      If MsgBox("Data in Bag Size Details will be lost. Are you want to Continue.", vbYesNo) = vbNo Then
         CmbBillingUnit.SetFocus
         Exit Sub
      End If
   End If
   Call Grid_Head1
   Call Grid_HeadDepoBillingCycle
   Call ClearDetail
   Call ClearDepoDetail
   Frame2.Enabled = False
   Frame4.Enabled = False
   TxtRateOnQty.Enabled = True
   TxtRateOnQty.SetFocus
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

Private Sub CmbSearchDefault_LostFocus()
If CmbSearchDefault.Text = "" Then
   mSearchDefault = Null
   Exit Sub
End If
If LCase(CmbSearchDefault.Text) <> "yes" And LCase(CmbSearchDefault.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSearchDefault.SetFocus
   Exit Sub
End If

If LCase(CmbSearchDefault.Text) = "yes" Then
   mSearchDefault = 1
ElseIf LCase(CmbSearchDefault.Text) = "no" Then
   mSearchDefault = 0
End If

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
   MSHFlexGrid2.TextMatrix(r, 3) = "" 'Val(TxtToBagSize)
End If

MSHFlexGrid2.TextMatrix(r, 0) = Val(TxtFromBagSize)
MSHFlexGrid2.TextMatrix(r, 1) = Val(TxtToBagSize)
MSHFlexGrid2.TextMatrix(r, 2) = Format(Val(TxtRatePerBag.Text), "#####0.0000")


Call ClearDetail
TxtFromBagSize.SetFocus

End Sub

Private Sub CmdBillingCycle_Click()
FrmMst_BillingCycle.Show
End Sub
Private Sub ClearDetail()
TxtFromBagSize = ""
TxtToBagSize = ""
TxtRatePerBag = ""
CmdAdd2List.Caption = "Add to List"

TxtDepoFromBagSize = ""
TxtDepoToBagSize = ""
TxtDepoRatePerBag = ""
CmdDepoAdd2List.Caption = "Add to List"

End Sub
Private Sub ClearDepoDetail()

TxtDepoFromBagSize = ""
TxtDepoToBagSize = ""
TxtDepoRatePerBag = ""
CmdDepoAdd2List.Caption = "Add to List"

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
If MSHFlexGrid4.Rows > 1 Then
   For u = 1 To MSHFlexGrid4.Rows - 1
       If (Val(TxtDepoFromBagSize) >= Val(MSHFlexGrid4.TextMatrix(u, 0)) And Val(TxtDepoFromBagSize) <= Val(MSHFlexGrid4.TextMatrix(u, 1))) Or Val(TxtDepoToBagSize) >= Val(MSHFlexGrid4.TextMatrix(u, 0)) And Val(TxtDepoToBagSize) <= Val(MSHFlexGrid4.TextMatrix(u, 1)) Or (Val(TxtDepoFromBagSize) <= Val(MSHFlexGrid4.TextMatrix(u, 0)) And Val(TxtDepoToBagSize) >= Val(MSHFlexGrid4.TextMatrix(u, 1))) Then
          If LCase(CmdDepoAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid4.RowSel Then
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
MSHFlexGrid2.Cols = 4

MSHFlexGrid2.TextMatrix(0, 0) = "Bag Size From"
MSHFlexGrid2.TextMatrix(0, 1) = "Bag Size To"
MSHFlexGrid2.TextMatrix(0, 2) = "Rate Per Bag"
MSHFlexGrid2.TextMatrix(0, 3) = "SRCDID"

MSHFlexGrid2.ColWidth(0) = 1400
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 1800
MSHFlexGrid2.ColWidth(3) = 0

End Sub
Private Sub Grid_HeadDepoBillingCycle()

MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 2
MSHFlexGrid4.FixedRows = 1
MSHFlexGrid4.FixedCols = 0
MSHFlexGrid4.Cols = 4

MSHFlexGrid4.TextMatrix(0, 0) = "Bag Size From"
MSHFlexGrid4.TextMatrix(0, 1) = "Bag Size To"
MSHFlexGrid4.TextMatrix(0, 2) = "Rate Per Bag"
MSHFlexGrid4.TextMatrix(0, 3) = "DepoSRCDID"

MSHFlexGrid4.ColWidth(0) = 1400
MSHFlexGrid4.ColWidth(1) = 1400
MSHFlexGrid4.ColWidth(2) = 1800
MSHFlexGrid4.ColWidth(3) = 0

End Sub

Private Sub CmdDeleteList_Click()
'i = MSHFlexGrid2.RowSel
'If i > 0 Then
'   If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
'      For RO = i To MSHFlexGrid2.Rows - 2
'          For c = 0 To MSHFlexGrid2.Cols - 1
'              MSHFlexGrid2.TextMatrix(RO, c) = MSHFlexGrid2.TextMatrix(RO + 1, c)
'          Next
'      Next
'      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
'   End If
'End If
'Call ClearDetail
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
   TxtToBagSize.SetFocus
   Exit Sub
End If

If TxtDepoRatePerBag.Text = "" Then
   MsgBox "Blank Rate Per Bag not Allowed!!!"
   TxtDepoRatePerBag.SetFocus
   Exit Sub
End If

If CheckDepoDuplicateBags = True Then
   MsgBox "'Bag Size From' or 'Bag Size To' in already used in range. Please change range!!! "
   TxtDepoFromBagSize.SetFocus
   Exit Sub
End If

If CmdDepoAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid4.RowSel
Else
   r = MSHFlexGrid4.Rows - 1
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
   MSHFlexGrid4.TextMatrix(r, 3) = "" 'Val(TxtToBagSize)
End If

MSHFlexGrid4.TextMatrix(r, 0) = Val(TxtDepoFromBagSize)
MSHFlexGrid4.TextMatrix(r, 1) = Val(TxtDepoToBagSize)
MSHFlexGrid4.TextMatrix(r, 2) = Format(Val(TxtDepoRatePerBag.Text), "#####0.0000")


Call ClearDepoDetail
TxtDepoFromBagSize.SetFocus



End Sub

Private Sub CmdDepositBillingCycle_Click()
FrmMst_BillingCycle.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_StorageRateCard.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_StorageRateCard.xls")
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim wc As String
Call Grid_Head
wc = ""


If ChkUpdated.Value = 1 Then
   wc = " Where MSRC.UpdatedDate between '" & Format(TxtUpdatedFrom.Value, Gen_DatFormat) & "' and '" & Format(TxtUpdatedTo.Value + 1, Gen_DatFormat) & "'  "
Else
   wc = " Where (MSRC.UpdatedDate between '" & Format(TxtUpdatedFrom.Value, Gen_DatFormat) & "' and '" & Format(TxtUpdatedTo.Value + 1, Gen_DatFormat) & "' Or (MSRC.UpdatedDate is null)) "
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
   wc = GenWc(wc, CmbSCommodityID.Text, "MCG.CommodityGroup", "S", "N")
End If
If CmbSExchangeTypeID.Text <> "" Then
   wc = GenWc(wc, CmbSExchangeTypeID.Text, "ME.ExchangeType", "S", "N")
End If
If CmbSBillingCycle.Text <> "" Then
   wc = GenWc(wc, CmbSBillingCycle.Text, "MB.BillingCycle", "S", "N")
End If

If CmbSBillingUnit.Text <> "" Then
   wc = GenWc(wc, Left(CmbSBillingUnit.Text, 1), "MSRC.Flag", "S", "N")
End If

If CmbSearchDefault <> "" Then
   wc = GenWc(wc, mSearchDefault, "MSRC.DefaultRate", "S", "N")
End If



tmp = "Select MSRC.SRCID,MSRC.SRCDate,MSRC.SM_Prefix,ML.LocationName,ME.ExchangeType,MCG.CommodityGroup,"
tmp = tmp & " MB.BillingCycle ,MSRC.Flag, MSRC.RateonQty,MB1.BillingCycle as DepoBillingCycle,MSRC.FromDate, MSRC.ToDate, MSRC.TillDate, MSRC.G_UID,MSRC.DefaultRate,MSRC.UpdatedBy,MSRC.UpdatedDate"
tmp = tmp & " From Mst_StorageRateCard MSRC"
tmp = tmp & " Inner Join Mst_Location ML On MSRC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME ON MSRC.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MSRC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Left Join Mst_BillingCycle MB On MSRC.BillingCycleID = MB.BillingCycleID "
tmp = tmp & " Left Join Mst_BillingCycle MB1 On MSRC.DepoBillingCycleID = MB1.BillingCycleID "

tmp = tmp & wc

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For i = 1 To TmpRs.RecordCount
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid1.TextMatrix(i, 1) = Format(MSHFlexGrid1.TextMatrix(i, 1), "dd-mmm-yyyy")
       If LCase(MSHFlexGrid1.TextMatrix(i, 7)) = "b" Then
          MSHFlexGrid1.TextMatrix(i, 7) = "Bags"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(i, 7)) = "q" Then
          MSHFlexGrid1.TextMatrix(i, 7) = "Quantity"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(i, 7)) = "s" Then
          MSHFlexGrid1.TextMatrix(i, 7) = "Sq Ft"
       Else
          MSHFlexGrid1.TextMatrix(i, 7) = ""
       End If
       MSHFlexGrid1.TextMatrix(i, 10) = "" 'Format(MSHFlexGrid1.TextMatrix(i, 9), "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(i, 11) = "" 'Format(MSHFlexGrid1.TextMatrix(i, 10), "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(i, 12) = "" 'Format(MSHFlexGrid1.TextMatrix(i, 11), "dd-mmm-yyyy")
       If MSHFlexGrid1.TextMatrix(i, 14) = 0 Then
          MSHFlexGrid1.TextMatrix(i, 14) = "No"
       Else
          MSHFlexGrid1.TextMatrix(i, 14) = "Yes"
       End If
       If MSHFlexGrid1.TextMatrix(i, 16) <> "" Then
          MSHFlexGrid1.TextMatrix(i, 16) = Format(MSHFlexGrid1.TextMatrix(i, 16), "dd-mmm-yyyy hh:mm:ss")
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
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_StorageRateCard.MoveFirst
RsMst_StorageRateCard.Find "SRCID = " & MSHFlexGrid1.TextMatrix(i, 0) & ""
If Not RsMst_StorageRateCard.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
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
Dim i As Variant
i = MSHFlexGrid3.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid3.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_StorageRateCard.MoveFirst
RsMst_StorageRateCard.Find "SRCID = " & MSHFlexGrid3.TextMatrix(i, 0) & ""
If Not RsMst_StorageRateCard.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid4_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid4.RowSel
If r < 1 Or r = MSHFlexGrid4.Rows - 1 Then
   Call ClearDepoDetail
   Exit Sub
End If
If MSHFlexGrid4.TextMatrix(r, 0) = "" Then
   Exit Sub
End If

TxtDepoFromBagSize = MSHFlexGrid4.TextMatrix(r, 0)
TxtDepoToBagSize = MSHFlexGrid4.TextMatrix(r, 1)
TxtDepoRatePerBag = MSHFlexGrid4.TextMatrix(r, 2)
CmdDepoAdd2List.Caption = "Update List"
TxtDepoFromBagSize.SetFocus
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_StorageRateCard.MoveNext
If RsMst_StorageRateCard.EOF Then
   RsMst_StorageRateCard.MoveLast
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
RsMst_StorageRateCard.MovePrevious
If RsMst_StorageRateCard.BOF Then
   RsMst_StorageRateCard.MoveFirst
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
RsMst_StorageRateCard.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_StorageRateCard.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
Frame2.Enabled = True
Frame4.Enabled = True
'CmbLocationID.Visible = True
'CmbCommodityGroupID.Visible = True
'CmbExchangeTypeID.Visible = True

CmbLocationID.Text = TxtLocationID.Text
CmbExchangeTypeID.Text = TxtExchangeTypeID.Text
CmbCommodityGroupID.Text = TxtCommodityGroupID.Text

'CmdCommodityGroupID.Visible = True
'CmdLocationID.Visible = True

'CmdBillingCycle.Visible = True

'CmbBillingCycle.Visible = True
'CmbBillingUnit.Visible = True


CmbDepositBillingCycle.Visible = True
TxtDepositBillingCycle.Visible = False


CmbBillingCycle.Text = TxtBillingCycle.Text
CmbBillingUnit.Text = TxtBillingUnit.Text
CmbDepositBillingCycle.Text = TxtDepositBillingCycle

If LCase(CmbBillingUnit.Text) = "bags" Then
   TxtRateOnQty.Enabled = False
   Frame2.Enabled = True
   Frame4.Enabled = True
Else
   TxtRateOnQty.Enabled = True
   Frame2.Enabled = False
   Frame4.Enabled = False
End If

Call GButtonLock(Me, False)

End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError

tmp = "Select * from Mst_SpecialRateCard Where SRCID = " & TxtSRCID & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   MsgBox "Storage Rate Card used in Special Rate Card!!!"
   Exit Sub
End If

Ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If Ans = vbYes Then
   Call DeleteChildTableDetail
   RsMst_StorageRateCard.Delete
   RsMst_StorageRateCard.Update
   RsMst_StorageRateCard.Requery
   If RsMst_StorageRateCard.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_StorageRateCard.MoveNext
   If RsMst_StorageRateCard.EOF() Then
      RsMst_StorageRateCard.MoveLast
   End If
   Call InData
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_StorageRateCard.CancelUpdate
End Sub
Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   If Not RsMst_StorageRateCard.EOF Then
      RsMst_StorageRateCard.MoveFirst
      RsMst_StorageRateCard.Find "SRCID = " & TxtSRCID & ""
      Call InData
   End If
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 300
Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1800

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'ChkUpdated.Value = 0
'TxtUpdatedFrom.Value = "01/01/2008"
'TxtUpdatedTo.Value = Date
'CmbSCommodityID.Text = ""
'CmbSLocationID.Text = ""
'CmbSExchangeTypeID.Text = ""
'LblTotalRecords.Caption = ""
'CmbSBillingCycle.Text = ""
'CmbSBillingUnit.Text = ""
'CmbSearchDefault.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Call Grid_Head1
Frame1.Enabled = True
Frame2.Enabled = True
TxtSRCDate.Enabled = True

TxtLocationID.Visible = False
TxtCommodityGroupID.Visible = False
TxtExchangeTypeID.Visible = False
TxtBillingCycle.Visible = False
TxtBillingUnit.Visible = False
TxtDepositBillingCycle.Visible = False

CmbLocationID.Visible = True
CmbExchangeTypeID.Visible = True
CmbCommodityGroupID.Visible = True
CmbBillingCycle.Visible = True
CmbBillingUnit.Visible = True
CmbDepositBillingCycle.Visible = True

CmdLocationID.Visible = True
CmdCommodityGroupID.Visible = True
CmdDepositBillingCycle.Visible = True
CmdBillingCycle.Visible = True


CmbLocationID.Text = ""
CmbExchangeTypeID.Text = ""
CmbCommodityGroupID.Text = ""
CmbBillingCycle.Text = ""
CmbBillingUnit.Text = ""
CmbDepositBillingCycle.Text = ""
'ChkDefault.Value = 0
ChkDefault.Value = 0
ChkDefault.Enabled = True


TxtSRCID.Text = ""

TxtSRCDate.Value = Date

'Commneted on 15th April at 11:59 am . Backup in Store-Man Backup folder RAR file 'Store-Man with SRC and SContract'
'TxtFromDate.Value = Date
'TxtToDate.Value = Date
'TxtTillDate.Value = Date



TxtToBagSize.Text = ""
TxtRatePerBag.Text = ""
TxtFromBagSize.Text = ""
TxtPreFix = ""
TxtRateOnQty = ""
TxtTotDays = ""
TxtCreated = ""
TxtUpdated = ""
Call ClearDetail
Call ClearDepoDetail
Call Grid_Head1
Call Grid_Head2
Call Grid_HeadDepoBillingCycle

If RsMst_StorageRateCard.RecordCount > 0 Then
   TxtSRCDate.SetFocus
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
For i = 1 To RsMst_CommodityGroup.RecordCount
   CmbSCommodityID.AddItem RsMst_CommodityGroup!CommodityGroup
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
For i = 1 To RsMst_CommodityGroup.RecordCount
   CmbCommodityGroupID.AddItem RsMst_CommodityGroup!CommodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbCommodityGroupID.Text = tmp
    
End Sub

Private Sub CmbCommodityGroupID_LostFocus()
If CmbCommodityGroupID.Text = "" Then
   mCommodityGroupId = Null
   Call Grid_Head2
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
Call GetSRCGridData

End Sub
Private Sub CmbsExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbSExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
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
Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = Null
   Call Grid_Head2
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
Call GetSRCGridData

'If mExchangeTypeID = 2 Then 'NCDEX
'   OptRatePerBag.Enabled = True
'   OptRatePerQuantity.Enabled = True
'   OptRatePerBag.Value = True
'Else
'   OptRatePerBag.Value = True
'   OptRatePerBag.Enabled = False
'   OptRatePerQuantity.Enabled = False
'End If

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
For i = 1 To RsMst_Location.RecordCount
   CmbLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
    
End Sub

'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   TxtPreFix = ""
   Call Grid_Head2
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
TxtPreFix = RsMst_Location!SM_Prefix
Call GetSRCGridData
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
For i = 1 To RsMst_Location.RecordCount
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
'Dim FrmLoadAccess() As Boolean
'FrmLoadAccess = GetFrmLoadAccess(Me.Name)
'If FrmLoadAccess(0) = False Then
'   Call GButtonLockAD(Me)
'   Frame1.Enabled = False
'   MsgBox "Access denied !!!!!!!!!"
'   Exit Sub
'End If
'Call Grid_Head
'ChkUpdated.Value = 0
'TxtUpdatedFrom.Value = "01/01/2008"
'TxtUpdatedTo.Value = Date
'CmbSCommodityID.Text = ""
'CmbSLocationID.Text = ""
'CmbSExchangeTypeID.Text = ""
'LblTotalRecords.Caption = ""
'CmbSBillingCycle.Text = ""
'CmbSBillingUnit.Text = ""
'CmbSearchDefault.Text = ""
'
'Frame1.Enabled = True
'If Gen_WcLocation = "" Then
'   Call TableMst_StorageCardRate
'Else
'   Call TableMst_StorageCardRate(" Where " & Gen_WcLocation)
'End If
'If RsMst_StorageRateCard.RecordCount = 0 Then
'   Call AddCmd_Click
'   Exit Sub
'End If
'
'Call GButtonLock(Me, True)
'
'Call Indata

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
ChkUpdated.Value = 0
TxtUpdatedFrom.Value = "01/01/2008"
TxtUpdatedTo.Value = Date
CmbSCommodityID.Text = ""
CmbSLocationID.Text = ""
CmbSExchangeTypeID.Text = ""
LblTotalRecords.Caption = ""
CmbSBillingCycle.Text = ""
CmbSBillingUnit.Text = ""
CmbSearchDefault.Text = ""

Frame1.Enabled = True
If Gen_WcLocation = "" Then
   Call TableMst_StorageCardRate
Else
   Call TableMst_StorageCardRate(" Where " & Gen_WcLocation)
End If
If RsMst_StorageRateCard.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call GButtonLock(Me, True)
If Val(Gen_GodownInspection(0)) > 0 Then
   RsMst_StorageRateCard.Filter = "LocationID = " & Val(Gen_GodownInspection(0)) & " And CommodityGroupID = " & Val(Gen_GodownInspection(1))
   If RsMst_StorageRateCard.RecordCount > 0 Then
      RsMst_StorageRateCard.MoveLast
      tmp = RsMst_StorageRateCard!SRCID
      RsMst_StorageRateCard.Filter = ""
      RsMst_StorageRateCard.MoveFirst
      RsMst_StorageRateCard.Find "SRCID = " & tmp
   Else
      RsMst_StorageRateCard.Filter = ""
   End If
End If

Call InData

End Sub
Private Sub InData()

Call Grid_Head1
Call Grid_Head2
Call Grid_HeadDepoBillingCycle

Frame1.Enabled = False
Frame2.Enabled = False
Frame4.Enabled = False

CmbLocationID.Visible = False
CmbExchangeTypeID.Visible = False
CmbCommodityGroupID.Visible = False

TxtLocationID.Visible = True
TxtExchangeTypeID.Visible = True
TxtCommodityGroupID.Visible = True
TxtBillingCycle.Visible = True
TxtBillingUnit.Visible = True
TxtDepositBillingCycle.Visible = True

CmdLocationID.Visible = False
CmdCommodityGroupID.Visible = False
CmbBillingCycle.Visible = False
CmbBillingUnit.Visible = False
CmbDepositBillingCycle.Visible = False

CmdBillingCycle.Visible = False
CmdDepositBillingCycle.Visible = False

mLocationID = RsMst_StorageRateCard!LocationID
mCommodityGroupId = RsMst_StorageRateCard!CommodityGroupID
mExchangeTypeID = RsMst_StorageRateCard!ExchangeTypeID
mBillingCycleID = RsMst_StorageRateCard!BillingCycleID
mDepositBillingCycleID = RsMst_StorageRateCard!DepoBillingCycleID

TxtSRCID.Text = RsMst_StorageRateCard!SRCID
TxtSRCDate.Value = RsMst_StorageRateCard!SRCDate
TxtLocationID.Text = GetLocationName(RsMst_StorageRateCard!LocationID)
TxtExchangeTypeID.Text = GetExchangeName(RsMst_StorageRateCard!ExchangeTypeID)
TxtCommodityGroupID.Text = GetCommodityGroupName(RsMst_StorageRateCard!CommodityGroupID)

CmbLocationID.Text = TxtLocationID.Text
CmbExchangeTypeID.Text = TxtExchangeTypeID.Text
CmbCommodityGroupID.Text = TxtCommodityGroupID.Text

ChkDefault.Enabled = True
Call GetSRCGridData
'TxtFromDate.Value = RsMst_StorageRateCard!FromDate
'TxtToDate.Value = RsMst_StorageRateCard!ToDate
'TxtTillDate.Value = RsMst_StorageRateCard!TillDate

If IsNull(RsMst_StorageRateCard!BillingCycleID) = False Then
   TxtBillingCycle.Text = GetBillingCycle(RsMst_StorageRateCard!BillingCycleID, "")
   TxtTotDays = GetBillingCycle(RsMst_StorageRateCard!BillingCycleID, "D")
Else
   TxtBillingCycle.Text = "" 'GetBillingCycle(RsMst_StorageRateCard!BillingCycleID, "")
   TxtTotDays = "" 'GetBillingCycle(RsMst_StorageRateCard!BillingCycleID, "D")
End If

If IsNull(RsMst_StorageRateCard!DepoBillingCycleID) = False Then
   TxtDepositBillingCycle.Text = GetBillingCycle(RsMst_StorageRateCard!DepoBillingCycleID, "")
   TxtDepositTotDays = GetBillingCycle(RsMst_StorageRateCard!DepoBillingCycleID, "D")
Else
   TxtDepositBillingCycle.Text = "" 'GetBillingCycle(RsMst_StorageRateCard!BillingCycleID, "")
   TxtDepositTotDays = "" 'GetBillingCycle(RsMst_StorageRateCard!BillingCycleID, "D")
End If


ChkDefault.Value = IIf(IsNull(RsMst_StorageRateCard!DefaultRate), 0, RsMst_StorageRateCard!DefaultRate)

mFlag = RsMst_StorageRateCard!Flag
If LCase(mFlag) = "b" Then
   TxtBillingUnit = "Bags"
ElseIf LCase(mFlag) = "q" Then
   TxtBillingUnit = "Quantity"
ElseIf LCase(mFlag) = "s" Then
   TxtBillingUnit = "Sq Ft"
Else
   TxtBillingUnit = ""
End If

TxtRateOnQty.Text = IIf(IsNull(RsMst_StorageRateCard!RateOnQty), 0, RsMst_StorageRateCard!RateOnQty)
TxtPreFix = IIf(IsNull(RsMst_StorageRateCard!SM_Prefix), "", RsMst_StorageRateCard!SM_Prefix)

TxtCreated = IIf(IsNull(RsMst_StorageRateCard!CreatedBy), "", RsMst_StorageRateCard!CreatedBy) & " :- " & IIf(IsNull(RsMst_StorageRateCard!CreatedDate), "", RsMst_StorageRateCard!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_StorageRateCard!UpdatedBy), "", RsMst_StorageRateCard!UpdatedBy) & " :- " & IIf(IsNull(RsMst_StorageRateCard!UpdatedDate), "", RsMst_StorageRateCard!UpdatedDate)

If LCase(TxtBillingUnit) = "bags" Then
   Call InDataBagSize
   Call IndataDepoBagSize
End If

Call GButtonLock(Me, True)

End Sub
Private Sub InDataBagSize()
Call TableMst_StorageCardRateDetail(" where SRCID=" & Val(TxtSRCID.Text) & "")
MSHFlexGrid2.Rows = RsMst_StorageCardRateDetail.RecordCount + 2
For i = 1 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(i, 0) = RsMst_StorageCardRateDetail!FromBagSize
    MSHFlexGrid2.TextMatrix(i, 1) = RsMst_StorageCardRateDetail!ToBagSize
    MSHFlexGrid2.TextMatrix(i, 2) = Format(RsMst_StorageCardRateDetail!RatePerBag, "#####0.0000")
    MSHFlexGrid2.TextMatrix(i, 3) = RsMst_StorageCardRateDetail!SRCDID
    RsMst_StorageCardRateDetail.MoveNext
Next
End Sub
Private Sub IndataDepoBagSize()
Call TableMst_StorageCardDepoRateDetail(" where SRCID=" & Val(TxtSRCID.Text) & "")
MSHFlexGrid4.Rows = RsMst_StorageCardDepoRateDetail.RecordCount + 2
For i = 1 To MSHFlexGrid4.Rows - 2
    MSHFlexGrid4.TextMatrix(i, 0) = RsMst_StorageCardDepoRateDetail!FromBagSize
    MSHFlexGrid4.TextMatrix(i, 1) = RsMst_StorageCardDepoRateDetail!ToBagSize
    MSHFlexGrid4.TextMatrix(i, 2) = Format(RsMst_StorageCardDepoRateDetail!RatePerBag, "#####0.0000")
    MSHFlexGrid4.TextMatrix(i, 3) = RsMst_StorageCardDepoRateDetail!DepoSRCDID
    RsMst_StorageCardDepoRateDetail.MoveNext
Next
End Sub


Private Sub SaveCmd_Click()

'If TxtSRCDate.Value < Date Then
'   MsgBox "SRC Date Should Not be Less than Current Date !!!"
'   TxtSRCDate.SetFocus
'   Exit Sub
'End If

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Name not allowed !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Blank Exchange Type not allowed !!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbCommodityGroupID.Text = "" Then
   MsgBox "Blank Commodity Name not allowed !!!"
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If

'Commneted on 15th April at 11:59 am . Backup in Store-Man Backup folder RAR file 'Store-Man with SRC and SContract'

'If TxtToDate.Value < TxtFromDate.Value Then
'   MsgBox "To Date Should be Greater than From Date !!!"
'   TxtToDate.SetFocus
'   Exit Sub
'End If

'If TxtTillDate.Value < TxtToDate.Value Then
'   MsgBox "Till Date Should be Greater than From Date And To Date !!!"
'   TxtTillDate.SetFocus
'   Exit Sub
'End If


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

If CmbDepositBillingCycle.Text = "" Then
   MsgBox "Blank Deposit Billing Cycle not allowed!!!"
   CmbDepositBillingCycle.SetFocus
   Exit Sub
End If


If LCase(CmbBillingUnit.Text) = "bags" Then
   If MSHFlexGrid2.Rows < 3 Then
      MsgBox "Bag Size details missing in 'Rate Details As per Bags'."
      TxtFromBagSize.SetFocus
      Exit Sub
   End If
   
   If MSHFlexGrid4.Rows < 3 Then
      MsgBox "Bag Size details missing in 'Deposit Rate Details As per Bags'"
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

If Edit_Flg = "A" Then
   If RsMst_StorageRateCard.RecordCount > 0 Then
      RsMst_StorageRateCard.MoveFirst
   End If
   'RsMst_StorageRateCard.Filter = "LocationID = " & mLocationID & " And ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & " And FromDate = '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And ToDate = '" & Format(TxtToDate.Value, Gen_DatFormat) & "' And Flag = '" & mFlag & "' And BillingCycleID = " & mBillingCycleID & " And DefaultRate = " & ChkDefault.Value & " "
   RsMst_StorageRateCard.Filter = "LocationID = " & mLocationID & " And ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & "  And BillingCycleID = " & mBillingCycleID & " And Flag = '" & mFlag & "' "
   If RsMst_StorageRateCard.RecordCount > 0 Then
      MsgBox "Dulicate Storage Rate Card Not allowed !!!"
      RsMst_StorageRateCard.Filter = ""
      Exit Sub
   End If
   RsMst_StorageRateCard.Filter = ""
   RsMst_StorageRateCard.AddNew
   RsMst_StorageRateCard!SRCID = GetSRCID
   RsMst_StorageRateCard!G_UID = GetGUID
   TxtSRCID.Text = RsMst_StorageRateCard!SRCID
Else
   RsMst_StorageRateCard.MoveFirst
   'RsMst_StorageRateCard.Filter = "LocationID = " & mLocationID & " And  ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & " And FromDate = '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And ToDate = '" & Format(TxtToDate.Value, Gen_DatFormat) & "'  And Flag = '" & mFlag & "' And BillingCycleID = " & mBillingCycleID & " And DefaultRate = " & ChkDefault.Value & " "
   RsMst_StorageRateCard.Filter = "LocationID = " & mLocationID & " And ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & "  And BillingCycleID = " & mBillingCycleID & " And Flag = '" & mFlag & "' " 'And DefaultRate = " & ChkDefault.Value & " "
   If RsMst_StorageRateCard.RecordCount > 0 Then
      If RsMst_StorageRateCard!SRCID <> TxtSRCID Then
         MsgBox "Dulicate Storage Rate Card Not allowed !!!"
         RsMst_StorageRateCard.Filter = ""
         Exit Sub
      End If
   End If
   RsMst_StorageRateCard.Filter = ""
   RsMst_StorageRateCard.MoveFirst
   RsMst_StorageRateCard.Find " SRCID= " & TxtSRCID.Text
End If
RsMst_StorageRateCard!SRCID = TxtSRCID.Text
RsMst_StorageRateCard!SRCDate = TxtSRCDate.Value
RsMst_StorageRateCard!LocationID = mLocationID
RsMst_StorageRateCard!CommodityGroupID = mCommodityGroupId
RsMst_StorageRateCard!ExchangeTypeID = mExchangeTypeID
RsMst_StorageRateCard!SM_Prefix = TxtPreFix
RsMst_StorageRateCard!Flag = mFlag

'Commneted on 15th April at 11:59 am . Backup in Store-Man Backup folder RAR file 'Store-Man with SRC and SContract'
'RsMst_StorageRateCard!FromDate = TxtFromDate.Value
'RsMst_StorageRateCard!ToDate = TxtToDate.Value
'RsMst_StorageRateCard!TillDate = TxtTillDate.Value


RsMst_StorageRateCard!BillingCycleID = mBillingCycleID

RsMst_StorageRateCard!DepoBillingCycleID = mDepositBillingCycleID

RsMst_StorageRateCard!RateOnQty = Val(TxtRateOnQty.Text)
RsMst_StorageRateCard!DefaultRate = ChkDefault.Value

If IsNull(RsMst_StorageRateCard!CreatedBy) = True Or RsMst_StorageRateCard!CreatedBy = "" Then
   RsMst_StorageRateCard!CreatedBy = Gen_UserLoginID
   RsMst_StorageRateCard!CreatedDate = Now
Else
   RsMst_StorageRateCard!UpdatedBy = Gen_UserLoginID
   RsMst_StorageRateCard!UpdatedDate = Now
End If

RsMst_StorageRateCard.Update

If CmbBillingUnit.Text = "Bags" Then
   Call SaveStorageRateCardDetails
   Call SaveStorageRateDepoCardDetails
End If

Call GButtonLock(Me, True)

RsMst_StorageRateCard.MoveFirst
RsMst_StorageRateCard.Find "SRCID = " & TxtSRCID & ""

Call InData

End Sub
Private Sub SaveStorageRateCardDetails()
Call TableMst_StorageCardRateDetail(" Where SRCID=" & Val(TxtSRCID.Text) & "")

For i = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
       If MSHFlexGrid2.TextMatrix(i, 3) <> "" Then
          If RsMst_StorageCardRateDetail.RecordCount > 0 Then
             RsMst_StorageCardRateDetail.MoveFirst
          End If
          RsMst_StorageCardRateDetail.Find "SRCDID = " & MSHFlexGrid2.TextMatrix(i, 3) & ""
       Else
          RsMst_StorageCardRateDetail.AddNew
          RsMst_StorageCardRateDetail!SRCDID = GetMaxSRCDID
       End If
       RsMst_StorageCardRateDetail!SRCID = Val(TxtSRCID.Text)
       RsMst_StorageCardRateDetail!FromBagSize = Val(MSHFlexGrid2.TextMatrix(i, 0))
       RsMst_StorageCardRateDetail!ToBagSize = Val(MSHFlexGrid2.TextMatrix(i, 1))
       RsMst_StorageCardRateDetail!RatePerBag = Val(MSHFlexGrid2.TextMatrix(i, 2))
       RsMst_StorageCardRateDetail.Update
       RsMst_StorageCardRateDetail.Requery
    End If
Next
End Sub
Private Sub SaveStorageRateDepoCardDetails()
Call TableMst_StorageCardDepoRateDetail(" Where SRCID=" & Val(TxtSRCID.Text) & "")

For i = 1 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(i, 1) <> "" Then
       If MSHFlexGrid4.TextMatrix(i, 3) <> "" Then
          If RsMst_StorageCardDepoRateDetail.RecordCount > 0 Then
             RsMst_StorageCardDepoRateDetail.MoveFirst
          End If
          RsMst_StorageCardDepoRateDetail.Find "DepoSRCDID = " & MSHFlexGrid4.TextMatrix(i, 3) & ""
       Else
          RsMst_StorageCardDepoRateDetail.AddNew
          RsMst_StorageCardDepoRateDetail!DepoSRCDID = GetMaxDepoSRCDID
       End If
       RsMst_StorageCardDepoRateDetail!SRCID = Val(TxtSRCID.Text)
       RsMst_StorageCardDepoRateDetail!FromBagSize = Val(MSHFlexGrid4.TextMatrix(i, 0))
       RsMst_StorageCardDepoRateDetail!ToBagSize = Val(MSHFlexGrid4.TextMatrix(i, 1))
       RsMst_StorageCardDepoRateDetail!RatePerBag = Val(MSHFlexGrid4.TextMatrix(i, 2))
       RsMst_StorageCardDepoRateDetail.Update
       RsMst_StorageCardDepoRateDetail.Requery
    End If
Next
End Sub
Sub DeleteChildTableDetail()

tmp = "Delete from Mst_StorageRateCardDetail Where SRCID = " & Val(TxtSRCID.Text) & ""
Call TmpTable

tmp = "Delete from Mst_StorageCardDepoRateDetail Where SRCID = " & Val(TxtSRCID.Text) & ""
Call TmpTable

End Sub

Private Function GetSRCID()
Dim Retval As Double
tmp = "Select max(SRCID) from Mst_StorageRateCard "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetSRCID = Retval
End Function
Private Function GetMaxSRCDID()
Dim Retval As Double
tmp = "Select max(SRCDID) from Mst_StorageRateCardDetail "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxSRCDID = Retval
End Function
Private Function GetMaxDepoSRCDID()
Dim Retval As Double
tmp = "Select max(DepoSRCDID) from Mst_StorageRateCardDepoDetail "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxDepoSRCDID = Retval
End Function
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
If IsNumeric(TxtRateOnQty) = False Or Val(TxtRateOnQty) <= 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtRateOnQty.SetFocus
   Exit Sub
End If

TxtRateOnQty = Round(TxtRateOnQty, 4)

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
Private Sub GetSRCGridData()

Call Grid_Head2
If CmbLocationID.Text = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub
If CmbCommodityGroupID.Text = "" Then Exit Sub



tmp = "Select MSRC.SRCID,MSRC.SRCDate,ML.LocationName,ME.ExchangeType,MCG.CommodityGroup,"
tmp = tmp & " MB.BillingCycle ,MSRC.Flag, MSRC.RateonQty,MB1.BillingCycle ,MSRC.DefaultRate,MSRC.G_UID"
tmp = tmp & " From Mst_StorageRateCard MSRC"
tmp = tmp & " Inner Join Mst_Location ML On MSRC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME ON MSRC.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MSRC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Left Join Mst_BillingCycle MB On MSRC.BillingCycleID = MB.BillingCycleID "
tmp = tmp & " Left Join Mst_BillingCycle MB1 On MSRC.DepoBillingCycleID = MB1.BillingCycleID "

tmp = tmp & "Where ML.LocationName = '" & CmbLocationID.Text & "' And ME.ExchangeType = '" & CmbExchangeTypeID.Text & "' And MCG.CommodityGroup = '" & CmbCommodityGroupID.Text & "'"


Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
   For i = 1 To TmpRs.RecordCount
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid3.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid3.TextMatrix(i, 1) = Format(MSHFlexGrid3.TextMatrix(i, 1), "dd-mmm-yyyy")
       If LCase(MSHFlexGrid3.TextMatrix(i, 6)) = "b" Then
          MSHFlexGrid3.TextMatrix(i, 6) = "Bags"
       ElseIf LCase(MSHFlexGrid3.TextMatrix(i, 6)) = "q" Then
          MSHFlexGrid3.TextMatrix(i, 6) = "Quantity"
       ElseIf LCase(MSHFlexGrid3.TextMatrix(i, 6)) = "s" Then
          MSHFlexGrid3.TextMatrix(i, 6) = "Sq ft"
       Else
          MSHFlexGrid3.TextMatrix(i, 6) = "" '"Sq ft"
       End If
       If MSHFlexGrid3.TextMatrix(i, 9) = 0 Then
          MSHFlexGrid3.TextMatrix(i, 9) = "No"
       Else
          MSHFlexGrid3.TextMatrix(i, 9) = "Yes"
          If TxtSRCID.Text <> MSHFlexGrid3.TextMatrix(i, 0) Then
             ChkDefault.Value = 0
             ChkDefault.Enabled = False
          End If
       End If
       TmpRs.MoveNext
   Next
Else
   Call Grid_Head2
   ChkDefault.Enabled = True
End If


End Sub

Private Sub Grid_Head2()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.Cols = 11


MSHFlexGrid3.TextMatrix(0, 0) = "SRC ID"
MSHFlexGrid3.TextMatrix(0, 1) = "SRC Date"
MSHFlexGrid3.TextMatrix(0, 2) = "Location "
MSHFlexGrid3.TextMatrix(0, 3) = "Exchange Type"
MSHFlexGrid3.TextMatrix(0, 4) = "Commodity Group"
MSHFlexGrid3.TextMatrix(0, 5) = "Billing Cycle"
MSHFlexGrid3.TextMatrix(0, 6) = "Billing Unit"
MSHFlexGrid3.TextMatrix(0, 7) = "Rate on Units"
MSHFlexGrid3.TextMatrix(0, 8) = "Deposit Billing Cycle"
MSHFlexGrid3.TextMatrix(0, 9) = "Default"
MSHFlexGrid3.TextMatrix(0, 10) = "G_UID"

MSHFlexGrid3.ColWidth(0) = 500
MSHFlexGrid3.ColWidth(1) = 0
MSHFlexGrid3.ColWidth(2) = 1000
MSHFlexGrid3.ColWidth(3) = 1000
MSHFlexGrid3.ColWidth(4) = 800
MSHFlexGrid3.ColWidth(5) = 1000
MSHFlexGrid3.ColWidth(6) = 1000
MSHFlexGrid3.ColWidth(7) = 900
MSHFlexGrid3.ColWidth(8) = 1000
MSHFlexGrid3.ColWidth(9) = 600
MSHFlexGrid3.ColWidth(10) = 0

MSHFlexGrid3.RowHeight(0) = 600


End Sub
