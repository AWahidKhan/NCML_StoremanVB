VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Commodity 
   Caption         =   "Commodity Variety Master"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   9120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1365
      Left            =   240
      TabIndex        =   16
      Top             =   10080
      Visible         =   0   'False
      Width           =   15180
      Begin VB.ComboBox CmbSYearlyBilling 
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
         ItemData        =   "Frm_Mst_Commodity.frx":0000
         Left            =   6240
         List            =   "Frm_Mst_Commodity.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   92
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox CmbSRiskType 
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
         ItemData        =   "Frm_Mst_Commodity.frx":0017
         Left            =   2760
         List            =   "Frm_Mst_Commodity.frx":0024
         Sorted          =   -1  'True
         TabIndex        =   90
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox CmbSCommtrackFlag 
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
         ItemData        =   "Frm_Mst_Commodity.frx":003B
         Left            =   12420
         List            =   "Frm_Mst_Commodity.frx":0045
         Sorted          =   -1  'True
         TabIndex        =   85
         Top             =   165
         Width           =   975
      End
      Begin VB.ComboBox CmbSCommodityGroup 
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
         Left            =   6915
         Sorted          =   -1  'True
         TabIndex        =   62
         Top             =   180
         Width           =   3945
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   13080
         TabIndex        =   58
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox TxtSearchCommodityVariety 
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
         Height          =   345
         Left            =   2745
         TabIndex        =   57
         Top             =   150
         Width           =   2415
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   14085
         TabIndex        =   32
         Top             =   600
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   26
         Top             =   1080
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label21 
         Caption         =   "Yearly Billing Flag"
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
         Left            =   4440
         TabIndex        =   93
         Top             =   660
         Width           =   1740
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Risk Type"
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
         Left            =   1560
         TabIndex        =   89
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label Label24 
         Caption         =   "Commtrack Flag"
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
         Left            =   10890
         TabIndex        =   86
         Top             =   225
         Width           =   1530
      End
      Begin VB.Label Label5 
         Caption         =   "Commodity Group"
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
         Left            =   5250
         TabIndex        =   63
         Top             =   195
         Width           =   1785
      End
      Begin VB.Label Label22 
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
         Left            =   9870
         TabIndex        =   61
         Top             =   645
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
         Left            =   7470
         TabIndex        =   60
         Top             =   645
         Width           =   2355
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Commodity Variety"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   59
         Top             =   180
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8205
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   8880
      Begin VB.TextBox TxtHSNNo 
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
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   12
         Top             =   3360
         Width           =   1680
      End
      Begin VB.ComboBox CmbType 
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
         ItemData        =   "Frm_Mst_Commodity.frx":0052
         Left            =   5550
         List            =   "Frm_Mst_Commodity.frx":005C
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   3360
         Width           =   1200
      End
      Begin VB.ComboBox CmbHRVarietyID 
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   3795
         Width           =   6855
      End
      Begin VB.TextBox TxtHRVarietyID 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   3795
         Width           =   6840
      End
      Begin VB.CheckBox ChkYearlyBilling 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3705
         TabIndex        =   10
         Top             =   3075
         Width           =   195
      End
      Begin VB.ComboBox CmbRiskType 
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
         ItemData        =   "Frm_Mst_Commodity.frx":006D
         Left            =   6840
         List            =   "Frm_Mst_Commodity.frx":007A
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   3030
         Width           =   1950
      End
      Begin VB.TextBox TxtRiskType 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   75
         TabIndex        =   88
         Top             =   3030
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.CheckBox ChkCommtrackFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   3045
         Width           =   255
      End
      Begin VB.TextBox TxtMaxRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6840
         MaxLength       =   75
         TabIndex        =   5
         Top             =   1650
         Width           =   1920
      End
      Begin VB.TextBox TxtMinRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         MaxLength       =   75
         TabIndex        =   4
         Top             =   1650
         Width           =   2040
      End
      Begin VB.TextBox TxtEffectiveDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6840
         MaxLength       =   75
         TabIndex        =   8
         Top             =   2610
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.TextBox TxtServiceTaxNAVCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1920
      End
      Begin VB.ComboBox CmbServiceTaxID 
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2130
         Width           =   2055
      End
      Begin VB.CommandButton CmdServiceTaxID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2130
         Visible         =   0   'False
         Width           =   375
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
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4215
         Width           =   6840
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
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   4620
         Width           =   6840
      End
      Begin VB.CommandButton CmdCommodityGroupID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   375
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
         Left            =   4560
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox TxtCommodityGroupID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   240
         Width           =   3720
      End
      Begin VB.TextBox TxtMinimum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         MaxLength       =   75
         TabIndex        =   2
         Top             =   1200
         Width           =   2040
      End
      Begin VB.TextBox TxtMaximum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6840
         MaxLength       =   75
         TabIndex        =   3
         Top             =   1200
         Width           =   1920
      End
      Begin VB.Frame Frame2 
         Height          =   1530
         Left            =   120
         TabIndex        =   41
         Top             =   5040
         Width           =   8670
         Begin VB.CommandButton CmdDeleteList 
            Caption         =   "Delete from List"
            Height          =   375
            Left            =   7080
            TabIndex        =   64
            Top             =   1050
            Width           =   1335
         End
         Begin VB.TextBox TxtSCRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3840
            TabIndex        =   21
            Top             =   1050
            Width           =   1920
         End
         Begin VB.TextBox TxtGunnyWt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            TabIndex        =   17
            Top             =   600
            Width           =   1680
         End
         Begin VB.TextBox TxtStdDeduction 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3840
            TabIndex        =   18
            Top             =   600
            Width           =   1920
         End
         Begin VB.TextBox TxtVariation 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   7080
            TabIndex        =   20
            Top             =   600
            Width           =   1320
         End
         Begin VB.TextBox TxtLotSize 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   5760
            TabIndex        =   19
            Top             =   600
            Width           =   1320
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
            ItemData        =   "Frm_Mst_Commodity.frx":0091
            Left            =   120
            List            =   "Frm_Mst_Commodity.frx":0093
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   600
            Width           =   2055
         End
         Begin VB.CommandButton CmdClientAdd2List 
            Caption         =   "Add to List"
            Height          =   375
            Left            =   5760
            TabIndex        =   22
            Top             =   1050
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Storage Charges Rate Per Day Per MT --->"
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
            Left            =   120
            TabIndex        =   49
            Top             =   1125
            Width           =   3720
         End
         Begin VB.Label LblGunnyWt 
            Caption         =   "Gunny Wt in KG"
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
            Left            =   2280
            TabIndex        =   46
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label LblStdDeduction 
            Caption         =   "Std. Deduction in %"
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
            Left            =   3840
            TabIndex        =   45
            Top             =   255
            Width           =   1800
         End
         Begin VB.Label LblLotSize 
            Caption         =   "Lot Size in MT"
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
            Left            =   5760
            TabIndex        =   44
            Top             =   270
            Width           =   1440
         End
         Begin VB.Label LblVariation 
            Caption         =   "Variation %"
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
            Left            =   7200
            TabIndex        =   43
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label Label7 
            Caption         =   "Mode"
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
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox TxtCommodityID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TxtCommodity 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         MaxLength       =   75
         TabIndex        =   1
         Top             =   735
         Width           =   6840
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1575
         Left            =   135
         TabIndex        =   40
         Top             =   6585
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   2778
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox TxtServiceTaxID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2130
         Width           =   2055
      End
      Begin VB.ComboBox CmbFumigation 
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
         ItemData        =   "Frm_Mst_Commodity.frx":0095
         Left            =   1920
         List            =   "Frm_Mst_Commodity.frx":009F
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2565
         Width           =   2055
      End
      Begin VB.TextBox TxtFumigation 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2565
         Width           =   2055
      End
      Begin VB.TextBox TxtType 
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
         Left            =   5550
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label Label27 
         Caption         =   "HSN No"
         Height          =   210
         Left            =   240
         TabIndex        =   98
         Top             =   3435
         Width           =   1800
      End
      Begin VB.Label Label26 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   97
         Top             =   3405
         Width           =   735
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "HR Variety"
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
         Left            =   240
         TabIndex        =   95
         Top             =   3885
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Yearly Billing"
         Height          =   195
         Left            =   2520
         TabIndex        =   91
         Top             =   3060
         Width           =   885
      End
      Begin VB.Label Label17 
         Caption         =   "Risk Type"
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
         Left            =   4680
         TabIndex        =   87
         Top             =   3083
         Width           =   1800
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Commtrack Flag"
         Height          =   195
         Left            =   240
         TabIndex        =   84
         Top             =   3075
         Width           =   1140
      End
      Begin VB.Label Label13 
         Caption         =   "Maximum Price in MT"
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
         Left            =   4680
         TabIndex        =   74
         Top             =   1725
         Width           =   2040
      End
      Begin VB.Label Label11 
         Caption         =   "Minimum price in MT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   73
         Top             =   1643
         Width           =   1320
      End
      Begin VB.Label Label12 
         Caption         =   "Fumigation Flag"
         Height          =   240
         Left            =   240
         TabIndex        =   71
         Top             =   2595
         Width           =   1320
      End
      Begin VB.Label LblEffeDays 
         Caption         =   "Effective Days"
         Height          =   300
         Left            =   4680
         TabIndex        =   70
         Top             =   2640
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label8 
         Caption         =   "GST Group NAV Code"
         Height          =   180
         Left            =   4680
         TabIndex        =   69
         Top             =   2190
         Width           =   1800
      End
      Begin VB.Label Label6 
         Caption         =   "GST Group"
         Height          =   240
         Left            =   240
         TabIndex        =   67
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label Label10 
         Caption         =   "Created  :"
         Height          =   240
         Left            =   240
         TabIndex        =   56
         Top             =   4275
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Commodity"
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
         Left            =   3600
         TabIndex        =   51
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Minimum bag size in Kg"
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
         Left            =   240
         TabIndex        =   48
         Top             =   1170
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Maximum bag size in Kg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4680
         TabIndex        =   47
         Top             =   1260
         Width           =   2040
      End
      Begin VB.Label LblCommodityID 
         Caption         =   "Commodity Variety ID"
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label LblCommodity 
         Caption         =   "Commodity Variety"
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
         Left            =   240
         TabIndex        =   36
         Top             =   765
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   37
      Top             =   8160
      Width           =   8880
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   39
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1740
         TabIndex        =   38
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   6885
         TabIndex        =   25
         Top             =   660
         Width           =   1875
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   6885
         TabIndex        =   23
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   5130
         TabIndex        =   27
         Top             =   660
         Width           =   1755
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   5130
         TabIndex        =   28
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   3375
         TabIndex        =   29
         Top             =   660
         Width           =   1755
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   3375
         TabIndex        =   30
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1740
         TabIndex        =   24
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   31
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   9120
      TabIndex        =   75
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox Check1 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton CmdLocationID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CmbLocationID 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   78
         Top             =   600
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton CmdRemoveFromList 
         Caption         =   "Remove From List"
         Height          =   400
         Left            =   4440
         TabIndex        =   77
         Top             =   580
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton CmdAdd2List 
         Caption         =   "Add To List"
         Height          =   400
         Left            =   4440
         TabIndex        =   76
         Top             =   180
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Location Mapping"
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
         TabIndex        =   83
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label Label14 
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
         Left            =   1477
         TabIndex        =   80
         Top             =   255
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
      Height          =   7815
      Left            =   9120
      TabIndex        =   81
      Top             =   1080
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   13785
      _Version        =   393216
      WordWrap        =   -1  'True
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
Attribute VB_Name = "FrmMst_Commodity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mExchangeTypeID, mCommodityGroupId, mServiceTaxID, mFumFlag, mLocationID, mHRCommodityID, mHRVarietyID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
   For I = 1 To MSHFlexGrid3.Rows - 2
      MSHFlexGrid3.row = I
      MSHFlexGrid3.Col = 0
      MSHFlexGrid3.CellFontName = "Wingdings"
      MSHFlexGrid3.CellFontSize = 12
      MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid3.Text = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid3.Rows - 2
      MSHFlexGrid3.row = I
      MSHFlexGrid3.Col = 0
      MSHFlexGrid3.CellFontName = "Wingdings"
      MSHFlexGrid3.CellFontSize = 12
      MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid3.Text = strUnChecked
   Next
End If
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

'----
Private Sub CmbCommodityGroupID_LostFocus()
If CmbCommodityGroupID.Text = "" Then
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
mHRCommodityID = IIf(IsNull(RsMst_CommodityGroup!HRCommodityID), 0, RsMst_CommodityGroup!HRCommodityID)

If mHRCommodityID = 0 Then
   MsgBox "No HR Commodity Mapperd to this Group!!!"
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If

End Sub


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

Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
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
End Sub
Function CheckDuplicate() As Boolean
Dim Retval As Boolean
For j = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(j, 0) = CmbExchangeTypeID.Text Then
       If CmdClientAdd2List.Caption = "Add to List" Then
          Retval = True
          Exit For
       Else
          If j <> MSHFlexGrid2.RowSel Then
             Retval = True
             Exit For
          End If
       End If
    End If
Next
CheckDuplicate = Retval
  
End Function

Private Sub CmbFumigation_LostFocus()
If CmbFumigation.Text = "" Then Exit Sub
If CmbFumigation.Text <> "Yes" And CmbFumigation.Text <> "No" Then
   MsgBox "Invalid selection "
   CmbFumigation.SetFocus
   Exit Sub
End If
If CmbFumigation.Text = "Yes" Then
   LblEffeDays.Visible = True
   TxtEffectiveDays.Visible = True
Else
   LblEffeDays.Visible = False
   TxtEffectiveDays.Visible = False
End If
End Sub


Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location("Where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub



Private Sub CmbRiskType_LostFocus()
If CmbRiskType.Text = "" Then Exit Sub

If LCase(CmbRiskType.Text) <> "low" And LCase(CmbRiskType.Text) <> "high" And LCase(CmbRiskType.Text) <> "medium" Then
   MsgBox "Invalid selection!!! "
   CmbRiskType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSCommodityGroup_GotFocus()

tmp = CmbSCommodityGroup.Text
Call TableMst_CommodityGroup
CmbSCommodityGroup.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_CommodityGroup.RecordCount
   CmbSCommodityGroup.AddItem RsMst_CommodityGroup!COmmodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbSCommodityGroup.Text = tmp

End Sub

Private Sub CmbSCommodityGroup_LostFocus()

If CmbSCommodityGroup.Text = "" Then
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbSCommodityGroup.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodityGroup.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSCommtrackFlag_LostFocus()
If CmbSCommtrackFlag.Text = "" Then Exit Sub

If LCase(CmbSCommtrackFlag.Text) <> "yes" And LCase(CmbSCommtrackFlag.Text) <> "no" Then
   MsgBox "Invalid selection!!! "
   CmbSCommtrackFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSRiskType_LostFocus()
If CmbSRiskType.Text = "" Then Exit Sub

If LCase(CmbSRiskType.Text) <> "low" And LCase(CmbSRiskType.Text) <> "high" And LCase(CmbSRiskType.Text) <> "medium" Then
   MsgBox "Invalid selection!!! "
   CmbSRiskType.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmbSYearlyBilling_LostFocus()

If CmbSYearlyBilling.Text = "" Then Exit Sub

If LCase(CmbSYearlyBilling.Text) <> "yes" And LCase(CmbSYearlyBilling.Text) <> "no" Then
   MsgBox "Invalid selection!!! "
   CmbSYearlyBilling.SetFocus
   Exit Sub
End If


End Sub

Private Sub CmbType_LostFocus()

If LCase(CmbType.Text) = "" Then Exit Sub

If LCase(CmbType.Text) <> "agri" And LCase(CmbType.Text) <> "nagri" Then
   MsgBox "Invalid Selection"
   CmbType.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdAdd2List_Click()
If CmbLocationID.Text = "" Then
    MsgBox "Blank Location not Allowed!!!"
    CmbLocationID.SetFocus
    Exit Sub
End If
For I = 1 To MSHFlexGrid3.Rows - 1
    If CmbLocationID.Text = MSHFlexGrid3.TextMatrix(I, 1) Then
        MsgBox "Location Already Exist!!!"
        Exit Sub
    End If
Next
r = MSHFlexGrid3.Rows - 1
MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1

MSHFlexGrid3.TextMatrix(r, 0) = mLocationID
MSHFlexGrid3.TextMatrix(r, 1) = CmbLocationID.Text

Call ClearLocationControl
End Sub

Private Sub ClearLocationControl()
CmdAdd2List.Enabled = True
CmdRemoveFromList.Enabled = False
CmbLocationID.Enabled = True
CmdLocationID.Enabled = True
CmbLocationID.Text = ""
End Sub

Private Sub CmdClientAdd2List_Click()
Dim I As Variant

'    If CmdClientAdd2List.Caption = "Add to List" Then
'       If MSHFlexGrid2.Rows > 4 Then
'         MsgBox "All Exchange Are Added"
'         Exit Sub
'       End If
'    End If
    If CmbExchangeTypeID.Text = "" Then
       MsgBox "Exchange Type Cant Be Left Blank!!!"
       CmbExchangeTypeID.SetFocus
       Exit Sub
    End If
    If TxtGunnyWt.Text = "" Then
       MsgBox "Gunny Weight Can't Be Left Blank!!!"
       TxtGunnyWt.SetFocus
       Exit Sub
    End If
    If TxtLotSize.Text = "" Then
       MsgBox "Lot Size Can't Be Left Blank!!!"
       TxtLotSize.SetFocus
       Exit Sub
    End If
    If TxtStdDeduction.Text = "" Then
       MsgBox "Standard Deduction Can't Be Left Blank!!!"
       TxtStdDeduction.SetFocus
       Exit Sub
    End If
    If TxtVariation.Text = "" Then
       MsgBox "Variation Can't Be Left Blank!!!"
       TxtVariation.SetFocus
       Exit Sub
    End If
'    If TxtSCRate.Text = "" Then
'       MsgBox "Storage Charges Can't Be Left Blank!!!"
'       TxtSCRate.SetFocus
'       Exit Sub
'    End If
    
    If CheckDuplicate = True Then
       MsgBox "Duplicate Exchange Type !!"
       CmbExchangeTypeID.SetFocus
       Exit Sub
    End If
    If CmdClientAdd2List.Caption = "Add to List" Then
       I = MSHFlexGrid2.Rows - 1
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    Else
       I = MSHFlexGrid2.RowSel
    End If
    MSHFlexGrid2.TextMatrix(I, 0) = CmbExchangeTypeID.Text
    MSHFlexGrid2.TextMatrix(I, 1) = TxtGunnyWt.Text
    MSHFlexGrid2.TextMatrix(I, 2) = TxtStdDeduction.Text
    MSHFlexGrid2.TextMatrix(I, 3) = TxtLotSize.Text
    MSHFlexGrid2.TextMatrix(I, 4) = TxtVariation.Text
    MSHFlexGrid2.TextMatrix(I, 5) = TxtSCRate.Text
    
Call ClearControl
End Sub
Sub ClearControl()
CmbExchangeTypeID = ""
CmbExchangeTypeID.Locked = False
TxtGunnyWt = ""
TxtStdDeduction = ""
TxtLotSize = ""
TxtVariation = ""
TxtSCRate = ""
CmdClientAdd2List.Caption = "Add to List"
CmdDeleteList.Enabled = False
    
End Sub

Private Sub CmdCommodityGroupID_Click()
FrmMst_CommodityGroup.Show
End Sub

Private Sub CmdDeleteList_Click()
If CmdClientAdd2List.Caption <> "Add to List" Then
   I = MSHFlexGrid2.RowSel
End If
    
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
ans = MsgBox("Are you sure, you want to delete this Exchange Type !!!", vbYesNo)

If ans = vbYes Then
   If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   End If
   Call ClearControl
End If
End Sub

Private Sub CmdGo_Click()

Call Grid_Head

'MCV.FumigationFlag,MCV.EffectiveDays
tmp = "Select MCV.CommodityID, MCV.Commodity, MCV.Maximum, MCV.Minimum, MCG.CommodityGroup, MST.Description, MST.NavCode, MCV.G_UID, MCV.MaxRate, MCV.MinRate, MCV.CommTrackFlag, MCV.RiskType,MCV.YearlyBillingFlg,Isnull(MCV.HRVarietyID,0) 'MCV.HRVarietyID',MCV.HSNNo,MCV.AccountType,MCV.FumigationFlag,MCV.EffectiveDays " & _
      " From Mst_Commodity MCV " & _
      " Left Join Mst_CommodityGroup MCG On MCV.CommodityGroupID = MCG.CommodityGroupID " & _
      " Left Join Mst_ServiceTax MST On MCV.ServiceTaxID = MST.ServiceTaxID "
Dim wc As Variant
wc = ""
If TxtSearchCommodityVariety <> "" Then
   wc = " Where MCV.Commodity Like '%" & TxtSearchCommodityVariety & "%' "
End If

If CmbSCommodityGroup.Text <> "" Then
   If TxtSearchCommodityVariety = "" Then
      wc = " Where MCG.CommodityGroup = '" & CmbSCommodityGroup.Text & "'"
   Else
      wc = wc & " And MCG.CommodityGroup = '" & CmbSCommodityGroup.Text & "'"
   End If
End If

If LCase(CmbSCommtrackFlag.Text) = "yes" Then
   If wc <> "" Then
      wc = wc & " AND MCV.CommTrackFlag = 1 "
   Else
      wc = " Where MCV.CommTrackFlag= 1 "
   End If
ElseIf LCase(CmbSCommtrackFlag.Text) = "no" Then
   If wc <> "" Then
      wc = wc & " AND MCV.CommTrackFlag = 0 "
   Else
      wc = " Where MCV.CommTrackFlag = 0 "
   End If
End If

If CmbSRiskType.Text <> "" Then
   If wc <> "" Then
      wc = wc & " AND MCV.RiskType = '" & CmbSRiskType.Text & "'"
   Else
      wc = " Where MCV.RiskType = '" & CmbSRiskType.Text & "'"
   End If
End If

If LCase(CmbSYearlyBilling.Text) = "yes" Then
   If wc <> "" Then
      wc = wc & " AND MCV.YearlyBillingFlg = 1 "
   Else
      wc = " Where MCV.YearlyBillingFlg= 1 "
   End If
ElseIf LCase(CmbSYearlyBilling.Text) = "no" Then
   If wc <> "" Then
      wc = wc & " AND MCV.YearlyBillingFlg = 0 "
   Else
      wc = " Where MCV.YearlyBillingFlg = 0 "
   End If
End If


tmp = tmp & wc & " Order by MCV.Commodity"

Call TmpTable

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 13) = GetHRVariety(MSHFlexGrid1.TextMatrix(iu, 13))
    If LCase(MSHFlexGrid1.TextMatrix(iu, 16)) = "y" Then
       MSHFlexGrid1.TextMatrix(iu, 16) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 16) = "No"
    End If
    TmpRs.MoveNext
Next

Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdRemoveFromList_Click()
I = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
   For RO = I To MSHFlexGrid3.Rows - 2
       For C = 0 To MSHFlexGrid3.Cols - 1
            MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
       Next
   Next
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
End If
Call ClearLocationControl
End Sub

Private Sub CmdServiceTaxID_Click()
FrmMst_ServiceTax.Show
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
TxtSearchCommodityVariety = ""
CmbSCommodityGroup.Text = ""
Label22.Caption = ""

Call TableMst_Commodity(" order By Commodity")
Call TableMst_CommodityDetail
Call Grid_Head2
If RsMst_Commodity.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call Indata
End Sub
Private Sub Grid_Head2()
MSHFlexGrid2.Clear
MSHFlexGrid2.ClearStructure
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 6
MSHFlexGrid2.TextMatrix(0, 0) = "Exchange"
MSHFlexGrid2.TextMatrix(0, 1) = "Gunny Weight"
MSHFlexGrid2.TextMatrix(0, 2) = "Standard Deduction"
MSHFlexGrid2.TextMatrix(0, 3) = "Lot Size"
MSHFlexGrid2.TextMatrix(0, 4) = "Variation"
MSHFlexGrid2.TextMatrix(0, 5) = "Storage Charges Rate Per Day Per MT"

MSHFlexGrid2.ColWidth(0) = 1400
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 1600
MSHFlexGrid2.ColWidth(3) = 1600
MSHFlexGrid2.ColWidth(4) = 1000
MSHFlexGrid2.ColWidth(5) = 1400

MSHFlexGrid2.RowHeight(0) = 600

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Commodity.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Commodity.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 2
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



Private Sub MSHFlexGrid2_Click()
  CmdClientAdd2List.Caption = "Update List"
  I = MSHFlexGrid2.RowSel
  If I <= 0 Then Exit Sub
  CmbExchangeTypeID = MSHFlexGrid2.TextMatrix(I, 0)
  CmbExchangeTypeID.Locked = True
  TxtGunnyWt = MSHFlexGrid2.TextMatrix(I, 1)
  TxtStdDeduction = MSHFlexGrid2.TextMatrix(I, 2)
  TxtLotSize = MSHFlexGrid2.TextMatrix(I, 3)
  TxtVariation = MSHFlexGrid2.TextMatrix(I, 4)
  TxtSCRate.Text = MSHFlexGrid2.TextMatrix(I, 5)
  CmdDeleteList.Enabled = True
End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
'Call ClearLocationControl
'If MSHFlexGrid3.TextMatrix(MSHFlexGrid3.RowSel, 0) = "" Then Exit Sub
'CmbLocationID.Text = MSHFlexGrid3.TextMatrix(MSHFlexGrid3.RowSel, 1)
'CmbLocationID.Enabled = False
'CmdLocationID.Enabled = False
'CmdAdd2List.Enabled = False
'CmdRemoveFromList.Enabled = True
With MSHFlexGrid3
    If .Col = 0 Then  '.Row = 1 And
      If .TextMatrix(.row, .Col) = "" Then Exit Sub
      .Col = 0
         If .TextMatrix(.row, .Col) = strUnChecked Then
            .TextMatrix(.row, .Col) = strChecked
         Else
         .TextMatrix(.row, .Col) = strUnChecked
      End If
   End If
End With

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Commodity.MoveNext
If RsMst_Commodity.EOF Then
   RsMst_Commodity.MoveLast
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
RsMst_Commodity.MovePrevious
If RsMst_Commodity.BOF Then
   RsMst_Commodity.MoveFirst
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
RsMst_Commodity.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Commodity.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_Commodity.RecordCount = 0 Then Exit Sub
 
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 800
Frame0.Left = Frame1.Left ' - 100
Frame0.Top = Frame1.Top ' - 100
Frame0.Width = Me.Width - 250
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200
RsMst_Commodity.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'TxtSearchCommodityVariety = ""
'CmbSCommodityGroup.Text = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

'MSHFlexGrid1.Rows = RsMst_Commodity.RecordCount + 1
'
'For i = 1 To RsMst_Commodity.RecordCount
'    For c = 0 To MSHFlexGrid1.Cols - 1
'       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsMst_Commodity.Fields(c)), "", RsMst_Commodity.Fields(c))
'    Next
'    RsMst_Commodity.MoveNext
'Next
End Sub
Public Sub Indata()

Me.Refresh
Frame2.Enabled = False

Frame3.Enabled = False
Call ClearLocationControl
Call Grid_Head1

Call ClearControl

CmdCommodityGroupID.Visible = False
CmbCommodityGroupID.Visible = False
TxtCommodityGroupID.Visible = True
CmbCommodityGroupID.Text = ""

CmdServiceTaxID.Visible = False
CmbServiceTaxID.Visible = Fasle
TxtServiceTaxID.Visible = True
CmbServiceTaxID.Text = ""

TxtHRVarietyID.Visible = True
CmbHRVarietyID.Visible = False

CmbRiskType.Visible = False
TxtRiskType.Visible = True

TxtCommodityID.Locked = True
TxtCommodity.Locked = True
TxtEffectiveDays.Locked = True
TxtMaximum.Locked = True
TxtMinimum.Locked = True
TxtMaxRate.Locked = True
TxtMinRate.Locked = True

TxtGunnyWt.Locked = True
TxtStdDeduction.Locked = True
TxtLotSize.Locked = True
TxtVariation.Locked = True
CmbExchangeTypeID.Locked = True
CmbExchangeTypeID.Text = ""
TxtGunnyWt = ""
TxtStdDeduction = ""
TxtLotSize = ""
TxtVariation = ""
TxtMaximum = ""
TxtMinimum = ""
TxtMaxRate = ""
TxtMinRate = ""
TxtServiceTaxNAVCode = ""

CmbFumigation.Visible = False
TxtFumigation.Visible = True
TxtFumigation.Locked = True
TxtFumigation.Text = ""

MSHFlexGrid2.Enabled = False

ChkCommtrackFlag.Enabled = False
ChkYearlyBilling.Enabled = False

TxtCommodityID = IIf(IsNull(RsMst_Commodity!CommodityID), "", RsMst_Commodity!CommodityID)
TxtCommodity = IIf(IsNull(RsMst_Commodity!Commodity), "", RsMst_Commodity!Commodity)
TxtMaximum = IIf(IsNull(RsMst_Commodity!Maximum), "0", RsMst_Commodity!Maximum)
TxtMinimum = IIf(IsNull(RsMst_Commodity!Minimum), "0", RsMst_Commodity!Minimum)
TxtMaxRate = IIf(IsNull(RsMst_Commodity!MaxRate), "", RsMst_Commodity!MaxRate)
TxtMinRate = IIf(IsNull(RsMst_Commodity!MinRate), "", RsMst_Commodity!MinRate)
TxtCommodityGroupID = GetCommodityGroupName(IIf(IsNull(RsMst_Commodity!CommodityGroupID), 0, RsMst_Commodity!CommodityGroupID))
mCommodityGroupId = IIf(IsNull(RsMst_Commodity!CommodityGroupID), 0, RsMst_Commodity!CommodityGroupID)
mServiceTaxID = IIf(IsNull(RsMst_Commodity!ServiceTaxID), 0, RsMst_Commodity!ServiceTaxID)
mFumFlag = IIf(IsNull(RsMst_Commodity!FumigationFlag), "N", RsMst_Commodity!FumigationFlag)
ChkCommtrackFlag = IIf(IsNull(RsMst_Commodity!CommTrackFlag), "0", RsMst_Commodity!CommTrackFlag)
ChkYearlyBilling = IIf(IsNull(RsMst_Commodity!YearlyBillingFlg), "0", RsMst_Commodity!YearlyBillingFlg)
TxtRiskType = IIf(IsNull(RsMst_Commodity!RiskType), "", RsMst_Commodity!RiskType)
CmbCommodityGroupID.Text = TxtCommodityGroupID

TxtHSNNo.Locked = True
TxtHSNNo = IIf(IsNull(RsMst_Commodity!HSNNo), "", RsMst_Commodity!HSNNo)
CmbType.Visible = False
TxtType.Visible = True
TxtType = IIf(IsNull(RsMst_Commodity!AccountType), "", RsMst_Commodity!AccountType)
CmbType.Text = TxtType


If mFumFlag = "Y" Then
   TxtFumigation.Text = "Yes"
   TxtEffectiveDays.Visible = True
   LblEffeDays.Visible = True
   TxtEffectiveDays.Text = IIf(IsNull(RsMst_Commodity!EffectiveDays), "0", RsMst_Commodity!EffectiveDays)
Else
   TxtFumigation.Text = "No"
   TxtEffectiveDays.Visible = False
   LblEffeDays.Visible = False
   TxtEffectiveDays.Text = ""
End If
If mServiceTaxID <> 0 Then
   TxtServiceTaxID = GetServiceTaxDescription(mServiceTaxID)
   TxtServiceTaxNAVCode = GetServiceTaxNAVCode(mServiceTaxID)
Else
   TxtServiceTaxID = "" 'GetServiceTaxDescription(mServiceTaxID)
   TxtServiceTaxNAVCode = "" 'GetServiceTaxNAVCode(mServiceTaxID)
End If

mHRVarietyID = IIf(IsNull(RsMst_Commodity!HRVarietyID), 0, RsMst_Commodity!HRVarietyID)
TxtHRVarietyID = GetHRVariety(mHRVarietyID)

TxtCreated = IIf(IsNull(RsMst_Commodity!CreatedBy), "", RsMst_Commodity!CreatedBy) & " :- " & IIf(IsNull(RsMst_Commodity!CreatedDate), "", RsMst_Commodity!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Commodity!UpdatedBy), "", RsMst_Commodity!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Commodity!UpdatedDate), "", RsMst_Commodity!UpdatedDate)

Call TableMst_CommodityDetail("Where CommodityID = " & TxtCommodityID & " ")

Call Grid_Head2
Dim I As Variant
I = 1
For ix = 1 To RsMst_CommodityDetail.RecordCount
    
    MSHFlexGrid2.TextMatrix(I, 0) = GetExchangeName(RsMst_CommodityDetail!ExchangeTypeID)
    MSHFlexGrid2.TextMatrix(I, 1) = RsMst_CommodityDetail!GunnyWt
    MSHFlexGrid2.TextMatrix(I, 2) = RsMst_CommodityDetail!StdDeduction
    MSHFlexGrid2.TextMatrix(I, 3) = RsMst_CommodityDetail!LotSize
    MSHFlexGrid2.TextMatrix(I, 4) = RsMst_CommodityDetail!Variation
    MSHFlexGrid2.TextMatrix(I, 5) = RsMst_CommodityDetail!ScRate
    RsMst_CommodityDetail.MoveNext
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    I = I + 1
Next

Call IndataLocationMapping

If RsMst_Commodity.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub IndataLocationMapping()
Call GetAllLocation
'tmp = " Select MLCM.LocationCommodityMappingID, ML.LocationID , ML.LocationName"
'tmp = tmp & " From Mst_LocationCOmmodityMapping MLCM"
'tmp = tmp & " Inner Join Mst_Location ML on MLCM.LocationID = ML.LocationID"
'tmp = tmp & " Where MLCM.CommodityID = " & Val(TxtCommodityID)
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + TmpRs.RecordCount
'    For I = 1 To TmpRs.RecordCount
'        MSHFlexGrid3.TextMatrix(I, 0) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
'        MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
'        MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!LocationCommodityMappingID), "", TmpRs!LocationCommodityMappingID)
'        TmpRs.MoveNext
'    Next
'End If

End Sub


Private Sub TxtCommodity_LostFocus()
TxtCommodity = Trim(TxtCommodity)
End Sub

Private Sub TxtCommodityID_LostFocus()
If TxtCommodityID = "" Then Exit Sub
If IsNumeric(TxtCommodityID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtCommodityID.SetFocus
   Exit Sub
End If
TxtCommodityID = Val(TxtCommodityID)
End Sub

Private Sub TxtEffectiveDays_LostFocus()
If TxtEffectiveDays = "" Then Exit Sub
If Not IsNumeric(TxtEffectiveDays) Then
   MsgBox "Invalid numeric format !!!! "
   TxtEffectiveDays.SetFocus
End If

End Sub

Private Sub TxtGunnyWt_LostFocus()
If TxtGunnyWt = "" Then Exit Sub
If IsNumeric(TxtGunnyWt) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtGunnyWt.SetFocus
   Exit Sub
End If
TxtGunnyWt = Val(TxtGunnyWt)
End Sub

Private Sub TxtMaximum_LostFocus()
If TxtMaximum = "" Then Exit Sub
If Not IsNumeric(TxtMaximum) Then
   MsgBox "Invalid numeric format !!!! "
   TxtMaximum.SetFocus
End If

End Sub


Private Sub TxtMaxRate_LostFocus()
If TxtMaxRate = "" Then Exit Sub
If Not IsNumeric(TxtMaxRate) Then
   MsgBox "Invalid numeric format !!!! "
   TxtMaxRate.SetFocus
End If
End Sub

Private Sub TxtMinimum_LostFocus()
If TxtMinimum = "" Then Exit Sub
If Not IsNumeric(TxtMinimum) Then
   MsgBox "Invalid numeric format !!!! "
   TxtMinimum.SetFocus
End If
End Sub

Private Sub TxtMinRate_LostFocus()
If TxtMinRate = "" Then Exit Sub
If Not IsNumeric(TxtMinRate) Then
   MsgBox "Invalid numeric format !!!! "
   TxtMinRate.SetFocus
End If
End Sub

Private Sub TxtSCRate_LostFocus()
If TxtSCRate = "" Then Exit Sub
If IsNumeric(TxtSCRate) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtSCRate.SetFocus
   Exit Sub
End If
TxtSCRate = Val(TxtSCRate)
End Sub

Private Sub TxtStdDeduction_LostFocus()
If TxtStdDeduction = "" Then Exit Sub
If IsNumeric(TxtStdDeduction) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtStdDeduction.SetFocus
   Exit Sub
End If
TxtStdDeduction = Val(TxtStdDeduction)
End Sub
Private Sub TxtLotSize_LostFocus()
If TxtLotSize = "" Then Exit Sub
If IsNumeric(TxtLotSize) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtLotSize.SetFocus
   Exit Sub
End If
TxtLotSize = Val(TxtLotSize)
End Sub
Private Sub TxtVariation_LostFocus()
If TxtVariation = "" Then Exit Sub
If IsNumeric(TxtVariation) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtVariation.SetFocus
   Exit Sub
End If
TxtVariation = Val(TxtVariation)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame2.Enabled = True

CmdCommodityGroupID.Visible = True
CmbCommodityGroupID.Visible = True
TxtCommodityGroupID.Visible = False
CmbCommodityGroupID.Text = ""

CmdServiceTaxID.Visible = True
CmbServiceTaxID.Visible = True
TxtServiceTaxID.Visible = False
CmbServiceTaxID.Text = ""


TxtCommodityID.Locked = True
TxtCommodity.Locked = False
TxtGunnyWt.Locked = False
TxtStdDeduction.Locked = False
TxtLotSize.Locked = False
TxtVariation.Locked = False
CmbExchangeTypeID.Locked = False
MSHFlexGrid2.Enabled = True
CmbExchangeTypeID = ""
TxtMaximum.Locked = False
TxtEffectiveDays.Locked = False
TxtMinimum.Locked = False
TxtMaxRate.Locked = False
TxtMinRate.Locked = False

CmbFumigation.Visible = True
TxtFumigation.Visible = False
TxtFumigation.Text = ""
CmbFumigation.Text = ""

CmbRiskType.Visible = True
TxtRiskType.Visible = False
CmbRiskType.Text = ""

TxtHRVarietyID.Visible = False
CmbHRVarietyID.Visible = True
CmbHRVarietyID.Text = ""

TxtCommodityID = ""
TxtCommodity = ""
TxtGunnyWt = ""
TxtStdDeduction = ""
TxtLotSize = ""
TxtVariation = ""
TxtMinimum = ""
TxtMaximum = ""
TxtMinRate = ""
TxtMaxRate = ""

TxtCreated = ""
TxtUpdated = ""
TxtServiceTaxNAVCode = ""

TxtHSNNo.Locked = False
TxtHSNNo = ""
CmbType.Visible = True
TxtType.Visible = False
CmbType.Text = ""


ChkCommtrackFlag.Enabled = True
ChkCommtrackFlag.Value = 0

ChkYearlyBilling.Enabled = True
ChkYearlyBilling.Value = 0

Call ClearLocationControl
Call Grid_Head1
'''''''''''''''''''''''''''''
Call GetAllLocation
'''''''''''''''''''''''''''''
Frame3.Enabled = True


Call ClearControl
Call Grid_Head2
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
MSHFlexGrid2.TextMatrix(1, 0) = "Non NCDEX"
MSHFlexGrid2.TextMatrix(1, 1) = 0
MSHFlexGrid2.TextMatrix(1, 2) = 0
MSHFlexGrid2.TextMatrix(1, 3) = 0
MSHFlexGrid2.TextMatrix(1, 4) = 0
'TxtCommodity.SetFocus

End Sub

Private Sub GetAllLocation()
tmp = " Select ML.LocationID , ML.LocationName "
tmp = tmp & " From Mst_Location ML"
'tmp = tmp & " Left Join Mst_LocationCOmmodityMapping MLCM on MLCM.LocationID = ML.LocationID"
tmp = tmp & " Order by LocationName"
Call TmpTable

tmp = " Select ML.LocationID , ML.LocationName, MLCM.CommodityID, MLCM.LocationCommodityMappingID "
tmp = tmp & " From Mst_Location ML"
tmp = tmp & " Inner Join Mst_LocationCOmmodityMapping MLCM on MLCM.LocationID = ML.LocationID"
'tmp = tmp & " Where ML.LocationID not in (Select LocationID from Mst_LocationCOmmodityMapping MLCM Where CommodityID =  )"
tmp = tmp & " Where MLCM.CommodityID = " & Val(TxtCommodityID)
tmp = tmp & " Order by LocationName"
Call Tmp1Table


If TmpRs.RecordCount > 0 Then
    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + TmpRs.RecordCount
    For I = 1 To TmpRs.RecordCount
        MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
        MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
        TmpRs1.Filter = " LocationID = " & TmpRs!LocationID & " And CommodityID = " & Val(TxtCommodityID) & " "
        If TmpRs1.EOF = True Then
            MSHFlexGrid3.row = I
            MSHFlexGrid3.Col = 0
            MSHFlexGrid3.CellFontName = "Wingdings"
            MSHFlexGrid3.CellFontSize = 12
            MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
            MSHFlexGrid3.Text = strUnChecked
        Else
            MSHFlexGrid3.TextMatrix(I, 3) = TmpRs1!LocationCommodityMappingID
            MSHFlexGrid3.row = I
            MSHFlexGrid3.Col = 0
            MSHFlexGrid3.CellFontName = "Wingdings"
            MSHFlexGrid3.CellFontSize = 12
            MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
            MSHFlexGrid3.Text = strChecked
        End If
            TmpRs.MoveNext
    Next
End If
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
If LCase(Gen_UserRole) = "power" Then
   Check1.Enabled = True
   MSHFlexGrid3.Enabled = True
   Frame2.Enabled = True
   Frame3.Enabled = True
   TxtEffectiveDays.Locked = False
   TxtMaximum.Locked = False
   TxtMinimum.Locked = False
   TxtMaxRate.Locked = False
   TxtMinRate.Locked = False
   TxtCommodityID.Locked = True
   TxtCommodity.Locked = False
   TxtGunnyWt.Locked = False
   TxtStdDeduction.Locked = False
   TxtLotSize.Locked = False
   TxtVariation.Locked = False
   CmbExchangeTypeID.Locked = False
   MSHFlexGrid2.Enabled = True
   
   CmdCommodityGroupID.Visible = True
   CmbCommodityGroupID.Visible = True
   TxtCommodityGroupID.Visible = False
   Call CmbCommodityGroupID_GotFocus
   Call CmbCommodityGroupID_LostFocus
   
   CmbFumigation.Visible = True
   TxtFumigation.Visible = False
   
   CmbRiskType.Visible = True
   TxtRiskType.Visible = False
   
   CmdServiceTaxID.Visible = True
   CmbServiceTaxID.Visible = True
   TxtServiceTaxID.Visible = False
   TxtCommodity.SetFocus
   TxtHRVarietyID.Visible = False
   CmbHRVarietyID.Visible = True
   
   TxtHSNNo.Locked = False
   CmbType.Visible = True
   TxtType.Visible = False
   CmbType.Text = TxtType
   
ElseIf LCase(Gen_UserRole) = "headcmg" Then
   TxtMaxRate.Locked = False
   TxtMinRate.Locked = False
   Check1.Enabled = False
   MSHFlexGrid3.Enabled = False
   TxtMinRate.SetFocus
End If
CmbHRVarietyID.Text = TxtHRVarietyID
ChkCommtrackFlag.Enabled = True
ChkYearlyBilling.Enabled = True


If ChkCommtrackFlag.Value = 1 Then
   tmp = "Select * from Txn_CxTF_Details Where TxnType = 'D' And CommodityID = " & TxtCommodityID & ""
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      ChkCommtrackFlag.Enabled = False
   End If
End If

CmbCommodityGroupID.Text = TxtCommodityGroupID
CmbFumigation.Text = TxtFumigation.Text
CmbServiceTaxID.Text = TxtServiceTaxID
CmbRiskType.Text = TxtRiskType

Call ClearControl
  
End Sub
Function CheckDetele() As Boolean
Dim Retval As Boolean
Retval = True
tmp = "select * from Txn_CxTF_Details where CommodityID = " & TxtCommodityID
Call TmpTable

If TmpRs.RecordCount > 0 Then
   Retval = False
End If
TmpRs.Close
CheckDetele = Retval
End Function

Private Sub DeleteCmd_Click()


If CheckDetele = False Then
   MsgBox "You Can't Delete This Commodity !!!!! "
   Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   tmp = " Delete From Mst_LocationCommodityMapping Where CommodityID = " & Val(TxtCommodityID.Text)
   Call TmpTable
   
   tmp = "Delete from Mst_CommodityDetail where CommodityID = " & Val(TxtCommodityID)
   Call TmpTable
   tmp = "Select * from Mst_CommodityDetail where CommodityID = " & Val(TxtCommodityID)
   Call TmpTable
   TmpRs.Close
   
   RsMst_Commodity.Requery
   RsMst_Commodity.Find "CommodityID = " & TxtCommodityID
   If RsMst_Commodity.EOF Then Exit Sub
   
   RsMst_Commodity.Delete
   RsMst_Commodity.Update
   If RsMst_Commodity.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Commodity.MoveNext
   If RsMst_Commodity.EOF() Then
      RsMst_Commodity.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Commodity.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If CmbCommodityGroupID.Text = "" Then
   MsgBox "Blank Commodity Not Allowed !!! "
   TxtCommodity.SetFocus
   Exit Sub
End If

If TxtCommodity.Text = "" Then
   MsgBox "Blank Commodity Variety Not Allowed !!! "
   TxtCommodity.SetFocus
   Exit Sub
End If

If Val(TxtMaximum) = 0 Then
   MsgBox "Blank Maximum Bag size Not Allowed !!! "
   TxtMaximum.SetFocus
   Exit Sub
End If

If Val(TxtMinimum) = 0 Then
   MsgBox "Blank Minimum Bag size Not Allowed !!! "
   TxtMinimum.SetFocus
   Exit Sub
End If

If Val(TxtMinRate) = 0 Then
   MsgBox "Blank Minimum Rate Not Allowed !!! "
   TxtMinRate.SetFocus
   Exit Sub
End If

If Val(TxtMaxRate) = 0 Then
   MsgBox "Blank Maximum Rate Not Allowed !!! "
   TxtMaxRate.SetFocus
   Exit Sub
End If

If Val(TxtMinimum) > Val(TxtMaximum) Then
   MsgBox "Minimum Bag size can't be more than Maximum bag size !!! "
   TxtMinimum.SetFocus
   Exit Sub
End If

If Val(TxtMinRate) > Val(TxtMaxRate) Then
   MsgBox "Minimum Rate can't be more than Maximum Rate !!! "
   TxtMinRate.SetFocus
   Exit Sub
End If

If CmbFumigation.Text = "" Then
   MsgBox "Select Fumigation flag!!! "
   CmbFumigation.SetFocus
   Exit Sub
End If

If TxtEffectiveDays.Visible = True Then
   If TxtEffectiveDays.Text = "" Then
      MsgBox "Blank Effective days Not Allowed !!! "
      TxtEffectiveDays.SetFocus
      Exit Sub
   End If
End If

If CmbRiskType.Text = "" Then
   MsgBox "Blank Risk Type not allowed!!! "
   CmbRiskType.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows <= 2 Then
   MsgBox "Please Add Atleast One Exchange!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbHRVarietyID.Text = "" Then
   MsgBox "Blank HR Variety not Allowed"
   CmbHRVarietyID.SetFocus
   Exit Sub
End If


If CmbType.Text = "" Then
   MsgBox "Blank Type Not Allowed !!! "
   CmbType.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Commodity.RecordCount > 0 Then
      RsMst_Commodity.MoveFirst
      RsMst_Commodity.Find "Commodity= '" & TxtCommodity.Text & "'"
      If Not RsMst_Commodity.EOF Then
        MsgBox "Duplicate Commodity Name Not Allowed !!! "
         TxtCommodity.SetFocus
         Exit Sub
      End If
   End If
      RsMst_Commodity.AddNew
      RsMst_Commodity!CommodityID = GetCommodityID
      TxtCommodityID.Text = RsMst_Commodity!CommodityID
      RsMst_Commodity!G_UID = GetGUID
      RsMst_CommodityDetail.AddNew
      RsMst_CommodityDetail!G_UID = GetGUID
Else
    RsMst_Commodity.MoveFirst
    RsMst_Commodity.Find "Commodity= '" & TxtCommodity.Text & "'"
    If Not RsMst_Commodity.EOF Then
       If RsMst_Commodity!CommodityID <> TxtCommodityID.Text Then
          MsgBox "Duplicate Commodity Name Not Allowed !!! "
          TxtCommodity.SetFocus
          Exit Sub
       End If
    End If
    RsMst_Commodity.MoveFirst
    RsMst_Commodity.Find " CommodityID= " & TxtCommodityID.Text
End If

'RsMst_Commodity!CommodityID = TxtCommodityID.Text
RsMst_Commodity!Commodity = TxtCommodity.Text
RsMst_Commodity!Maximum = TxtMaximum.Text
RsMst_Commodity!Minimum = TxtMinimum.Text
RsMst_Commodity!MaxRate = TxtMaxRate.Text
RsMst_Commodity!MinRate = TxtMinRate.Text
RsMst_Commodity!CommodityGroupID = mCommodityGroupId
RsMst_Commodity!ServiceTaxID = mServiceTaxID
RsMst_Commodity!CommTrackFlag = ChkCommtrackFlag.Value
RsMst_Commodity!RiskType = CmbRiskType.Text
RsMst_Commodity!YearlyBillingFlg = ChkYearlyBilling.Value
If LCase(CmbFumigation.Text) = "yes" Then
   RsMst_Commodity!FumigationFlag = "Y"
   RsMst_Commodity!EffectiveDays = Val(TxtEffectiveDays.Text)
Else
   RsMst_Commodity!FumigationFlag = "N"
End If
RsMst_Commodity!HRVarietyID = mHRVarietyID

RsMst_Commodity!HSNNo = Trim(TxtHSNNo.Text)
RsMst_Commodity!AccountType = Trim(CmbType.Text)

If IsNull(RsMst_Commodity!CreatedBy) = True Or RsMst_Commodity!CreatedBy = "" Then
   RsMst_Commodity!CreatedBy = Gen_UserLoginID
   RsMst_Commodity!CreatedDate = Now
Else
   RsMst_Commodity!UpdatedBy = Gen_UserLoginID
   RsMst_Commodity!UpdatedDate = Now
End If

RsMst_Commodity.Update

tmp = "Select * from Mst_CommodityDetail Where CommodityId = " & TxtCommodityID
Call TmpTable

For I = 1 To TmpRs.RecordCount
    TmpRs.Delete
    TmpRs.Update
    TmpRs.MoveNext
Next

'tmp = "delete from Mst_CommodityDetail Where CommodityId = " & TxtCommodityID
'Call TmpTable

Call TableMst_CommodityDetail("Where CommodityId = " & TxtCommodityID)
If MSHFlexGrid2.Rows > 2 Then
   For I = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit For
       mExchangeTypeID = GetExchangeID(MSHFlexGrid2.TextMatrix(I, 0))
       If RsMst_CommodityDetail.RecordCount = 0 Then
         RsMst_CommodityDetail.AddNew
         RsMst_CommodityDetail!G_UID = GetGUID
       Else
         RsMst_CommodityDetail.MoveFirst
         RsMst_CommodityDetail.Find ("ExchangeTypeId = " & mExchangeTypeID)
         If RsMst_CommodityDetail.EOF = True Then
            RsMst_CommodityDetail.AddNew
            RsMst_CommodityDetail!G_UID = GetGUID
         Else
             
         End If
         x = RsMst_CommodityDetail.RecordCount
       End If
       
       RsMst_CommodityDetail!CommodityID = TxtCommodityID.Text
       RsMst_CommodityDetail!ExchangeTypeID = GetExchangeID(MSHFlexGrid2.TextMatrix(I, 0))
       RsMst_CommodityDetail!GunnyWt = Val(MSHFlexGrid2.TextMatrix(I, 1))
       RsMst_CommodityDetail!StdDeduction = Val(MSHFlexGrid2.TextMatrix(I, 2))
       RsMst_CommodityDetail!LotSize = Val(MSHFlexGrid2.TextMatrix(I, 3))
       RsMst_CommodityDetail!Variation = Val(MSHFlexGrid2.TextMatrix(I, 4))
       RsMst_CommodityDetail!ScRate = Val(MSHFlexGrid2.TextMatrix(I, 5))
       RsMst_CommodityDetail.Update
       RsMst_CommodityDetail.MoveNext
  Next I
End If
Call SaveLocationMapping
Call Indata
End Sub

Private Sub SaveLocationMapping()
tmp = " Delete From Mst_LocationCommodityMapping Where CommodityID = " & Val(TxtCommodityID.Text)
Call TmpTable

tmp = "Select * From Mst_LocationCommodityMapping Where CommodityID = " & Val(TxtCommodityID.Text)
Call Tmp1Table

For I = 1 To MSHFlexGrid3.Rows - 2
   If MSHFlexGrid3.TextMatrix(I, 0) = strChecked Then
      TmpRs1.AddNew
'      TmpRs1!LocationCommodityMappingID = GetCommodityMappingID
      TmpRs1!G_UID = GetGUID
      TmpRs1!LocationID = Val(MSHFlexGrid3.TextMatrix(I, 2))
      TmpRs1!CommodityID = Val(TxtCommodityID.Text)
      If IsNull(TmpRs1!CreatedBy) = True Or TmpRs1!CreatedBy = "" Then
         TmpRs1!CreatedBy = Gen_UserLoginID
         TmpRs1!CreatedDate = Now
      Else
         TmpRs1!UpdatedBy = Gen_UserLoginID
         TmpRs1!UpdatedDate = Now
      End If
      TmpRs1.Update
      TmpRs1.Requery
   End If
Next

'If MSHFlexGrid3.Rows = 2 Then Exit Sub
'
'tmp = "Select * From Mst_LocationCommodityMapping Where CommodityID = " & Val(TxtCommodityID.Text)
'Call Tmp1Table
'
'tmp1 = " (0, "
'For I = 1 To MSHFlexGrid3.Rows - 2
'   If MSHFlexGrid3.TextMatrix(I, 2) = "" Then
'      TmpRs1.AddNew
'      TmpRs1!LocationCommodityMappingID = GetCommodityMappingID
'      tmp1 = tmp1 & TmpRs1!LocationCommodityMappingID & " ,"
'      TmpRs1!G_UID = GetGUID
'   Else
'      TmpRs1.MoveFirst
'      TmpRs1.Find " LocationCommodityMappingID= " & Val(MSHFlexGrid3.TextMatrix(I, 2))
'      tmp1 = tmp1 & Val(MSHFlexGrid3.TextMatrix(I, 2)) & " ,"
'   End If
'   TmpRs1!LocationID = Val(MSHFlexGrid3.TextMatrix(I, 0))
'   TmpRs1!CommodityID = Val(TxtCommodityID.Text)
'   If IsNull(TmpRs1!CreatedBy) = True Or TmpRs1!CreatedBy = "" Then
'      TmpRs1!CreatedBy = Gen_UserLoginID
'      TmpRs1!CreatedDate = Now
'   Else
'      TmpRs1!UpdatedBy = Gen_UserLoginID
'      TmpRs1!UpdatedDate = Now
'   End If
'   TmpRs1.Update
'   TmpRs1.Requery
'Next
'
'tmp1 = Mid(tmp1, 1, Len(tmp1) - 1) & " )"
'tmp = " Delete From Mst_LocationCommodityMapping Where CommodityID = " & Val(TxtCommodityID.Text) & " And LocationCommodityMappingID not in " & tmp1 & ""
'Call TmpTable

End Sub

Private Function GetCommodityMappingID() As Double
Dim Retval As Double
tmp = "Select max(LocationCommodityMappingID) from Mst_LocationCOmmodityMapping"
Call TmpTable

Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCommodityMappingID = Retval
End Function

Private Sub Grid_Head1()
'With MSHFlexGrid3
'    .Clear
'    .Rows = 2
'    .Cols = 3
'
'    .TextMatrix(0, 0) = "Location ID"
'    .TextMatrix(0, 1) = "Location Name"
'    .TextMatrix(0, 2) = "Location Commodity Mapping ID"
'    .ColWidth(0) = 1200
'    .ColWidth(1) = 4500
'    .ColWidth(2) = 0
'End With
Check1.Value = vbUnchecked
With MSHFlexGrid3
    .Clear
    .Rows = 2
    .Cols = 4
    .FixedCols = 0

    .TextMatrix(0, 0) = "Select"
    .TextMatrix(0, 1) = "Location Name"
    .TextMatrix(0, 2) = "Location ID"
    .TextMatrix(0, 3) = "Location Commodity Mapping ID"
    
    .ColWidth(0) = 700
    .ColWidth(1) = 4500
    .ColWidth(2) = 1200
    .ColWidth(3) = 0
End With

End Sub


Private Function GetCommodityID() As Double
Dim Retval As Double
tmp = "Select max(CommodityID) from Mst_Commodity"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCommodityID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 18 '16 '14 '13 '12
MSHFlexGrid1.TextMatrix(0, 0) = "Commodity ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 2) = "Maximum bag size in Kg "
MSHFlexGrid1.TextMatrix(0, 3) = "Minimum bag size in Kg "
MSHFlexGrid1.TextMatrix(0, 4) = "Commodity Group"
MSHFlexGrid1.TextMatrix(0, 5) = "Service Tax Description"
MSHFlexGrid1.TextMatrix(0, 6) = "Service Tax NAV Code"
MSHFlexGrid1.TextMatrix(0, 7) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 8) = "Maximum Rate "
MSHFlexGrid1.TextMatrix(0, 9) = "Minimum Rate "
MSHFlexGrid1.TextMatrix(0, 10) = "Commtrack Flag"
MSHFlexGrid1.TextMatrix(0, 11) = "Risk Type"
MSHFlexGrid1.TextMatrix(0, 12) = "Yearly Billing Flag"
MSHFlexGrid1.TextMatrix(0, 13) = "HR Variety"
MSHFlexGrid1.TextMatrix(0, 14) = "HSN No"
MSHFlexGrid1.TextMatrix(0, 15) = "Type"
MSHFlexGrid1.TextMatrix(0, 16) = "Fumigation Flag"
MSHFlexGrid1.TextMatrix(0, 17) = "Effective Days"


'MSHFlexGrid1.TextMatrix(0, 4) = "Lot Size in MT"
'MSHFlexGrid1.TextMatrix(0, 5) = "Variation"


MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 3200
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 3000
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 0
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1400

'MSHFlexGrid1.ColWidth(4) = 1200

MSHFlexGrid1.RowHeight(0) = 600

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "CommodityID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_Commodity.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub CmbServiceTaxID_GotFocus()
tmp = CmbServiceTaxID.Text
Call TableMst_ServiceTax(" WHere Flag = 'A' ")
CmbServiceTaxID.Clear
If RsMst_ServiceTax.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ServiceTax.RecordCount
    CmbServiceTaxID.AddItem RsMst_ServiceTax!Description
    RsMst_ServiceTax.MoveNext
Next
CmbServiceTaxID.Text = tmp
End Sub

Private Sub CmbServiceTaxID_LostFocus()
If CmbServiceTaxID.Text = "" Then
   mServiceTaxID = Null
   TxtServiceTaxNAVCode = ""
   Exit Sub
End If
RsMst_ServiceTax.MoveFirst
RsMst_ServiceTax.Find "Description = '" & CmbServiceTaxID.Text & "'"
If RsMst_ServiceTax.EOF Then
   MsgBox "Invalid selection "
   CmbServiceTaxID.SetFocus
   Exit Sub
End If
mServiceTaxID = RsMst_ServiceTax!ServiceTaxID
TxtServiceTaxNAVCode = IIf(IsNull(RsMst_ServiceTax!NavCode), "", RsMst_ServiceTax!NavCode)

End Sub

Private Sub CmbHRVarietyID_GotFocus()
If CmbCommodityGroupID.Text = "" Then
   MsgBox "Select Commodity Group!!!"
   Exit Sub
End If
tmp1 = CmbHRVarietyID.Text
tmp = "select * from Mst_Variety Where CommodityID = " & mHRCommodityID & " And Flag = 'A'"
Call TmpTableHRPro

CmbHRVarietyID.Clear
If TmpRsHRPro.RecordCount = 0 Then
   MsgBox "No HR Variety found for this Commodity"
   CmbHRVarietyID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRsHRPro.RecordCount
    CmbHRVarietyID.AddItem TmpRsHRPro!VarietyName
    TmpRsHRPro.MoveNext
Next
CmbHRVarietyID.Text = tmp1
End Sub

Private Sub CmbHRVarietyID_LostFocus()
If CmbHRVarietyID.Text = "" Then
   Exit Sub
End If
TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "VarietyName = '" & CmbHRVarietyID.Text & "'"

If TmpRsHRPro.EOF Then
   MsgBox "Invalid HR Variety selection "
   CmbHRVarietyID.SetFocus
   Exit Sub
End If
mHRVarietyID = TmpRsHRPro!VarietyID
End Sub
