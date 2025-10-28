VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmDeposit 
   Caption         =   "Deposit"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Caption         =   "Hologram No "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5520
      TabIndex        =   122
      Top             =   5160
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox Check2 
         BackColor       =   &H008080FF&
         Caption         =   "Generate Excel"
         Height          =   255
         Left            =   120
         TabIndex        =   140
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtHologramNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   125
         Top             =   360
         Width           =   2880
      End
      Begin VB.CommandButton CmdSaveHologram 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         Height          =   400
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   840
         Width           =   1440
      End
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         Height          =   400
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   $"FrmDeposit.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1260
         Left            =   3120
         TabIndex        =   146
         Top             =   1800
         Width           =   4380
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Hologram No"
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
         TabIndex        =   126
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame Frame0 
      Height          =   1845
      Left            =   360
      TabIndex        =   109
      Top             =   5280
      Visible         =   0   'False
      Width           =   13875
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   2040
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   450
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59375617
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   480
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   450
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59375617
         CurrentDate     =   36494
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   11760
         TabIndex        =   130
         Top             =   870
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   5160
         Sorted          =   -1  'True
         TabIndex        =   40
         Top             =   450
         Width           =   3645
      End
      Begin VB.ComboBox CmbSAgent 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   825
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.ComboBox CmbSExchangeType 
         Height          =   315
         Left            =   11610
         Sorted          =   -1  'True
         TabIndex        =   42
         Top             =   450
         Width           =   2205
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   8880
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   450
         Width           =   2685
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   3525
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   450
         Width           =   1605
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12795
         TabIndex        =   110
         Top             =   870
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   90
         TabIndex        =   111
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   2340
         TabIndex        =   137
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   780
         TabIndex        =   136
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label42 
         Caption         =   "CMP"
         Height          =   300
         Left            =   6615
         TabIndex        =   135
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label39 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   9855
         TabIndex        =   134
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label38 
         Caption         =   "Exchange Type"
         Height          =   225
         Left            =   12165
         TabIndex        =   133
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label37 
         Caption         =   "Agent Name"
         Height          =   300
         Left            =   120
         TabIndex        =   132
         Top             =   855
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label36 
         Caption         =   "Location ID"
         Height          =   300
         Left            =   3960
         TabIndex        =   131
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.TextBox TxtG_UID 
      Height          =   495
      Left            =   720
      TabIndex        =   138
      Top             =   10080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   8655
      Left            =   60
      ScaleHeight     =   8595
      ScaleWidth      =   15075
      TabIndex        =   48
      Top             =   120
      Width           =   15135
      Begin VB.PictureBox Picture2 
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1875
         ScaleWidth      =   14835
         TabIndex        =   63
         Top             =   120
         Width           =   14895
         Begin VB.TextBox TxtManualCDTFNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
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
            MaxLength       =   18
            TabIndex        =   2
            Top             =   1350
            Width           =   1815
         End
         Begin VB.ComboBox CmbAgent 
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
            Left            =   4800
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   1320
            Width           =   6135
         End
         Begin VB.CommandButton CmdAgent 
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
            Left            =   11040
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox TxtPreFix 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   3
            EndProperty
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
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox TxtCVariation 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   3
            EndProperty
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
            Left            =   13680
            Locked          =   -1  'True
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox TxtCLotSize 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   3
            EndProperty
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
            Left            =   13680
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   900
            Width           =   1095
         End
         Begin VB.CommandButton CmdContractID 
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
            Left            =   11040
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton CmdCMPID 
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
            Left            =   11040
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   105
            Width           =   375
         End
         Begin VB.CommandButton CmdCommodityID 
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
            Left            =   8160
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   510
            Width           =   375
         End
         Begin VB.TextBox TxtCStdDed 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   3
            EndProperty
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
            Left            =   13680
            Locked          =   -1  'True
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   510
            Width           =   1095
         End
         Begin VB.TextBox TxtCGunnyWt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   3
            EndProperty
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
            Left            =   13680
            Locked          =   -1  'True
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   105
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   375
            Left            =   4800
            TabIndex        =   5
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox TxtCxTFNo 
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   150
            Width           =   1815
         End
         Begin VB.TextBox TxtAgent 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   3
            EndProperty
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
            Left            =   4800
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1365
            Width           =   6135
         End
         Begin MSComCtl2.DTPicker TxtCxTFDate 
            Height          =   360
            Left            =   1200
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   555
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            Format          =   59375617
            CurrentDate     =   36494
         End
         Begin VB.ComboBox CmbContractID 
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
            Left            =   5040
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   960
            Width           =   5895
         End
         Begin VB.TextBox TxtContractID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   3
            EndProperty
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
            Left            =   5040
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   960
            Width           =   5895
         End
         Begin VB.ComboBox CmbCommodityID 
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
            Left            =   4800
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   510
            Width           =   3255
         End
         Begin VB.ComboBox CmbCMPID 
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
            Left            =   4800
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   105
            Width           =   6135
         End
         Begin VB.TextBox TxtCommodityID 
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
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   510
            Width           =   3255
         End
         Begin VB.TextBox TxtCMPID 
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
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   105
            Width           =   6135
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
            ItemData        =   "FrmDeposit.frx":0092
            Left            =   1200
            List            =   "FrmDeposit.frx":0094
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox TxtExchangeTypeID 
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   106
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox TxtCxTFDispalyDate 
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   510
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label35 
            Caption         =   "Manual CDTF #"
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
            Left            =   120
            TabIndex        =   129
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "Location ID"
            Height          =   255
            Left            =   11640
            TabIndex        =   121
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "% Variation "
            Height          =   255
            Left            =   11880
            TabIndex        =   105
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "Lot Size"
            Height          =   255
            Left            =   11880
            TabIndex        =   103
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Std. Deduction"
            Height          =   255
            Left            =   11880
            TabIndex        =   94
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Gunny Wt. per Bag"
            Height          =   255
            Left            =   11880
            TabIndex        =   92
            Top             =   180
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "CDTF No"
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
            Left            =   120
            TabIndex        =   73
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "CMP"
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
            Left            =   3120
            TabIndex        =   72
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label4 
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
            Height          =   375
            Left            =   3120
            TabIndex        =   71
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "CDTF Date"
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
            Left            =   120
            TabIndex        =   70
            Top             =   600
            Width           =   1215
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
            Left            =   120
            TabIndex        =   69
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label28 
            Caption         =   "space reservation"
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
            Left            =   3120
            TabIndex        =   68
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label29 
            Caption         =   "Agent"
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
            Left            =   3120
            TabIndex        =   67
            Top             =   1440
            Width           =   855
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6375
         Left            =   120
         TabIndex        =   49
         Top             =   2160
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   11245
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Client Detail"
         TabPicture(0)   =   "FrmDeposit.frx":0096
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(1)=   "Pic_Client"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Truck Details"
         TabPicture(1)   =   "FrmDeposit.frx":00B2
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "MSHFlexGrid3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Pic_TTD"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "GSL Details"
         TabPicture(2)   =   "FrmDeposit.frx":00CE
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "MSHFlexGrid4"
         Tab(2).Control(1)=   "Pic_GCTD"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Log History"
         TabPicture(3)   =   "FrmDeposit.frx":00EA
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "TxtUpdated"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "TxtCreated"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label47"
         Tab(3).Control(3)=   "Label46"
         Tab(3).ControlCount=   4
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
            Left            =   -73515
            Locked          =   -1  'True
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   1245
            Width           =   9315
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
            Left            =   -73515
            Locked          =   -1  'True
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   480
            Width           =   9330
         End
         Begin VB.PictureBox Pic_TTD 
            Height          =   1455
            Left            =   120
            ScaleHeight     =   1395
            ScaleWidth      =   14595
            TabIndex        =   82
            Top             =   780
            Width           =   14655
            Begin VB.CommandButton CmdTTDDeleteList 
               Caption         =   "Delete From List"
               Height          =   375
               Left            =   12240
               TabIndex        =   118
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txtTTDRemarks 
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
               Left            =   10215
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Top             =   420
               Width           =   3975
            End
            Begin VB.CommandButton CmdTTDAdd2List 
               Caption         =   "Add to List"
               Height          =   375
               Left            =   10680
               TabIndex        =   21
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TxtTTDTareWt 
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
               Left            =   6600
               TabIndex        =   18
               Top             =   420
               Width           =   1815
            End
            Begin VB.TextBox TxtTTDTotalWt 
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
               Left            =   5040
               TabIndex        =   17
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox TxtWeighBridgNo 
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
               Left            =   2520
               TabIndex        =   16
               Top             =   420
               Width           =   2535
            End
            Begin VB.TextBox TxtTTDTruckNo 
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
               Left            =   120
               TabIndex        =   15
               Top             =   420
               Width           =   2415
            End
            Begin VB.TextBox TxtTTDGrossWt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   8400
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   420
               Width           =   1815
            End
            Begin VB.Label Label6 
               Caption         =   "Remarks"
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
               Left            =   10920
               TabIndex        =   88
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label17 
               Caption         =   "Tare Weight"
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
               Left            =   6840
               TabIndex        =   87
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label16 
               Caption         =   "Total Weight"
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
               Left            =   5160
               TabIndex        =   86
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label15 
               Caption         =   "Weighbridge Slip No."
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
               Left            =   2880
               TabIndex        =   85
               Top             =   120
               Width           =   1935
            End
            Begin VB.Label Label30 
               Caption         =   "Truck No."
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
               Left            =   600
               TabIndex        =   84
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label18 
               Caption         =   "Gross Weight"
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
               Left            =   8640
               TabIndex        =   83
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_GCTD 
            Height          =   1935
            Left            =   -74880
            ScaleHeight     =   1875
            ScaleWidth      =   14595
            TabIndex        =   74
            Top             =   660
            Width           =   14655
            Begin VB.TextBox TxtGodAdd 
               Height          =   285
               Left            =   5280
               TabIndex        =   149
               Top             =   1440
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.CommandButton CmdGCTDeleteList 
               Caption         =   "Delete From List"
               Height          =   350
               Left            =   13200
               TabIndex        =   117
               Top             =   1395
               Width           =   1335
            End
            Begin VB.TextBox TxtAdjGrossWt 
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
               Left            =   8160
               Locked          =   -1  'True
               TabIndex        =   115
               TabStop         =   0   'False
               Top             =   420
               Width           =   1575
            End
            Begin VB.CommandButton CmdGSL 
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
               Left            =   720
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox TxtLotStock 
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
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   108
               TabStop         =   0   'False
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox TxtAdjWt 
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
               Left            =   6600
               TabIndex        =   27
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox TxtGCTDDematWt 
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
               Left            =   12840
               Locked          =   -1  'True
               TabIndex        =   100
               TabStop         =   0   'False
               Top             =   420
               Width           =   1680
            End
            Begin VB.TextBox TxtGCTDStdDeduction 
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
               Left            =   11280
               Locked          =   -1  'True
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox TxtGCTDGunnyWt 
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
               Left            =   9720
               Locked          =   -1  'True
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   420
               Width           =   1575
            End
            Begin VB.ComboBox CmbGCTDClient 
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
               ItemData        =   "FrmDeposit.frx":0106
               Left            =   5190
               List            =   "FrmDeposit.frx":0108
               Sorted          =   -1  'True
               TabIndex        =   28
               Top             =   900
               Width           =   3975
            End
            Begin VB.TextBox TxtGCTDNoofBags 
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
               Left            =   3840
               TabIndex        =   25
               Top             =   420
               Width           =   1215
            End
            Begin VB.TextBox TxtGCTDGrossWt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   5040
               TabIndex        =   26
               Top             =   420
               Width           =   1575
            End
            Begin VB.CommandButton CmdGCTDAdd2List 
               Caption         =   "Add to List"
               Height          =   350
               Left            =   11880
               TabIndex        =   30
               Top             =   1395
               Width           =   1335
            End
            Begin VB.TextBox TxtGCTDremarks 
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
               Left            =   10080
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   29
               Top             =   840
               Width           =   4455
            End
            Begin VB.ComboBox CmbGodownNo 
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
               ItemData        =   "FrmDeposit.frx":010A
               Left            =   120
               List            =   "FrmDeposit.frx":010C
               Sorted          =   -1  'True
               TabIndex        =   22
               Top             =   420
               Width           =   1335
            End
            Begin VB.ComboBox CmbStack 
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
               ItemData        =   "FrmDeposit.frx":010E
               Left            =   1440
               List            =   "FrmDeposit.frx":0110
               Sorted          =   -1  'True
               TabIndex        =   23
               Top             =   420
               Width           =   1215
            End
            Begin VB.ComboBox CmbLot 
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
               ItemData        =   "FrmDeposit.frx":0112
               Left            =   2640
               List            =   "FrmDeposit.frx":0114
               Sorted          =   -1  'True
               TabIndex        =   24
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label19 
               Caption         =   "Adj Gross Wt."
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
               Left            =   8400
               TabIndex        =   116
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label14 
               Caption         =   "Adjusted Wt."
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
               Left            =   6840
               TabIndex        =   114
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Lot Balance"
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
               Left            =   1440
               TabIndex        =   107
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label40 
               Caption         =   "Electronic Weight"
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
               Left            =   12870
               TabIndex        =   101
               Top             =   105
               Width           =   1575
            End
            Begin VB.Label Label34 
               Caption         =   "Std. Deduction"
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
               Left            =   11400
               TabIndex        =   96
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label27 
               Caption         =   "Gunny Weight"
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
               Left            =   9960
               TabIndex        =   90
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label20 
               Caption         =   "Godown No"
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
               Left            =   120
               TabIndex        =   81
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label21 
               Caption         =   "Stack No"
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
               Left            =   1680
               TabIndex        =   80
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label22 
               Caption         =   "Lot No"
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
               Left            =   2880
               TabIndex        =   79
               Top             =   120
               Width           =   735
            End
            Begin VB.Label Label23 
               Caption         =   "No Of Bags"
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
               Left            =   3960
               TabIndex        =   78
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label24 
               Caption         =   "Gross Weight"
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
               Left            =   5280
               TabIndex        =   77
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label LblClient_WHR_RRN 
               Caption         =   "Client Name ~ DP Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3960
               TabIndex        =   76
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label25 
               Caption         =   "Remarks"
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
               Left            =   9240
               TabIndex        =   75
               Top             =   960
               Width           =   975
            End
         End
         Begin VB.PictureBox Pic_Client 
            Height          =   1215
            Left            =   -74880
            ScaleHeight     =   1155
            ScaleWidth      =   14595
            TabIndex        =   50
            Top             =   660
            Width           =   14655
            Begin VB.ComboBox CmbNonNcdexClient 
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
               Left            =   1680
               Sorted          =   -1  'True
               TabIndex        =   13
               Top             =   120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CommandButton CmdClient 
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
               Height          =   240
               Left            =   1200
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   120
               Width           =   375
            End
            Begin VB.ComboBox CmbClientDP 
               Height          =   315
               Left            =   5160
               Sorted          =   -1  'True
               TabIndex        =   12
               Top             =   480
               Visible         =   0   'False
               Width           =   5055
            End
            Begin VB.CommandButton CmdClientAdd2List 
               Caption         =   "Add to List"
               Height          =   375
               Left            =   13200
               TabIndex        =   14
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtClientWeight 
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
               Left            =   11640
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox TxtClientNoBags 
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
               Left            =   10200
               Locked          =   -1  'True
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox TxtDPName 
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
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   480
               Width           =   3855
            End
            Begin VB.TextBox TxtDPID 
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
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   480
               Width           =   1215
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
               Height          =   375
               Left            =   120
               TabIndex        =   11
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox TxtClientName 
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
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   480
               Width           =   3495
            End
            Begin VB.Label lblNameDid 
               Caption         =   "ClientDPID"
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
               Left            =   6240
               TabIndex        =   62
               Top             =   180
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label13 
               Caption         =   "Weight"
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
               Left            =   12000
               TabIndex        =   61
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label12 
               Caption         =   "No Bags"
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
               Left            =   10440
               TabIndex        =   60
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label11 
               Caption         =   "DP Name "
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
               Left            =   7800
               TabIndex        =   59
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Label10 
               Caption         =   "DP ID"
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
               Left            =   5400
               TabIndex        =   58
               Top             =   180
               Width           =   735
            End
            Begin VB.Label Label8 
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
               Left            =   2400
               TabIndex        =   57
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Label9 
               Caption         =   "Client ID"
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
               Left            =   360
               TabIndex        =   56
               Top             =   180
               Width           =   975
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   145
            Top             =   2640
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   6376
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   4335
            Left            =   -74880
            TabIndex        =   147
            Top             =   1920
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7646
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   3930
            Left            =   120
            TabIndex        =   148
            Top             =   2280
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   6932
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label47 
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
            Height          =   255
            Left            =   -74760
            TabIndex        =   144
            Top             =   1305
            Width           =   975
         End
         Begin VB.Label Label46 
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
            Left            =   -74760
            TabIndex        =   143
            Top             =   540
            Width           =   1695
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   930
      Left            =   75
      ScaleHeight     =   870
      ScaleWidth      =   14955
      TabIndex        =   0
      Top             =   8880
      Width           =   15015
      Begin VB.CommandButton CmdPrintPreview 
         Caption         =   "Print Preview"
         Height          =   375
         Left            =   12165
         TabIndex        =   139
         Top             =   75
         Width           =   2655
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   390
         Left            =   12165
         TabIndex        =   119
         Top             =   435
         Width           =   2655
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   9765
         TabIndex        =   34
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   7350
         TabIndex        =   47
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   4935
         TabIndex        =   46
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2520
         TabIndex        =   45
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   9765
         TabIndex        =   33
         Top             =   75
         Width           =   2415
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   7350
         TabIndex        =   36
         Top             =   75
         Width           =   2415
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4935
         TabIndex        =   35
         Top             =   75
         Width           =   2415
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2520
         TabIndex        =   32
         Top             =   75
         Width           =   2415
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   75
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mLocationID, mSLocationId, mClientIDRow, mContractClientIDRow, mCMPID, mCommodityID, mContractID, mGodownID, mGSLID, mGCTDClientID As Variant
Dim mAgentID As Variant
Dim mClientIDRowGsl As Variant
Dim mExchangeTypeID As Double
Dim mCMPName As Double
Dim mCmbStack, mCmbLot, mTimeStamp, msSM_prefix As Variant
Dim PntPrv  As Variant
Dim Wc As Variant


Sub UpdateClientGridStock()
For i = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 7) <> "" Then
       MSHFlexGrid2.TextMatrix(i, 5) = 0 '"No of Bags"
       MSHFlexGrid2.TextMatrix(i, 6) = 0 '"Weight"
    End If
Next
MSHFlexGrid2.TextMatrix(0, 5) = 0 '"No of Bags"
MSHFlexGrid2.TextMatrix(0, 6) = 0 '"Weight"

For i = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(i, 16) <> "" Then
       Call UpdateDataClient(MSHFlexGrid4.TextMatrix(i, 16), MSHFlexGrid4.TextMatrix(i, 4), MSHFlexGrid4.TextMatrix(i, 10))
    End If
Next
End Sub

Sub UpdateDataClient(ClientRowID As Variant, NoofBags As Variant, EWeight As Variant)
For i = 2 To MSHFlexGrid2.Rows - 1
    If Val(MSHFlexGrid2.TextMatrix(i, 7)) = Val(ClientRowID) Then
       MSHFlexGrid2.TextMatrix(i, 5) = Val(MSHFlexGrid2.TextMatrix(i, 5)) + NoofBags '"No of Bags"
       MSHFlexGrid2.TextMatrix(i, 6) = Val(MSHFlexGrid2.TextMatrix(i, 6)) + EWeight '"Weight"
       
       MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + NoofBags '"No of Bags"
       MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + EWeight '"Weight"
       
       Exit For
    End If
Next

End Sub


Private Sub Check1_Click()
CmbContractID.Enabled = False
Check1.Enabled = False
If mExchangeTypeID <> 1 Then Exit Sub
Check1.Enabled = True
If Check1.Value = 1 Then
   CmbContractID.Enabled = True
End If
CmdClientAdd2List.Enabled = True
Call Grid_Head_GCTD
Call Grid_Head_Client

End Sub

Private Sub CmbAgent_GotFocus()
tmp = CmbAgent.Text
Call TableMst_Agent
CmbAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Agent.RecordCount
    CmbAgent.AddItem RsMst_Agent!AgentName
    RsMst_Agent.MoveNext
Next
CmbAgent.Text = tmp
End Sub
'----
Private Sub CmbAgent_LostFocus()
If CmbAgent.Text = "" Then
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "agentName = '" & CmbAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid selection "
   CmbAgent.SetFocus
   Exit Sub
End If
mAgentID = RsMst_Agent!AgentID
End Sub

Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      TxtClientID.SetFocus
   End If
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow
CmdClientAdd2List.SetFocus
End Sub



Private Sub CmbGCTDClient_GotFocus()
tmp = CmbGCTDClient.Text
CmbGCTDClient.Clear
If MSHFlexGrid2.Rows < 4 Then Exit Sub

For i = 2 To MSHFlexGrid2.Rows - 2
   CmbGCTDClient.AddItem MSHFlexGrid2.TextMatrix(i, 2) & " ~ " & MSHFlexGrid2.TextMatrix(i, 4)
Next i
CmbGCTDClient.Text = tmp
End Sub


Private Sub CmbGCTDClient_LostFocus()
If CmbGCTDClient.Text = "" Then Exit Sub

For i = 2 To MSHFlexGrid2.Rows - 2
    
   If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(i, 2) & " ~ " & MSHFlexGrid2.TextMatrix(i, 4) Then
      mGCTDClientID = Val(MSHFlexGrid2.TextMatrix(i, 7))
      Exit Sub
   End If
Next i

MsgBox "Invalid selection "
CmbGCTDClient.SetFocus

End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text <> "" Then
    tmp = CmbGodownNo.Text
    Call TableMst_Godown(" Where CloseDate is null And CMPID = " & mCMPID)
    CmbGodownNo.Clear
    If RsMst_Godown.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For i = 1 To RsMst_Godown.RecordCount
        CmbGodownNo.AddItem RsMst_Godown!GodownNo
        RsMst_Godown.MoveNext
    Next
    CmbGodownNo.Text = tmp
Else
   MsgBox "Please Select Commodity!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If
End Sub
'----
Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbGodownNo.Enabled = True Then
      CmbGodownNo.SetFocus
   End If
   Exit Sub
End If
If mGodownID <> RsMst_Godown!GodownID Then
   CmbStack.Text = ""
   CmbLot.Text = ""
   mCmbStack = ""
   mCmbLot = ""
End If
mGodownID = RsMst_Godown!GodownID
TxtGodAdd = IIf(IsNull(RsMst_Godown!Address), "", RsMst_Godown!Address)

End Sub

Private Sub CmbLot_GotFocus()
If CmbStack.Text = "" Then
  MsgBox "Please Select Stack No !!!"
  CmbStack.SetFocus
  Exit Sub
End If
If CmbCommodityID.Text = "" Then
  MsgBox "Please Select Commodity !!!"
  CmbCommodityID.SetFocus
  Exit Sub
End If
tmp = CmbLot.Text
Call TableMst_GSL(" Where SpillageFlag <> '1' And  " & Gen_ExpireGSLFlag & "  And  CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "'")
CmbLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If

For i = 1 To RsMst_GSL.RecordCount
    CmbLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next i



'For i = 1 To RsMst_GSL.RecordCount
'    CmbLot.AddItem RsMst_GSL!LotNo
'    RsMst_GSL.MoveNext
'Next
CmbLot.Text = tmp
End Sub
'----
Private Sub CmbLot_LostFocus()
Dim LotBagStock  As Variant
If CmbLot.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "LotNo = '" & CmbLot.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbLot.SetFocus
   Exit Sub
End If
TxtLotStock = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight)
LotBagStock = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags)


If Val(TxtLotStock) = 0 And LotBagStock = 0 Then
   mGSLID = RsMst_GSL!GSLID
   mClientIDRowGsl = Null
   Exit Sub
End If

If RsMst_GSL!CommodityID <> mCommodityID Then
   MsgBox "GSL Contain Stock of other Commodity!!! "
   CmbLot.SetFocus
   Exit Sub
End If
If RsMst_GSL!ExchangeTypeId <> mExchangeTypeID Then
   MsgBox "GSL Contain Stock of other Exchange Type!!! "
   CmbLot.SetFocus
   Exit Sub
End If

If mExchangeTypeID = 2 Then
   If RsMst_GSL!KRNo <> 0 Then
      MsgBox "KR No is Mapped against Selected Lot!!! "
      CmbLot.SetFocus
      Exit Sub
   End If
End If
If mExchangeTypeID = 1 Then
   If RsMst_GSL!DepositeWeight <> 0 Or RsMst_GSL!DepositeBags <> 0 Then
      MsgBox "This Lot is already having stock as shown. You can not add further stock into this Lot!!! "
      CmbLot.SetFocus
      Exit Sub
   End If
End If

If Not IsNull(RsMst_GSL!ISINID) And RsMst_GSL!ISINID <> "" And RsMst_GSL!ISINID <> 0 Then
   If mExchangeTypeID = 2 Then
      MsgBox "ISIN Mapped against Selected Lot "
   Else
      MsgBox "CDC Mapped against Selected Lot "
   End If
   CmbLot.SetFocus
   Exit Sub
End If


mGSLID = RsMst_GSL!GSLID
mClientIDRowGsl = RsMst_GSL!ClientIDRow

End Sub

Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub

Private Sub CmbNonNcdexClient_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 1
   TxtTTDTruckNo.SetFocus
End If
End Sub

Private Sub CmbNonNcdexClient_LostFocus()
If CmbNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow
Dim ans As Variant
ans = GetClientNameByRow(RsMst_Client!ClientIDRow)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)

End Sub
Private Sub CmbStack_GotFocus()

If CmbGodownNo.Text = "" Or mGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  Exit Sub
End If
tmp = CmbStack.Text
Call TableMst_GSL(" Where SpillageFlag <> '1' And CMPID = " & mCMPID & " And GodownID = " & mGodownID, "D")
CmbStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_GSL.RecordCount
    CmbStack.AddItem RsMst_GSL!StackNo
    RsMst_GSL.MoveNext
Next

CmbStack.Text = tmp
End Sub
'----
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
If mCmbStack <> CmbStack.Text Then
   CmbLot.Text = ""
   
   mCmbLot = ""
End If
mCmbStack = CmbStack.Text
End Sub

Private Sub CmdAgent_Click()
FrmMst_Agent.Show
End Sub

Private Sub CmdCancel_Click()
TxtHologramNo = ""
Frame3.Visible = False
End Sub

Private Sub CmdClient_Click()
FrmMst_Client.Show
End Sub
Function CheckDuplicateClient() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = TxtClientID And MSHFlexGrid2.TextMatrix(u, 3) = TxtDPID Then
          If LCase(CmdClientAdd2List.Caption) = "add to list" Then
             RetVal = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                RetVal = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateClient = RetVal
End Function

Private Sub CmdClientAdd2List_Click()
Dim r As Variant

If CmbClientDP.Visible = True Then
   MsgBox "Invalid selection !!! "
   Exit Sub
End If
If CmbNonNcdexClient.Visible = True And Check1.Value = 0 Then
    If CmbNonNcdexClient.Text = "" Then
       MsgBox "Invalid selection !!! "
       CmbNonNcdexClient.SetFocus
       Exit Sub
    End If
End If
If TxtClientID = "" Then
   MsgBox "Invalid Client ID !!! "
   Exit Sub
End If

If CheckDuplicateClient = True Then
  MsgBox "Duplicate Client ID  not allowd !!!"
  TxtClientID.SetFocus
  Exit Sub
End If

If CmdClientAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
Else
  r = MSHFlexGrid2.Rows - 1
  MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If
If MSHFlexGrid2.Rows > 7 Then
   MsgBox "You can add only 4 clients detail in one CDTF "
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   Exit Sub
End If

MSHFlexGrid2.TextMatrix(r, 1) = TxtClientID
MSHFlexGrid2.TextMatrix(r, 2) = TxtClientName
MSHFlexGrid2.TextMatrix(r, 3) = TxtDPID
MSHFlexGrid2.TextMatrix(r, 4) = TxtDPName
MSHFlexGrid2.TextMatrix(r, 5) = TxtClientNoBags
MSHFlexGrid2.TextMatrix(r, 6) = TxtClientWeight
MSHFlexGrid2.TextMatrix(r, 7) = mClientIDRow

MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(r, 5))
MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(r, 6))
Call ClearClientTabControl
TxtClientID.Locked = False
If mExchangeTypeID <> 1 Then
   TxtClientID.SetFocus
Else
   CmbNonNcdexClient.SetFocus
End If
End Sub
Sub ClearClientTabControl()

TxtClientID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
TxtClientNoBags = ""
TxtClientWeight = ""
CmbClientDP.Text = ""
CmbClientDP.Visible = False
CmdClientAdd2List.Caption = "Add to List"
End Sub
Function CheckDuplicateGSL() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid4.Rows > 3 Then
   For u = 1 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(u, 1) = CmbGodownNo.Text And _
          MSHFlexGrid4.TextMatrix(u, 2) = CmbStack.Text And _
          MSHFlexGrid4.TextMatrix(u, 3) = CmbLot.Text Then
          If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
             RetVal = True
             Exit For
          Else
             If u <> MSHFlexGrid4.RowSel Then
                RetVal = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateGSL = RetVal
End Function


Private Sub CmdGCTDeleteList_Click()
i = MSHFlexGrid4.RowSel
If i > 1 Then
   If MSHFlexGrid4.TextMatrix(i, 1) <> "" Then
      MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(i, 4))
      MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(i, 5))
      MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) - Val(MSHFlexGrid4.TextMatrix(i, 6))
      MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) - Val(MSHFlexGrid4.TextMatrix(i, 7))
      MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) - Val(MSHFlexGrid4.TextMatrix(i, 8))
      MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) - Val(MSHFlexGrid4.TextMatrix(i, 9))
      MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) - Val(MSHFlexGrid4.TextMatrix(i, 10))
      For RO = i To MSHFlexGrid4.Rows - 2
          For c = 0 To MSHFlexGrid4.Cols - 1
               MSHFlexGrid4.TextMatrix(RO, c) = MSHFlexGrid4.TextMatrix(RO + 1, c)
          Next
      Next
      MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 1
   End If
End If
Call ClearGCTDTabControl
End Sub

Private Sub CmdGCTDAdd2List_Click()
Dim r As Variant

If (CmbGodownNo.Text) = "" Then
   MsgBox "Blank Godown No Not Allowed !!! "
   CmbGodownNo.SetFocus
   Exit Sub
End If

If (CmbStack.Text) = "" Then
   MsgBox "Blank Stack  Not Allowed !!! "
   CmbStack.SetFocus
   Exit Sub
End If
If (CmbLot.Text) = "" Then
   MsgBox "Blank Lot Not Allowed !!! "
   CmbLot.SetFocus
   Exit Sub
End If

If CheckDuplicateGSL() = True Then
   MsgBox "Duplicate G.S.L.  Not Allowed !!! "
   CmbGodownNo.SetFocus
   Exit Sub
End If

If Val(TxtGCTDNoofBags.Text) <= 0 Then
   MsgBox "Invalid No of Bags !!! "
   TxtGCTDNoofBags.SetFocus
   Exit Sub
End If

If (CmbGCTDClient.Text) = "" Then
   MsgBox "Blank Client ID Not Allowed !!! "
   CmbGCTDClient.SetFocus
   Exit Sub
End If

If Val(TxtGCTDGrossWt) <= 0 Then
   MsgBox "Invalid Weight !!! "
   TxtGCTDGrossWt.SetFocus
   Exit Sub
End If

If Val(TxtCLotSize) > 0 Then
   tmp = 100 + Val(TxtCVariation)
   tmp = (tmp / 100) * Val(TxtCLotSize)
   If Val(TxtGCTDDematWt) + Val(TxtLotStock) >= tmp Then
      MsgBox "Lot balance + Electronic weight exceeds lot size limit in Commodity Master !!! "
      TxtGCTDGrossWt.SetFocus
      Exit Sub
   End If
End If
tmp = Validate_ComBags(mCommodityID, Val(TxtGCTDNoofBags), Val(TxtGCTDGrossWt))
If Len(tmp) > 1 Then
   MsgBox tmp
   Exit Sub
End If
If mClientIDRowGsl <> 0 Then
   If mClientIDRowGsl <> (mGCTDClientID) Then
      MsgBox "GSL Contain Stock of other Client  "
      If CmbLot.Enabled = True Then
         CmbLot.SetFocus
      End If
      Exit Sub
   End If
End If

'MSHFlexGrid4.TextMatrix(1, 6) = "Adjusted Wt."
'MSHFlexGrid4.TextMatrix(1, 7) = "Adj.Gross Wt."
'MSHFlexGrid4.TextMatrix(1, 8) = "Gunny Weight"
'MSHFlexGrid4.TextMatrix(1, 9) = "Std. Deduction"
'MSHFlexGrid4.TextMatrix(1, 10) = "Demat Weight"
'MSHFlexGrid4.TextMatrix(1, 11) = "Client ID"
'MSHFlexGrid4.TextMatrix(1, 12) = "Remark"
'MSHFlexGrid4.TextMatrix(1, 13) = "GSL ID"
'MSHFlexGrid4.TextMatrix(1, 14) = "Godown Close"


If CmdGCTDAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid4.RowSel
   MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(r, 4))
   MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(r, 5))
   MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) - Val(MSHFlexGrid4.TextMatrix(r, 6))
   MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) - Val(MSHFlexGrid4.TextMatrix(r, 7))
   MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) - Val(MSHFlexGrid4.TextMatrix(r, 8))
   MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) - Val(MSHFlexGrid4.TextMatrix(r, 9))
   MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) - Val(MSHFlexGrid4.TextMatrix(r, 10))
Else
  r = MSHFlexGrid4.Rows - 1
  MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
End If

MSHFlexGrid4.TextMatrix(r, 1) = CmbGodownNo.Text
MSHFlexGrid4.TextMatrix(r, 2) = CmbStack.Text
MSHFlexGrid4.TextMatrix(r, 3) = CmbLot.Text
MSHFlexGrid4.TextMatrix(r, 4) = TxtGCTDNoofBags

MSHFlexGrid4.TextMatrix(r, 5) = TxtGCTDGrossWt
MSHFlexGrid4.TextMatrix(r, 6) = TxtAdjWt
MSHFlexGrid4.TextMatrix(r, 7) = TxtAdjGrossWt

MSHFlexGrid4.TextMatrix(r, 8) = TxtGCTDGunnyWt
MSHFlexGrid4.TextMatrix(r, 9) = TxtGCTDStdDeduction
MSHFlexGrid4.TextMatrix(r, 10) = TxtGCTDDematWt
MSHFlexGrid4.TextMatrix(r, 11) = CmbGCTDClient.Text
MSHFlexGrid4.TextMatrix(r, 12) = TxtGCTDremarks
MSHFlexGrid4.TextMatrix(r, 13) = mGSLID
MSHFlexGrid4.TextMatrix(r, 15) = mGodownID
MSHFlexGrid4.TextMatrix(r, 16) = mGCTDClientID
MSHFlexGrid4.TextMatrix(r, 18) = TxtGodAdd


MSHFlexGrid4.Col = 5
MSHFlexGrid4.Row = r
MSHFlexGrid4.CellBackColor = &HFFC0C0
 
 
MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(r, 4))
MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(r, 5))
MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(r, 6))
MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(r, 7))
MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(r, 8))
MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(r, 9))
MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(r, 10))

Call ClearGCTDTabControl
CmbGodownNo.SetFocus
End Sub
Private Sub ClearGCTDTabControl()
TxtGodAdd = ""
CmbGodownNo.Text = ""
CmbStack.Text = ""
CmbLot.Text = ""
CmbGCTDClient.Text = ""
TxtGCTDNoofBags = ""
TxtGCTDGrossWt = ""
TxtGCTDGunnyWt = ""
TxtGCTDStdDeduction = ""
TxtGCTDDematWt = ""
TxtGCTDremarks = ""
TxtRecClient = ""
TxtRecNCMSL = ""
TxtAdjGrossWt = ""
TxtAdjWt = ""
TxtAdjWt.Locked = True
CmbNonNcdexClient.Text = ""
mGSLID = Null
CmdGCTDAdd2List.Caption = "Add to List"
CmdGCTDeleteList.Enabled = False


CmbGodownNo.Enabled = True
CmbStack.Enabled = True
CmbLot.Enabled = True

End Sub

Private Sub CmdGSL_Click()

Gen_GSLID(0) = False
Gen_GSLID(1) = mCMPID
Gen_GSLID(2) = mGodownID
Gen_GSLID(3) = CmbStack.Text
Gen_GSLID(4) = CmbLot.Text
Gen_GSLID(5) = CmbCMPID.Text
Gen_GSLID(6) = CmbGodownNo.Text


FrmMst_GSL.Show
End Sub

Private Sub CmdPrint_Click()

tmp = Format(Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(0, 5)), "#######0.000")

If Val(tmp) <> 0 Then
   MsgBox "Gross Weight in 'Truck Detail' Tab and 'GSL Detail' tab are not equal !!!!"
   Exit Sub
End If

If Val(MSHFlexGrid2.TextMatrix(0, 5)) <> Val(MSHFlexGrid4.TextMatrix(0, 4)) Then
   MsgBox "No of Bags in 'Client Detail' Tab and 'GSL Detail' tab are not equal !!!!"
   Exit Sub
End If
If Val(MSHFlexGrid2.TextMatrix(0, 6)) <> Val(MSHFlexGrid4.TextMatrix(0, 10)) Then
   MsgBox "Electronic Weight in 'Client Detail' Tab and 'GSL Detail' tab are not equal !!!!"
   Exit Sub
End If

Check2.Visible = False
If LCase(UserType) = "power" Then
   Check2.Visible = True
End If

If RsTxn_CxTF_Details!HologramNo = "" Or IsNull(RsTxn_CxTF_Details!HologramNo) Then
   If PntPrv = "N" Then
      ans = MsgBox("We strongly recommend that you take a look at Print-Preview first before going ahead with this Print option. Do you still want to continue?", vbYesNo)
      If ans = vbNo Then
         Exit Sub
      End If
   End If
   Frame3.Visible = True
   TxtHologramNo.SetFocus
   Call TableMst_Hologram
   Exit Sub
Else
   TxtHologramNo = ""
   Frame3.Visible = False
End If
RsTxn_CxTF_Details!Flag = "P"
RsTxn_CxTF_Details.Update

Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
Call Xls_Print_Rpt
End Sub

Private Sub CmdPrintPreview_Click()

tmp = Format(Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(0, 5)), "#######0.000")

If Val(tmp) <> 0 Then
   MsgBox "Gross Weight in 'Truck Detail' Tab and 'GSL Detail' tab are not equal !!!!"
   Exit Sub
End If
If Val(MSHFlexGrid2.TextMatrix(0, 5)) <> Val(MSHFlexGrid4.TextMatrix(0, 4)) Then
   MsgBox "No of Bags in 'Client Detail' Tab and 'GSL Detail' tab are not equal !!!!"
   Exit Sub
End If
If Val(MSHFlexGrid2.TextMatrix(0, 6)) <> Val(MSHFlexGrid4.TextMatrix(0, 10)) Then
   MsgBox "Electronic Weight in 'Client Detail' Tab and 'GSL Detail' tab are not equal !!!!"
   Exit Sub
End If

Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
Call Xls_Print_Rpt
PntPrv = "Y"
End Sub

Private Sub CmdSearch_Click()

Dim Wc As String
Call Grid_Head

If STo.Value < SFrom.Value Then
   MsgBox "To date must be always greater than from date!!!"
   STo.SetFocus
   Exit Sub
End If

If CmbSCmp.Text <> "" Then
   Wc = GenWc(Wc, mCMPName, "Mst_CMP.CMPID", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp = Replace(Gen_WcCMP, "CMPID", "mst_CMP.CMPID")
      Wc = GenWc(Wc, tmp, "", "I")
   End If
End If

If CmbSCommodity.Text <> "" Then
   Wc = GenWc(Wc, mCommodityID, "Mst_Commodity.CommodityID", "N")
End If
If CmbSExchangeType.Text <> "" Then
   Wc = GenWc(Wc, mExchangeTypeID, "Mst_ExchangeType.ExchangeTypeID", "N")
End If
If CmbPreFix.Text <> "" Then
   Wc = GenWc(Wc, msSM_prefix, "Txn_CxTf_Details.sm_prefix", "S")
End If

tmp = "Select Txn_CxTf_Details.CxTFNo,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,Txn_CxTf_Details.SM_Prefix,"
tmp = tmp & " Mst_Commodity.Commodity,Txn_CxTf_Details.Flag,"
tmp = tmp & " Txn_CxTf_Details.Agent,Txn_CxTf_Details.Remark,Txn_CxTf_Details.HologramNo,"
tmp = tmp & " Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo, "
tmp = tmp & " Sum (Txn_CxTf_Gsl.NoofBags) 'NoOfBags',sum(Txn_CxTf_Gsl.GrossWeight) 'GrossWeight',Txn_CxTf_Details.G_UID "
tmp = tmp & " From Txn_CxTf_Details"
tmp = tmp & " Inner Join Txn_CxTf_Gsl  on Txn_CxTf_Details.SM_Prefix = Txn_CxTf_Gsl.SM_Prefix And"
tmp = tmp & " Txn_CxTf_Details.CxTFNo = Txn_CxTf_Gsl.CxTFNo And Txn_CxTf_Details.TxnType = Txn_CxTf_Gsl.TxnType"
tmp = tmp & " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID"
tmp = tmp & " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
tmp = tmp & " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID"
tmp = tmp & " Where Txn_CxTf_Details.TxnType = 'D' "

If Gen_DBType = "MDB" Then
   tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
    tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If Wc <> "" Then
   tmp = tmp & Wc
End If
tmp = tmp & " group by Txn_CxTf_Details.CxTFNo,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,"
tmp = tmp & " Txn_CxTf_Details.SM_Prefix, Mst_Commodity.Commodity,Txn_CxTf_Details.Flag, Txn_CxTf_Details.Agent,"
tmp = tmp & " Txn_CxTf_Details.Remark,Txn_CxTf_Details.HologramNo, Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo,"
tmp = tmp & " Txn_CxTf_Details.G_UID"
tmp = tmp & " ORDER BY Txn_CxTf_Details.sm_prefix,Mst_CMP.CMPName,Txn_CxTf_Details.CxTFNo "

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
For i = 1 To TmpRs.RecordCount
    For c = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    TmpRs.MoveNext
Next

RsTxn_CxTF_Details.Close
Wc = ""
If Gen_WcCMP <> "" Then
   Wc = "And " & Gen_WcCMP
End If

Call TableTxn_CxTF_Details(" Where TxnType = 'D' " & Wc & "Order By SM_Prefix,CMPID,CxTFNo ")

End Sub
Function GenWc(Wc_ As String, ID As Variant, Field_ As String, Type_ As String) As String
Dim RetVal
RetVal = Wc_
If RetVal = "" Then
   If Type_ = "S" Then
      RetVal = "And " & Field_ & "= '" & ID & "'"
   ElseIf Type_ = "I" Then
      RetVal = "And " & Field_ & " " & ID
   Else
      RetVal = "And " & Field_ & "= " & ID
   End If
Else
   If Type_ = "S" Then
      RetVal = RetVal & " And " & Field_ & "= '" & ID & "'"
   ElseIf Type_ = "I" Then
      RetVal = RetVal & " And " & Field_ & " " & ID
   Else
      RetVal = RetVal & " And " & Field_ & "= " & ID
   End If
End If
GenWc = RetVal
End Function

Private Sub TxtHologramNo_LostFocus()
If TxtHologramNo = "" Then Exit Sub
If Check_HologramNo(TxtHologramNo) = False Then
   TxtHologramNo.SetFocus
   CmdSaveHologram.Enabled = False
   Exit Sub
End If
CmdSaveHologram.Enabled = True
CmdSaveHologram.SetFocus
End Sub
Private Sub CmdSaveHologram_Click()
Dim ans As Variant
If TxtHologramNo = "" Then
   MsgBox "Blank Hologram No not allowed!!!"
   TxtHologramNo.SetFocus
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to save this HologramNo", vbYesNo)
If ans = vbYes Then
   Call TableMst_Hologram
   RsMst_Hologram.AddNew
   RsMst_Hologram!HologramNo = TxtHologramNo
   RsMst_Hologram!DocType = "CDTF"
   RsMst_Hologram!DocNo = TxtCxTFNo
   RsMst_Hologram!SM_Prefix = TxtPreFix
   RsMst_Hologram.Update
   RsTxn_CxTF_Details!HologramNo = TxtHologramNo
   RsTxn_CxTF_Details.Update
   TxtHologramNo = ""
   Frame3.Visible = False
   RsTxn_CxTF_Details!Flag = "P"
   RsTxn_CxTF_Details.Update
   Gen_mTimeStamp = GetTimeStamp
   If Check2.Value = 1 Then
      Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
      Call Xls_Print_Rpt
   End If
   RsMst_Hologram.Close
Else
   TxtHologramNo.SetFocus
End If

End Sub
Public Sub Create_Xls(FileName As String)
'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim Pat As String
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Gen_FLName = FileName
oBook.SaveAs (Pat & "\excel\" & FileName)

Set oSheet = oBook.Worksheets(1)
With oExcel
'.Application.Sheets.Delete
.SheetsInNewWorkbook = 15 'MSHFlexGrid2.Rows - 2
.Workbooks.Add
End With
'oBook.SaveAs (Pat & "\" & FileName)

oExcel.Quit
'oExcel.Close
End Sub
Public Sub Xls_Print_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Integer
'Dim mCol As Integer
Dim ans As Variant
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
'Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
 
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")

i = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
'oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 0.05
'oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 4.6
oWS.Name = TxtPreFix & "-" & TxtCxTFNo  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 4

oWS.Range("a1:g1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 12.86
oWS.Range("c:c").ColumnWidth = 12.86
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 12.86 '32.14
oWS.Range("f:f").ColumnWidth = 12.86
oWS.Range("g:g").ColumnWidth = 12.86

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = GetCMPAddress(RsTxn_CxTF_Details!CMPID, "D") '& "-" & GetCMPName(RsTxn_CxTF_Details!CMPID) ' & "-" & GetCMPAddress(RsTxn_CxTF_Details!CMPID)
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10
'oWS.Cells(mRow, c + 1).RowHeight = 24
mRow = mRow + 1


tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Commodity Deposit Transaction Form (CDTF)"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10
oWS.Cells(mRow, c + 1).RowHeight = 24

mRow = mRow + 1

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "(This is NOT an instrument of title to the goods)"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7


mRow = mRow + 1
oWS.Cells(mRow, c + 1) = "CDTF No"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
'oWS.Range("b5:c5").Merge
tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = TxtPreFix & "-" & TxtCxTFNo
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 7


oWS.Cells(mRow, c + 5) = "Date"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True

'oWS.Range(tmp).Merge
tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 6) = Format(TxtCxTFDate, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Location"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
tmp = "b" & mRow & ":" & "d" & mRow

oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = GetLocationFromCMP(RsTxn_CxTF_Details!CMPID)
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 7

oWS.Cells(mRow, c + 5) = "State"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True

tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 6) = GetStateFromLocation(GetLocationIDFromCMP(RsTxn_CxTF_Details!CMPID))
oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "CMP ID & Address"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).RowHeight = 40

tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = RsTxn_CxTF_Details!CMPID & "-" & GetCMPAddress(RsTxn_CxTF_Details!CMPID, "A")
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 2).WrapText = True


oWS.Cells(mRow, c + 5) = "Commodity"
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).Font.Size = 7

tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 6) = TxtCommodityID
oWS.Cells(mRow, c + 6).Font.Size = 7


mRow = mRow + 1


oWS.Cells(mRow, c + 1) = "Transaction Type"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).WrapText = True

tmp = "b" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtExchangeTypeID
oWS.Cells(mRow, c + 2).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Name Of Agent"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtAgent
oWS.Cells(mRow, c + 2).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Depositor Details"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Client-ID"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 2) = "Client Name"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).Font.Size = 7

If TxtExchangeTypeID = "NCDEX" Then
   oWS.Cells(mRow, c + 4) = "DP ID"
   oWS.Cells(mRow, c + 5) = "DP Name"
ElseIf TxtExchangeTypeID = "Non NCDEX" Then
   tmp = "b" & mRow & ":" & "e" & mRow
   oWS.Range(tmp).Merge
ElseIf TxtExchangeTypeID = "SPOT" Then
   oWS.Cells(mRow, c + 4) = "TM ID"
   oWS.Cells(mRow, c + 5) = "TM Name"
End If
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).Font.Size = 7

oWS.Cells(mRow, c + 6) = "No of Bags"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).Font.Size = 7
oWS.Cells(mRow, c + 7) = "Electronic Weight in MT"
oWS.Cells(mRow, c + 7).Font.Bold = True
oWS.Cells(mRow, c + 7).Font.Size = 7
oWS.Cells(mRow, c + 7).RowHeight = 24

mRow = mRow + 1

For k = 2 To MSHFlexGrid2.Rows - 2
'k = 1
    oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(k, 1)
    oWS.Cells(mRow, c + 1).Font.Size = 7
    tmp = "b" & mRow & ":" & "c" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 2) = MSHFlexGrid2.TextMatrix(k, 2)
    oWS.Cells(mRow, c + 2).Font.Size = 7
    If TxtExchangeTypeID <> "Non NCDEX" Then
       oWS.Cells(mRow, c + 4) = MSHFlexGrid2.TextMatrix(k, 3)
       oWS.Cells(mRow, c + 4).Font.Size = 7
       oWS.Cells(mRow, c + 5) = MSHFlexGrid2.TextMatrix(k, 4)
       oWS.Cells(mRow, c + 5).Font.Size = 7
    Else
       tmp = "b" & mRow & ":" & "e" & mRow
       oWS.Range(tmp).Merge
    End If
    oWS.Cells(mRow, c + 6) = MSHFlexGrid2.TextMatrix(k, 5)
    oWS.Cells(mRow, c + 6).Font.Size = 7
    oWS.Cells(mRow, c + 7) = Format(MSHFlexGrid2.TextMatrix(k, 6), "#######0.000")
    'oWS.Cells(mRow, c + 7).FormatCell = "##0.000"
    
    oWS.Cells(mRow, c + 7).Font.Size = 7
    mRow = mRow + 1
Next k

If TxtExchangeTypeID <> "Non NCDEX" Then
   tmp = "b" & mRow & ":" & "c" & mRow
Else
   tmp = "b" & mRow & ":" & "e" & mRow
End If
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Total"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108

oWS.Cells(mRow, c + 6) = MSHFlexGrid2.TextMatrix(0, 5)
oWS.Cells(mRow, c + 6).Font.Size = 7
oWS.Cells(mRow, c + 7) = MSHFlexGrid2.TextMatrix(0, 6)
oWS.Cells(mRow, c + 7).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Truck-wise Deposit Details"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Truck No"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = "Total Weight(MT)"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 3) = "Tare Weight(MT)"
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 4) = "Gross Weight"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 5) = "Weighbridge Slip No"
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).RowHeight = 25
tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 6) = "Remarks"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

For k = 2 To MSHFlexGrid3.Rows - 2
'k = 1
    oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(k, 1)
    oWS.Cells(mRow, c + 1).Font.Size = 7
    oWS.Cells(mRow, c + 2) = Format(MSHFlexGrid3.TextMatrix(k, 3), ".000")
    oWS.Cells(mRow, c + 2) = Format(MSHFlexGrid3.TextMatrix(k, 3), "#######0.000")
    oWS.Cells(mRow, c + 2).Font.Size = 7
    oWS.Cells(mRow, c + 3) = Format(MSHFlexGrid3.TextMatrix(k, 4), "#######0.000")
    oWS.Cells(mRow, c + 3).Font.Size = 7
    oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid3.TextMatrix(k, 5), "#######0.000")
    oWS.Cells(mRow, c + 4).Font.Size = 7
    oWS.Cells(mRow, c + 5) = MSHFlexGrid3.TextMatrix(k, 2)
    oWS.Cells(mRow, c + 5).Font.Size = 7
    tmp = "f" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 6) = MSHFlexGrid3.TextMatrix(k, 6)
    oWS.Cells(mRow, c + 6).Font.Size = 7
    mRow = mRow + 1
Next k

tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

mRow = mRow + 1



tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Total"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108

oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid3.TextMatrix(0, 5), "#######0.000")
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge


oWS.Cells(mRow, c + 5) = "Affix Hologram Here"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge

mRow = mRow + 1
'mCol = 5

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Less Gunny Weight(MT)"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108

oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 8), "#######0.000")
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Less Standard Deductions(MT)"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108

oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 9), "#######0.000")
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge
mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Add Adjustments(MT)"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108

oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 6), "#######0.000")
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge


mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Electronic Weight(MT)"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108

oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 10), "#######0.000")

oWS.Cells(mRow, c + 4).Font.Size = 7


oWS.Cells(mRow, c + 5) = "Hologram No"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True

tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
If RsTxn_CxTF_Details!HologramNo = "" Or IsNull(RsTxn_CxTF_Details!HologramNo) Then
   oWS.Cells(mRow, c + 6) = "Print Preview !!!! "
Else
   oWS.Cells(mRow, c + 6) = RsTxn_CxTF_Details!HologramNo
End If
oWS.Cells(mRow, c + 6).Font.Size = 7
oWS.Cells(mRow, c + 6).HorizontalAlignment = -4108

mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Name & Signature of CMP Manager/CMP DRC:"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).RowHeight = 40
mRow = mRow + 1

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

Dim SDesc As String

SDesc = "I/We hereby declare that the commodities being deposited herein belong " & _
        "to the clients as stated above and are free from any encumbrances. " & _
        "I/we request you to take custody of the said commodities and I/we undertake to " & _
        "indemnify you against any losses incurred by you due to any false information provided by me/us." & _
        "I/we also undertake to pay all required charges as per your charges schedule from time to time. " & _
        "I have gone through the attached 'Terms and Conditions'** of storage and agree to abide by the same. "
If mExchangeTypeID = 2 Then
   SDesc = SDesc & " I/We hereby request the warehouse to credit the above commodities to the account of the beneficiaries " & _
           "stated above after receiving the assayer's certificates for the above commodities  and after deducting " & _
           "standard deductions and  gunny bags weight as per Commodity Exchange norm."
           ' Added on 22.03.2010
   SDesc = SDesc & " We understand that NCMSL may not enter the deposit details in NCDEX system for the purpose of demat unless all charges, including pre-demat storage charges and assaying charges are paid in full"
           
End If
oWS.Cells(mRow, c + 1) = SDesc
oWS.Cells(mRow, c + 1).Font.Size = 6
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).RowHeight = 50

mRow = mRow + 1
mRow = mRow + 1

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "** 'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html "
oWS.Cells(mRow, c + 1).Font.Size = 8
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).RowHeight = 24
oWS.Cells(mRow, c + 1).Font.Name = "Zurich BT"

mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1



tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Name & Signature of Depositor/Agent of Depositor:"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Name = "Zurich BT"

k = mRow - 6

For Cell = 4 To k
    tmp = "a" & Cell & ":" & "g" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell
   
If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If




' --------------------- GSL Detail


Set oWS = oWB.Worksheets("Sheet" & 2)
'oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 0.05
'oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 4.6
oWS.Name = TxtPreFix & "-" & TxtCxTFNo & " - GCTD Detail"  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 1

oWS.Range("a1:g1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 12.86
oWS.Range("c:c").ColumnWidth = 12.86
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 12.86 '32.14
oWS.Range("f:f").ColumnWidth = 12.86
oWS.Range("g:g").ColumnWidth = 12.86

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = GetCMPAddress(RsTxn_CxTF_Details!CMPID, "D") '& "-" & GetCMPName(RsTxn_CxTF_Details!CMPID) ' & "-" & GetCMPAddress(RsTxn_CxTF_Details!CMPID)
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10
'oWS.Cells(mRow, c + 1).RowHeight = 24
mRow = mRow + 1


tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Godown Stack Lot Detail "
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10
oWS.Cells(mRow, c + 1).RowHeight = 24

'mRow = mRow + 1
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "(This is NOT an instrument of title to the goods)"
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 7


mRow = mRow + 1
oWS.Cells(mRow, c + 1) = "CDTF No"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
'oWS.Range("b5:c5").Merge
tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = TxtPreFix & "-" & TxtCxTFNo
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 7


oWS.Cells(mRow, c + 5) = "Date"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True

'oWS.Range(tmp).Merge
tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 6) = Format(TxtCxTFDate, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1
oWS.Cells(mRow, c + 1) = "Godown No"
oWS.Cells(mRow, c + 1).Font.Size = 7: oWS.Cells(mRow, c + 1).Font.Bold = True

oWS.Cells(mRow, c + 2) = "Godown Address"
oWS.Cells(mRow, c + 2).Font.Size = 7: oWS.Cells(mRow, c + 2).Font.Bold = True

oWS.Cells(mRow, c + 3) = "Stack No "
oWS.Cells(mRow, c + 3).Font.Size = 7: oWS.Cells(mRow, c + 3).Font.Bold = True

oWS.Cells(mRow, c + 4) = "Lot No"
oWS.Cells(mRow, c + 4).Font.Size = 7: oWS.Cells(mRow, c + 4).Font.Bold = True

oWS.Cells(mRow, c + 5) = "No of Bags "
oWS.Cells(mRow, c + 5).Font.Size = 7: oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 6) = "Gross Weight"
oWS.Cells(mRow, c + 6).Font.Size = 7: oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 7) = "client Name"
oWS.Cells(mRow, c + 7).Font.Size = 7: oWS.Cells(mRow, c + 7).Font.Bold = True

mRow = mRow + 1

For k = 2 To MSHFlexGrid4.Rows - 2
'k = 1
    oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid4.TextMatrix(k, 1)
    oWS.Cells(mRow, c + 1).Font.Size = 7
    oWS.Cells(mRow, c + 2) = "'" & MSHFlexGrid4.TextMatrix(k, 18)
    oWS.Cells(mRow, c + 2).Font.Size = 7
    oWS.Cells(mRow, c + 3) = "'" & MSHFlexGrid4.TextMatrix(k, 2)
    oWS.Cells(mRow, c + 3).Font.Size = 7
    oWS.Cells(mRow, c + 4) = (MSHFlexGrid4.TextMatrix(k, 3))
    oWS.Cells(mRow, c + 4).Font.Size = 7
    oWS.Cells(mRow, c + 5) = MSHFlexGrid4.TextMatrix(k, 4)
    oWS.Cells(mRow, c + 5).Font.Size = 7
'    tmp = "f" & mRow & ":" & "g" & mRow
'    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 6) = MSHFlexGrid4.TextMatrix(k, 5)
    oWS.Cells(mRow, c + 6).Font.Size = 7
    oWS.Cells(mRow, c + 7) = MSHFlexGrid4.TextMatrix(k, 11)
    oWS.Cells(mRow, c + 7).Font.Size = 7
    
    mRow = mRow + 1
Next k
k = mRow

For Cell = 1 To k
    tmp = "a" & Cell & ":" & "g" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell

If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
'-------- End GSL Detail

oWB.Save

Call Log_Print_History(TxtPreFix, "CDTF", TxtCxTFNo, Gen_FLName)
    
    

'oWB.Close
'oXL.Quit
MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub




Private Sub CmdPrintold_Click()
On Error Resume Next

RsTxn_CxTF_Details!Flag = "P"
RsTxn_CxTF_Details.Update
tmp = "Delete from PrintDetail "
Call TmpTable

tmp = "Delete from PrintMain "
Call TmpTable
tmp = "select * from PrintMain"
Call TmpTable
TmpRs.AddNew
TmpRs!ID = TxtCxTFNo
TmpRs!Type = "CDTF"
TmpRs!F1 = TxtCxTFDate
TmpRs!F2 = TxtExchangeTypeID
TmpRs!F3 = TxtCMPID
TmpRs!F4 = TxtAgent
TmpRs!F5 = TxtExchangeTypeID
TmpRs!F6 = TxtCommodityID
TmpRs!F7 = TxtExchangeTypeID

TmpRs!F10 = TxtExchangeTypeID
TmpRs!F11 = TxtExchangeTypeID
TmpRs!F12 = TxtExchangeTypeID
TmpRs!F13 = TxtExchangeTypeID


TmpRs.Update
tmp = "select * from PrintDetail "
Call TmpTable


For i = 2 To MSHFlexGrid3.Rows - 1
    TmpRs.AddNew
    TmpRs!ID = TxtCxTFNo
    TmpRs!F1 = MSHFlexGrid3.TextMatrix(i, 1)
    TmpRs!F2 = MSHFlexGrid3.TextMatrix(i, 3)
    TmpRs!F3 = MSHFlexGrid3.TextMatrix(i, 4)
    TmpRs!F4 = MSHFlexGrid3.TextMatrix(i, 5)
    TmpRs!F5 = MSHFlexGrid3.TextMatrix(i, 2)
    TmpRs!F6 = MSHFlexGrid3.TextMatrix(i, 6)
    TmpRs.Update
Next
TmpRs.Close
cn.Close
If DataEnvironment1.rsCommand1.State <> 0 Then
  DataEnvironment1.rsCommand1.Close
End If

Rpt_CDTF.Refresh
Rpt_CDTF.Width = 15200
Rpt_CDTF.Show
End Sub

Private Sub CmdTTDDeleteList_Click()
i = MSHFlexGrid3.RowSel
If i > 1 Then
   If MSHFlexGrid3.TextMatrix(i, 1) <> "" Then
      MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) - Val(MSHFlexGrid3.TextMatrix(i, 3))
      MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) - Val(MSHFlexGrid3.TextMatrix(i, 4))
      MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid3.TextMatrix(i, 5))
      For RO = i To MSHFlexGrid3.Rows - 2
          For c = 0 To MSHFlexGrid3.Cols - 1
               MSHFlexGrid3.TextMatrix(RO, c) = MSHFlexGrid3.TextMatrix(RO + 1, c)
          Next
      Next
      MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
   End If
End If
Call ClearTTDTabControl
End Sub

Private Sub CmdTTDAdd2List_Click()
Dim r As Variant
If TxtTTDTruckNo = "" Then
   MsgBox "Blank Truck No !!! "
   TxtTTDTruckNo.SetFocus
   Exit Sub
End If

If Val(TxtTTDTotalWt) <= 0 Then
   MsgBox "Invalid Weight !!! "
   TxtTTDTotalWt.SetFocus
   Exit Sub
End If

If Val(TxtTTDGrossWt) <= 0 Then
   MsgBox "Invalid Weight !!! "
   TxtTTDGrossWt.SetFocus
   Exit Sub
End If
'tmp = Format((Val(TxtTTDTotalWt) - Val(TxtTTDTareWt)) - Val(TxtTTDGrossWt), "#####0.000")
'If Val(tmp) <> 0 Then
'   MsgBox "Invalid Weight !!! "
'   Exit Sub
'End If

If CmdTTDAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid3.RowSel
   MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) - Val(MSHFlexGrid3.TextMatrix(r, 3))
   MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) - Val(MSHFlexGrid3.TextMatrix(r, 4))
   MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid3.TextMatrix(r, 5))
Else
  r = MSHFlexGrid3.Rows - 1
  MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If
MSHFlexGrid3.TextMatrix(r, 1) = TxtTTDTruckNo
MSHFlexGrid3.TextMatrix(r, 2) = TxtWeighBridgNo
MSHFlexGrid3.TextMatrix(r, 3) = TxtTTDTotalWt
MSHFlexGrid3.TextMatrix(r, 4) = TxtTTDTareWt
MSHFlexGrid3.TextMatrix(r, 5) = TxtTTDGrossWt
MSHFlexGrid3.TextMatrix(r, 6) = txtTTDRemarks

MSHFlexGrid3.Col = 5
MSHFlexGrid3.Row = r
MSHFlexGrid3.CellBackColor = &HFFC0C0


MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(r, 3))
MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(r, 4))
MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + Val(MSHFlexGrid3.TextMatrix(r, 5))
Call ClearTTDTabControl
TxtTTDTruckNo.SetFocus
End Sub
Sub ClearTTDTabControl()
TxtTTDTruckNo = ""
TxtWeighBridgNo = ""
TxtTTDTotalWt = ""
TxtTTDTareWt = ""
TxtTTDGrossWt = ""
txtTTDRemarks = ""
CmdTTDAdd2List.Caption = "Add to List"
CmdTTDDeleteList.Enabled = False
TxtTTDGrossWt.Locked = True
End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Wc = ""
If Gen_WcCMP <> "" Then
   Wc = "And " & Gen_WcCMP
End If

Call TableTxn_CxTF_Details(" Where TxnType = 'D' " & Wc & " Order By SM_Prefix,CMPID,CxTFNo ")
If RsTxn_CxTF_Details.RecordCount = 0 Then
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
Call InData

End Sub
Private Sub MSHFlexGrid4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyC And Shift = vbCtrlMask Then
   With MSHFlexGrid4
        .Redraw = False              'Find range to copy
        min_row = .Row:   max_row = .RowSel
        If min_row > max_row Then tmp = max_row: max_row = min_row: min_row = tmp
        min_col = .Col:   max_col = .ColSel
        If min_col > max_col Then tmp = max_col: max_col = min_col: min_col = tmp

        my_text = ""                   'Get the text
        For i = min_row To max_row
            For j = min_col To max_col
                my_text = my_text & .TextMatrix(i, j)
                If j < max_col Then my_text = my_text & vbTab
            Next j
            If i < max_row Then my_text = my_text & vbCrLf
        Next i
        .Redraw = True
   End With
   If my_text <> "" Then
      Clipboard.Clear
      Clipboard.SetText (my_text)
   Else
      Beep
      MsgBox "Please select some text.", vbExclamation, "Nothing selected!"
   End If
End If
End Sub
Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Txn_CxTF_Details.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\Txn_CxTF_Details.xls")
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

Private Sub MSHFlexGrid2_Click()
    
If SaveCmd.Enabled = False Then Exit Sub
    
If MSHFlexGrid2.Rows = 3 Then Exit Sub
r = MSHFlexGrid2.RowSel
If r <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(r, 1) = "" Then
   Exit Sub
End If
CmdClientAdd2List.Caption = "Update List"
TxtClientID.Locked = False


TxtClientID = MSHFlexGrid2.TextMatrix(r, 1)
TxtClientName = MSHFlexGrid2.TextMatrix(r, 2)
TxtDPID = MSHFlexGrid2.TextMatrix(r, 3)
TxtDPName = MSHFlexGrid2.TextMatrix(r, 4)
TxtClientNoBags = MSHFlexGrid2.TextMatrix(r, 5)
TxtClientWeight = MSHFlexGrid2.TextMatrix(r, 6)

If CmbNonNcdexClient.Visible = True Then
   CmbNonNcdexClient.Text = MSHFlexGrid2.TextMatrix(r, 2)
   CmbNonNcdexClient.SetFocus
Else
   If TxtClientID.Enabled = True Then
      TxtClientID.SetFocus
   End If
End If

'    MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) - Val(MSHFlexGrid2.TextMatrix(r, 5))
'    MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(r, 6))

    
End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub

CmdTTDAdd2List.Caption = "Update List"
r = MSHFlexGrid3.RowSel
If r <= 1 Then Exit Sub
TxtTTDTruckNo = MSHFlexGrid3.TextMatrix(r, 1)
TxtWeighBridgNo = MSHFlexGrid3.TextMatrix(r, 2)
TxtTTDTotalWt = MSHFlexGrid3.TextMatrix(r, 3)
TxtTTDTareWt = MSHFlexGrid3.TextMatrix(r, 4)
TxtTTDGrossWt = MSHFlexGrid3.TextMatrix(r, 5)
txtTTDRemarks = MSHFlexGrid3.TextMatrix(r, 6)
CmdTTDDeleteList.Enabled = True
End Sub

Private Sub MSHFlexGrid4_Click()
     
Call ClearGCTDTabControl
If SaveCmd.Enabled = False Then Exit Sub ' Code Added on 16 jan 10 for View mode navigation '

r = MSHFlexGrid4.RowSel
If r <= 1 Then Exit Sub
If MSHFlexGrid4.TextMatrix(r, 14) <> "" Then
   MsgBox "This godown is Closed you can't Edit !!!!!"
   Exit Sub
End If
If MSHFlexGrid4.TextMatrix(r, 17) <> "" Then
   MsgBox "This ISIN is Mapped against this GSL you can't Edit !!!!!"
   Exit Sub
End If
If MSHFlexGrid4.TextMatrix(r, 1) = "" Then
   Call ClearGCTDTabControl
   Exit Sub
End If

CmdGCTDAdd2List.Caption = "Update List"


CmbGodownNo.Text = MSHFlexGrid4.TextMatrix(r, 1)
CmbStack.Text = MSHFlexGrid4.TextMatrix(r, 2)
CmbLot.Text = MSHFlexGrid4.TextMatrix(r, 3)
TxtGCTDNoofBags = MSHFlexGrid4.TextMatrix(r, 4)
TxtGCTDGrossWt = MSHFlexGrid4.TextMatrix(r, 5)

TxtAdjWt = MSHFlexGrid4.TextMatrix(r, 6)
TxtAdjGrossWt = MSHFlexGrid4.TextMatrix(r, 7)

TxtGCTDGunnyWt = MSHFlexGrid4.TextMatrix(r, 8)
TxtGCTDStdDeduction = MSHFlexGrid4.TextMatrix(r, 9)
TxtGCTDDematWt = MSHFlexGrid4.TextMatrix(r, 10)

CmbGCTDClient.Text = MSHFlexGrid4.TextMatrix(r, 11)
TxtGCTDremarks = MSHFlexGrid4.TextMatrix(r, 12)
mGCTDClientID = MSHFlexGrid4.TextMatrix(r, 16)

mGSLID = MSHFlexGrid4.TextMatrix(r, 13)
mGodownID = Val(MSHFlexGrid4.TextMatrix(r, 15))

Call CmbGodownNo_GotFocus
Call CmbGodownNo_LostFocus
CmdGCTDeleteList.Enabled = True

If MSHFlexGrid4.TextMatrix(r, 0) <> "" Then
   CmbGodownNo.Enabled = False
   CmbStack.Enabled = False
   CmbLot.Enabled = False
Else
   CmbGodownNo.Enabled = True
   CmbStack.Enabled = True
   CmbLot.Enabled = True
End If
     
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_Details.MoveNext
If RsTxn_CxTF_Details.EOF Then
   RsTxn_CxTF_Details.MoveLast
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
RsTxn_CxTF_Details.MovePrevious
If RsTxn_CxTF_Details.BOF Then
   RsTxn_CxTF_Details.MoveFirst
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
RsTxn_CxTF_Details.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_Details.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
If SaveCmd.Enabled = True Then
   ans = MsgBox("Do you really want to exit?", vbYesNo)
   If ans = vbNo Then
      Exit Sub
   End If
End If
Unload Me
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CxTF_Details.RecordCount = 0 Then Exit Sub
   ans = MsgBox("Do you really want to Cancel entered details?", vbYesNo)
   If ans = vbNo Then
      Exit Sub
   End If
   Call GButtonLock(Me, True)
   Call InData
   Exit Sub
End If

Frame0.Height = Me.Height - 550
Frame0.Left = Picture1.Left - 120
Frame0.Top = Picture1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1600
RsTxn_CxTF_Details.MoveFirst
Call Grid_Head


CmbSCmp.Text = ""
CmbSCommodity.Text = ""
CmbSExchangeType.Text = ""
CmbSAgent.Text = ""
SFrom.Value = Date - 730
STo.Value = Date

'MSHFlexGrid1.Rows = RsTxn_CxTF_Details.RecordCount + 1
'MSHFlexGrid1.ColWidth(0) = 1000
'For i = 1 To RsTxn_CxTF_Details.RecordCount
'    MSHFlexGrid1.TextMatrix(i, 0) = RsTxn_CxTF_Details!CxTFNo
'    MSHFlexGrid1.TextMatrix(i, 1) = RsTxn_CxTF_Details!CxTFDate
'    MSHFlexGrid1.TextMatrix(i, 2) = GetCMPName(RsTxn_CxTF_Details!CMPID)
'    MSHFlexGrid1.TextMatrix(i, 3) = GetCommodityName(RsTxn_CxTF_Details!CommodityID)
'    MSHFlexGrid1.TextMatrix(i, 4) = GetExchangeName(RsTxn_CxTF_Details!ExchangeTypeID)
'    MSHFlexGrid1.TextMatrix(i, 5) = RsTxn_CxTF_Details!SM_Prefix
'    MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(RsTxn_CxTF_Details!Flag), "", RsTxn_CxTF_Details!Flag)
'    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(RsTxn_CxTF_Details!Agent), "", RsTxn_CxTF_Details!Agent)
'    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(RsTxn_CxTF_Details!Remark), "", RsTxn_CxTF_Details!Remark)
'    MSHFlexGrid1.TextMatrix(i, 9) = "Deposit"
'    MSHFlexGrid1.TextMatrix(i, 11) = IIf(IsNull(RsTxn_CxTF_Details!G_UID), "", RsTxn_CxTF_Details!G_UID)
'    RsTxn_CxTF_Details.MoveNext
'Next

End Sub
Public Sub InData()

Pic_Client.Enabled = False
Pic_GCTD.Enabled = False
Pic_TTD.Enabled = False
PntPrv = "N"
Call Grid_Head_TTD
Call Grid_Head_Client
Call Grid_Head_GCTD

TxtCxTFNo.Locked = True
TxtManualCDTFNo.Locked = True
TxtManualCDTFNo = ""

TxtCxTFDate.Enabled = False
TxtCMPID.Locked = True
CmbCMPID.Visible = False
TxtCMPID.Visible = True
TxtCommodityID.Locked = True
CmbCommodityID.Visible = False
TxtCommodityID.Visible = True
TxtExchangeTypeID.Locked = True
CmbExchangeTypeID.Visible = False
TxtExchangeTypeID.Visible = True

TxtAgent.Locked = True
CmbAgent.Visible = False
CmdAgent.Visible = False
TxtAgent.Visible = True

CmdClient.Visible = False
CmdCMPID.Visible = False
CmdCommodityID.Visible = False
CmdContractID.Visible = False
Picture2.Enabled = False
'TxtBookedSpaceFlag.Locked = True
'TxtFlag.Locked = True
TxtAgent.Locked = True
'TxtRemark.Locked = True
'TxtTxnType.Locked = True
TxtContractID.Locked = True
CmbContractID.Visible = False
TxtContractID.Visible = True
TxtContractID = ""
TxtCxTFNo = IIf(IsNull(RsTxn_CxTF_Details!CxTFNo), "", RsTxn_CxTF_Details!CxTFNo)
TxtPreFix = IIf(IsNull(RsTxn_CxTF_Details!SM_Prefix), "", RsTxn_CxTF_Details!SM_Prefix)
TxtManualCDTFNo = IIf(IsNull(RsTxn_CxTF_Details!ManualCDTFNo), "", RsTxn_CxTF_Details!ManualCDTFNo)

TxtCxTFDate = IIf(IsNull(RsTxn_CxTF_Details!CxTFDate), "", RsTxn_CxTF_Details!CxTFDate)
TxtCMPID = GetCMPName(IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID))
mCMPID = IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID)
TxtCommodityID = GetCommodityName(IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID))
mCommodityID = IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID)
TxtExchangeTypeID = GetExchangeName(IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeId), "", RsTxn_CxTF_Details!ExchangeTypeId))
mExchangeTypeID = IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeId), "", RsTxn_CxTF_Details!ExchangeTypeId)
TxtBookedSpaceFlag = IIf(IsNull(RsTxn_CxTF_Details!BookedSpaceFlag), "", RsTxn_CxTF_Details!BookedSpaceFlag)
TxtFlag = IIf(IsNull(RsTxn_CxTF_Details!Flag), "", RsTxn_CxTF_Details!Flag)
TxtAgent = IIf(IsNull(RsTxn_CxTF_Details!Agent), "", RsTxn_CxTF_Details!Agent)


If Not IsNull(RsTxn_CxTF_Details!AgentID) Then
   mAgentID = IIf(IsNull(RsTxn_CxTF_Details!AgentID), "", RsTxn_CxTF_Details!AgentID)
   TxtAgent = GetAgentName(IIf(IsNull(RsTxn_CxTF_Details!AgentID), "", RsTxn_CxTF_Details!AgentID))
End If

TxtRemark = IIf(IsNull(RsTxn_CxTF_Details!Remark), "", RsTxn_CxTF_Details!Remark)
TxtTxnType = IIf(IsNull(RsTxn_CxTF_Details!TxnType), "", RsTxn_CxTF_Details!TxnType)
Check1.Value = 0
If Not IsNull(RsTxn_CxTF_Details!ContractID) Then
   TxtContractID = GetContract(IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID))
   Check1.Value = 1
End If
mContractID = IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID)
CmbCMPID.Text = TxtCMPID
CmbExchangeTypeID.Text = TxtExchangeTypeID



TxtCreated = IIf(IsNull(RsTxn_CxTF_Details!CreatedBy), "", RsTxn_CxTF_Details!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CxTF_Details!CreatedDate), "", RsTxn_CxTF_Details!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_CxTF_Details!UpdatedBy), "", RsTxn_CxTF_Details!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CxTF_Details!UpdatedDate), "", RsTxn_CxTF_Details!UpdatedDate)

TxtG_UID = IIf(IsNull(RsTxn_CxTF_Details!G_UID), "", RsTxn_CxTF_Details!G_UID)

Call CmbExchangeTypeID_GotFocus
Call CmbExchangeTypeID_LostFocus
Call CmbCMPID_GotFocus
Call CmbCMPID_LostFocus
CmbCommodityID.Text = TxtCommodityID
Call CmbCommodityID_GotFocus
Call CmbCommodityID_LostFocus

Call InDataTruck
Call InDataClient
Call InDataGSL
If RsTxn_CxTF_Details.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
Call UpdateClientGridStock
TxtCxTFDate = IIf(IsNull(RsTxn_CxTF_Details!CxTFDate), "", RsTxn_CxTF_Details!CxTFDate)
TxtCxTFDispalyDate = IIf(IsNull(RsTxn_CxTF_Details!CxTFDate), "", RsTxn_CxTF_Details!CxTFDate)
Frame3.Visible = False

CmdPrint.Enabled = True
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
If RsTxn_CxTF_Details.RecordCount = 0 Then Exit Sub

If SSTab1.Tab = 0 Then
   If SaveCmd.Enabled = True Then
      TxtClientID.SetFocus
   End If
   Call UpdateClientGridStock

End If

If SaveCmd.Enabled = False Then
   Exit Sub
End If
If SSTab1.Tab = 1 Then
   TxtTTDTruckNo.SetFocus
End If
If SSTab1.Tab = 2 Then
   If CmbGodownNo.Enabled = True Then
      CmbGodownNo.SetFocus
   End If
End If

End Sub

Private Sub TxtAdjWt_LostFocus()
If TxtAdjWt = "" Then
   TxtAdjWt = Val(TxtAdjWt)
Else
    If IsNumeric(TxtAdjWt) = False Then
       MsgBox "Invalid numeric format !!!"
       TxtAdjWt.SetFocus
       Exit Sub
    End If
End If
Call CalDmatWt
End Sub

Private Sub TxtClientID_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 1
   TxtTTDTruckNo.SetFocus
End If
End Sub

Private Sub TxtTTDTruckNo_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 2
   CmbGodownNo.SetFocus
End If
End Sub
Private Sub CmbGodownNo_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 0
   If mExchangeTypeID <> 1 Then
      TxtClientID.SetFocus
   Else
       CmbNonNcdexClient.SetFocus
   End If
End If
End Sub

Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim ans As Variant
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtClientID.SetFocus
   Exit Sub
End If
If ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
   CmbClientDP.SetFocus
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

'      "WHERE (((Mst_Depository.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

Call Tmp1Table

For op = 1 To TmpRs1.RecordCount
    If CmbClientDP.Text = "" Then
       CmbClientDP.Text = TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    End If
    CmbClientDP.AddItem TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    TmpRs1.MoveNext
Next
If ans(3, 0) <= 1 Then
   CmbClientDP_LostFocus
End If
End Sub

Private Sub TxtCxTFNo_LostFocus()
If TxtCxTFNo = "" Then Exit Sub
If IsNumeric(TxtCxTFNo) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtCxTFNo.SetFocus
   Exit Sub
End If
TxtCxTFNo = Val(TxtCxTFNo)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
SSTab1.Tab = 0
Call GButtonLock(Me, False)
CmdPrint.Enabled = False
Call Grid_Head_TTD
Call Grid_Head_Client
Call Grid_Head_GCTD
Picture2.Enabled = True
SSTab1.Tab = 0
Pic_Client.Enabled = True
Pic_GCTD.Enabled = True
Pic_TTD.Enabled = True
TxtCxTFDispalyDate = ""
TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = True
Check1.Value = 0
Check1.Enabled = False
TxtCMPID.Locked = False
CmbCMPID.Visible = True
TxtCMPID.Visible = False
TxtCommodityID.Locked = False
CmbCommodityID.Visible = True
TxtCommodityID.Visible = False
TxtExchangeTypeID.Locked = False
CmbExchangeTypeID.Visible = True
TxtExchangeTypeID.Visible = False
TxtManualCDTFNo.Locked = False
TxtManualCDTFNo = ""

'TxtBookedSpaceFlag.Locked = False
'TxtFlag.Locked = False
TxtAgent.Locked = False
CmbAgent.Visible = True
CmdAgent.Visible = True
TxtAgent.Visible = False
CmbAgent.Text = ""

'TxtRemark.Locked = False
'TxtTxnType.Locked = False
CmbContractID.Text = ""

TxtPreFix = ""
CmdClient.Visible = True
CmdCMPID.Visible = True
CmdCommodityID.Visible = True
CmdContractID.Visible = True


TxtContractID.Locked = False
CmbContractID.Visible = True
TxtContractID.Visible = False
TxtClientID.Locked = False
TxtCxTFNo = ""
TxtCxTFDate = CDate("01/01/2001")
TxtCMPID = ""
TxtCommodityID = ""
TxtExchangeTypeID = ""
TxtBookedSpaceFlag = ""
TxtFlag = ""
TxtAgent = ""
TxtRemark = ""
TxtTxnType = ""
TxtContractID = ""
CmbCMPID.Text = ""
CmbExchangeTypeID.Text = ""
CmbCommodityID.Text = ""

ClearClientTabControl
ClearTTDTabControl
ClearGCTDTabControl

TxtCGunnyWt = ""
TxtCStdDed = ""
TxtCLotSize = ""
TxtCVariation = ""
TxtCreated = ""
TxtUpdated = ""
CmbContractID.Enabled = False
If RsTxn_CxTF_Details.RecordCount > 0 Then
   If CmbExchangeTypeID.Enabled = True Then
      CmbExchangeTypeID.SetFocus
   End If
End If

End Sub
Private Sub EditCmd_Click()
If RsTxn_CxTF_Details!Flag = "P" Then
   MsgBox "Print out is Generated, You Can't Edit This CDTF   !!!! "
   Exit Sub
End If

For i = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(i, 17) <> "" Then
       MsgBox "Some of lot is trasfer, You Can't Edit This CDTF!!! "
       Exit Sub
    End If
Next

Edit_Flg = "E"
Call GButtonLock(Me, False)

Pic_Client.Enabled = True
Pic_GCTD.Enabled = True
Pic_TTD.Enabled = True


TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = True
TxtCMPID.Locked = True
'CmbCMPID.Visible = True
'TxtCMPID.Visible = False
TxtCommodityID.Locked = True
'CmbCommodityID.Visible = True
'TxtCommodityID.Visible = False
TxtExchangeTypeID.Locked = True
'CmbExchangeTypeID.Visible = True
'TxtExchangeTypeID.Visible = False
'TxtBookedSpaceFlag.Locked = False
'TxtFlag.Locked = False
TxtAgent.Locked = False
'TxtRemark.Locked = False
'TxtTxnType.Locked = False
TxtContractID.Locked = False
CmbContractID.Visible = True
TxtContractID.Visible = False
TxtClientID.Locked = False
CmbCMPID.Text = TxtCMPID
TxtCxTFDate.Enabled = False

'If LCase(CmbCMPID.Text) = "unjha" Or LCase(CmbCMPID.Text) = "rajkot" Then
'   TxtCxTFDate.Enabled = True
'Else
'   TxtCxTFDate.Enabled = False
'End If
CmbCommodityID.Text = TxtCommodityID
CmbExchangeTypeID.Text = TxtExchangeTypeID
CmbContractID.Text = TxtContractID


TxtAgent.Locked = False
CmbAgent.Visible = True
CmdAgent.Visible = True
TxtAgent.Visible = False
CmbAgent.Text = TxtAgent


End Sub
Private Sub DeleteCmd_Click()
If RsTxn_CxTF_Details!Flag = "P" Then
   MsgBox "Print out is Generated, You Can't Edit This CDTF   !!!! "
   Exit Sub
End If
For i = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(i, 17) <> "" Then
       MsgBox "Some of lot is trasfer, You Can't Edit This CDTF!!! "
       Exit Sub
    End If
Next

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

   Call DeleteChildTableDetail
   tmp = "select * from Txn_CxTF_TruckDetail where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
   Call TmpTable
   TmpRs.Close
   
   RsTxn_CxTF_Details.Delete
   RsTxn_CxTF_Details.Update
   If RsTxn_CxTF_Details.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CxTF_Details.MoveNext
   If RsTxn_CxTF_Details.EOF() Then
      RsTxn_CxTF_Details.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_CxTF_Details.CancelUpdate
End Sub

Sub DeleteChildTableDetail()
Dim adjWt As Variant
Call TableMst_GSL("where SM_Prefix = '" & TxtPreFix & "'")
tmp = "select * from Txn_CxTF_GSL where TxnType = 'D' And CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
For i = 1 To TmpRs.RecordCount
    RsMst_GSL.MoveFirst
    RsMst_GSL.Find "GSLID = " & TmpRs!GSLID
    If Not RsMst_GSL.EOF Then
       RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - TmpRs!NoofBags
       RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - TmpRs!DematWeight
       RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - TmpRs!NoofBags
       RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - TmpRs!DematWeight
       RsMst_GSL!GrossWt = IIf(IsNull(RsMst_GSL!GrossWt), 0, RsMst_GSL!GrossWt) - TmpRs!GrossWeight
       If RsMst_GSL!DepositeWeight = 0 Then
          RsMst_GSL!Flag = "B"
          RsMst_GSL!ClientIDRow = 0
          RsMst_GSL!CommodityID = Null
          RsMst_GSL!ExchangeTypeId = Null
          RsMst_GSL!CdtfNo = Null
          RsMst_GSL!CDTFDate = Null
       Else
          RsMst_GSL!Flag = "O"
       End If
       RsMst_GSL.Update
    End If
    TmpRs.MoveNext
Next
tmp = "Delete from Txn_CxTF_ClientDetail where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
tmp = "Delete from Txn_CxTF_GSL where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
tmp = "Delete from Txn_CxTF_TruckDetail where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
tmp = "select * from Txn_CxTF_GSL where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
TmpRs.Close
End Sub

Private Sub SaveCmd_Click()
Call UpdateClientGridStock

If TxtCxTFDate.Value > Date Then
   MsgBox "Future Date is not Allowed !!!! "
   TxtCxTFDate.SetFocus
   Exit Sub
End If

If CmbCMPID.Text <> "" Then
   If TxtCxTFDate.Value < CDate(TxtCxTFDispalyDate) Then
      MsgBox "Date less than BOD Date is not Allowed !!!! "
      TxtCxTFDate.SetFocus
      Exit Sub
   End If
End If
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select mode "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
If CmbAgent.Enabled = True Then
    If CmbAgent.Text = "" Then
       MsgBox "Select Agent "
       
       CmbAgent.SetFocus
       Exit Sub
    End If
End If

If CmbCMPID.Text = "" Then
   MsgBox "Select CMP "
   CmbCMPID.SetFocus
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Select Commodity "
   CmbCommodityID.SetFocus
   Exit Sub
End If
Check1.Enabled = False
If Check1.Value = 1 Then
   If CmbContractID.Text = "" Then
      MsgBox "Select Contract  "
      CmbContractID.SetFocus
      Exit Sub
   End If
End If

If MSHFlexGrid2.Rows <= 3 Then
   MsgBox "Missing Client Detail !!!! "
   Exit Sub
End If
If MSHFlexGrid3.Rows <= 3 Then
   MsgBox "Missing Truck wise Detail !!!! "
   Exit Sub
End If
If MSHFlexGrid4.Rows <= 3 Then
   MsgBox "Missing Godown Stack Lot wise Detail !!!! "
   Exit Sub
End If

'If Val(MSHFlexGrid3.TextMatrix(0, 5)) <> Val(MSHFlexGrid4.TextMatrix(0, 5)) Then ' Done this change on 24/11/2009

tmp = Format(Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(0, 5)), "#######0.000")
  
If Val(tmp) <> 0 Then
   MsgBox "Gross Weight in 'Truck Detail' Tab and 'GSL Detail' tab are not equal !!!!"
   Exit Sub
End If

MSHFlexGrid2.TextMatrix(1, 5) = "No of Bags"
MSHFlexGrid4.TextMatrix(1, 4) = "Electronic Weight   (A+B-C-D) "

If Val(MSHFlexGrid2.TextMatrix(0, 5)) <> Val(MSHFlexGrid4.TextMatrix(0, 4)) Then
   MsgBox "No of Bags in 'Client Detail' Tab and 'GSL Detail' tab are not equal !!!!"
   Exit Sub
End If
If Val(MSHFlexGrid2.TextMatrix(0, 6)) <> Val(MSHFlexGrid4.TextMatrix(0, 10)) Then
   MsgBox "Electronic Weight in 'Client Detail' Tab and 'GSL Detail' tab are not equal !!!!"
   Exit Sub
End If


If Edit_Flg = "A" Then
   RsTxn_CxTF_Details.AddNew
   RsTxn_CxTF_Details!CxTFNo = GetCxTFNo
   RsTxn_CxTF_Details!G_UID = GetGUID
   TxtCxTFNo.Text = RsTxn_CxTF_Details!CxTFNo
Else
    RsTxn_CxTF_Details.MoveFirst
    RsTxn_CxTF_Details.Find " G_UID= '" & TxtG_UID.Text & "'"
    Call DeleteChildTableDetail
    RsTxn_CxTF_Details.MoveFirst
    RsTxn_CxTF_Details.Find " G_UID= '" & TxtG_UID.Text & "'"
End If

RsTxn_CxTF_Details!CxTFNo = TxtCxTFNo.Text
RsTxn_CxTF_Details!ManualCDTFNo = TxtManualCDTFNo
RsTxn_CxTF_Details!CxTFDate = TxtCxTFDate.Value
RsTxn_CxTF_Details!CMPID = mCMPID
RsTxn_CxTF_Details!CommodityID = mCommodityID
RsTxn_CxTF_Details!ExchangeTypeId = mExchangeTypeID
RsTxn_CxTF_Details!BookedSpaceFlag = Check1.Value   'CmbContractID.Text

'RsTxn_CxTF_Details!Flag = TxtFlag.Text
RsTxn_CxTF_Details!Agent = TxtAgent.Text
RsTxn_CxTF_Details!AgentID = mAgentID

'RsTxn_CxTF_Details!Remark = TxtRemark.Text
RsTxn_CxTF_Details!TxnType = "D"
RsTxn_CxTF_Details!SM_Prefix = TxtPreFix

If mContractID <> "" Then
   RsTxn_CxTF_Details!ContractID = mContractID
End If


If IsNull(RsTxn_CxTF_Details!CreatedBy) = True Or RsTxn_CxTF_Details!CreatedBy = "" Then
   RsTxn_CxTF_Details!CreatedBy = Gen_UserLoginID
   RsTxn_CxTF_Details!CreatedDate = Now
Else
   RsTxn_CxTF_Details!UpdatedBy = Gen_UserLoginID
   RsTxn_CxTF_Details!UpdatedDate = Now
End If


RsTxn_CxTF_Details.Update

SaveClient
SaveTruck
SaveGSL
SaveCmd.Enabled = False
Call InData
End Sub
Function getCxtf_ClinetMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(ClientDetailID) from Txn_CxTF_ClientDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCxtf_ClinetMax = RetVal

End Function
Private Sub SaveClient()
Call TableTxn_CxTF_ClientDetail(" Where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")

For i = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
       RsTxn_CxTF_ClientDetail.AddNew
       RsTxn_CxTF_ClientDetail!G_UID = GetGUID
       RsTxn_CxTF_ClientDetail!ClientDetailID = getCxtf_ClinetMax
       RsTxn_CxTF_ClientDetail!CxTFNo = TxtCxTFNo
       RsTxn_CxTF_ClientDetail!ClientIDRow = MSHFlexGrid2.TextMatrix(i, 7)
       RsTxn_CxTF_ClientDetail!TxnType = "D"
       RsTxn_CxTF_ClientDetail!SM_Prefix = TxtPreFix
       RsTxn_CxTF_ClientDetail.Update
       RsTxn_CxTF_ClientDetail.Requery
    End If
Next
End Sub

Private Sub InDataClient()
Dim ans As Variant

Call TableTxn_CxTF_ClientDetail(" Where TxnType = 'D'  And CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_CxTF_ClientDetail.RecordCount
For i = 2 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(i, 0) = RsTxn_CxTF_ClientDetail!ClientDetailID
    ans = GetClientNameByRow(RsTxn_CxTF_ClientDetail!ClientIDRow)
    MSHFlexGrid2.TextMatrix(i, 1) = ans(0)
    MSHFlexGrid2.TextMatrix(i, 2) = ans(1) 'TmpRs!ClientName
    MSHFlexGrid2.TextMatrix(i, 3) = ans(2) 'TmpRs!DepositoryID
    MSHFlexGrid2.TextMatrix(i, 4) = ans(3) 'TmpRs!DepositoryName
    MSHFlexGrid2.TextMatrix(i, 7) = RsTxn_CxTF_ClientDetail!ClientIDRow
    RsTxn_CxTF_ClientDetail.MoveNext
Next
'MsgBox TmpRs.State

End Sub
Function getCxtf_TruckMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(TruckDetailID) from Txn_CxTF_TruckDetail where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCxtf_TruckMax = RetVal

End Function
Private Sub SaveTruck()
Call TableTxn_CxTF_TruckDetail(" Where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")
For i = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(i, 1) <> "" Then
       RsTxn_CxTF_TruckDetail.AddNew
       RsTxn_CxTF_TruckDetail!TruckDetailID = getCxtf_TruckMax
       RsTxn_CxTF_TruckDetail!G_UID = GetGUID
       RsTxn_CxTF_TruckDetail!CxTFNo = TxtCxTFNo
       RsTxn_CxTF_TruckDetail!TruckNo = MSHFlexGrid3.TextMatrix(i, 1)
       
       RsTxn_CxTF_TruckDetail!WeighBridgeSlipno = MSHFlexGrid3.TextMatrix(i, 2)
       RsTxn_CxTF_TruckDetail!TotalWeight = MSHFlexGrid3.TextMatrix(i, 3)
       RsTxn_CxTF_TruckDetail!TareWeight = Val(MSHFlexGrid3.TextMatrix(i, 4))
       RsTxn_CxTF_TruckDetail!GrossWeight = MSHFlexGrid3.TextMatrix(i, 5)
       RsTxn_CxTF_TruckDetail!Remarks = MSHFlexGrid3.TextMatrix(i, 6)
       RsTxn_CxTF_TruckDetail!TxnType = "D"
       RsTxn_CxTF_TruckDetail!SM_Prefix = TxtPreFix
       RsTxn_CxTF_TruckDetail.Update
       RsTxn_CxTF_TruckDetail.Requery
    End If
Next
End Sub
Function getCxtf_GSLMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(GSLDetailID) from Txn_CxTF_GSL where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCxtf_GSLMax = RetVal

End Function

Private Sub InDataGSL()
tmp = "select tgsl.GSLDetailID 'Row ID',"
tmp = tmp & " god.GodownNo 'GodownNo',"
tmp = tmp & " gsl.StackNo,"
tmp = tmp & " gsl.LotNo,"
tmp = tmp & " tgsl.NoofBags,"
tmp = tmp & " tgsl.GrossWeight,"
tmp = tmp & " tgsl.AdjNClientWeight,"
tmp = tmp & " tgsl.AdjNCMSLWeight,"
tmp = tmp & " tgsl.GunnyWeight,"
tmp = tmp & " tgsl.StdDeduction,"
tmp = tmp & " tgsl.DematWeight,"
tmp = tmp & " c.ClientName + ' ' + d.DepositoryName,"
tmp = tmp & " tgsl.Remarks,"
tmp = tmp & " tgsl.GSLID,"
tmp = tmp & " god.CloseDate,"
tmp = tmp & " god.GodownID,"
tmp = tmp & " tgsl.ClientIDRow,"
tmp = tmp & " gsl.ISINID,"
tmp = tmp & " gsl.FGSLID, "
tmp = tmp & " god.address"
tmp = tmp & " from"
tmp = tmp & " Txn_CxTF_GSL tgsl"
tmp = tmp & " inner join mst_gsl gsl on gsl.GSLID = tgsl.GSLID and gsl.SM_Prefix = tgsl.SM_Prefix"
tmp = tmp & " inner join Mst_Godown god on gsl.GodownID = god.GodownID"
tmp = tmp & " inner join mst_client c on tgsl.ClientIDRow = c.ClientIDRow"
tmp = tmp & " inner join mst_depository d on d.DepositoryID = c.DepositoryID  and d.ExchangeTypeID = " & mExchangeTypeID
tmp = tmp & " Where tgsl.TxnType = 'D'  And  tgsl.CxTFNo = " & TxtCxTFNo & " And tgsl.SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable

MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs.RecordCount
If TmpRs.RecordCount = 0 Then Exit Sub



Dim ans As Variant

For i = 2 To MSHFlexGrid4.Rows - 2
    For c = 0 To 16
        MSHFlexGrid4.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If Not IsNull(TmpRs!FGSLID) Then
       If MSHFlexGrid4.TextMatrix(i, 17) = "" Or MSHFlexGrid4.TextMatrix(i, 17) = "0" Then
          MSHFlexGrid4.TextMatrix(i, 17) = IIf(IsNull(TmpRs!FGSLID), "", TmpRs!FGSLID)
       End If
    End If
    MSHFlexGrid4.TextMatrix(i, 18) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
    
    MSHFlexGrid4.Col = 5
    MSHFlexGrid4.Row = i
    MSHFlexGrid4.CellBackColor = &HFFC0C0
    
    MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(i, 4))
    MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(i, 5))
    MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(i, 6))
    MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(i, 7))
    MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(i, 8))
    MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(i, 9))
    MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(i, 10))
    
    MSHFlexGrid4.TextMatrix(i, 4) = Format(MSHFlexGrid4.TextMatrix(i, 4), "#######0")
    MSHFlexGrid4.TextMatrix(i, 5) = Format(MSHFlexGrid4.TextMatrix(i, 5), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 6) = Format(MSHFlexGrid4.TextMatrix(i, 6), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 7) = Format(MSHFlexGrid4.TextMatrix(i, 7), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 8) = Format(MSHFlexGrid4.TextMatrix(i, 8), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 9) = Format(MSHFlexGrid4.TextMatrix(i, 9), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 10) = Format(MSHFlexGrid4.TextMatrix(i, 10), "#######0.000")

    MSHFlexGrid4.TextMatrix(i, 0) = Format(MSHFlexGrid4.TextMatrix(0, 4), "#######0")
    MSHFlexGrid4.TextMatrix(i, 0) = Format(MSHFlexGrid4.TextMatrix(0, 5), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 0) = Format(MSHFlexGrid4.TextMatrix(0, 6), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 0) = Format(MSHFlexGrid4.TextMatrix(0, 7), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 0) = Format(MSHFlexGrid4.TextMatrix(0, 8), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 0) = Format(MSHFlexGrid4.TextMatrix(0, 9), "#######0.000")
    MSHFlexGrid4.TextMatrix(i, 0) = Format(MSHFlexGrid4.TextMatrix(0, 10), "#######0.000")
    
    TmpRs.MoveNext
Next


End Sub


Private Sub InDataGSL_Old()
Call TableTxn_CxTF_GSL(" Where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + RsTxn_CxTF_GSL.RecordCount
If RsTxn_CxTF_GSL.RecordCount = 0 Then Exit Sub

Dim ans As Variant

For i = 2 To MSHFlexGrid4.Rows - 2
    
    ans = GetGSLByGLSID(RsTxn_CxTF_GSL!GSLID, TxtPreFix)
    MSHFlexGrid4.TextMatrix(i, 0) = RsTxn_CxTF_GSL!GSLDetailID
    
    MSHFlexGrid4.TextMatrix(i, 1) = ans(3) ' TmpRs!GodownNo
    MSHFlexGrid4.TextMatrix(i, 2) = ans(4) 'TmpRs!StackNo
    MSHFlexGrid4.TextMatrix(i, 3) = ans(5) 'TmpRs!LotNo
    MSHFlexGrid4.TextMatrix(i, 14) = IIf(IsNull(ans(6)), "", ans(6))
    MSHFlexGrid4.TextMatrix(i, 15) = IIf(IsNull(ans(2)), "", ans(2))
    MSHFlexGrid4.TextMatrix(i, 17) = IIf(IsNull(ans(10)) Or ans(10) = 0, "", ans(10))
    
    MSHFlexGrid4.TextMatrix(i, 4) = RsTxn_CxTF_GSL!NoofBags
    MSHFlexGrid4.TextMatrix(i, 5) = RsTxn_CxTF_GSL!GrossWeight
    MSHFlexGrid4.TextMatrix(i, 6) = RsTxn_CxTF_GSL!AdjNClientWeight
    MSHFlexGrid4.TextMatrix(i, 7) = RsTxn_CxTF_GSL!AdjNCMSLWeight
    MSHFlexGrid4.TextMatrix(i, 8) = RsTxn_CxTF_GSL!GunnyWeight
    MSHFlexGrid4.TextMatrix(i, 9) = RsTxn_CxTF_GSL!StdDeduction
    MSHFlexGrid4.TextMatrix(i, 10) = RsTxn_CxTF_GSL!DematWeight
    
    ans = GetClientNameByRow(RsTxn_CxTF_GSL!ClientIDRow)
    MSHFlexGrid4.TextMatrix(i, 11) = ans(1) & " ~ " & ans(3)
    MSHFlexGrid4.TextMatrix(i, 12) = RsTxn_CxTF_GSL!Remarks
    
    MSHFlexGrid4.TextMatrix(i, 13) = RsTxn_CxTF_GSL!GSLID
   
    'MSHFlexGrid4.TextMatrix(i, 14) = 'look at top Used Up
    'MSHFlexGrid4.TextMatrix(i, 15) = 'look at top Used Up
    'MSHFlexGrid4.TextMatrix(i, 17) = 'look at top Used Up
    
    MSHFlexGrid4.TextMatrix(i, 16) = RsTxn_CxTF_GSL!ClientIDRow
    
    MSHFlexGrid4.Col = 5
    MSHFlexGrid4.Row = i
    MSHFlexGrid4.CellBackColor = &HFFC0C0

    
    MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(i, 4))
    MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(i, 5))
    MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(i, 6))
    MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(i, 7))
    MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(i, 8))
    MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(i, 9))
    MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(i, 10))
    
    RsTxn_CxTF_GSL.MoveNext
Next
End Sub

Private Sub SaveGSL()
Dim adjWt As Variant

Call TableTxn_CxTF_GSL(" Where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")
Call TableMst_GSL("where SM_Prefix = '" & TxtPreFix & "'")
For i = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(i, 1) <> "" Then
       RsTxn_CxTF_GSL.AddNew
       RsTxn_CxTF_GSL!GSLDetailID = getCxtf_GSLMax
       RsTxn_CxTF_GSL!G_UID = GetGUID
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid4.TextMatrix(i, 13))
'       If Not RsMst_GSL.EOF Then
'          RsMst_GSL!CommodityID = mCommodityID
'          RsMst_GSL!ExchangeTypeID = mExchangeTypeID
'          RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsTxn_CxTF_GSL!NoofBags), 0, RsTxn_CxTF_GSL!NoofBags)
'          RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsTxn_CxTF_GSL!DematWeight), 0, RsTxn_CxTF_GSL!DematWeight)
'          RsMst_GSL.Update
'       End If
        
       RsTxn_CxTF_GSL!CxTFNo = TxtCxTFNo
       RsTxn_CxTF_GSL!NoofBags = MSHFlexGrid4.TextMatrix(i, 4)
       RsTxn_CxTF_GSL!GrossWeight = MSHFlexGrid4.TextMatrix(i, 5)
       RsTxn_CxTF_GSL!AdjNClientWeight = Val(MSHFlexGrid4.TextMatrix(i, 6))
       RsTxn_CxTF_GSL!AdjNCMSLWeight = Val(MSHFlexGrid4.TextMatrix(i, 7))
       RsTxn_CxTF_GSL!GunnyWeight = Val(MSHFlexGrid4.TextMatrix(i, 8))
       RsTxn_CxTF_GSL!StdDeduction = Val(MSHFlexGrid4.TextMatrix(i, 9))
       RsTxn_CxTF_GSL!DematWeight = Val(MSHFlexGrid4.TextMatrix(i, 10))
       RsTxn_CxTF_GSL!ClientIDRow = Val(MSHFlexGrid4.TextMatrix(i, 16))
       RsTxn_CxTF_GSL!Remarks = MSHFlexGrid4.TextMatrix(i, 12)
       RsTxn_CxTF_GSL!GSLID = MSHFlexGrid4.TextMatrix(i, 13)
       RsTxn_CxTF_GSL!TxnType = "D"
       RsTxn_CxTF_GSL!SM_Prefix = TxtPreFix
       RsTxn_CxTF_GSL.Update
       RsTxn_CxTF_GSL.Requery
       
       If Not RsMst_GSL.EOF Then
          RsMst_GSL!CommodityID = mCommodityID
          RsMst_GSL!ExchangeTypeId = mExchangeTypeID
          RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) + Val(MSHFlexGrid4.TextMatrix(i, 4))
          RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(MSHFlexGrid4.TextMatrix(i, 10))
          
          RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(MSHFlexGrid4.TextMatrix(i, 4))
          RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(MSHFlexGrid4.TextMatrix(i, 10))
          
          adjWt = Val(MSHFlexGrid4.TextMatrix(i, 6)) - Val(MSHFlexGrid4.TextMatrix(i, 9))
          RsMst_GSL!GrossWt = IIf(IsNull(RsMst_GSL!GrossWt), 0, RsMst_GSL!GrossWt) + Val(MSHFlexGrid4.TextMatrix(i, 5))
          
          RsMst_GSL!LCDTF_Date = TxtCxTFDate.Value
          RsMst_GSL!KRNo = IIf(IsNull(RsMst_GSL!KRNo), 0, RsMst_GSL!KRNo)
          
          
          If IsNull(RsMst_GSL!CDTFDate) Then
             RsMst_GSL!CdtfNo = Val(TxtCxTFNo)
             RsMst_GSL!CDTFDate = TxtCxTFDate.Value
          End If
          RsMst_GSL!Flag = "O"
          
          RsMst_GSL!ClientIDRow = Val(MSHFlexGrid4.TextMatrix(i, 16))
          RsMst_GSL.Update
       End If
       
    End If
Next
End Sub



Private Sub InDataTruck()
Call TableTxn_CxTF_TruckDetail(" Where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")
MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + RsTxn_CxTF_TruckDetail.RecordCount
For i = 2 To MSHFlexGrid3.Rows - 2
   MSHFlexGrid3.TextMatrix(i, 0) = RsTxn_CxTF_TruckDetail!TruckDetailID
   MSHFlexGrid3.TextMatrix(i, 1) = RsTxn_CxTF_TruckDetail!TruckNo
   MSHFlexGrid3.TextMatrix(i, 2) = RsTxn_CxTF_TruckDetail!WeighBridgeSlipno
   MSHFlexGrid3.TextMatrix(i, 3) = RsTxn_CxTF_TruckDetail!TotalWeight
   MSHFlexGrid3.TextMatrix(i, 4) = RsTxn_CxTF_TruckDetail!TareWeight
   MSHFlexGrid3.TextMatrix(i, 5) = RsTxn_CxTF_TruckDetail!GrossWeight
   MSHFlexGrid3.TextMatrix(i, 6) = RsTxn_CxTF_TruckDetail!Remarks
   
   MSHFlexGrid3.Col = 5
   MSHFlexGrid3.Row = i
   MSHFlexGrid3.CellBackColor = &HFFC0C0
       
   MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + RsTxn_CxTF_TruckDetail!TotalWeight
   MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + RsTxn_CxTF_TruckDetail!TareWeight
   MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + RsTxn_CxTF_TruckDetail!GrossWeight
   
   RsTxn_CxTF_TruckDetail.MoveNext
Next
End Sub

Private Function GetCxTFNo() As Double
Dim RetVal As Double
tmp = "Select max(CxTFNo) from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetCxTFNo = RetVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 15
MSHFlexGrid1.TextMatrix(0, 0) = "CDTF No"
MSHFlexGrid1.TextMatrix(0, 1) = "CDTF Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 3) = "CMP"
MSHFlexGrid1.TextMatrix(0, 4) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 5) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 6) = "Flag"
MSHFlexGrid1.TextMatrix(0, 7) = "Agent"
MSHFlexGrid1.TextMatrix(0, 8) = "Remark"
MSHFlexGrid1.TextMatrix(0, 9) = "Hologram No"
MSHFlexGrid1.TextMatrix(0, 10) = "Agent Name"
MSHFlexGrid1.TextMatrix(0, 11) = "Manual CDTFNo"
MSHFlexGrid1.TextMatrix(0, 12) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 13) = "Gross Wt."
MSHFlexGrid1.TextMatrix(0, 14) = "Guid"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 2600
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 0 '1200
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 0 '1600
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 0 '1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(14) = 0

MSHFlexGrid1.RowHeight(0) = 700
End Sub
Private Sub Grid_Head_Client()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.Cols = 8
MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Client ID"
MSHFlexGrid2.TextMatrix(1, 2) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 3) = "DP ID"
MSHFlexGrid2.TextMatrix(1, 4) = "DP Name "
MSHFlexGrid2.TextMatrix(1, 5) = "No of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Weight"
MSHFlexGrid2.TextMatrix(1, 7) = "ClientRowID"


MSHFlexGrid2.ColWidth(0) = 500
MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 3500
MSHFlexGrid2.ColWidth(3) = 1800
MSHFlexGrid2.ColWidth(4) = 3500
MSHFlexGrid2.ColWidth(5) = 1500
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(7) = 0

End Sub
Private Sub Grid_Head_TTD()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 2

MSHFlexGrid3.Cols = 7
MSHFlexGrid3.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid3.TextMatrix(1, 1) = "Truck No"
MSHFlexGrid3.TextMatrix(1, 2) = "Weighbridge Slip No"
MSHFlexGrid3.TextMatrix(1, 3) = "Total Weight"
MSHFlexGrid3.TextMatrix(1, 4) = "Tare Weight"
MSHFlexGrid3.TextMatrix(1, 5) = "Gross Weight"
MSHFlexGrid3.TextMatrix(1, 6) = "Remark"




MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 2000
MSHFlexGrid3.ColWidth(2) = 2500
MSHFlexGrid3.ColWidth(3) = 1500
MSHFlexGrid3.ColWidth(4) = 1500
MSHFlexGrid3.ColWidth(5) = 1500
MSHFlexGrid3.ColWidth(6) = 4500

'MSHFlexGrid3.Col = 5
'MSHFlexGrid3.Row = 0
'MSHFlexGrid3.CellBackColor = vbRed '&HFFC0C0


End Sub
Private Sub Grid_Head_GCTD()
MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 3
MSHFlexGrid4.FixedRows = 2

MSHFlexGrid4.Cols = 19
MSHFlexGrid4.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid4.TextMatrix(1, 1) = "Godown No"
MSHFlexGrid4.TextMatrix(1, 2) = "Stack No"
MSHFlexGrid4.TextMatrix(1, 3) = "Lot No"
MSHFlexGrid4.TextMatrix(1, 4) = "No of Bags"
MSHFlexGrid4.TextMatrix(1, 5) = "Gross Weight    (A) "
MSHFlexGrid4.TextMatrix(1, 6) = "Adjusted Wt.    (B) "
MSHFlexGrid4.TextMatrix(1, 7) = "Adj.Gross Wt."
MSHFlexGrid4.TextMatrix(1, 8) = "Gunny Weight    (C) "
MSHFlexGrid4.TextMatrix(1, 9) = "Std. Deduction    (D) "
MSHFlexGrid4.TextMatrix(1, 10) = "Electronic Weight   (A+B-C-D) "
MSHFlexGrid4.TextMatrix(1, 11) = "Client ID"
MSHFlexGrid4.TextMatrix(1, 12) = "Remark"
MSHFlexGrid4.TextMatrix(1, 13) = "GSL ID"
MSHFlexGrid4.TextMatrix(1, 14) = "Godown Close"
MSHFlexGrid4.TextMatrix(1, 15) = "Godown ID"
MSHFlexGrid4.TextMatrix(1, 16) = "ClientRowID"
MSHFlexGrid4.TextMatrix(1, 17) = "ISINID"
MSHFlexGrid4.TextMatrix(1, 18) = "Godown Address"



MSHFlexGrid4.ColWidth(0) = 0
MSHFlexGrid4.ColWidth(1) = 1200
MSHFlexGrid4.ColWidth(2) = 1200
MSHFlexGrid4.ColWidth(3) = 1200
MSHFlexGrid4.ColWidth(4) = 1000
MSHFlexGrid4.ColWidth(7) = 0
MSHFlexGrid4.ColWidth(11) = 3000
MSHFlexGrid4.ColWidth(12) = 1500
MSHFlexGrid4.ColWidth(13) = 0
MSHFlexGrid4.ColWidth(14) = 0
MSHFlexGrid4.ColWidth(15) = 0
MSHFlexGrid4.ColWidth(16) = 0
MSHFlexGrid4.ColWidth(17) = 0
MSHFlexGrid4.RowHeight(1) = 750
MSHFlexGrid4.ColWidth(18) = 0


'MSHFlexGrid4.Col = 5
'MSHFlexGrid4.Row = 0
'MSHFlexGrid4.CellBackColor = &HFFC0C0
'



End Sub

Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CxTF_Details.MoveFirst
RsTxn_CxTF_Details.Find "g_uid = '" & MSHFlexGrid1.TextMatrix(i, 14) & "'"
If Not RsTxn_CxTF_Details.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub CmdCMPID_Click()
    FrmMst_CMP.Show
End Sub
'----
Private Sub CmbCMPID_GotFocus()

Wc = ""
If Gen_WcCMP <> "" Then
   Wc = " And " & Gen_WcCMP
End If

    tmp = CmbCMPID.Text
    Call TableMst_CMP(" where CloseDate is null " & Wc)
    CmbCMPID.Clear
    If RsMst_CMP.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For i = 1 To RsMst_CMP.RecordCount
        CmbCMPID.AddItem RsMst_CMP!CMPName
        RsMst_CMP.MoveNext
    Next
    CmbCMPID.Text = tmp

End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   Call Grid_Head_GCTD
   CmbContractID.Text = ""
End If

mCMPID = RsMst_CMP!CMPID
TxtCxTFDate.Value = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate) + 1
TxtCxTFDispalyDate = TxtCxTFDate.Value
TxtCxTFDate.Enabled = False
'If LCase(CmbCMPID.Text) = "unjha" Or LCase(CmbCMPID.Text) = "rajkot" Then
'   TxtCxTFDate.Enabled = True
'Else
'   TxtCxTFDate.Enabled = False
'End If

'If SaveCmd.Enabled = True Then
'   MsgBox "CDTF Date is " & TxtCxTFDate
'End If
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub
'----
Private Sub CmdCommodityId_Click()
    FrmMst_Commodity.Show
End Sub
'----
Private Sub CmbCommodityID_GotFocus()

 If CmbExchangeTypeID.Text <> "" Then
'    If CmbCMPID.Text <> "" Then
       tmp = CmbCommodityID.Text
       Call TableMst_Commodity
       CmbCommodityID.Clear
       If RsMst_Commodity.RecordCount = 0 Then
          MsgBox "No Record found !!!! "
          Exit Sub
       End If
       For i = 1 To RsMst_Commodity.RecordCount
           CmbCommodityID.AddItem RsMst_Commodity!Commodity
           RsMst_Commodity.MoveNext
       Next
       CmbCommodityID.Text = tmp
    
'    Else
'       MsgBox "Please Select CMP!!!!!"
'       CmbCMPID.SetFocus
'       Exit Sub
'    End If
 Else
    MsgBox "Please Select Exchange Mode !!!!"
    CmbExchangeTypeID.SetFocus
    Exit Sub
 End If
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If

If mCommodityID <> RsMst_Commodity!CommodityID Then
   Call Grid_Head_GCTD
   CmbContractID.Text = ""
End If
mCommodityID = RsMst_Commodity!CommodityID

Call TableMst_CommodityDetail("Where CommodityId = " & mCommodityID & " and ExchangeTypeId = " & mExchangeTypeID)
 If RsMst_CommodityDetail.RecordCount > 0 Then
    TxtCGunnyWt = RsMst_CommodityDetail!GunnyWt
    TxtCStdDed = IIf(IsNull(RsMst_CommodityDetail!StdDeduction), "", RsMst_CommodityDetail!StdDeduction)
    TxtCLotSize = IIf(IsNull(RsMst_CommodityDetail!LotSize), "", RsMst_CommodityDetail!LotSize)
    TxtCVariation = IIf(IsNull(RsMst_CommodityDetail!Variation), "", RsMst_CommodityDetail!Variation)
 Else
    MsgBox "No Record Found For Exchange" & CmbExchangeTypeID.Text & " - " & CmbCommodityID.Text
    CmbCommodityID.SetFocus
    Exit Sub
 End If
'SSTab1.Tab = 0
End Sub
'----
Private Sub CmdExchangeTypeID_Click()
    Frm_Mst_ExchangeType.Show
End Sub
'----
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
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeId Then
   Call Grid_Head_Client
   Call Grid_Head_GCTD
   Call ClearClientTabControl
   Call ClearGCTDTabControl
   CmbContractID.Text = ""
   Check1.Value = 0
   CmdClientAdd2List.Enabled = True
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeId


CmbCommodityID.Text = ""
TxtCGunnyWt = ""
TxtCStdDed = ""
TxtCLotSize = ""
TxtCVariation = ""
If mExchangeTypeID = 1 Then
   CmbNonNcdexClient.Width = 10095
   CmbNonNcdexClient.Top = TxtClientID.Top
   CmbNonNcdexClient.Left = TxtClientID.Left
   Label9.Visible = False
   Label10.Visible = False
   Label11.Visible = False
   
   CmbNonNcdexClient.Visible = True
Else
   Label9.Visible = True
   Label10.Visible = True
   Label11.Visible = True
   CmbNonNcdexClient.Visible = False
End If
Call Check1_Click
End Sub
'----
Private Sub CmdContractID_Click()
    FrmMst_StorageContract.Show
End Sub
'----
Private Sub CmbContractID_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select CMP !!!!! "
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Select CMP !!!!! "
   Exit Sub
End If
Call Grid_Head_GCTD
Call Grid_Head_Client
tmp = CmbContractID.Text
mLocationID = GetLocationIDFromCMP(Val(mCMPID))
Call TableMst_StorageContract(" where (CommodityID is null or CommodityID = " & mCommodityID & " ) And LocationID = " & mLocationID)
CmbContractID.Clear
If RsMst_StorageContract.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Check1.Value = 0
   Exit Sub
End If
For i = 1 To RsMst_StorageContract.RecordCount
    CmbContractID.AddItem RsMst_StorageContract!ContractName
    RsMst_StorageContract.MoveNext
Next
CmbContractID.Text = tmp
End Sub
'----
Private Sub CmbContractID_LostFocus()
If CmbContractID.Text = "" Then
   Exit Sub
End If
RsMst_StorageContract.MoveFirst
RsMst_StorageContract.Find "ContractName = '" & CmbContractID.Text & "'"
If RsMst_StorageContract.EOF Then
   MsgBox "Invalid selection "
   CmbContractID.SetFocus
   Exit Sub
End If
Call Grid_Head_Client
mContractID = RsMst_StorageContract!ContractID
mContractClientIDRow = RsMst_StorageContract!ClientIDRow
ans = GetClientNameByRow(Val(mContractClientIDRow))
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)
mClientIDRow = mContractClientIDRow
Call CmdClientAdd2List_Click
CmdClientAdd2List.Enabled = False
End Sub
'----

Private Sub TxtGCTDGrossWt_LostFocus()
If TxtGCTDGrossWt = "" Then Exit Sub

If IsNumeric(TxtGCTDGrossWt) = False Then
   MsgBox "Invalid numeric format !!!"
   TxtGCTDGrossWt.SetFocus
   Exit Sub
End If
Call CalDmatWt

End Sub

Private Sub TxtGCTDNoofBags_LostFocus()
If TxtGCTDNoofBags = "" Then Exit Sub

If IsNumeric(TxtGCTDNoofBags) = False Then
   MsgBox "Invalid numeric format !!!"
   TxtGCTDNoofBags.SetFocus
   Exit Sub
End If
Call CalDmatWt
End Sub

Private Sub CalDmatWt()

TxtGCTDGunnyWt = Format((Val(TxtGCTDNoofBags) * Val(TxtCGunnyWt)) / 1000, "#######0.000")
TxtAdjGrossWt = Val(TxtGCTDGrossWt) + Val(TxtAdjWt)
tmp = Val(TxtAdjGrossWt) - Val(TxtGCTDGunnyWt)
TxtGCTDStdDeduction = Format(((tmp) * Val(TxtCStdDed)) / 100, "#######0.000")
TxtGCTDDematWt = Format(Val(TxtAdjGrossWt) - (Val(TxtGCTDStdDeduction) + Val(TxtGCTDGunnyWt)), "#######0.000")


End Sub


Private Sub TxtTTDTareWt_LostFocus()
If TxtTTDTareWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDTareWt) Then
   MsgBox "Invalid numeric format"
   TxtTTDTareWt.SetFocus
End If
Call CalTruckWt
End Sub
Private Sub CalTruckWt()
'If Val(TxtTTDTotalWt) = 0 Then Exit Sub
'If Val(TxtTTDTareWt) = 0 Then Exit Sub

TxtTTDGrossWt = Val(TxtTTDTotalWt) - Val(TxtTTDTareWt)

End Sub
Private Sub TxtTTDGrossWt_LostFocus()
If TxtTTDGrossWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDGrossWt) Then
   MsgBox "Invalid numeric format"
   TxtTTDGrossWt.SetFocus
End If

End Sub

Private Sub TxtTTDTotalWt_LostFocus()
If TxtTTDTotalWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDTotalWt) Then
   MsgBox "Invalid numeric format"
   TxtTTDTotalWt.SetFocus
   Exit Sub
End If
Call CalTruckWt
End Sub

' Location Pre-fix change to Location name on 16 jan 10

'Private Sub CmbPreFix_GotFocus()
'tmp = CmbPreFix.Text
'Call TableMst_StoreManLocations("")
'
'CmbPreFix.Clear
'If RsMst_StoreManLocations.RecordCount = 0 Then
'   MsgBox "No Bank found"
'   CmbPreFix.SetFocus
'   Exit Sub
'End If
'For i = 1 To RsMst_StoreManLocations.RecordCount
'    CmbPreFix.AddItem RsMst_StoreManLocations!SM_Prefix
'    RsMst_StoreManLocations.MoveNext
'Next
'CmbPreFix.Text = tmp
'End Sub
'
'Private Sub CmbPreFix_LostFocus()
'If CmbPreFix.Text = "" Then
'   Exit Sub
'End If
'RsMst_StoreManLocations.MoveFirst
'RsMst_StoreManLocations.Find "SM_Prefix = '" & CmbPreFix.Text & "'"
'
'If RsMst_StoreManLocations.EOF Then
'   MsgBox "Invalid PreFix Selection !!!!"
'   CmbPreFix.SetFocus
'   Exit Sub
'End If
'msSM_prefix = RsMst_StoreManLocations!SM_Prefix
'
'End Sub


'Private Sub CmbPreFix_GotFocus()
'tmp = CmbPreFix.Text
'Call TableMst_Location
'CmbPreFix.Clear
'If RsMst_Location.RecordCount = 0 Then
'   MsgBox "No Record found !!!! "
'   Exit Sub
'End If
'For i = 1 To RsMst_Location.RecordCount
'    CmbPreFix.AddItem RsMst_Location!LocationName
'    RsMst_Location.MoveNext
'Next
'CmbPreFix.Text = tmp
'End Sub
'
'Private Sub CmbPreFix_LostFocus()
'If CmbPreFix.Text = "" Then
'   mLocationID = Null
'   mSM_Prefix = Null
'   Exit Sub
'End If
'RsMst_Location.MoveFirst
'RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"
'If RsMst_Location.EOF Then
'   MsgBox "Invalid selection "
'   CmbPreFix.SetFocus
'   Exit Sub
'End If
'
'If tmp <> CmbPreFix.Text Then
'   CmbSCmp.Text = ""
'   mCMPID = Null
'End If
'mSLocationId = RsMst_Location!LocationID
'mLocationName = RsMst_Location!LocationName
'msSM_prefix = RsMst_Location!SM_Prefix
'End Sub



Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If


CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   mLocationID = Null
   mSM_Prefix = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbPreFix.SetFocus
   Exit Sub
End If

If tmp <> CmbPreFix.Text Then
   CmbSCmp.Text = ""
   mCMPID = Null
End If
mSLocationId = RsMst_Location!LocationID
mLocationName = RsMst_Location!LocationName
msSM_prefix = RsMst_Location!SM_Prefix
End Sub


Private Sub CmbSCmp_GotFocus()
If msSM_prefix = "" Then
   MsgBox "Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationId & "")
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbSCmp.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp
End Sub
Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
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
For i = 1 To RsMst_Commodity.RecordCount
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
For i = 1 To RsMst_ExchangeType.RecordCount
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
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeId

End Sub
