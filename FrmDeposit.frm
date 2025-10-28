VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmDeposit 
   Caption         =   "Deposit"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   154
      Top             =   9600
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   153
      Top             =   0
      Width           =   255
   End
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
      TabIndex        =   125
      Top             =   4440
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox ChkRecAnnexure 
         BackColor       =   &H008080FF&
         Caption         =   "Received Annexure-III"
         Height          =   255
         Left            =   1680
         TabIndex        =   168
         Top             =   1320
         Width           =   2115
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H008080FF&
         Caption         =   "Generate PDF"
         Height          =   255
         Left            =   120
         TabIndex        =   141
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         Left            =   240
         TabIndex        =   147
         Top             =   1680
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
         TabIndex        =   129
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame Frame0 
      Height          =   1725
      Left            =   480
      TabIndex        =   115
      Top             =   5640
      Visible         =   0   'False
      Width           =   13635
      Begin VB.ComboBox CmbSSCMFlag 
         Height          =   315
         ItemData        =   "FrmDeposit.frx":0092
         Left            =   9960
         List            =   "FrmDeposit.frx":009C
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   840
         Width           =   1125
      End
      Begin VB.TextBox TxtSTruckNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6705
         MaxLength       =   25
         TabIndex        =   51
         Top             =   840
         Width           =   2460
      End
      Begin VB.TextBox TxtSManualCxTFNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4530
         MaxLength       =   25
         TabIndex        =   50
         Top             =   810
         Width           =   1320
      End
      Begin VB.TextBox TxtSCxTFNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1650
         MaxLength       =   25
         TabIndex        =   49
         Top             =   810
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   1680
         TabIndex        =   44
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
         Format          =   38338561
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   120
         TabIndex        =   43
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
         Format          =   38338561
         CurrentDate     =   36494
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   12360
         TabIndex        =   53
         Top             =   870
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   450
         Width           =   3645
      End
      Begin VB.ComboBox CmbSExchangeType 
         Height          =   315
         Left            =   12090
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   450
         Width           =   2205
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   8520
         Sorted          =   -1  'True
         TabIndex        =   47
         Top             =   450
         Width           =   3525
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   3165
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   450
         Width           =   1605
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13395
         TabIndex        =   116
         Top             =   870
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   90
         TabIndex        =   117
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
      Begin VB.ComboBox CmbSAgent 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1305
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.Label Label54 
         Caption         =   "SCM Deposit"
         Height          =   435
         Left            =   9390
         TabIndex        =   169
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label51 
         Caption         =   "Truck No"
         Height          =   225
         Left            =   6000
         TabIndex        =   164
         Top             =   915
         Width           =   750
      End
      Begin VB.Label Label50 
         Caption         =   "Manual CDTF No"
         Height          =   225
         Left            =   3240
         TabIndex        =   152
         Top             =   900
         Width           =   1530
      End
      Begin VB.Label Label49 
         Caption         =   "Store Man CDTF No"
         Height          =   225
         Left            =   120
         TabIndex        =   151
         Top             =   870
         Width           =   1530
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   1980
         TabIndex        =   139
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   420
         TabIndex        =   138
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label42 
         Caption         =   "CMP"
         Height          =   300
         Left            =   6382
         TabIndex        =   137
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label39 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   9802
         TabIndex        =   136
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label38 
         Caption         =   "Exchange Type"
         Height          =   225
         Left            =   12592
         TabIndex        =   135
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label37 
         Caption         =   "Agent Name"
         Height          =   300
         Left            =   120
         TabIndex        =   134
         Top             =   1335
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label36 
         Caption         =   "Location ID"
         Height          =   300
         Left            =   3502
         TabIndex        =   133
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.TextBox TxtG_UID 
      Height          =   495
      Left            =   720
      TabIndex        =   140
      Top             =   10080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Frame4 
      Height          =   9435
      Left            =   60
      ScaleHeight     =   9375
      ScaleWidth      =   14910
      TabIndex        =   55
      Top             =   120
      Width           =   14970
      Begin VB.PictureBox Picture3 
         Height          =   930
         Left            =   120
         ScaleHeight     =   870
         ScaleWidth      =   14655
         TabIndex        =   155
         Top             =   8400
         Width           =   14715
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   375
            Left            =   2520
            TabIndex        =   34
            Top             =   75
            Width           =   2400
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   4920
            TabIndex        =   163
            Top             =   75
            Width           =   2400
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   7320
            TabIndex        =   162
            Top             =   75
            Width           =   2400
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "&Search"
            Height          =   375
            Left            =   9720
            TabIndex        =   36
            Top             =   75
            Width           =   2400
         End
         Begin VB.CommandButton CmdPrintPreview 
            Caption         =   "Print Preview"
            Height          =   375
            Left            =   12120
            TabIndex        =   161
            Top             =   75
            Width           =   2400
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Print"
            Height          =   375
            Left            =   12120
            TabIndex        =   160
            Top             =   435
            Width           =   2400
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   9720
            TabIndex        =   35
            Top             =   450
            Width           =   2400
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   375
            Left            =   7320
            TabIndex        =   159
            Top             =   450
            Width           =   2400
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   375
            Left            =   4920
            TabIndex        =   158
            Top             =   450
            Width           =   2400
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   2520
            TabIndex        =   157
            Top             =   450
            Width           =   2400
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   375
            Left            =   120
            TabIndex        =   156
            Top             =   450
            Width           =   2400
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New "
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   75
            Width           =   2400
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1905
         ScaleWidth      =   14475
         TabIndex        =   69
         Top             =   120
         Width           =   14505
         Begin VB.CheckBox ChkSCMFlag 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   10680
            TabIndex        =   5
            Top             =   585
            Width           =   225
         End
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
            Top             =   1455
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
            TabIndex        =   8
            Top             =   1455
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
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   1455
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
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   1455
            Width           =   795
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
            Left            =   13320
            Locked          =   -1  'True
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   1455
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
            Left            =   13320
            Locked          =   -1  'True
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   1020
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
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   1020
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
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   150
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
            Left            =   8100
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   600
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
            Left            =   13320
            Locked          =   -1  'True
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   585
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
            Left            =   13320
            Locked          =   -1  'True
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   150
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   6
            Top             =   1065
            Width           =   225
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
            TabIndex        =   38
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
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1455
            Width           =   6135
         End
         Begin MSComCtl2.DTPicker TxtCxTFDate 
            Height          =   360
            Left            =   1200
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   585
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   38338561
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
            TabIndex        =   7
            Top             =   1020
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
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   1020
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
            Top             =   585
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
            Top             =   150
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
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   585
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
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   150
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
            ItemData        =   "FrmDeposit.frx":00A9
            Left            =   1200
            List            =   "FrmDeposit.frx":00AB
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   1020
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
            TabIndex        =   112
            Top             =   1020
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
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   585
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label52 
            Caption         =   "SCM Deposit"
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
            Left            =   9105
            TabIndex        =   165
            Top             =   660
            Width           =   1575
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
            TabIndex        =   132
            Top             =   1388
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "Location ID"
            Height          =   255
            Left            =   11505
            TabIndex        =   124
            Top             =   1125
            Width           =   855
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "% Variation "
            Height          =   255
            Left            =   11550
            TabIndex        =   111
            Top             =   1500
            Width           =   1695
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "Lot Size"
            Height          =   255
            Left            =   11520
            TabIndex        =   109
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Std. Deduction"
            Height          =   255
            Left            =   11520
            TabIndex        =   100
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Gunny Wt. per Bag"
            Height          =   255
            Left            =   11535
            TabIndex        =   98
            Top             =   210
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
            TabIndex        =   79
            Top             =   143
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
            TabIndex        =   78
            Top             =   203
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
            Height          =   255
            Left            =   3120
            TabIndex        =   77
            Top             =   638
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
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   638
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
            TabIndex        =   75
            Top             =   1073
            Width           =   735
         End
         Begin VB.Label Label28 
            Caption         =   "Space Reservation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3120
            TabIndex        =   74
            Top             =   960
            Width           =   1200
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
            TabIndex        =   73
            Top             =   1508
            Width           =   855
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6135
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   10821
         _Version        =   393216
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Client Detail"
         TabPicture(0)   =   "FrmDeposit.frx":00AD
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(1)=   "Pic_Client"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Truck Details"
         TabPicture(1)   =   "FrmDeposit.frx":00C9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid3"
         Tab(1).Control(1)=   "Pic_TTD"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "GSL Details"
         TabPicture(2)   =   "FrmDeposit.frx":00E5
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "MSHFlexGrid4"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Pic_GCTD"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Log History"
         TabPicture(3)   =   "FrmDeposit.frx":0101
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label46"
         Tab(3).Control(1)=   "Label47"
         Tab(3).Control(2)=   "TxtCreated"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "TxtUpdated"
         Tab(3).Control(3).Enabled=   0   'False
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
            TabIndex        =   143
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
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   480
            Width           =   9330
         End
         Begin VB.PictureBox Pic_TTD 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   -74880
            ScaleHeight     =   1785
            ScaleWidth      =   14505
            TabIndex        =   88
            Top             =   435
            Width           =   14535
            Begin VB.TextBox TxtForeignMatter 
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
               Left            =   8160
               TabIndex        =   16
               Top             =   780
               Width           =   1575
            End
            Begin VB.TextBox TxtMinimum 
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
               Height          =   360
               Left            =   9510
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   175
               Top             =   75
               Width           =   1200
            End
            Begin VB.TextBox TxtMaximum 
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
               Height          =   360
               Left            =   13200
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   174
               Top             =   75
               Width           =   1245
            End
            Begin VB.ComboBox CmbVehicleTypeID 
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
               TabIndex        =   11
               Top             =   60
               Width           =   5295
            End
            Begin VB.TextBox TxtMoisturePer 
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
               Left            =   12990
               TabIndex        =   19
               Top             =   780
               Width           =   1455
            End
            Begin VB.TextBox TxtTTDNoOfBags 
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
               Left            =   11415
               TabIndex        =   18
               Top             =   780
               Width           =   1575
            End
            Begin VB.CommandButton CmdTTDDeleteList 
               Caption         =   "Delete From List"
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
               Left            =   12750
               TabIndex        =   22
               Top             =   1320
               Width           =   1695
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
               Left            =   2145
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Top             =   1215
               Width           =   9015
            End
            Begin VB.CommandButton CmdTTDAdd2List 
               Caption         =   "Add to List"
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
               TabIndex        =   21
               Top             =   1320
               Width           =   1455
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
               TabIndex        =   15
               Top             =   780
               Width           =   1575
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
               TabIndex        =   14
               Top             =   780
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
               TabIndex        =   13
               Top             =   780
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
               TabIndex        =   12
               Top             =   780
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
               Left            =   9720
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   780
               Width           =   1695
            End
            Begin VB.Label Label61 
               Alignment       =   2  'Center
               Caption         =   "Foreign Matter"
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
               Left            =   8160
               TabIndex        =   178
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label60 
               Caption         =   "Min Tare Weight in MT"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7320
               TabIndex        =   177
               Top             =   113
               Width           =   2100
            End
            Begin VB.Label Label59 
               Caption         =   "Max Tare Weight in MT"
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
               Left            =   10830
               TabIndex        =   176
               Top             =   128
               Width           =   2145
            End
            Begin VB.Label Label57 
               Caption         =   "Vehicle Model"
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
               Left            =   45
               TabIndex        =   173
               Top             =   105
               Width           =   1875
            End
            Begin VB.Label Label55 
               Caption         =   "Moisture %"
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
               Left            =   13125
               TabIndex        =   170
               Top             =   480
               Width           =   1305
            End
            Begin VB.Label Label53 
               Caption         =   "No of Bags"
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
               Left            =   11670
               TabIndex        =   166
               Top             =   480
               Width           =   1305
            End
            Begin VB.Label Label6 
               Caption         =   "Quality Parameters / Remarks"
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
               Left            =   180
               TabIndex        =   94
               Top             =   1230
               Width           =   1935
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
               Left            =   6720
               TabIndex        =   93
               Top             =   480
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
               TabIndex        =   92
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label15 
               Caption         =   "Weighbridge Slip No."
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
               TabIndex        =   91
               Top             =   480
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
               Left            =   990
               TabIndex        =   90
               Top             =   480
               Width           =   660
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
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
               Left            =   9900
               TabIndex        =   89
               Top             =   480
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_GCTD 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   90
            ScaleHeight     =   1905
            ScaleWidth      =   14505
            TabIndex        =   80
            Top             =   420
            Width           =   14535
            Begin VB.TextBox TxtAverageBagSize 
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
               Left            =   10320
               Locked          =   -1  'True
               TabIndex        =   179
               TabStop         =   0   'False
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox TxtMoisturePer1 
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
               Left            =   3960
               Locked          =   -1  'True
               TabIndex        =   171
               TabStop         =   0   'False
               Top             =   840
               Width           =   1095
            End
            Begin VB.ComboBox CmbGSLTruckNo 
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
               ItemData        =   "FrmDeposit.frx":011D
               Left            =   3900
               List            =   "FrmDeposit.frx":011F
               Sorted          =   -1  'True
               TabIndex        =   26
               Top             =   420
               Width           =   2910
            End
            Begin VB.TextBox TxtGodAdd 
               Height          =   285
               Left            =   9240
               TabIndex        =   150
               Top             =   1800
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.CommandButton CmdGCTDeleteList 
               Caption         =   "Delete From List"
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
               Left            =   12750
               TabIndex        =   33
               Top             =   1380
               Width           =   1695
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
               Left            =   13440
               Locked          =   -1  'True
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   1980
               Visible         =   0   'False
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
               Left            =   120
               TabIndex        =   119
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
               Left            =   1710
               Locked          =   -1  'True
               TabIndex        =   114
               TabStop         =   0   'False
               Top             =   840
               Width           =   1095
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
               Left            =   11880
               TabIndex        =   29
               Top             =   1980
               Visible         =   0   'False
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
               Height          =   360
               Left            =   12840
               Locked          =   -1  'True
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   420
               Width           =   1560
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
               Height          =   360
               Left            =   11280
               Locked          =   -1  'True
               TabIndex        =   101
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
               Height          =   360
               Left            =   9720
               Locked          =   -1  'True
               TabIndex        =   95
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
               ItemData        =   "FrmDeposit.frx":0121
               Left            =   7740
               List            =   "FrmDeposit.frx":0123
               Sorted          =   -1  'True
               TabIndex        =   30
               Top             =   900
               Width           =   6660
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
               Height          =   360
               Left            =   6840
               TabIndex        =   27
               Top             =   420
               Width           =   1305
            End
            Begin VB.TextBox TxtGCTDGrossWt 
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
               Left            =   8160
               TabIndex        =   28
               Top             =   420
               Width           =   1575
            End
            Begin VB.CommandButton CmdGCTDAdd2List 
               Caption         =   "Add to List"
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
               Left            =   11280
               TabIndex        =   32
               Top             =   1380
               Width           =   1455
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
               Left            =   1965
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   1320
               Width           =   7290
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
               ItemData        =   "FrmDeposit.frx":0125
               Left            =   120
               List            =   "FrmDeposit.frx":0127
               Sorted          =   -1  'True
               TabIndex        =   23
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
               ItemData        =   "FrmDeposit.frx":0129
               Left            =   1440
               List            =   "FrmDeposit.frx":012B
               Sorted          =   -1  'True
               TabIndex        =   24
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
               ItemData        =   "FrmDeposit.frx":012D
               Left            =   2640
               List            =   "FrmDeposit.frx":012F
               Sorted          =   -1  'True
               TabIndex        =   25
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label62 
               Caption         =   "Average Bag Size"
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
               Left            =   9480
               TabIndex        =   180
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label56 
               Caption         =   "Moisture %"
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
               TabIndex        =   172
               Top             =   900
               Width           =   1095
            End
            Begin VB.Label Label58 
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
               Left            =   4800
               TabIndex        =   167
               Top             =   120
               Width           =   1095
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
               Left            =   13680
               TabIndex        =   122
               Top             =   1680
               Visible         =   0   'False
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
               Left            =   12120
               TabIndex        =   120
               Top             =   1560
               Visible         =   0   'False
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
               Left            =   555
               TabIndex        =   113
               Top             =   900
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
               TabIndex        =   107
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
               TabIndex        =   102
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
               Left            =   9870
               TabIndex        =   96
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
               TabIndex        =   87
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
               Left            =   1560
               TabIndex        =   86
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
               TabIndex        =   85
               Top             =   120
               Width           =   735
            End
            Begin VB.Label Label23 
               Caption         =   "No Of Bags"
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
               Left            =   6900
               TabIndex        =   84
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label24 
               Caption         =   "Gross Weight"
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
               Left            =   8265
               TabIndex        =   83
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
               Height          =   255
               Left            =   5160
               TabIndex        =   82
               Top             =   960
               Width           =   2535
            End
            Begin VB.Label Label25 
               Caption         =   "Quality Parameters / Remarks"
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
               TabIndex        =   81
               Top             =   1335
               Width           =   2055
            End
         End
         Begin VB.PictureBox Pic_Client 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   -74880
            ScaleHeight     =   1905
            ScaleWidth      =   14505
            TabIndex        =   56
            Top             =   450
            Width           =   14535
            Begin VB.TextBox TxtCGSTNo 
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
               Left            =   1170
               Locked          =   -1  'True
               TabIndex        =   185
               TabStop         =   0   'False
               Top             =   1440
               Width           =   3015
            End
            Begin VB.TextBox TxtCAddress 
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
               Height          =   975
               Left            =   5265
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   183
               TabStop         =   0   'False
               Top             =   900
               Width           =   9135
            End
            Begin VB.TextBox TxtCNAVCode 
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
               Left            =   1170
               Locked          =   -1  'True
               TabIndex        =   181
               TabStop         =   0   'False
               Top             =   960
               Width           =   3015
            End
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
               TabIndex        =   42
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
               TabIndex        =   118
               TabStop         =   0   'False
               Top             =   120
               Width           =   375
            End
            Begin VB.ComboBox CmbClientDP 
               Height          =   315
               Left            =   5160
               Sorted          =   -1  'True
               TabIndex        =   41
               Top             =   480
               Visible         =   0   'False
               Width           =   5055
            End
            Begin VB.CommandButton CmdClientAdd2List 
               Caption         =   "Add to List"
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
               Left            =   13200
               TabIndex        =   10
               Top             =   480
               Width           =   1215
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
               TabIndex        =   61
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
               TabIndex        =   60
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
               TabIndex        =   59
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
               TabIndex        =   58
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
               TabIndex        =   9
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
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   480
               Width           =   3495
            End
            Begin VB.Label Label65 
               Caption         =   "GST No"
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
               Left            =   120
               TabIndex        =   186
               Top             =   1500
               Width           =   1095
            End
            Begin VB.Label Label64 
               Caption         =   "Address"
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
               Left            =   4320
               TabIndex        =   184
               Top             =   1260
               Width           =   1095
            End
            Begin VB.Label Label63 
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
               Left            =   120
               TabIndex        =   182
               Top             =   1020
               Width           =   1095
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               TabIndex        =   62
               Top             =   180
               Width           =   975
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   3600
            Left            =   90
            TabIndex        =   146
            Top             =   2415
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   6350
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3540
            Left            =   -74895
            TabIndex        =   148
            Top             =   2460
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   6244
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   149
            Top             =   2355
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   6376
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
            TabIndex        =   145
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
            TabIndex        =   144
            Top             =   540
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "FrmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mLocationID, mSLocationID, mClientIDRow, mContractClientIDRow, mCommodityID, mContractID, mGodownID, mGSLID, mGCTDClientID As Variant
Dim mAgentID As Variant
Dim mClientIDRowGsl As Variant
Dim mExchangeTypeID As Double
Dim mCMPID As Double
Dim mCMPName As Double
Dim mCmbStack, mCmbLot, mTimeStamp, msSM_prefix As Variant
Dim PntPrv  As Variant
Dim wc As Variant

''Declaration for PDF File

Dim pgval As Double
Dim objPDF As New mjwPDF
Dim mLeft As Double
Dim mTop As Double
Dim mHeight As Double
Dim mRowHeight As Double
Dim mRow As Double
Dim mGlobalLeft As Double
Dim mTxtTop As Double
Dim mTxtLeft As Double
Dim mFontSize As Integer
Dim mTxtFromTop As Double
Dim mPDFFileName, mPDFFileName1 As Variant

Dim mReserveGodown As Variant
Dim mTruckWtRatio As Variant
Dim mBODDate As Date

Dim Totpages As Double
Dim mCommtrackFlag, mCMSPCode, mKYCFlag As Variant
Dim mCommodityGroupId As Variant
Dim mCalMoisturePer As Double
Dim mAdjGSLID As Variant
Dim mVehicleTypeID As Variant

Dim mCommodityYearlyFlag, mGodownYearlyFlag As Variant
Dim mBlockDays, mBlockMessage As Variant
Dim mBlockDate As Date
''-- Ullhas Patil 21 April 2018
Dim mSiloLocation As Boolean
Dim mTmp1 As Variant
''-----------------------------

Dim mBOEIDNew As Variant
Dim mInvAmountOnPrint  As Variant
Sub UpdateClientGridStock()
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 7) <> "" Then
       MSHFlexGrid2.TextMatrix(I, 5) = 0 '"No of Bags"
       MSHFlexGrid2.TextMatrix(I, 6) = 0 '"Weight"
    End If
Next
MSHFlexGrid2.TextMatrix(0, 5) = 0 '"No of Bags"
MSHFlexGrid2.TextMatrix(0, 6) = 0 '"Weight"

For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 16) <> "" Then
       Call UpdateDataClient(MSHFlexGrid4.TextMatrix(I, 16), MSHFlexGrid4.TextMatrix(I, 4), MSHFlexGrid4.TextMatrix(I, 10))
    End If
Next
End Sub

Sub UpdateDataClient(ClientRowID As Variant, NoofBags As Variant, EWeight As Variant)
For I = 2 To MSHFlexGrid2.Rows - 1
    If Val(MSHFlexGrid2.TextMatrix(I, 7)) = Val(ClientRowID) Then
       MSHFlexGrid2.TextMatrix(I, 5) = Val(MSHFlexGrid2.TextMatrix(I, 5)) + NoofBags '"No of Bags"
       MSHFlexGrid2.TextMatrix(I, 6) = Val(MSHFlexGrid2.TextMatrix(I, 6)) + EWeight '"Weight"
       
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
Else
   CmbContractID.Text = ""
End If
CmdClientAdd2List.Enabled = True
Call Grid_Head_GCTD
Call Grid_Head_Client

End Sub

Private Sub ChkSCMFlag_Click()
Call Grid_Head_Client
Call Grid_Head_GCTD
If CmbContractID.Enabled = True Then
   CmbContractID.Text = ""
End If
If CmbNonNcdexClient.Visible = True Then
   CmbNonNcdexClient.Text = ""
End If

End Sub

Private Sub CmbAgent_GotFocus()
tmp = CmbAgent.Text
Call TableMst_Agent
CmbAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Agent.RecordCount
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
mClientIDRow = TmpRs1!ClientID
mCMSPCode = IIf(IsNull(TmpRs1!CMSPCode), "", TmpRs1!CMSPCode)
mKYCFlag = IIf(IsNull(TmpRs1!KYCFlag), "", TmpRs1!KYCFlag)

If Check1.Value = 0 Then
   TxtCNAVCode = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "GNAV")
Else
   TxtCNAVCode = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "RGNAV")
End If
TxtCAddress = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "a")
TxtCGSTNo = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "g")

If CmdClientAdd2List.Enabled = True Then CmdClientAdd2List.SetFocus

End Sub



Private Sub CmbGCTDClient_GotFocus()
tmp = CmbGCTDClient.Text
CmbGCTDClient.Clear
If MSHFlexGrid2.Rows < 4 Then Exit Sub

For I = 2 To MSHFlexGrid2.Rows - 2
   CmbGCTDClient.AddItem MSHFlexGrid2.TextMatrix(I, 2) & " ~ " & MSHFlexGrid2.TextMatrix(I, 4)
Next I
CmbGCTDClient.Text = tmp
End Sub


Private Sub CmbGCTDClient_LostFocus()
If CmbGCTDClient.Text = "" Then Exit Sub

For I = 2 To MSHFlexGrid2.Rows - 2
    
   If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(I, 2) & " ~ " & MSHFlexGrid2.TextMatrix(I, 4) Then
      mGCTDClientID = Val(MSHFlexGrid2.TextMatrix(I, 7))
      Exit Sub
   End If
Next I

MsgBox "Invalid selection "
CmbGCTDClient.SetFocus

End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPID.SetFocus
   Exit Sub
End If

If Check1.Value = 1 Then
   If CmbContractID.Text = "" Then
      MsgBox "Select Contract!!!"
      CmbContractID.SetFocus
      Exit Sub
   End If
End If
If CmbCommodityID.Text <> "" Then
    tmp = CmbGodownNo.Text
    'Call TableMst_Godown(" Where CloseDate is null And CMPID = " & mCMPID)
    '----Added on 30 Nov 2011 411 PM Parag For Reservation Deposit
    If Check1.Value = 0 Then
       Call TableMst_Godown(" Where CloseDate is null And CMPID = " & mCMPID)
    Else
       If mReserveGodown <> "" Then
          Call TableMst_Godown(" Where CloseDate is null And CMPID = " & mCMPID & " And GodownID in  " & mReserveGodown & "  ")
       Else
          MsgBox "Please map Godown to reservation contract !!!"
          Exit Sub
       End If
    End If
    CmbGodownNo.Clear
    If RsMst_Godown.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To RsMst_Godown.RecordCount
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
mGodownYearlyFlag = IIf(IsNull(RsMst_Godown!YearlyBillingFlg), 0, RsMst_Godown!YearlyBillingFlg)
TxtGodAdd = IIf(IsNull(RsMst_Godown!Address), "", RsMst_Godown!Address)

If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
   For I = 2 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(I, 21) = "" Then Exit For
       If Val(MSHFlexGrid4.TextMatrix(I, 21)) <> Val(mGodownYearlyFlag) Then
          MsgBox "All godown in which deposit is going should have same yearly billing flag!!!"
          CmbGodownNo.SetFocus
          Exit Sub
       End If
   Next
End If


End Sub



Private Sub CmbLot_GotFocus()
If CmbStack.Text = "" Then
  MsgBox "Please Select Stack No !!!"
  CmbStack.SetFocus
  Exit Sub
End If
If CmbCommodityID.Text = "" Then
  MsgBox "Please Select Commodity !!!"
  CmbLot.Text = ""
  CmbCommodityID.SetFocus
  Exit Sub
End If
tmp = CmbLot.Text
Call TableMst_GSL(" Where SpillageFlag <> '1' And  MadeUpFlag <> '1' And  " & Gen_ExpireGSLFlag & "  And  CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "'")
CmbLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If

For I = 1 To RsMst_GSL.RecordCount
    CmbLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next I



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

LotBagStock = Val(LotBagStock) + getDepositedGSLBag(RsMst_GSL!Mst_GSL_ID)
TxtLotStock = Format(Val(TxtLotStock) + getDepositedGSLWt(RsMst_GSL!Mst_GSL_ID), "#####0.000")
If Val(TxtLotStock) = 0 And LotBagStock = 0 Then
   mGSLID = RsMst_GSL!Mst_GSL_ID
   mClientIDRowGsl = Null
   Exit Sub
End If

If RsMst_GSL!CommodityID <> mCommodityID Then
   MsgBox "GSL Contain Stock of other Commodity!!! "
   CmbLot.SetFocus
   Exit Sub
End If
If RsMst_GSL!ExchangeTypeID <> mExchangeTypeID Then
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


mGSLID = RsMst_GSL!Mst_GSL_ID
mClientIDRowGsl = RsMst_GSL!ClientID

End Sub

'Private Sub CmbNonNcdexClient_GotFocus()
'tmp = CmbNonNcdexClient.Text
'Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
'CmbNonNcdexClient.Clear
'If RsMst_Client.RecordCount = 0 Then
'   MsgBox "No Record found !!!! "
'   Exit Sub
'End If
'For I = 1 To RsMst_Client.RecordCount
'    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
'    RsMst_Client.MoveNext
'Next
'CmbNonNcdexClient.Text = tmp
'End Sub

Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text

If ChkSCMFlag.Value = 0 Then
   wc = "SCMFlag = 0 "
Else
   wc = "SCMFlag = 1 "
End If

If Edit_Flg = "A" Then
   Call TableMst_Client("where " & wc & " And Status = 'A' and ExchangeTypeId = " & mExchangeTypeID & " And Converted = 0 ")
ElseIf Edit_Flg = "E" Then
   Call TableMst_Client("where " & wc & " And (Status = 'A' or ClientID in (select distinct ClientID from Txn_CxTF_ClientDetail where CxTF_DetailsID = " & TxtCxTFNo.Text & " and SM_Prefix = '" & TxtPreFix.Text & "' And TxnType = 'D')) and ExchangeTypeId = " & mExchangeTypeID & " And Converted = 0  ")
End If

CmbNonNcdexClient.Clear

If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
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
mClientIDRow = RsMst_Client!ClientID
mCMSPCode = IIf(IsNull(RsMst_Client!CMSPCode), "", RsMst_Client!CMSPCode)
mKYCFlag = IIf(IsNull(RsMst_Client!KYCFlag), "", RsMst_Client!KYCFlag)

Dim ans As Variant

ans = GetClientNameByRow(RsMst_Client!ClientID)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)

If Check1.Value = 0 Then
   TxtCNAVCode = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "GNAV")
Else
   TxtCNAVCode = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "RGNAV")
End If
TxtCAddress = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "a")
TxtCGSTNo = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "g")



End Sub

Private Sub CmbSSCMFlag_LostFocus()
If CmbSSCMFlag.Text = "" Then Exit Sub

If LCase(CmbSSCMFlag.Text) <> "yes" And LCase(CmbSSCMFlag.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSSCMFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbStack_GotFocus()

If CmbGodownNo.Text = "" Or mGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  Exit Sub
End If
tmp = CmbStack.Text
Call TableMst_GSL(" Where SpillageFlag <> '1' And MadeUpFlag <> '1' And CMPID = " & mCMPID & " And GodownID = " & mGodownID, "D")
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
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = TxtClientID And MSHFlexGrid2.TextMatrix(u, 3) = TxtDPID Then
          If LCase(CmdClientAdd2List.Caption) = "add to list" Then
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

CheckDuplicateClient = Retval
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

If mExchangeTypeID <> 1 Then
   If mCommtrackFlag = 1 Then
      If mCMSPCode = "" Then
         MsgBox "CMS-P Code not mapped to client. Can't select this client for deposit!!!"
         Exit Sub
      End If
   End If
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
MSHFlexGrid2.TextMatrix(r, 8) = mKYCFlag

MSHFlexGrid2.TextMatrix(r, 9) = TxtCNAVCode
MSHFlexGrid2.TextMatrix(r, 10) = TxtCAddress
MSHFlexGrid2.TextMatrix(r, 11) = TxtCGSTNo


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
mKYCFlag = 0
TxtCAddress = ""
TxtCNAVCode = ""
TxtCGSTNo = ""
CmdClientAdd2List.Caption = "Add to List"
End Sub

Function CheckDuplicateGSL() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid4.Rows > 3 Then
   For u = 1 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(u, 1) = CmbGodownNo.Text And _
          MSHFlexGrid4.TextMatrix(u, 2) = CmbStack.Text And _
          MSHFlexGrid4.TextMatrix(u, 3) = CmbLot.Text And MSHFlexGrid4.TextMatrix(u, 19) = CmbGSLTruckNo.Text Then
          If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
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

CheckDuplicateGSL = Retval
End Function


Private Sub CmdGCTDeleteList_Click()
I = MSHFlexGrid4.RowSel
If I > 1 Then
   If MSHFlexGrid4.TextMatrix(I, 1) <> "" Then
      MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(I, 4))
      MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(I, 5))
      MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) - Val(MSHFlexGrid4.TextMatrix(I, 6))
      MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) - Val(MSHFlexGrid4.TextMatrix(I, 7))
      MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) - Val(MSHFlexGrid4.TextMatrix(I, 8))
      MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) - Val(MSHFlexGrid4.TextMatrix(I, 9))
      MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) - Val(MSHFlexGrid4.TextMatrix(I, 10))
      For RO = I To MSHFlexGrid4.Rows - 2
          For C = 0 To MSHFlexGrid4.Cols - 1
               MSHFlexGrid4.TextMatrix(RO, C) = MSHFlexGrid4.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 1
   End If
End If
Call ClearGCTDTabControl
End Sub

Private Sub CmdGCTDAdd2List_Click()
Dim r As Variant
Dim mstr As Variant


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
   MsgBox "Duplicate G.S.L. & Truck No.  Not Allowed !!! "
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

If CmbGSLTruckNo.Text = "" Then
   MsgBox "Blank Truck No. not allowed!!! "
   CmbGSLTruckNo.SetFocus
   Exit Sub
End If

If CheckPreviousGSLClient = True Then
   MsgBox "Client Name should be same as previous of this GSL!!! "
   CmbGCTDClient.SetFocus
   Exit Sub
End If

If Val(TxtCLotSize) > 0 Then
   tmp = 100 + Val(TxtCVariation)
   tmp = (tmp / 100) * Val(TxtCLotSize)
   If Val(TxtGCTDDematWt) + (Val(TxtLotStock)) >= tmp Then
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

mBag = getRemainingTruckBag(CmbGSLTruckNo.Text)
'If getRemainingTruckBag(CmbGSLTruckNo.Text) < Val(TxtGCTDNoofBags.Text) Then
If Val(mBag) < Val(TxtGCTDNoofBags.Text) Then
   MsgBox "Deposited number of bags are more than available number of bags of selected truck!!!"
   TxtGCTDNoofBags.SetFocus
   Exit Sub
End If

mWt = Format(getRemainingTruckWt(CmbGSLTruckNo.Text), "#######0.000")

'If getRemainingTruckWt(CmbGSLTruckNo.Text) < Val(TxtGCTDGrossWt.Text) Then
If Val(mWt) < Val(TxtGCTDGrossWt.Text) Then
   MsgBox "Deposited gross weight is more than available gross weight of selected truck!!!"
   TxtGCTDGrossWt.SetFocus
   Exit Sub
End If

If mExchangeTypeID = 2 Then
   mstr = Check7DaysInvoice(mGSLID)
   If mstr <> "" Then
      MsgBox "7 days invoice (Invoice No : SD" & mstr & " ) is generated for selected GSL, You can't deposite in this GSL!!"
      CmbLot.SetFocus
      Exit Sub
   End If
End If

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
MSHFlexGrid4.TextMatrix(r, 19) = CmbGSLTruckNo.Text
MSHFlexGrid4.TextMatrix(r, 20) = TxtMoisturePer1.Text
MSHFlexGrid4.TextMatrix(r, 21) = Val(mGodownYearlyFlag)
MSHFlexGrid4.TextMatrix(r, 22) = Val(TxtAverageBagSize)

MSHFlexGrid4.Col = 5
MSHFlexGrid4.row = r
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
CmbGSLTruckNo.Text = ""
CmbNonNcdexClient.Text = ""
TxtLotStock = ""
TxtMoisturePer1.Text = ""
TxtAverageBagSize = ""
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
Private Sub UpdateBillingDetails()

tmp = "Select CxTF_DetailsID,SM_Prefix,CxTFDate,CDTFDate,NoOfDay,ClientID,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,NoOfBags,DematWeight,BagSize,NContractID,BillingCycleID,DepBillingCycleID,BagSizeFlag ,FinalPrice,DepositFinalPrice,G_UID,DepConDID"
tmp = tmp & " From (Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
tmp = tmp & " From  Txn_CxTF_GSL TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((TCG.DematWeight/TCG.NoOfBags)*1000) Between  MinBagSize And MaxBagSize"
tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.CxTFDate Between StartDate and EndDate And ContractType = 'N'"
tmp = tmp & " Where TC.SM_Prefix = '" & TxtPreFix & "' And TC.CxTF_DetailsID = " & Val(TxtCxTFNo) & " And TC.CxTFDate = '" & CDate(TxtCxTFDate.Value) & "'  And TC.ExchangeTypeID=1 And  TC.TxnType='D'   And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y' " 'And TCG.DepConDID is  null"
tmp = tmp & " Union All"
tmp = tmp & " Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
tmp = tmp & " From  Txn_CxTF_GSL TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.GodownID = TDCD.GodownId"
tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And  TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.CxTFDate Between StartDate and EndDate And ContractType = 'N'"
tmp = tmp & " Where TC.SM_Prefix = '" & TxtPreFix & "' And TC.CxTF_DetailsID = " & Val(TxtCxTFNo) & " And TC.CxTFDate = '" & CDate(TxtCxTFDate.Value) & "' And TC.ExchangeTypeID=1 And  TC.TxnType='D'   And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'N' )A" ' And TCG.DepConDID is null) A" '
tmp = tmp & " Order By SM_Prefix,CxTF_DetailsID,Mst_GSL_ID"

Call TmpTable


For I = 1 To TmpRs.RecordCount
    mDepoBillingCycleID = ""
    mDepoBillingUnit = ""
    mNoofDays = ""
    mDepConDID = ""



    mDepoBillingCycleID = IIf(IsNull(TmpRs!DepBillingCycleID), "", TmpRs!DepBillingCycleID)
    If TmpRs!BagSizeFlag = "Y" Then
       mDepoBillingUnit = "B"
    Else
       mDepoBillingUnit = "Q"
    End If
    mDepoBillingRate = IIf(IsNull(TmpRs!DepositFinalPrice), 0, TmpRs!DepositFinalPrice)
    mDepConDID = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)

    If mDepoBillingCycleID = 1 Then
       mNoofDays = a
       mFromDate = Year(TmpRs!CxTFDate) & "-" & Month(TmpRs!CxTFDate) & "-01"
       mSMonth = Month(TmpRs!CxTFDate)
       mSYear = Year(TmpRs!CxTFDate)
       mTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
       mNoofDays = (Val(mTotalDays) - Val(Day(TmpRs!CxTFDate))) + 1
    ElseIf mDepoBillingCycleID = 2 Then
       If Day(TmpRs!CxTFDate) < 8 Then
          mNoofDays = "4"
       ElseIf Day(TmpRs!CxTFDate) < 15 Then
          mNoofDays = "3"
       ElseIf Day(TmpRs!CxTFDate) < 22 Then
          mNoofDays = "2"
       Else
          mNoofDays = "1"
       End If
    ElseIf mDepoBillingCycleID = 3 Then
       If Day(TmpRs!CxTFDate) < 16 Then
          mNoofDays = "2"
       Else
          mNoofDays = "1"
       End If
    ElseIf mDepoBillingCycleID = 4 Then
       mNoofDays = "1"
    End If

    If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofDays <> "" Then
       tmp = "Update Txn_CxTF_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID = '" & TmpRs!CxTF_DetailsID & "' And TxnType='D'  And G_UID = '" & TmpRs!G_UID & "' " 'And DepConDID is null
       Call Tmp1Table
    
       'tmp = "Update Txn_CxTF_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID = '" & TmpRs!CxTF_DetailsID & "' And TxnType='D' And DepConDID is null And G_UID = '" & TmpRs!G_UID & "' "
       'Call Tmp1Table

'       tmp = "Update Txn_InvMonGSLDetail Set DRNo =" & mDepConDID & ",BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",NooFDays = " & mNoofDays & " Where YearMonth = '2020-12-01' And CxTF_DetailsID ='" & TmpRs!CxTF_DetailsID & "' And Tflag='D' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       Call Tmp1Table
'
'       If LCase(mDepoBillingUnit) = "q" Then
'          tmp = "Update Txn_InvMonGSLDetail Set InvoiceAmount = BalanceWeight * NoOfDays * Rate,InvoiceDate =  '2020-12-31'  Where YearMonth = '2020-12-01' And ExchangeTypeID = 1 And BillingUnit = 'Q' And (ContractID = 0 Or ContractID is Null)  And TFLag = 'D' And InvoiceAmount is null and SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID ='" & TmpRs!CxTF_DetailsID & "'"
'          Call Tmp1Table
'       Else
'          tmp = "Update Txn_InvMonGSLDetail Set InvoiceAmount = BalanceBags * NoOfDays * Rate,InvoiceDate =  '2020-12-31'  Where YearMonth = '2020-12-01' And ExchangeTypeID = 1 And BillingUnit = 'B' And (ContractID = 0 Or ContractID is Null)  And TFLag = 'D' And InvoiceAmount is null and SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID ='" & TmpRs!CxTF_DetailsID & "'"
'          Call Tmp1Table
'       End If
    End If

    TmpRs.MoveNext
Next


tmp = " Select SM_Prefix,ClientID,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,BalanceBags,BalanceWeight,BagSize,ContractID,NContractID,BillingCycleID,WBillingCycleID,BagSizeFlag ,FinalPrice,WithdrawalFinalPrice,G_UID,DepConDID,CDTFDate,DepConID,GSLDCDID"
tmp = tmp & " From (Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID 'GSLDCDID'"
tmp = tmp & " From  Mst_GSL GSL"
tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') And GSL.CDTFDate Between StartDate and EndDate And ContractType = 'N'"
tmp = tmp & " Where GSL.SM_Prefix = '" & TxtPreFix & "' And GSL.CxTF_DetailsID = " & TxtCxTFNo & " And GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1   And GSL.SpillageFlag = 0  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'" 'And GSL.DepConDID is null
tmp = tmp & " Union All"
tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID,GSL.DepConDID 'GSLDCDID'"
tmp = tmp & " From  Mst_GSL GSL"
tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And GSL.CDTFDate Between StartDate and EndDate And ContractType = 'N'"
tmp = tmp & " Where GSL.SM_Prefix = '" & TxtPreFix & "' And GSL.CxTF_DetailsID = " & TxtCxTFNo & " And GSL.DepositeBags > 0 And GSL.BalanceBags > 0  And GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.SpillageFlag = 0  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N' " 'And  GSL.DepConDID is null
tmp = tmp & " Union All "
tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID 'GSLDCDID'"
tmp = tmp & " From  Mst_GSL GSL"
tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
tmp = tmp & " Left Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') And GSL.CDTFDate Between StartDate and EndDate And ContractType = 'N' "
tmp = tmp & " Where GSL.SM_Prefix = '" & TxtPreFix & "' And GSL.CxTF_DetailsID = " & TxtCxTFNo & " And GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1   And GSL.SpillageFlag = 0  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y' And Convert(Varchar,GSL.Mst_GSL_ID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.Mst_GSL_ID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where a.TxnType = 'W' And b.CxTFDate >= '2017-12-01') " 'And GSL.DepConDID is null
tmp = tmp & " Union All"
tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID,GSL.DepConDID 'GSLDCDID'"
tmp = tmp & " From  Mst_GSL GSL"
tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And GSL.CDTFDate Between StartDate and EndDate And ContractType = 'N' "
tmp = tmp & " Where GSL.SM_Prefix = '" & TxtPreFix & "' And GSL.CxTF_DetailsID = " & TxtCxTFNo & " And GSL.DepositeBags > 0 And GSL.BalanceBags = 0  And GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1 And   GSL.SpillageFlag = 0  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N' And Convert(Varchar,GSL.Mst_GSL_ID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.Mst_GSL_ID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where a.TxnType = 'W' And b.CxTFDate >= '2017-12-01')) A" 'GSL.DepConDID is null And
tmp = tmp & " Order By SM_Prefix,Mst_GSL_ID"

Call TmpTable


For I = 1 To TmpRs.RecordCount
    mDepoBillingCycleID = ""
    mDepoBillingUnit = ""
    mNoofDays = ""
    mDepConDID = ""
    mDepConID = ""
    mDepoBillingCycleID = IIf(IsNull(TmpRs!BillingCycleID), "", TmpRs!BillingCycleID)
    If TmpRs!BagSizeFlag = "Y" Then
       mDepoBillingUnit = "B"
    Else
       mDepoBillingUnit = "Q"
    End If
    mDepoBillingRate = IIf(IsNull(TmpRs!FinalPrice), 0, TmpRs!FinalPrice)
    mDepConDID = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)
    mDepConID = IIf(IsNull(TmpRs!DepConID), "", TmpRs!DepConID)
    
    If mDepoBillingCycleID = 1 Then
       mNoofDays = "0"
    ElseIf mDepoBillingCycleID = 2 Then
       mNoofDays = "4"
    ElseIf mDepoBillingCycleID = 3 Then
       mNoofDays = "2"
    ElseIf mDepoBillingCycleID = 4 Then
       mNoofDays = "1"
    End If
    
    If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofDays <> "" Then
       tmp = "Update Mst_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where ExchangeTypeID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " " 'And DepConDID is null "
       Call Tmp1Table
    End If
    
    TmpRs.MoveNext
Next


End Sub
Private Sub CmdPrint_Click()


If mExchangeTypeID = 2 Then
   If mCommtrackFlag = 1 Then
      If ChkClientKYCFlag = False Then
         MsgBox "One or more client(s) have not submitted KYC-P forms to CMG. Printout can be taken only after submission!!!"
         Exit Sub
      End If
   End If
End If

If mExchangeTypeID = 1 Then
   If ChkOneGSLMultiCDTF = True Then
      'MsgBox "One or more client(s) have not submitted KYC-P forms to CMG. Printout can be taken only after submission!!!"
      MsgBox "Multiple CDTF for GSL Marked in Red. Please delete this CDTF!!!"
      Exit Sub
   End If
End If

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

If CheckRateCards = False Then
   MsgBox "Storage Rate Card is not updated for this location and commodity in rate card master. Can't take print till rate card updated in Store-Man. Contact CMG department for same !!!"
   Exit Sub
End If


Check2.Visible = False

If LCase(UserType) = "power" Then
   Check2.Visible = True
End If

If mExchangeTypeID = 2 Then
   If mCommtrackFlag = 1 Then
      ChkRecAnnexure.Visible = True
      ChkRecAnnexure.Value = 0
   End If
End If


If RsTxn_CxTF_Details!HologramNo = "" Or IsNull(RsTxn_CxTF_Details!HologramNo) Then
   
   tmp = "Select Distinct TCD.CxTFDate,TCD.CxTF_DetailsID,TCD.Flag  From Txn_CxTF_Details TCD"
   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTF_DetailsID = TCG.CxTF_DetailsID And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType "
   tmp = tmp & " Where TCD.TxnType = 'D' and TCD.ExchangeTypeID = " & mExchangeTypeID & " And TCD.SM_Prefix='" & TxtPreFix & "' And TCD.CxTF_DetailsID<" & Val(TxtCxTFNo) & " And TCG.Mst_GSL_ID in "
   tmp = tmp & " (Select Mst_GSL_ID From Txn_CxTF_GSL TCG Where TCG.TxnType = 'D' And TCG.SM_Prefix='" & TxtPreFix & "' And TCG.CxTF_DetailsID=" & Val(TxtCxTFNo) & ") And TCD.Flag is Null "
   tmp = tmp & " order by TCD.CxTF_DetailsID Desc"
    
   Call TmpTable
   mstr = ""
   If TmpRs.RecordCount > 0 Then
      For I = 1 To TmpRs.RecordCount
          If mstr = "" Then
             mstr = TmpRs!CxTF_DetailsID
          Else
             mstr = mstr & "," & TmpRs!CxTF_DetailsID
          End If
          TmpRs.MoveNext
      Next
      MsgBox " Previous CDTF No pending for printing " & mstr
      Exit Sub
   End If

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

'RsTxn_CxTF_Details!Flag = "P"
'RsTxn_CxTF_Details.Update

Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
'Call Xls_Print_Rpt

PntPrv = "N"

'Call PDF_Print_Rpt
'''
'Call Gen_Email

Call SearchCmd_Click

End Sub

Private Sub CmdPrintPreview_Click()

'Commented after UAT by Mayur Sir on 6th April 2013
'If mExchangeTypeID = 2 Then
'   If mCommtrackFlag = 1 Then
'      If ChkClientKYCFlag = False Then
'         MsgBox "One or more client(s) have not submitted KYC-P forms to CMG. Printout can be taken only after submission!!!"
'         Exit Sub
'      End If
'   End If
'End If

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

'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
'Call Xls_Print_Rpt

PntPrv = "Y"

Call PDF_Print_Rpt

End Sub
Private Function ChkClientKYCFlag() As Boolean

Dim Retval As Boolean
Retval = True

For k = 2 To MSHFlexGrid2.Rows
    If MSHFlexGrid2.TextMatrix(k, 7) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(k, 8) = 0 Then
       Retval = False
       Exit For
    End If
Next
ChkClientKYCFlag = Retval

End Function

Private Sub CmdSearch_Click()

Dim wc As String
Call Grid_Head
wc = ""
If STo.Value < SFrom.Value Then
   MsgBox "To date must be always greater than from date!!!"
   STo.SetFocus
   Exit Sub
End If

If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, mCMPName, "Mst_CMP.CMPID", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp = Replace(Gen_WcCMP, "CMPID", "mst_CMP.CMPID")
      wc = GenWc(wc, tmp, "", "I")
   End If
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "Mst_Commodity.CommodityID", "N")
End If
If CmbSExchangeType.Text <> "" Then
   wc = GenWc(wc, mExchangeTypeID, "Mst_ExchangeType.ExchangeTypeID", "N")
End If
If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, msSM_prefix, "Txn_CxTf_Details.sm_prefix", "S")
End If

If TxtSCxTFNo <> "" Then
   wc = GenWc(wc, TxtSCxTFNo, "Txn_CxTf_Details.CxTF_DetailsID", "N")
End If

If TxtSManualCxTFNo <> "" Then
   wc = GenWc(wc, TxtSManualCxTFNo, "Txn_CxTf_Details.ManualCDTFNo", "S")
End If


If TxtSTruckNo <> "" Then
   'Wc = GenWc(Wc, TxtSTruckNo, "Txn_CxTF_TruckDetail.TruckNo", "S")
   
   wc = wc & " And Txn_CxTF_TruckDetail.TruckNo like '%" & TxtSTruckNo & "%' "
   
End If

If CmbSSCMFlag.Text <> "" Then
   If LCase(CmbSSCMFlag.Text) = "yes" Then
      wc = wc & " And isnull(Txn_CxTf_Details.SCMFlag,0) = '1' "
   ElseIf LCase(CmbSSCMFlag.Text) = "no" Then
      wc = wc & " And isnull(Txn_CxTf_Details.SCMFlag,0) = '0' "
   End If
End If


tmp = "Select Txn_CxTf_Details.CxTF_DetailsID,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,Txn_CxTf_Details.SM_Prefix,"
tmp = tmp & " Mst_Commodity.Commodity,Txn_CxTf_Details.Flag,"
tmp = tmp & " Txn_CxTf_Details.Agent,Txn_CxTf_Details.Remark,Txn_CxTf_Details.HologramNo,"
tmp = tmp & " Mst_Agent.AgentName,MC.ClientName+'~'+MC.ClientIDRow 'ClientName' , Txn_CxTf_Details.ManualCDTFNo, "
tmp = tmp & " Sum (Txn_CxTf_Gsl.NoofBags) 'NoOfBags',sum(Txn_CxTf_Gsl.GrossWeight) 'GrossWeight',Txn_CxTf_Details.G_UID,Txn_CxTf_Details.RecAnnexure,Txn_CxTf_Details.SCMFlag,Txn_CxTf_Details.CreatedBy,Txn_CxTf_Details.CreatedDate "
tmp = tmp & " From Txn_CxTf_Details"
tmp = tmp & " Inner Join Txn_CxTf_Gsl  on Txn_CxTf_Details.SM_Prefix = Txn_CxTf_Gsl.SM_Prefix And"
tmp = tmp & " Txn_CxTf_Details.CxTF_DetailsID = Txn_CxTf_Gsl.CxTF_DetailsID And Txn_CxTf_Details.TxnType = Txn_CxTf_Gsl.TxnType"
tmp = tmp & " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID"
tmp = tmp & " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
tmp = tmp & " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID"
tmp = tmp & " Left Join Mst_GSL MGSL ON Txn_CxTf_Gsl.Mst_GSL_ID = MGSL.Mst_GSL_ID And Txn_CxTf_Gsl.SM_Prefix = MGSL.SM_Prefix "
tmp = tmp & " Left Join Mst_Client MC ON MGSL.ClientID = MC.ClientID"

If TxtSTruckNo <> "" Then
   tmp = tmp & " Inner Join Txn_CxTF_TruckDetail  on Txn_CxTf_Details.SM_Prefix = Txn_CxTF_TruckDetail.SM_Prefix And"
   tmp = tmp & " Txn_CxTf_Details.CxTF_DetailsID = Txn_CxTF_TruckDetail.CxTF_DetailsID And Txn_CxTf_Details.TxnType = Txn_CxTF_TruckDetail.TxnType"
End If

tmp = tmp & " Where Txn_CxTf_Details.TxnType = 'D' "

If Gen_DBType = "MDB" Then
   tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
    tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If wc <> "" Then
   tmp = tmp & wc
End If
tmp = tmp & " group by Txn_CxTf_Details.CxTF_DetailsID,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,"
tmp = tmp & " Txn_CxTf_Details.SM_Prefix, Mst_Commodity.Commodity,Txn_CxTf_Details.Flag, Txn_CxTf_Details.Agent,"
tmp = tmp & " Txn_CxTf_Details.Remark,Txn_CxTf_Details.HologramNo, Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo,"
tmp = tmp & " Txn_CxTf_Details.G_UID,Txn_CxTf_Details.RecAnnexure,Txn_CxTf_Details.SCMFlag,Txn_CxTf_Details.CreatedBy,Txn_CxTf_Details.CreatedDate,MC.ClientName,MC.ClientIDRow "
tmp = tmp & " ORDER BY Txn_CxTf_Details.sm_prefix,Mst_CMP.CMPName,Txn_CxTf_Details.CxTF_DetailsID "

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    For C = 0 To TmpRs.Fields.Count - 1
        'MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    
        If InStr(TmpRs.Fields(C), "/") > 0 Then
           If IsDate(TmpRs.Fields(C)) Then
              MSHFlexGrid1.TextMatrix(I, C) = Format(TmpRs.Fields(C), "dd-mmm-yyyy")
           Else
              MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C)) 'TmpRs.Fields(c)
           End If
        Else
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
        End If
    Next
    If MSHFlexGrid1.TextMatrix(I, 16) = "1" Then
       MSHFlexGrid1.TextMatrix(I, 16) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(I, 16) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(I, 17) = "1" Then
       MSHFlexGrid1.TextMatrix(I, 17) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(I, 17) = "No"
    End If
    TmpRs.MoveNext
Next

RsTxn_CxTF_Details.Close
wc = ""
If Gen_WcCMP <> "" Then
   wc = "And " & Gen_WcCMP
End If

Call TableTxn_CxTF_Details(" Where TxnType = 'D' " & wc & "Order By SM_Prefix,CMPID,CxTF_DetailsID ")

End Sub
Function GenWc(wc_ As String, ID As Variant, Field_ As String, Type_ As String) As String
Dim Retval
Retval = wc_
If Retval = "" Then
   If Type_ = "S" Then
      Retval = "And " & Field_ & "= '" & ID & "'"
   ElseIf Type_ = "I" Then
      Retval = "And " & Field_ & " " & ID
   Else
      Retval = "And " & Field_ & "= " & ID
   End If
Else
   If Type_ = "S" Then
      Retval = Retval & " And " & Field_ & "= '" & ID & "'"
   ElseIf Type_ = "I" Then
      Retval = Retval & " And " & Field_ & " " & ID
   Else
      Retval = Retval & " And " & Field_ & "= " & ID
   End If
End If
GenWc = Retval
End Function

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub



Private Sub TxtCxTFDate_LostFocus()
If TxtCxTFDate.Value > Date Then
   MsgBox "Future date is not allowed!!! "
   TxtCxTFDate.SetFocus
End If
If TxtCxTFDate.Value < mBODDate Then
   MsgBox mBlockMessage '"CDTF Date less than BOD Date is not allowed!!! "
   TxtCxTFDate.SetFocus
End If

TxtCxTFDispalyDate = TxtCxTFDate.Value
If Check1.Value = 1 Then
   CmbContractID.Text = ""
End If
End Sub

''--- Ullhas Patil 21 APril 2018
Private Sub TxtForeignMatter_LostFocus()
If TxtForeignMatter = "" Then Exit Sub
If Not IsNumeric(TxtForeignMatter) Then
   MsgBox "Invalid numeric format"
   TxtForeignMatter.SetFocus
End If
Call CalTruckWt
End Sub
Private Sub CalForeignMatter()
TxtTTDGrossWt = Format(Val(TxtTTDTotalWt) - (Val(TxtTTDTareWt) + Val(TxtForeignMatter)), "#####0.000")
End Sub
''-----------------------------

Private Sub TxtHologramNo_LostFocus()
'If TxtHologramNo = "" Then Exit Sub
'If Check_HologramNo(TxtHologramNo) = False Then
'   TxtHologramNo.SetFocus
'   CmdSaveHologram.Enabled = False
'   Exit Sub
'End If
'CmdSaveHologram.Enabled = True
'CmdSaveHologram.SetFocus


If TxtHologramNo = "" Then Exit Sub
'If Check_HologramNo(TxtHologramNo) = False Then
'   TxtHologramNo.SetFocus
'   CmdSaveHologram.Enabled = False
'   Exit Sub
'End If


If TxtHologramNo <> "0" Then

   If Check_HologramNo(TxtHologramNo) = False Then
      TxtHologramNo.SetFocus
      CmdSaveHologram.Enabled = False
      Exit Sub
   End If
    
   If CheckHologramInRange(TxtHologramNo, mLocationID) = False Then
      MsgBox "Hologram No not found in any range."
      TxtHologramNo.SetFocus
      CmdSaveHologram.Enabled = False
      Exit Sub
   End If
End If

CmdSaveHologram.Enabled = True
CmdSaveHologram.SetFocus

End Sub
Private Sub CmdSaveHologram_Click()
On Error GoTo msgError
Dim ans As Variant
If TxtHologramNo = "" Then
   MsgBox "Blank Hologram No not allowed!!!"
   TxtHologramNo.SetFocus
   Exit Sub
End If

If mExchangeTypeID = 2 Then
   'Commented on 06 April 2013 after UAT and Mail from Sandeep Mohite
   'If mCommtrackFlag = 1 Then
      If ChkRecAnnexure.Value = 0 Then
         MsgBox "Please confirm having received Annexure-III by clicking the check box!. You cannot take print out without receiving Annesure-III"
         ChkRecAnnexure.SetFocus
         Exit Sub
      End If
   'End If
End If

ans = MsgBox("Are you Sure you want to save this HologramNo", vbYesNo)
If ans = vbYes Then
   If mExchangeTypeID = 2 Then
      If TxtCxTFDate.Value >= CDate("02/18/2012") Then
         tmp = "Select Collect7Days from Mst_Location where LocationID = " & mLocationID
         Call Tmp2Table
         If TmpRs2.RecordCount > 0 Then
            If TmpRs2!Collect7Days = True Then
            
               tmp = "Select GodownID,Rate from Mst_NCDEXStorageRateCardDetail MNRC"
               tmp = tmp & " inner Join Mst_NCDEXStorageRateCard MNR on MNRC.NSRCID=MNR.NSRCID"
               tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MNR.CommodityGroupID = MCG.CommodityGroupID"
               tmp = tmp & " Inner Join Mst_Commodity MComm On MCG.CommodityGroupID = MComm.CommodityGroupID"
               tmp = tmp & " Where MComm.CommodityID=" & mCommodityID & "  And MNR.LocationID=" & mLocationID
                
               Call Tmp2Table
              
               For I = 2 To MSHFlexGrid4.Rows - 1
                   If MSHFlexGrid4.TextMatrix(I, 13) <> "" Then
                      If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
                         TmpRs2.Find "GodownID=" & Val(MSHFlexGrid4.TextMatrix(I, 15))
                         If TmpRs2.EOF = False Then
                            mRate = TmpRs2!Rate
                         Else
                            mRate = 0
                         End If
                         If Val(mRate) = 0 Then
                            MsgBox "No Rate updated in NCDEX rate card master. Please contact CMG Dept for updating NCDEX rate card master!!!" 'for Location " & GetLocationName(mLocationID) & " And Commodity Group " & GetCommodityGroupName(  & " "
                            TxtHologramNo.SetFocus
                            Exit Sub
                         End If
                   End If
               Next
            End If
         End If
      End If
   End If
   If TxtHologramNo <> "0" Then
      Call TableMst_Hologram
      RsMst_Hologram.AddNew
      RsMst_Hologram!NewHologramNo = TxtHologramNo
      RsMst_Hologram!DocType = "CDTF"
      RsMst_Hologram!DocNo = TxtCxTFNo
      RsMst_Hologram!SM_Prefix = TxtPreFix
      RsMst_Hologram!CreatedBy = Gen_UserLoginID
      RsMst_Hologram.Update
      RsMst_Hologram.Close
   End If
   
   RsTxn_CxTF_Details!HologramNo = TxtHologramNo
   RsTxn_CxTF_Details!Flag = "P"
   RsTxn_CxTF_Details!RecAnnexure = ChkRecAnnexure.Value
   RsTxn_CxTF_Details.Update
   RsTxn_CxTF_Details.Requery
   RsTxn_CxTF_Details.MoveFirst
   RsTxn_CxTF_Details.Find "G_UID = '" & TxtG_UID & "'"
   
   TxtHologramNo = ""
   Frame3.Visible = False
   RsTxn_CxTF_Details.Update
   Gen_mTimeStamp = GetTimeStamp
   If TxtCxTFDate.Value >= CDate("02/18/2012") Then
      Call SaveCDTFBOE
   End If
   If Check2.Value = 1 Then
'      Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
'      Call Xls_Print_Rpt
       PntPrv = "N"
       Call PDF_Print_Rpt
       Call Gen_Email
   End If

Else
   TxtHologramNo.SetFocus
End If

Exit Sub

msgError:

MsgBox Err.Description, vbExclamation, "Error" '"Some error occurred during process!!!"


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
'Public Sub Xls_Print_Rpt()
'Dim oXL As Excel.Application
'Dim oWB As Excel.Workbook
'Dim oWS As Excel.Worksheet
'Dim mRow As Integer
''Dim mCol As Integer
'Dim ans As Variant
'MsgBox ("Wait till next message")
'Set oXL = New Excel.Application
''Dim Pat As String
'Pat = App.Path
''Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
'
'Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPrefix) & "" & Trim(TxtCxTFNo) & ".xls")
'
'I = 2
'Set oWS = oWB.Worksheets("Sheet" & 1)
''oWS.PageSetup.RightMargin = 0.05
''oWS.PageSetup.TopMargin = 0.05
''oWS.PageSetup.RightMargin = 0.05
''oWS.PageSetup.TopMargin = 4.6
'oWS.Name = TxtPrefix & "-" & TxtCxTFNo  'MSHFlexGrid2.TextMatrix(i, 23)
'oWS.ClearArrows
'mRow = 4
'
'oWS.Range("a1:g1").Merge
'oWS.Range("a:a").ColumnWidth = 10.86
'oWS.Range("b:b").ColumnWidth = 12.86
'oWS.Range("c:c").ColumnWidth = 12.86
'oWS.Range("d:d").ColumnWidth = 12.86
'oWS.Range("e:e").ColumnWidth = 12.86 '32.14
'oWS.Range("f:f").ColumnWidth = 12.86
'oWS.Range("g:g").ColumnWidth = 12.86
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = GetCMPAddress(RsTxn_CxTF_Details!CMPID, "D") '& "-" & GetCMPName(RsTxn_CxTF_Details!CMPID) ' & "-" & GetCMPAddress(RsTxn_CxTF_Details!CMPID)
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 10
''oWS.Cells(mRow, c + 1).RowHeight = 24
'mRow = mRow + 1
'
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Commodity Deposit Transaction Form (CDTF)"
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 10
'oWS.Cells(mRow, c + 1).RowHeight = 24
'
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "(This is NOT an instrument of title to the goods)"
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 7
'
'
'mRow = mRow + 1
'oWS.Cells(mRow, c + 1) = "CDTF No"
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).WrapText = True
''oWS.Range("b5:c5").Merge
'tmp = "b" & mRow & ":" & "d" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 2) = TxtPrefix & "-" & TxtCxTFNo
'oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
'oWS.Cells(mRow, c + 2).Font.Size = 7
'
'
'oWS.Cells(mRow, c + 5) = "Date"
'oWS.Cells(mRow, c + 5).Font.Size = 7
'oWS.Cells(mRow, c + 5).Font.Bold = True
'
''oWS.Range(tmp).Merge
'tmp = "f" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 6) = Format(TxtCxTFDate, "dd-mmm-yyyy")
'oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
'oWS.Cells(mRow, c + 6).Font.Size = 7
'
'mRow = mRow + 1
'
'oWS.Cells(mRow, c + 1) = "Location"
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).Font.Bold = True
'tmp = "b" & mRow & ":" & "d" & mRow
'
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 2) = GetLocationFromCMP(RsTxn_CxTF_Details!CMPID)
'oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
'oWS.Cells(mRow, c + 2).Font.Size = 7
'
'oWS.Cells(mRow, c + 5) = "State"
'oWS.Cells(mRow, c + 5).Font.Size = 7
'oWS.Cells(mRow, c + 5).Font.Bold = True
'
'tmp = "f" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 6) = GetStateFromLocation(GetLocationIDFromCMP(RsTxn_CxTF_Details!CMPID))
'oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
'oWS.Cells(mRow, c + 6).Font.Size = 7
'
'mRow = mRow + 1
'
'oWS.Cells(mRow, c + 1) = "CMP ID & Address"
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).WrapText = True
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).RowHeight = 40
'
'tmp = "b" & mRow & ":" & "d" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 2) = RsTxn_CxTF_Details!CMPID & "-" & GetCMPAddress(RsTxn_CxTF_Details!CMPID, "A")
'oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
'oWS.Cells(mRow, c + 2).Font.Size = 7
'oWS.Cells(mRow, c + 2).WrapText = True
'
'
'oWS.Cells(mRow, c + 5) = "Commodity"
'oWS.Cells(mRow, c + 5).Font.Bold = True
'oWS.Cells(mRow, c + 5).Font.Size = 7
'
'tmp = "f" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 6) = TxtCommodityID
'oWS.Cells(mRow, c + 6).Font.Size = 7
'
'
'mRow = mRow + 1
'
'
'oWS.Cells(mRow, c + 1) = "Transaction Type"
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).WrapText = True
'
'tmp = "b" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 2) = TxtExchangeTypeID
'oWS.Cells(mRow, c + 2).Font.Size = 7
'
'mRow = mRow + 1
'
'oWS.Cells(mRow, c + 1) = "Name Of Agent"
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 7
'
'tmp = "b" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 2) = TxtAgent
'oWS.Cells(mRow, c + 2).Font.Size = 7
'
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Depositor Details"
'oWS.Cells(mRow, c + 1).Font.Size = 9
'oWS.Cells(mRow, c + 1).Font.Bold = True
'
'mRow = mRow + 1
'
'oWS.Cells(mRow, c + 1) = "Client-ID"
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 7
'tmp = "b" & mRow & ":" & "c" & mRow
'oWS.Range(tmp).Merge
'
'oWS.Cells(mRow, c + 2) = "Client Name"
'oWS.Cells(mRow, c + 2).Font.Bold = True
'oWS.Cells(mRow, c + 2).Font.Size = 7
'
'If TxtExchangeTypeID = "NCDEX" Then
'   oWS.Cells(mRow, c + 4) = "DP ID"
'   oWS.Cells(mRow, c + 5) = "DP Name"
'ElseIf TxtExchangeTypeID = "Non NCDEX" Then
'   tmp = "b" & mRow & ":" & "e" & mRow
'   oWS.Range(tmp).Merge
'ElseIf TxtExchangeTypeID = "SPOT" Then
'   oWS.Cells(mRow, c + 4) = "TM ID"
'   oWS.Cells(mRow, c + 5) = "TM Name"
'End If
'oWS.Cells(mRow, c + 4).Font.Bold = True
'oWS.Cells(mRow, c + 4).Font.Size = 7
'oWS.Cells(mRow, c + 5).Font.Bold = True
'oWS.Cells(mRow, c + 5).Font.Size = 7
'
'oWS.Cells(mRow, c + 6) = "No of Bags"
'oWS.Cells(mRow, c + 6).Font.Bold = True
'oWS.Cells(mRow, c + 6).Font.Size = 7
'oWS.Cells(mRow, c + 7) = "Electronic Weight in MT"
'oWS.Cells(mRow, c + 7).Font.Bold = True
'oWS.Cells(mRow, c + 7).Font.Size = 7
'oWS.Cells(mRow, c + 7).RowHeight = 24
'
'mRow = mRow + 1
'
'For k = 2 To MSHFlexGrid2.Rows - 2
''k = 1
'    oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(k, 1)
'    oWS.Cells(mRow, c + 1).Font.Size = 7
'    tmp = "b" & mRow & ":" & "c" & mRow
'    oWS.Range(tmp).Merge
'    oWS.Cells(mRow, c + 2) = MSHFlexGrid2.TextMatrix(k, 2)
'    oWS.Cells(mRow, c + 2).Font.Size = 7
'    If TxtExchangeTypeID <> "Non NCDEX" Then
'       oWS.Cells(mRow, c + 4) = MSHFlexGrid2.TextMatrix(k, 3)
'       oWS.Cells(mRow, c + 4).Font.Size = 7
'       oWS.Cells(mRow, c + 5) = MSHFlexGrid2.TextMatrix(k, 4)
'       oWS.Cells(mRow, c + 5).Font.Size = 7
'    Else
'       tmp = "b" & mRow & ":" & "e" & mRow
'       oWS.Range(tmp).Merge
'    End If
'    oWS.Cells(mRow, c + 6) = MSHFlexGrid2.TextMatrix(k, 5)
'    oWS.Cells(mRow, c + 6).Font.Size = 7
'    oWS.Cells(mRow, c + 7) = Format(MSHFlexGrid2.TextMatrix(k, 6), "#######0.000")
'    'oWS.Cells(mRow, c + 7).FormatCell = "##0.000"
'
'    oWS.Cells(mRow, c + 7).Font.Size = 7
'    mRow = mRow + 1
'Next k
'
'If TxtExchangeTypeID <> "Non NCDEX" Then
'   tmp = "b" & mRow & ":" & "c" & mRow
'Else
'   tmp = "b" & mRow & ":" & "e" & mRow
'End If
'oWS.Range(tmp).Merge
'
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "e" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Total"
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'
'oWS.Cells(mRow, c + 6) = MSHFlexGrid2.TextMatrix(0, 5)
'oWS.Cells(mRow, c + 6).Font.Size = 7
'oWS.Cells(mRow, c + 7) = MSHFlexGrid2.TextMatrix(0, 6)
'oWS.Cells(mRow, c + 7).Font.Size = 7
'
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Truck-wise Deposit Details"
'oWS.Cells(mRow, c + 1).Font.Size = 9
'oWS.Cells(mRow, c + 1).Font.Bold = True
'
'mRow = mRow + 1
'
'oWS.Cells(mRow, c + 1) = "Truck No"
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 2) = "Total Weight(MT)"
'oWS.Cells(mRow, c + 2).Font.Bold = True
'oWS.Cells(mRow, c + 2).Font.Size = 7
'oWS.Cells(mRow, c + 3) = "Tare Weight(MT)"
'oWS.Cells(mRow, c + 3).Font.Bold = True
'oWS.Cells(mRow, c + 3).Font.Size = 7
'oWS.Cells(mRow, c + 4) = "Gross Weight"
'oWS.Cells(mRow, c + 4).Font.Bold = True
'oWS.Cells(mRow, c + 4).Font.Size = 7
'oWS.Cells(mRow, c + 5) = "Weighbridge Slip No"
'oWS.Cells(mRow, c + 5).Font.Bold = True
'oWS.Cells(mRow, c + 5).Font.Size = 7
'oWS.Cells(mRow, c + 5).RowHeight = 25
'tmp = "f" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 6) = "Remarks"
'oWS.Cells(mRow, c + 6).Font.Bold = True
'oWS.Cells(mRow, c + 6).Font.Size = 7
'
'mRow = mRow + 1
'
'For k = 2 To MSHFlexGrid3.Rows - 2
''k = 1
'    oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(k, 1)
'    oWS.Cells(mRow, c + 1).Font.Size = 7
'    oWS.Cells(mRow, c + 2) = Format(MSHFlexGrid3.TextMatrix(k, 3), ".000")
'    oWS.Cells(mRow, c + 2) = Format(MSHFlexGrid3.TextMatrix(k, 3), "#######0.000")
'    oWS.Cells(mRow, c + 2).Font.Size = 7
'    oWS.Cells(mRow, c + 3) = Format(MSHFlexGrid3.TextMatrix(k, 4), "#######0.000")
'    oWS.Cells(mRow, c + 3).Font.Size = 7
'    oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid3.TextMatrix(k, 5), "#######0.000")
'    oWS.Cells(mRow, c + 4).Font.Size = 7
'    oWS.Cells(mRow, c + 5) = MSHFlexGrid3.TextMatrix(k, 2)
'    oWS.Cells(mRow, c + 5).Font.Size = 7
'    tmp = "f" & mRow & ":" & "g" & mRow
'    oWS.Range(tmp).Merge
'    oWS.Cells(mRow, c + 6) = MSHFlexGrid3.TextMatrix(k, 6)
'    oWS.Cells(mRow, c + 6).Font.Size = 7
'    mRow = mRow + 1
'Next k
'
'tmp = "f" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'
'mRow = mRow + 1
'
'
'
'tmp = "a" & mRow & ":" & "c" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Total"
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'
'oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid3.TextMatrix(0, 5), "#######0.000")
'oWS.Cells(mRow, c + 4).Font.Size = 7
'
'tmp = "e" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'
'
'oWS.Cells(mRow, c + 5) = "Affix Hologram Here"
'oWS.Cells(mRow, c + 5).Font.Size = 7
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'
'tmp = "e" & mRow & ":" & "e" & (mRow + 1)
'oWS.Range(tmp).Merge
'
'mRow = mRow + 1
''mCol = 5
'
'tmp = "e" & mRow & ":" & "e" & (mRow + 1)
'oWS.Range(tmp).Merge
'
'tmp = "a" & mRow & ":" & "c" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Less Gunny Weight(MT)"
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'
'oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 8), "#######0.000")
'oWS.Cells(mRow, c + 4).Font.Size = 7
'
'tmp = "e" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'
'tmp = "e" & mRow & ":" & "e" & (mRow + 1)
'oWS.Range(tmp).Merge
'
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "c" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Less Standard Deductions(MT)"
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'
'oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 9), "#######0.000")
'oWS.Cells(mRow, c + 4).Font.Size = 7
'
'tmp = "e" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'
'tmp = "e" & mRow & ":" & "e" & (mRow + 1)
'oWS.Range(tmp).Merge
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "c" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Add Adjustments(MT)"
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'
'oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 6), "#######0.000")
'oWS.Cells(mRow, c + 4).Font.Size = 7
'
'tmp = "e" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'
'
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "c" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Electronic Weight(MT)"
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'
'oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 10), "#######0.000")
'
'oWS.Cells(mRow, c + 4).Font.Size = 7
'
'
'oWS.Cells(mRow, c + 5) = "Hologram No"
'oWS.Cells(mRow, c + 5).Font.Size = 7
'oWS.Cells(mRow, c + 5).Font.Bold = True
'
'tmp = "f" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'If RsTxn_CxTF_Details!HologramNo = "" Or IsNull(RsTxn_CxTF_Details!HologramNo) Then
'   oWS.Cells(mRow, c + 6) = "Print Preview !!!! "
'Else
'   oWS.Cells(mRow, c + 6) = RsTxn_CxTF_Details!HologramNo
'End If
'oWS.Cells(mRow, c + 6).Font.Size = 7
'oWS.Cells(mRow, c + 6).HorizontalAlignment = -4108
'
'mRow = mRow + 1
'mRow = mRow + 1
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Name & Signature of CMP Manager/CMP DRC:"
'oWS.Cells(mRow, c + 1).Font.Size = 9
'oWS.Cells(mRow, c + 1).RowHeight = 40
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'
'Dim SDesc As String
'
'SDesc = "I/We hereby declare that the commodities being deposited herein belong " & _
'        "to the clients as stated above and are free from any encumbrances. " & _
'        "I/we request you to take custody of the said commodities and I/we undertake to " & _
'        "indemnify you against any losses incurred by you due to any false information provided by me/us." & _
'        "I/we also undertake to pay all required charges as per your charges schedule from time to time. " & _
'        "I have gone through the attached 'Terms and Conditions'** of storage and agree to abide by the same. "
'If mExchangeTypeID = 2 Then
'   SDesc = SDesc & " I/We hereby request the warehouse to credit the above commodities to the account of the beneficiaries " & _
'           "stated above after receiving the assayer's certificates for the above commodities  and after deducting " & _
'           "standard deductions and  gunny bags weight as per Commodity Exchange norm."
'           ' Added on 22.03.2010
'   SDesc = SDesc & " We understand that NCMSL may not enter the deposit details in NCDEX system for the purpose of demat unless all charges, including pre-demat storage charges and assaying charges are paid in full"
'
'End If
'oWS.Cells(mRow, c + 1) = SDesc
'oWS.Cells(mRow, c + 1).Font.Size = 6
'oWS.Cells(mRow, c + 1).WrapText = True
'oWS.Cells(mRow, c + 1).RowHeight = 50
'
'mRow = mRow + 1
'mRow = mRow + 1
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "** 'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html "
'oWS.Cells(mRow, c + 1).Font.Size = 8
'oWS.Cells(mRow, c + 1).WrapText = True
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).RowHeight = 24
'oWS.Cells(mRow, c + 1).Font.Name = "Zurich BT"
'
'mRow = mRow + 1
'mRow = mRow + 1
'mRow = mRow + 1
'mRow = mRow + 1
'
'
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Name & Signature of Depositor/Agent of Depositor:"
'oWS.Cells(mRow, c + 1).Font.Size = 9
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Name = "Zurich BT"
'
'k = mRow - 6
'
'For Cell = 4 To k
'    tmp = "a" & Cell & ":" & "g" & Cell
'    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
'    oWS.Range(tmp).Font.Name = "Zurich BT"
'    'oWS.Range(tmp).Font.Size = 7
'    oWS.Range(tmp).WrapText = True
'Next Cell
'
'If Gen_ExcelPassword <> "" Then
'   oWS.Protect Gen_ExcelPassword, , , , , , True, True
'End If
'
'
'
'
'' --------------------- GSL Detail
'
'
'Set oWS = oWB.Worksheets("Sheet" & 2)
''oWS.PageSetup.RightMargin = 0.05
''oWS.PageSetup.TopMargin = 0.05
''oWS.PageSetup.RightMargin = 0.05
''oWS.PageSetup.TopMargin = 4.6
'oWS.Name = TxtPrefix & "-" & TxtCxTFNo & " - GCTD Detail"  'MSHFlexGrid2.TextMatrix(i, 23)
'oWS.ClearArrows
'mRow = 1
'
'oWS.Range("a1:g1").Merge
'oWS.Range("a:a").ColumnWidth = 10.86
'oWS.Range("b:b").ColumnWidth = 12.86
'oWS.Range("c:c").ColumnWidth = 12.86
'oWS.Range("d:d").ColumnWidth = 12.86
'oWS.Range("e:e").ColumnWidth = 12.86 '32.14
'oWS.Range("f:f").ColumnWidth = 12.86
'oWS.Range("g:g").ColumnWidth = 12.86
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = GetCMPAddress(RsTxn_CxTF_Details!CMPID, "D") '& "-" & GetCMPName(RsTxn_CxTF_Details!CMPID) ' & "-" & GetCMPAddress(RsTxn_CxTF_Details!CMPID)
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 10
''oWS.Cells(mRow, c + 1).RowHeight = 24
'mRow = mRow + 1
'
'
'tmp = "a" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "Godown Stack Lot Detail "
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 10
'oWS.Cells(mRow, c + 1).RowHeight = 24
'
''mRow = mRow + 1
''tmp = "a" & mRow & ":" & "g" & mRow
''oWS.Range(tmp).Merge
''oWS.Cells(mRow, c + 1) = "(This is NOT an instrument of title to the goods)"
''oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
''oWS.Cells(mRow, c + 1).Font.Bold = True
''oWS.Cells(mRow, c + 1).Font.Size = 7
'
'
'mRow = mRow + 1
'oWS.Cells(mRow, c + 1) = "CDTF No"
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).WrapText = True
''oWS.Range("b5:c5").Merge
'tmp = "b" & mRow & ":" & "d" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 2) = TxtPrefix & "-" & TxtCxTFNo
'oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
'oWS.Cells(mRow, c + 2).Font.Size = 7
'
'
'oWS.Cells(mRow, c + 5) = "Date"
'oWS.Cells(mRow, c + 5).Font.Size = 7
'oWS.Cells(mRow, c + 5).Font.Bold = True
'
''oWS.Range(tmp).Merge
'tmp = "f" & mRow & ":" & "g" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 6) = Format(TxtCxTFDate, "dd-mmm-yyyy")
'oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
'oWS.Cells(mRow, c + 6).Font.Size = 7
'
'mRow = mRow + 1
'oWS.Cells(mRow, c + 1) = "Godown No"
'oWS.Cells(mRow, c + 1).Font.Size = 7: oWS.Cells(mRow, c + 1).Font.Bold = True
'
'oWS.Cells(mRow, c + 2) = "Godown Address"
'oWS.Cells(mRow, c + 2).Font.Size = 7: oWS.Cells(mRow, c + 2).Font.Bold = True
'
'oWS.Cells(mRow, c + 3) = "Stack No "
'oWS.Cells(mRow, c + 3).Font.Size = 7: oWS.Cells(mRow, c + 3).Font.Bold = True
'
'oWS.Cells(mRow, c + 4) = "Lot No"
'oWS.Cells(mRow, c + 4).Font.Size = 7: oWS.Cells(mRow, c + 4).Font.Bold = True
'
'oWS.Cells(mRow, c + 5) = "No of Bags "
'oWS.Cells(mRow, c + 5).Font.Size = 7: oWS.Cells(mRow, c + 5).Font.Bold = True
'oWS.Cells(mRow, c + 6) = "Gross Weight"
'oWS.Cells(mRow, c + 6).Font.Size = 7: oWS.Cells(mRow, c + 6).Font.Bold = True
'oWS.Cells(mRow, c + 7) = "client Name"
'oWS.Cells(mRow, c + 7).Font.Size = 7: oWS.Cells(mRow, c + 7).Font.Bold = True
'
'mRow = mRow + 1
'
'For k = 2 To MSHFlexGrid4.Rows - 2
''k = 1
'    oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid4.TextMatrix(k, 1)
'    oWS.Cells(mRow, c + 1).Font.Size = 7
'    oWS.Cells(mRow, c + 2) = "'" & MSHFlexGrid4.TextMatrix(k, 18)
'    oWS.Cells(mRow, c + 2).Font.Size = 7
'    oWS.Cells(mRow, c + 3) = "'" & MSHFlexGrid4.TextMatrix(k, 2)
'    oWS.Cells(mRow, c + 3).Font.Size = 7
'    oWS.Cells(mRow, c + 4) = (MSHFlexGrid4.TextMatrix(k, 3))
'    oWS.Cells(mRow, c + 4).Font.Size = 7
'    oWS.Cells(mRow, c + 5) = MSHFlexGrid4.TextMatrix(k, 4)
'    oWS.Cells(mRow, c + 5).Font.Size = 7
''    tmp = "f" & mRow & ":" & "g" & mRow
''    oWS.Range(tmp).Merge
'    oWS.Cells(mRow, c + 6) = MSHFlexGrid4.TextMatrix(k, 5)
'    oWS.Cells(mRow, c + 6).Font.Size = 7
'    oWS.Cells(mRow, c + 7) = MSHFlexGrid4.TextMatrix(k, 11)
'    oWS.Cells(mRow, c + 7).Font.Size = 7
'
'    mRow = mRow + 1
'Next k
'k = mRow
'
'For Cell = 1 To k
'    tmp = "a" & Cell & ":" & "g" & Cell
'    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
'    oWS.Range(tmp).Font.Name = "Zurich BT"
'    'oWS.Range(tmp).Font.Size = 7
'    oWS.Range(tmp).WrapText = True
'Next Cell
'
'If Gen_ExcelPassword <> "" Then
'   oWS.Protect Gen_ExcelPassword, , , , , , True, True
'End If
''-------- End GSL Detail
'
'oWB.Save
'
'Call Log_Print_History(TxtPrefix, "CDTF", TxtCxTFNo, Gen_FLName)
'
'
'
''oWB.Close
''oXL.Quit
'MsgBox ("Excel file created !!! ")
'oXL.Visible = True
'
'End Sub




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


For I = 2 To MSHFlexGrid3.Rows - 1
    TmpRs.AddNew
    TmpRs!ID = TxtCxTFNo
    TmpRs!F1 = MSHFlexGrid3.TextMatrix(I, 1)
    TmpRs!F2 = MSHFlexGrid3.TextMatrix(I, 3)
    TmpRs!F3 = MSHFlexGrid3.TextMatrix(I, 4)
    TmpRs!F4 = MSHFlexGrid3.TextMatrix(I, 5)
    TmpRs!F5 = MSHFlexGrid3.TextMatrix(I, 2)
    TmpRs!F6 = MSHFlexGrid3.TextMatrix(I, 6)
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

If MsgBox("This will clear GSL details of selected Truck, You wan't to continue?", vbYesNo) = vbNo Then
   Call ClearTTDTabControl
   Exit Sub
End If
      
I = MSHFlexGrid3.RowSel
If I > 1 Then
   If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
      Call DeleteFromGSLDetail
      Call UpdateClientGridStock
      Call ClearGCTDTabControl
      MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) - Val(MSHFlexGrid3.TextMatrix(I, 3))
      MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) - Val(MSHFlexGrid3.TextMatrix(I, 4))
      MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid3.TextMatrix(I, 5))
      MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) - Val(MSHFlexGrid3.TextMatrix(I, 6))
      MSHFlexGrid3.TextMatrix(0, 13) = Val(MSHFlexGrid3.TextMatrix(0, 13)) - Val(MSHFlexGrid3.TextMatrix(I, 13)) ''-- Ullhas Patil 21 April 2018
      
      For RO = I To MSHFlexGrid3.Rows - 2
          For C = 0 To MSHFlexGrid3.Cols - 1
               MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
          Next
      Next
        

      MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
   End If
End If
Call ClearTTDTabControl
End Sub

Private Sub CmdTTDAdd2List_Click()
Dim r As Variant

If CmbVehicleTypeID.Text = "" Then
   MsgBox "Blank Vehicle Model not allowed!!!"
   CmbVehicleTypeID.SetFocus
   Exit Sub
End If

If TxtTTDTruckNo = "" Then
   MsgBox "Blank Truck No !!! "
   TxtTTDTruckNo.SetFocus
   Exit Sub
End If

If TxtWeighBridgNo = "" Then
   MsgBox "Blank WeighBridge No !!! "
   TxtWeighBridgNo.SetFocus
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

''--- Ullhas Patil 21 April 2018
If mSiloLocation = True Then
    If Val(TxtForeignMatter) <= 0 Then
     MsgBox "Invalid Foreign Matter !!! "
     TxtForeignMatter.SetFocus
     Exit Sub
    End If
End If
''-------------------------------

If Val(TxtTTDNoOfBags) <= 0 Then
   MsgBox "Invalid No of Bags !!! "
   TxtTTDNoOfBags.SetFocus
   Exit Sub
End If

'Commented on 26 Feb 2015 mail from Sandeep Mohite
'If mExchangeTypeID = 2 Then
   If TxtMoisturePer.Text = "" Then
      MsgBox "Blank Moisture % Not Allowed !!!"
      TxtMoisturePer.SetFocus
      Exit Sub
   End If
'End If

If CheckDuplicateTruck() = True Then
   MsgBox "Duplicate Truck No. and Weighbridge Slip No. not Allowed!!! "
   TxtTTDTruckNo.SetFocus
   Exit Sub
End If

If Val(TxtTTDTareWt) < Val(TxtMinimum) Or Val(TxtTTDTareWt) > Val(TxtMaximum) Then
   MsgBox "The Tare weight of the Truck appears to be beyond the limits set for this Vehicle Model. Please make an enquiry and ensure that there is no fraudulent activity taking place before going ahead with this data entry!!" '" Tare weight must be between " & Val(TxtMinimum) & " and " & Val(TxtMaximum) & "  range of Vehicle Model Master !!!", vbExclamation, "Message"
   'TxtTTDTareWt.SetFocus
   'Exit Sub
End If

'tmp = Format((Val(TxtTTDTotalWt) - Val(TxtTTDTareWt)) - Val(TxtTTDGrossWt), "#####0.000")
'If Val(tmp) <> 0 Then
'   MsgBox "Invalid Weight !!! "
'   Exit Sub
'End If

If CmdTTDAdd2List.Caption = "Update List" Then
   If MsgBox("This will clear GSL details of selected Truck, You wan't to continue?", vbYesNo) = vbNo Then
      Call ClearTTDTabControl
      Exit Sub
   End If
   r = MSHFlexGrid3.RowSel
   MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) - Val(MSHFlexGrid3.TextMatrix(r, 3))
   MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) - Val(MSHFlexGrid3.TextMatrix(r, 4))
   MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid3.TextMatrix(r, 5))
   MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) - Val(MSHFlexGrid3.TextMatrix(r, 6))
   MSHFlexGrid3.TextMatrix(0, 13) = Val(MSHFlexGrid3.TextMatrix(0, 13)) - Val(MSHFlexGrid3.TextMatrix(r, 13)) ''-- Ullhas Patil 21 April 2018
   Call DeleteFromGSLDetail
   Call UpdateClientGridStock
   Call ClearGCTDTabControl
Else
  r = MSHFlexGrid3.Rows - 1
  MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If
MSHFlexGrid3.TextMatrix(r, 1) = TxtTTDTruckNo
MSHFlexGrid3.TextMatrix(r, 2) = TxtWeighBridgNo
MSHFlexGrid3.TextMatrix(r, 3) = TxtTTDTotalWt
MSHFlexGrid3.TextMatrix(r, 4) = TxtTTDTareWt
MSHFlexGrid3.TextMatrix(r, 5) = TxtTTDGrossWt
MSHFlexGrid3.TextMatrix(r, 6) = Val(TxtTTDNoOfBags)
MSHFlexGrid3.TextMatrix(r, 7) = txtTTDRemarks
MSHFlexGrid3.TextMatrix(r, 8) = TxtMoisturePer.Text
MSHFlexGrid3.TextMatrix(r, 9) = CmbVehicleTypeID.Text
MSHFlexGrid3.TextMatrix(r, 10) = mVehicleTypeID
MSHFlexGrid3.TextMatrix(r, 11) = Val(TxtMinimum)
MSHFlexGrid3.TextMatrix(r, 12) = Val(TxtMaximum)
MSHFlexGrid3.TextMatrix(r, 13) = TxtForeignMatter  ''-- Ullhas Patil 21 April 2018

MSHFlexGrid3.Col = 5
MSHFlexGrid3.row = r
MSHFlexGrid3.CellBackColor = &HFFC0C0

MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(r, 3))
MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(r, 4))
MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + Val(MSHFlexGrid3.TextMatrix(r, 5))
MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(r, 6))
MSHFlexGrid3.TextMatrix(0, 13) = Val(MSHFlexGrid3.TextMatrix(0, 13)) + Val(MSHFlexGrid3.TextMatrix(r, 13))  ''-- Ullhas Patil 21 April 2018

Call ClearTTDTabControl
'TxtTTDTruckNo.SetFocus
CmbVehicleTypeID.SetFocus
End Sub

Sub ClearTTDTabControl()
CmbVehicleTypeID.Text = ""
TxtMinimum = ""
TxtMaximum = ""
TxtTTDTruckNo = ""
TxtWeighBridgNo = ""
TxtTTDTotalWt = ""
TxtTTDTareWt = ""
TxtTTDGrossWt = ""
TxtTTDNoOfBags = ""
txtTTDRemarks = ""
TxtMoisturePer.Text = ""
mVehicleTypeID = ""
CmdTTDAdd2List.Caption = "Add to List"
CmdTTDDeleteList.Enabled = False
TxtTTDGrossWt.Locked = True
''--- Ullhas Patil 21 April 2018
If mSiloLocation = False Then
    TxtForeignMatter = "": TxtForeignMatter.BackColor = &H80000000: TxtForeignMatter.Enabled = False
    Label18.Caption = "Gross Weight"
Else
    TxtForeignMatter = "": TxtForeignMatter.BackColor = &H80000005: TxtForeignMatter.Enabled = True
    Label18.Caption = "Net Weight"
End If
''------------------------------
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

wc = ""
If Gen_WcCMP <> "" Then
   wc = "And " & Gen_WcCMP
End If
Call Grid_Head
CmbSCmp.Text = ""
CmbSCommodity.Text = ""
CmbSExchangeType.Text = ""
CmbSAgent.Text = ""
SFrom.Value = Date - 730
STo.Value = Date
TxtSManualCxTFNo = ""
TxtSCxTFNo = ""
CmbSSCMFlag.Text = ""

Call TableTxn_CxTF_Details(" Where TxnType = 'D' " & wc & " Order By SM_Prefix,CMPID,CxTF_DetailsID ")
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
RsTxn_CxTF_Details.MoveLast
Call Indata

End Sub
Private Sub MSHFlexGrid4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyC And Shift = vbCtrlMask Then
   With MSHFlexGrid4
        .Redraw = False              'Find range to copy
        min_row = .row:   max_row = .RowSel
        If min_row > max_row Then tmp = max_row: max_row = min_row: min_row = tmp
        min_col = .Col:   max_col = .ColSel
        If min_col > max_col Then tmp = max_col: max_col = min_col: min_col = tmp

        my_text = ""                   'Get the text
        For I = min_row To max_row
            For j = min_col To max_col
                my_text = my_text & .TextMatrix(I, j)
                If j < max_col Then my_text = my_text & vbTab
            Next j
            If I < max_row Then my_text = my_text & vbCrLf
        Next I
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

Gen_mTimeStamp = GetTimeStamp

Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_CxTF_Details.xls")
Call Xls_Detail_Rpt

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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_CxTF_Details.xls")
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
mKYCFlag = Val(MSHFlexGrid2.TextMatrix(r, 8))

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

r = MSHFlexGrid3.RowSel
If r <= 1 Or r = MSHFlexGrid3.Rows - 1 Then
   Call ClearTTDTabControl
   Exit Sub
End If
TxtTTDTruckNo = MSHFlexGrid3.TextMatrix(r, 1)
TxtWeighBridgNo = MSHFlexGrid3.TextMatrix(r, 2)
TxtTTDTotalWt = MSHFlexGrid3.TextMatrix(r, 3)
TxtTTDTareWt = MSHFlexGrid3.TextMatrix(r, 4)
TxtTTDGrossWt = MSHFlexGrid3.TextMatrix(r, 5)
TxtTTDNoOfBags = MSHFlexGrid3.TextMatrix(r, 6)
txtTTDRemarks = MSHFlexGrid3.TextMatrix(r, 7)
TxtMoisturePer = MSHFlexGrid3.TextMatrix(r, 8)
CmbVehicleTypeID.Text = MSHFlexGrid3.TextMatrix(r, 9)
mVehicleTypeID = Val(MSHFlexGrid3.TextMatrix(r, 10))
TxtMinimum = Val(MSHFlexGrid3.TextMatrix(r, 11))
TxtMaximum = Val(MSHFlexGrid3.TextMatrix(r, 12))

''-- Ullhas Patil 21 April 2018
If mSiloLocation = False Then
    TxtForeignMatter = "": TxtForeignMatter.BackColor = &H80000000: TxtForeignMatter.Enabled = False
    Label18.Caption = "Gross Weight"
Else
    TxtForeignMatter = "": TxtForeignMatter.BackColor = &H80000005: TxtForeignMatter.Enabled = True
    Label18.Caption = "Net Weight"
End If
TxtForeignMatter = MSHFlexGrid3.TextMatrix(r, 13)
''----------------------------

CmdTTDAdd2List.Caption = "Update List"
CmdTTDDeleteList.Enabled = True

End Sub

Private Sub MSHFlexGrid4_Click()
     
Call ClearGCTDTabControl
If SaveCmd.Enabled = False Then Exit Sub ' Code Added on 16 jan 10 for View mode navigation '

r = MSHFlexGrid4.RowSel
If r <= 1 Or r = MSHFlexGrid4.Rows - 1 Then
   Call ClearGCTDTabControl
   Exit Sub
End If
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
CmbGSLTruckNo.Text = MSHFlexGrid4.TextMatrix(r, 19)
Call CmbGSLTruckNo_LostFocus
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

TxtMoisturePer1.Text = MSHFlexGrid4.TextMatrix(r, 20)
TxtAverageBagSize = Val(MSHFlexGrid4.TextMatrix(r, 22))

Call CmbGodownNo_GotFocus
Call CmbGodownNo_LostFocus

Call CmbLot_GotFocus
Call CmbLot_LostFocus

CmdGCTDeleteList.Enabled = True


TxtLotStock = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight)
LotBagStock = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags)

LotBagStock = Val(LotBagStock) + getDepositedGSLBag(mGSLID)
TxtLotStock = Format(Val(TxtLotStock) + getDepositedGSLWt(mGSLID), "#####0.000")

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
RsTxn_CxTF_Details.MovePrevious
If RsTxn_CxTF_Details.BOF Then
   RsTxn_CxTF_Details.MoveFirst
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
RsTxn_CxTF_Details.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_Details.MoveLast
Call Indata
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
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 550
Frame0.Left = Frame4.Left - 120
Frame0.Top = Frame4.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1600
RsTxn_CxTF_Details.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'CmbSCmp.Text = ""
'CmbSCommodity.Text = ""
'CmbSExchangeType.Text = ""
'CmbSAgent.Text = ""
'SFrom.Value = Date - 730
'STo.Value = Date
'TxtSManualCxTFNo = ""
'TxtSCxTFNo = ""
''===========comment on 7 july 2012 for not clearing search grid=============


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
Public Sub Indata()
Edit_Flg = ""
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
TxtCxTFNo = IIf(IsNull(RsTxn_CxTF_Details!CxTF_DetailsID), "", RsTxn_CxTF_Details!CxTF_DetailsID)
TxtPreFix = IIf(IsNull(RsTxn_CxTF_Details!SM_Prefix), "", RsTxn_CxTF_Details!SM_Prefix)
TxtManualCDTFNo = IIf(IsNull(RsTxn_CxTF_Details!ManualCDTFNo), "", RsTxn_CxTF_Details!ManualCDTFNo)

TxtCxTFDate = IIf(IsNull(RsTxn_CxTF_Details!CxTFDate), "", RsTxn_CxTF_Details!CxTFDate)
TxtCMPID = GetCMPName(IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID))
mCMPID = IIf(IsNull(RsTxn_CxTF_Details!CMPID), 0, RsTxn_CxTF_Details!CMPID)
TxtCommodityID = GetCommodityName(IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID))
mCommtrackFlag = GetCommodityCommtrackFlag(IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID))
mLocationID = GetLocationIDFromCMP(mCMPID)

mCommodityID = IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID)
TxtExchangeTypeID = GetExchangeName(IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID))
mExchangeTypeID = IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID)
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
ChkSCMFlag.Value = IIf(IsNull(RsTxn_CxTF_Details!SCMFlag), 0, RsTxn_CxTF_Details!SCMFlag)
TxtCreated = IIf(IsNull(RsTxn_CxTF_Details!CreatedBy), "", RsTxn_CxTF_Details!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CxTF_Details!CreatedDate), "", RsTxn_CxTF_Details!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_CxTF_Details!UpdatedBy), "", RsTxn_CxTF_Details!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CxTF_Details!UpdatedDate), "", RsTxn_CxTF_Details!UpdatedDate)

TxtG_UID = IIf(IsNull(RsTxn_CxTF_Details!G_UID), "", RsTxn_CxTF_Details!G_UID)

Call CmbExchangeTypeID_GotFocus
Call CmbExchangeTypeID_LostFocus

'Call CmbCMPID_GotFocus
'Call CmbCMPID_LostFocus

Call TableMst_CMP(" where cmpid = " & mCMPID)
TxtCxTFDate.Value = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate) + 1
TxtCxTFDispalyDate = TxtCxTFDate.Value
TxtCxTFDate.Enabled = False
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")

''-- Ullhas Patil 21 April 2018
mSiloLocation = CBool(GetPreFix(RsMst_CMP!CMPID, "S"))
If mSiloLocation = False Then
    TxtForeignMatter = "": TxtForeignMatter.BackColor = &H80000000: TxtForeignMatter.Enabled = False
    Label18.Caption = "Gross Weight"
Else
    TxtForeignMatter = "": TxtForeignMatter.BackColor = &H80000005: TxtForeignMatter.Enabled = True
    Label18.Caption = "Net Weight"
End If
Call ClearTTDTabControl  ''-- Ullhas Patil 21 April 2018
''------------------------------

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

CmdPrintPreview.Enabled = True
If RsTxn_CxTF_Details!HologramNo <> "" Or IsNull(RsTxn_CxTF_Details!HologramNo) = False Then
   'MsgBox "Print out already generated for this CDTF . Can't preview!!!"
   'Exit Sub
   CmdPrintPreview.Enabled = False
End If

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
   CmbVehicleTypeID.SetFocus
   'TxtTTDTruckNo.SetFocus
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

Private Sub TxtMoisturePer_LostFocus()
If TxtMoisturePer.Text = "" Then Exit Sub
If Not IsNumeric(TxtMoisturePer.Text) Then
   MsgBox "Invalid Format"
   TxtMoisturePer.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtTTDNoOfBags_LostFocus()
If TxtTTDNoOfBags = "" Then Exit Sub
If IsNumeric(TxtTTDNoOfBags) = False Or Val(TxtTTDNoOfBags) < 0 Then
   MsgBox "Invalid numeric format"
   TxtTTDNoOfBags.SetFocus
   Exit Sub
End If
TxtTTDNoOfBags = Format(TxtTTDNoOfBags, "#######0")
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

'Private Sub TxtClientID_LostFocus()
'CmbClientDP.Visible = False
'If TxtClientID = "" Then Exit Sub
'If CmbExchangeTypeID.Text = "" Then Exit Sub
'
'Dim ans As Variant
'ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)
'
'If ans(0, 0) = False Then
'   MsgBox "Not Found !!!"
'   TxtClientID.SetFocus
'   Exit Sub
'End If
'If ans(3, 0) > 1 Then
'   CmbClientDP.Visible = True
'   CmbClientDP.SetFocus
'End If
'CmbClientDP.Text = ""
'CmbClientDP.Clear
'tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
'      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
'      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
'      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
'      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
'
''      "WHERE (((Mst_Depository.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
'
'Call Tmp1Table
'
'For op = 1 To TmpRs1.RecordCount
'    If CmbClientDP.Text = "" Then
'       CmbClientDP.Text = TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
'    End If
'    CmbClientDP.AddItem TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
'    TmpRs1.MoveNext
'Next
'If ans(3, 0) <= 1 Then
'   CmbClientDP_LostFocus
'End If
'End Sub

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
tmp = " SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, "
tmp = tmp & " Mst_Client.ClientName,  Mst_Client.ClientID,Mst_Client.CMSPCode,Mst_Client.KYCFlag "
tmp = tmp & " FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID "
tmp = tmp & " WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID
If Edit_Flg = "A" Then
   tmp = tmp & " And Mst_CLient.Status = 'A' "
ElseIf Edit_Flg = "E" Then
   tmp = tmp & " And (Mst_CLient.Status = 'A' or ClientID in (select distinct ClientID from Txn_CxTF_ClientDetail where CxTF_DetailsID = " & TxtCxTFNo.Text & " and SM_Prefix = '" & TxtPreFix.Text & "' And TxnType = 'D')) "
End If
tmp = tmp & " And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
'     " WHERE (((Mst_Depository.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
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
mSiloLocation = False ''--- Ullhas Patil 21 April 2018
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
mContractID = ""

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
ChkSCMFlag.Value = 0

If RsTxn_CxTF_Details.RecordCount > 0 Then
   If CmbExchangeTypeID.Enabled = True Then
      CmbExchangeTypeID.SetFocus
   End If
End If

End Sub

Private Sub EditCmd_Click()

If TxtCxTFDate.Value < Date - 2 Then
   MsgBox "This CDTF is more than 2 days old. You Can't Edit This CDTF!!! "
   Exit Sub
End If

mAdjGSLID = "(0"

For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 13) <> "" Then
       mAdjGSLID = mAdjGSLID & "," & Val(MSHFlexGrid4.TextMatrix(I, 13))
    End If
Next

mAdjGSLID = mAdjGSLID & ")"

tmp = "Select distinct a.Mst_GSL_ID from Txn_GSLAdjTransaction a"
tmp = tmp & " Inner Join Txn_AdjTransaction b On  a.TxnID = b.TxnID And b.SM_prefix = '" & TxtPreFix & "'"
tmp = tmp & " Where a.Mst_GSL_ID in " & mAdjGSLID & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   MsgBox "Some of lot is/are adjusted through Adjustement Module. You Can't Edit This CDTF!!! "
   Exit Sub
End If

If RsTxn_CxTF_Details!Flag = "P" Then
   MsgBox "Print out is Generated, You Can't Edit This CDTF   !!!! "
   Exit Sub
End If

For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 17) <> "" Then
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

If Check1.Value = 1 Then
   Call GetGodownMapped
End If

End Sub
Private Sub DeleteCmd_Click()

'GSLID from grid MSHFlexGrid4.TextMatrix(1, 13)

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" Then
   If TxtCxTFDate.Value < Date - 2 Then
      MsgBox "This CDTF is more than 2 days old. You Can't Delete This CDTF!!! "
      Exit Sub
   End If
End If

mAdjGSLID = "(0"

For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 13) <> "" Then
       mAdjGSLID = mAdjGSLID & "," & Val(MSHFlexGrid4.TextMatrix(I, 13))
    End If
Next

mAdjGSLID = mAdjGSLID & ")"

tmp = "Select distinct a.Mst_GSL_ID from Txn_GSLAdjTransaction a"
tmp = tmp & " Inner Join Txn_AdjTransaction b On  a.TxnID = b.TxnID And b.SM_prefix = '" & TxtPreFix & "'"
tmp = tmp & " Where a.Mst_GSL_ID in " & mAdjGSLID & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   MsgBox "Some of lot is/are adjusted through Adjustement Module. You Can't Delete This CDTF!!! "
   Exit Sub
End If


If RsTxn_CxTF_Details!Flag = "P" Then
   MsgBox "Print out is Generated, You Can't Edit This CDTF   !!!! "
   Exit Sub
End If

For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 17) <> "" Then
       MsgBox "Some of lot is trasfer, You Can't Delete This CDTF!!! "
       Exit Sub
    End If
Next

'tmp = "Select * from Txn_CashReceiptCDTFDetails Where SM_Prefix = '" & TxtPreFix & "' And CxTF_DetailsID = " & TxtCxTFNo & " "
'Call TmpTable

'If TmpRs.RecordCount > 0 Then
'   MsgBox "Cash Receipt generated for this CDTF. Can't delete!!!"
'   Exit Sub
'End If

Dim ans As Variant
'On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

   Call TableLog_CxTF_Details

   tmp = "Select * from Log_CxTF_Details Where TxnType = 'D' And SM_prefix = '" & RsTxn_CxTF_Details!SM_Prefix & "' And CxTF_DetailsID = " & RsTxn_CxTF_Details!CxTF_DetailsID & ""
   Call TmpTable
   
   If TmpRs.RecordCount = 0 Then
      RsLog_CxTF_Details.AddNew
      RsLog_CxTF_Details!CxTF_DetailsID = RsTxn_CxTF_Details!CxTF_DetailsID
      RsLog_CxTF_Details!TxnType = RsTxn_CxTF_Details!TxnType
      RsLog_CxTF_Details!SM_Prefix = RsTxn_CxTF_Details!SM_Prefix
      RsLog_CxTF_Details!CxTFDate = RsTxn_CxTF_Details!CxTFDate
      RsLog_CxTF_Details!CMPID = RsTxn_CxTF_Details!CMPID
      RsLog_CxTF_Details!CommodityID = RsTxn_CxTF_Details!CommodityID
      RsLog_CxTF_Details!ExchangeTypeID = RsTxn_CxTF_Details!ExchangeTypeID
      RsLog_CxTF_Details!BookedSpaceFlag = RsTxn_CxTF_Details!BookedSpaceFlag
      RsLog_CxTF_Details!Flag = RsTxn_CxTF_Details!Flag
      RsLog_CxTF_Details!Agent = RsTxn_CxTF_Details!Agent
      RsLog_CxTF_Details!Remark = RsTxn_CxTF_Details!Remark
      RsLog_CxTF_Details!ContractID = RsTxn_CxTF_Details!ContractID
      RsLog_CxTF_Details!G_UID = RsTxn_CxTF_Details!G_UID
      RsLog_CxTF_Details!HologramNo = RsTxn_CxTF_Details!HologramNo
      RsLog_CxTF_Details!AgentID = RsTxn_CxTF_Details!AgentID
      RsLog_CxTF_Details!ManualCDTFNo = RsTxn_CxTF_Details!ManualCDTFNo
      RsLog_CxTF_Details!CreatedBy = RsTxn_CxTF_Details!CreatedBy
      RsLog_CxTF_Details!CreatedDate = RsTxn_CxTF_Details!CreatedDate
      RsLog_CxTF_Details!UpdatedDate = RsTxn_CxTF_Details!UpdatedDate
      RsLog_CxTF_Details!UpdatedBy = RsTxn_CxTF_Details!UpdatedBy
      RsLog_CxTF_Details!DeletedBy = Gen_UserLoginID
      RsLog_CxTF_Details!CxTFNO = 0
      RsLog_CxTF_Details.Update
      RsLog_CxTF_Details.Close
   End If
   
   Call DeleteChildTableDetail
   
   tmp = "Delete from Txn_GSL where CxTF_DetailsID = " & TxtCxTFNo
   Call TmpTable
   
   
   tmp = "Select * from Txn_CxTF_TruckDetail where TxnType = 'D'  And  CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
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
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_CxTF_Details.CancelUpdate
End Sub

Sub DeleteChildTableDetail()
Dim adjWt As Variant
Dim I As Integer

Call TableMst_GSL("where SM_Prefix = '" & TxtPreFix & "'")
tmp = "Select * from Txn_CxTF_GSL where TxnType = 'D' And CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
For I = 1 To TmpRs.RecordCount
    RsMst_GSL.MoveFirst
    RsMst_GSL.Find "Mst_GSL_ID = " & TmpRs!Mst_GSL_ID
    If Not RsMst_GSL.EOF Then
    
    
       If mExchangeTypeID = 2 Then
          mCalMoisturePer = CheckMoisturePerDelete(I)
          RsMst_GSL!MoisturePer = Format(mCalMoisturePer, "#######0.000")
       End If
      
       
       RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - TmpRs!NoofBags
       RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - TmpRs!DematWeight
       RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - TmpRs!NoofBags
       RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - TmpRs!DematWeight
       RsMst_GSL!GrossWt = IIf(IsNull(RsMst_GSL!GrossWt), 0, RsMst_GSL!GrossWt) - TmpRs!GrossWeight
       
       
       
       
       If RsMst_GSL!DepositeWeight = 0 Then
          RsMst_GSL!Status = "B"
          RsMst_GSL!ClientID = Null
          RsMst_GSL!CommodityID = Null
          RsMst_GSL!ExchangeTypeID = Null
          RsMst_GSL!CdtfNo = Null
          RsMst_GSL!CDTFDate = Null
          RsMst_GSL!SCMFlag = "False"
          RsMst_GSL!MoisturePer = 0
          RsMst_GSL!CxTF_DetailsID = Null
       Else
          RsMst_GSL!Status = "O"
          
       End If
       RsMst_GSL.Update
    End If
    TmpRs.MoveNext
Next
tmp = "Delete from Txn_CxTF_ClientDetail where TxnType = 'D'  And  CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
tmp = "Delete from Txn_CxTF_GSL where TxnType = 'D'  And  CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
tmp = "Delete from Txn_CxTF_TruckDetail where TxnType = 'D'  And  CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
tmp = "select * from Txn_CxTF_GSL where TxnType = 'D'  And  CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'"
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

tmp = Format(Val(MSHFlexGrid3.TextMatrix(0, 6)) - Val(MSHFlexGrid4.TextMatrix(0, 4)), "#######0")
  
If Val(tmp) <> 0 Then
''   If MsgBox("No. of Bags in 'Truck Detail' Tab and 'GSL Detail' tab are not equal, do you wan't to continue? ", vbYesNo) = vbNo Then
''      Exit Sub
''   End If
   MsgBox "No. of Bags in 'Truck Detail' tab and 'GSL Detail' tab are not equal!!!"
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
   RsTxn_CxTF_Details!CxTFNO = GetCxTFNo
   RsTxn_CxTF_Details!G_UID = GetGUID
   TxtCxTFNo.Text = RsTxn_CxTF_Details!CxTFNO
   TxtG_UID = RsTxn_CxTF_Details!G_UID
Else
    RsTxn_CxTF_Details.MoveFirst
    RsTxn_CxTF_Details.Find " G_UID= '" & TxtG_UID.Text & "'"
    Call DeleteChildTableDetail
    RsTxn_CxTF_Details.MoveFirst
    RsTxn_CxTF_Details.Find " G_UID= '" & TxtG_UID.Text & "'"
End If

RsTxn_CxTF_Details!CxTFNO = TxtCxTFNo.Text
RsTxn_CxTF_Details!ManualCDTFNo = TxtManualCDTFNo
RsTxn_CxTF_Details!CxTFDate = TxtCxTFDate.Value
RsTxn_CxTF_Details!CMPID = mCMPID
RsTxn_CxTF_Details!CommodityID = mCommodityID
RsTxn_CxTF_Details!ExchangeTypeID = mExchangeTypeID
RsTxn_CxTF_Details!BookedSpaceFlag = Check1.Value   'CmbContractID.Text

'RsTxn_CxTF_Details!Flag = TxtFlag.Text
RsTxn_CxTF_Details!Agent = TxtAgent.Text
RsTxn_CxTF_Details!AgentID = mAgentID

'RsTxn_CxTF_Details!Remark = TxtRemark.Text
RsTxn_CxTF_Details!TxnType = "D"
RsTxn_CxTF_Details!SM_Prefix = TxtPreFix
RsTxn_CxTF_Details!SCMFlag = ChkSCMFlag.Value

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

tmp = "Select CxTF_DetailsID From Txn_CxTF_Details Where G_UID = '" & TxtG_UID & "'"
Call TmpTable

If TmpRs.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
   TxtCxTFNo.Text = TmpRs!CxTF_DetailsID
End If


SaveClient
SaveTruck
SaveGSL
SaveCmd.Enabled = False

RsTxn_CxTF_Details.Requery
RsTxn_CxTF_Details.Find " G_UID= '" & TxtG_UID.Text & "' "

If ChkTareWeightDiscripency = True Then
   Call Gen_MailTareWeightDiscripency
End If

Call Indata

End Sub
Function getCxtf_ClinetMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(ClientDetailID) from Txn_CxTF_ClientDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCxtf_ClinetMax = Retval

End Function
Private Sub SaveClient()
Call TableTxn_CxTF_ClientDetail(" Where TxnType = 'D'  And  CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsTxn_CxTF_ClientDetail.AddNew
       RsTxn_CxTF_ClientDetail!G_UID = GetGUID
       RsTxn_CxTF_ClientDetail!ClientDetailID = getCxtf_ClinetMax
       RsTxn_CxTF_ClientDetail!CxTF_DetailsID = TxtCxTFNo
       RsTxn_CxTF_ClientDetail!ClientID = MSHFlexGrid2.TextMatrix(I, 7)
       RsTxn_CxTF_ClientDetail!TxnType = "D"
       RsTxn_CxTF_ClientDetail!SM_Prefix = TxtPreFix
       RsTxn_CxTF_ClientDetail.Update
       RsTxn_CxTF_ClientDetail.Requery
    End If
Next
End Sub

Private Sub InDataClient()
Dim ans As Variant

Call TableTxn_CxTF_ClientDetail(" Where TxnType = 'D'  And CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_CxTF_ClientDetail.RecordCount
For I = 2 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(I, 0) = RsTxn_CxTF_ClientDetail!ClientDetailID
    'Parag
    'ans = GetClientNameByRow(RsTxn_CxTF_ClientDetail!ClientIDRow)
    ans = GetClientNameByRow(RsTxn_CxTF_ClientDetail!ClientID)
    MSHFlexGrid2.TextMatrix(I, 1) = ans(0)
    MSHFlexGrid2.TextMatrix(I, 2) = ans(1) 'TmpRs!ClientName
    MSHFlexGrid2.TextMatrix(I, 3) = ans(2) 'TmpRs!DepositoryID
    MSHFlexGrid2.TextMatrix(I, 4) = ans(3) 'TmpRs!DepositoryName
    'Parag
    'MSHFlexGrid2.TextMatrix(I, 7) = RsTxn_CxTF_ClientDetail!ClientIDRow
    MSHFlexGrid2.TextMatrix(I, 7) = RsTxn_CxTF_ClientDetail!ClientID
    MSHFlexGrid2.TextMatrix(I, 8) = ans(8) 'IIf(IsNull(RsTxn_CxTF_ClientDetail!KYCFlag), 0, RsTxn_CxTF_ClientDetail!KYCFlag)
    If Check1.Value = 0 Then
       MSHFlexGrid2.TextMatrix(I, 9) = GetCLientGSTDetails(Val(MSHFlexGrid2.TextMatrix(I, 7)), mCMPID, "GNAV")
    Else
       MSHFlexGrid2.TextMatrix(I, 9) = GetCLientGSTDetails(Val(MSHFlexGrid2.TextMatrix(I, 7)), mCMPID, "RGNAV")
    End If
    MSHFlexGrid2.TextMatrix(I, 10) = GetCLientGSTDetails(Val(MSHFlexGrid2.TextMatrix(I, 7)), mCMPID, "a")
    MSHFlexGrid2.TextMatrix(I, 11) = GetCLientGSTDetails(Val(MSHFlexGrid2.TextMatrix(I, 7)), mCMPID, "g")
    
    RsTxn_CxTF_ClientDetail.MoveNext
Next
'MsgBox TmpRs.State

End Sub
Function getCxtf_TruckMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(TruckDetailID) from Txn_CxTF_TruckDetail where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCxtf_TruckMax = Retval

End Function
Private Sub SaveTruck()
Call TableTxn_CxTF_TruckDetail(" Where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
       RsTxn_CxTF_TruckDetail.AddNew
       RsTxn_CxTF_TruckDetail!TruckDetailID = getCxtf_TruckMax
       RsTxn_CxTF_TruckDetail!G_UID = GetGUID
       RsTxn_CxTF_TruckDetail!CxTF_DetailsID = TxtCxTFNo
       RsTxn_CxTF_TruckDetail!TruckNo = MSHFlexGrid3.TextMatrix(I, 1)
       
       RsTxn_CxTF_TruckDetail!WeighBridgeSlipno = MSHFlexGrid3.TextMatrix(I, 2)
       RsTxn_CxTF_TruckDetail!TotalWeight = MSHFlexGrid3.TextMatrix(I, 3)
       RsTxn_CxTF_TruckDetail!TareWeight = Val(MSHFlexGrid3.TextMatrix(I, 4))
       RsTxn_CxTF_TruckDetail!GrossWeight = MSHFlexGrid3.TextMatrix(I, 5)
       ''-- Ullhas Patil 21 April 2018
       RsTxn_CxTF_TruckDetail!ForeignMatter = IIf(MSHFlexGrid3.TextMatrix(I, 13) = "", Null, MSHFlexGrid3.TextMatrix(I, 13))
       ''-----------------------------
       RsTxn_CxTF_TruckDetail!NoofBags = Val(MSHFlexGrid3.TextMatrix(I, 6))
       RsTxn_CxTF_TruckDetail!Remarks = MSHFlexGrid3.TextMatrix(I, 7)
       RsTxn_CxTF_TruckDetail!TxnType = "D"
       RsTxn_CxTF_TruckDetail!MoisturePer = Val(MSHFlexGrid3.TextMatrix(I, 8))
       RsTxn_CxTF_TruckDetail!SM_Prefix = TxtPreFix
       RsTxn_CxTF_TruckDetail!VehicleTypeID = Val(MSHFlexGrid3.TextMatrix(I, 10))
       RsTxn_CxTF_TruckDetail.Update
       RsTxn_CxTF_TruckDetail.Requery
    End If
Next
End Sub
Function getCxtf_GSLMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(GSLDetailID) from Txn_CxTF_GSL where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCxtf_GSLMax = Retval

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
tmp = tmp & " c.ClientName + ' ~ ' + d.DepositoryName,"
tmp = tmp & " tgsl.Remarks,"
tmp = tmp & " tgsl.Mst_GSL_ID,"
tmp = tmp & " god.CloseDate,"
tmp = tmp & " god.GodownID,"
tmp = tmp & " tgsl.ClientID,"
tmp = tmp & " gsl.ISINID,"
tmp = tmp & " gsl.FGSLID_Mst_GSL_ID, "
tmp = tmp & " god.address,Tgsl.WeighBridgeSlipno,tgsl.TruckNo,god.YearlyBillingFlg "
tmp = tmp & " from"
tmp = tmp & " Txn_CxTF_GSL tgsl"
tmp = tmp & " inner join mst_gsl gsl on gsl.Mst_GSL_ID = tgsl.Mst_GSL_ID and gsl.SM_Prefix = tgsl.SM_Prefix"
tmp = tmp & " inner join Mst_Godown god on gsl.GodownID = god.GodownID"
tmp = tmp & " inner join mst_client c on tgsl.ClientID = c.ClientID"
tmp = tmp & " inner join mst_depository d on d.DepositoryID = c.DepositoryID  and d.ExchangeTypeID = " & mExchangeTypeID
tmp = tmp & " Where tgsl.TxnType = 'D'  And  tgsl.CxTF_DetailsID = " & TxtCxTFNo & " And tgsl.SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable

Call Grid_Head_GCTD
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs.RecordCount
If TmpRs.RecordCount = 0 Then Exit Sub

Dim ans As Variant

For I = 2 To MSHFlexGrid4.Rows - 2
    For C = 0 To 16
        MSHFlexGrid4.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    If Not IsNull(TmpRs!FGSLID_Mst_GSL_ID) Then
       If MSHFlexGrid4.TextMatrix(I, 17) = "" Or MSHFlexGrid4.TextMatrix(I, 17) = "0" Then
          MSHFlexGrid4.TextMatrix(I, 17) = IIf(IsNull(TmpRs!FGSLID_Mst_GSL_ID), "", TmpRs!FGSLID_Mst_GSL_ID)
       End If
    End If
    MSHFlexGrid4.TextMatrix(I, 18) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
    MSHFlexGrid4.TextMatrix(I, 21) = IIf(IsNull(TmpRs!YearlyBillingFlg), 0, TmpRs!YearlyBillingFlg)
    MSHFlexGrid4.TextMatrix(I, 22) = Format(Val(MSHFlexGrid4.TextMatrix(I, 10)) * 1000 / Val(MSHFlexGrid4.TextMatrix(I, 4)), "#######0.00")
    MSHFlexGrid4.Col = 5
    MSHFlexGrid4.row = I
    MSHFlexGrid4.CellBackColor = &HFFC0C0
    
    MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(I, 4))
    MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(I, 5))
    MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(I, 6))
    MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(I, 7))
    MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(I, 8))
    MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(I, 9))
    MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(I, 10))
    
    MSHFlexGrid4.TextMatrix(I, 4) = Format(MSHFlexGrid4.TextMatrix(I, 4), "#######0")
    MSHFlexGrid4.TextMatrix(I, 5) = Format(MSHFlexGrid4.TextMatrix(I, 5), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 6) = Format(MSHFlexGrid4.TextMatrix(I, 6), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 7) = Format(MSHFlexGrid4.TextMatrix(I, 7), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 8) = Format(MSHFlexGrid4.TextMatrix(I, 8), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 9) = Format(MSHFlexGrid4.TextMatrix(I, 9), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 10) = Format(MSHFlexGrid4.TextMatrix(I, 10), "#######0.000")
    
    MSHFlexGrid4.TextMatrix(I, 19) = IIf(IsNull(TmpRs!TruckNo), "", TmpRs!TruckNo) & "~" & IIf(IsNull(TmpRs!WeighBridgeSlipno), "", TmpRs!WeighBridgeSlipno)

    MSHFlexGrid4.TextMatrix(I, 0) = Format(MSHFlexGrid4.TextMatrix(0, 4), "#######0")
    MSHFlexGrid4.TextMatrix(I, 0) = Format(MSHFlexGrid4.TextMatrix(0, 5), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 0) = Format(MSHFlexGrid4.TextMatrix(0, 6), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 0) = Format(MSHFlexGrid4.TextMatrix(0, 7), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 0) = Format(MSHFlexGrid4.TextMatrix(0, 8), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 0) = Format(MSHFlexGrid4.TextMatrix(0, 9), "#######0.000")
    MSHFlexGrid4.TextMatrix(I, 0) = Format(MSHFlexGrid4.TextMatrix(0, 10), "#######0.000")
    
    For j = 2 To MSHFlexGrid3.Rows - 1
        If MSHFlexGrid4.TextMatrix(I, 19) = MSHFlexGrid3.TextMatrix(j, 1) & "~" & MSHFlexGrid3.TextMatrix(j, 2) Then
           MSHFlexGrid4.TextMatrix(I, 20) = MSHFlexGrid3.TextMatrix(j, 8)
        End If
    Next
    
    TmpRs.MoveNext
Next


End Sub


Private Sub InDataGSL_Old()

Call TableTxn_CxTF_GSL(" Where TxnType = 'D'  And  CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + RsTxn_CxTF_GSL.RecordCount
If RsTxn_CxTF_GSL.RecordCount = 0 Then Exit Sub

Dim ans As Variant

For I = 2 To MSHFlexGrid4.Rows - 2
    
    ans = GetGSLByGLSID(RsTxn_CxTF_GSL!Mst_GSL_ID, TxtPreFix)
    MSHFlexGrid4.TextMatrix(I, 0) = RsTxn_CxTF_GSL!GSLDetailID
    
    MSHFlexGrid4.TextMatrix(I, 1) = ans(3) ' TmpRs!GodownNo
    MSHFlexGrid4.TextMatrix(I, 2) = ans(4) 'TmpRs!StackNo
    MSHFlexGrid4.TextMatrix(I, 3) = ans(5) 'TmpRs!LotNo
    MSHFlexGrid4.TextMatrix(I, 14) = IIf(IsNull(ans(6)), "", ans(6))
    MSHFlexGrid4.TextMatrix(I, 15) = IIf(IsNull(ans(2)), "", ans(2))
    MSHFlexGrid4.TextMatrix(I, 17) = IIf(IsNull(ans(10)) Or ans(10) = 0, "", ans(10))
    
    MSHFlexGrid4.TextMatrix(I, 4) = RsTxn_CxTF_GSL!NoofBags
    MSHFlexGrid4.TextMatrix(I, 5) = RsTxn_CxTF_GSL!GrossWeight
    MSHFlexGrid4.TextMatrix(I, 6) = RsTxn_CxTF_GSL!AdjNClientWeight
    MSHFlexGrid4.TextMatrix(I, 7) = RsTxn_CxTF_GSL!AdjNCMSLWeight
    MSHFlexGrid4.TextMatrix(I, 8) = RsTxn_CxTF_GSL!GunnyWeight
    MSHFlexGrid4.TextMatrix(I, 9) = RsTxn_CxTF_GSL!StdDeduction
    MSHFlexGrid4.TextMatrix(I, 10) = RsTxn_CxTF_GSL!DematWeight
    
    ans = GetClientNameByRow(RsTxn_CxTF_GSL!ClientID)
    MSHFlexGrid4.TextMatrix(I, 11) = ans(1) & " ~ " & ans(3)
    MSHFlexGrid4.TextMatrix(I, 12) = RsTxn_CxTF_GSL!Remarks
    
    MSHFlexGrid4.TextMatrix(I, 13) = RsTxn_CxTF_GSL!Mst_GSL_ID
   
    'MSHFlexGrid4.TextMatrix(i, 14) = 'look at top Used Up
    'MSHFlexGrid4.TextMatrix(i, 15) = 'look at top Used Up
    'MSHFlexGrid4.TextMatrix(i, 17) = 'look at top Used Up
    
    MSHFlexGrid4.TextMatrix(I, 16) = RsTxn_CxTF_GSL!ClientID
    
    MSHFlexGrid4.Col = 5
    MSHFlexGrid4.row = I
    MSHFlexGrid4.CellBackColor = &HFFC0C0

    
    MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(I, 4))
    MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(I, 5))
    MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(I, 6))
    MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(I, 7))
    MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(I, 8))
    MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(I, 9))
    MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(I, 10))
    
    RsTxn_CxTF_GSL.MoveNext
Next
End Sub

Private Sub SaveGSL()
Dim adjWt As Variant
Dim strSplitString() As String
Dim I As Integer
Call TableTxn_CxTF_GSL(" Where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")
Call TableMst_GSL("where SM_Prefix = '" & TxtPreFix & "'")
For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 1) <> "" Then
       RsTxn_CxTF_GSL.AddNew
       RsTxn_CxTF_GSL!GSLDetailID = getCxtf_GSLMax
       RsTxn_CxTF_GSL!G_UID = GetGUID
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "Mst_GSL_ID = " & Val(MSHFlexGrid4.TextMatrix(I, 13))
       
'       If Not RsMst_GSL.EOF Then
'          RsMst_GSL!CommodityID = mCommodityID
'          RsMst_GSL!ExchangeTypeID = mExchangeTypeID
'          RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsTxn_CxTF_GSL!NoofBags), 0, RsTxn_CxTF_GSL!NoofBags)
'          RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsTxn_CxTF_GSL!DematWeight), 0, RsTxn_CxTF_GSL!DematWeight)
'          RsMst_GSL.Update
'       End If
        
       RsTxn_CxTF_GSL!CxTF_DetailsID = TxtCxTFNo
       RsTxn_CxTF_GSL!NoofBags = MSHFlexGrid4.TextMatrix(I, 4)
       RsTxn_CxTF_GSL!GrossWeight = MSHFlexGrid4.TextMatrix(I, 5)
       RsTxn_CxTF_GSL!AdjNClientWeight = Val(MSHFlexGrid4.TextMatrix(I, 6))
       RsTxn_CxTF_GSL!AdjNCMSLWeight = Val(MSHFlexGrid4.TextMatrix(I, 7))
       RsTxn_CxTF_GSL!GunnyWeight = Val(MSHFlexGrid4.TextMatrix(I, 8))
       RsTxn_CxTF_GSL!StdDeduction = Val(MSHFlexGrid4.TextMatrix(I, 9))
       RsTxn_CxTF_GSL!DematWeight = Val(MSHFlexGrid4.TextMatrix(I, 10))
       RsTxn_CxTF_GSL!ClientID = Val(MSHFlexGrid4.TextMatrix(I, 16))
       RsTxn_CxTF_GSL!Remarks = MSHFlexGrid4.TextMatrix(I, 12)
       RsTxn_CxTF_GSL!Mst_GSL_ID = MSHFlexGrid4.TextMatrix(I, 13)
       RsTxn_CxTF_GSL!TxnType = "D"
       RsTxn_CxTF_GSL!SM_Prefix = TxtPreFix
       
       
       strSplitString = Split(MSHFlexGrid4.TextMatrix(I, 19), "~")
       RsTxn_CxTF_GSL!TruckNo = strSplitString(0)
       RsTxn_CxTF_GSL!WeighBridgeSlipno = strSplitString(1)
       RsTxn_CxTF_GSL!MoisturePer = IIf(MSHFlexGrid4.TextMatrix(I, 20) <> "", Format(MSHFlexGrid4.TextMatrix(I, 20), "#######0.000"), 0)
       
       RsTxn_CxTF_GSL.Update
       RsTxn_CxTF_GSL.Requery
       
       If Not RsMst_GSL.EOF Then
          '  Added Moisture Persentage calculation Request date : 10.10.2014 NCDEX New Process by Mayur
          'If mExchangeTypeID = 2 Then
             mCalMoisturePer = CheckMoisturePer(I)
             RsMst_GSL!MoisturePer = Format(mCalMoisturePer, "#######0.000")
          'End If
          
          RsMst_GSL!CommodityID = mCommodityID
          RsMst_GSL!ExchangeTypeID = mExchangeTypeID
          RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) + Val(MSHFlexGrid4.TextMatrix(I, 4))
          RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(MSHFlexGrid4.TextMatrix(I, 10))
          RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(MSHFlexGrid4.TextMatrix(I, 4))
          RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(MSHFlexGrid4.TextMatrix(I, 10))
          adjWt = Val(MSHFlexGrid4.TextMatrix(I, 6)) - Val(MSHFlexGrid4.TextMatrix(I, 9))
          RsMst_GSL!GrossWt = IIf(IsNull(RsMst_GSL!GrossWt), 0, RsMst_GSL!GrossWt) + Val(MSHFlexGrid4.TextMatrix(I, 5))
          RsMst_GSL!LCDTF_Date = TxtCxTFDate.Value
          RsMst_GSL!KRNo = IIf(IsNull(RsMst_GSL!KRNo), 0, RsMst_GSL!KRNo)
             
          If Check1.Value = 1 Then
             RsMst_GSL!ContractID = mContractID
             RsMst_GSL!NContractID = mContractID
          End If
          
          'RsMst_GSL!SCMFlag = ChkSCMFlag.Value
          
          If IsNull(RsMst_GSL!CDTFDate) Then
             RsMst_GSL!CxTF_DetailsID = Val(TxtCxTFNo)
             RsMst_GSL!CDTFDate = TxtCxTFDate.Value
          End If
          RsMst_GSL!Status = "O"
          RsMst_GSL!State = "A"
          If ChkSCMFlag.Value = 0 Then
             RsMst_GSL!SCMFlag = "False"
          Else
             RsMst_GSL!SCMFlag = "True"
          End If
          RsMst_GSL!ClientID = Val(MSHFlexGrid4.TextMatrix(I, 16))
          RsMst_GSL!OriginalDepositFlag = "Y"
          RsMst_GSL.Update
       End If
       
    End If
Next
End Sub
Private Sub InDataTruck()

Call TableTxn_CxTF_TruckDetail(" Where TxnType = 'D'  And  CxTF_DetailsID = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")

MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + RsTxn_CxTF_TruckDetail.RecordCount
For I = 2 To MSHFlexGrid3.Rows - 2
   MSHFlexGrid3.TextMatrix(I, 0) = RsTxn_CxTF_TruckDetail!TruckDetailID
   MSHFlexGrid3.TextMatrix(I, 1) = RsTxn_CxTF_TruckDetail!TruckNo
   MSHFlexGrid3.TextMatrix(I, 2) = RsTxn_CxTF_TruckDetail!WeighBridgeSlipno
   MSHFlexGrid3.TextMatrix(I, 3) = RsTxn_CxTF_TruckDetail!TotalWeight
   MSHFlexGrid3.TextMatrix(I, 4) = RsTxn_CxTF_TruckDetail!TareWeight
   MSHFlexGrid3.TextMatrix(I, 5) = RsTxn_CxTF_TruckDetail!GrossWeight
   MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(RsTxn_CxTF_TruckDetail!NoofBags), 0, RsTxn_CxTF_TruckDetail!NoofBags)
   MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(RsTxn_CxTF_TruckDetail!Remarks), "", RsTxn_CxTF_TruckDetail!Remarks)
   MSHFlexGrid3.TextMatrix(I, 8) = IIf(IsNull(RsTxn_CxTF_TruckDetail!MoisturePer), "", RsTxn_CxTF_TruckDetail!MoisturePer)
   
   MSHFlexGrid3.TextMatrix(I, 9) = IIf(IsNull(RsTxn_CxTF_TruckDetail!VehicleTypeID), "", GetVehicleType(RsTxn_CxTF_TruckDetail!VehicleTypeID, "N"))
   MSHFlexGrid3.TextMatrix(I, 10) = IIf(IsNull(RsTxn_CxTF_TruckDetail!VehicleTypeID), "", RsTxn_CxTF_TruckDetail!VehicleTypeID)
   MSHFlexGrid3.TextMatrix(I, 11) = IIf(IsNull(RsTxn_CxTF_TruckDetail!VehicleTypeID), "", GetVehicleType(RsTxn_CxTF_TruckDetail!VehicleTypeID, "Min"))
   MSHFlexGrid3.TextMatrix(I, 12) = IIf(IsNull(RsTxn_CxTF_TruckDetail!VehicleTypeID), "", GetVehicleType(RsTxn_CxTF_TruckDetail!VehicleTypeID, "Max"))
   MSHFlexGrid3.TextMatrix(I, 13) = IIf(IsNull(RsTxn_CxTF_TruckDetail!ForeignMatter), "", RsTxn_CxTF_TruckDetail!ForeignMatter) ''-- Ullhas Patil 21 April 2018
   
   MSHFlexGrid3.Col = 5
   MSHFlexGrid3.row = I
   MSHFlexGrid3.CellBackColor = &HFFC0C0
       
   MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + RsTxn_CxTF_TruckDetail!TotalWeight
   MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + RsTxn_CxTF_TruckDetail!TareWeight
   MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + RsTxn_CxTF_TruckDetail!GrossWeight
   MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) + IIf(IsNull(RsTxn_CxTF_TruckDetail!NoofBags), 0, RsTxn_CxTF_TruckDetail!NoofBags)
   MSHFlexGrid3.TextMatrix(0, 13) = Val(MSHFlexGrid3.TextMatrix(0, 13)) + Val(IIf(IsNull(RsTxn_CxTF_TruckDetail!ForeignMatter), "", RsTxn_CxTF_TruckDetail!ForeignMatter)) ''-- Ullhas Patil 21 April 2018
   
   RsTxn_CxTF_TruckDetail.MoveNext
Next
End Sub

'Private Function GetCxTFNo() As Double
'Dim RetVal As Double
'tmp = "Select max(CxTFNo) from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TxtPreFix & "'"
'Call TmpTable
'RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
'RetVal = RetVal + 1
'GetCxTFNo = RetVal
'End Function

Private Function GetCxTFNo() As Double

Dim Retval As Double
Dim LogRetVal As Double
tmp = "Select max(CxTFNo) from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable

tmp = "Select max(CxTFNo) from Log_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TxtPreFix & "'"
Call Tmp1Table

Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
LogRetVal = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))

If Retval > LogRetVal Then
   Retval = Retval + 1
Else
   Retval = LogRetVal + 1
End If

GetCxTFNo = Retval

End Function

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 20 '19
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
MSHFlexGrid1.TextMatrix(0, 11) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 12) = "Manual CDTFNo"
MSHFlexGrid1.TextMatrix(0, 13) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 14) = "Gross Wt."
MSHFlexGrid1.TextMatrix(0, 15) = "Guid"
MSHFlexGrid1.TextMatrix(0, 16) = "Received Annexure-III"
MSHFlexGrid1.TextMatrix(0, 17) = "SCM Flag"
MSHFlexGrid1.TextMatrix(0, 18) = "Created By"
MSHFlexGrid1.TextMatrix(0, 19) = "Created Date"


MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 2600
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 0 '1200
MSHFlexGrid1.ColWidth(7) = 0
MSHFlexGrid1.ColWidth(8) = 0 '1600
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1400 '1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(15) = 0
MSHFlexGrid1.ColWidth(17) = 800
MSHFlexGrid1.ColWidth(18) = 800
MSHFlexGrid1.ColWidth(19) = 1500


MSHFlexGrid1.RowHeight(0) = 700
End Sub
Private Sub Grid_Head_Client()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.Cols = 12 '9
MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Client ID"
MSHFlexGrid2.TextMatrix(1, 2) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 3) = "DP ID"
MSHFlexGrid2.TextMatrix(1, 4) = "DP Name "
MSHFlexGrid2.TextMatrix(1, 5) = "No of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Weight"
MSHFlexGrid2.TextMatrix(1, 7) = "ClientRowID"
MSHFlexGrid2.TextMatrix(1, 8) = "KYCFlag"

MSHFlexGrid2.TextMatrix(1, 9) = "NAVCode"
MSHFlexGrid2.TextMatrix(1, 10) = "Address"
MSHFlexGrid2.TextMatrix(1, 11) = "GSTNo"

MSHFlexGrid2.ColWidth(0) = 500
MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 3500
MSHFlexGrid2.ColWidth(3) = 1800
MSHFlexGrid2.ColWidth(4) = 3500
MSHFlexGrid2.ColWidth(5) = 1500
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(8) = 0
MSHFlexGrid2.ColWidth(9) = 1500
MSHFlexGrid2.ColWidth(10) = 1500
MSHFlexGrid2.ColWidth(11) = 1500

End Sub
Private Sub Grid_Head_TTD()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 2

MSHFlexGrid3.Cols = 14 '13 '9
MSHFlexGrid3.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid3.TextMatrix(1, 1) = "Truck No"
MSHFlexGrid3.TextMatrix(1, 2) = "Weighbridge Slip No"
MSHFlexGrid3.TextMatrix(1, 3) = "Total Weight"
MSHFlexGrid3.TextMatrix(1, 4) = "Tare Weight"
''MSHFlexGrid3.TextMatrix(1, 5) = "Gross Weight"
If mSiloLocation = False Then
    MSHFlexGrid3.TextMatrix(1, 5) = "Gross Weight"
Else
    MSHFlexGrid3.TextMatrix(1, 5) = "Net Weight"
End If
MSHFlexGrid3.TextMatrix(1, 6) = "No Of Bags"
MSHFlexGrid3.TextMatrix(1, 7) = "Quality Parameters / Remark"
MSHFlexGrid3.TextMatrix(1, 8) = "Moisture %"
MSHFlexGrid3.TextMatrix(1, 9) = "Vehicle Model"
MSHFlexGrid3.TextMatrix(1, 10) = "Vehicle Type ID"
MSHFlexGrid3.TextMatrix(1, 11) = "Minimum tare weight"
MSHFlexGrid3.TextMatrix(1, 12) = "Maximum tare weight"

MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 2000
MSHFlexGrid3.ColWidth(2) = 2500
MSHFlexGrid3.ColWidth(3) = 1200
MSHFlexGrid3.ColWidth(4) = 1200
MSHFlexGrid3.ColWidth(5) = 1200
MSHFlexGrid3.ColWidth(6) = 1200
MSHFlexGrid3.ColWidth(7) = 3800
MSHFlexGrid3.ColWidth(8) = 1000
MSHFlexGrid3.ColWidth(9) = 1000
MSHFlexGrid3.ColWidth(10) = 1200
MSHFlexGrid3.ColWidth(11) = 1000
MSHFlexGrid3.ColWidth(12) = 1000
MSHFlexGrid3.TextMatrix(1, 13) = "Freign Matter": MSHFlexGrid3.ColWidth(13) = 1100 ''-- Ullhas Patil 21 April 2018

'MSHFlexGrid3.Col = 5
'MSHFlexGrid3.Row = 0
'MSHFlexGrid3.CellBackColor = vbRed '&HFFC0C0

End Sub
Private Sub Grid_Head_GCTD()
MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 3
MSHFlexGrid4.FixedRows = 2

MSHFlexGrid4.Cols = 23 '22 '21
MSHFlexGrid4.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid4.TextMatrix(1, 1) = "Godown No"
MSHFlexGrid4.TextMatrix(1, 2) = "Stack No"
MSHFlexGrid4.TextMatrix(1, 3) = "Lot No"
MSHFlexGrid4.TextMatrix(1, 4) = "No of Bags"
MSHFlexGrid4.TextMatrix(1, 5) = "Gross Weight    (A) "
'MSHFlexGrid4.TextMatrix(1, 6) = "Adjusted Wt.    (B) "
'MSHFlexGrid4.TextMatrix(1, 7) = "Adj.Gross Wt."
MSHFlexGrid4.TextMatrix(1, 8) = "Gunny Weight    (B) "
MSHFlexGrid4.TextMatrix(1, 9) = "Std. Deduction    (C) "
MSHFlexGrid4.TextMatrix(1, 10) = "Electronic Weight   (A-B-C) "
MSHFlexGrid4.TextMatrix(1, 11) = "Client ID"
MSHFlexGrid4.TextMatrix(1, 12) = "Quality Parameters / Remark"
MSHFlexGrid4.TextMatrix(1, 13) = "GSL ID"
MSHFlexGrid4.TextMatrix(1, 14) = "Godown Close"
MSHFlexGrid4.TextMatrix(1, 15) = "Godown ID"
MSHFlexGrid4.TextMatrix(1, 16) = "ClientRowID"
MSHFlexGrid4.TextMatrix(1, 17) = "ISINID"
MSHFlexGrid4.TextMatrix(1, 18) = "Godown Address"
MSHFlexGrid4.TextMatrix(1, 19) = "Truck No"
MSHFlexGrid4.TextMatrix(1, 20) = "Moisture Per"
MSHFlexGrid4.TextMatrix(1, 21) = "Yearly Billing Flag"
MSHFlexGrid4.TextMatrix(1, 22) = "Bag Size"

MSHFlexGrid4.ColWidth(0) = 0
MSHFlexGrid4.ColWidth(1) = 1200
MSHFlexGrid4.ColWidth(2) = 1200
MSHFlexGrid4.ColWidth(3) = 1200
MSHFlexGrid4.ColWidth(4) = 1000
MSHFlexGrid4.ColWidth(7) = 0
MSHFlexGrid4.ColWidth(6) = 0
MSHFlexGrid4.ColWidth(8) = 1000
MSHFlexGrid4.ColWidth(11) = 3000
MSHFlexGrid4.ColWidth(12) = 1500
MSHFlexGrid4.ColWidth(13) = 0
MSHFlexGrid4.ColWidth(14) = 0
MSHFlexGrid4.ColWidth(15) = 0
MSHFlexGrid4.ColWidth(16) = 0
MSHFlexGrid4.ColWidth(17) = 0
MSHFlexGrid4.RowHeight(1) = 750
MSHFlexGrid4.ColWidth(18) = 0
MSHFlexGrid4.ColWidth(19) = 1500
MSHFlexGrid4.ColWidth(20) = 1500

'MSHFlexGrid4.Col = 5
'MSHFlexGrid4.Row = 0
'MSHFlexGrid4.CellBackColor = &HFFC0C0

End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_CxTF_Details.MoveFirst
RsTxn_CxTF_Details.Find "g_uid = '" & MSHFlexGrid1.TextMatrix(I, 15) & "'"
If Not RsTxn_CxTF_Details.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub CmdCMPID_Click()
    FrmMst_CMP.Show
End Sub
'----
Private Sub CmbCMPID_GotFocus()
wc = ""
If Gen_WcCMP <> "" Then
   wc = " And " & Gen_WcCMP
End If

    mTmp1 = CmbCMPID.Text
    Call TableMst_CMP(" where CloseDate is null " & wc)
    CmbCMPID.Clear
    If RsMst_CMP.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To RsMst_CMP.RecordCount
        CmbCMPID.AddItem RsMst_CMP!CMPName
        RsMst_CMP.MoveNext
    Next
    CmbCMPID.Text = mTmp1

End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   CmbContractID.Text = ""
   CmbCommodityID.Text = ""
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
   CmbCommodityID.Text = ""
End If

mCMPID = RsMst_CMP!CMPID
mLocationID = RsMst_CMP!LocationID 'GetLocationIDFromCMP(mCmpID)
'If RsMst_CMP!CurrentDate + 1 >= RsMst_CMP!CMPLockedDate Then
'   MsgBox "Month end process is pending for this CMP, Please Call CMG Department to Execute Month End Process !!! "
'   CmbCMPID.SetFocus
'   Exit Sub
'End If

TxtCxTFDate.Value = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate) + 1
mBODDate = TxtCxTFDate.Value
TxtCxTFDispalyDate = TxtCxTFDate.Value
mBlockDays = RsMst_CMP!BlockDays
mBlockDate = Date - Val(mBlockDays)
mBlockMessage = "CDTF Date less than BOD Date is not allowed!!!"

If mBODDate < mBlockDate Then
   TxtCxTFDate.Value = mBlockDate
   mBODDate = TxtCxTFDate.Value
   TxtCxTFDispalyDate = TxtCxTFDate.Value
   mBlockMessage = "CDTF Date less than Block Date (which is back " & mBlockDays & " no. of days from todays date) is not allowed!!! "
End If

'====Commented on 22 May 2012
''TxtCxTFDate.Enabled = False

'If LCase(CmbCMPID.Text) = "unjha" Or LCase(CmbCMPID.Text) = "rajkot" Then
'   TxtCxTFDate.Enabled = True
'Else
'   TxtCxTFDate.Enabled = False
'End If

'If SaveCmd.Enabled = True Then
'   MsgBox "CDTF Date is " & TxtCxTFDate
'End If
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
''-- Ullhas Patil 21 April 2018
mSiloLocation = CBool(GetPreFix(RsMst_CMP!CMPID, "S"))
If mSiloLocation = False Then
    TxtForeignMatter = "": TxtForeignMatter.BackColor = &H80000000: TxtForeignMatter.Enabled = False
    Label18.Caption = "Gross Weight"
Else
    TxtForeignMatter = "": TxtForeignMatter.BackColor = &H80000005: TxtForeignMatter.Enabled = True
    Label18.Caption = "Net Weight"
End If
If LCase(mTmp1) <> LCase(CmbCMPID.Text) Then
Call Grid_Head_TTD
End If
''------------------------------
Call CheckForReservationMessage

End Sub
'----
Private Sub CmdCommodityID_Click()
    FrmMst_Commodity.Show
End Sub
'----
'Private Sub CmbCommodityID_GotFocus()
'
' If CmbExchangeTypeID.Text <> "" Then
''    If CmbCMPID.Text <> "" Then
'       tmp = CmbCommodityID.Text
'       Call TableMst_Commodity
'       CmbCommodityID.Clear
'       If RsMst_Commodity.RecordCount = 0 Then
'          MsgBox "No Record found !!!! "
'          Exit Sub
'       End If
'       For I = 1 To RsMst_Commodity.RecordCount
'           CmbCommodityID.AddItem RsMst_Commodity!Commodity
'           RsMst_Commodity.MoveNext
'       Next
'       CmbCommodityID.Text = tmp
'
''    Else
''       MsgBox "Please Select CMP!!!!!"
''       CmbCMPID.SetFocus
''       Exit Sub
''    End If
' Else
'    MsgBox "Please Select Exchange Mode !!!!"
'    CmbExchangeTypeID.SetFocus
'    Exit Sub
' End If
'End Sub
''----
'Private Sub CmbCommodityID_LostFocus()
'If CmbCommodityID.Text = "" Then
'   Exit Sub
'End If
'
'RsMst_Commodity.MoveFirst
'RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
'If RsMst_Commodity.EOF Then
'   MsgBox "Invalid selection "
'   CmbCommodityID.SetFocus
'   Exit Sub
'End If
'
'If mCommodityID <> RsMst_Commodity!CommodityID Then
'   Call Grid_Head_GCTD
'   CmbContractID.Text = ""
'End If
'mCommodityID = RsMst_Commodity!CommodityID
'
'Call TableMst_CommodityDetail("Where CommodityId = " & mCommodityID & " and ExchangeTypeId = " & mExchangeTypeID)
' If RsMst_CommodityDetail.RecordCount > 0 Then
'    TxtCGunnyWt = RsMst_CommodityDetail!GunnyWt
'    TxtCStdDed = IIf(IsNull(RsMst_CommodityDetail!StdDeduction), "", RsMst_CommodityDetail!StdDeduction)
'    TxtCLotSize = IIf(IsNull(RsMst_CommodityDetail!LotSize), "", RsMst_CommodityDetail!LotSize)
'    TxtCVariation = IIf(IsNull(RsMst_CommodityDetail!Variation), "", RsMst_CommodityDetail!Variation)
' Else
'    MsgBox "No Record Found For Exchange" & CmbExchangeTypeID.Text & " - " & CmbCommodityID.Text
'    CmbCommodityID.SetFocus
'    Exit Sub
' End If
''SSTab1.Tab = 0
'End Sub

'----------------------------------------------------------

Private Sub CmbCommodityID_GotFocus()

If CmbExchangeTypeID.Text = "" Then
    MsgBox "Please Select Exchange Mode !!!!"
    CmbExchangeTypeID.SetFocus
    Exit Sub
End If

If CmbCMPID.Text = "" Then
    MsgBox "Please Select CMP!!!!!"
    CmbCMPID.SetFocus
    Exit Sub
End If

tmp1 = CmbCommodityID.Text
'Call TableMst_Commodity

If Edit_Flg = "" Then
    tmp = "Select MCom.CommodityID, MCom.Commodity,MCom.CommtrackFlag,MCom.CommodityGroupID,Mcom.YearlyBillingFlg From Mst_Commodity MCom Where MCom.CommodityID = " & mCommodityID
Else
    tmp = "Select MCom.CommodityID, MCom.Commodity,MCom.CommtrackFlag,MCom.CommodityGroupID,Mcom.YearlyBillingFlg From Mst_Commodity MCom Inner Join Mst_LocationCOmmodityMapping MLCM on MLCM.CommodityID = MCom.CommodityID "
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MLCM.LocationID "
    tmp = tmp & " Where ML.SM_Prefix = '" & TxtPreFix & "'"
End If

Call Tmp7Table
CmbCommodityID.Clear
'If RsMst_Commodity.RecordCount = 0 Then
If TmpRs7.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
'For I = 1 To RsMst_Commodity.RecordCount
'    CmbCommodityID.AddItem RsMst_Commodity!Commodity
'    RsMst_Commodity.MoveNext
'Next
For I = 1 To TmpRs7.RecordCount
    CmbCommodityID.AddItem TmpRs7!Commodity
    TmpRs7.MoveNext
Next
CmbCommodityID.Text = tmp1
    
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then Exit Sub

'RsMst_Commodity.MoveFirst
'RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
'If RsMst_Commodity.EOF Then
'   MsgBox "Invalid selection "
'   CmbCommodityID.SetFocus
'   Exit Sub
'End If

TmpRs7.MoveFirst
TmpRs7.Find "Commodity = '" & CmbCommodityID.Text & "'"
If TmpRs7.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If


If mCommodityID <> TmpRs7!CommodityID Then
   Call Grid_Head_GCTD
   CmbContractID.Text = ""
End If
mCommodityID = TmpRs7!CommodityID
mCommodityGroupId = TmpRs7!CommodityGroupID
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

mCommtrackFlag = IIf(IsNull(TmpRs7!CommTrackFlag), 0, TmpRs7!CommTrackFlag)
mCommodityYearlyFlag = IIf(IsNull(TmpRs7!YearlyBillingFlg), 0, TmpRs7!YearlyBillingFlg)



Call CheckForReservationMessage


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
For I = 1 To RsMst_ExchangeType.RecordCount
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

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   Call Grid_Head_Client
   Call Grid_Head_GCTD
   Call ClearClientTabControl
   Call ClearGCTDTabControl
   CmbContractID.Text = ""
   CmdClientAdd2List.Enabled = True
   'Added on 30 Nov 2011 for Reservation Deposit from CDTF
   Check1.Value = 0
   Check1.Enabled = False
   
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID


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
'Added on 30 Nov 2011 for Reservation Deposit from CDTF
   Check1.Enabled = True
   ChkSCMFlag.Enabled = True
Else
   Label9.Visible = True
   Label10.Visible = True
   Label11.Visible = True
   CmbNonNcdexClient.Visible = False
   ChkSCMFlag.Value = 0
   ChkSCMFlag.Enabled = False
End If

Call CheckForReservationMessage
Call Check1_Click
End Sub
'----
Private Sub CmdContractID_Click()
    FrmMst_StorageContract.Show
End Sub
'----
'Private Sub CmbContractID_GotFocus()
'If CmbCMPID.Text = "" Then
'   MsgBox "Select CMP !!!!! "
'   Exit Sub
'End If
'
'If CmbCommodityID.Text = "" Then
'   MsgBox "Select Commodity !!!!! "
'   Exit Sub
'End If
'Call Grid_Head_GCTD
'Call Grid_Head_Client
'tmp = CmbContractID.Text
'mLocationID = GetLocationIDFromCMP(Val(mCMPID))
'Call TableMst_StorageContract(" where (CommodityID is null or CommodityID = " & mCommodityID & " ) And LocationID = " & mLocationID)
'CmbContractID.Clear
'If RsMst_StorageContract.RecordCount = 0 Then
'   MsgBox "No Record found !!!! "
'   Check1.Value = 0
'   Exit Sub
'End If
'For I = 1 To RsMst_StorageContract.RecordCount
'    CmbContractID.AddItem RsMst_StorageContract!ContractName
'    RsMst_StorageContract.MoveNext
'Next
'CmbContractID.Text = tmp
'End Sub
''----
'Private Sub CmbContractID_LostFocus()
'If CmbContractID.Text = "" Then
'   Exit Sub
'End If
'RsMst_StorageContract.MoveFirst
'RsMst_StorageContract.Find "ContractName = '" & CmbContractID.Text & "'"
'If RsMst_StorageContract.EOF Then
'   MsgBox "Invalid selection "
'   CmbContractID.SetFocus
'   Exit Sub
'End If
'Call Grid_Head_Client
'mContractID = RsMst_StorageContract!ContractID
'mContractClientID = RsMst_StorageContract!ClientIDRow
'ans = GetClientNameByRow(Val(mContractClientIDRow))
'TxtClientID = ans(0)
'TxtClientName = ans(1)
'TxtDPID = ans(2)
'TxtDPName = ans(3)
'mClientIDRow = mContractClientIDRow
'
'Call CmdClientAdd2List_Click
'CmdClientAdd2List.Enabled = False
'
'Call GetGodownMapped
'
'
'End Sub
Private Sub CmbContractID_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select CMP !!!!! "
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Select Commodity !!!!! "
   Exit Sub
End If

Call Grid_Head_GCTD
Call Grid_Head_Client

tmp1 = CmbContractID.Text

mLocationID = GetLocationIDFromCMP(Val(mCMPID))
'Call TableMst_StorageContract(" where (CommodityID is null or CommodityID = " & mCommodityID & " ) And LocationID = " & mLocationID)
tmp = "Select Distinct MSC.ContractID,MSC.ContractName,MSC.ClientID"
tmp = tmp & " from Mst_StorageContract MSC"
tmp = tmp & " Inner Join Txn_SRGodownMapping TSGM on MSC.ContractID  = TSGM.ContractID"
tmp = tmp & " Inner Join Txn_SRGodownMapDetails TSGMD on TSGM.SRGodownMappingID  = TSGMD.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_Godown MG On TSGMD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Where MSC.ExtendedPeriod > '" & Format(TxtCxTFDate, Gen_DatFormat) & "' And MC.CMPName = '" & CmbCMPID.Text & "' "
tmp = tmp & " And (MSC.CommodityID is null or MSC.CommodityID = " & mCommodityID & " ) And MSC.Flag = 'A' And MSC.SCMFlag = '" & ChkSCMFlag.Value & "' And isnull(MSC.BlockDeposit,0) = '0' "

Call Tmp7Table

CmbContractID.Clear

'If RsMst_StorageContract.RecordCount = 0 Then
'   MsgBox "No Record found !!!! "
'   Check1.Value = 0
'   Exit Sub
'End If

If TmpRs7.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Check1.Value = 0
   TxtCNAVCode = ""
   TxtCAddress = ""
   TxtCGSTNo = ""
   Exit Sub
End If


'For I = 1 To RsMst_StorageContract.RecordCount
'    CmbContractID.AddItem RsMst_StorageContract!ContractName
'    RsMst_StorageContract.MoveNext
'Next

For I = 1 To TmpRs7.RecordCount
    CmbContractID.AddItem TmpRs7!ContractName
    TmpRs7.MoveNext
Next

CmbContractID.Text = tmp1

End Sub
'----
Private Sub CmbContractID_LostFocus()
If CmbContractID.Text = "" Then
   TxtCNAVCode = ""
   TxtCAddress = ""
   TxtCGSTNo = ""
   Exit Sub
End If

'RsMst_StorageContract.MoveFirst
'RsMst_StorageContract.Find "ContractName = '" & CmbContractID.Text & "'"
'If RsMst_StorageContract.EOF Then
'   MsgBox "Invalid selection "
'   CmbContractID.SetFocus
'   Exit Sub
'End If
'Call Grid_Head_Client
'
'mContractID = RsMst_StorageContract!ContractID
'mContractClientIDRow = RsMst_StorageContract!ClientIDRow

TmpRs7.MoveFirst
TmpRs7.Find "ContractName = '" & CmbContractID.Text & "'"
If TmpRs7.EOF Then
   MsgBox "Invalid selection "
   CmbContractID.SetFocus
   Exit Sub
End If
Call Grid_Head_Client

mContractID = TmpRs7!ContractID
mContractClientIDRow = TmpRs7!ClientID



ans = GetClientNameByRow(Val(mContractClientIDRow))
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)
mClientIDRow = mContractClientIDRow

TxtCNAVCode = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "RGNAV")
TxtCAddress = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "a")
TxtCGSTNo = GetCLientGSTDetails(mClientIDRow, Val(mCMPID), "g")


Call CmdClientAdd2List_Click
CmdClientAdd2List.Enabled = False

Call GetGodownMapped


End Sub

'----Added on 30 Nov 2011 411 PM Parag For Reservation Deposit
Private Sub GetGodownMapped()

mReserveGodown = ""

tmp = "Select Distinct MG.GodownNo,MG.GodownID "
tmp = tmp & " From Txn_SRGodownMapDetails TSRGD"
tmp = tmp & " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_UnitType MU On MG.UnitTypeID = MU.UnitTypeID"
tmp = tmp & " Inner Join Txn_SRGodownMapping TSGM On TSRGD.SRGodownMappingID = TSGM.SRGodownMappingID"
tmp = tmp & " Where TSGM.ContractID = " & mContractID & " And MC.CMPID = " & mCMPID & ""

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       If mReserveGodown = "" Then
          mReserveGodown = "(" & TmpRs!GodownID
       Else
          mReserveGodown = mReserveGodown & "," & TmpRs!GodownID
       End If
       TmpRs.MoveNext
   Next
   
   mReserveGodown = mReserveGodown & ")" '& TmpRs!GodownID
End If

End Sub


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

TxtGCTDNoofBags = Format(TxtGCTDNoofBags, "#######0")
TxtGCTDGrossWt = Format(Val(TxtGCTDNoofBags) * mTruckWtRatio, "#####0.000")

Call CalDmatWt
End Sub

Private Sub CalDmatWt()

TxtGCTDGunnyWt = Format((Val(TxtGCTDNoofBags) * Val(TxtCGunnyWt)) / 1000, "#######0.000")
TxtAdjGrossWt = Val(TxtGCTDGrossWt) + Val(TxtAdjWt)
tmp = Val(TxtAdjGrossWt) - Val(TxtGCTDGunnyWt)
TxtGCTDStdDeduction = Format(((tmp) * Val(TxtCStdDed)) / 100, "#######0.000")
TxtGCTDDematWt = Format(Val(TxtAdjGrossWt) - (Val(TxtGCTDStdDeduction) + Val(TxtGCTDGunnyWt)), "#######0.000")
TxtAverageBagSize = 0
If Val(TxtGCTDNoofBags) > 0 Then
   TxtAverageBagSize = Format(Val(TxtGCTDDematWt) * 1000 / Val(TxtGCTDNoofBags), "#######0.00")
End If

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
''TxtTTDGrossWt = Format(Val(TxtTTDTotalWt) - Val(TxtTTDTareWt), "#####0.000")
TxtTTDGrossWt = Format(Val(TxtTTDTotalWt) - (Val(TxtTTDTareWt) + Val(TxtForeignMatter)), "#####0.000") ''--- Ullhas Patil 21 April 2018
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
For I = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!Locationname
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
   mCMPID = 0
End If
mSLocationID = RsMst_Location!LocationID
mLocationName = RsMst_Location!Locationname
msSM_prefix = RsMst_Location!SM_Prefix
End Sub


Private Sub CmbSCmp_GotFocus()
If msSM_prefix = "" Then
   MsgBox "Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "")
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
Private Sub Form_Activate()
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
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True
If Frame4.Width < Me.Width Then HScroll1.Enabled = False
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

'Private Sub PDF_Print_Rpt()
'Set objPDF = New mjwPDF
'
'' Create a simple PDF file using the mjwPDF class
'
'Gen_mTimeStamp = GetTimeStamp
''Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
'
'
'' Set the PDF title and filename
'objPDF.PDFTitle = "CDTF PDF Document"
'objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
'mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
'' We must tell the class where the PDF fonts are located
'objPDF.PDFLoadAfm = App.Path & "\Fonts"
'
'' Set the file properties
'objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
'objPDF.PDFFormatPage = FORMAT_A4
'objPDF.PDFOrientation = ORIENT_PORTRAIT
'objPDF.PDFSetUnit = UNIT_PT
'objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU
'
'
'' Lets us set see the bookmark pane when we view the PDF
''objPDF.PDFUseOutlines = True
'
'' View the PDF file after we create it
'objPDF.PDFView = True
'
'' Begin our PDF document
'
'objPDF.PDFBeginDoc
'    ' Lets add a heading
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'objPDF.PDFCell " ", 15, 15, _
'objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'
'If PntPrv = "Y" Then
'   objPDF.PDFWatermark = "Preview"
'Else
'   objPDF.PDFWatermark = "NCMSL"
'End If
'
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'
'
'pgval = ((objPDF.PDFGetPageWidth - 18))
'
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'
'mRowHeight = 15
'mGlobalLeft = 12
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'mLeft = mGlobalLeft
'mTop = 15
'mHeight = mTop + mRowHeight
'mTxtFromTop = mRowHeight + 2
'
' 'Call DrawTable(10, 2, 15, 30, 12)
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("NCMSL " & TxtCMPID & "")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 12
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 10
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Deposit Transaction Form (CDTF)")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("(This is NOT an instrument of title of the goods)")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 6
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
'
''--Code added By Kishor
''Row 1
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'Dim mDispTextSize  As Integer
'mDispTextSize = 6
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))
''--
'mTxtTop = mTxtTop + mRowHeight
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCmpID))
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "State")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCmpID)))
''--
'mTxtTop = mTxtTop + mRowHeight
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
'Dim mCMPIDnAddr As String
'mCMPIDnAddr = mCmpID & "-" & GetCMPAddress(mCmpID, "A")
'Dim mLen As Integer, mMax As Integer, mPara As Integer
'Dim t1 As Double
'mMax = 35
'mPara = 0
't1 = 1
'mLen = Len(mCMPIDnAddr)
'mLen = Round(mLen / mMax) '+ 1
'
'For t = 1 To mLen
'    If t > 1 Then
'       t1 = mMax + t1
'       'mMax '= mMax + 35
'       mPara = mPara + 7
'    End If
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, Mid$(mCMPIDnAddr, t1, mMax))
'Next
'
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Commodity")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCommodityID)
'
'Call DrawTable(2, 4, mTop, mHeight, mLeft, mRowHeight)
'mHeight = mHeight + 15
'Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
'
'
''Row 2
'
'mTxtTop = mHeight - (mTxtFromTop) + 15 'mTop ' + 15
'mStrLength = Len("Transaction Tyepe")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtExchangeTypeID)
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)
'
'objPDF.PDFDrawLine mGlobalLeft + 132, mTop - 4 * mRowHeight, mGlobalLeft + 132, mHeight + mRowHeight
'Call DrawTable(2, , mTop, mHeight, mLeft, mRowHeight)
'
'
''Row 3
'
'mTxtTop = mHeight '- (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 8
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop - 2, "Depositor Details")
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'objPDF.PDFDrawLine mGlobalLeft + 132, mTop, mGlobalLeft + 132, mTop + mRowHeight
'objPDF.PDFDrawLine 3 * (mGlobalLeft + 132), mTop, 3 * (mGlobalLeft + 132), mTop + mRowHeight
'objPDF.PDFDrawLine 3.5 * (mGlobalLeft + 132), mTop, 3.5 * (mGlobalLeft + 132), mTop + mRowHeight
'
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, "Client ID")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, "Client Name")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, "No of Bags")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 6, "Electronic Weight")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "in MT")
'
'
'objPDF.PDFDrawLine mGlobalLeft + 132, mTop, mGlobalLeft + 132, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
'objPDF.PDFDrawLine 3 * (mGlobalLeft + 132), mTop, 3 * (mGlobalLeft + 132), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'objPDF.PDFDrawLine 3.5 * (mGlobalLeft + 132), mTop, 3.5 * (mGlobalLeft + 132), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'
'Call DrawTable(MSHFlexGrid2.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
'
'
''FILL GRID WITH DETAILS
'For t = 2 To MSHFlexGrid2.Rows - 1
'    mTxtTop = mTxtTop + mRowHeight
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid2.TextMatrix(t, 1))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, MSHFlexGrid2.TextMatrix(t, 2))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, MSHFlexGrid2.TextMatrix(t, 5))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, MSHFlexGrid2.TextMatrix(t, 6))
'Next
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 220, mTxtTop, "Total")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, MSHFlexGrid2.TextMatrix(0, 5))
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, MSHFlexGrid2.TextMatrix(0, 6))
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop - 2, "Truck-wise Deposit Details")
'
'objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + ((MSHFlexGrid3.Rows) * mRowHeight)
'objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
'objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''''
''''
'''''Dim mTempTop As Integer
'''''mTempTop = mTxtTop - 3
'''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), "Total")
''''
'''''----Commented on 18-01-2011 by Sagar
''''
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), "Less Gunny Weight(MT)")
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), "Less Standard Deductions(MT)")
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), "Add Adjustments(MT)")
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Electronic Weight(MT)")
''''
''''objPDF.PDFDrawLine (mGlobalLeft), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight), (mGlobalLeft + 256), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight)
''''
'''''pp 'Hologram No
''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Hologram No")
''''
'''''Actual Values
''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), MSHFlexGrid3.TextMatrix(0, 5))
''''
'''''----Commented on 18-01-2011 by Sagar
''''
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 8))
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 9))
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 6))
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 10))
'''''Hologram No -- Use Record Set
''''
''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), " " & RsTxn_CxTF_Details!HologramNo & " ")
''''
'Dim mTemp As Integer
'For t = 1 To MSHFlexGrid3.Rows - 1
'    If t = 1 Then
'       mTemp = mFontSize
'    Else
'       mTemp = mDispTextSize
'    End If
'    mTxtTop = mTxtTop + mRowHeight
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(t, 1))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(t, 3))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(t, 4))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(t, 5))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(t, 2))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(t, 6))
'Next
'
'Call DrawTable(MSHFlexGrid3.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTop + mRowHeight - 3, "Total")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTop + mRowHeight - 3, MSHFlexGrid3.TextMatrix(0, 5))
'objPDF.PDFDrawLine (mGlobalLeft), mTop + mRowHeight, (mGlobalLeft + 256), mTop + mRowHeight
'mTop = mTop + (mRowHeight * 4)
'
'objPDF.PDFDrawLine mGlobalLeft + 50, mTop + mRowHeight, mGlobalLeft + 50, mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 120), mTop + mRowHeight, (mGlobalLeft + 120), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 190), mTop + mRowHeight, (mGlobalLeft + 190), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 420), mTop + mRowHeight, (mGlobalLeft + 420), mTop + (mRowHeight * 3)
'
'Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 265, mTop - 3, "Affix Hologram Here")
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTop + mRowHeight - 3, "Hologram No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTop + mRowHeight - 3, " " & RsTxn_CxTF_Details!HologramNo & " ")
'
'Call DrawTable(4, , mTop, mTop, mLeft, mRowHeight)
'
'If mHeight > objPDF.PDFGetPageHeight - 50 Then
'   objPDF.PDFEndPage
'   objPDF.PDFNewPage
'   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'   objPDF.PDFSetDrawColor = vbWhite
'   objPDF.PDFSetTextColor = vbBlack
'   objPDF.PDFSetAlignement = ALIGN_CENTER
'   objPDF.PDFSetBorder = BORDER_ALL
'   objPDF.PDFSetFill = True
'
'   objPDF.PDFCell " ", 15, 15, _
'   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'   mLeft = mGlobalLeft
'   mTop = 15
'   mHeight = mTop + mRowHeight
'   mTxtFromTop = mRowHeight + 2
'End If
'mHeight = mTop
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'mTxtTop = mHeight - (mTxtFromTop)
''''
'Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
''''
'mHeight = mHeight + 50
''''
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'mTxtTop = mHeight - (mTxtFromTop) '+ 15
''''
'Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
''''
'mHeight = mHeight + 30
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'mTxtTop = mHeight - (mTxtFromTop)
''''
'Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 17, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
'Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
'Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 3, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
''''
'
'mHeight = mHeight + 50
'mTxtTop = mHeight - (mTxtFromTop)
'
'If mExchangeTypeID = 1 Then
'    Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Note: Please collect from us Warehouse Receipt for the deposited commodities at the earliest. Please note that deposited goods can be withdrawn only against production of valid ")
'    Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 10, "Warehouse Receipt issued by us and not against this CDTF.")
'
'    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'End If
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'mHeight = mHeight + 100
'mTxtTop = mHeight - (mTxtFromTop)
'Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Depositor/Agent of Depositor :")
'
'mHeight = mHeight + 50
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'mTxtTop = mHeight - (mTxtFromTop)
'Call PrintText(8, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "**'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html")
'
'If mHeight > objPDF.PDFGetPageHeight Then
'   objPDF.PDFEndPage
'   objPDF.PDFNewPage
'   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'   objPDF.PDFSetDrawColor = vbWhite
'   objPDF.PDFSetTextColor = vbBlack
'   objPDF.PDFSetAlignement = ALIGN_CENTER
'   objPDF.PDFSetBorder = BORDER_ALL
'   objPDF.PDFSetFill = True
'
'   objPDF.PDFCell " ", 15, 15, _
'   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'   mLeft = mGlobalLeft
'   mTop = 15
'   mHeight = mTop + mRowHeight
'   mTxtFromTop = mRowHeight + 2
'End If
'
'
'objPDF.PDFEndDoc
''    objPDF.PDFNewPage
'
''For GSL Details New Doc--------------------------------------------------------------------------------
'
'
'Set objPDF = New mjwPDF
'
'' Set the PDF title and filename
'objPDF.PDFTitle = "CDTF Godowm Stack Lot Details PDF Document"
'objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & "GSLDetails.pdf"
'mPDFFileName1 = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & "GSLDetails.pdf"
'
'' We must tell the class where the PDF fonts are located
'objPDF.PDFLoadAfm = App.Path & "\Fonts"
'
'' Set the file properties
'objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
'objPDF.PDFFormatPage = FORMAT_A4
'objPDF.PDFOrientation = ORIENT_PORTRAIT
'objPDF.PDFSetUnit = UNIT_PT
'objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU
'
'
'' Lets us set see the bookmark pane when we view the PDF
''objPDF.PDFUseOutlines = True
'
'' View the PDF file after we create it
'objPDF.PDFView = False
'
'' Begin our PDF document
'
'objPDF.PDFBeginDoc
'    ' Lets add a heading
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'objPDF.PDFCell " ", 15, 15, _
'objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'
'If PntPrv = "Y" Then
'   objPDF.PDFWatermark = "Preview"
'Else
'   objPDF.PDFWatermark = "NCMSL"
'End If
'
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'
'
'pgval = ((objPDF.PDFGetPageWidth - 18))
'
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'
'mRowHeight = 15
'mGlobalLeft = 12
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'mLeft = mGlobalLeft
'mTop = 15
'mHeight = mTop + mRowHeight
'mTxtFromTop = mRowHeight + 2
'
' 'Call DrawTable(10, 2, 15, 30, 12)
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("NCMSL " & TxtCMPID & "")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 12
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID & "")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Godown Stack Lot Detail")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength) ' * 2) '* 5) '270
'mFontSize = 10
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Godown Stack Lot Detail")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
''Dim mDispTextSize  As Integer
'mDispTextSize = 6
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCxTFDate)
'
'
'Call DrawTable(1, 4, mTop, mHeight, mLeft, mRowHeight)
'mHeight = mHeight + 15
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Godown No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, "Godown Address")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mmGlobalLeft + 180, mTxtTop, "Stack No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 260, mTxtTop, "Lot No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 340, mTxtTop, "No Of Bags")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 420, mTxtTop, "Gross Weight")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "Client Name")
'
'mHeight = mTop + mRowHeight
'Call DrawTable(1, 7, mTop, mHeight, mLeft, mRowHeight)
''    mHeight = mHeight + 15
'
'
''Call DrawTable(MSHFlexGrid4.Rows - 1, 7, mTop, mHeight, mLeft, mRowHeight)
'
'
''FILL GRID WITH DETAILS
''mHeight = mTop + mRowHeight
'Dim mTxtTopforNextRow As Double
''mValuAddtoVertiLine = 75
'Dim mUpValuAddtoVertiLine As Double
'mUpValuAddtoVertiLine = 75
'mDownValuAddtoVertiLine = mUpValuAddtoVertiLine
'blnnew = True
'For t = 2 To MSHFlexGrid4.Rows - 2
'    If MSHFlexGrid4.TextMatrix(t, 1) = "" Then Exit For
'    If blnnew = True Then
'       mTxtTop = mTxtTop + mRowHeight ' * 2.8
'    Else
'       mTxtTop = mTxtTopforNextRow + mRowHeight
'    End If
'    blnnew = False
''        mTxtTop = mHeight - (mTxtFromTop) + 15
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid4.TextMatrix(t, 1))
''        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, MSHFlexGrid4.TextMatrix(t, 18))
'
'    mMax = 20
'    mPara = 0
'
'    t1 = 1
'    mLen = Len(MSHFlexGrid4.TextMatrix(t, 18))
'    mLen = Round(mLen / mMax) '+ 1
'
'    For x = 1 To mLen
'        If x > 1 Then
'           t1 = mMax + t1
'           'mMax '= mMax + 35
'           mPara = mPara + 7
'        End If
'        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop + mPara, Mid$(MSHFlexGrid4.TextMatrix(t, 18), t1, mMax))
'    Next
'
'
'    objPDF.PDFDrawLine mGlobalLeft + 70, mUpValuAddtoVertiLine, mGlobalLeft + 70, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine      ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 152.5, mUpValuAddtoVertiLine, mGlobalLeft + 152.5, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine      ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 235.25, mUpValuAddtoVertiLine, mGlobalLeft + 235.25, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine     ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 317.5, mUpValuAddtoVertiLine, mGlobalLeft + 317.5, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine     ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 400.15, mUpValuAddtoVertiLine, mGlobalLeft + 400.15, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 482.5, mUpValuAddtoVertiLine, mGlobalLeft + 482.5, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    'objPDF.PDFDrawLine mLeft + 45, mLeft + 75, mLeft + 45, mTxtTopforNextRow + 75   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'
'    mDownValuAddtoVertiLine = mUpValuAddtoVertiLine
'    mTxtTopforNextRow = mTxtTop + mPara
'
'    mUpValuAddtoVertiLine = mTxtTop + mPara + 8
'
'
'
''        objPDF.PDFDrawLine mGlobalLeft + 70, mGlobalLeft + 70, mGlobalLeft + 70, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 152.5, mGlobalLeft + 152.5, mGlobalLeft + 152.5, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 235.25, mGlobalLeft + 235.25, mGlobalLeft + 235.25, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 317.5, mGlobalLeft + 317.5, mGlobalLeft + 317.5, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 400.15, mGlobalLeft + 400.15, mGlobalLeft + 400.15, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 482.5, mGlobalLeft + 482.5, mGlobalLeft + 482.5, mTxtTopforNextRow + mValuAddtoVertiLine  ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        'objPDF.PDFDrawLine mLeft + 45, mLeft + 75, mLeft + 45, mTxtTopforNextRow + 75   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'
'
'
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop, MSHFlexGrid4.TextMatrix(t, 2))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 260, mTxtTop, MSHFlexGrid4.TextMatrix(t, 3))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 340, mTxtTop, MSHFlexGrid4.TextMatrix(t, 4))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 420, mTxtTop, MSHFlexGrid4.TextMatrix(t, 5))
'
'    mMax = 15
'    mPara = 0
'    t1 = 1
'    mLen = Len(MSHFlexGrid4.TextMatrix(t, 11))
'    mLen = Round(mLen / mMax) '+ 1
'
'    For x = 1 To mLen
'        If x > 1 Then
'           t1 = mMax + t1
'           'mMax '= mMax + 35
'           mPara = mPara + 7
'        End If
'        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 490, mTxtTop + mPara, Mid$(MSHFlexGrid4.TextMatrix(t, 11), t1, mMax))
'    Next
'
'
'    mHeight = mTop + mRowHeight
'
'    'Call DrawTable(MSHFlexGrid4.Rows - 1, 7, mTop, mHeight, mLeft, mRowHeight)
'
'    objPDF.PDFDrawLine mGlobalLeft, mTxtTopforNextRow + 8, mGlobalLeft + (objPDF.PDFGetPageWidth - 30), mTxtTopforNextRow + 8 ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'
'    If mTxtTopforNextRow > (objPDF.PDFGetPageHeight - 50) Then 'objPDF.PDFGetPageHeight Then
'       objPDF.PDFEndPage
'       objPDF.PDFNewPage
'
'    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'    objPDF.PDFSetDrawColor = vbWhite
'    objPDF.PDFSetTextColor = vbBlack
'    objPDF.PDFSetAlignement = ALIGN_CENTER
'    objPDF.PDFSetBorder = BORDER_ALL
'    objPDF.PDFSetFill = True
'
'    objPDF.PDFCell " ", 15, 15, _
'    objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'
''    If PntPrv = "Y" Then
''       objPDF.PDFWatermark = "Preview"
''    Else
''       objPDF.PDFWatermark = "NCMSL"
''    End If
'
''    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
''    objPDF.PDFSetDrawColor = vbWhite
''    objPDF.PDFSetTextColor = vbBlack
''    objPDF.PDFSetAlignement = ALIGN_CENTER
''    objPDF.PDFSetBorder = BORDER_ALL
''    objPDF.PDFSetFill = True
''
'
'    pgval = ((objPDF.PDFGetPageWidth - 18))
'
'    'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'
'    mRowHeight = 15
'    mGlobalLeft = 12
'    'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'    mLeft = mGlobalLeft
'    mTop = 15
'    mHeight = mTop + mRowHeight
'    mTxtFromTop = mRowHeight + 2
'
'     'Call DrawTable(10, 2, 15, 30, 12)
'
'    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'    mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'    mStrLength = Len("NCMSL " & TxtCMPID & "")
'    mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'    mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'    mFontSize = 12
'
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID & "")
'
'       mHeight = mTop + mRowHeight
'       mLeft = mGlobalLeft
'
'
'       Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'       mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'       mStrLength = Len("Godown Stack Lot Detail")
'       mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'       mTxtLeft = mPrintPoint - (mStrLength) ' * 2) '* 5) '270
'       mFontSize = 10
'
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Godown Stack Lot Detail")
'
'       mHeight = mTop + mRowHeight
'       mLeft = mGlobalLeft
'
'
'       mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'       mStrLength = Len("Depositor Details")
'       mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'       mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'       mFontSize = 7
'       'Dim mDispTextSize  As Integer
'       mDispTextSize = 6
'
'       mTxtTop = mTxtTop + mRowHeight
'
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
'       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
'       '--
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
'       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCxTFDate)
'
'
'
'       Call DrawTable(1, 4, mTop, mHeight, mLeft, mRowHeight)
'       mHeight = mHeight + 15
'
'       mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'       mStrLength = Len("Depositor Details")
'       mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'       mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'       mFontSize = 7
'
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Godown No")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, "Godown Address")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mmGlobalLeft + 180, mTxtTop, "Stack No")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 260, mTxtTop, "Lot No")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 340, mTxtTop, "No Of Bags")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 420, mTxtTop, "Gross Weight")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "Client Name")
'
'       mHeight = mTop + mRowHeight
'       Call DrawTable(1, 7, mTop, mHeight, mLeft, mRowHeight)
'       blnnew = True
'       mTxtTopforNextRow = 0
'       mUpValuAddtoVertiLine = 75
'       mDownValuAddtoVertiLine = mUpValuAddtoVertiLine
'   End If
'
'Next
'
'mTxtTop = mTxtTop + mRowHeight
'
''mTxtTop = mTxtTop + mRowHeight
'
'objPDF.PDFEndDoc
'
'End Sub

Private Sub DrawTable(Optional mROw_ As Double, Optional mCol_ As Double, Optional mTop_ As Double, Optional mHeight_ As Double, Optional mLeft_ As Double, Optional mRowHeight_ As Double)

If mROw_ = 0 Then mROw_ = 1
If mCol_ = 0 Then mCol_ = 1
If mLeft_ = 0 Then mLeft_ = mmasterLeft
If mTop_ = 0 Then mTop_ = 0
If mHeight_ = 0 Then mHeight_ = mTop_ + mRowHeight_

Dim pgval_ As Double

For ir = 1 To mROw_
    pgval_ = pgval
    pgval_ = pgval_ / mCol_
    mWidth_ = pgval_
    For ic = 1 To mCol_
        objPDF.PDFDrawPolygon Array(mLeft_, mTop_, mWidth_, mTop_, mWidth_, mHeight_, mLeft_, mHeight_)
        mLeft_ = pgval_ * ic
        mWidth_ = pgval_ * (ic + 1)
    Next
    mTop_ = mHeight_
    mHeight_ = mHeight_ + mRowHeight_
    'mLeft_ = mmasterLeft

    mLeft_ = mGlobalLeft

Next

mTop = mTop_
mRow = mROw_ + 1

End Sub

'
Private Sub PrintText(mFontSize_ As Integer, mFontFormat_ As PDFFontStl, mAlign_ As Variant, mTxtLeft_ As Double, mTxtTop_ As Double, mTexttoPrint As String)

'        objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'        objPDF.PDFSetAlignement = ALIGN_CENTER
        
objPDF.PDFSetFont FONT_ARIAL, mFontSize_, mFontFormat_
objPDF.PDFSetAlignement = mAlign_
objPDF.PDFTextOut mTexttoPrint, mTxtLeft_, mTxtTop_


End Sub
Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
TOADD = ""
TOCC = ""


mMsg = "<table border = 1><tr><td colspan=2>CDTF Details (From Store-Man)</td></tr>"
mMsg = mMsg & "<tr><td>CDTF No</td><td>" & TxtCxTFNo.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>CDTF Date</td><td>" & Format(TxtCxTFDate.Value, "dd-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & "<tr><td>Location</td><td>" & GetLocationFromCMP(RsTxn_CxTF_Details!CMPID) & "</td> </tr>"
mMsg = mMsg & "<tr><td>State</td><td>" & GetStateFromLocation(GetLocationIDFromCMP(mCMPID)) & "</td></tr>"
mMsg = mMsg & "<tr><td>CMPID & Address</td><td>" & mCMPID & "-" & GetCMPAddress(mCMPID, "A") & "</td></tr>"
mMsg = mMsg & "<tr><td>Commodity</td><td>" & TxtCommodityID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Transaction Type</td><td>" & TxtExchangeTypeID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Name Of Agent </td><td>" & TxtAgent.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Total No of Bags</td><td>" & MSHFlexGrid2.TextMatrix(0, 5) & "</td></tr>"
mMsg = mMsg & "<tr><td>Total Electronic Weight in MT </td><td>" & MSHFlexGrid2.TextMatrix(0, 6) & "</td></tr>"
mMsg = mMsg & "<tr><td>Hologram No</td><td>" & IIf(IsNull(RsTxn_CxTF_Details!HologramNo), "", RsTxn_CxTF_Details!HologramNo) & "</td></tr></table>"

tmp = " Select EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID,EMP1.EmployeeID as 'AMEmpID',EMP1.EmployeeNo as 'AMEmpNo',EMP1.EmployeeName as 'AMName',EMP1.EmployeeEmailID as 'AMEmailID'" & _
      " From Mst_EmpCMPMapping EGM" & _
      " Inner Join Mst_Employee EMP on EGM.EmployeeID = EMP.EmployeeID" & _
      " Inner Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID" & _
      " Where EMP.Flag = 'Y' And EGM.CMPID = " & mCMPID
'      " Group By EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID"

Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
'   TOADD = TmpRs3!EmployeeEmailID
'   TOCC = IIf(IsNull(TmpRs3!AMEmailID), "", TmpRs3!AMEmailID) '""

   For ab = 1 To TmpRs3.RecordCount
       
        If TmpRs3!EmployeeEmailID <> "" Then
           If InStr(1, TmpRs3!EmployeeEmailID, "@") Then
              If TOADD <> "" Then
                 If InStr(1, LCase(TOADD), LCase(TmpRs3!EmployeeEmailID)) = 0 Then
                    TOADD = TOADD & "," & TmpRs3!EmployeeEmailID
                 End If
              Else
                 TOADD = TmpRs3!EmployeeEmailID
              End If
           End If
        End If
       
        If TmpRs3!AMEmailID <> "" Then
           If InStr(1, TmpRs3!AMEmailID, "@") Then
              If TOCC <> "" Then
                 If InStr(1, LCase(TOCC), LCase(TmpRs3!AMEmailID)) = 0 Then
                    TOCC = TOCC & "," & TmpRs3!AMEmailID
                 End If
              Else
                 TOCC = TmpRs3!AMEmailID
              End If
           End If
        End If
       
'       TOADD = TmpRs3!EmployeeEmailID
'       TOCC = IIf(IsNull(TmpRs3!AMEmailID), "", TmpRs3!AMEmailID) '""
       TmpRs3.MoveNext
   Next


End If

If TOADD = "" Then
   TOADD = "kuldip.s@ncml.com"
End If
 
 '"cmg@ncml.com" & "," & Gen_UserEmailID  'Gen_UserEmailID
If TOCC <> "" Then
   TOCC = TOCC & "," & Gen_UserEmailID
Else
   TOCC = Gen_UserEmailID
End If

If ChkSCMFlag.Value = 0 Then
   mMsgSubject = "NCMSL " & TxtCMPID.Text & " CDTF Details"
Else
   TOCC = TOCC & "," & "supplychain@ncml.com,scmquality@ncml.com"
   mMsgSubject = "NCMSL CDTF No-" & TxtCxTFNo.Text & " for " & TxtCMPID.Text & "-Agent name: " & TxtAgent.Text & ""
End If

'TOADD = "parag.r@ncml.com"
'TOCC = "koyal.d@ncml.com,ganesh.h@ncml.com"

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "N")

End Sub

'Private Sub Gen_Email()
'Dim TOADD As Variant
'Dim TOCC, mMsg, mGridMsg As Variant
'Dim strSplitString() As String
'Dim EmpNo_, EmpName_ As String
'
'mMsg = ""
'TOADD = ""
'TOCC = ""
'
'
'mMsg = "<tr><td colspan=2>CDTF Details (From Store-Man)</td></tr>"
'mMsg = mMsg & "<tr><td>CDTF No</td><td>" & TxtCxTFNo.Text & "</td> </tr>"
'mMsg = mMsg & "<tr><td>CDTF Date</td><td>" & Format(TxtCxTFDate.Value, "dd-MMM-yyyy") & "</td></tr>"
'mMsg = mMsg & "<tr><td>Location</td><td>" & GetLocationFromCMP(RsTxn_CxTF_Details!CMPID) & "</td> </tr>"
'mMsg = mMsg & "<tr><td>State</td><td>" & GetStateFromLocation(GetLocationIDFromCMP(mCMPID)) & "</td></tr>"
'mMsg = mMsg & "<tr><td>CMPID & Address</td><td>" & mCMPID & "-" & GetCMPAddress(mCMPID, "A") & "</td></tr>"
'mMsg = mMsg & "<tr><td>Commodity</td><td>" & TxtCommodityID.Text & "</td></tr>"
'mMsg = mMsg & "<tr><td>Transaction Type</td><td>" & TxtExchangeTypeID.Text & "</td></tr>"
'mMsg = mMsg & "<tr><td>Name Of Agent </td><td>" & TxtAgent.Text & "</td></tr>"
'mMsg = mMsg & "<tr><td>Total No of Bags</td><td>" & MSHFlexGrid2.TextMatrix(0, 5) & "</td></tr>"
'mMsg = mMsg & "<tr><td>Total Electronic Weight in MT </td><td>" & MSHFlexGrid2.TextMatrix(0, 6) & "</td></tr>"
'mMsg = mMsg & "<tr><td>Hologram No</td><td>" & IIf(IsNull(RsTxn_CxTF_Details!HologramNo), "", RsTxn_CxTF_Details!HologramNo) & "</td></tr>"
'
'tmp = " Select EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID,EMP1.EmployeeID as 'AMEmpID',EMP1.EmployeeNo as 'AMEmpNo',EMP1.EmployeeName as 'AMName',EMP1.EmployeeEmailID as 'AMEmailID'" & _
'      " From Mst_EmpCMPMapping EGM" & _
'      " Inner Join Mst_Employee EMP on EGM.EmployeeID = EMP.EmployeeID" & _
'      " Inner Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID" & _
'      " Where EMP.Flag = 'Y' And EGM.CMPID = " & mCMPID
''      " Group By EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID"
'
'Call Tmp3Table
'If TmpRs3.RecordCount > 0 Then
'   TOADD = TmpRs3!EmployeeEmailID
'   TOCC = IIf(IsNull(TmpRs3!AMEmailID), "", TmpRs3!AMEmailID) '""
'End If
'
'If TOADD = "" Then
'   TOADD = "kuldip.s@ncml.com"
'End If
'
' '"cmg@ncml.com" & "," & Gen_UserEmailID  'Gen_UserEmailID
'If TOCC <> "" Then
'   TOCC = TOCC & "," & Gen_UserEmailID
'Else
'   TOCC = Gen_UserEmailID
'End If
'mMsgSubject = "NCMSL " & TxtCMPID.Text & " CDTF Details"
'
'Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)
'
'
'End Sub

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant, mMailTypeFlg As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Candara color = Blue>" '<table border = 1>"
htmlStringtail = "</Font></body></html>" '</table>
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  If LCase(mMailTypeFlg) = "n" Then
'     .MailFrom = "info@ncml.com" 'Gen_UserEmailID 'TOCC_ '"mayur.d@ncml.com"
'  Else
'     .MailFrom = Gen_UserEmailID
'  End If
'  .SendTo = ToADD_
''  TOCC_ = TOCC_ '& "," & mVendorEmailID
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
''  .BCC = "mayur.d@ncml.com"
'  If LCase(mMailTypeFlg) = "n" Then
'     .Attachments.Add (mPDFFileName)
'     .Attachments.Add (mPDFFileName1)
'  End If
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"

If LCase(mMailTypeFlg) = "n" Then
   objSMTP.FromAddr = "info@ncml.com"
Else
   objSMTP.FromAddr = Gen_UserEmailID
End If

If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If
If LCase(mMailTypeFlg) = "n" Then
   objSMTP.AddAttachment mPDFFileName
   objSMTP.AddAttachment mPDFFileName1
End If

objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function

'Private Sub PDF_Print_Rpt()
'Set objPDF = New mjwPDF
'
'' Create a simple PDF file using the mjwPDF class
'
'Gen_mTimeStamp = GetTimeStamp
''Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
'
'
'' Set the PDF title and filename
'objPDF.PDFTitle = "CDTF PDF Document"
'objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
'mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
'' We must tell the class where the PDF fonts are located
'objPDF.PDFLoadAfm = App.Path & "\Fonts"
'
'' Set the file properties
'objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
'objPDF.PDFFormatPage = FORMAT_A4
'objPDF.PDFOrientation = ORIENT_PORTRAIT
'objPDF.PDFSetUnit = UNIT_PT
'objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU
'
'' Lets us set see the bookmark pane when we view the PDF
''objPDF.PDFUseOutlines = True
'
'' View the PDF file after we create it
'objPDF.PDFView = True
'
'' Begin our PDF document
'
'objPDF.PDFBeginDoc
'    ' Lets add a heading
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'objPDF.PDFCell " ", 15, 15, _
'objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'
'If PntPrv = "Y" Then
'   objPDF.PDFWatermark = "Preview"
'Else
'   objPDF.PDFWatermark = "NCMSL"
'End If
'
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'pgval = ((objPDF.PDFGetPageWidth - 18))
'
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'
'mRowHeight = 15
'mGlobalLeft = 12
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'mLeft = mGlobalLeft
'mTop = 15
'mHeight = mTop + mRowHeight
'mTxtFromTop = mRowHeight + 2
'
' 'Call DrawTable(10, 2, 15, 30, 12)
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("NCMSL " & TxtCMPID & "")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 12
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 10
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Deposit Transaction Form (CDTF)")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'tmp1 = mTop
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("(This is NOT an instrument of title of the goods)")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 6
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
'
''--Code added By Kishor
''Row 1
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'Dim mDispTextSize  As Integer
'mDispTextSize = 6
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
''--
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Date")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))
'
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCMPID))
''--
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "State")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCMPID)))
''--
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
'Dim mCMPIDnAddr As String
'mCMPIDnAddr = mCMPID & "-" & GetCMPAddress(mCMPID, "A")
'Dim mLen As Integer, mMax As Integer, mPara As Integer
'Dim t1 As Double
'mMax = 45
'mPara = 0
't1 = 1
'mLen = Len(mCMPIDnAddr)
'mLen = Round(mLen / mMax) '+ 1
'
'For t = 1 To mLen
'    If t > 1 Then
'       t1 = mMax + t1
'       'mMax '= mMax + 35
'       mPara = mPara + 7
'    End If
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, Mid$(mCMPIDnAddr, t1, mMax))
'Next
'
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Commodity")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, TxtCommodityID)
'
'mTxtTop = mTxtTop + mPara + mRowHeight
'
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'
'objPDF.PDFDrawLine mGlobalLeft + 345, Val(tmp1), mGlobalLeft + 345, mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 445, Val(tmp1), mGlobalLeft + 445, mTxtTop - 10
'
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtExchangeTypeID)
'
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)
'
''objPDF.PDFDrawLine mGlobalLeft + 132, mTop - 4 * mRowHeight, mGlobalLeft + 132, mHeight + mRowHeight
'
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 135, Val(tmp1), mGlobalLeft + 135, mTxtTop - 10
'
'''Call DrawTable(2, , mTop, mHeight, mLeft, mRowHeight)
'
'Call PrintText(mFontSize, FONT_BOLD & FONT_UNDERLINE, ALIGN_LEFT, mLeft, mTxtTop - 2, "Depositor Details")
'
'''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'mTxtTop = mTxtTop + mRowHeight
'
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'
'''objPDF.PDFDrawLine mGlobalLeft + 132, mTop, mGlobalLeft + 132, mTop + mRowHeight
'''objPDF.PDFDrawLine 3 * (mGlobalLeft + 132), mTop, 3 * (mGlobalLeft + 132), mTop + mRowHeight
'''objPDF.PDFDrawLine 3.5 * (mGlobalLeft + 132), mTop, 3.5 * (mGlobalLeft + 132), mTop + mRowHeight
'''
'''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''
'''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'''mStrLength = Len("Depositor Details")
'''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'''mFontSize = 7
'
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Client ID")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, "Client Name")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, "No of Bags")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "Gross Weight")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop + 10, "in MT")
'
'mTxtTop = mTxtTop + mRowHeight + 10
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 135, mTxtTop - (mRowHeight * 2 + 5), mGlobalLeft + 135, mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 425, mTxtTop - (mRowHeight * 2 + 5), mGlobalLeft + 425, mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 495, mTxtTop - (mRowHeight * 2 + 5), mGlobalLeft + 495, mTxtTop - 10
'
'
'Dim mTemp As Integer
'
''FILL GRID WITH DETAILS
'For t = 2 To MSHFlexGrid2.Rows - 1
'    If mHeight > objPDF.PDFGetPageHeight Then
'       Call CheckForNewPage
'       mTop = 15
'       mHeight = mHeight - mRowHeight
'       mTxtTop = 10
'    End If
'    'mTxtTop = mTxtTop + mRowHeight
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid2.TextMatrix(t, 1))
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, MSHFlexGrid2.TextMatrix(t, 2))
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, MSHFlexGrid2.TextMatrix(t, 5))
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, MSHFlexGrid2.TextMatrix(t, 6))
'
'    mTxtTop = mTxtTop + mRowHeight
'    objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'    objPDF.PDFDrawLine mGlobalLeft + 135, mTxtTop - (mRowHeight + 10), mGlobalLeft + 135, mTxtTop - 10
'    objPDF.PDFDrawLine mGlobalLeft + 425, mTxtTop - (mRowHeight + 10), mGlobalLeft + 425, mTxtTop - 10
'    objPDF.PDFDrawLine mGlobalLeft + 495, mTxtTop - (mRowHeight + 10), mGlobalLeft + 495, mTxtTop - 10
'    mTop = mTop + mRowHeight
'    mHeight = mHeight + mRowHeight
'Next
'
'Call CheckForNewPage
'mTxtTop = mTxtTop - mRowHeight
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 220, mTxtTop, "Total")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, MSHFlexGrid2.TextMatrix(0, 5))
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, MSHFlexGrid2.TextMatrix(0, 6))
'
'Call CheckForNewPage
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_BOLD & FONT_UNDERLINE, ALIGN_LEFT, mLeft, mTxtTop, "Truck-wise Deposit Details")
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Truck No")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Weighbridge Slip No")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Totol Weight")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop + 10, "in MT")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, "Tare Weight")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop + 10, "in MT")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
'
'mTxtTop = mTxtTop + mRowHeight + 10
'
'objPDF.PDFDrawLine mGlobalLeft + 65, mTxtTop - (mRowHeight + 20), mGlobalLeft + 65, (mTxtTop - 10)
'objPDF.PDFDrawLine (mGlobalLeft + 145), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 145), (mTxtTop - 10)
'objPDF.PDFDrawLine (mGlobalLeft + 195), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 195), (mTxtTop - 10)
'objPDF.PDFDrawLine (mGlobalLeft + 245), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 245), (mTxtTop - 10)
'objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
'objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'
'For t = 2 To MSHFlexGrid3.Rows - 1
'    If mHeight > objPDF.PDFGetPageHeight Then
'       Call CheckForNewPage
'       mTop = 15
'       mTxtTop = 10
'    End If
'
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(t, 1))
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, MSHFlexGrid3.TextMatrix(t, 2))
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, MSHFlexGrid3.TextMatrix(t, 3))
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(t, 4))
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(t, 5))
'    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(t, 6))
'    'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, MSHFlexGrid3.TextMatrix(t, 7))
'    mMax = 62
'    mPara = 7
'    t1 = 1
'    mLen = Len(MSHFlexGrid3.TextMatrix(t, 7))
'
'    If mLen > 70 Then
'       mLen = RoundUp(mLen / mMax)
'    Else
'       mLen = 1
'    End If
'    For S = 1 To mLen
'        If S > 1 Then
'           t1 = mMax + t1
'           mPara = mPara + 7
'        End If
'        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop + mPara - 5, Mid$(MSHFlexGrid3.TextMatrix(t, 7), t1, mMax))
'    Next
'
'    a1 = mPara + 5
'
'    If a1 > mRowHeight Then
'       tmp = a1 + 5
'    Else
'       tmp = mRowHeight
'    End If
'    mTxtTop = mTxtTop + tmp
'    objPDF.PDFDrawLine mGlobalLeft + 65, mTxtTop - (tmp + 10), mGlobalLeft + 65, (mTxtTop - 10)
'    objPDF.PDFDrawLine (mGlobalLeft + 145), mTxtTop - (tmp + 10), (mGlobalLeft + 145), (mTxtTop - 10)
'    objPDF.PDFDrawLine (mGlobalLeft + 195), mTxtTop - (tmp + 10), (mGlobalLeft + 195), (mTxtTop - 10)
'    objPDF.PDFDrawLine (mGlobalLeft + 245), mTxtTop - (tmp + 10), (mGlobalLeft + 245), (mTxtTop - 10)
'    objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (tmp + 10), (mGlobalLeft + 300), (mTxtTop - 10)
'    objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (tmp + 10), (mGlobalLeft + 355), (mTxtTop - 10)
'    objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'    mTop = mTop + mRowHeight
'    mHeight = mHeight + mRowHeight
'Next
'mTxtTop = mTxtTop - mRowHeight
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Total")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(0, 5))
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(0, 6))
'
'Call CheckForNewPage
'
'mTxtTop = mTxtTop + (mRowHeight * 4)
'
'Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Affix Hologram Here")
'mTxtTop = mTxtTop + mRowHeight
'
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight * 4 + 10), mGlobalLeft + 300, (mTxtTop - 10)
'
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Hologram No")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 405, mTxtTop, " " & RsTxn_CxTF_Details!HologramNo & " ")
'
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight + 10), mGlobalLeft + 300, (mTxtTop - 10)
'objPDF.PDFDrawLine mGlobalLeft + 400, mTxtTop - (mRowHeight + 10), mGlobalLeft + 400, (mTxtTop - 10)
'
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 100, mTxtTop - (mRowHeight + 10), mGlobalLeft + 100, (mTxtTop - 10)
'objPDF.PDFDrawLine mGlobalLeft + 200, mTxtTop - (mRowHeight + 10), mGlobalLeft + 200, (mTxtTop - 10)
'objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight + 10), mGlobalLeft + 300, (mTxtTop - 10)
'objPDF.PDFDrawLine mGlobalLeft + 400, mTxtTop - (mRowHeight + 10), mGlobalLeft + 400, (mTxtTop - 10)
'
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 100, mTxtTop - (mRowHeight + 10), mGlobalLeft + 100, (mTxtTop - 10)
'objPDF.PDFDrawLine mGlobalLeft + 200, mTxtTop - (mRowHeight + 10), mGlobalLeft + 200, (mTxtTop - 10)
'objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight + 10), mGlobalLeft + 300, (mTxtTop - 10)
'objPDF.PDFDrawLine mGlobalLeft + 400, mTxtTop - (mRowHeight + 10), mGlobalLeft + 400, (mTxtTop - 10)
'
'mTxtTop = mTxtTop + mRowHeight
'Call CheckForNewPage
'
'Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'
''''mHeight = mHeight + 50
''''Call CheckForNewPage
''''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''''mTxtTop = mHeight - (mTxtFromTop) '+ 15
'Call CheckForNewPage
'mTxtTop = mTxtTop + mRowHeight * 3
'
'Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'
'Call CheckForNewPage
'
'Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
'Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 8, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
'Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 16, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
'
'
'mTxtTop = mTxtTop + (mRowHeight * 4) + 16
'
'''''''
''''
''''''
''''''
''''''mHeight = mHeight + 50
''''''mTxtTop = mHeight - (mTxtFromTop)
''''''
''''''If mExchangeTypeID = 1 Then
''''''   Call CheckForNewPage
''''''    Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Note: Please collect from us Warehouse Receipt for the deposited commodities at the earliest. Please note that deposited goods can be withdrawn only against production of valid ")
''''''    Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 10, "Warehouse Receipt issued by us and not against this CDTF.")
''''''
''''''    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''''''
''''''End If
'''''''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''''''mHeight = mHeight + 100
''''''mTxtTop = mHeight - (mTxtFromTop)
''''''Call CheckForNewPage
''''''Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Depositor/Agent of Depositor :")
''''''
''''''mHeight = mHeight + 50
''''''mTxtTop = mHeight - (mTxtFromTop)
''''''Call CheckForNewPage
''''
''''mHeight = mHeight + 50
''''mTxtTop = mHeight - (mTxtFromTop)
'Call CheckForNewPage
'Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Depositor/Agent of Depositor :")
'mTxtTop = mTxtTop + mRowHeight
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'
''''mHeight = mHeight + 50
''''mTxtTop = mHeight - (mTxtFromTop)
'
'If mExchangeTypeID = 1 Then
'   Call CheckForNewPage
'   mHeight = mHeight + 50
'
'   Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Note: Please collect from us Warehouse Receipt for the deposited commodities at the earliest. Please note that deposited goods can be withdrawn only against production of valid ")
'   Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 8, "Warehouse Receipt issued by us and not against this CDTF.")
'   mTxtTop = mTxtTop + mRowHeight + 8
'End If
'''''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''''
'Call CheckForNewPage
''''
''''mHeight = mHeight + 50
''''mTxtTop = mHeight - (mTxtFromTop)
''''
'Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "**'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html")
''''
'objPDF.PDFEndDoc
''    objPDF.PDFNewPage
'
''For GSL Details New Doc--------------------------------------------------------------------------------
'
'
'Set objPDF = New mjwPDF
'
'' Set the PDF title and filename
'objPDF.PDFTitle = "CDTF Godowm Stack Lot Details PDF Document"
'objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & "GSLDetails.pdf"
'mPDFFileName1 = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & "GSLDetails.pdf"
'
'' We must tell the class where the PDF fonts are located
'objPDF.PDFLoadAfm = App.Path & "\Fonts"
'
'' Set the file properties
'objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
'objPDF.PDFFormatPage = FORMAT_A4
'objPDF.PDFOrientation = ORIENT_PORTRAIT
'objPDF.PDFSetUnit = UNIT_PT
'objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU
'
'
'' Lets us set see the bookmark pane when we view the PDF
''objPDF.PDFUseOutlines = True
'
'' View the PDF file after we create it
'objPDF.PDFView = False
'
'' Begin our PDF document
'
'objPDF.PDFBeginDoc
'    ' Lets add a heading
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'objPDF.PDFCell " ", 15, 15, _
'objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'
'If PntPrv = "Y" Then
'   objPDF.PDFWatermark = "Preview"
'Else
'   objPDF.PDFWatermark = "NCMSL"
'End If
'
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'
'
'pgval = ((objPDF.PDFGetPageWidth - 18))
'
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'
'mRowHeight = 15
'mGlobalLeft = 12
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'mLeft = mGlobalLeft
'mTop = 15
'mHeight = mTop + mRowHeight
'mTxtFromTop = mRowHeight + 2
'
' 'Call DrawTable(10, 2, 15, 30, 12)
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("NCMSL " & TxtCMPID & "")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 12
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID & "")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Godown Stack Lot Detail")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength) ' * 2) '* 5) '270
'mFontSize = 7
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Godown Stack Lot Detail")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
''--
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCxTFDate)
'
'
'Call DrawTable(1, 4, mTop, mHeight, mLeft, mRowHeight)
'mHeight = mHeight + 15
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Godown No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTxtTop, "Godown Address")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Stack No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, "Lot No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "No Of Bags")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 300, mTxtTop, "Gross Weight")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Client Name")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, "Quality parameters")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop + 10, "/ Remarks")
'
'mHeight = mTop + mRowHeight
'
'mTxtTop = mTxtTop + mRowHeight + 10
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 45, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 45, mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 145, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 145, mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 195, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 195, mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 245, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 245, mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 295, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 295, mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 345, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 345, mTxtTop - 10
'objPDF.PDFDrawLine mGlobalLeft + 445, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 445, mTxtTop - 10
'
''Call DrawTable(1, 7, mTop, mHeight, mLeft, mRowHeight)
''    mHeight = mHeight + 15
'
'
''Call DrawTable(MSHFlexGrid4.Rows - 1, 7, mTop, mHeight, mLeft, mRowHeight)
'
'
''FILL GRID WITH DETAILS
''mHeight = mTop + mRowHeight
'Dim mTxtTopforNextRow As Double
''mValuAddtoVertiLine = 75
'Dim mUpValuAddtoVertiLine As Double
'mUpValuAddtoVertiLine = 75
'mDownValuAddtoVertiLine = mUpValuAddtoVertiLine
'blnnew = True
'mTxtTop = mTxtTop - mRowHeight
'For t = 2 To MSHFlexGrid4.Rows - 2
'    If MSHFlexGrid4.TextMatrix(t, 1) = "" Then Exit For
'    If blnnew = True Then
'       mTxtTop = mTxtTop + mRowHeight ' * 2.8
'    Else
'       mTxtTop = mTxtTopforNextRow + mRowHeight
'    End If
'    blnnew = False
''        mTxtTop = mHeight - (mTxtFromTop) + 15
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid4.TextMatrix(t, 1))
''        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, MSHFlexGrid4.TextMatrix(t, 18))
'
'    mMax = 20
'    mPara = 0
'
'    t1 = 1
'    mLen = Len(MSHFlexGrid4.TextMatrix(t, 18))
'    mLen = Round(mLen / mMax) '+ 1
'
'    For x = 1 To mLen
'        If x > 1 Then
'           t1 = mMax + t1
'           'mMax '= mMax + 35
'           mPara = mPara + 7
'        End If
'        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTxtTop + mPara, Mid$(MSHFlexGrid4.TextMatrix(t, 18), t1, mMax))
'    Next
'
'
'    objPDF.PDFDrawLine mGlobalLeft + 45, mUpValuAddtoVertiLine, mGlobalLeft + 45, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine      ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 145, mUpValuAddtoVertiLine, mGlobalLeft + 145, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine      ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 195, mUpValuAddtoVertiLine, mGlobalLeft + 195, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine     ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 245, mUpValuAddtoVertiLine, mGlobalLeft + 245, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine     ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 295, mUpValuAddtoVertiLine, mGlobalLeft + 295, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 345, mUpValuAddtoVertiLine, mGlobalLeft + 345, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft + 445, mUpValuAddtoVertiLine, mGlobalLeft + 445, mTxtTop + mPara + 8
'    'objPDF.PDFDrawLine mLeft + 45, mLeft + 75, mLeft + 45, mTxtTopforNextRow + 75   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'
'    mDownValuAddtoVertiLine = mUpValuAddtoVertiLine
'    mTxtTopforNextRow = mTxtTop + mPara
'
'    mUpValuAddtoVertiLine = mTxtTop + mPara + 8
'
'
'
''        objPDF.PDFDrawLine mGlobalLeft + 70, mGlobalLeft + 70, mGlobalLeft + 70, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 152.5, mGlobalLeft + 152.5, mGlobalLeft + 152.5, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 235.25, mGlobalLeft + 235.25, mGlobalLeft + 235.25, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 317.5, mGlobalLeft + 317.5, mGlobalLeft + 317.5, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 400.15, mGlobalLeft + 400.15, mGlobalLeft + 400.15, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        objPDF.PDFDrawLine mGlobalLeft + 482.5, mGlobalLeft + 482.5, mGlobalLeft + 482.5, mTxtTopforNextRow + mValuAddtoVertiLine  ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''        'objPDF.PDFDrawLine mLeft + 45, mLeft + 75, mLeft + 45, mTxtTopforNextRow + 75   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'
'
'
'
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, MSHFlexGrid4.TextMatrix(t, 2))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid4.TextMatrix(t, 3))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid4.TextMatrix(t, 4))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 300, mTxtTop, MSHFlexGrid4.TextMatrix(t, 5))
'    mMax = 15
'    mPara = 0
'    t1 = 1
'    mLen = Len(MSHFlexGrid4.TextMatrix(t, 11))
'    mLen = Round(mLen / mMax) '+ 1
'
'    For x = 1 To mLen
'        If x > 1 Then
'           t1 = mMax + t1
'           'mMax '= mMax + 35
'           mPara = mPara + 7
'        End If
'        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop + mPara, Mid$(MSHFlexGrid4.TextMatrix(t, 11), t1, mMax))
'    Next
'
'
'    mMax = 15
'    mPara = 0
'    t1 = 1
'    mLen = Len(MSHFlexGrid4.TextMatrix(t, 12))
'    mLen = Round(mLen / mMax) '+ 1
'
'    For x = 1 To mLen
'        If x > 1 Then
'           t1 = mMax + t1
'           'mMax '= mMax + 35
'           mPara = mPara + 7
'        End If
'        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop + mPara, Mid$(MSHFlexGrid4.TextMatrix(t, 12), t1, mMax))
'    Next
'
'    mHeight = mTop + mRowHeight
'
'    'Call DrawTable(MSHFlexGrid4.Rows - 1, 7, mTop, mHeight, mLeft, mRowHeight)
'
'    objPDF.PDFDrawLine mGlobalLeft, mTxtTopforNextRow + 8, mGlobalLeft + (objPDF.PDFGetPageWidth - 30), mTxtTopforNextRow + 8 ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'
'    If mTxtTopforNextRow > (objPDF.PDFGetPageHeight - 50) Then 'objPDF.PDFGetPageHeight Then
'        objPDF.PDFWatermark = ""
'
'       objPDF.PDFEndPage
'       objPDF.PDFNewPage
'
'    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'    objPDF.PDFSetDrawColor = vbWhite
'    objPDF.PDFSetTextColor = vbBlack
'    objPDF.PDFSetAlignement = ALIGN_CENTER
'    objPDF.PDFSetBorder = BORDER_ALL
'    objPDF.PDFSetFill = True
'
'    objPDF.PDFCell " ", 15, 15, _
'    objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'
'    If PntPrv = "Y" Then
'       objPDF.PDFWatermark = "Preview"
'    Else
'       objPDF.PDFWatermark = "NCMSL"
'    End If
'
'    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'    objPDF.PDFSetDrawColor = vbWhite
'    objPDF.PDFSetTextColor = vbBlack
'    objPDF.PDFSetAlignement = ALIGN_CENTER
'    objPDF.PDFSetBorder = BORDER_ALL
'    objPDF.PDFSetFill = True
'
'    pgval = ((objPDF.PDFGetPageWidth - 18))
'
'    'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'
'    mRowHeight = 15
'    mGlobalLeft = 12
'    'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'    mLeft = mGlobalLeft
'    mTop = 15
'    mHeight = mTop + mRowHeight
'    mTxtFromTop = mRowHeight + 2
'
'     'Call DrawTable(10, 2, 15, 30, 12)
'
'    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'    mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'    mStrLength = Len("NCMSL " & TxtCMPID & "")
'    mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'    mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'    mFontSize = 12
'
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID & "")
'
'       mHeight = mTop + mRowHeight
'       mLeft = mGlobalLeft
'
'
'       Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'       mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'       mStrLength = Len("Godown Stack Lot Detail")
'       mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'       mTxtLeft = mPrintPoint - (mStrLength) ' * 2) '* 5) '270
'       mFontSize = 10
'
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Godown Stack Lot Detail")
'
'       mHeight = mTop + mRowHeight
'       mLeft = mGlobalLeft
'
'
'       mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'       mStrLength = Len("Depositor Details")
'       mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'       mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'       mFontSize = 7
'       'Dim mDispTextSize  As Integer
'       mDispTextSize = 6
'
'       mTxtTop = mTxtTop + mRowHeight
'
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
'       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
'       '--
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
'       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCxTFDate)
'
'
'
'       Call DrawTable(1, 4, mTop, mHeight, mLeft, mRowHeight)
'       mHeight = mHeight + 15
'
'       mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'       mStrLength = Len("Depositor Details")
'       mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'       mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'       mFontSize = 7
'
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Godown No")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, "Godown Address")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mmGlobalLeft + 180, mTxtTop, "Stack No")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 260, mTxtTop, "Lot No")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 340, mTxtTop, "No Of Bags")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 420, mTxtTop, "Gross Weight")
'       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "Client Name")
'
'       mHeight = mTop + mRowHeight
'       Call DrawTable(1, 7, mTop, mHeight, mLeft, mRowHeight)
'       blnnew = True
'       mTxtTopforNextRow = 0
'       mUpValuAddtoVertiLine = 75
'       mDownValuAddtoVertiLine = mUpValuAddtoVertiLine
'   End If
'
'Next
'
'mTxtTop = mTxtTop + mRowHeight
'
''mTxtTop = mTxtTop + mRowHeight
'
'objPDF.PDFEndDoc
'
'End Sub

Private Sub PDF_Print_Rpt()
'Dim Totpages As Double
Totpages = GetTotalPDFPages()
Set objPDF = New mjwPDF
' Create a simple PDF file using the mjwPDF class
Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
   
' Set the PDF title and filename
objPDF.PDFTitle = "CDTF PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
' We must tell the class where the PDF fonts are located
objPDF.PDFLoadAfm = App.Path & "\Fonts"

' Set the file properties
objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
objPDF.PDFFormatPage = FORMAT_A4
objPDF.PDFOrientation = ORIENT_PORTRAIT
objPDF.PDFSetUnit = UNIT_PT
objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU

' Lets us set see the bookmark pane when we view the PDF
'objPDF.PDFUseOutlines = True

' View the PDF file after we create it
objPDF.PDFView = True
' Begin our PDF document

objPDF.PDFBeginDoc
    ' Lets add a heading
objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

objPDF.PDFCell " ", 15, 15, _
objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20

If PntPrv = "Y" Then
   objPDF.PDFWatermark = "Preview"
Else
   objPDF.PDFWatermark = "NCML" '"NCMSL"
End If

objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

pgval = ((objPDF.PDFGetPageWidth - 18))

'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)

mRowHeight = 15
mGlobalLeft = 12
'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
mLeft = mGlobalLeft
mTop = 15
mHeight = mTop + mRowHeight
mTxtFromTop = mRowHeight + 2
 
'Call DrawTable(10, 2, 15, 30, 12)
Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("NCMSL " & TxtCMPID & "")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 12

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")

If Totpages > 1 Then
   Call PrintText(6, FONT_NORMAL, ALIGN_RIGHT, objPDF.PDFGetPageWidth - 70, mTxtTop - 5, "Page " & objPDF.PDFPageNumber & " out of " & Totpages)
Else
   
End If
mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Deposit Transaction Form (CDTF)")

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
 
tmp1 = mTop
mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("(This is NOT an instrument of title of the goods)")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 6

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")

'--Code added By Kishor
'Row 1

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Depositor Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7
Dim mDispTextSize  As Integer
mDispTextSize = 6

mTxtTop = mTxtTop + mRowHeight

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Date")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCMPID))
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "State")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCMPID)))
'--
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
Dim mCMPIDnAddr As String
mCMPIDnAddr = mCMPID & "-" & GetCMPAddress(mCMPID, "A")
Dim mLen As Integer, mMax As Integer, mPara As Integer
Dim t1 As Double
mMax = 45
mPara = 0
t1 = 1
mLen = Len(mCMPIDnAddr)
mLen = Round(mLen / mMax) '+ 1

For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       'mMax '= mMax + 35
       mPara = mPara + 7
    End If
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, mID$(mCMPIDnAddr, t1, mMax))
Next

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Commodity")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, TxtCommodityID)

mTxtTop = mTxtTop + mPara + mRowHeight

objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

objPDF.PDFDrawLine mGlobalLeft + 345, Val(tmp1), mGlobalLeft + 345, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 445, Val(tmp1), mGlobalLeft + 445, mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
If ChkSCMFlag.Value = 1 And Check1.Value = 1 Then
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtExchangeTypeID & " - Reservation and SCM Deposit")
ElseIf ChkSCMFlag.Value = 1 And Check1.Value = 0 Then
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtExchangeTypeID & " - SCM Deposit")
ElseIf ChkSCMFlag.Value = 0 And Check1.Value = 1 Then
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtExchangeTypeID & " - Reservation Deposit")
ElseIf ChkSCMFlag.Value = 0 And Check1.Value = 0 Then
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtExchangeTypeID)
End If

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)

'objPDF.PDFDrawLine mGlobalLeft + 132, mTop - 4 * mRowHeight, mGlobalLeft + 132, mHeight + mRowHeight

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 135, Val(tmp1), mGlobalLeft + 135, mTxtTop - 10

''Call DrawTable(2, , mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD & FONT_UNDERLINE, ALIGN_LEFT, mLeft, mTxtTop - 2, "Depositor Details")

''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
mTxtTop = mTxtTop + mRowHeight

objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

''objPDF.PDFDrawLine mGlobalLeft + 132, mTop, mGlobalLeft + 132, mTop + mRowHeight
''objPDF.PDFDrawLine 3 * (mGlobalLeft + 132), mTop, 3 * (mGlobalLeft + 132), mTop + mRowHeight
''objPDF.PDFDrawLine 3.5 * (mGlobalLeft + 132), mTop, 3.5 * (mGlobalLeft + 132), mTop + mRowHeight
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("Depositor Details")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
''mFontSize = 7

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Client ID")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, "Client Name")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, "No of Bags")
If LCase(TxtExchangeTypeID) = "ncdex" Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 495, mTxtTop, "Electronic weight")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 495, mTxtTop + 10, "after Applicable")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 495, mTxtTop + 20, "deductions in MT")
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "Gross Weight")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop + 10, "in MT")
End If
If LCase(TxtExchangeTypeID) = "ncdex" Then
   mTxtTop = mTxtTop + mRowHeight + 20
Else
   mTxtTop = mTxtTop + mRowHeight + 10
End If

If LCase(TxtExchangeTypeID) = "ncdex" Then
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
   objPDF.PDFDrawLine mGlobalLeft + 135, mTxtTop - (mRowHeight * 2 + 15), mGlobalLeft + 135, mTxtTop - 10
   objPDF.PDFDrawLine mGlobalLeft + 425, mTxtTop - (mRowHeight * 2 + 15), mGlobalLeft + 425, mTxtTop - 10
   objPDF.PDFDrawLine mGlobalLeft + 495, mTxtTop - (mRowHeight * 2 + 15), mGlobalLeft + 495, mTxtTop - 10
Else
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
   objPDF.PDFDrawLine mGlobalLeft + 135, mTxtTop - (mRowHeight * 2 + 5), mGlobalLeft + 135, mTxtTop - 10
   objPDF.PDFDrawLine mGlobalLeft + 425, mTxtTop - (mRowHeight * 2 + 5), mGlobalLeft + 425, mTxtTop - 10
   objPDF.PDFDrawLine mGlobalLeft + 495, mTxtTop - (mRowHeight * 2 + 5), mGlobalLeft + 495, mTxtTop - 10
End If

Dim mTemp As Integer

'FILL GRID WITH DETAILS
For T = 2 To MSHFlexGrid2.Rows - 1
    If mHeight > objPDF.PDFGetPageHeight - 50 Then
       Call CheckForNewPage
       mTop = 15
       mHeight = mHeight - mRowHeight
       mTxtTop = 10
    End If
    'mTxtTop = mTxtTop + mRowHeight
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid2.TextMatrix(T, 1))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, MSHFlexGrid2.TextMatrix(T, 2))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, MSHFlexGrid2.TextMatrix(T, 5))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, MSHFlexGrid2.TextMatrix(T, 6))

    mTxtTop = mTxtTop + mRowHeight
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
    objPDF.PDFDrawLine mGlobalLeft + 135, mTxtTop - (mRowHeight + 10), mGlobalLeft + 135, mTxtTop - 10
    objPDF.PDFDrawLine mGlobalLeft + 425, mTxtTop - (mRowHeight + 10), mGlobalLeft + 425, mTxtTop - 10
    objPDF.PDFDrawLine mGlobalLeft + 495, mTxtTop - (mRowHeight + 10), mGlobalLeft + 495, mTxtTop - 10
    mTop = mTop + mRowHeight
    mHeight = mHeight + mRowHeight
Next

Call CheckForNewPage
mTxtTop = mTxtTop - mRowHeight
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 220, mTxtTop, "Total")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, MSHFlexGrid2.TextMatrix(0, 5))
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, MSHFlexGrid2.TextMatrix(0, 6))

Call CheckForNewPage
mTxtTop = mTxtTop + mRowHeight

Call PrintText(mFontSize, FONT_BOLD & FONT_UNDERLINE, ALIGN_LEFT, mLeft, mTxtTop, "Truck-wise Deposit Details")
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Truck No")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Weighbridge Slip No")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Totol Weight")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop + 10, "in MT")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, "Tare Weight")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop + 10, "in MT")
''-- Ullhas Patil 21 APril 2018
If mSiloLocation = True Then
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Foreign")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "Matter")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Net Weight")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop + 10, "in MT")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "No Of Bags")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 410, mTxtTop, "Quality Parametes / Remark")
Else
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
End If
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
''-----------------------------

mTxtTop = mTxtTop + mRowHeight + 10

objPDF.PDFDrawLine mGlobalLeft + 65, mTxtTop - (mRowHeight + 20), mGlobalLeft + 65, (mTxtTop - 10)
objPDF.PDFDrawLine (mGlobalLeft + 145), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 145), (mTxtTop - 10)
objPDF.PDFDrawLine (mGlobalLeft + 195), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 195), (mTxtTop - 10)
objPDF.PDFDrawLine (mGlobalLeft + 245), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 245), (mTxtTop - 10)
''-- Ullhas Patil 21 APril 2018
If mSiloLocation = True Then
    objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 405), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 405), (mTxtTop - 10)
Else
    objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
End If
''objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
''objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
''-----------------------------
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

mHeight = mTxtTop ' mTop + mRowHeight
'mTxtTop = mHeight
For T = 2 To MSHFlexGrid3.Rows - 1

    If mHeight > objPDF.PDFGetPageHeight - 50 Then
       Call CheckForNewPage
       mTop = 15
       
      mRowHeight = 15
      mGlobalLeft = 12
      'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
      mLeft = mGlobalLeft
      mTop = 15
      mHeight = mTop + mRowHeight
      mTxtFromTop = mRowHeight + 2
      
      'Call DrawTable(10, 2, 15, 30, 12)
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("NCMSL " & TxtCMPID & "")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
      mFontSize = 12
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
      
      mHeight = mTop + mRowHeight
      mLeft = mGlobalLeft
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
      mFontSize = 10
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Deposit Transaction Form (CDTF)")
      
      mHeight = mTop + mRowHeight
      mLeft = mGlobalLeft
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      tmp1 = mTop
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("(This is NOT an instrument of title of the goods)")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
      mFontSize = 6
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
      
      '--Code added By Kishor
      'Row 1
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("Depositor Details")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
      mFontSize = 7
      'Dim mDispTextSize  As Integer
      mDispTextSize = 6
      
      mTxtTop = mTxtTop + mRowHeight
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
      '--
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Date")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))
      
      mTxtTop = mTxtTop + mRowHeight
      
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCMPID))
      '--
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "State")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCMPID)))
      '--
      mTxtTop = mTxtTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

     
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtExchangeTypeID)
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Commodity")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, TxtCommodityID)
      
      mTxtTop = mTxtTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

      objPDF.PDFDrawLine mGlobalLeft + 345, Val(tmp1), mGlobalLeft + 345, mTxtTop - 10
      objPDF.PDFDrawLine mGlobalLeft + 445, Val(tmp1), mGlobalLeft + 445, mTxtTop - 10
      objPDF.PDFDrawLine mGlobalLeft + 135, Val(tmp1), mGlobalLeft + 135, mTxtTop - 10
      
      mTxtTop = mTxtTop + (mRowHeight * 2)
      
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Truck No")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Weighbridge Slip No")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Totol Weight")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop + 10, "in MT")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, "Tare Weight")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop + 10, "in MT")
        ''-- Ullhas Patil 21 APril 2018
        If mSiloLocation = True Then
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Foreign")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "Matter")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Net Weight")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop + 10, "in MT")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "No Of Bags")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 410, mTxtTop, "Quality Parametes / Remark")
        Else
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
        End If
        ''        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
        ''        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
        ''        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
        ''        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
        ''-----------------------------
      
      mTxtTop = mTxtTop + mRowHeight + 10
      mHeight = mTxtTop
      objPDF.PDFDrawLine mGlobalLeft + 65, mTxtTop - (mRowHeight + 20), mGlobalLeft + 65, (mTxtTop - 10)
      objPDF.PDFDrawLine (mGlobalLeft + 145), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 145), (mTxtTop - 10)
      objPDF.PDFDrawLine (mGlobalLeft + 195), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 195), (mTxtTop - 10)
      objPDF.PDFDrawLine (mGlobalLeft + 245), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 245), (mTxtTop - 10)
      ''-- Ullhas Patil 21 APril 2018
      If mSiloLocation = True Then
        objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 405), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 405), (mTxtTop - 10)
      Else
        objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
      End If
      ''      objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
      ''      objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
      ''-----------------------------
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

    End If
   
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(T, 1))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, MSHFlexGrid3.TextMatrix(T, 2))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, MSHFlexGrid3.TextMatrix(T, 3))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(T, 4))
    ''-- Ullhas Patil 21 APril 2018
    If mSiloLocation = True Then
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(T, 13))
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(T, 5))
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, MSHFlexGrid3.TextMatrix(T, 6))
    Else
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(T, 5))
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(T, 6))
    End If
    ''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(T, 5))
    ''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(T, 6))
    ''------------------------------
    'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, MSHFlexGrid3.TextMatrix(t, 7))
    
    mMax = 62
    mPara = 7
    t1 = 1
    mLen = Len(MSHFlexGrid3.TextMatrix(T, 7))
    
    If mLen > 70 Then
       mLen = RoundUp(mLen / mMax)
    Else
       mLen = 1
    End If
    For S = 1 To mLen
        If S > 1 Then
           t1 = mMax + t1
           mPara = mPara + 7
        End If
        ''-- Ullhas Patil 21 APril 2018
        If mSiloLocation = True Then
            Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 410, mTxtTop + mPara - 5, mID$(MSHFlexGrid3.TextMatrix(T, 7), t1, mMax))
        Else
            Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop + mPara - 5, mID$(MSHFlexGrid3.TextMatrix(T, 7), t1, mMax))
        End If
        ''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop + mPara - 5, mID$(MSHFlexGrid3.TextMatrix(T, 7), t1, mMax))
        ''------------------------------
    Next
    
    a1 = mPara + 5
    
    If a1 > mRowHeight Then
       tmp = a1 + 5
    Else
       tmp = mRowHeight
    End If
    mTxtTop = mTxtTop + tmp
    objPDF.PDFDrawLine mGlobalLeft + 65, mTxtTop - (tmp + 10), mGlobalLeft + 65, (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 145), mTxtTop - (tmp + 10), (mGlobalLeft + 145), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 195), mTxtTop - (tmp + 10), (mGlobalLeft + 195), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 245), mTxtTop - (tmp + 10), (mGlobalLeft + 245), (mTxtTop - 10)
    ''-- Ullhas Patil 21 APril 2018
    If mSiloLocation = True Then
        objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (tmp + 10), (mGlobalLeft + 300), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (tmp + 10), (mGlobalLeft + 355), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 405), mTxtTop - (tmp + 10), (mGlobalLeft + 405), (mTxtTop - 10)
    Else
        objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (tmp + 10), (mGlobalLeft + 300), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (tmp + 10), (mGlobalLeft + 355), (mTxtTop - 10)
    End If
    ''objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (tmp + 10), (mGlobalLeft + 300), (mTxtTop - 10)
    ''objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (tmp + 10), (mGlobalLeft + 355), (mTxtTop - 10)
    ''------------------------------
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
    
    mTop = mTop + mRowHeight
    mHeight = mHeight + mRowHeight
Next
mTxtTop = mTxtTop - mRowHeight
mHeight = mTxtTop

''-- Ullhas Patil 21 APril 2018
If mSiloLocation = True Then
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Total")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(0, 13))
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(0, 5))
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, MSHFlexGrid3.TextMatrix(0, 6))
Else
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Total")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(0, 5))
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(0, 6))
End If
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Total")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(0, 5))
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(0, 6))
''------------------------------

Call CheckForNewPage

mTxtTop = mTxtTop + (mRowHeight * 4)
mHeight = mTxtTop
Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Affix Hologram Here")
mTxtTop = mTxtTop + mRowHeight

objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight * 4 + 10), mGlobalLeft + 300, (mTxtTop - 10)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Hologram No")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 405, mTxtTop, " " & RsTxn_CxTF_Details!HologramNo & " ")

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight + 10), mGlobalLeft + 300, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 400, mTxtTop - (mRowHeight + 10), mGlobalLeft + 400, (mTxtTop - 10)

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 100, mTxtTop - (mRowHeight + 10), mGlobalLeft + 100, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 200, mTxtTop - (mRowHeight + 10), mGlobalLeft + 200, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight + 10), mGlobalLeft + 300, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 400, mTxtTop - (mRowHeight + 10), mGlobalLeft + 400, (mTxtTop - 10)

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 100, mTxtTop - (mRowHeight + 10), mGlobalLeft + 100, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 200, mTxtTop - (mRowHeight + 10), mGlobalLeft + 200, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight + 10), mGlobalLeft + 300, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 400, mTxtTop - (mRowHeight + 10), mGlobalLeft + 400, (mTxtTop - 10)

mTxtTop = mTxtTop + mRowHeight
mHeight = mTxtTop
Call CheckForNewPage

Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

'''mHeight = mHeight + 50
'''Call CheckForNewPage
'''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''mTxtTop = mHeight - (mTxtFromTop) '+ 15
mHeight = mTxtTop
Call CheckForNewPage
mTxtTop = mTxtTop + mRowHeight * 3

Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
mHeight = mTxtTop
Call CheckForNewPage

Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 8, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 16, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")

'mHeight = mHeight + 50

mTxtTop = mTxtTop + mRowHeight + 8
mTxtTop = mTxtTop + mRowHeight
'mTxtTop = mTxtTop + mRowHeight + 8

mHeight = mTxtTop
Call CheckForNewPage
mHeight = mHeight + 50
   
Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "I / We confirm that the commodities deposited by me/us conform to all applicable rules, regulations and notifications of FSSAI, Essential Commodities Act (ECA), State Warehousing Acts  and other statutory acts ")
Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 8, "applicable on storage and warehousing of agricultural commodities.")
mTxtTop = mTxtTop + mRowHeight + 8

mTxtTop = mTxtTop + (mRowHeight * 4) + 16

''''''
'''
'''''
'''''
'''''mHeight = mHeight + 50
'''''mTxtTop = mHeight - (mTxtFromTop)
'''''
'''''If mExchangeTypeID = 1 Then
'''''   Call CheckForNewPage
'''''    Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Note: Please collect from us Warehouse Receipt for the deposited commodities at the earliest. Please note that deposited goods can be withdrawn only against production of valid ")
'''''    Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 10, "Warehouse Receipt issued by us and not against this CDTF.")
'''''
'''''    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''''
'''''End If
''''''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''''mHeight = mHeight + 100
'''''mTxtTop = mHeight - (mTxtFromTop)
'''''Call CheckForNewPage
'''''Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Depositor/Agent of Depositor :")
'''''
'''''mHeight = mHeight + 50
'''''mTxtTop = mHeight - (mTxtFromTop)
'''''Call CheckForNewPage
'''
'''mHeight = mHeight + 50
'''mTxtTop = mHeight - (mTxtFromTop)
mHeight = mTxtTop
Call CheckForNewPage
Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Depositor/Agent of Depositor :")
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

'''mHeight = mHeight + 50
'''mTxtTop = mHeight - (mTxtFromTop)


If mExchangeTypeID = 1 Then
   mHeight = mTxtTop
   Call CheckForNewPage
   mHeight = mHeight + 50
   
   Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Note: Please collect from us Warehouse Receipt for the deposited commodities at the earliest. Please note that deposited goods can be withdrawn only against production of valid ")
   Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 8, "Warehouse Receipt issued by us and not against this CDTF.")
   mTxtTop = mTxtTop + mRowHeight + 8
End If
''''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
mHeight = mTxtTop
Call CheckForNewPage
'''
'''mHeight = mHeight + 50
'''mTxtTop = mHeight - (mTxtFromTop)
'''
Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "**'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html")
'''

If mExchangeTypeID = 1 Then
   Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 510, mTxtTop + (20 * 9), "F&F 112")
ElseIf mExchangeTypeID = 2 Then
   Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 510, mTxtTop + (20 * 9), "F&F 111")
End If

For I = 1 To objPDF.PDFPageNumber
    
Next
objPDF.PDFEndDoc
'    objPDF.PDFNewPage

'For GSL Details New Doc--------------------------------------------------------------------------------

   
Set objPDF = New mjwPDF
 
' Set the PDF title and filename
objPDF.PDFTitle = "CDTF Godowm Stack Lot Details PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & "GSLDetails.pdf"
mPDFFileName1 = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & "GSLDetails.pdf"

' We must tell the class where the PDF fonts are located
objPDF.PDFLoadAfm = App.Path & "\Fonts"

' Set the file properties
objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
objPDF.PDFFormatPage = FORMAT_A4
objPDF.PDFOrientation = ORIENT_PORTRAIT
objPDF.PDFSetUnit = UNIT_PT
objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU


' Lets us set see the bookmark pane when we view the PDF
'objPDF.PDFUseOutlines = True

' View the PDF file after we create it
objPDF.PDFView = False

' Begin our PDF document

objPDF.PDFBeginDoc
    ' Lets add a heading
objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

objPDF.PDFCell " ", 15, 15, _
objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20

If PntPrv = "Y" Then
   objPDF.PDFWatermark = "Preview"
Else
   objPDF.PDFWatermark = "NCML" '"NCMSL"
End If

objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True



pgval = ((objPDF.PDFGetPageWidth - 18))

'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)

mRowHeight = 15
mGlobalLeft = 12
'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
mLeft = mGlobalLeft
mTop = 15
mHeight = mTop + mRowHeight
mTxtFromTop = mRowHeight + 2
 
 'Call DrawTable(10, 2, 15, 30, 12)

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("NCMSL " & TxtCMPID & "")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 12

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID & "")

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Godown Stack Lot Detail")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength) ' * 2) '* 5) '270
mFontSize = 7

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Godown Stack Lot Detail")

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft


mTxtTop = mTxtTop + mRowHeight

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCxTFDate)


Call DrawTable(1, 4, mTop, mHeight, mLeft, mRowHeight)
mHeight = mHeight + 15

mTxtTop = mTxtTop + mRowHeight

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Godown No")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTxtTop, "Godown Address")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Stack No")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, "Lot No")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "No Of Bags")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 300, mTxtTop, "Gross Weight")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Client Name")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, "Quality parameters")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop + 10, "/ Remarks")
                    
mHeight = mTop + mRowHeight

mTxtTop = mTxtTop + mRowHeight + 10
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 45, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 45, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 145, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 145, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 195, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 195, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 245, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 245, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 295, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 295, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 345, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 345, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 445, mTxtTop - (mRowHeight * 2 + 8), mGlobalLeft + 445, mTxtTop - 10

'Call DrawTable(1, 7, mTop, mHeight, mLeft, mRowHeight)
'    mHeight = mHeight + 15
                    
                    
'Call DrawTable(MSHFlexGrid4.Rows - 1, 7, mTop, mHeight, mLeft, mRowHeight)


'FILL GRID WITH DETAILS
'mHeight = mTop + mRowHeight
Dim mTxtTopforNextRow As Double
'mValuAddtoVertiLine = 75
Dim mUpValuAddtoVertiLine As Double
mUpValuAddtoVertiLine = 75
mDownValuAddtoVertiLine = mUpValuAddtoVertiLine
blnnew = True
mTxtTop = mTxtTop - mRowHeight
For T = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(T, 1) = "" Then Exit For
    If blnnew = True Then
       mTxtTop = mTxtTop + mRowHeight ' * 2.8
    Else
       mTxtTop = mTxtTopforNextRow + mRowHeight
    End If
    blnnew = False
'        mTxtTop = mHeight - (mTxtFromTop) + 15
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid4.TextMatrix(T, 1))
'        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, MSHFlexGrid4.TextMatrix(t, 18))
           
    mMax = 20
    mPara = 0

    t1 = 1
    mLen = Len(MSHFlexGrid4.TextMatrix(T, 18))
    mLen = Round(mLen / mMax) '+ 1

    For x = 1 To mLen
        If x > 1 Then
           t1 = mMax + t1
           'mMax '= mMax + 35
           mPara = mPara + 7
        End If
        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTxtTop + mPara, mID$(MSHFlexGrid4.TextMatrix(T, 18), t1, mMax))
    Next
    
    
    objPDF.PDFDrawLine mGlobalLeft + 45, mUpValuAddtoVertiLine, mGlobalLeft + 45, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine      ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
    objPDF.PDFDrawLine mGlobalLeft + 145, mUpValuAddtoVertiLine, mGlobalLeft + 145, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine      ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
    objPDF.PDFDrawLine mGlobalLeft + 195, mUpValuAddtoVertiLine, mGlobalLeft + 195, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine     ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
    objPDF.PDFDrawLine mGlobalLeft + 245, mUpValuAddtoVertiLine, mGlobalLeft + 245, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine     ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
    objPDF.PDFDrawLine mGlobalLeft + 295, mUpValuAddtoVertiLine, mGlobalLeft + 295, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
    objPDF.PDFDrawLine mGlobalLeft + 345, mUpValuAddtoVertiLine, mGlobalLeft + 345, mTxtTop + mPara + 8 '+ mDownValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
    objPDF.PDFDrawLine mGlobalLeft + 445, mUpValuAddtoVertiLine, mGlobalLeft + 445, mTxtTop + mPara + 8
    'objPDF.PDFDrawLine mLeft + 45, mLeft + 75, mLeft + 45, mTxtTopforNextRow + 75   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)

    mDownValuAddtoVertiLine = mUpValuAddtoVertiLine
    mTxtTopforNextRow = mTxtTop + mPara

    mUpValuAddtoVertiLine = mTxtTop + mPara + 8
    
         
   
'        objPDF.PDFDrawLine mGlobalLeft + 70, mGlobalLeft + 70, mGlobalLeft + 70, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'        objPDF.PDFDrawLine mGlobalLeft + 152.5, mGlobalLeft + 152.5, mGlobalLeft + 152.5, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'        objPDF.PDFDrawLine mGlobalLeft + 235.25, mGlobalLeft + 235.25, mGlobalLeft + 235.25, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'        objPDF.PDFDrawLine mGlobalLeft + 317.5, mGlobalLeft + 317.5, mGlobalLeft + 317.5, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'        objPDF.PDFDrawLine mGlobalLeft + 400.15, mGlobalLeft + 400.15, mGlobalLeft + 400.15, mTxtTopforNextRow + mValuAddtoVertiLine   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'        objPDF.PDFDrawLine mGlobalLeft + 482.5, mGlobalLeft + 482.5, mGlobalLeft + 482.5, mTxtTopforNextRow + mValuAddtoVertiLine  ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'        'objPDF.PDFDrawLine mLeft + 45, mLeft + 75, mLeft + 45, mTxtTopforNextRow + 75   ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
    
    
    
    
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, MSHFlexGrid4.TextMatrix(T, 2))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid4.TextMatrix(T, 3))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid4.TextMatrix(T, 4))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 300, mTxtTop, MSHFlexGrid4.TextMatrix(T, 5))
    mMax = 15
    mPara = 0
    t1 = 1
    mLen = Len(MSHFlexGrid4.TextMatrix(T, 11))
    mLen = Round(mLen / mMax) '+ 1

    For x = 1 To mLen
        If x > 1 Then
           t1 = mMax + t1
           'mMax '= mMax + 35
           mPara = mPara + 7
        End If
        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop + mPara, mID$(MSHFlexGrid4.TextMatrix(T, 11), t1, mMax))
    Next
    
   
    mMax = 15
    mPara = 0
    t1 = 1
    mLen = Len(MSHFlexGrid4.TextMatrix(T, 12))
    mLen = Round(mLen / mMax) '+ 1

    For x = 1 To mLen
        If x > 1 Then
           t1 = mMax + t1
           'mMax '= mMax + 35
           mPara = mPara + 7
        End If
        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop + mPara, mID$(MSHFlexGrid4.TextMatrix(T, 12), t1, mMax))
    Next
       
    mHeight = mTop + mRowHeight
    
    'Call DrawTable(MSHFlexGrid4.Rows - 1, 7, mTop, mHeight, mLeft, mRowHeight)
    
    objPDF.PDFDrawLine mGlobalLeft, mTxtTopforNextRow + 8, mGlobalLeft + (objPDF.PDFGetPageWidth - 30), mTxtTopforNextRow + 8 ' + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
    
    If mTxtTopforNextRow > (objPDF.PDFGetPageHeight - 50) Then 'objPDF.PDFGetPageHeight Then
        objPDF.PDFWatermark = ""
       
       objPDF.PDFEndPage
       objPDF.PDFNewPage
      
    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
    objPDF.PDFSetDrawColor = vbWhite
    objPDF.PDFSetTextColor = vbBlack
    objPDF.PDFSetAlignement = ALIGN_CENTER
    objPDF.PDFSetBorder = BORDER_ALL
    objPDF.PDFSetFill = True
    
    objPDF.PDFCell " ", 15, 15, _
    objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
    
    If PntPrv = "Y" Then
       objPDF.PDFWatermark = "Preview"
    Else
       objPDF.PDFWatermark = "NCML" '"NCMSL"
    End If
    
    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
    objPDF.PDFSetDrawColor = vbWhite
    objPDF.PDFSetTextColor = vbBlack
    objPDF.PDFSetAlignement = ALIGN_CENTER
    objPDF.PDFSetBorder = BORDER_ALL
    objPDF.PDFSetFill = True
    
    pgval = ((objPDF.PDFGetPageWidth - 18))
    
    'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
    
    mRowHeight = 15
    mGlobalLeft = 12
    'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
    mLeft = mGlobalLeft
    mTop = 15
    mHeight = mTop + mRowHeight
    mTxtFromTop = mRowHeight + 2
     
     'Call DrawTable(10, 2, 15, 30, 12)
    
    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
    
    mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
    mStrLength = Len("NCMSL " & TxtCMPID & "")
    mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
    mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
    mFontSize = 12

       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID & "")
       
       mHeight = mTop + mRowHeight
       mLeft = mGlobalLeft
   
   
       Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
       mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
       mStrLength = Len("Godown Stack Lot Detail")
       mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
       mTxtLeft = mPrintPoint - (mStrLength) ' * 2) '* 5) '270
       mFontSize = 10
       
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Godown Stack Lot Detail")
       
       mHeight = mTop + mRowHeight
       mLeft = mGlobalLeft
   
   
       mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
       mStrLength = Len("Depositor Details")
       mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
       mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
       mFontSize = 7
       'Dim mDispTextSize  As Integer
       mDispTextSize = 6
       
       mTxtTop = mTxtTop + mRowHeight
       
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
       '--
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCxTFDate)
       
       
       
       Call DrawTable(1, 4, mTop, mHeight, mLeft, mRowHeight)
       mHeight = mHeight + 15
   
       mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
       mStrLength = Len("Depositor Details")
       mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
       mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
       mFontSize = 7
   
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Godown No")
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, "Godown Address")
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mmGlobalLeft + 180, mTxtTop, "Stack No")
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 260, mTxtTop, "Lot No")
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 340, mTxtTop, "No Of Bags")
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 420, mTxtTop, "Gross Weight")
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "Client Name")
                           
       mHeight = mTop + mRowHeight
       Call DrawTable(1, 7, mTop, mHeight, mLeft, mRowHeight)
       blnnew = True
       mTxtTopforNextRow = 0
       mUpValuAddtoVertiLine = 75
       mDownValuAddtoVertiLine = mUpValuAddtoVertiLine
   End If
       
Next

mTxtTop = mTxtTop + mRowHeight

'mTxtTop = mTxtTop + mRowHeight

objPDF.PDFEndDoc
    
End Sub




'Private Sub CheckForNewPage()
'
'If mHeight > objPDF.PDFGetPageHeight Then
'   objPDF.PDFWatermark = ""
'   objPDF.PDFEndPage
'   objPDF.PDFNewPage
'   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'   objPDF.PDFSetDrawColor = vbWhite
'   objPDF.PDFSetTextColor = vbBlack
'   objPDF.PDFSetAlignement = ALIGN_CENTER
'   objPDF.PDFSetBorder = BORDER_ALL
'   objPDF.PDFSetFill = True
'
'   objPDF.PDFCell " ", 15, 15, _
'   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'   If PntPrv = "Y" Then
'      objPDF.PDFWatermark = "Preview"
'   Else
'      objPDF.PDFWatermark = "NCMSL"
'   End If
'
'    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'    objPDF.PDFSetDrawColor = vbWhite
'    objPDF.PDFSetTextColor = vbBlack
'    objPDF.PDFSetAlignement = ALIGN_CENTER
'    objPDF.PDFSetBorder = BORDER_ALL
'    objPDF.PDFSetFill = True
'
'   mLeft = mGlobalLeft
'   mTop = 30
'   mHeight = mTop + mRowHeight
'   mTxtFromTop = mRowHeight + 2
'End If
'
'End Sub

Private Sub CheckForNewPage()

If mHeight > objPDF.PDFGetPageHeight - 50 Then
   objPDF.PDFWatermark = ""
   objPDF.PDFEndPage
   objPDF.PDFNewPage
   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
   objPDF.PDFSetDrawColor = vbWhite
   objPDF.PDFSetTextColor = vbBlack
   objPDF.PDFSetAlignement = ALIGN_CENTER
   objPDF.PDFSetBorder = BORDER_ALL
   objPDF.PDFSetFill = True

   objPDF.PDFCell " ", 15, 15, _
   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
   If PntPrv = "Y" Then
      objPDF.PDFWatermark = "Preview"
   Else
      objPDF.PDFWatermark = "NCML" '"NCMSL"
   End If
   
    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
    objPDF.PDFSetDrawColor = vbWhite
    objPDF.PDFSetTextColor = vbBlack
    objPDF.PDFSetAlignement = ALIGN_CENTER
    objPDF.PDFSetBorder = BORDER_ALL
    objPDF.PDFSetFill = True

   mLeft = mGlobalLeft
   mTop = 30
   mHeight = mTop + mRowHeight
   mTxtFromTop = mRowHeight + 2
   mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
   If Totpages > 1 Then
      Call PrintText(6, FONT_NORMAL, ALIGN_RIGHT, objPDF.PDFGetPageWidth - 70, mTxtTop - 5, "Page " & objPDF.PDFPageNumber & " out of " & Totpages)
   Else
   End If
   
End If

End Sub

Private Sub SaveCDTFBOE()
Dim mVar As Variant
Dim mTotDays, mInvAmt, mLotSize, mMinLotSize As Variant

If mExchangeTypeID <> 2 Then Exit Sub

tmp = "Select Collect7Days from Mst_Location where LocationID=" & mLocationID
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If TmpRs!Collect7Days <> True Then
      Exit Sub
   End If
End If

''Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CDTFNo=" & Val(TxtCxTFNo))

Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='D'")


If RsTxn_CDTFBOE.RecordCount > 0 Then Exit Sub

Call TableTxn_CDTFBOE_GSL(" Where SM_Prefix='" & TxtPreFix & "'")

mVar = Val(TxtCVariation)
mLotSize = Val(TxtCLotSize)

If mVar > 0 Then
   mMinLotSize = mLotSize - Format(mLotSize * (mVar / 100), "#####0.000")
End If


tmp = "Select GodownID,Rate from Mst_NCDEXStorageRateCardDetail MNRC"
tmp = tmp & " inner Join Mst_NCDEXStorageRateCard MNR on MNRC.NSRCID=MNR.NSRCID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MNR.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_Commodity MComm On MCG.CommodityGroupID = MComm.CommodityGroupID"
tmp = tmp & " Where MComm.CommodityID=" & mCommodityID & "  And MNR.LocationID=" & mLocationID

Call Tmp2Table

tmp = "select god.GodownNo 'GodownNo',gsl.StackNo,gsl.LotNo,"
tmp = tmp & " sum(tgsl.NoofBags)NoofBags,sum(tgsl.GrossWeight)GrossWeight,sum(tgsl.AdjNClientWeight)AdjNClientWeight,"
tmp = tmp & " sum(tgsl.AdjNCMSLWeight)AdjNCMSLWeight,sum(tgsl.GunnyWeight)GunnyWeight,sum(tgsl.StdDeduction)StdDeduction,"
tmp = tmp & " sum(tgsl.DematWeight)DematWeight,c.ClientName + ' ' + d.DepositoryName ClientName,tgsl.Remarks,tgsl.Mst_GSL_ID,"
tmp = tmp & " god.CloseDate , god.GodownID, tgsl.ClientID, gsl.ISINID, gsl.FGSLID_Mst_GSL_ID, god.Address"
tmp = tmp & " From"
tmp = tmp & " Txn_CxTF_GSL tgsl"
tmp = tmp & " inner join mst_gsl gsl on gsl.Mst_GSL_ID = tgsl.Mst_GSL_ID and gsl.SM_Prefix = tgsl.SM_Prefix"
tmp = tmp & " inner join Mst_Godown god on gsl.GodownID = god.GodownID"
tmp = tmp & " inner join mst_client c on tgsl.ClientID = c.ClientID"
tmp = tmp & " inner join mst_depository d on d.DepositoryID = c.DepositoryID  and d.ExchangeTypeID =" & mExchangeTypeID
tmp = tmp & " Where tgsl.TxnType = 'D'  And  tgsl.CxTF_DetailsID =  " & TxtCxTFNo & "  And tgsl.SM_Prefix = '" & TxtPreFix & "'"
tmp = tmp & " Group By god.GodownNo,gsl.StackNo,gsl.LotNo,c.ClientName ,d.DepositoryName,"
tmp = tmp & " tgsl.Remarks,tgsl.Mst_GSL_ID,god.CloseDate,god.GodownID,tgsl.ClientID,gsl.ISINID,gsl.FGSLID_Mst_GSL_ID,"
tmp = tmp & " god.Address"

Call Tmp3Table

For I = 1 To TmpRs3.RecordCount

    If TmpRs3!Mst_GSL_ID <> "" Then
       tmp = "Select TCD.CxTFDate,TCD.CxTF_DetailsID From Txn_CxTF_Details TCD"
       tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTF_DetailsID = TCG.CxTF_DetailsID And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType "
       tmp = tmp & " Where TCD.TxnType = 'D' and TCD.ExchangeTypeID = " & mExchangeTypeID & " And TCG.Mst_GSL_ID=" & TmpRs3!Mst_GSL_ID & " And TCD.SM_Prefix='" & TxtPreFix & "' And TCD.CxTF_DetailsID>" & Val(TxtCxTFNo)
       Call Tmp5Table
       
       If TmpRs5.RecordCount <= 0 Then
            If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
            TmpRs2.Find "GodownID=" & TmpRs3!GodownID
            If TmpRs2.EOF = False Then
               mRate = TmpRs2!Rate
            Else
               mRate = 0
            End If
            If Val(TmpRs3!DematWeight) >= mMinLotSize Then
               If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
               RsTxn_CDTFBOE.Find "ClientID=" & TmpRs3!ClientID
               If RsTxn_CDTFBOE.EOF = True Then
                  Call AddCDTFBOE(Val(TmpRs3!ClientID))
               End If
               Call SaveCDTFBOEGSL(Val(TmpRs3!Mst_GSL_ID), TmpRs3!NoofBags, TmpRs3!DematWeight, 7, mRate)
            Else
               tmp = "Select DepositeWeight,DepositeBags,CDTFDate,LCDTF_Date From Mst_GSL Where CDTFDate<>LCDTF_Date And Mst_GSL_ID=" & Val(TmpRs3!Mst_GSL_ID) & " And SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID<" & Val(TxtCxTFNo)
               Call Tmp4Table
               If TmpRs4.RecordCount > 0 Then
                  If Val(TmpRs4!DepositeWeight) >= mMinLotSize Then
                     If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
                     RsTxn_CDTFBOE.Find "ClientID=" & Val(TmpRs3!ClientID)
                     If RsTxn_CDTFBOE.EOF = True Then
                        Call AddCDTFBOE(Val(TmpRs3!ClientID))
                     End If
                     mTotDays = DateDiff("d", TmpRs4!CDTFDate, DateAdd("d", 7, TmpRs4!LCDTF_Date)) '+ 1
                     Call SaveCDTFBOEGSL(Val(TmpRs3!Mst_GSL_ID), TmpRs4!DepositeBags, TmpRs4!DepositeWeight, mTotDays, mRate)
                  End If
               End If
            End If
            tmp = " Update  Txn_CDTFBOE Set Amount = " & Val(mInvAmountOnPrint) & " Where BOEID =  " & Val(mBOEIDNew) & "   "
            Call Tmp7Table
       End If
    End If
    TmpRs3.MoveNext
Next


''For I = 2 To MSHFlexGrid4.Rows - 1
''
''    If MSHFlexGrid4.TextMatrix(I, 13) <> "" Then
''       tmp = "Select TCD.CxTFDate,TCD.CxTFNo From Txn_CxTF_Details TCD"
''       tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType "
''       tmp = tmp & " Where TCD.TxnType = 'D' and TCD.ExchangeTypeID = " & mExchangeTypeID & " And TCG.GSLID=" & Val(MSHFlexGrid4.TextMatrix(I, 13)) & " And TCD.SM_Prefix='" & TxtPreFix & "' And TCD.CxTFNo>" & Val(TxtCxTFNo)
''       Call Tmp5Table
''       If TmpRs5.RecordCount <= 0 Then
''
''            If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
''            TmpRs2.Find "GodownID=" & Val(MSHFlexGrid4.TextMatrix(I, 15))
''            If TmpRs2.EOF = False Then
''               mRate = TmpRs2!Rate
''            Else
''               mRate = 0
''            End If
''            If Val(MSHFlexGrid4.TextMatrix(I, 10)) >= mMinLotSize Then
''               If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
''               RsTxn_CDTFBOE.Find "ClientID=" & Val(MSHFlexGrid4.TextMatrix(I, 16))
''               If RsTxn_CDTFBOE.EOF = True Then
''                  Call AddCDTFBOE(Val(MSHFlexGrid4.TextMatrix(I, 16)))
''               End If
''               Call SaveCDTFBOEGSL(Val(MSHFlexGrid4.TextMatrix(I, 13)), MSHFlexGrid4.TextMatrix(I, 4), MSHFlexGrid4.TextMatrix(I, 10), 7, mRate)
''            Else
''               tmp = "Select DepositeWeight,DepositeBags,CDTFDate,LCDTF_Date From Mst_GSL Where CDTFDate<>LCDTF_Date And GSLID=" & Val(MSHFlexGrid4.TextMatrix(I, 13)) & " And SM_Prefix='" & TxtPreFix & "' And CDTFNo<" & Val(TxtCxTFNo)
''               Call Tmp3Table
''               If TmpRs3.RecordCount > 0 Then
''                  If Val(TmpRs3!DepositeWeight) >= mMinLotSize Then
''                     If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
''                     RsTxn_CDTFBOE.Find "ClientID=" & Val(MSHFlexGrid4.TextMatrix(I, 16))
''                     If RsTxn_CDTFBOE.EOF = True Then
''                        Call AddCDTFBOE(Val(MSHFlexGrid4.TextMatrix(I, 16)))
''                     End If
''                     mTotDays = DateDiff("d", TmpRs3!CDTFDate, DateAdd("d", 7, TmpRs3!LCDTF_Date)) '+ 1
''                     Call SaveCDTFBOEGSL(Val(MSHFlexGrid4.TextMatrix(I, 13)), TmpRs3!DepositeBags, TmpRs3!DepositeWeight, mTotDays, mRate)
''                  End If
''               End If
''            End If
''       End If
''    End If
''Next
''

End Sub


'Private Sub SaveCDTFBOE()
'Dim mVar As Variant
'Dim mTotDays, mInvAmt, mLotSize, mMinLotSize As Variant
'
'If mExchangeTypeID <> 2 Then Exit Sub
'
'tmp = "Select Collect7Days from Mst_Location where LocationID=" & mLocationID
'Call TmpTable
'If TmpRs.RecordCount > 0 Then
'   If TmpRs!Collect7Days <> 1 Then
'      Exit Sub
'   End If
'End If
'
'Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CDTFNo=" & Val(TxtCxTFNo))
'
'If RsTxn_CDTFBOE.RecordCount > 0 Then Exit Sub
'
'Call TableTxn_CDTFBOE_GSL(" Where SM_Prefix='" & TxtPreFix & "'")
'
'mVar = Val(TxtCVariation)
'mLotSize = Val(TxtCLotSize)
'
'If mVar > 0 Then
'   mMinLotSize = mLotSize - Format(mLotSize * (mVar / 100), "#####0.000")
'End If
'
'
'tmp = "Select GodownID,Rate from Mst_NCDEXStorageRateCardDetail MNRC"
'tmp = tmp & " inner Join Mst_NCDEXStorageRateCard MNR on MNRC.NSRCID=MNR.NSRCID"
'tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MNR.CommodityGroupID = MCG.CommodityGroupID"
'tmp = tmp & " Inner Join Mst_Commodity MComm On MCG.CommodityGroupID = MComm.CommodityGroupID"
'tmp = tmp & " Where MComm.CommodityID=" & mCommodityID & "  And MNR.LocationID=" & mLocationID
'
'Call Tmp2Table
'
''If TmpRs2.RecordCount > 0 Then
''   TmpRs2.MoveFirst
''   For I = 1 To TmpRs2.RecordCount
''       If TmpRs2!Rate = 0 Then
''          MsgBox "No Rate updated in rate card master!!!" 'for Location " & GetLocationName(mLocationID) & " And Commodity Group " & GetCommodityGroupName(  & " "
''          'TxtHologramNo.s
''          Exit Sub
''       End If
''   Next
''Else
''   MsgBox "No Rate updated in rate card master!!!" 'for Location " & GetLocationName(mLocationID) & " And Commodity Group " & GetCommodityGroupName(  & " "
''   Exit Sub
''End If
'
'For I = 2 To MSHFlexGrid4.Rows - 1
'    If MSHFlexGrid4.TextMatrix(I, 13) <> "" Then
'       tmp = "Select TCD.CxTFDate,TCD.CxTFNo From Txn_CxTF_Details TCD"
'       tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType "
'       tmp = tmp & " Where TCD.TxnType = 'D' and TCD.ExchangeTypeID = " & mExchangeTypeID & " And TCG.GSLID=" & Val(MSHFlexGrid4.TextMatrix(I, 13)) & " And TCD.SM_Prefix='" & TxtPreFix & "' And TCD.CxTFNo>" & Val(TxtCxTFNo)
'       Call Tmp5Table
'       If TmpRs5.RecordCount <= 0 Then
'
'            If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
'            TmpRs2.Find "GodownID=" & Val(MSHFlexGrid4.TextMatrix(I, 15))
'            If TmpRs2.EOF = False Then
'               mRate = TmpRs2!Rate
'            Else
'               mRate = 0
'            End If
'            If Val(MSHFlexGrid4.TextMatrix(I, 10)) >= mMinLotSize Then
'               If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
'               RsTxn_CDTFBOE.Find "ClientIDRow=" & Val(MSHFlexGrid4.TextMatrix(I, 16))
'               If RsTxn_CDTFBOE.EOF = True Then
'                  Call AddCDTFBOE(Val(MSHFlexGrid4.TextMatrix(I, 16)))
'               End If
'               Call SaveCDTFBOEGSL(Val(MSHFlexGrid4.TextMatrix(I, 13)), MSHFlexGrid4.TextMatrix(I, 4), MSHFlexGrid4.TextMatrix(I, 10), 7, mRate)
'            Else
'               tmp = "Select DepositeWeight,DepositeBags,CDTFDate,LCDTF_Date From Mst_GSL Where CDTFDate<>LCDTF_Date And GSLID=" & Val(MSHFlexGrid4.TextMatrix(I, 13)) & " And SM_Prefix='" & TxtPreFix & "' And CDTFNo<" & Val(TxtCxTFNo)
'               Call Tmp3Table
'               If TmpRs3.RecordCount > 0 Then
'                  If Val(TmpRs3!DepositeWeight) >= mMinLotSize Then
'                     If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
'                     RsTxn_CDTFBOE.Find "ClientIDRow=" & Val(MSHFlexGrid4.TextMatrix(I, 16))
'                     If RsTxn_CDTFBOE.EOF = True Then
'                        Call AddCDTFBOE(Val(MSHFlexGrid4.TextMatrix(I, 16)))
'                     End If
'                     mTotDays = DateDiff("d", TmpRs3!CDTFDate, DateAdd("d", 7, TmpRs3!LCDTF_Date)) '+ 1
'                     Call SaveCDTFBOEGSL(Val(MSHFlexGrid4.TextMatrix(I, 13)), TmpRs3!DepositeBags, TmpRs3!DepositeWeight, mTotDays, mRate)
'                  End If
'               End If
'            End If
'       End If
'    End If
'Next
'
'
'End Sub
Private Sub AddCDTFBOE(mClientID_ As Variant)

RsTxn_CDTFBOE.AddNew
'RsTxn_CDTFBOE!BOEID = getCDTFBOE_Max
'mBOEIDNew = RsTxn_CDTFBOE!BOEID
RsTxn_CDTFBOE!LocationID = mLocationID
RsTxn_CDTFBOE!SM_Prefix = TxtPreFix
RsTxn_CDTFBOE!CxTF_DetailsID = TxtCxTFNo
RsTxn_CDTFBOE!CDTFDate = TxtCxTFDate.Value
RsTxn_CDTFBOE!ClientID = mClientID_
RsTxn_CDTFBOE!InvoiceDate = Date
RsTxn_CDTFBOE!BillingMonth = CDate(Year(Date) & "-" & Month(Date) & "-01")
RsTxn_CDTFBOE!Discount = 0
RsTxn_CDTFBOE!Premium = 0
RsTxn_CDTFBOE!BatchID = 0
RsTxn_CDTFBOE!RBatchID = 0
RsTxn_CDTFBOE!Flag = "P"
RsTxn_CDTFBOE!G_UID = GetGUID
mInvG_UID = RsTxn_CDTFBOE!G_UID
RsTxn_CDTFBOE!CreatedDate = Now
RsTxn_CDTFBOE!CreatedBy = Gen_UserLoginID
RsTxn_CDTFBOE!TxnType = "D"
RsTxn_CDTFBOE.Update


tmp = "Select BOEID from Txn_CDTFBOE Where G_UID = '" & mInvG_UID & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mBOEIDNew = TmpRs!BOEID
End If


End Sub
Private Sub SaveCDTFBOEGSL(mGSLID_ As Variant, mNoofBags_ As Variant, mDematWeight_ As Variant, mNoOfDays_ As Variant, mRate_ As Variant)

RsTxn_CDTFBOE_GSL.AddNew
'RsTxn_CDTFBOE_GSL!BOEGSLID = getCDTFBOE_GSL_Max
RsTxn_CDTFBOE_GSL!BOEID = mBOEIDNew 'RsTxn_CDTFBOE!BOEID
RsTxn_CDTFBOE_GSL!SM_Prefix = TxtPreFix
RsTxn_CDTFBOE_GSL!Mst_GSL_ID = mGSLID_
RsTxn_CDTFBOE_GSL!CDTFBags = mNoofBags_
RsTxn_CDTFBOE_GSL!CDTFQty = mDematWeight_
RsTxn_CDTFBOE_GSL!NoOfDays = mNoOfDays_
RsTxn_CDTFBOE_GSL!Rate = mRate_ 'TmpRs2!Rate
RsTxn_CDTFBOE_GSL!BillingUnit = "Q" 'TmpRs2!Unit
RsTxn_CDTFBOE_GSL!BillingCycleID = 1 ''TmpRs2!BillingCycleID
RsTxn_CDTFBOE_GSL!Amount = Format(mNoOfDays_ * mDematWeight_ * mRate_, "#####0.00")
RsTxn_CDTFBOE_GSL.Update

'RsTxn_CDTFBOE!Amount = IIf(IsNull(RsTxn_CDTFBOE!Amount), 0, RsTxn_CDTFBOE!Amount) + RsTxn_CDTFBOE_GSL!Amount
'RsTxn_CDTFBOE.Update

mInvAmountOnPrint = Val(mInvAmountOnPrint) + RsTxn_CDTFBOE_GSL!Amount

tmp = "Update Txn_CxTF_GSL Set BillFlag='I' Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & TxtCxTFNo & " And TxnType='D' And Mst_GSL_ID=" & mGSLID_ & " And BillFlag is null"
Call Tmp5Table

End Sub
Function getCDTFBOE_Max() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(BOEID) from Txn_CDTFBOE"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCDTFBOE_Max = Retval
End Function

Function getCDTFBOE_GSL_Max() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(BOEGSLID) from Txn_CDTFBOE_GSL"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCDTFBOE_GSL_Max = Retval
End Function



Function CheckDuplicateTruck() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid3.Rows > 3 Then
   For u = 1 To MSHFlexGrid3.Rows - 1
       If LCase(MSHFlexGrid3.TextMatrix(u, 1)) = LCase(TxtTTDTruckNo.Text) And LCase(MSHFlexGrid3.TextMatrix(u, 2)) = LCase(TxtWeighBridgNo.Text) Then
          If LCase(CmdTTDAdd2List.Caption) = "add to list" Then
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

CheckDuplicateTruck = Retval
End Function


Private Sub DeleteFromGSLDetail()
I = MSHFlexGrid3.RowSel
Dim mDelRows As Variant
If I > 0 Then
   If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
      mDelRows = 0
Restart:
        For r = 2 To MSHFlexGrid4.Rows - 1
            If LCase(MSHFlexGrid4.TextMatrix(r, 19)) = LCase(MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2)) Then
               MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(r, 4))
               MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(r, 5))
               MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) - Val(MSHFlexGrid4.TextMatrix(r, 6))
               MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) - Val(MSHFlexGrid4.TextMatrix(r, 7))
               MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) - Val(MSHFlexGrid4.TextMatrix(r, 8))
               MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) - Val(MSHFlexGrid4.TextMatrix(r, 9))
               MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) - Val(MSHFlexGrid4.TextMatrix(r, 10))
               

               For RO = r To MSHFlexGrid4.Rows - 2
                   For C = 0 To MSHFlexGrid4.Cols - 1
                       MSHFlexGrid4.TextMatrix(RO, C) = MSHFlexGrid4.TextMatrix(RO + 1, C)
                   Next
               Next
               MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 1
               GoTo Restart
            End If
        Next
         

   End If
End If

End Sub


Private Sub CmbGSLTruckNo_GotFocus()
tmp = CmbGSLTruckNo.Text
CmbGSLTruckNo.Clear
If MSHFlexGrid3.Rows < 4 Then Exit Sub

For I = 2 To MSHFlexGrid3.Rows - 2
    CmbGSLTruckNo.AddItem MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2)
Next I
CmbGSLTruckNo.Text = tmp
End Sub

Private Sub CmbGSLTruckNo_LostFocus()
If CmbGSLTruckNo.Text = "" Then Exit Sub
For I = 2 To MSHFlexGrid3.Rows - 2
    If CmbGSLTruckNo.Text = MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2) Then
       'mTruckNo = MSHFlexGrid3.TextMatrix(I, 11)
       mTruckWtRatio = Val(MSHFlexGrid3.TextMatrix(I, 5)) / Val(MSHFlexGrid3.TextMatrix(I, 6))
       TxtGCTDNoofBags = getRemainingTruckBag(CmbGSLTruckNo.Text)
       TxtGCTDGrossWt = Format(getRemainingTruckWt(CmbGSLTruckNo.Text), "#######0.000")
       Call CalDmatWt
       'Add BY Prajakta to get moisture %
       For j = 2 To MSHFlexGrid3.Rows - 1
           If CmbGSLTruckNo.Text = MSHFlexGrid3.TextMatrix(j, 1) & "~" & MSHFlexGrid3.TextMatrix(j, 2) Then
              TxtMoisturePer1.Text = MSHFlexGrid3.TextMatrix(j, 8)
              Exit For
           End If
       Next
       Exit Sub
    End If
Next I

MsgBox "Invalid selection!!!"
CmbGSLTruckNo.SetFocus

End Sub

Private Function getRemainingTruckWt(mTruckNo_ As String) As Double
Dim mTotTruckwt, mTotUseWt As Double
Dim mstr() As String
mTotTruckwt = 0
mTotUseWt = 0
mstr = Split(mTruckNo_, "~")
For I = 2 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(I, 1) = mstr(0) And MSHFlexGrid3.TextMatrix(I, 2) = mstr(1) Then
       mTotTruckwt = mTotTruckwt + Val(MSHFlexGrid3.TextMatrix(I, 5))
    End If
Next I

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 19) = mTruckNo_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotUseWt = mTotUseWt + Val(MSHFlexGrid4.TextMatrix(I, 5))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotUseWt = mTotUseWt + Val(MSHFlexGrid4.TextMatrix(I, 5))
          End If
       End If
    End If
Next I
getRemainingTruckWt = mTotTruckwt - mTotUseWt
End Function

Private Function getRemainingTruckBag(mTruckNo_ As String) As Double
Dim mTotTruckBag, mTotUseBag As Double
Dim mstr() As String
mTotTruckBag = 0
mTotUseBag = 0
mstr = Split(mTruckNo_, "~")
For I = 2 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(I, 1) = mstr(0) And MSHFlexGrid3.TextMatrix(I, 2) = mstr(1) Then
       mTotTruckBag = mTotTruckBag + Val(MSHFlexGrid3.TextMatrix(I, 6))
    End If
Next I

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 19) = mTruckNo_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotUseBag = mTotUseBag + Val(MSHFlexGrid4.TextMatrix(I, 4))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotUseBag = mTotUseBag + Val(MSHFlexGrid4.TextMatrix(I, 4))
          End If
       End If
    End If
Next I
getRemainingTruckBag = mTotTruckBag - mTotUseBag
End Function

Private Function getRemainingGSLBag(mGSLID_ As Variant) As Double
Dim mTotWidBag As Double
mTotWidBag = 0

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 14) = mGSLID_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotWidBag = mTotWidBag + Val(MSHFlexGrid4.TextMatrix(I, 6))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotWidBag = mTotWidBag + Val(MSHFlexGrid4.TextMatrix(I, 6))
          End If
       End If
    End If
Next I

getRemainingGSLBag = mTotWidBag
End Function
Private Function getRemainingGSLWt(mGSLID_ As Variant) As Double
Dim mTotWidwt As Double
mTotWidwt = 0

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 14) = mGSLID_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotWidwt = mTotWidwt + Val(MSHFlexGrid4.TextMatrix(I, 7))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotWidwt = mTotWidwt + Val(MSHFlexGrid4.TextMatrix(I, 7))
          End If
       End If
    End If
Next I

getRemainingGSLWt = mTotWidwt
End Function


Function CheckPreviousGSLClient() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid4.Rows > 3 Then
   For u = 1 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(u, 1) = CmbGodownNo.Text And _
          MSHFlexGrid4.TextMatrix(u, 2) = CmbStack.Text And _
          MSHFlexGrid4.TextMatrix(u, 3) = CmbLot.Text Then
          If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
             If LCase(CmbGCTDClient.Text) <> LCase(MSHFlexGrid4.TextMatrix(u, 11)) Then
                Retval = True
                Exit For
             End If
          Else
             If u <> MSHFlexGrid4.RowSel Then
                If LCase(CmbGCTDClient.Text) <> LCase(MSHFlexGrid4.TextMatrix(u, 11)) Then
                   Retval = True
                   Exit For
                End If
             End If
          End If
       End If
   Next
End If

CheckPreviousGSLClient = Retval
End Function
Private Sub TxtTTDTruckNo_LostFocus()
TxtTTDTruckNo.Text = Trim$(TxtTTDTruckNo.Text)
End Sub
Private Sub TxtWeighBridgNo_LostFocus()
TxtWeighBridgNo.Text = Trim$(TxtWeighBridgNo.Text)
End Sub

Private Function getDepositedGSLBag(mGSLID_ As Variant) As Double
Dim mTotWidBag As Double
mTotWidBag = 0

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 13) = mGSLID_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotWidBag = mTotWidBag + Val(MSHFlexGrid4.TextMatrix(I, 4))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotWidBag = mTotWidBag + Val(MSHFlexGrid4.TextMatrix(I, 4))
          End If
       End If
    End If
Next I

getDepositedGSLBag = mTotWidBag
End Function
Private Function getDepositedGSLWt(mGSLID_ As Variant) As Double
Dim mTotWidwt As Double
mTotWidwt = 0

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 13) = mGSLID_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotWidwt = mTotWidwt + Val(MSHFlexGrid4.TextMatrix(I, 5))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotWidwt = mTotWidwt + Val(MSHFlexGrid4.TextMatrix(I, 5))
          End If
       End If
    End If
Next I

getDepositedGSLWt = mTotWidwt
End Function

Private Function Check7DaysInvoice(gslID_ As Variant) As Variant
Dim Retval As Variant
Retval = ""

tmp = "Select TC.BOEID,TCG.Mst_GSL_ID  from  Txn_CDTFBOE_GSL  TCG"
tmp = tmp & " Inner join Txn_CDTFBOE TC on TCG.BOEID=TC.BOEID"
tmp = tmp & " Where TCG.Mst_GSL_ID in (" & gslID_ & ") And TCG.SM_Prefix='" & TxtPreFix & "' And TC.TxnType='D'"

Call Tmp6Table
 
If TmpRs6.RecordCount > 0 Then
   Retval = TmpRs6!BOEID
End If
Check7DaysInvoice = Retval
End Function


Private Function GetTotalPDFPages() As Double
Set objPDF = New mjwPDF

' Create a simple PDF file using the mjwPDF class

Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
    
    
' Set the PDF title and filename
objPDF.PDFTitle = "CDTF PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
' We must tell the class where the PDF fonts are located
objPDF.PDFLoadAfm = App.Path & "\Fonts"

' Set the file properties
objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
objPDF.PDFFormatPage = FORMAT_A4
objPDF.PDFOrientation = ORIENT_PORTRAIT
objPDF.PDFSetUnit = UNIT_PT
objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU

' Lets us set see the bookmark pane when we view the PDF
'objPDF.PDFUseOutlines = True

' View the PDF file after we create it
objPDF.PDFView = False

' Begin our PDF document

objPDF.PDFBeginDoc
    ' Lets add a heading
objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

objPDF.PDFCell " ", 15, 15, _
objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20


If PntPrv = "Y" Then
   objPDF.PDFWatermark = "Preview"
Else
   objPDF.PDFWatermark = "NCML" '"NCMSL"
End If

objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

pgval = ((objPDF.PDFGetPageWidth - 18))

'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)

mRowHeight = 15
mGlobalLeft = 12
'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
mLeft = mGlobalLeft
mTop = 15
mHeight = mTop + mRowHeight
mTxtFromTop = mRowHeight + 2
 
 
 
 'Call DrawTable(10, 2, 15, 30, 12)

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("NCMSL " & TxtCMPID & "")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 12

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")

Call PrintText(6, FONT_NORMAL, ALIGN_RIGHT, objPDF.PDFGetPageWidth - 70, mTxtTop - 5, "Page No : " & objPDF.PDFPageNumber)

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft


Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Deposit Transaction Form (CDTF)")

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
 
tmp1 = mTop
mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("(This is NOT an instrument of title of the goods)")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 6

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")

'--Code added By Kishor
'Row 1

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Depositor Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7
Dim mDispTextSize  As Integer
mDispTextSize = 6

mTxtTop = mTxtTop + mRowHeight

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Date")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCMPID))
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "State")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCMPID)))
'--
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
Dim mCMPIDnAddr As String
mCMPIDnAddr = mCMPID & "-" & GetCMPAddress(mCMPID, "A")
Dim mLen As Integer, mMax As Integer, mPara As Integer
Dim t1 As Double
mMax = 45
mPara = 0
t1 = 1
mLen = Len(mCMPIDnAddr)
mLen = Round(mLen / mMax) '+ 1

For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       'mMax '= mMax + 35
       mPara = mPara + 7
    End If
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, mID$(mCMPIDnAddr, t1, mMax))
Next

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Commodity")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, TxtCommodityID)

mTxtTop = mTxtTop + mPara + mRowHeight

objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

objPDF.PDFDrawLine mGlobalLeft + 345, Val(tmp1), mGlobalLeft + 345, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 445, Val(tmp1), mGlobalLeft + 445, mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtExchangeTypeID)

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)

'objPDF.PDFDrawLine mGlobalLeft + 132, mTop - 4 * mRowHeight, mGlobalLeft + 132, mHeight + mRowHeight

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 135, Val(tmp1), mGlobalLeft + 135, mTxtTop - 10

''Call DrawTable(2, , mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD & FONT_UNDERLINE, ALIGN_LEFT, mLeft, mTxtTop - 2, "Depositor Details")

''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
mTxtTop = mTxtTop + mRowHeight

objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

''objPDF.PDFDrawLine mGlobalLeft + 132, mTop, mGlobalLeft + 132, mTop + mRowHeight
''objPDF.PDFDrawLine 3 * (mGlobalLeft + 132), mTop, 3 * (mGlobalLeft + 132), mTop + mRowHeight
''objPDF.PDFDrawLine 3.5 * (mGlobalLeft + 132), mTop, 3.5 * (mGlobalLeft + 132), mTop + mRowHeight
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("Depositor Details")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
''mFontSize = 7

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Client ID")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, "Client Name")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, "No of Bags")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "Gross Weight")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop + 10, "in MT")

mTxtTop = mTxtTop + mRowHeight + 10
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 135, mTxtTop - (mRowHeight * 2 + 5), mGlobalLeft + 135, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 425, mTxtTop - (mRowHeight * 2 + 5), mGlobalLeft + 425, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 495, mTxtTop - (mRowHeight * 2 + 5), mGlobalLeft + 495, mTxtTop - 10


Dim mTemp As Integer

'FILL GRID WITH DETAILS
For T = 2 To MSHFlexGrid2.Rows - 1
    If mHeight > objPDF.PDFGetPageHeight - 50 Then
       Call CheckForNewPage
       mTop = 15
       mHeight = mHeight - mRowHeight
       mTxtTop = 10
    End If
    'mTxtTop = mTxtTop + mRowHeight
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid2.TextMatrix(T, 1))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, MSHFlexGrid2.TextMatrix(T, 2))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, MSHFlexGrid2.TextMatrix(T, 5))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, MSHFlexGrid2.TextMatrix(T, 6))

    mTxtTop = mTxtTop + mRowHeight
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
    objPDF.PDFDrawLine mGlobalLeft + 135, mTxtTop - (mRowHeight + 10), mGlobalLeft + 135, mTxtTop - 10
    objPDF.PDFDrawLine mGlobalLeft + 425, mTxtTop - (mRowHeight + 10), mGlobalLeft + 425, mTxtTop - 10
    objPDF.PDFDrawLine mGlobalLeft + 495, mTxtTop - (mRowHeight + 10), mGlobalLeft + 495, mTxtTop - 10
    mTop = mTop + mRowHeight
    mHeight = mHeight + mRowHeight
Next

Call CheckForNewPage
mTxtTop = mTxtTop - mRowHeight
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 220, mTxtTop, "Total")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, MSHFlexGrid2.TextMatrix(0, 5))
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, MSHFlexGrid2.TextMatrix(0, 6))

Call CheckForNewPage
mTxtTop = mTxtTop + mRowHeight

Call PrintText(mFontSize, FONT_BOLD & FONT_UNDERLINE, ALIGN_LEFT, mLeft, mTxtTop, "Truck-wise Deposit Details")
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Truck No")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Weighbridge Slip No")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Totol Weight")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop + 10, "in MT")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, "Tare Weight")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop + 10, "in MT")
''-- Ullhas Patil 21 APril 2018
If mSiloLocation = True Then
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Foreign")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "Matter")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Net Weight")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop + 10, "in MT")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "No Of Bags")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 410, mTxtTop, "Quality Parametes / Remark")
Else
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
End If
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
''-----------------------------
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")

mTxtTop = mTxtTop + mRowHeight + 10

objPDF.PDFDrawLine mGlobalLeft + 65, mTxtTop - (mRowHeight + 20), mGlobalLeft + 65, (mTxtTop - 10)
objPDF.PDFDrawLine (mGlobalLeft + 145), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 145), (mTxtTop - 10)
objPDF.PDFDrawLine (mGlobalLeft + 195), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 195), (mTxtTop - 10)
objPDF.PDFDrawLine (mGlobalLeft + 245), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 245), (mTxtTop - 10)
''-- Ullhas Patil 21 APril 2018
If mSiloLocation = True Then
    objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 405), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 405), (mTxtTop - 10)
Else
    objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
End If
''objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
''objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
''-----------------------------
''objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
''objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

mHeight = mTxtTop ' mTop + mRowHeight
'mTxtTop = mHeight
For T = 2 To MSHFlexGrid3.Rows - 1

    If mHeight > objPDF.PDFGetPageHeight - 50 Then
       Call CheckForNewPage
       mTop = 15
       
      mRowHeight = 15
      mGlobalLeft = 12
      'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
      mLeft = mGlobalLeft
      mTop = 15
      mHeight = mTop + mRowHeight
      mTxtFromTop = mRowHeight + 2
      
      'Call DrawTable(10, 2, 15, 30, 12)
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("NCMSL " & TxtCMPID & "")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
      mFontSize = 12
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
      
      mHeight = mTop + mRowHeight
      mLeft = mGlobalLeft
      
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
      mFontSize = 10
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Deposit Transaction Form (CDTF)")
      
      mHeight = mTop + mRowHeight
      mLeft = mGlobalLeft
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      tmp1 = mTop
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("(This is NOT an instrument of title of the goods)")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
      mFontSize = 6
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
      
      '--Code added By Kishor
      'Row 1
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("Depositor Details")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
      mFontSize = 7
      'Dim mDispTextSize  As Integer
      mDispTextSize = 6
      
      mTxtTop = mTxtTop + mRowHeight
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CDTF No")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
      '--
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Date")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))
      
      mTxtTop = mTxtTop + mRowHeight
      
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCMPID))
      '--
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "State")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCMPID)))
      '--
      mTxtTop = mTxtTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

     
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtExchangeTypeID)
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Commodity")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, TxtCommodityID)
      
      mTxtTop = mTxtTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

      objPDF.PDFDrawLine mGlobalLeft + 345, Val(tmp1), mGlobalLeft + 345, mTxtTop - 10
      objPDF.PDFDrawLine mGlobalLeft + 445, Val(tmp1), mGlobalLeft + 445, mTxtTop - 10
      objPDF.PDFDrawLine mGlobalLeft + 135, Val(tmp1), mGlobalLeft + 135, mTxtTop - 10
      
      mTxtTop = mTxtTop + (mRowHeight * 2)
      
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Truck No")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Weighbridge Slip No")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Totol Weight")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop + 10, "in MT")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, "Tare Weight")
        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop + 10, "in MT")
        ''-- Ullhas Patil 21 APril 2018
        If mSiloLocation = True Then
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Foreign")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "Matter")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Net Weight")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop + 10, "in MT")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "No Of Bags")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 410, mTxtTop, "Quality Parametes / Remark")
        Else
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
            Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
        End If
        ''        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
        ''        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
        ''        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
        ''        Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
        ''-----------------------------
        ''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, "Gross Weight")
        ''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop + 10, "in MT")
        ''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "No Of Bags")
        ''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, "Quality Parametes / Remark")
      
      mTxtTop = mTxtTop + mRowHeight + 10
      mHeight = mTxtTop
      objPDF.PDFDrawLine mGlobalLeft + 65, mTxtTop - (mRowHeight + 20), mGlobalLeft + 65, (mTxtTop - 10)
      objPDF.PDFDrawLine (mGlobalLeft + 145), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 145), (mTxtTop - 10)
      objPDF.PDFDrawLine (mGlobalLeft + 195), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 195), (mTxtTop - 10)
      objPDF.PDFDrawLine (mGlobalLeft + 245), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 245), (mTxtTop - 10)
      ''-- Ullhas Patil 21 APril 2018
      If mSiloLocation = True Then
        objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 405), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 405), (mTxtTop - 10)
      Else
        objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
      End If
      ''      objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
      ''      objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
      ''-----------------------------
      ''objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 300), (mTxtTop - 10)
      ''objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (mRowHeight + 20), (mGlobalLeft + 355), (mTxtTop - 10)
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

    End If
   
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(T, 1))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, MSHFlexGrid3.TextMatrix(T, 2))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, MSHFlexGrid3.TextMatrix(T, 3))
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(T, 4))
    ''-- Ullhas Patil 21 APril 2018
    If mSiloLocation = True Then
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(T, 13))
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(T, 5))
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, MSHFlexGrid3.TextMatrix(T, 6))
    Else
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(T, 5))
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(T, 6))
    End If
    ''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(T, 5))
    ''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(T, 6))
    ''------------------------------
    ''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(T, 5))
    ''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(T, 6))
    'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, MSHFlexGrid3.TextMatrix(t, 7))
    mMax = 62
    mPara = 7
    t1 = 1
    mLen = Len(MSHFlexGrid3.TextMatrix(T, 7))
    
    If mLen > 70 Then
       mLen = RoundUp(mLen / mMax)
    Else
       mLen = 1
    End If
    For S = 1 To mLen
        If S > 1 Then
           t1 = mMax + t1
           mPara = mPara + 7
        End If
        ''-- Ullhas Patil 21 APril 2018
        If mSiloLocation = True Then
            Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 410, mTxtTop + mPara - 5, mID$(MSHFlexGrid3.TextMatrix(T, 7), t1, mMax))
        Else
            Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop + mPara - 5, mID$(MSHFlexGrid3.TextMatrix(T, 7), t1, mMax))
        End If
        ''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop + mPara - 5, mID$(MSHFlexGrid3.TextMatrix(T, 7), t1, mMax))
        ''------------------------------
        ''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop + mPara - 5, mID$(MSHFlexGrid3.TextMatrix(T, 7), t1, mMax))
    Next
    
    a1 = mPara + 5
    
    If a1 > mRowHeight Then
       tmp = a1 + 5
    Else
       tmp = mRowHeight
    End If
    mTxtTop = mTxtTop + tmp
    objPDF.PDFDrawLine mGlobalLeft + 65, mTxtTop - (tmp + 10), mGlobalLeft + 65, (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 145), mTxtTop - (tmp + 10), (mGlobalLeft + 145), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 195), mTxtTop - (tmp + 10), (mGlobalLeft + 195), (mTxtTop - 10)
    objPDF.PDFDrawLine (mGlobalLeft + 245), mTxtTop - (tmp + 10), (mGlobalLeft + 245), (mTxtTop - 10)
    ''-- Ullhas Patil 21 APril 2018
    If mSiloLocation = True Then
        objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (tmp + 10), (mGlobalLeft + 300), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (tmp + 10), (mGlobalLeft + 355), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 405), mTxtTop - (tmp + 10), (mGlobalLeft + 405), (mTxtTop - 10)
    Else
        objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (tmp + 10), (mGlobalLeft + 300), (mTxtTop - 10)
        objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (tmp + 10), (mGlobalLeft + 355), (mTxtTop - 10)
    End If
    ''objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (tmp + 10), (mGlobalLeft + 300), (mTxtTop - 10)
    ''objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (tmp + 10), (mGlobalLeft + 355), (mTxtTop - 10)
    ''------------------------------
    ''objPDF.PDFDrawLine (mGlobalLeft + 300), mTxtTop - (tmp + 10), (mGlobalLeft + 300), (mTxtTop - 10)
    ''objPDF.PDFDrawLine (mGlobalLeft + 355), mTxtTop - (tmp + 10), (mGlobalLeft + 355), (mTxtTop - 10)
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
    
    mTop = mTop + mRowHeight
    mHeight = mHeight + mRowHeight
Next
mTxtTop = mTxtTop - mRowHeight
mHeight = mTxtTop
''-- Ullhas Patil 21 APril 2018
If mSiloLocation = True Then
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Total")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(0, 13))
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(0, 5))
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop, MSHFlexGrid3.TextMatrix(0, 6))
Else
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Total")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(0, 5))
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(0, 6))
End If
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Total")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(0, 5))
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(0, 6))
''------------------------------
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 70, mTxtTop, "Total")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 250, mTxtTop, MSHFlexGrid3.TextMatrix(0, 5))
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, MSHFlexGrid3.TextMatrix(0, 6))
Call CheckForNewPage

mTxtTop = mTxtTop + (mRowHeight * 4)
mHeight = mTxtTop
Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Affix Hologram Here")
mTxtTop = mTxtTop + mRowHeight

objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight * 4 + 10), mGlobalLeft + 300, (mTxtTop - 10)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 305, mTxtTop, "Hologram No")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 405, mTxtTop, " " & RsTxn_CxTF_Details!HologramNo & " ")

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight + 10), mGlobalLeft + 300, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 400, mTxtTop - (mRowHeight + 10), mGlobalLeft + 400, (mTxtTop - 10)

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 100, mTxtTop - (mRowHeight + 10), mGlobalLeft + 100, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 200, mTxtTop - (mRowHeight + 10), mGlobalLeft + 200, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight + 10), mGlobalLeft + 300, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 400, mTxtTop - (mRowHeight + 10), mGlobalLeft + 400, (mTxtTop - 10)

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 100, mTxtTop - (mRowHeight + 10), mGlobalLeft + 100, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 200, mTxtTop - (mRowHeight + 10), mGlobalLeft + 200, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 300, mTxtTop - (mRowHeight + 10), mGlobalLeft + 300, (mTxtTop - 10)
objPDF.PDFDrawLine mGlobalLeft + 400, mTxtTop - (mRowHeight + 10), mGlobalLeft + 400, (mTxtTop - 10)

mTxtTop = mTxtTop + mRowHeight
mHeight = mTxtTop
Call CheckForNewPage

Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

mHeight = mTxtTop
Call CheckForNewPage
mTxtTop = mTxtTop + mRowHeight * 3

Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
mHeight = mTxtTop
Call CheckForNewPage

Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 8, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 16, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")


mTxtTop = mTxtTop + (mRowHeight * 4) + 16

mHeight = mTxtTop
Call CheckForNewPage
Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Depositor/Agent of Depositor :")
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

'''mHeight = mHeight + 50
'''mTxtTop = mHeight - (mTxtFromTop)

If mExchangeTypeID = 1 Then
   mHeight = mTxtTop
   Call CheckForNewPage
   mHeight = mHeight + 50
   
   Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Note: Please collect from us Warehouse Receipt for the deposited commodities at the earliest. Please note that deposited goods can be withdrawn only against production of valid ")
   Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop + 8, "Warehouse Receipt issued by us and not against this CDTF.")
   mTxtTop = mTxtTop + mRowHeight + 8
End If
''''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
mHeight = mTxtTop
Call CheckForNewPage
Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "**'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html")
'''
objPDF.PDFEndDoc

Retval = objPDF.PDFPageNumber
GetTotalPDFPages = Retval
End Function

Private Sub CheckForReservationMessage()

If CmbExchangeTypeID.Text = "" Then Exit Sub
If CmbCMPID.Text = "" Then Exit Sub
If CmbCommodityID.Text = "" Then Exit Sub

If Edit_Flg = "A" Then
   If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
      If Check1.Value = 0 Then
         ans = MsgBox("Is this deposit is a reservation deposit ? ", vbYesNo)
         If ans = vbYes Then
            Check1.Value = 1
         End If
      End If
   End If
End If

End Sub


Private Function CheckRateCards() As Boolean

Dim Retval As Boolean
Retval = True

tmp = "Select CommodityGroupID From Mst_Commodity Where CommodityID=" & mCommodityID
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mCommodityGroupId = TmpRs!CommodityGroupID
End If

If LCase(CmbExchangeTypeID) = "non ncdex" Then
   If Val(mCommodityYearlyFlag) = 1 And ChkGodownYearlyFlag = True Then
   
      tmp = "Select GodownID,Rate from Mst_YearlyStorageRateCardDetail MNRC"
      tmp = tmp & " inner Join Mst_YearlyStorageRateCard MNR on MNRC.YSRCID=MNR.YSRCID"
      tmp = tmp & " Where MNR.CommodityID=" & mCommodityID & "  And MNR.LocationID=" & mLocationID
       
      Call TmpTable
      
      If TmpRs.RecordCount = 0 Then
         Retval = False
         CheckRateCards = Retval
         Exit Function
      End If
   
   Else
   
      tmp = "Select TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,MC.ClientName,TCG.Mst_GSL_ID,MG.GodownID,GSL.CommodityID,MG.GodownNo,MG.GodownNo+'~'+"
      tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID"
      tmp = tmp & " From  Txn_CxTF_GSL TCG"
      tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
      tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
      tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
      tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientID=MC.ClientID"
      tmp = tmp & " Where TC.ExchangeTypeID=1 And TC.CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TC.TxnType='D' And TC.SM_Prefix='" & TxtPreFix.Text & "' And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0"
    
      Call TmpTable
    
      If TmpRs.RecordCount > 0 Then
         For I = 1 To TmpRs.RecordCount
             If CheckBillingDetails(mCommodityGroupId) = False Then
                Retval = False
                CheckRateCards = Retval
                Exit Function
             End If
             TmpRs.MoveNext
         Next
         Call UpdateBillingDetails
      End If
   End If

ElseIf LCase(CmbExchangeTypeID) = "ncdex" Then
   
   tmp = " Select TCD.CxTFDate,MG.CDTFDate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfDay,MG.ClientID,MClient.ClientName,"
   tmp = tmp & " TCG.Mst_GSL_ID, MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo as GSL,MG.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,0 as 'BagSize',MG.NContractID"
   tmp = tmp & " ,'Daily' Cycle,'Qty' Unit,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfCycle,MNRC.Rate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) * TCG.DematWeight * MNRC.Rate as 'Amount',ML.Collect7Days,MG.LCDTF_Date"
   tmp = tmp & " From Txn_CxTF_Details TCD"
   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTF_DetailsID = TCG.CxTF_DetailsID And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
   tmp = tmp & " Inner Join Mst_GSL MG On TCG.Mst_GSL_ID = MG.Mst_GSL_ID And TCG.SM_Prefix = MG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
   tmp = tmp & " Left Join Mst_Client MClient On MG.ClientID = MClient.ClientID"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCard MNR On ML.LocationID = MNR.LocationID And MComm.CommodityGroupID = MNR.CommodityGroupID"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And MG.GodownID = MNRC.GodownID"
   tmp = tmp & " Where TCD.TxnType = 'D' and TCD.ExchangeTypeID = 2 And TCD.CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TCD.SM_Prefix='" & TxtPreFix.Text & "' And MG.SpillageFlag = 0 And MG.MadeUpFlag= 0 "
      
   Call TmpTable
    
   If TmpRs.RecordCount > 0 Then
      For I = 1 To TmpRs.RecordCount
          If IIf(IsNull(TmpRs!Rate), 0, TmpRs!Rate) <= 0 Then
             MsgBox "NCDEX Rate is not updated for GSL " & TmpRs!GSL
             Retval = False
             CheckRateCards = Retval
             Exit Function
          End If
          TmpRs.MoveNext
      Next
   End If
End If

If CheckfranchiseeRateCard(mCommodityGroupId) = False Then
   Retval = False
End If

CheckRateCards = Retval

End Function

Private Function CheckBillingDetails(mCommodityGroupId_ As Variant) As Boolean

Dim Retval As Boolean
Retval = True

'tmp = "Select SPRCID,MS.BillingCycleID,MB.BillingCycle,MB.NoofDays,BillingUnit,Amount,MB1.BillingCycle as DepoBillingCycle,MS.DepoBillingCycleID,MB1.NoofDays as DepoNoOfDays"
'tmp = tmp & " From Mst_SpecialRateCard MS"
'tmp = tmp & " Inner Join Mst_BillingCycle MB on MS.BillingCycleID=MB.BillingCycleID"
'tmp = tmp & " left Join Mst_BillingCycle MB1 on MS.DepoBillingCycleID=MB1.BillingCycleID "
'tmp = tmp & " Where ClientIDRow = " & TmpRs!ClientIDRow & " And CommodityGroupID = " & mCommodityGroupId_ & " And LocationID =" & mLocationID & "  And ExchangeTypeID = 1 And MS.Flag = 'A'"


tmp = " Select Distinct TDCD.DepConID,TDC.ClientID,BagSizeFlag  from Txn_DepositContractGodownDetail TDCD"
tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID"
tmp = tmp & " Where TDC.SM_Prefix = '" & TxtPreFix & "' And TDC.ClientID = " & TmpRs!ClientID & " And TDCD.GodownId = " & TmpRs!GodownID & " And TDC.CommodityID = " & TmpRs!CommodityID & " And ApprovalFlag in ('S','A')  And '" & CDate(TxtCxTFDate) & "' Between StartDate and EndDate And Verified = 'Y' And ContractType = 'N' "

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   If LCase(TmpRs2!BagSizeFlag) = LCase("N") Then
   Else
      'tmp = " Select RatePerBag From Mst_SpecialRateCardDepoDetail Where " & Round(TmpRs!BagSize, 0) & " <= ToBagSize And " & Round(TmpRs!BagSize, 0) & " >= FromBagSize And  SPRCID = " & TmpRs2!SPRCID & " "
      tmp = " Select DepositFinalPrice From Txn_DepositContractGodownDetail Where GodownId = " & TmpRs!GodownID & " And DepConID  = " & TmpRs2!DepConID & " And " & Round(TmpRs!BagSize, 4) & " <= MaxBagSize And " & Round(TmpRs!BagSize, 4) & " >= MinBagSize"
      Call Tmp1Table
      
      If TmpRs1.RecordCount > 0 Then
      Else
         MsgBox "Bags size not in range in Deposit Contract " & TmpRs2!DepConID & "  for client " & TmpRs!ClientName & " !!!"
         Retval = False
         CheckBillingDetails = Retval
         Exit Function
      End If
   End If
Else
'Commented on 06 Dec 2017 By Parag after mail from Prit and Mayur Sir to implement the same
'   tmp = "Select SRCID,MS.BillingCycleID,MB.BillingCycle,MB.NoofDays,MS.Flag,RateOnQty,MB1.BillingCycle as DepoBillingCycle,MS.DepoBillingCycleID,MB1.NoofDays as DepoNoOfDays"
'   tmp = tmp & " From Mst_StorageRateCard MS "
'   tmp = tmp & " Inner Join Mst_BillingCycle MB on MS.BillingCycleID=MB.BillingCycleID "
'   tmp = tmp & " Inner Join Mst_BillingCycle MB1 on MS.DepoBillingCycleID=MB1.BillingCycleID "
'   tmp = tmp & " Where CommodityGroupID = " & mCommodityGroupId_ & " And LocationID =" & mLOcationID & "  And ExchangeTypeID = 1 And DefaultRate = 1 "
'
'   Call Tmp2Table
'
'   If TmpRs2.RecordCount > 0 Then
'      If TmpRs2!Flag = "Q" Then
'
'      Else
'         tmp = "Select RatePerBag From Mst_StorageRateCardDepoDetail Where " & Round(TmpRs!BagSize, 0) & " <= ToBagSize And " & Round(TmpRs!BagSize, 0) & ">=FromBagSize And SRCID=" & TmpRs2!SRCID
'         Call Tmp1Table
'         If TmpRs1.RecordCount > 0 Then
'
'         Else
'            MsgBox "Bags size not in range for client " & TmpRs!ClientName & " in Base Rate Card!!"
'            Retval = False
'            CheckBillingDetails = Retval
'            Exit Function
'         End If
'       End If
'   Else
'      MsgBox "Rate Card not found for client " & TmpRs!ClientName & "!!"
'      Retval = False
'      CheckBillingDetails = Retval
'      Exit Function
'   End If


      MsgBox "Approved and Verfied Normal Rate Card not found for Client " & TmpRs!ClientName & " , Godown - " & TmpRs!GodownNo & " , Commodity - " & TxtCommodityID & " !!"
      Retval = False
      CheckBillingDetails = Retval
      Exit Function

End If

CheckBillingDetails = Retval

End Function

Private Function CheckfranchiseeRateCard(mCommodityGroupId_ As Variant) As Boolean
Dim Retval As Boolean
Retval = True

tmp = "Select distinct MGod.GodownNo, mGod.GodownID"
tmp = tmp & " From Txn_CxTF_Details TCD"
tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTF_DetailsID = TCG.CxTF_DetailsID And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
tmp = tmp & " Inner Join Mst_GSL MG On TCG.Mst_GSL_ID = MG.Mst_GSL_ID And TCG.SM_Prefix = MG.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
tmp = tmp & " Where MGod.UnitTypeID=2 And TCD.TxnType = 'D' and TCD.CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TCD.SM_Prefix='" & TxtPreFix.Text & "' And MG.SpillageFlag = 0 And MG.MadeUpFlag= 0"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
   
       tmp = "Select  MNR.FF_RateCardID"
       tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
       tmp = tmp & " Where MNR.CommodityGroupID=" & mCommodityGroupId_ & " and b.GodownID=" & TmpRs!GodownID & " And b.BillingUnit in ('P')  And MNR.Flag='A'"
       Call Tmp1Table
       If TmpRs1.RecordCount = 0 Then
          tmp = "Select  MNR.FF_RateCardID"
          tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
          tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
          tmp = tmp & " Where MNR.CommodityGroupID=" & mCommodityGroupId_ & " and b.GodownID=" & TmpRs!GodownID & " And b.BillingUnit in ('Q')  And MNR.Flag='A'"
          Call Tmp1Table
          If TmpRs1.RecordCount = 0 Then
             tmp = "Select  MNR.FF_RateCardID"
             tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
             tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
             tmp = tmp & " inner Join Mst_FF_RateCard_Godown_Fix c on c.FF_RateCard_GodownID=b.FF_RateCard_GodownID"
             tmp = tmp & " Where b.GodownID=" & TmpRs!GodownID & " And b.BillingUnit in ('F')  And MNR.Flag='A'"
             tmp = tmp & " And '" & Format(TxtCxTFDate.Value, Gen_DatFormat) & "'>=c.StartDate And '" & Format(TxtCxTFDate.Value, Gen_DatFormat) & "'<=c.EndDate"
             Call Tmp1Table
             If TmpRs1.RecordCount = 0 Then
             
               
                tmp = "Select TC.CxTFDate,GSL.CDTFDate,GSL.ClientID,MC.ClientName,TCG.Mst_GSL_ID,MG.GodownNo+'~'+"
                tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,GSL.GodownID"
                tmp = tmp & " From  Txn_CxTF_GSL TCG"
                tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
                tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
                tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
                tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientID=MC.ClientID"
                tmp = tmp & " Where TC.CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TC.TxnType='D' And TC.SM_Prefix='" & TxtPreFix.Text & "' And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 And GSL.GodownID=" & TmpRs!GodownID
                               
                Call Tmp2Table
                For Ln = 1 To TmpRs2.RecordCount
                
                    tmp = "Select  MNR.FF_RateCardID"
                    tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
                    tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
                    tmp = tmp & " inner Join Mst_FF_RateCard_Godown_Bag c on c.FF_RateCard_GodownID=b.FF_RateCard_GodownID"
                    tmp = tmp & " Where " & Round(TmpRs2!BagSize) & " <= BagSizeTo And " & Round(TmpRs2!BagSize) & " >= BagSizeFrom  And MNR.CommodityGroupID=" & mCommodityGroupId_ & " And b.GodownID=" & TmpRs2!GodownID & " And b.BillingUnit in ('B')  And MNR.Flag='A'"

                    
                    Call Tmp3Table
                    If TmpRs3.RecordCount = 0 Then
                       Retval = False
                       Exit For
                    End If
                    TmpRs2.MoveNext
                Next
             End If
          End If
       End If
       TmpRs.MoveNext
   Next
End If
CheckfranchiseeRateCard = Retval
End Function

Private Function ChkOneGSLMultiCDTF() As Boolean

Dim Retval As Boolean
Dim mGSLChkforDoubleDepo As Variant

Retval = False

mGSLChkforDoubleDepo = ""

For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 1) = "" Then Exit For
    If mGSLChkforDoubleDepo = "" Then
       mGSLChkforDoubleDepo = "(" & Val(MSHFlexGrid4.TextMatrix(I, 13))
    Else
       mGSLChkforDoubleDepo = mGSLChkforDoubleDepo & "," & Val(MSHFlexGrid4.TextMatrix(I, 13))
    End If
Next
If mGSLChkforDoubleDepo <> "" Then
   mGSLChkforDoubleDepo = mGSLChkforDoubleDepo & ")"

   tmp = " Select ML.LocationName,TCG.Mst_GSL_ID,TCG.SM_Prefix,MG.Flag,Count(Distinct TCG.CxTF_DetailsID)"
   tmp = tmp & " From Txn_CxTF_Details TCD"
   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTF_DetailsID = TCG.CxTF_DetailsID ANd TCD.TxnType = TCG.TxnType And TCD.SM_Prefix = TCG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_GSL MG On TCG.Mst_GSL_ID = MG.Mst_GSL_ID ANd TCG.SM_Prefix = MG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Location ML On MG.SM_Prefix = ML.SM_Prefix"
   tmp = tmp & " Where TCD.ExchangeTypeID = 1 And TCG.TxnType = 'D' And TCG.SM_Prefix = '" & TxtPreFix & "' And TCG.Mst_GSL_ID in " & mGSLChkforDoubleDepo & ""
   tmp = tmp & " Group By ML.LocationName,TCG.Mst_GSL_ID ,TCG.SM_Prefix,MG.Flag"
   tmp = tmp & " Having Count(Distinct TCG.CxTF_DetailsID) > 1"

   Call TmpTable

   If TmpRs.RecordCount > 0 Then
      Retval = True
      For I = 1 To TmpRs.RecordCount
          For x = 2 To MSHFlexGrid4.Rows - 1
              If Val(MSHFlexGrid4.TextMatrix(x, 13)) = TmpRs!Mst_GSL_ID Then
                 For C = 0 To MSHFlexGrid4.Cols - 1
                     MSHFlexGrid4.row = x
                     MSHFlexGrid4.Col = C
                     MSHFlexGrid4.CellBackColor = vbRed
                 Next
              End If
          Next
          TmpRs.MoveNext
      Next
   End If
End If

ChkOneGSLMultiCDTF = Retval

End Function

Public Function CheckMoisturePer(RowVal As Integer) As Double
Dim Retval As Double
Retval = 0

If RsMst_GSL!Status = "B" Then
   Retval = MSHFlexGrid4.TextMatrix(RowVal, 20)
Else
   If RsMst_GSL!BalanceWeight + Val(MSHFlexGrid4.TextMatrix(RowVal, 10)) > 0 Then
      Retval = (RsMst_GSL!BalanceWeight * IIf(IsNull(RsMst_GSL!MoisturePer), 0, RsMst_GSL!MoisturePer)) + (Val(MSHFlexGrid4.TextMatrix(RowVal, 10)) * MSHFlexGrid4.TextMatrix(RowVal, 20))
      Retval = Retval / (RsMst_GSL!BalanceWeight + Val(MSHFlexGrid4.TextMatrix(RowVal, 10)))
   End If
End If
   
CheckMoisturePer = Retval
End Function
Public Function CheckMoisturePerDelete(RowVal As Integer) As Double
Dim Retval As Double
Retval = 0

If RsMst_GSL!Status <> "B" Then
   Retval = (RsMst_GSL!BalanceWeight * IIf(IsNull(RsMst_GSL!MoisturePer), 0, RsMst_GSL!MoisturePer)) - (TmpRs!DematWeight * TmpRs!MoisturePer)
   If Retval > 0 Then
      If RsMst_GSL!BalanceWeight - TmpRs!DematWeight > 0 Then
         Retval = Retval / (RsMst_GSL!BalanceWeight - TmpRs!DematWeight)
      End If
   End If
End If
   
CheckMoisturePerDelete = Retval
End Function

Private Sub CmbVehicleTypeID_GotFocus()
tmp = CmbVehicleTypeID.Text
Call TableMst_VehicleType
CmbVehicleTypeID.Clear
If RsMst_VehicleType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_VehicleType.RecordCount
    CmbVehicleTypeID.AddItem RsMst_VehicleType!VehicleType
    RsMst_VehicleType.MoveNext
Next
CmbVehicleTypeID.Text = tmp
End Sub
'----
Private Sub CmbVehicleTypeID_LostFocus()
If CmbVehicleTypeID.Text = "" Then
   Exit Sub
End If
RsMst_VehicleType.MoveFirst

If tmp <> CmbVehicleTypeID.Text Then
   TxtMinimum = ""
   TxtMaximum = ""
End If

RsMst_VehicleType.Find "VehicleType = '" & CmbVehicleTypeID.Text & "'"
If RsMst_VehicleType.EOF Then
   MsgBox "Invalid selection "
   CmbVehicleTypeID.SetFocus
   Exit Sub
End If
mVehicleTypeID = RsMst_VehicleType!VehicleTypeID
TxtMinimum = IIf(IsNull(RsMst_VehicleType!MinTareWt), 0, RsMst_VehicleType!MinTareWt)
TxtMaximum = IIf(IsNull(RsMst_VehicleType!MaxTareWt), 0, RsMst_VehicleType!MaxTareWt)

End Sub

Private Function ChkTareWeightDiscripency() As Boolean
Dim Retval As Boolean
Retval = False

For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
       If Val(MSHFlexGrid3.TextMatrix(I, 4)) < Val(MSHFlexGrid3.TextMatrix(I, 11)) Or Val(MSHFlexGrid3.TextMatrix(I, 4)) > Val(MSHFlexGrid3.TextMatrix(I, 12)) Then
          Retval = True
          Exit For
       End If
    End If
Next
ChkTareWeightDiscripency = Retval

End Function

'Gen_MailTareWeightDiscripency

Private Sub Gen_MailTareWeightDiscripency()

Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
TOADD = ""
TOCC = ""

mMsg = "The Tare weight entered for the following vehicles are beyond the approved range. Please look into the matter on priority basis"
mMsg = mMsg & "<br><br><br>"
mMsg = mMsg & "<table border = 1 face= Candara><tr><td colspan=2>CDTF Details (From Store-Man)</td></tr>"
mMsg = mMsg & "<tr><td>CDTF No</td><td>" & TxtCxTFNo.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>CDTF Date</td><td>" & Format(TxtCxTFDate.Value, "dd-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & "<tr><td>Location</td><td>" & GetLocationFromCMP(RsTxn_CxTF_Details!CMPID) & "</td> </tr>"
mMsg = mMsg & "<tr><td>State</td><td>" & GetStateFromLocation(GetLocationIDFromCMP(mCMPID)) & "</td></tr>"
mMsg = mMsg & "<tr><td>CMPID & Address</td><td>" & mCMPID & "-" & GetCMPAddress(mCMPID, "A") & "</td></tr>"
mMsg = mMsg & "<tr><td>Commodity</td><td>" & CmbCommodityID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Transaction Type</td><td>" & CmbExchangeTypeID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Name Of Agent </td><td>" & CmbAgent.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Total No of Bags</td><td>" & MSHFlexGrid2.TextMatrix(0, 5) & "</td></tr>"
mMsg = mMsg & "<tr><td>Total Electronic Weight in MT </td><td>" & MSHFlexGrid2.TextMatrix(0, 6) & "</td></tr>"

mGridMsg = ""

For I = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
       If I = 1 Then
          mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid3.TextMatrix(I, 9) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2) & " </td><td>Approved Tare Weight Range</td><td> " & MSHFlexGrid3.TextMatrix(I, 4) & " </td><td>Remarks</td></tr>"
       Else
          mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid3.TextMatrix(I, 9) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 11) & "~" & MSHFlexGrid3.TextMatrix(I, 12) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 7) & " </td></tr>"
       End If
    End If
Next

mMsg = mMsg & "<tr><td> TTD Details </td><td><table border = 1 face= Candara>" & mGridMsg & "</table></td></table>"


tmp = " Select EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID,EMP1.EmployeeID as 'AMEmpID',EMP1.EmployeeNo as 'AMEmpNo',EMP1.EmployeeName as 'AMName',EMP1.EmployeeEmailID as 'AMEmailID'" & _
      " From Mst_EmpCMPMapping EGM" & _
      " Inner Join Mst_Employee EMP on EGM.EmployeeID = EMP.EmployeeID" & _
      " Inner Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID" & _
      " Where EMP.Flag = 'Y' And EGM.CMPID = " & mCMPID
'      " Group By EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID"

Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
   TOADD = TmpRs3!EmployeeEmailID
   TOCC = IIf(IsNull(TmpRs3!AMEmailID), "", TmpRs3!AMEmailID) '""
End If

If TOADD = "" Then
   TOADD = "kuldip.y@ncml.com"
End If
 
 '"cmg@ncml.com" & "," & Gen_UserEmailID  'Gen_UserEmailID
If TOCC <> "" Then
   TOCC = TOCC & ",cmg@ncml.com," & Gen_UserEmailID
Else
   TOCC = Gen_UserEmailID
End If




mMsgSubject = "CDTF No - " & Val(TxtCxTFNo.Text) & " is having inconsistency in the tare weight entered in StoreMan at Location - " & GetLocationFromCMP(RsTxn_CxTF_Details!CMPID) & ""



'If ChkSCMFlag.Value = 0 Then
'   mMsgSubject = "NCMSL " & TxtCMPID.Text & " CDTF Details"
'Else
'   TOCC = TOCC & "," & "supplychain@ncml.com,scmquality@ncml.com"
'   mMsgSubject = "NCMSL CDTF No-" & TxtCxTFNo.Text & " for " & TxtCMPID.Text & "-Agent name: " & TxtAgent.Text & ""
'End If



Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "D")

End Sub


Private Function ChkGodownYearlyFlag() As Boolean
Dim Retval As Boolean

Retval = False

For I = 2 To MSHFlexGrid4.Rows - 1
    If Val(MSHFlexGrid4.TextMatrix(I, 21)) = 1 Then
       Retval = True
    End If
Next

ChkGodownYearlyFlag = Retval

End Function
