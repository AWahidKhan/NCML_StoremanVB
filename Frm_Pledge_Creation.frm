VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A4F98396-D7CE-459C-BA85-E874A5B1E44F}#3.3#0"; "OSSMTP.ocx"
Begin VB.Form Frm_Pledge_Creation 
   Caption         =   "Pledge Creation"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   8895
      Left            =   16080
      TabIndex        =   81
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   82
      Top             =   9960
      Width           =   10935
   End
   Begin VB.PictureBox Frame4 
      Height          =   9705
      Left            =   120
      ScaleHeight     =   9645
      ScaleWidth      =   15345
      TabIndex        =   16
      Top             =   120
      Width           =   15405
      Begin VB.Frame Frame0 
         Height          =   1965
         Left            =   240
         TabIndex        =   83
         Top             =   6480
         Visible         =   0   'False
         Width           =   7485
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   12210
            TabIndex        =   93
            Top             =   735
            Width           =   1000
         End
         Begin VB.ComboBox CmbSearchExchange 
            Height          =   315
            ItemData        =   "Frm_Pledge_Creation.frx":0000
            Left            =   10860
            List            =   "Frm_Pledge_Creation.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   92
            Top             =   330
            Width           =   2070
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   345
            Left            =   11205
            TabIndex        =   91
            Top             =   735
            Width           =   1005
         End
         Begin VB.TextBox TxtSearchPledge 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1410
            MaxLength       =   200
            TabIndex        =   90
            Top             =   675
            Width           =   1485
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "Clear"
            Height          =   345
            Left            =   13215
            TabIndex        =   89
            Top             =   735
            Width           =   1005
         End
         Begin VB.ComboBox CmbSBankName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4815
            Sorted          =   -1  'True
            TabIndex        =   88
            Top             =   330
            Width           =   2985
         End
         Begin VB.ComboBox CmbSBranchName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7845
            Sorted          =   -1  'True
            TabIndex        =   87
            Top             =   330
            Width           =   2985
         End
         Begin VB.ComboBox CmbPreFix 
            Height          =   315
            Left            =   2850
            Sorted          =   -1  'True
            TabIndex        =   86
            Top             =   315
            Width           =   1920
         End
         Begin VB.TextBox TxtSClientName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3810
            TabIndex        =   85
            Top             =   675
            Width           =   3540
         End
         Begin VB.ComboBox CmbSCommodity 
            Height          =   315
            Left            =   12960
            Sorted          =   -1  'True
            TabIndex        =   84
            Top             =   330
            Width           =   1830
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   615
            Left            =   120
            TabIndex        =   94
            Top             =   1200
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   1085
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   300
            Left            =   1470
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   51052545
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   120
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   330
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   51052545
            CurrentDate     =   36494
         End
         Begin VB.Label Label22 
            Caption         =   "Exchange Type"
            Height          =   225
            Left            =   11288
            TabIndex        =   107
            Top             =   120
            Width           =   1215
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
            Left            =   7605
            TabIndex        =   106
            Top             =   750
            Width           =   2355
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "_"
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
            Left            =   10005
            TabIndex        =   105
            Top             =   750
            Width           =   135
         End
         Begin VB.Label Label25 
            Caption         =   "WHR/CMSE Lot"
            Height          =   240
            Left            =   120
            TabIndex        =   104
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label29 
            Caption         =   "Bank Name"
            Height          =   255
            Left            =   5872
            TabIndex        =   103
            Top             =   120
            Width           =   870
         End
         Begin VB.Label Label30 
            Caption         =   "Branch Name"
            Height          =   225
            Left            =   8805
            TabIndex        =   102
            Top             =   120
            Width           =   1065
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   1800
            TabIndex        =   101
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
            Height          =   255
            Left            =   345
            TabIndex        =   100
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label36 
            Caption         =   "Location"
            Height          =   300
            Left            =   3473
            TabIndex        =   99
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label32 
            Caption         =   "Client Name"
            Height          =   210
            Left            =   2955
            TabIndex        =   98
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label33 
            Caption         =   "Commodity"
            Height          =   195
            Left            =   13440
            TabIndex        =   97
            Top             =   120
            Width           =   825
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   930
         Left            =   120
         ScaleHeight     =   870
         ScaleWidth      =   14205
         TabIndex        =   70
         Top             =   8565
         Width           =   14265
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Pledge Letter"
            Height          =   750
            Left            =   10965
            TabIndex        =   110
            Top             =   45
            Width           =   3150
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New "
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   45
            Width           =   2175
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   375
            Left            =   2295
            TabIndex        =   79
            Top             =   45
            Width           =   2175
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   4470
            TabIndex        =   78
            Top             =   45
            Width           =   2160
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   6630
            TabIndex        =   77
            Top             =   45
            Width           =   2160
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "&Search"
            Height          =   375
            Left            =   8790
            TabIndex        =   76
            Top             =   45
            Width           =   2175
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   375
            Left            =   120
            TabIndex        =   75
            Top             =   420
            Width           =   2175
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   2295
            TabIndex        =   74
            Top             =   420
            Width           =   2175
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   375
            Left            =   4470
            TabIndex        =   73
            Top             =   420
            Width           =   2160
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   375
            Left            =   6630
            TabIndex        =   72
            Top             =   420
            Width           =   2160
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   8790
            TabIndex        =   71
            Top             =   420
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1770
         Left            =   120
         TabIndex        =   56
         Top             =   3270
         Width           =   14220
         Begin VB.TextBox TxtBalFee 
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
            Left            =   9780
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   1320
            Width           =   1920
         End
         Begin VB.TextBox TxtReceivableFee 
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
            Left            =   5910
            Locked          =   -1  'True
            TabIndex        =   116
            Top             =   1320
            Width           =   1920
         End
         Begin VB.TextBox TxtRecFee 
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
            Left            =   7845
            Locked          =   -1  'True
            TabIndex        =   114
            Top             =   1320
            Width           =   1920
         End
         Begin VB.TextBox TxtDiscount 
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
            Left            =   4305
            TabIndex        =   10
            Top             =   1320
            Width           =   1590
         End
         Begin VB.TextBox TxtPledgeFee 
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
            Left            =   2370
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1920
         End
         Begin VB.TextBox TxtPledgeReqNo 
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
            TabIndex        =   8
            Top             =   600
            Width           =   1905
         End
         Begin VB.ComboBox CmbBankName 
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
            Left            =   2385
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   180
            Width           =   5505
         End
         Begin VB.ComboBox CmbBranchName 
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
            Left            =   2385
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   5505
         End
         Begin VB.CommandButton CmdBankName 
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
            Height          =   330
            Left            =   7935
            TabIndex        =   58
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton CmdBranchName 
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
            Height          =   330
            Left            =   7950
            TabIndex        =   57
            Top             =   600
            Width           =   375
         End
         Begin MSComCtl2.DTPicker PledgeRequestDate 
            Height          =   375
            Left            =   11880
            TabIndex        =   9
            Top             =   180
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51052545
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtBankName 
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
            Left            =   2400
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   180
            Width           =   5490
         End
         Begin VB.TextBox TxtBranchName 
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
            Left            =   2385
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   600
            Width           =   5505
         End
         Begin VB.Label Label42 
            Caption         =   "Balance Fee"
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
            Left            =   10110
            TabIndex        =   119
            Top             =   1020
            Width           =   1455
         End
         Begin VB.Label Label41 
            Caption         =   "Receivable Fee"
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
            Left            =   6165
            TabIndex        =   117
            Top             =   1005
            Width           =   1455
         End
         Begin VB.Label Label40 
            Caption         =   "Received Fee"
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
            Left            =   8160
            TabIndex        =   115
            Top             =   1020
            Width           =   1455
         End
         Begin VB.Label Label35 
            Caption         =   "Discount"
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
            Left            =   4680
            TabIndex        =   111
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label Label11 
            Caption         =   "Pledge Creation Fee"
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
            Left            =   2370
            TabIndex        =   65
            Top             =   1035
            Width           =   1950
         End
         Begin VB.Label Label5 
            Caption         =   "Pledge Request No "
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
            Left            =   9120
            TabIndex        =   64
            Top             =   660
            Width           =   2775
         End
         Begin VB.Label Label38 
            Caption         =   "Bank Name"
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
            TabIndex        =   63
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label Label39 
            Caption         =   "Branch Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   62
            Top             =   705
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Pledge Request Date"
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
            Left            =   9120
            TabIndex        =   61
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   3150
         Left            =   120
         ScaleHeight     =   3090
         ScaleWidth      =   14160
         TabIndex        =   17
         Top             =   120
         Width           =   14220
         Begin VB.TextBox TxtCarNo 
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
            Height          =   375
            Left            =   12450
            Locked          =   -1  'True
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   480
            Width           =   1515
         End
         Begin VB.TextBox TxtG_UID 
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
            Height          =   375
            Left            =   12450
            Locked          =   -1  'True
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1515
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   2640
            Width           =   4650
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
            Left            =   8505
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   2640
            Width           =   5340
         End
         Begin VB.TextBox TxtWHRNoNew 
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
            TabIndex        =   2
            Top             =   975
            Width           =   2670
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
            Height          =   375
            Left            =   10155
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   45
            Width           =   1515
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1575
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
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1800
            Width           =   6135
         End
         Begin VB.ComboBox CmbClientDP 
            Height          =   315
            Left            =   6120
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   960
            Visible         =   0   'False
            Width           =   7215
         End
         Begin VB.TextBox TxtGodownNo 
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
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   2670
         End
         Begin VB.ComboBox CmbExchange 
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
            ItemData        =   "Frm_Pledge_Creation.frx":0004
            Left            =   1920
            List            =   "Frm_Pledge_Creation.frx":0006
            Sorted          =   -1  'True
            TabIndex        =   0
            Top             =   120
            Width           =   2655
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
            Height          =   435
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2175
            Width           =   11940
         End
         Begin VB.TextBox TxtWHRQuantity 
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
            Left            =   12315
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1400
            Width           =   1515
         End
         Begin VB.TextBox TxtWHRNoOfBags 
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
            Left            =   9585
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1400
            Width           =   1605
         End
         Begin VB.TextBox TxtLicenseNo 
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
            Left            =   13560
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1395
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox TxtPledgeID 
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
            Left            =   6105
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   45
            Width           =   2700
         End
         Begin VB.TextBox TxtClientID 
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   500
            Width           =   2670
         End
         Begin VB.TextBox TxtCommodity 
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1395
            Width           =   2670
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   960
            Width           =   7710
         End
         Begin MSComCtl2.DTPicker PledgeDate 
            Height          =   375
            Left            =   1905
            TabIndex        =   1
            Top             =   525
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51052545
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtWHRNo 
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
            Left            =   1560
            TabIndex        =   15
            Top             =   975
            Visible         =   0   'False
            Width           =   510
         End
         Begin MSComCtl2.DTPicker WHRDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   29
            Top             =   1400
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   51052545
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker WHRValidDate 
            Height          =   375
            Left            =   10155
            TabIndex        =   108
            Top             =   480
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   51052545
            CurrentDate     =   39884
         End
         Begin VB.Label Label43 
            Caption         =   "Car No"
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
            Left            =   11760
            TabIndex        =   121
            Top             =   495
            Width           =   735
         End
         Begin VB.Label Label37 
            Caption         =   "G_UID"
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
            Left            =   11760
            TabIndex        =   113
            Top             =   90
            Width           =   645
         End
         Begin VB.Label Label34 
            Caption         =   "WHR Valid Date"
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
            Left            =   8970
            TabIndex        =   109
            Top             =   465
            Width           =   1095
         End
         Begin VB.Label Label28 
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
            Left            =   120
            TabIndex        =   69
            Top             =   2707
            Width           =   1695
         End
         Begin VB.Label Label27 
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
            Left            =   7470
            TabIndex        =   68
            Top             =   2700
            Width           =   975
         End
         Begin VB.Label Label31 
            Caption         =   "Location ID"
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
            Left            =   9000
            TabIndex        =   43
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label lblDpIDName 
            Caption         =   "DP ID : Name"
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
            Left            =   4800
            TabIndex        =   41
            Top             =   1860
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Godown No"
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
            TabIndex        =   38
            Top             =   1860
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Exchange Type"
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
            TabIndex        =   36
            Top             =   150
            Width           =   1680
         End
         Begin VB.Label Label19 
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
            Left            =   135
            TabIndex        =   35
            Top             =   2265
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "WHR Qty."
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
            Left            =   11250
            TabIndex        =   34
            Top             =   1460
            Width           =   945
         End
         Begin VB.Label Label16 
            Caption         =   "WHR No of Bags"
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
            Left            =   7980
            TabIndex        =   32
            Top             =   1460
            Width           =   1590
         End
         Begin VB.Label Label14 
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
            Height          =   255
            Left            =   4800
            TabIndex        =   30
            Top             =   1460
            Width           =   1215
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   12240
            TabIndex        =   28
            Top             =   1455
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Pledge ID"
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
            TabIndex        =   27
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label Label6 
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
            Left            =   4800
            TabIndex        =   26
            Top             =   660
            Width           =   1590
         End
         Begin VB.Label Label26 
            Caption         =   "Pledge  Date"
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
            TabIndex        =   25
            Top             =   585
            Width           =   1215
         End
         Begin VB.Label Label7 
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
            Height          =   255
            Left            =   135
            TabIndex        =   24
            Top             =   1455
            Width           =   1455
         End
         Begin VB.Label Label8 
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
            Height          =   255
            Left            =   4800
            TabIndex        =   23
            Top             =   1020
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "WHR No"
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
            TabIndex        =   22
            Top             =   1035
            Width           =   1335
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2685
         Left            =   120
         TabIndex        =   55
         Top             =   5880
         Width           =   14220
         _ExtentX        =   25083
         _ExtentY        =   4736
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   150
         TabIndex        =   44
         Top             =   4980
         Width           =   14220
         Begin VB.TextBox TxtUnpledgeWeight 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   7200
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox TxtUnpledgeNoOfBags 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   5475
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   420
            Width           =   1785
         End
         Begin VB.TextBox TxtCMP 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   420
            Width           =   3075
         End
         Begin VB.TextBox TxtGSL 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   3120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   420
            Width           =   2415
         End
         Begin VB.CommandButton CmdUpdate 
            Caption         =   "Update List"
            Height          =   360
            Left            =   12345
            TabIndex        =   14
            Top             =   420
            Width           =   1380
         End
         Begin VB.TextBox TxtPledgeNoOfBags 
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
            Left            =   8730
            TabIndex        =   12
            Top             =   420
            Width           =   1785
         End
         Begin VB.TextBox TxtPledgeQuantity 
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
            Left            =   10545
            TabIndex        =   13
            Top             =   420
            Width           =   1620
         End
         Begin VB.Label Label15 
            Caption         =   "Unpledged No Of Bags"
            Height          =   300
            Left            =   5520
            TabIndex        =   54
            Top             =   135
            Width           =   1710
         End
         Begin VB.Label Label10 
            Caption         =   "CMP"
            Height          =   300
            Left            =   1267
            TabIndex        =   53
            Top             =   165
            Width           =   1020
         End
         Begin VB.Label Label12 
            Caption         =   "G-S-L"
            Height          =   300
            Left            =   4050
            TabIndex        =   52
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label13 
            Caption         =   "Unpledged  Qty."
            Height          =   300
            Left            =   7560
            TabIndex        =   51
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label Label3 
            Caption         =   "Pledged No Of Bags"
            Height          =   255
            Left            =   9000
            TabIndex        =   50
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Pledged Qty."
            Height          =   255
            Left            =   10920
            TabIndex        =   49
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin OSSMTPControl.SMTP SMTP 
      Left            =   120
      Top             =   10320
      _ExtentX        =   1296
      _ExtentY        =   873
      AuthenticationType=   0
      BCC             =   ""
      CC              =   ""
      Charset         =   "us-ascii"
      ClientHostName  =   ""
      ContentTransferEncoding=   "7bit"
      ContentType     =   "text/plain"
      ExpiresAfter    =   ""
      Importance      =   0
      LastSMTPCommand =   ""
      MailFrom        =   ""
      MessageHTML     =   ""
      MessageSubject  =   ""
      MessageText     =   ""
      Notification    =   0
      Password        =   ""
      POPServer       =   ""
      Port            =   25
      RaiseError      =   0   'False
      ReplyTo         =   ""
      SendTo          =   ""
      Sensitivity     =   0
      Server          =   ""
      Status          =   ""
      Timeout         =   10000
      TimeStamp       =   ""
      Username        =   ""
   End
End
Attribute VB_Name = "Frm_Pledge_Creation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mClientIDRow As Double
Dim mCommodityID As Double
Dim mExchangeTypeID As Double
Dim mCMPID As Double
Dim mBankID As Double
Dim mBranchID As Double
Dim mWHRNo As Double
Dim mISINID As Variant
Dim mGodownID As Variant
Dim mExchangeID As Variant
Dim mSBankID As Double
Dim mPreMsg, mPostMsg, mGodownAddress, mRate, mValueOfCommodity As Variant
Dim mCMAgreement As Variant
Dim mPrintFlag As Variant
Dim mCARNo As Variant
Dim mPledgeFlag As Variant
Dim mFontSize As Integer
Dim mTxtLeft As Double
Dim mTxtTop As Double
Dim mSCMFlag As Variant
Dim mNCMLFinance As Variant
Dim mSendToMktYard, mMarketYard, mMarketYardPledge As Variant
Dim mLotWiseWHR As Variant
Dim mWHR_Detail_ID As Variant

Private Sub AddCmd_Click()

Edit_Flg = "A"

Picture2.Enabled = True
Frame3.Enabled = True
Frame1.Enabled = True

TxtWHRNo.Locked = False
TxtWHRNoNew.Locked = False
CmbBankName.Visible = True
CmbBranchName.Visible = True
CmdBankName.Visible = True
CmdBranchName.Visible = True
TxtBankName.Visible = False
TxtBranchName.Visible = False
CmbExchange.Enabled = True
PledgeDate.Value = Date
PledgeDate.Enabled = True
PledgeRequestDate.Value = Date
CmdUpdate.Enabled = False
TxtPledgeNoOfBags.Locked = True
TxtPledgeQuantity.Locked = True
TxtPledgeReqNo.Locked = False
'TxtPledgeFee.Locked = False
TxtDiscount.Locked = False
PledgeRequestDate.Enabled = True
TxtRemarks.Locked = False
CmbExchange.Text = ""
mExchangeTypeID = 0
TxtDPID.Visible = False
TxtDPName.Visible = False
TxtClientID.BackColor = -2147483648#
TxtG_UID = ""
TxtPreFix = ""
TxtWHRNoNew = ""
TxtCreated = ""
TxtUpdated = ""
CmdPrint.Enabled = False
mPrintFlag = Null
TxtDiscount = ""
TxtReceivableFee = ""
TxtRecFee = ""
TxtBalFee = ""
mGodownAddress = "" 'GetGodownName(IIf(IsNull(RsTxn_WHRDetail!GodownID), "", RsTxn_WHRDetail!GodownID), "A")
mRate = 0 'IIf(IsNull(RsTxn_WHRDetail!Rate), "", RsTxn_WHRDetail!Rate)
mValueOfCommodity = 0 'Val(TxtWHRQuantity) * Val(mRate)

mSCMFlag = 0
Call ClearData
Call GButtonLock(Me, False)
Call TableTxn_WHRDetail(" Where SCMFlag = 0 And isnull(CARNo,0) = 0 ")
Call TableMst_ISIN
If Gen_WHRDetail(0) = False Then
   CmbExchange = Gen_WHRDetail(4)
   Call CmbExchange_LostFocus
   TxtWHRNoNew = Gen_WHRDetail(3)
   Call TxtWHRNoNew_LostFocus
End If

End Sub

Private Sub CmbBranchName_GotFocus()
If CmbBankName.Text = "" Then Exit Sub
If IsNull(mBankID) Then Exit Sub
tmp = CmbBranchName.Text
Call TableMst_Branch(" Where BankID =  " & mBankID & " Order by  BranchID")
CmbBranchName.Clear
For I = 1 To RsMst_Branch.RecordCount
    CmbBranchName.AddItem RsMst_Branch!BranchName
    RsMst_Branch.MoveNext
Next
CmbBranchName.Text = tmp
End Sub
Private Sub CmbBranchName_LostFocus()
If CmbBranchName.Text = "" Then
   Exit Sub
End If
RsMst_Branch.MoveFirst
RsMst_Branch.Find "BranchName = '" & CmbBranchName.Text & "'"
If RsMst_Branch.EOF Then
   MsgBox "Invalid selection!!!"
   CmbBranchName.SetFocus
   Exit Sub
End If
mBranchID = RsMst_Branch!BranchID
End Sub

Private Sub CmbExchange_GotFocus()
tmp = CmbExchange.Text
Call TableMst_ExchangeType
CmbExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    If RsMst_ExchangeType!ExchangeType <> "SPOT" Then
       CmbExchange.AddItem RsMst_ExchangeType!ExchangeType
    End If
    RsMst_ExchangeType.MoveNext
Next
CmbExchange.Text = tmp
TxtCommodity = ""
TxtLicenseNo = ""
End Sub

Private Sub CmbExchange_LostFocus()
If CmbExchange.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection!!! "
   CmbExchange.SetFocus
   Exit Sub
End If
If CmbExchange.Text = "NCDEX" Then
   Call EnDn("NCDEX", False)
   mClientIDRow = 0
   TxtClientID.BackColor = -2147483643
   TxtClientID.Enabled = True
   TxtClientID.Locked = False
   lblDpIDName.Visible = True
   TxtDPID.Visible = True
   TxtDPName.Visible = True
Else
   Call EnDn("Non NCDEX", True)
   WHRDate.Enabled = False
   TxtClientID.BackColor = -2147483648#
   TxtClientID.Enabled = False
   TxtClientID.Locked = True
  
   lblDpIDName.Visible = False
   TxtDPID.Visible = False
   TxtDPName.Visible = False
End If
If tmp <> CmbExchange.Text Then
   Call Grid_Head
End If
'TxtWHRNoNew.SetFocus
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
Call TxtWHRNo_LostFocus

End Sub
Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_Bank(" Where Isnull(MarketYard,0) = " & mSendToMktYard & " ")

CmbBankName.Clear
If RsMst_Bank.RecordCount = 0 Then
   MsgBox "No Bank found!!!"
   'CmdInsuranceCo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Bank.RecordCount
    CmbBankName.AddItem RsMst_Bank!BankName
    RsMst_Bank.MoveNext
Next
CmbBankName.Text = tmp
End Sub
Private Sub CmbBankName_LostFocus()
If CmbBankName.Text = "" Then
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankName = '" & CmbBankName.Text & "'"
If RsMst_Bank.EOF Then
   MsgBox "Invalid Bank selection!!!"
   CmbBankName.SetFocus
   Exit Sub
End If
mBankID = RsMst_Bank!BankID
If CmbBankName.Text <> tmp Then
   CmbBranchName.Text = ""
   TxtPledgeReqNo = ""
   TxtPledgeFee = ""
End If
mCMAgreement = IIf(IsNull(RsMst_Bank!CMAgreement), 0, RsMst_Bank!CMAgreement)
mNCMLFinance = IIf(IsNull(RsMst_Bank!NCMLFinance), 0, RsMst_Bank!NCMLFinance)
mMarketYard = IIf(IsNull(RsMst_Bank!MarketYard), 0, RsMst_Bank!MarketYard)
'If TxtWHRNoNew.Text <> "" Then
Call GetPledgeFee(mCMAgreement)
'End If
   
End Sub
Private Sub GetPledgeFee(mCm_Agee_ As Variant)

TxtPledgeFee = 0
If TxtWHRNoNew.Text = "" Then Exit Sub
If CmbBankName.Text = "" Then Exit Sub

tmp = "Select FeesForCM,FeesForWithoutCM From Mst_Location Where SM_Prefix='" & TxtPreFix & " '"
Call Tmp3Table
If TmpRs3.RecordCount > 0 Then
   If mCm_Agee_ = 1 Then
      TxtPledgeFee.Text = TmpRs3!FeesForCM
   Else
      TxtPledgeFee.Text = TmpRs3!FeesForWithoutCM
   End If
'Else
'   TxtPledgeFee.Text = 0
End If
TxtBalFee = TxtPledgeFee
TxtReceivableFee = TxtPledgeFee
End Sub

Private Sub CmbSBankName_GotFocus()

tmp = CmbSBankName.Text
Call TableMst_Bank("")

CmbSBankName.Clear
If RsMst_Bank.RecordCount = 0 Then
   MsgBox "No Bank found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Bank.RecordCount
    CmbSBankName.AddItem RsMst_Bank!BankName
    RsMst_Bank.MoveNext
Next
CmbSBankName.Text = tmp

End Sub

Private Sub CmbSBankName_LostFocus()

If CmbSBankName.Text = "" Then
   CmbSBranchName.Text = ""
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankName = '" & CmbSBankName.Text & "'"
If RsMst_Bank.EOF Then
   MsgBox "Invalid Bank selection!!!"
   CmbSBankName.SetFocus
   Exit Sub
End If
mSBankID = RsMst_Bank!BankID
If CmbSBankName.Text <> tmp Then
   CmbSBranchName.Text = ""
End If

End Sub

Private Sub CmbSBranchName_GotFocus()
If CmbSBankName.Text = "" Then
   MsgBox "Select Bank!!!"
   CmbSBankName.SetFocus
   Exit Sub
End If
tmp = CmbSBranchName.Text
Call TableMst_Branch(" Where BankID =  " & mSBankID & " Order by  BranchID")
CmbSBranchName.Clear
For I = 1 To RsMst_Branch.RecordCount
    CmbSBranchName.AddItem RsMst_Branch!BranchName
    RsMst_Branch.MoveNext
Next
CmbSBranchName.Text = tmp
End Sub

Private Sub CmbSBranchName_LostFocus()
If CmbSBranchName.Text = "" Then
   Exit Sub
End If
RsMst_Branch.MoveFirst
RsMst_Branch.Find "BranchName = '" & CmbSBranchName.Text & "'"
If RsMst_Branch.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSBranchName.SetFocus
   Exit Sub
End If
mSBranchID = RsMst_Branch!BranchID
End Sub

Private Sub CmdBankName_Click()
FrmMst_Bank.Show
End Sub

Private Sub CmdBranchName_Click()
FrmMst_Branch.Show
End Sub

Private Sub CmdClear_Click()
SFrom.Value = Date - 365
STo.Value = Date
CmbPreFix.Text = ""
CmbSBankName.Text = ""
CmbSBranchName.Text = ""
CmbSearchExchange.Text = ""
CmbSCommodity.Text = ""
TxtSearchPledge = ""
TxtSearchPledge.Locked = True
Label25.Caption = "WHR/CMSE Lot No"
Label24.Caption = ""
Call Grid_Head1

End Sub

Private Sub CmdGo_Click()
Dim wc As String
Dim tmp1 As String

Call Grid_Head1


If STo.Value < SFrom.Value Then
   MsgBox "'To date' must be always greater than 'From date'!!!"
   STo.SetFocus
   Exit Sub
End If

wc = ""

tmp = " Select PM.Txn_PledgeCreation_ID,ML.LocationName,PM.PledgeDate,E.ExchangeType,WD.WHRNo,WD.CMPID,PM.ISINID as ISINID,Bank.BankName,"
tmp = tmp & " Branch.BranchName,  MC.ClientName , PM.PledgeReqNo, PM.PledgeFees, PM.PledgeReqDate, PM.Remarks,"
tmp = tmp & " PM.G_UID,Sum(PledgeNoOfBags) as 'PledgeNoOfBags',Sum(PledgeQuantity) as 'PledgeQuantity',MCom.Commodity,MCom1.Commodity 'NCDEXCommodity',PM.Discount,PM.RecFees,PM.NCMLFinance"
tmp = tmp & " From Txn_PledgeCreation PM"
tmp = tmp & " Left Join Txn_WHR_Detail WD on PM.WHRNo = WD.WHRNo"
tmp = tmp & " Inner Join Mst_Bank Bank On PM.BankID = Bank.BankID"
tmp = tmp & " Inner Join Mst_Branch Branch On PM.BranchID = Branch.BranchID"
tmp = tmp & " Inner Join Mst_ExchangeType E On PM.ExchangeTypeID = E.ExchangeTypeID"
tmp = tmp & " Left  Join Mst_Client MC On PM.ClientID = MC.ClientID"
'tmp = tmp & " Left Join Txn_ImportNcdexLotPleadge I On PM.ISINID = I.Code"
tmp = tmp & " Inner Join Mst_Location ML On PM.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Left Join Txn_PledgeCreation_GSL PG On PM.Txn_PledgeCreation_ID = PG.Txn_PledgeCreation_ID And PM.SM_Prefix = PG.SM_Prefix"
tmp = tmp & " Left join Mst_GSL MGL on PG.Mst_GSL_ID=MGL.Mst_GSL_ID and PG.SM_Prefix = MGL.SM_Prefix"
tmp = tmp & " Left Join Mst_Commodity MCom On WD.CommodityID = MCom.CommodityID " 'Or MGL.CommodityID = MCom.CommodityID"
tmp = tmp & " Left Join Mst_Commodity MCom1 On MGL.CommodityID = MCom1.CommodityID "
If Gen_DBType = "MDB" Then
   wc = " Where PM.PledgeDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   wc = " Where PM.PledgeDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, CmbPreFix.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbSBankName.Text <> "" Then
   wc = GenWc(wc, CmbSBankName.Text, "Bank.BankName", "S", "N")
End If

If CmbSBranchName.Text <> "" Then
   wc = GenWc(wc, CmbSBranchName.Text, "Branch.BranchName", "S", "N")
End If

If CmbSearchExchange.Text <> "" Then
   wc = GenWc(wc, CmbSearchExchange.Text, "E.ExchangeType", "S", "N")
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, CmbSCommodity.Text, "MCom.Commodity", "S", "N")
End If

If TxtSearchPledge <> "" Then
   If LCase(CmbSearchExchange.Text) = "non ncdex" Then
      Dim strSplitString() As String
      Dim mSWHRNo As Variant
      strSplitString = Split(TxtSearchPledge.Text, "-")
      tmp1 = UBound(strSplitString)
      If tmp1 <> 1 Then
         MsgBox "Invalid WHR No selection"
         TxtSearchPledge.SetFocus
         Exit Sub
      End If
      mSWHRNo = Val(strSplitString(0)) * 1000000 + Val(strSplitString(1))
      wc = GenWc(wc, mSWHRNo, "WD.WHRNo", "", "")
      'Wc = GenWc(Wc, TxtSearchPledge, "WD.WHRNo", "S", "I")
   ElseIf LCase(CmbSearchExchange.Text) = "ncdex" Then
      wc = GenWc(wc, TxtSearchPledge, "PM.ISINID", "S", "L")
   End If
   
End If


If TxtSClientName <> "" Then
   wc = GenWc(wc, TxtSClientName, "MC.ClientName", "S", "L")
End If

tmp = tmp & " " & wc & " "
tmp = tmp & " Group By PM.Txn_PledgeCreation_ID,ML.LocationName,PM.PledgeDate,E.ExchangeType,WD.WHRNo,WD.CMPID,PM.ISINID,Bank.BankName, Branch.BranchName , MC.ClientName, PM.PledgeReqNo, PM.PledgeFees, PM.PledgeReqDate, PM.Remarks,PM.G_UID , MCom.Commodity,MCom1.Commodity, PM.Discount, PM.RecFees,PM.NCMLFinance"
tmp = tmp & " Order By PM.Txn_PledgeCreation_ID,ML.LocationName"

Call TmpTable
Call Grid_Head1
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    MSHFlexGrid1.TextMatrix(iu, 0) = IIf(IsNull(TmpRs!Txn_PledgeCreation_ID), "", TmpRs!Txn_PledgeCreation_ID)
    MSHFlexGrid1.TextMatrix(iu, 1) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
    MSHFlexGrid1.TextMatrix(iu, 2) = Format(TmpRs!PledgeDate, "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(iu, 3) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
    If IsNull(TmpRs!WHRNo) <> True Then
       MSHFlexGrid1.TextMatrix(iu, 4) = TmpRs!CMPID & "-" & ((TmpRs!WHRNo) - TmpRs!CMPID * 1000000)
    Else
       MSHFlexGrid1.TextMatrix(iu, 4) = IIf(IsNull(TmpRs!ISINID), "", TmpRs!ISINID)
    End If
    MSHFlexGrid1.TextMatrix(iu, 5) = IIf(IsNull(TmpRs!BankName), "", TmpRs!BankName)
    MSHFlexGrid1.TextMatrix(iu, 6) = IIf(IsNull(TmpRs!BranchName), "", TmpRs!BranchName)
    MSHFlexGrid1.TextMatrix(iu, 7) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
    MSHFlexGrid1.TextMatrix(iu, 8) = IIf(IsNull(TmpRs!PledgeReqNo), "", TmpRs!PledgeReqNo)
    MSHFlexGrid1.TextMatrix(iu, 9) = IIf(IsNull(TmpRs!PledgeFees), "", TmpRs!PledgeFees)
    MSHFlexGrid1.TextMatrix(iu, 10) = Format(TmpRs!PledgeReqDate, "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(iu, 11) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
    MSHFlexGrid1.TextMatrix(iu, 12) = IIf(IsNull(TmpRs!G_UID), "", TmpRs!G_UID)
    MSHFlexGrid1.TextMatrix(iu, 13) = IIf(IsNull(TmpRs!PledgeNoOfBags), 0, TmpRs!PledgeNoOfBags)
    MSHFlexGrid1.TextMatrix(iu, 14) = IIf(IsNull(TmpRs!PledgeQuantity), "", TmpRs!PledgeQuantity)
    If IsNull(TmpRs!WHRNo) <> True Then
       MSHFlexGrid1.TextMatrix(iu, 15) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
    Else
       MSHFlexGrid1.TextMatrix(iu, 15) = IIf(IsNull(TmpRs!NCDEXCommodity), "", TmpRs!NCDEXCommodity)
    End If
    MSHFlexGrid1.TextMatrix(iu, 16) = IIf(IsNull(TmpRs!Discount), "", TmpRs!Discount)
    MSHFlexGrid1.TextMatrix(iu, 17) = Val(MSHFlexGrid1.TextMatrix(iu, 9)) - Val(MSHFlexGrid1.TextMatrix(iu, 16))
    MSHFlexGrid1.TextMatrix(iu, 18) = IIf(IsNull(TmpRs!RecFees), "", TmpRs!RecFees)
    MSHFlexGrid1.TextMatrix(iu, 19) = Val(MSHFlexGrid1.TextMatrix(iu, 17)) - Val(MSHFlexGrid1.TextMatrix(iu, 18))
    MSHFlexGrid1.TextMatrix(iu, 20) = IIf(IsNull(TmpRs!NCMLFinance), "0", TmpRs!NCMLFinance)
    If Val(MSHFlexGrid1.TextMatrix(iu, 20)) = 1 Then
       MSHFlexGrid1.TextMatrix(iu, 20) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 20) = "No"
    End If
    TmpRs.MoveNext
Next


Label24.Refresh
Label24.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdPrint_Click()

If Val(TxtBalFee) > 0 Then
   MsgBox "Total Receivable pledge fee not received. Can't print Pledge letter!!!"
   Exit Sub
End If

mPrintFlag = "P"

RsTxn_PledgeCreation!Flag = "P"
RsTxn_PledgeCreation.Update

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "_PledgeConfirmation" & Trim(TxtPledgeID) & ".xls")
Call Xls_Print_Rpt

End Sub

Public Sub Xls_Print_Rpt()
'Dim ans As Variant
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Pat = App.Path
 
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_PledgeConfirmation" & Trim(TxtPledgeID) & ".xls")

Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Pledge Confirmation"
oWS.ClearArrows

oWS.Range("a:a").ColumnWidth = 9.3
oWS.Range("b:b").ColumnWidth = 9.3
oWS.Range("c:c").ColumnWidth = 9.3
oWS.Range("d:d").ColumnWidth = 18.29
oWS.Range("e:e").ColumnWidth = 9.3
oWS.Range("f:f").ColumnWidth = 8
oWS.Range("g:g").ColumnWidth = 10.3
oWS.Range("h:h").ColumnWidth = 9.3

mRow = 1

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "National Collateral Management Services Limited"
oWS.Cells(mRow, 1).HorizontalAlignment = -4108
oWS.Range(tmp).Font.Bold = True
mRow = mRow + 1

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Gayatri Towers,954,Appasaheb Marathe Marg,"
oWS.Cells(mRow, 1).HorizontalAlignment = -4108
mRow = mRow + 1

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Prabhadevi, Mumbai-400025"
oWS.Cells(mRow, 1).HorizontalAlignment = -4108
mRow = mRow + 1

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Tel.No-022 40419191, Fax. No-022 40419193"
oWS.Cells(mRow, 1).HorizontalAlignment = -4108
mRow = mRow + 2

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = "Pledge ID"
'oWS.Cells(mRow, 1).Font.Bold = True

tmp = "c" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = ": " & TxtPledgeID
mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = "Pledge Date"
'oWS.Cells(mRow, 1).Font.Bold = True

tmp = "c" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = ": " & Format(PledgeDate, "dd-MMM-yyyy")
mRow = mRow + 3

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "To,"
mRow = mRow + 1

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "The Branch Manager"
mRow = mRow + 1

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = TxtBankName
mRow = mRow + 1

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Branch : " & TxtBranchName
mRow = mRow + 1

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = GetBranchName(RsTxn_PledgeCreation!BranchID, "A")
mRow = mRow + 2

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = "SUB : Pledge Confirmation Letter "
oWS.Cells(mRow, 1).HorizontalAlignment = -4108
oWS.Cells(mRow, 1).Font.Bold = True
oWS.Cells(mRow, 1).Font.Underline = True
mRow = mRow + 1

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Dear Sir,"
mRow = mRow + 1


tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Range(tmp).RowHeight = 30
oWS.Cells(mRow, 1) = "           We confirm that the following Warehouse Receipt(WHR) is pledged in your favour, vide your Pledge Request No  " & TxtPledgeReqNo & "  dated   " & Format(PledgeDate, "dd-MMM-yyyy") & " "
mRow = mRow + 2

'tmp = "a" & mRow & ":i" & mRow
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow, 1) = "  WHR No " & TxtWHRNoNew & " ,  WHRDate " & Format(WHRDate, "dd-MMM-yyyy") & " , commodity " & TxtCommodity & "   ,  " & TxtWHRNoOfBags & " No Of Bags ,   " & TxtWHRQuantity & "  Ton Weight  ,  WHR Valid Date  " & Format(WHRValidDate, "dd-MMM-yyyy") & ""
'oWS.Cells(mRow, 1).RowHeight = 45
'mRow = mRow + 2

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = "Client Name "
'oWS.Cells(mRow, 1).Font.Bold = True

tmp = "c" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = ": " & TxtClientName
mRow = mRow + 1


tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = "WHR No"
'oWS.Cells(mRow, 1).Font.Bold = True

tmp = "c" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = ": " & TxtWHRNoNew
mRow = mRow + 1


tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = "WHR Date"
'oWS.Cells(mRow, 1).Font.Bold = True

tmp = "c" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = ": " & Format(WHRDate, "dd-MMM-yyyy")
mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = "Commodity"
'oWS.Cells(mRow, 1).Font.Bold = True

tmp = "c" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = ": " & TxtCommodity
mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = "No Of Bags"
'oWS.Cells(mRow, 1).Font.Bold = True

tmp = "c" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = ": " & TxtWHRNoOfBags
mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1) = "Quantity"
'oWS.Cells(mRow, 1).Font.Bold = True

tmp = "c" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = ": " & TxtWHRQuantity & " MT"
mRow = mRow + 2

tmp = "a" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "The details of pledged lots are given below ."
mRow = mRow + 2

tmp = "a" & mRow & ":d" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, 1) = "Godown Address"
mRow1 = mRow + 1

If MSHFlexGrid2.Rows < 5 Then
    oWS.Cells(mRow1, 1).RowHeight = 30
End If

For I = 1 To MSHFlexGrid2.Rows - 2
    If I <> 1 Then
        
        tmp = "a" & mRow1 & ":d" & mRow
        oWS.Range(tmp).Merge
        oWS.Cells(mRow1, 1) = mGodownAddress
        oWS.Cells(mRow1, 1).VerticalAlignment = xlTop
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    End If
    'If I <> 1 Then
    '    tmp = "a" & mRow & ":d" & mRow
    '    oWS.Range(tmp).Merge
    '    oWS.Cells(mRow1, 1) = mGodownAddress
    '    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    '    oWS.Cells(mRow1, 1).WrapText = True
    'End If
    
    oWS.Cells(mRow, 5) = MSHFlexGrid2.TextMatrix(I, 3)
    oWS.Cells(mRow, 5).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 5).WrapText = True
    
    oWS.Cells(mRow, 6) = MSHFlexGrid2.TextMatrix(I, 4)
    oWS.Cells(mRow, 6).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 6).WrapText = True
    
    oWS.Cells(mRow, 7) = MSHFlexGrid2.TextMatrix(I, 7)
    oWS.Cells(mRow, 7).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 7).WrapText = True
    
    oWS.Cells(mRow, 8) = MSHFlexGrid2.TextMatrix(I, 8)
    oWS.Cells(mRow, 8).Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, 8).WrapText = True
    mRow = mRow + 1
Next

tmp = "a" & mRow & ":f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Total"
oWS.Cells(mRow, 1).HorizontalAlignment = xlRight
oWS.Cells(mRow, 1).Font.Bold = True
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, 7) = MSHFlexGrid2.TextMatrix(0, 7)
oWS.Cells(mRow, 7).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, 8) = MSHFlexGrid2.TextMatrix(0, 8)
oWS.Cells(mRow, 8).Borders.Color = RGB(0, 0, 0)
mRow = mRow + 5

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Yours faithfully"
mRow = mRow + 5

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "(CMP Manager)"
mRow = mRow + 1

If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If

oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdUpdate_Click()

Dim strSplitString() As String

If TxtPledgeNoOfBags = "" Then
   MsgBox "Please enter the number of pledged bags!!!"
   TxtPledgeNoOfBags.SetFocus
   Exit Sub
End If
If TxtPledgeQuantity = "" Then
   MsgBox "Please enter Pledged quantity!!!"
   TxtPledgeQuantity.SetFocus
   Exit Sub
End If


If Val(TxtUnpledgeNoOfBags) = TxtPledgeNoOfBags Then
   If Val(TxtUnpledgeWeight) <> Val(TxtPledgeQuantity) Then
      MsgBox "When all the bags are pledged, all the quantity are also required to be pledged!!! "
      TxtPledgeQuantity.SetFocus
      Exit Sub
   End If
End If

If Val(TxtUnpledgeWeight) = TxtPledgeQuantity Then
   If Val(TxtUnpledgeNoOfBags) <> Val(TxtPledgeNoOfBags) Then
      MsgBox "When all the quantity are pledged, all the  bags are also required to be pledged!!! "
      TxtPledgeQuantity.SetFocus
      Exit Sub
   End If
End If

If TxtUnpledgeNoOfBags = 0 And TxtPledgeNoOfBags = 0 Then

Else
    If Val(TxtUnpledgeNoOfBags) <> TxtPledgeNoOfBags Then
       tmp = Validate_ComBags(mCommodityID, Val(TxtPledgeNoOfBags), Val(TxtPledgeQuantity))
       If Len(tmp) > 1 Then
          MsgBox tmp
          TxtPledgeQuantity.SetFocus
          Exit Sub
       End If
    End If
End If

MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))

If Val(TxtPledgeNoOfBags) = 0 And Val(TxtPledgeQuantity) = 0 Then
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
Else
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
End If

MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = TxtCMP
mCMPID = GetCMPID(TxtCMP)
strSplitString = Split(TxtGSL, "~")
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) = Trim(strSplitString(0))
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3) = Trim(strSplitString(1))
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4) = Trim(strSplitString(2))
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5) = Val(TxtUnpledgeNoOfBags) - Val(TxtPledgeNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6) = Format(Val(TxtUnpledgeWeight) - Val(TxtPledgeQuantity), "########0.000")
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7) = Val(TxtPledgeNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8) = Val(TxtPledgeQuantity)


MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))

TxtCMP = ""
TxtGSL = ""
TxtUnpledgeNoOfBags = ""
TxtUnpledgeWeight = ""
TxtPledgeNoOfBags = ""
TxtPledgeQuantity = ""
TxtPledgeNoOfBags.Locked = True
TxtPledgeQuantity.Locked = True
CmdUpdate.Enabled = False

End Sub

Private Sub DeleteCmd_Click()
Dim strSplitString() As String
Dim ans As Variant
On Error GoTo msgError


If Val(TxtRecFee.Text) > 0 Then
   MsgBox "Pledge fee received. Can't delete!!!"
   Exit Sub
End If

If mPrintFlag = "P" Then
   MsgBox "Pledge Confirmation letter printed. Can't delete!!!"
   Exit Sub
End If

If mCARNo <> "" And mCARNo <> 0 Then
   MsgBox "SR generated against this WHR. Can't delete!!!"
   Exit Sub
End If

If mExchangeTypeID = 1 Then
   tmp = " Select * from Txn_SCMDO Where WHRNo = " & TxtWHRNo & " "
   Call TmpTable
    If TmpRs.RecordCount > 0 Then
       MsgBox "DO generated against this WHR. Can't delete!!!"
       Exit Sub
    End If
End If




Call TableTxn_PledgeRelease("Where Txn_PledgeCreation_ID = " & TxtPledgeID & " And Sm_prefix = '" & TxtPreFix & "'")

If RsTxn_PledgeRelease.RecordCount > 0 Then
   MsgBox "Released Pledge. Can't delete!!!"
   Exit Sub
End If


ans = MsgBox("Do you really want to DELETE this record?", vbYesNo)

If ans = vbYes Then
   
   'tmp = "Select * from Mst_GSL where sm_prefix = '" & TxtPreFix & "' And GSLID = " & MSHFlexGrid2.TextMatrix(2, 9) & " "
   RsMst_GSL.MoveFirst
   For x = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(x, 1) = "" Then Exit For
       RsMst_GSL.Filter = " sm_prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & MSHFlexGrid2.TextMatrix(x, 9) & " And ExchangeTypeID = " & mExchangeTypeID & " "
       If RsMst_GSL.RecordCount > 0 Then
          strSplitString = Split(RsMst_GSL!PledgeNo, "~")
          For I = 0 To UBound(strSplitString)
              If TxtPledgeID = strSplitString(I) Then
                 If I = 0 Then
                    RsMst_GSL!PledgeNo = ""
                 End If
              Else
                 If I = 0 Then
                    RsMst_GSL!PledgeNo = strSplitString(I)
                 Else
                    If Trim(RsMst_GSL!PledgeNo) = "" Then
                       RsMst_GSL!PledgeNo = strSplitString(I)
                    Else
                       RsMst_GSL!PledgeNo = Trim(RsMst_GSL!PledgeNo) & "~" & strSplitString(I)
                    End If
                 End If
              End If
          Next
          If RsMst_GSL!PledgeBags > 0 Then
             RsMst_GSL!PledgeBags = Val(RsMst_GSL!PledgeBags) - Val(MSHFlexGrid2.TextMatrix(x, 7))
             'RsMst_GSL!PledgeWeight = Val(IIf(IsNull(RsMst_GSL!PledgeWeight), "0", RsMst_GSL!PledgeWeight)) + Val(MSHFlexGrid2.TextMatrix(i, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
             'RsMst_GSL.Update
          End If
          If RsMst_GSL!PledgeWeight > 0 Then
             RsMst_GSL!PledgeWeight = Val(RsMst_GSL!PledgeWeight) - Val(MSHFlexGrid2.TextMatrix(x, 8))
          End If
          RsMst_GSL.Update
          RsMst_GSL.Filter = " Mst_GSL_ID <> Null"
       Else
          Exit Sub
       End If
   Next
   If mExchangeTypeID = 1 Then
      RsTxn_WHRDetail.MoveFirst
      RsTxn_WHRDetail.Find "WHRNo = " & mWHRNo
      If RsTxn_WHRDetail.EOF = False Then
         RsTxn_WHRDetail!Txn_PledgeCreation_ID = Null
         RsTxn_WHRDetail!Flag = "N"
         RsTxn_WHRDetail.Update
      End If
   End If
   
   tmp = "Delete from Txn_PledgeCreation_GSL where Txn_PledgeCreation_ID = " & TxtPledgeID & " And Sm_prefix = '" & TxtPreFix & "'"
   Call TmpTable
   
   RsTxn_PledgeCreation.Delete
   RsTxn_PledgeCreation.Update
   
   
   tmp = "Delete from Txn_PledgeCreation_GSL where Txn_PledgeCreation_ID = " & TxtPledgeID & " And Sm_prefix = '" & TxtPreFix & "'"
   Call TmpTableAzureStoreMan
   
   tmp = "Delete from Txn_PledgeCreation where Txn_PledgeCreation_ID = " & TxtPledgeID & " And Sm_prefix = '" & TxtPreFix & "'"
   Call TmpTableAzureStoreMan
   
   
   'RsTxn_PledgeCreation_GSL.MoveFirst
   If RsTxn_PledgeCreation.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_PledgeCreation.MoveNext
   If RsTxn_PledgeCreation.EOF() Then
      RsTxn_PledgeCreation.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record present!!!"
RsTxn_PledgeCreation.CancelUpdate
End Sub

Private Sub EditCmd_Click()

If Val(mSCMFlag) = 1 Then
   MsgBox "This Pledge is SCM Pledge. Can't Edit!!!"
   Exit Sub
End If

If mExchangeTypeID = 2 Then
   MsgBox "This Pledge is NCDEX Pledge. Can't Edit!!!"
   Exit Sub
End If

Edit_Flg = "E"
Call GButtonLock(Me, False)
Picture2.Enabled = False
Frame3.Enabled = True
Frame1.Enabled = False
'''CmbBankName.Visible = True
'''CmbBranchName.Visible = True
'''CmdBankName.Visible = True
'''CmdBranchName.Visible = True
'''TxtBankName.Visible = False
'''TxtBranchName.Visible = False
'''CmdPrint.Enabled = False
'''CmbBankName.Text = TxtBankName
'''CmbBranchName.Text = TxtBranchName
'''TxtPledgeReqNo.Locked = False
'''PledgeRequestDate.Enabled = True
'''TxtDiscount.Locked = False
'''mPledgeFlag = "N"

CmbBankName.Visible = False
CmbBranchName.Visible = False
TxtBankName.Visible = True
TxtBranchName.Visible = True
TxtPledgeReqNo.Locked = True
PledgeRequestDate.Enabled = False
TxtDiscount.Locked = False
mPledgeFlag = "Y"


''If mCARNo <> "" And mCARNo <> 0 Then
''   'MsgBox "SR generated against this WHR. Can't Edit!!!"
''   'Exit Sub
''   CmbBankName.Visible = False
''   CmbBranchName.Visible = False
''   TxtBankName.Visible = True
''   TxtBranchName.Visible = True
''   TxtPledgeReqNo.Locked = True
''   PledgeRequestDate.Enabled = False
''   mPledgeFlag = "Y"
''End If
''
''If Val(MSHFlexGrid2.TextMatrix(0, 7)) = Val(TxtWHRNoOfBags) And Val(MSHFlexGrid2.TextMatrix(0, 8)) = Val(TxtWHRQuantity) Then
''
''Else
''''   MsgBox "Partial or full release of pledge done . Can't edit!!!"
''''   Exit Sub
''    CmbBankName.Visible = False
''    CmbBranchName.Visible = False
''    TxtBankName.Visible = True
''    TxtBranchName.Visible = True
''    TxtPledgeReqNo.Locked = True
''    PledgeRequestDate.Enabled = False
''    mPledgeFlag = "Y"
''End If
   
End Sub

Private Sub ExitCmd_Click()
Gen_WHRDetail(0) = True
Unload Me
End Sub

Public Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 10
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "  "
MSHFlexGrid2.TextMatrix(1, 1) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 3) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 4) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 5) = "Unpledged  No Of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Unpledged  Qty"
MSHFlexGrid2.TextMatrix(1, 7) = "Pledged No Of Bags"
MSHFlexGrid2.TextMatrix(1, 8) = "Pledged Qty"
MSHFlexGrid2.TextMatrix(1, 9) = "PledgeGSLID"


MSHFlexGrid2.ColWidth(0) = 400
MSHFlexGrid2.ColWidth(1) = 3200
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 1000
MSHFlexGrid2.ColWidth(5) = 1800
MSHFlexGrid2.ColWidth(6) = 1800
MSHFlexGrid2.ColWidth(7) = 1600
MSHFlexGrid2.ColWidth(8) = 1600
MSHFlexGrid2.ColWidth(9) = 0

MSHFlexGrid2.RowHeight(1) = 400
End Sub
Public Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 21
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "Pledge ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Pledge Date"
MSHFlexGrid1.TextMatrix(0, 3) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 4) = "WHR No/CMSE Lot No" '"WHR No/ISINID"
MSHFlexGrid1.TextMatrix(0, 5) = "Bank Name"
MSHFlexGrid1.TextMatrix(0, 6) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 7) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 8) = "Pledge Request No"
MSHFlexGrid1.TextMatrix(0, 9) = "Pledge Fees"
MSHFlexGrid1.TextMatrix(0, 10) = "Pledge Request Date"
MSHFlexGrid1.TextMatrix(0, 11) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 12) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 13) = "Pledge No Of Bags"
MSHFlexGrid1.TextMatrix(0, 14) = "Pledge Qty"
MSHFlexGrid1.TextMatrix(0, 15) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 16) = "Discount"
MSHFlexGrid1.TextMatrix(0, 17) = "Receivable Fee"
MSHFlexGrid1.TextMatrix(0, 18) = "Received Fee"
MSHFlexGrid1.TextMatrix(0, 19) = "Balance Fee"
MSHFlexGrid1.TextMatrix(0, 20) = "NCML Finance Flag"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1800
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1600
MSHFlexGrid1.ColWidth(12) = 0
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1400
MSHFlexGrid1.ColWidth(15) = 2000
MSHFlexGrid1.ColWidth(16) = 2000
MSHFlexGrid1.RowHeight(0) = 600

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied!!!"
   Exit Sub
End If

Call Grid_Head
'SSTab1.Tab = 0
Call Grid_Head1
Call CmdClear_Click

Call TableTxn_PledgeCreation

If Gen_WHRDetail(0) = False Then
   tmp = ""
   If Gen_WHRDetail(1) <> "" Then
      tmp = tmp & "WHRNo = " & Gen_WHRDetail(1)
   End If
   If Gen_WHRDetail(2) <> "" Then
      tmp = tmp & " And ExchangeTypeID = " & Gen_WHRDetail(2)
   End If
   RsTxn_PledgeCreation.Filter = tmp
''   If RsTxn_PledgeCreation.RecordCount > 0 Then
''      tmp = RsTxn_PledgeCreation!PledgeID
''      RsTxn_PledgeCreation.Filter = ""
''      RsTxn_PledgeCreation.MoveFirst
''      RsTxn_PledgeCreation.Find "PledgeID = " & tmp
''   Else
''      RsTxn_PledgeCreation.Filter = ""
''      RsTxn_PledgeCreation.MoveFirst
''   End If
End If

If RsTxn_PledgeCreation.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call Indata

End Sub

Private Sub ClearData()
TxtPledgeID = ""
TxtWHRNo = ""
TxtGodownNo = ""
TxtWHRNoOfBags = ""
TxtWHRQuantity = ""
TxtClientID = ""
TxtClientName = ""
TxtCommodity = ""
TxtLicenseNo = ""
CmbBankName.Text = ""
CmbBranchName.Text = ""
TxtPledgeReqNo = ""
TxtPledgeFee = ""
TxtRemarks = ""
TxtDPID = ""
TxtDPName = ""

Call Grid_Head
End Sub

Private Sub GetGridData(mExchangeTypeID_, mCommodityID_, mClientIDRow_, mCMPID_)
Call Grid_Head
If mExchangeTypeID_ = "" Or mExchangeTypeID_ = 0 Then Exit Sub
If mCommodityID_ = "" Or mCommodityID_ = 0 Then Exit Sub

If Label1.Caption = "WHR No" Then
   If mClientIDRow_ = "" Or mClientIDRow_ = 0 Then Exit Sub
   If mGodownID = "" Or mGodownID = 0 Then Exit Sub
End If
If mCMPID_ = "" Or mCMPID_ = 0 Then Exit Sub

If CmbExchange.Text = "Non NCDEX" Then
'  Changes done on 20.04.2010 : Remove GodownID filter : show all balance Irrespective  of Godwon ID
'
'   tmp = "SELECT Mst_GSL.GSLID, Mst_GSL.StackNo, Mst_GSL.LotNo, (Mst_GSL.BalanceBags-Mst_GSL.PledgeBags) as UnpledgeBags, (Mst_GSL.BalanceWeight-Mst_GSL.PledgeWeight) as UnpledgeWeight, " & _
'         "Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight, Mst_GSL.PledgeNo, Mst_CMP.CMPName, Mst_Godown.GodownNo " & _
'         "FROM Mst_Godown INNER JOIN (Mst_CMP INNER JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) ON Mst_Godown.GodownID = Mst_GSL.GodownID " & _
'         "WHERE SM_Prefix = '" & TxtPreFix & "' And ( " & Gen_ExpireGSLFlag & "  ) And (((Mst_GSL.ExchangeTypeID)= " & mExchangeTypeID_ & ")AND ((Mst_GSL.CMPID)=" & mCMPID_ & ") AND  " & _
'         "((Mst_GSL.CommodityID)=" & mCommodityID_ & ") And Mst_GSL.GodownID = " & mGodownID & _
'         "AND ((Mst_GSL.DepositeWeight - Mst_GSL.PledgeWeight) > 0)) AND ((Mst_GSL.ClientIDRow)=" & mClientIDRow_ & ")"
  
      If mLotWiseWHR = "0" Then
         tmp = "SELECT Mst_GSL.Mst_GSL_ID, Mst_GSL.StackNo, Mst_GSL.LotNo, (Mst_GSL.BalanceBags-Mst_GSL.PledgeBags) as UnpledgeBags, (Mst_GSL.BalanceWeight-Mst_GSL.PledgeWeight) as UnpledgeWeight, " & _
               "Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight, Mst_GSL.PledgeNo, Mst_CMP.CMPName, Mst_Godown.GodownNo " & _
               "FROM Mst_Godown INNER JOIN (Mst_CMP INNER JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) ON Mst_Godown.GodownID = Mst_GSL.GodownID " & _
               "WHERE SM_Prefix = '" & TxtPreFix & "' And ( " & Gen_ExpireGSLFlag & "  ) And (((Mst_GSL.ExchangeTypeID)= " & mExchangeTypeID_ & ")AND ((Mst_GSL.CMPID)=" & mCMPID_ & ") AND  " & _
               "((Mst_GSL.CommodityID)=" & mCommodityID_ & ")  " & _
               "AND ((Mst_GSL.DepositeWeight - Mst_GSL.PledgeWeight) > 0)) AND ((Mst_GSL.ClientID)=" & mClientIDRow_ & ") AND ((Mst_GSL.GodownID)=" & mGodownID & ") And Mst_GSL.MadeupFlag = '0' And Mst_GSL.WHRNo is null "
      Else
         tmp = "SELECT Mst_GSL.Mst_GSL_ID, Mst_GSL.StackNo, Mst_GSL.LotNo, (Mst_GSL.BalanceBags-Mst_GSL.PledgeBags) as UnpledgeBags, (Mst_GSL.BalanceWeight-Mst_GSL.PledgeWeight) as UnpledgeWeight, " & _
               "Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight, Mst_GSL.PledgeNo, Mst_CMP.CMPName, Mst_Godown.GodownNo " & _
               "FROM Mst_Godown INNER JOIN (Mst_CMP INNER JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) ON Mst_Godown.GodownID = Mst_GSL.GodownID " & _
               "WHERE SM_Prefix = '" & TxtPreFix & "' And ( " & Gen_ExpireGSLFlag & "  ) And (((Mst_GSL.ExchangeTypeID)= " & mExchangeTypeID_ & ")AND ((Mst_GSL.CMPID)=" & mCMPID_ & ") AND  " & _
               "((Mst_GSL.CommodityID)=" & mCommodityID_ & ")  " & _
               "AND ((Mst_GSL.DepositeWeight - Mst_GSL.PledgeWeight) > 0)) AND ((Mst_GSL.ClientID)=" & mClientIDRow_ & ")  And Mst_GSL.MadeupFlag = '0' And Mst_GSL.WHRNo = " & TxtWHRNo & " " 'AND ((Mst_GSL.GodownID)=" & mGodownID & ")
      End If
Else
   
   tmp = "SELECT Mst_GSL.Mst_GSL_ID, Mst_GSL.StackNo, Mst_GSL.LotNo, (Mst_GSL.BalanceBags-Mst_GSL.PledgeBags) as UnpledgeBags, (Mst_GSL.BalanceWeight-Mst_GSL.PledgeWeight) as UnpledgeWeight, " & _
         "Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight, Mst_GSL.PledgeNo, Mst_CMP.CMPName, Mst_Godown.GodownNo " & _
         "FROM Mst_Godown INNER JOIN (Mst_CMP INNER JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) ON Mst_Godown.GodownID = Mst_GSL.GodownID " & _
         "WHERE SM_Prefix = '" & TxtPreFix & "' And ( " & Gen_ExpireGSLFlag & "  ) And (((Mst_GSL.ExchangeTypeID)= " & mExchangeTypeID_ & ")AND ((Mst_GSL.CMPID)=" & mCMPID_ & ") AND  " & _
         "((Mst_GSL.CommodityID)=" & mCommodityID_ & ") AND ((Mst_GSL.ISINID) = '" & TxtWHRNo.Text & "')) AND ((Mst_GSL.DepositeWeight - Mst_GSL.PledgeWeight) > 0) And Mst_GSL.MadeupFlag = '0'"
End If

Call TmpTable

If TmpRs.RecordCount = 0 Then
   MsgBox "No Lots available for Pledge!!!"
   Exit Sub
End If

k = 1
For I = 0 To TmpRs.RecordCount - 1
    k = MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    MSHFlexGrid2.TextMatrix(k, 0) = ""
    MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!CMPName
    MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!GodownNo
    MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!StackNo
    MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!LotNo
    MSHFlexGrid2.TextMatrix(k, 5) = TmpRs!UnPledgeBags
    MSHFlexGrid2.TextMatrix(k, 6) = TmpRs!UnPledgeWeight
    If mLotWiseWHR = "1" Then
       MSHFlexGrid2.TextMatrix(k, 7) = Val(MSHFlexGrid2.TextMatrix(k, 5))
       MSHFlexGrid2.TextMatrix(k, 8) = Val(MSHFlexGrid2.TextMatrix(k, 6))
       MSHFlexGrid2.TextMatrix(k, 0) = "*"
       MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(k, 7))
       MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(k, 8))
    End If
    MSHFlexGrid2.TextMatrix(k, 9) = TmpRs!Mst_GSL_ID
    TmpRs.MoveNext
Next

End Sub




Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
If Edit_Flg = "E" Then Exit Sub
TxtCMP = ""
TxtGSL = ""
TxtUnpledgeNoOfBags = ""
TxtUnpledgeWeight = ""
TxtPledgeQuantity = ""
TxtPledgeNoOfBags = ""
CmdUpdate.Enabled = False
If mLotWiseWHR = "1" Then
   MsgBox "Partial Pledge on Lot mapped during WHR creation not allowed . If you want to generate Pledge for whole WHR click Save Button, else delete WHR and split GSL using GSL Transfer screen, generate seprate WHR and further pledge transaction!!!"
   Exit Sub
End If
If MSHFlexGrid2.Rows > 3 Then
   If MSHFlexGrid2.RowSel <> "0" Then
      If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then Exit Sub
      TxtPledgeNoOfBags.Locked = False
      TxtPledgeQuantity.Locked = False
      TxtCMP = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1)
      TxtGSL = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) & "~" & MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3) & "~" & MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4)
      TxtUnpledgeNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
      TxtUnpledgeWeight = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
       
      If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7)) = 0 Then
         TxtPledgeNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5))
      Else
         TxtPledgeNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
      End If
      If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8)) = 0 Then
         TxtPledgeQuantity = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6))
      Else
         TxtPledgeQuantity = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
      End If
       
      CmdUpdate.Enabled = True
   End If
End If

End Sub

Private Sub PledgeDate_LostFocus()
PledgeRequestDate.Value = PledgeDate.Value
End Sub

Private Sub SaveCmd_Click()

If Edit_Flg = "A" Then
   If PledgeDate.Value > Date Then
      MsgBox "Future date not allowed!!!"
      PledgeDate.SetFocus
      Exit Sub
   End If
   
   If CmbExchange.Text = "Non NCDEX" Then
      If PledgeDate.Value < WHRDate.Value Then
         MsgBox "Pledge Date must be greater than WHR Date!!!"
         PledgeDate.SetFocus
         Exit Sub
      End If
   Else
      If TxtClientID = "" Then
         MsgBox "Blank Client ID not allowed!!!"
         TxtClientID.SetFocus
         Exit Sub
      End If
   
   End If
   If PledgeRequestDate > PledgeDate.Value Then
      MsgBox "Pledge Request Date must be less than or equal to Pledge date!!!"
      PledgeRequestDate.SetFocus
      Exit Sub
   End If
   If TxtWHRNo.Text = "" Then
      If Label1.Caption = "WHR No" Then
         MsgBox "Blank WHR Number not allowed!!!"
      Else
         MsgBox "Blank CMSE Lot Number not allowed!!!"
      End If
      TxtWHRNo.SetFocus
      Exit Sub
   End If
    
   If MSHFlexGrid2.Rows < 4 Then
      MsgBox "No Lot available for pledge!!!"
      TxtWHRNo.SetFocus
      Exit Sub
   End If
    
   If ChkSelection() = False Then
      MsgBox "Please select atleast one Lot number!!!"
      MSHFlexGrid2.SetFocus
      Exit Sub
   End If
    
   If Label1.Caption = "WHR No" Then
   '   If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" Then
      If Val(MSHFlexGrid2.TextMatrix(0, 7)) <> Val(TxtWHRNoOfBags) Then
         MsgBox "Number of bags pledged should be same as WHR number of bags!!!"
         MSHFlexGrid2.SetFocus
         Exit Sub
      End If
      If Val(MSHFlexGrid2.TextMatrix(0, 8)) <> Val(TxtWHRQuantity) Then
         MsgBox "Pledged quantity should be same as WHR Quantity!!!"
         MSHFlexGrid2.SetFocus
         Exit Sub
      End If
    '   End If
   End If

End If

If PledgeRequestDate.Value > Date Then
   MsgBox "'Pledge Request Date' should not be a future date!!!"
   PledgeRequestDate.SetFocus
   Exit Sub
End If

If CmbBankName.Text = "" Then
   MsgBox "Blank 'Pledged to Bank' not allowed!!!"
   CmbBankName.SetFocus
   Exit Sub
End If
If CmbBranchName.Text = "" Then
   MsgBox "Blank 'Branch Name' not allowed!!!"
   CmbBranchName.SetFocus
   Exit Sub
End If

If TxtPledgeReqNo = "" Then
   MsgBox "Blank 'Pledge Request No' not allowed!!!"
   TxtPledgeReqNo.SetFocus
   Exit Sub
End If

If Val(TxtBalFee) < 0 Then
   MsgBox "Balance fee must be grater than or equal to zero!!!"
   TxtDiscount.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_PledgeCreation.AddNew
   RsTxn_PledgeCreation!PledgeID = GetPledgeID
   RsTxn_PledgeCreation!G_UID = GetGUID
   TxtG_UID = RsTxn_PledgeCreation!G_UID
   'TxtPledgeID = RsTxn_PledgeCreation!PledgeID

   RsTxn_PledgeCreation!PledgeDate = PledgeDate.Value
   If Label1.Caption = "WHR No" Then
      RsTxn_PledgeCreation!WHRNo = mWHRNo
      RsTxn_PledgeCreation!ISINID = "0"
      RsTxn_PledgeCreation!Txn_WHR_Detail_ID = mWHR_Detail_ID
   Else
      RsTxn_PledgeCreation!WHRNo = 0
      RsTxn_PledgeCreation!ISINID = mISINID
   End If
    
   RsTxn_PledgeCreation!ExchangeTypeID = mExchangeTypeID

   RsTxn_PledgeCreation!Remarks = IIf(IsNull(TxtRemarks), "", TxtRemarks)
   RsTxn_PledgeCreation!ClientID = mClientIDRow
   RsTxn_PledgeCreation!SM_Prefix = TxtPreFix
   
Else
   RsTxn_PledgeCreation.MoveFirst
   RsTxn_PledgeCreation.Find "G_UID = '" & TxtG_UID & "'"
   
End If

RsTxn_PledgeCreation!BankID = mBankID
RsTxn_PledgeCreation!BranchID = mBranchID
RsTxn_PledgeCreation!PledgeReqNo = TxtPledgeReqNo
RsTxn_PledgeCreation!PledgeFees = Val(TxtPledgeFee)
RsTxn_PledgeCreation!PledgeReqDate = PledgeRequestDate.Value
RsTxn_PledgeCreation!Discount = Val(TxtDiscount)
RsTxn_PledgeCreation!SCMFlag = 0
RsTxn_PledgeCreation!NCMLFinance = mNCMLFinance
RsTxn_PledgeCreation!MarketYardFlag = mMarketYard

If IsNull(RsTxn_PledgeCreation!CreatedBy) = True Or RsTxn_PledgeCreation!CreatedBy = "" Then
   RsTxn_PledgeCreation!CreatedBy = Gen_UserLoginID
   RsTxn_PledgeCreation!CreatedDate = Now
Else
   RsTxn_PledgeCreation!UpdatedBy = Gen_UserLoginID
   RsTxn_PledgeCreation!UpdatedDate = Now
End If

RsTxn_PledgeCreation.Update

tmp = "Select * from  Txn_PledgeCreation Where G_UID = '" & TxtG_UID & "'"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   TxtPledgeID = TmpRs1!Txn_PledgeCreation_ID
End If

If Edit_Flg = "A" Then
   If Label1.Caption = "WHR No" Then
      'RsTxn_WHRDetail!PledgeID = TxtPledgeID
      'RsTxn_WHRDetail.Update
      tmp = "Update Txn_WHR_Detail Set Txn_PledgeCreation_ID = " & Val(TxtPledgeID) & " Where WHRNo =  " & mWHRNo & " "
      Call Tmp1Table
      Call SavePledgeCreation_GSL
      If Val(mMarketYard) = 1 Then
         tmp = "Update Txn_WHRForMktYard Set Txn_PledgeCreation_ID = " & Val(TxtPledgeID) & ",SM_Prefix = '" & TxtPreFix & "' Where WHRNo =  " & mWHRNo & " And Isnull(TransferedFlag,'') = '' And Isnull(Txn_PledgeCreation_ID,0) = 0"
         Call Tmp1Table
      End If
   End If
End If

Call Indata
Call Gen_Email

End Sub
Private Sub SavePledgeCreation_GSL()
Call TableTxn_PledgeCreation_GSL("Where SM_Prefix = '" & TxtPreFix & "' And Txn_PledgeCreation_ID = " & TxtPledgeID)
If mLotWiseWHR = "0" Then
   Call TableMst_GSL(" Where SM_Prefix = '" & TxtPreFix & "'")
Else
   Call TableMst_GSL(" Where SM_Prefix = '" & TxtPreFix & "' And WHRNo = " & TxtWHRNo & "")
End If
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(I, 0) = "*" Then
       RsTxn_PledgeCreation_GSL.AddNew
       RsTxn_PledgeCreation_GSL!G_UID = GetGUID
       RsTxn_PledgeCreation_GSL!PledgeNo = GetPledgeNo
       RsTxn_PledgeCreation_GSL!Txn_PledgeCreation_ID = TxtPledgeID
       RsTxn_PledgeCreation_GSL!Mst_GSL_ID = MSHFlexGrid2.TextMatrix(I, 9)
       RsTxn_PledgeCreation_GSL!PledgeNoOfBags = MSHFlexGrid2.TextMatrix(I, 7)
       RsTxn_PledgeCreation_GSL!PledgeQuantity = MSHFlexGrid2.TextMatrix(I, 8)
       RsTxn_PledgeCreation_GSL!SM_Prefix = TxtPreFix
       RsTxn_PledgeCreation_GSL.Update
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "Mst_GSL_ID = " & Val(MSHFlexGrid2.TextMatrix(I, 9))
       If Not RsMst_GSL.EOF Then
          If IsNull(RsMst_GSL!PledgeNo) Or RsMst_GSL!PledgeNo = 0 Or RsMst_GSL!PledgeNo = "" Then
             RsMst_GSL!PledgeNo = TxtPledgeID
          Else
             RsMst_GSL!PledgeNo = Trim(RsMst_GSL!PledgeNo) & "~" & TxtPledgeID
          End If
          RsMst_GSL!PledgeBags = Val(IIf(IsNull(RsMst_GSL!PledgeBags), "0", RsMst_GSL!PledgeBags)) + Val(MSHFlexGrid2.TextMatrix(I, 7))
          RsMst_GSL!PledgeWeight = Val(IIf(IsNull(RsMst_GSL!PledgeWeight), "0", RsMst_GSL!PledgeWeight)) + Val(MSHFlexGrid2.TextMatrix(I, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
          RsMst_GSL.Update
       End If
       RsTxn_PledgeCreation_GSL.Requery
    End If
Next

End Sub
Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""

'mMsg = "<tr>Pledge Details (From Store-Man)</tr>"
mPreMsg = ""
mPostMsg = ""
'tmp = "Select * from Mst_Location Where SM_Prefix = '" & TxtPreFix & "'"


tmp = " Select ML.LocationName,MS.StateName,MS.StateID from Mst_Location ML "
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Where ML.SM_Prefix = '" & TxtPreFix & "'"

Call TmpTable

mLocationName = TmpRs!Locationname
mStateName = TmpRs!StateName
mStateID = TmpRs!StateID



'mPreMsg = "<P>The below mentioned WHR/ISIN is pledged today."
'mPreMsg = mPreMsg & " The details of the pledge are given below. Please arrange to issue SR to the lending Bank at the earliest."

'mPreMsg = mPreMsg & " The details of the pledge are given below. Please arrange to issue SR to the lending Bank at the earliest."
mPreMsg = "<P> Dear DRC , <br> "
mPreMsg = mPreMsg & " The below mentioned WHR/CMSE Lot No is pledged today in mentioned CMP warehouse. The details of the pledge are given below. Kindly generate CAR No for the mentioned WHR to the lending Bank at the earliest."

mPreMsg = mPreMsg & "<br>"

mMsg = mMsg & "<tr><td> Pledge ID </td><td> " & TxtPreFix & " - " & TxtPledgeID & " </td></tr>"

mMsg = mMsg & "<tr><td> CMP ID </td><td> " & mCMPID & " </td></tr>"
mMsg = mMsg & "<tr><td> CMP Name </td><td> " & MSHFlexGrid2.TextMatrix(2, 1) & "  </td></tr>"
mMsg = mMsg & "<tr><td> Godown Address </td><td> " & mGodownAddress & "  </td></tr>"
mMsg = mMsg & "<tr><td> Location Name </td><td> " & mLocationName & "  </td></tr>"
mMsg = mMsg & "<tr><td> State Name </td><td> " & mStateName & "  </td></tr>"


If LCase(CmbExchange.Text) = "non ncdex" Then
   mMsg = mMsg & "<tr><td> WHR No </td><td> " & TxtWHRNoNew & "   </td></tr>"
   mMsg = mMsg & "<tr><td> WHR Date </td><td> " & Format(WHRDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
   mMsg = mMsg & "<tr><td> WHR Valid Date </td><td> " & Format(WHRValidDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
ElseIf LCase(CmbExchange.Text) = "ncdex" Then
   mMsg = mMsg & "<tr><td> CMSE Lot No </td><td> TxtWHRNo </td></tr>"
   mMsg = mMsg & "<tr><td> CMSE Lot Valid Date </td><td> " & Format(WHRValidDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
End If

mMsg = mMsg & "<tr><td> Borrower Name and ID </td><td> " & TxtClientName & " - " & TxtClientID & " </td></tr>"
mMsg = mMsg & "<tr><td> Commodity </td><td> " & TxtCommodity & " </td></tr>"

If LCase(CmbExchange.Text) = "non ncdex" Then
   mMsg = mMsg & "<tr><td> Value of Commodity </td><td> " & Val(mValueOfCommodity) & " </td></tr>"
End If


mMsg = mMsg & "<tr><td> Bank Name </td><td> " & TxtBankName & " </td></tr>"

mMsg = mMsg & "<tr><td> Branch Name </td><td> " & TxtBranchName & " </td></tr>"
mMsg = mMsg & "<tr><td> Pledge Created by  </td><td> " & Gen_UserName & " - " & Gen_UserLoginID & " (" & GetDesignaion(Gen_DesignationID) & ")</td></tr>"

If LCase(CmbExchange.Text) = "non ncdex" Then

   tmp = " Select ME.EmployeeName,ME.EmployeeNo,ME.Flag from Mst_EmpCMPMapping MEC "
   tmp = tmp & " Inner Join Mst_Employee ME On MEC.EMployeeID = ME.EmployeeID"
   tmp = tmp & " Where MEC.CMPID = " & mCMPID & ""
   
   Call Tmp1Table
   
   If TmpRs1.RecordCount > 0 Then
      For I = 1 To TmpRs1.RecordCount
          mMsg = mMsg & "<tr><td> CMP Manager  </td><td> " & TmpRs1!EmployeeName & " ~ " & TmpRs1!EmployeeNo & " ( Status - " & TmpRs1!Flag & ")</td></tr>"
          TmpRs1.MoveNext
      Next
   End If
End If

mGridMsg = ""

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid2.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 7) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 8) & " </td></tr>"
    Else
       mGridMsg = mGridMsg & "<tr><td> &nbsp </td><td> &nbsp </td><td> Total Pledge </td><td> " & MSHFlexGrid2.TextMatrix(0, 7) & " </td><td> " & MSHFlexGrid2.TextMatrix(0, 8) & " </td></tr>"
    End If
Next

mMsg = mMsg & "<tr><td> Details </td><td><table border = 1>" & mGridMsg & "</table></td></tr>"

'mMsg = mMsg & "<br><br><br>"

'mPostMsg = "<br><br><br>"
'mPostMsg = mPostMsg & "<P>Please issue SR to the Bank immediately."



'TOADD = "sr.ccg@ncml.com"
'TOADD = mForwardToEmailID ' mEmployeeEmailID

tmp = "Select * from Mst_StateDRCMapping Where StateID = " & mStateID & " And Flag = 'A' "
Call TmpTable

TOADD = ""
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       If TOADD = "" Then
          TOADD = TmpRs!EmailID
       Else
          TOADD = TOADD & "," & TmpRs!EmailID
       End If
       TmpRs.MoveNext
   Next
End If
'TOCC = "manish.g@ncml.com,jayaprakash.g@ncml.com,state.cmdrc@ncml.com,whr.ccg@ncml.com," 'Gen_UserEmailID '& "," & mForwardToEmailID

TOCC = "sr.ccg@ncml.com"

If Edit_Flg = "A" Then
   mMsgSubject = "New Pledge created in StoreMan - PledgeID is " & TxtPledgeID & " in Location  " & mLocationName & ""
Else
   mMsgSubject = "Pledge Modified in StoreMan - PledgeID is " & TxtPledgeID & " in Location  " & mLocationName & ""
End If

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)


End Sub
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "10.0.0.13"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT> " & mPreMsg & " <table border = 1>"
htmlStringtail = "</table></Font></body></html>"

'htmlStringtail = "</table> " & mPostMsg & " </Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  .MailFrom = Gen_UserEmailID
'  'message
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "amol.pa@ncml.com" '"sr.ccg@ncml.com"
'  End If
'  .SendTo = ToADD_
'
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_ & "," & Gen_UserEmailID
'  End If
'
'  .BCC = "parag.r@ncml.com"
'  .MessageSubject = MsgSubject_ '& " e-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com"  '"mail.ncml.com"
objSMTP.FromAddr = Gen_UserEmailID
objSMTP.EnableLogging = True
objSMTP.LogFilePath = "C:\temp\StoreMan\PledgeRelease.txt"
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_ & "," & Gen_UserEmailID
End If

objSMTP.Subject = MsgSubject_ '& " (from HR-Pro)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function

Private Sub Indata()

mExchangeTypeID = RsTxn_PledgeCreation!ExchangeTypeID
CmbExchange.Text = GetExchangeName(mExchangeTypeID)
CmdPrint.Enabled = True
TxtWHRNoNew.Locked = True
If CmbExchange.Text = "NCDEX" Then
   Call EnDn("NCDEX", False)
   mClientIDRow = 0
   TxtClientID.BackColor = -2147483643
   TxtClientID.Enabled = True
   TxtClientID.Locked = True
   lblDpIDName.Visible = True
   TxtDPID.Visible = True
   TxtDPName.Visible = True
Else
   Call EnDn("Non NCDEX", True)
   WHRDate.Enabled = False
   mClientIDRow = 0
   TxtClientID.BackColor = -2147483648#
   TxtClientID.Enabled = False
   TxtClientID.Locked = True
   lblDpIDName.Visible = False
   TxtDPID.Visible = False
   TxtDPName.Visible = False
End If

TxtPledgeID = IIf(IsNull(RsTxn_PledgeCreation!Txn_PledgeCreation_ID), "0", RsTxn_PledgeCreation!Txn_PledgeCreation_ID)
PledgeDate = RsTxn_PledgeCreation!PledgeDate
TxtPreFix = RsTxn_PledgeCreation!SM_Prefix
mSCMFlag = IIf(IsNull(RsTxn_PledgeCreation!SCMFlag), 0, RsTxn_PledgeCreation!SCMFlag)
mNCMLFinance = IIf(IsNull(RsTxn_PledgeCreation!NCMLFinance), 0, RsTxn_PledgeCreation!NCMLFinance)
'If LCase(TxtPreFix) = "bknrp" Then
'   mSM_Prefix = "BKNER"
'End If
If LCase(TxtPreFix) = "bknrp" Then
   mSM_Prefix = "BKNER"
ElseIf LCase(TxtPreFix) = "ujjap" Then
   mSM_Prefix = "Ujjan"
ElseIf LCase(TxtPreFix) = "dewsp" Then
   mSM_Prefix = "DEWAS"
ElseIf LCase(TxtPreFix) = "vdshp" Then
   mSM_Prefix = "VDISA"
ElseIf LCase(TxtPreFix) = "adlbp" Then
   mSM_Prefix = "ADLBD"
ElseIf LCase(TxtPreFix) = "nzmdp" Then
   mSM_Prefix = "NZMBD"
ElseIf LCase(TxtPreFix) = "unjap" Then
   mSM_Prefix = "unjha"
ElseIf LCase(TxtPreFix) = "gntrp" Then
   mSM_Prefix = "guntr"
ElseIf LCase(TxtPreFix) = "dvgrp" Then
   mSM_Prefix = "dvngr"
ElseIf LCase(TxtPreFix) = "jdhrp" Then
   mSM_Prefix = "jdhpr"
ElseIf LCase(TxtPreFix) = "hngtp" Then
   mSM_Prefix = "hghat"
ElseIf LCase(TxtPreFix) = "kadip" Then
   mSM_Prefix = "kadi"
ElseIf LCase(TxtPreFix) = "lturp" Then
   mSM_Prefix = "latur"
ElseIf LCase(TxtPreFix) = "ktap" Then
   mSM_Prefix = "kota"
ElseIf LCase(TxtPreFix) = "rsgrp" Then
   mSM_Prefix = "rsngr"
ElseIf LCase(TxtPreFix) = "rsngr" Then
   mSM_Prefix = "rsgrp"
End If

If Label1.Caption = "WHR No" Then
   TxtWHRNo = RsTxn_PledgeCreation!WHRNo
   mWHRNo = TxtWHRNo
   If LCase(TxtPreFix) = "bknrp" Or LCase(TxtPreFix) = "ujjap" Or LCase(TxtPreFix) = "dewsp" Or LCase(TxtPreFix) = "vdshp" Or LCase(TxtPreFix) = "adlbp" Or LCase(TxtPreFix) = "nzmdp" Or LCase(TxtPreFix) = "unjap" Or LCase(TxtPreFix) = "gntrp" Or LCase(TxtPreFix) = "dvgrp" Or LCase(TxtPreFix) = "jdhrp" Or LCase(TxtPreFix) = "hngtp" Or LCase(TxtPreFix) = "kadip" Or LCase(TxtPreFix) = "lturp" Or LCase(TxtPreFix) = "ktap" Or LCase(TxtPreFix) = "rsgrp" Or LCase(TxtPreFix) = "rsngr" Then
      Call TableTxn_WHRDetail("Where WHRNo = " & TxtWHRNo & " And (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mSM_Prefix & "') ")
   Else
      Call TableTxn_WHRDetail("Where WHRNo = " & TxtWHRNo & " And sm_prefix = '" & TxtPreFix & "'")
   End If
   WHRDate.Value = RsTxn_WHRDetail!WHRDate
   mClientIDRow = RsTxn_WHRDetail!ClientID
   TxtWHRNoNew = RsTxn_WHRDetail!CMPID & "-" & (Val(TxtWHRNo) - (RsTxn_WHRDetail!CMPID * 1000000))
   ans = GetClientNameByRow(mClientIDRow)
   If ans(0) = "" Then
      MsgBox "Client ID not found!!!"
      Exit Sub
   End If
   TxtClientID.Text = ans(0)
   TxtClientName = ans(1)
   mCommodityID = IIf(IsNull(RsTxn_WHRDetail!CommodityID), "", RsTxn_WHRDetail!CommodityID)
   mLotWiseWHR = IIf(IsNull(RsTxn_WHRDetail!LotWiseWHR), 0, RsTxn_WHRDetail!LotWiseWHR)
   TxtLicenseNo = GetLicenseNo(RsTxn_WHRDetail!LicenseID)
   TxtWHRNoOfBags = IIf(IsNull(RsTxn_WHRDetail!WHRNoOfBags), "", RsTxn_WHRDetail!WHRNoOfBags)
   TxtWHRQuantity = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "", RsTxn_WHRDetail!WHRQuantity)
   WHRValidDate.Value = RsTxn_WHRDetail!StorageTo
   TxtGodownNo = GetGodownName(IIf(IsNull(RsTxn_WHRDetail!GodownID), "", RsTxn_WHRDetail!GodownID), "N")
   mCMPID = IIf(IsNull(RsTxn_WHRDetail!CMPID), "", RsTxn_WHRDetail!CMPID)
   mGodownID = IIf(IsNull(RsTxn_WHRDetail!GodownID), "", RsTxn_WHRDetail!GodownID)
   If mGodownID = 2283 Then
      TxtPreFix = "BKNRP"
      mCMPID = 678
      mGodownID = 2316
   ElseIf mGodownID = 25 Then
      TxtPreFix = "ujjap"
      mCMPID = 705
      mGodownID = 2378
   End If
   
   mGodownAddress = GetGodownName(IIf(IsNull(RsTxn_WHRDetail!GodownID), "", RsTxn_WHRDetail!GodownID), "A")
   mRate = IIf(IsNull(RsTxn_WHRDetail!Rate), "", RsTxn_WHRDetail!Rate)
   mValueOfCommodity = Val(TxtWHRQuantity) * Val(mRate)
   mCARNo = IIf(IsNull(RsTxn_WHRDetail!CARNo), "", RsTxn_WHRDetail!CARNo)
   If mCARNo = "" Or mCARNo = 0 Then
      CmdPrint.Enabled = False
   End If
   If mNCMLFinance = 1 Then
      CmdPrint.Enabled = True
   End If
Else
   mClientIDRow = RsTxn_PledgeCreation!ClientID
   ans = GetClientNameByRow(mClientIDRow)
   If ans(0) = "" Then
      'MsgBox "Client ID Not Found!!!"
      Exit Sub
   End If
   TxtClientID.Text = ans(0)
   TxtClientName = ans(1)
   TxtDPID.Text = ans(2)
   TxtDPName = ans(3)
   
   TxtWHRNoNew = RsTxn_PledgeCreation!ISINID
   TxtWHRNo = RsTxn_PledgeCreation!ISINID
   
   If LCase(TxtPreFix) = "bknrp" Or LCase(TxtPreFix) = "ujjap" Or LCase(TxtPreFix) = "dewsp" Or LCase(TxtPreFix) = "vdshp" Or LCase(TxtPreFix) = "adlbp" Or LCase(TxtPreFix) = "nzmdp" Or LCase(TxtPreFix) = "unjap" Or LCase(TxtPreFix) = "gntrp" Or LCase(TxtPreFix) = "dvgrp" Or LCase(TxtPreFix) = "jdhrp" Or LCase(TxtPreFix) = "hngtp" Or LCase(TxtPreFix) = "kadip" Or LCase(TxtPreFix) = "lturp" Or LCase(TxtPreFix) = "ktap" Or LCase(TxtPreFix) = "rsgrp" Or LCase(TxtPreFix) = "rsngr" Then
      'Call TableTxn_WHRDetail("Where WHRNo = " & TxtWHRNo & " And (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mSM_Prefix & "') ")
      tmp = " Select * from Txn_ImportNcdexLotPleadge Where Txn_PledgeCreation_ID=" & Val(TxtPledgeID.Text) & " And (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mSM_Prefix & "')"
   Else
      'Call TableTxn_WHRDetail("Where WHRNo = " & TxtWHRNo & " And sm_prefix = '" & TxtPreFix & "'")
      tmp = " Select * from Txn_ImportNcdexLotPleadge Where Txn_PledgeCreation_ID=" & Val(TxtPledgeID.Text) & " And (sm_prefix = '" & TxtPreFix & "')"
   End If
   
   Call Tmp2Table
   If TmpRs2.RecordCount > 0 Then
      TxtCarNo.Text = IIf(IsNull(TmpRs2!CARNo), "", TmpRs2!CARNo)
      mISINID = TmpRs2!CMSELotId
      Call TableMst_ISIN("Where ISINID = '" & mISINID & "'")
      mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
      mCMPID = IIf(IsNull(RsMst_ISIN!CMPID), "", RsMst_ISIN!CMPID)
      TxtLicenseNo = GetCMPName(mCMPID)
      WHRValidDate.Value = RsMst_ISIN!FED
      mGodownAddress = ""
   End If
End If

TxtCommodity = GetCommodityName(mCommodityID)
TxtRemarks = IIf(IsNull(RsTxn_PledgeCreation!Remarks), "", RsTxn_PledgeCreation!Remarks)
mBankID = RsTxn_PledgeCreation!BankID
mBranchID = RsTxn_PledgeCreation!BranchID
TxtBankName = GetBankName(RsTxn_PledgeCreation!BankID, "N")
TxtBranchName = GetBranchName(RsTxn_PledgeCreation!BranchID, "N")
TxtPledgeReqNo = IIf(IsNull(RsTxn_PledgeCreation!PledgeReqNo), "", RsTxn_PledgeCreation!PledgeReqNo)
TxtPledgeFee = IIf(IsNull(RsTxn_PledgeCreation!PledgeFees), "", RsTxn_PledgeCreation!PledgeFees)
PledgeRequestDate.Value = RsTxn_PledgeCreation!PledgeReqDate
TxtCreated = IIf(IsNull(RsTxn_PledgeCreation!CreatedBy), "", RsTxn_PledgeCreation!CreatedBy) & " :- " & IIf(IsNull(RsTxn_PledgeCreation!CreatedDate), "", RsTxn_PledgeCreation!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_PledgeCreation!UpdatedBy), "", RsTxn_PledgeCreation!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_PledgeCreation!UpdatedDate), "", RsTxn_PledgeCreation!UpdatedDate)
TxtDiscount = IIf(IsNull(RsTxn_PledgeCreation!Discount), "", RsTxn_PledgeCreation!Discount)
TxtReceivableFee = Val(TxtPledgeFee) - Val(TxtDiscount)
TxtRecFee = IIf(IsNull(RsTxn_PledgeCreation!RecFees), 0, RsTxn_PledgeCreation!RecFees)
TxtBalFee = Val(TxtReceivableFee) - Val(TxtRecFee)
TxtG_UID = IIf(IsNull(RsTxn_PledgeCreation!G_UID), "", RsTxn_PledgeCreation!G_UID)
mMarketYardPledge = IIf(IsNull(RsTxn_PledgeCreation!MarketYardFlag), 0, RsTxn_PledgeCreation!MarketYardFlag)

CmbBankName.Visible = False
CmbBranchName.Visible = False
TxtWHRNo.Visible = False
TxtWHRNo.Locked = True
TxtBankName.Visible = True
TxtBranchName.Visible = True
TxtPledgeFee.Locked = True
TxtPledgeReqNo.Locked = True
PledgeDate.Enabled = False
PledgeRequestDate.Enabled = False
CmdBankName.Visible = False
CmdBranchName.Visible = False
TxtBankName.Locked = True
TxtBranchName.Locked = True
TxtRemarks.Locked = True
TxtPledgeNoOfBags.Locked = True
TxtPledgeQuantity.Locked = True
TxtDiscount.Locked = True

mPrintFlag = Null

If IsNull(RsTxn_PledgeCreation!Flag) = False Then
   mPrintFlag = "P"
End If

CmbExchange.Enabled = False
CmdUpdate.Enabled = False
Call Grid_Head
If mLotWiseWHR = "0" Then
   Call TableMst_GSL(" Where Sm_prefix = '" & TxtPreFix & "'")
Else
   Call TableMst_GSL(" Where Sm_prefix = '" & TxtPreFix & "' And WHRNo = " & TxtWHRNo & " ")
End If
Call TableTxn_PledgeCreation_GSL("Where sm_prefix = '" & TxtPreFix & "' And Txn_PledgeCreation_ID = " & TxtPledgeID)
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_PledgeCreation_GSL.RecordCount
k = 1
MSHFlexGrid2.TextMatrix(0, 7) = 0 'RsTxn_PledgeCreation_GSL!PledgeNoOfBags
MSHFlexGrid2.TextMatrix(0, 8) = 0 'RsTxn_PledgeCreation_GSL!PledgeQuantity
For I = 0 To RsTxn_PledgeCreation_GSL.RecordCount - 1
    k = k + 1
    RsMst_GSL.Find "Mst_GSL_ID = " & RsTxn_PledgeCreation_GSL!Mst_GSL_ID

    If Not RsMst_GSL.EOF Then
            
       MSHFlexGrid2.TextMatrix(k, 1) = GetCMPName(RsMst_GSL!CMPID)
       MSHFlexGrid2.TextMatrix(k, 2) = GetGodownName(RsMst_GSL!GodownID, "N")
       MSHFlexGrid2.TextMatrix(k, 3) = RsMst_GSL!StackNo
       MSHFlexGrid2.TextMatrix(k, 4) = RsMst_GSL!LotNo
       MSHFlexGrid2.TextMatrix(k, 5) = "0" 'RsMst_GSL!WHRBags
       MSHFlexGrid2.TextMatrix(k, 6) = "0" 'RsMst_GSL!WHRWeight
       MSHFlexGrid2.TextMatrix(k, 7) = RsTxn_PledgeCreation_GSL!PledgeNoOfBags
       MSHFlexGrid2.TextMatrix(k, 8) = RsTxn_PledgeCreation_GSL!PledgeQuantity
       MSHFlexGrid2.TextMatrix(k, 9) = RsTxn_PledgeCreation_GSL!Mst_GSL_ID
       MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(k, 7)) 'RsTxn_PledgeCreation_GSL!PledgeNoOfBags
       MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(k, 8)) '0 'RsTxn_PledgeCreation_GSL!PledgeQuantity
       RsMst_GSL.MoveFirst
    End If
    RsTxn_PledgeCreation_GSL.MoveNext
Next I
Call GButtonLock(Me, True)
'EditCmd.Enabled = False
'DeleteCmd.Enabled = False

'If Val(MSHFlexGrid2.TextMatrix(0, 7)) = Val(TxtWHRNoOfBags) And Val(MSHFlexGrid2.TextMatrix(0, 8)) = Val(TxtWHRQuantity) Then
'    CmdPrint.Enabled = True
'Else
'    CmdPrint.Enabled = False
'End If

End Sub
Private Function GetPledgeID() As Double

Dim Retval As Double
tmp = "Select max(PledgeID) from Txn_PledgeCreation where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetPledgeID = Retval

End Function
Private Function GetPledgeNo() As Double
Dim Retval As Double
tmp = "Select max(PledgeNo) from Txn_PledgeCreation_GSL where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetPledgeNo = Retval
End Function

Private Sub SearchCmd_Click()
Dim ans As Variant
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_PledgeCreation.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
''   Call FirstCmd_Click
   RsTxn_PledgeCreation.MoveFirst
   RsTxn_PledgeCreation.Find "G_UID= '" & TxtG_UID & "'"
   If RsTxn_PledgeCreation.EOF Then
      RsTxn_PledgeCreation.MoveLast
   End If
   Call Indata
   Exit Sub
End If
'Frame0.Height = Me.Height - 600
'Frame0.Left = Frame4.Left - 100
'Frame0.Top = Frame4.Top - 100
'Frame0.Width = Me.Width - 150
'Frame0.Visible = True


'Picture2
Frame0.Height = Frame4.Height - 200 'Me.Height - 500
Frame0.Left = Picture2.Left '- 100
Frame0.Top = Picture2.Top
Frame0.Width = Frame4.Width - 150 'Me.Width - 200
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 600
MSHFlexGrid1.Height = Frame0.Height - 1400

RsTxn_PledgeCreation.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head1
'Call CmdClear_Click
''===========comment on 7 july 2012 for not clearing search grid=============


'MSHFlexGrid1.Rows = RsTxn_PledgeCreation.RecordCount + 1
'RsTxn_PledgeCreation.MoveFirst
'For i = 1 To RsTxn_PledgeCreation.RecordCount
'    MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(RsTxn_PledgeCreation!PledgeID), "", RsTxn_PledgeCreation!PledgeID)
'    MSHFlexGrid1.TextMatrix(i, 1) = IIf(IsNull(RsTxn_PledgeCreation!PledgeDate), "", RsTxn_PledgeCreation!PledgeDate)
'    If RsTxn_PledgeCreation!WHRNo <> 0 Then
'      'MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(RsTxn_PledgeCreation!WHRNo), "", RsTxn_PledgeCreation!WHRNo)
'      Call TableTxn_WHRDetail("Where sm_prefix = '" & RsTxn_PledgeCreation!SM_Prefix & "' And WHRNo = " & RsTxn_PledgeCreation!WHRNo)
'      MSHFlexGrid1.TextMatrix(i, 2) = RsTxn_WHRDetail!CMPID & "-" & ((RsTxn_PledgeCreation!WHRNo) - RsTxn_WHRDetail!CMPID * 1000000)
'    Else
'      MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(RsTxn_PledgeCreation!ISINID), "", RsTxn_PledgeCreation!ISINID)
'    End If
'    mBankID = IIf(IsNull(RsTxn_PledgeCreation!BankID), "", RsTxn_PledgeCreation!BankID)
'    MSHFlexGrid1.TextMatrix(i, 3) = GetBankName(mBankID, "N")
'    mBranchID = IIf(IsNull(RsTxn_PledgeCreation!BranchID), "", RsTxn_PledgeCreation!BranchID)
'    MSHFlexGrid1.TextMatrix(i, 4) = GetBranchName(mBranchID, "N")
'    MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(RsTxn_PledgeCreation!PledgeReqNo), "", RsTxn_PledgeCreation!PledgeReqNo)
'    MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(RsTxn_PledgeCreation!PledgeFees), "", RsTxn_PledgeCreation!PledgeFees)
'    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(RsTxn_PledgeCreation!PledgeReqDate), "", RsTxn_PledgeCreation!PledgeReqDate)
'    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(RsTxn_PledgeCreation!Remarks), "", RsTxn_PledgeCreation!Remarks)
'    MSHFlexGrid1.TextMatrix(i, 9) = GetExchangeName(RsTxn_PledgeCreation!ExchangeTypeID)
'    RsTxn_PledgeCreation.MoveNext
'Next
End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_PledgeCreation.MoveFirst
RsTxn_PledgeCreation.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 12) & "'"
If Not RsTxn_PledgeCreation.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientID
End Sub

Private Sub TxtDiscount_LostFocus()

If TxtDiscount = "" Then Exit Sub
If IsNumeric(TxtDiscount) = False Or Val(TxtDiscount) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtDiscount.SetFocus
   Exit Sub
End If
TxtReceivableFee = Val(TxtPledgeFee) - Val(TxtDiscount)
TxtBalFee = Val(TxtReceivableFee) - Val(TxtRecFee)

If Val(TxtBalFee) < 0 Then
   MsgBox "Balance fee must be grater than or equal to zero!!!"
   TxtDiscount.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub

Dim ans As Variant
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not found!!!"
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
      "Mst_Client.ClientName,  Mst_Client.ClientID " & _
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

Private Sub TxtPledgeFee_LostFocus()
''If TxtPledgeFee = "" Then Exit Sub
''If IsNumeric(TxtPledgeFee) = False Then
''   MsgBox "Invalid numeric format!!!"
''   TxtPledgeFee.SetFocus
''   Exit Sub
''End If
End Sub
Private Sub TxtPledgeNoOfBags_LostFocus()
If TxtPledgeNoOfBags = "" Then Exit Sub
If IsNumeric(TxtPledgeNoOfBags) = False Then
   MsgBox "Invalid numeric format!!!"
   TxtPledgeNoOfBags.SetFocus
   Exit Sub
End If
If Val(TxtPledgeNoOfBags) < 0 Then
   MsgBox "Nobody can pledge negative number of bags!!!"
   TxtPledgeNoOfBags.SetFocus
   Exit Sub
End If

If InStr(1, TxtPledgeNoOfBags, ".") > 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtPledgeNoOfBags.SetFocus
   Exit Sub
End If

If Val(TxtPledgeNoOfBags) > Val(TxtUnpledgeNoOfBags) Then
   MsgBox "Number of bags pledged cannot be more than the number of bags available for pledge!!!"
   TxtPledgeNoOfBags.SetFocus
   Exit Sub
End If
End Sub
Private Sub TxtPledgeQuantity_LostFocus()
If TxtPledgeQuantity = "" Then Exit Sub
If IsNumeric(TxtPledgeQuantity) = False Then
   MsgBox "Invalid numeric format!!!"
   TxtPledgeQuantity.SetFocus
   Exit Sub
End If
If Val(TxtPledgeQuantity) < 0 Then
   MsgBox "Negative Quantity can't be pledged!!!"
   TxtPledgeQuantity.SetFocus
   Exit Sub
End If

' Given Access to Super user for Entering more or less Qty on 30.03.2010
If LCase(Gen_UserRole) = "power" And LCase(Gen_UserRole) = "superadmin" Then
   If Val(TxtPledgeQuantity) > Val(TxtUnpledgeWeight) Then
      MsgBox "Warning : Pledged quantity can't be more than the quantity available for pledge !!!"
      'TxtPledgeQuantity.SetFocus
      'Exit Sub
   End If
Else
   If Val(TxtPledgeQuantity) > Val(TxtUnpledgeWeight) Then
      MsgBox "Pledged quantity can't be more than the quantity available for pledge!!!"
      TxtPledgeQuantity.SetFocus
      Exit Sub
   End If
End If
'CmdUpdate.Enabled = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_PledgeCreation.MoveNext
If RsTxn_PledgeCreation.EOF Then
   RsTxn_PledgeCreation.MoveLast
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
RsTxn_PledgeCreation.MovePrevious
If RsTxn_PledgeCreation.BOF Then
   RsTxn_PledgeCreation.MoveFirst
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
RsTxn_PledgeCreation.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_PledgeCreation.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Function ChkSelection() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 3 Then
   For I = 2 To MSHFlexGrid2.Rows - 1
      If MSHFlexGrid2.TextMatrix(I, 0) = "*" Then
         Retval = True
         Exit For
      End If
   Next
End If
ChkSelection = Retval
End Function
Private Sub EnDn(Exg, Yn)

If Exg = "Non NCDEX" Then
   Label1.Caption = "WHR No"
   Label2.Caption = "License No"
ElseIf Exg = "NCDEX" Then
   Label1.Caption = "CMSE Lot No"
   Label2.Caption = "CMP"
End If
TxtClientID.Enabled = Yn
TxtClientName.Enabled = Yn
WHRDate.Enabled = Yn
TxtWHRNoOfBags.Enabled = Yn
TxtWHRQuantity.Enabled = Yn
'Label6.Enabled = Yn
'Label8.Enabled = Yn
Label16.Enabled = Yn
Label17.Enabled = Yn
Label14.Enabled = Yn

TxtClientID.Text = ""
TxtClientName.Text = ""
TxtWHRNoOfBags.Text = ""
TxtWHRQuantity.Text = ""

End Sub

Private Sub TxtWHRNo_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtWHRNoNew.Text = "" Then
   Call Grid_Head
   Exit Sub
End If
If CmbExchange.Text = "Non NCDEX" Then
   If RsTxn_WHRDetail.RecordCount <= 0 Then
      MsgBox "No WHR Data found!!! "
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
   Dim strSplitString() As String
   If TxtWHRNoNew.Text = "" Then
      Exit Sub
   End If
   strSplitString = Split(TxtWHRNoNew.Text, "-")
   tmp = UBound(strSplitString)
   If tmp <> 1 Then
      MsgBox "Invalid WHR Number!!!"
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
   TxtWHRNo = (Val(strSplitString(0)) * 1000000) + Val(strSplitString(1))
   RsTxn_WHRDetail.MoveFirst
   RsTxn_WHRDetail.Find "WHRNo = " & TxtWHRNo.Text & ""
   If RsTxn_WHRDetail.EOF Then
      MsgBox "Invalid selection!!!"
      Call Grid_Head
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
'   If RsTxn_WHRDetail!Flag <> "P" Or IsNull(RsTxn_WHRDetail!Flag) Then
'      MsgBox " WHR is not yet printed!!! "
'      TxtWHRNoNew.SetFocus
'      Exit Sub
'   End If
   If RsTxn_WHRDetail!DeliveredNoOfBags = RsTxn_WHRDetail!WHRNoOfBags Then
      MsgBox "WHR already withdrawn!!!"
      If Gen_WHRDetail(0) <> False Then
         TxtWHRNoNew.SetFocus
      End If
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
   
   If RsTxn_WHRDetail!Txn_PledgeCreation_ID <> 0 Then
      MsgBox "WHR already pledged. Pledge ID is " & RsTxn_WHRDetail!Txn_PledgeCreation_ID & "!!!"
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
    
   If LCase(RsTxn_WHRDetail!Flag) = "n" Then
      MsgBox "Print Out of WHR which was pledge earlier is mandatory to make it Non Pledgable WHR !!!"
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
    
   mWHRNo = RsTxn_WHRDetail!WHRNo
   mWHR_Detail_ID = RsTxn_WHRDetail!Txn_WHR_Detail_ID
   WHRDate.Value = RsTxn_WHRDetail!WHRDate
   WHRValidDate.Value = RsTxn_WHRDetail!StorageTo  'RsTxn_WHRDetail!WHRDate
   mClientIDRow = RsTxn_WHRDetail!ClientID
   ans = GetClientNameByRow(mClientIDRow)
   If ans(0) = "" Then
      MsgBox "Not found!!!"
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
   TxtClientID.Text = ans(0)
   TxtClientName = ans(1)
   mLotWiseWHR = IIf(IsNull(RsTxn_WHRDetail!LotWiseWHR), 0, RsTxn_WHRDetail!LotWiseWHR)
   mCommodityID = IIf(IsNull(RsTxn_WHRDetail!CommodityID), "", RsTxn_WHRDetail!CommodityID)
   TxtCommodity = GetCommodityName(mCommodityID)
   TxtLicenseNo = GetLicenseNo(RsTxn_WHRDetail!LicenseID)
   TxtWHRNoOfBags = IIf(IsNull(RsTxn_WHRDetail!WHRNoOfBags), "", RsTxn_WHRDetail!WHRNoOfBags)
   TxtWHRQuantity = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "", RsTxn_WHRDetail!WHRQuantity)
   TxtPreFix = IIf(IsNull(RsTxn_WHRDetail!SM_Prefix), "", RsTxn_WHRDetail!SM_Prefix)
   TxtWHRNoOfBags = TxtWHRNoOfBags - IIf(IsNull(RsTxn_WHRDetail!DeliveredNoOfBags), 0, RsTxn_WHRDetail!DeliveredNoOfBags)
   TxtWHRQuantity = TxtWHRQuantity - IIf(IsNull(RsTxn_WHRDetail!DeliveredQuantity), 0, RsTxn_WHRDetail!DeliveredQuantity)
   TxtWHRQuantity = Format(TxtWHRQuantity, "#########0.000")
   mCMPID = IIf(IsNull(RsTxn_WHRDetail!CMPID), "", RsTxn_WHRDetail!CMPID)
   mGodownID = IIf(IsNull(RsTxn_WHRDetail!GodownID), "", RsTxn_WHRDetail!GodownID)
   mSendToMktYard = IIf(IsNull(RsTxn_WHRDetail!SendToMktYard), 0, RsTxn_WHRDetail!SendToMktYard) 'SendToMktYard
   If mSendToMktYard = True Then
      mSendToMktYard = 1
   Else
      mSendToMktYard = 0
   End If
   If mGodownID = 2283 Then
      TxtPreFix = "BKNRP"
      mCMPID = 678
      mGodownID = 2316
   ElseIf mGodownID = 25 Then
      TxtPreFix = "ujjap"
      mCMPID = 705
      mGodownID = 2378
   End If
   
   If Not IsNull(RsTxn_WHRDetail!GodownID) Then
      TxtGodownNo = GetGodownName(RsTxn_WHRDetail!GodownID, "N")
      mGodownAddress = GetGodownName(IIf(IsNull(RsTxn_WHRDetail!GodownID), "", RsTxn_WHRDetail!GodownID), "A")
   End If
   mRate = IIf(IsNull(RsTxn_WHRDetail!Rate), "", RsTxn_WHRDetail!Rate)
   mValueOfCommodity = Val(TxtWHRQuantity) * Val(mRate)
   CmbBankName.Text = ""
   CmbBranchName.Text = ""
   Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow, mCMPID)
Else
   TxtWHRNo = TxtWHRNoNew
   RsMst_ISIN.MoveFirst
   RsMst_ISIN.Find "ISINID = '" & TxtWHRNo.Text & "'"
   If RsMst_ISIN.EOF Then
      MsgBox "Invalid selection!!!"
      Call Grid_Head
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
   mISINID = RsMst_ISIN!ISINID
   mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
   TxtCommodity = GetCommodityName(mCommodityID)
   TxtLicenseNo = GetCMPName(RsMst_ISIN!CMPID)
   mCMPID = IIf(IsNull(RsMst_ISIN!CMPID), "", RsMst_ISIN!CMPID)
   TxtPreFix = GetPreFix(RsMst_ISIN!CMPID, "P")
   WHRValidDate.Value = RsMst_ISIN!FED
   mGodownAddress = ""
   CmbBankName.Text = ""
   CmbBranchName.Text = ""
   Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow, mCMPID)
End If

End Sub

Private Sub TxtWHRNoNew_LostFocus()

If SaveCmd.Enabled = False Then Exit Sub
Call TxtWHRNo_LostFocus ' = TxtWHRNoNew
TxtCMP = ""
TxtGSL = ""
TxtUnpledgeNoOfBags = ""
TxtUnpledgeWeight = ""
TxtPledgeNoOfBags = ""
TxtPledgeQuantity = ""
TxtPledgeNoOfBags.Locked = True
TxtPledgeQuantity.Locked = True
CmdUpdate.Enabled = False
'If TxtWHRNoNew.Text <> "" And CmbBankName.Text <> "" Then
Call GetPledgeFee(mCMAgreement)
'End If
End Sub
Private Sub CmbSearchExchange_GotFocus()
mExchangeID = Null
tmp = CmbSearchExchange.Text
Call TableMst_ExchangeType
CmbSearchExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSearchExchange.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSearchExchange.Text = tmp

End Sub

Private Sub CmbSearchExchange_LostFocus()
Dim x As Integer
If CmbSearchExchange.Text = "" Then
   TxtSearchPledge = ""
   TxtSearchPledge.Locked = True
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSearchExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSearchExchange.SetFocus
   Exit Sub
End If



mSExchangeID = RsMst_ExchangeType!ExchangeTypeID
TxtSearchPledge = ""
TxtSearchPledge.Locked = False


End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_PledgeCreation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_PledgeCreation.xls")
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
           If C = "4" Then
              oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

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

mSLocationID = RsMst_Location!LocationID
mLocationName = RsMst_Location!Locationname
msSM_prefix = RsMst_Location!SM_Prefix

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

End Sub


Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub
Private Sub SetScrollBars()
'Frame1.Width = 10000
'Frame1.Height = 20000
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

Private Function PDF_Print_Rpt()
Dim mstr As String
Dim mTotTop As Double
    
Set objPDF = New mjwPDF

' Create a simple PDF file using the mjwPDF class

Gen_mTimeStamp = GetTimeStamp
' Set the PDF title and filename
objPDF.PDFTitle = "Quotation PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_PledgeConfirmation-" & Trim(TxtPledgeID) & ".pdf"

' We must tell the class where the PDF fonts are located
objPDF.PDFLoadAfm = App.Path & "\Fonts"

' Set the file properties
objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
objPDF.PDFFormatPage = FORMAT_A4
objPDF.PDFOrientation = ORIENT_PORTRAIT
objPDF.PDFSetUnit = UNIT_PT
objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI


' Lets us set see the bookmark pane when we view the PDF
'objPDF.PDFUseOutlines = True

' View the PDF file after we create it
objPDF.PDFView = True
'    objPDF.PDFWatermark = "NCMSL"
' Begin our PDF document

objPDF.PDFBeginDoc
    ' Lets add a heading
objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFWatermark = "NCMSL"
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL

objPDF.PDFSetFill = True

''
''    objPDF.PDFCell " ", 15, 15, _
''    objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
''

pgval = ((objPDF.PDFGetPageWidth - 18))

'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)

mRowHeight = 15
mGlobalLeft = 12
'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
mLeft = mGlobalLeft
mTop = 15
mHeight = mTop + mRowHeight
mTxtFromTop = mRowHeight + 2
 
'objPDF.PDFImage App.Path & "\images\Logo.gif", 20, 17, 60, 30, ""

 'Call DrawTable(10, 2, 15, 30, 12)

'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) + 35
mStrLength = Len("QMS certified as per ISO 9001:2008")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = 20 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 6

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

mTxtTop = mTop + mRowHeight
mStrLength = Len("National Collateral Management Services Ltd.")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2)
mFontSize = 8

Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mTxtLeft, mTxtTop, "National Collateral Management Services Ltd.")

mTxtTop = mTxtTop + mRowHeight
mStrLength = Len("Lodha Supremus, Unit No 505 to 509, 5th Floor,")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 7

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Lodha Supremus, Unit No 505 to 509, 5th Floor,")

mTxtTop = mTxtTop + mRowHeight
mStrLength = Len("Kanjurmarg (East), Mumbai-400042")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 7

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Kanjurmarg (East), Mumbai-400042")
  
mTxtTop = mTxtTop + mRowHeight
mStrLength = Len("Tel.No-022 40419191, Fax. No-022 40419193")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 7

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Tel.No-022 40419191, Fax. No-022 40419193")

mTxtTop = mTxtTop + 10

objPDF.PDFDrawLine mGlobalLeft + 20, mTxtTop, mGlobalLeft + 550, mTxtTop

mStrLength = Len("Quotation")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

mTxtTop = mTxtTop + mRowHeight

mTxtTop = mTxtTop + mRowHeight
mStrLength = Len("Pledge ID :" & TxtPledgeID)
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
mTxtTop = mTxtTop + 10
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "Pledge ID :" & TxtPledgeID)
   
mStrLength = Len("Pledge Date : " & Format(PledgeDate.Value, "dd-MMM-yyyy"))
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
mTxtTop = mTxtTop + 15
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "Pledge Date : " & Format(PledgeDate.Value, "dd-MMM-yyyy"))

mTxtTop = mTxtTop + mRowHeight + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "To,")

mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "The Branch Manager")

mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, TxtBankName)

mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "Branch : " & TxtBranchName)

mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270

Dim mLen As Integer, mMax As Integer, mPara As Integer, mAdrHt As Double
Dim t1 As Double
Dim mAddr As String
mAddr = GetBranchName(RsTxn_PledgeCreation!BranchID, "A")
mMax = 90
mPara = 0
t1 = 1
mLen = Len(mAddr)
mLen = Round(mLen / mMax) + 1

For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       'mMax '= mMax + 35
       mPara = mPara + 7
    End If
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop + mPara, mID$(mAddr, t1, mMax))
    mAdrHt = mTxtTop + 10 + mPara
Next
        
mTxtTop = mAdrHt + mRowHeight
mStrLength = Len("SUB : Pledge Confirmation Letter")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10
mTxtTop = mTxtTop + mRowHeight
Call PrintText(mFontSize, FONT_BOLD & FONT_UNDERLINE, ALIGN_CENTER, mTxtLeft, mTxtTop, "SUB : Pledge Confirmation Letter")
   
mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "Dear Sir,")
   
mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 65 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "We confirm that the following Warehouse Receipt(WHR) is pledged in your favour, vide your Pledge Request No ")


mTxtTop = mTxtTop + mRowHeight - 5
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, TxtPledgeReqNo & " dated  " & Format(PledgeRequestDate.Value, "dd-MMM-yyyy"))

mTxtTop = mTxtTop + mRowHeight + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "Client Name")
mTxtLeft = mGlobalLeft + 90
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, ": " & TxtClientName)
mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "WHR No")
mTxtLeft = mGlobalLeft + 90
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, ": " & TxtWHRNoNew.Text)

mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "WHR Date")
mTxtLeft = mGlobalLeft + 90
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, ": " & Format(WHRValidDate.Value, "dd-MMM-yyyy"))

mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "Commodity")
mTxtLeft = mGlobalLeft + 90
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, ": " & TxtCommodity)


mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "No Of Bags ")
mTxtLeft = mGlobalLeft + 90
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, ": " & TxtWHRNoOfBags)

mTxtTop = mTxtTop + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "Quantity ")
mTxtLeft = mGlobalLeft + 90
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, ": " & TxtWHRQuantity & " MT")

mTxtTop = mTxtTop + mRowHeight + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "The details of pledged lots are given below .")
   
mTxtTop = mTxtTop + mRowHeight + 10

objPDF.PDFDrawLine mGlobalLeft + 30, mTxtTop, mGlobalLeft + 540, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 30, mTxtTop + mRowHeight, mGlobalLeft + 540, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 30, mTxtTop, mGlobalLeft + 30, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 260, mTxtTop, mGlobalLeft + 260, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 320, mTxtTop, mGlobalLeft + 320, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 380, mTxtTop, mGlobalLeft + 380, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 470, mTxtTop, mGlobalLeft + 470, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 540, mTxtTop, mGlobalLeft + 540, mTxtTop + mRowHeight

mHeight = mTop + mRowHeight
mTxtLeft = mGlobalLeft + 33 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, "Godown Address")

mTxtLeft = mGlobalLeft + 263 'mPrintPoint - (mStrLength * 2) '* 5) '270
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, "Stack No")

mTxtLeft = mGlobalLeft + 323 'mPrintPoint - (mStrLength * 2) '* 5) '270
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, "Lot No")

mTxtLeft = mGlobalLeft + 383 'mPrintPoint - (mStrLength * 2) '* 5) '270
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, "Pledged No Of Bags")
             
mTxtLeft = mGlobalLeft + 473 'mPrintPoint - (mStrLength * 2) '* 5) '270
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, "Pledged Qty")
      
For I = 2 To MSHFlexGrid2.Rows - 2

    mTxtTop = mTxtTop + mRowHeight

    If I = 2 Then
       mTxtLeft = mGlobalLeft + 33
       
        mAddr = mGodownAddress
        mMax = 45
        mPara = 0
        t1 = 1
        mLen = Len(mAddr)
        mLen = Round(mLen / mMax) + 1
    
        For T = 1 To mLen
            If T > 1 Then
               t1 = mMax + t1
               'mMax '= mMax + 35
               mPara = mPara + 7
            End If
            Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop + 10 + mPara, mID$(mAddr, t1, mMax))
            mAdrHt = mTxtTop + 10 + mPara
        Next
       'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, mStr)
       
    End If
    
    objPDF.PDFDrawLine mGlobalLeft + 260, mTxtTop, mGlobalLeft + 540, mTxtTop
    
    If I = MSHFlexGrid2.Rows - 2 Then
       If mAdrHt > mTxtTop + mRowHeight Then
          objPDF.PDFDrawLine mGlobalLeft + 30, mAdrHt, mGlobalLeft + 540, mAdrHt
          
          objPDF.PDFDrawLine mGlobalLeft + 30, mTxtTop, mGlobalLeft + 30, mAdrHt
          objPDF.PDFDrawLine mGlobalLeft + 260, mTxtTop, mGlobalLeft + 260, mAdrHt
          objPDF.PDFDrawLine mGlobalLeft + 320, mTxtTop, mGlobalLeft + 320, mAdrHt
          objPDF.PDFDrawLine mGlobalLeft + 380, mTxtTop, mGlobalLeft + 380, mAdrHt
          objPDF.PDFDrawLine mGlobalLeft + 470, mTxtTop, mGlobalLeft + 470, mAdrHt
          objPDF.PDFDrawLine mGlobalLeft + 540, mTxtTop, mGlobalLeft + 540, mAdrHt
            
          mTotTop = mAdrHt
       Else
          objPDF.PDFDrawLine mGlobalLeft + 30, mTxtTop + mRowHeight, mGlobalLeft + 540, mTxtTop + mRowHeight
          
          objPDF.PDFDrawLine mGlobalLeft + 30, mTxtTop, mGlobalLeft + 30, mTxtTop + mRowHeight
          objPDF.PDFDrawLine mGlobalLeft + 260, mTxtTop, mGlobalLeft + 260, mTxtTop + mRowHeight
          objPDF.PDFDrawLine mGlobalLeft + 320, mTxtTop, mGlobalLeft + 320, mTxtTop + mRowHeight
          objPDF.PDFDrawLine mGlobalLeft + 380, mTxtTop, mGlobalLeft + 380, mTxtTop + mRowHeight
          objPDF.PDFDrawLine mGlobalLeft + 470, mTxtTop, mGlobalLeft + 470, mTxtTop + mRowHeight
          objPDF.PDFDrawLine mGlobalLeft + 540, mTxtTop, mGlobalLeft + 540, mTxtTop + mRowHeight
          mTotTop = mTxtTop + mRowHeight
             
       End If
    Else
       objPDF.PDFDrawLine mGlobalLeft + 260, mTxtTop + mRowHeight, mGlobalLeft + 540, mTxtTop + mRowHeight
       objPDF.PDFDrawLine mGlobalLeft + 30, mTxtTop, mGlobalLeft + 30, mTxtTop + mRowHeight
       objPDF.PDFDrawLine mGlobalLeft + 260, mTxtTop, mGlobalLeft + 260, mTxtTop + mRowHeight
       objPDF.PDFDrawLine mGlobalLeft + 320, mTxtTop, mGlobalLeft + 320, mTxtTop + mRowHeight
       objPDF.PDFDrawLine mGlobalLeft + 380, mTxtTop, mGlobalLeft + 380, mTxtTop + mRowHeight
       objPDF.PDFDrawLine mGlobalLeft + 470, mTxtTop, mGlobalLeft + 470, mTxtTop + mRowHeight
       objPDF.PDFDrawLine mGlobalLeft + 540, mTxtTop, mGlobalLeft + 540, mTxtTop + mRowHeight
        
    End If
           
    mTxtLeft = mGlobalLeft + 263
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, MSHFlexGrid2.TextMatrix(I, 3))
    
    mTxtLeft = mGlobalLeft + 323
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, MSHFlexGrid2.TextMatrix(I, 4))
    
    mTxtLeft = mGlobalLeft + 383
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, MSHFlexGrid2.TextMatrix(I, 7))
                 
    mTxtLeft = mGlobalLeft + 473
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, MSHFlexGrid2.TextMatrix(I, 8))
    
Next

mTxtTop = mTotTop
objPDF.PDFDrawLine mGlobalLeft + 30, mTotTop + mRowHeight, mGlobalLeft + 540, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 30, mTxtTop, mGlobalLeft + 30, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 380, mTxtTop, mGlobalLeft + 380, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 470, mTxtTop, mGlobalLeft + 470, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft + 540, mTxtTop, mGlobalLeft + 540, mTxtTop + mRowHeight
 
mTxtLeft = mGlobalLeft + 328
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, "Total")
mTxtLeft = mGlobalLeft + 383
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, MSHFlexGrid2.TextMatrix(0, 7))
mTxtLeft = mGlobalLeft + 473
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop + 10, MSHFlexGrid2.TextMatrix(0, 8))

              
mTxtTop = mTxtTop + mRowHeight + mRowHeight + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "Yours faithfully")
 
mTxtTop = mTxtTop + mRowHeight + mRowHeight + mRowHeight
mTxtLeft = mGlobalLeft + 30 'mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 8
mHeight = mTop + mRowHeight
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mTxtLeft, mTxtTop, "(CMP Manager)")

objPDF.PDFEndDoc

End Function

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


Private Sub PrintText(mFontSize_ As Integer, mFontFormat_ As PDFFontStl, mAlign_ As Variant, mTxtLeft_ As Double, mTxtTop_ As Double, mTexttoPrint As String)

'        objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'        objPDF.PDFSetAlignement = ALIGN_CENTER
        
objPDF.PDFSetFont FONT_ARIAL, mFontSize_, mFontFormat_
objPDF.PDFSetAlignement = mAlign_
objPDF.PDFTextOut mTexttoPrint, mTxtLeft_, mTxtTop_


End Sub


