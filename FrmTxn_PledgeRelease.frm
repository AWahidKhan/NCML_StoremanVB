VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_PledgeRelease 
   Caption         =   "Pledge Release"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   8895
      Left            =   15000
      TabIndex        =   81
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   300
      Left            =   0
      TabIndex        =   80
      Top             =   9720
      Width           =   12495
   End
   Begin VB.PictureBox Frame4 
      Height          =   9540
      Left            =   60
      ScaleHeight     =   9480
      ScaleWidth      =   15045
      TabIndex        =   9
      Top             =   120
      Width           =   15105
      Begin VB.Frame Frame0 
         Height          =   1605
         Left            =   1800
         TabIndex        =   82
         Top             =   6000
         Visible         =   0   'False
         Width           =   10635
         Begin VB.ComboBox CmbSPreFix 
            Height          =   315
            Left            =   2850
            Sorted          =   -1  'True
            TabIndex        =   91
            Top             =   315
            Width           =   1920
         End
         Begin VB.ComboBox CmbSBranchName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7845
            Sorted          =   -1  'True
            TabIndex        =   90
            Top             =   330
            Width           =   2985
         End
         Begin VB.ComboBox CmbSBankName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4815
            Sorted          =   -1  'True
            TabIndex        =   89
            Top             =   330
            Width           =   2985
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "Clear"
            Height          =   345
            Left            =   13455
            TabIndex        =   88
            Top             =   735
            Width           =   1005
         End
         Begin VB.TextBox TxtSearchPledge 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   12930
            MaxLength       =   200
            TabIndex        =   87
            Top             =   315
            Width           =   1845
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   345
            Left            =   11445
            TabIndex        =   86
            Top             =   735
            Width           =   1005
         End
         Begin VB.ComboBox CmbSearchExchange 
            Height          =   315
            ItemData        =   "FrmTxn_PledgeRelease.frx":0000
            Left            =   10860
            List            =   "FrmTxn_PledgeRelease.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   85
            Top             =   330
            Width           =   2070
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   12450
            TabIndex        =   84
            Top             =   735
            Width           =   1000
         End
         Begin VB.TextBox TxtSClientName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1080
            TabIndex        =   83
            Top             =   720
            Width           =   4665
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   92
            Top             =   1200
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   2566
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   300
            Left            =   1470
            TabIndex        =   93
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
            Format          =   111804417
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   120
            TabIndex        =   94
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
            Format          =   111804417
            CurrentDate     =   36494
         End
         Begin VB.Label Label36 
            Caption         =   "Location"
            Height          =   300
            Left            =   3473
            TabIndex        =   104
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
            Height          =   255
            Left            =   345
            TabIndex        =   103
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   1800
            TabIndex        =   102
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label30 
            Caption         =   "Branch Name"
            Height          =   225
            Left            =   8805
            TabIndex        =   101
            Top             =   120
            Width           =   1065
         End
         Begin VB.Label Label29 
            Caption         =   "Bank Name"
            Height          =   255
            Left            =   5872
            TabIndex        =   100
            Top             =   120
            Width           =   870
         End
         Begin VB.Label Label25 
            Caption         =   "WHR/ISIN"
            Height          =   240
            Left            =   13500
            TabIndex        =   99
            Top             =   150
            Width           =   810
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
            Left            =   8805
            TabIndex        =   98
            Top             =   750
            Width           =   135
         End
         Begin VB.Label Label27 
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
            Left            =   6285
            TabIndex        =   97
            Top             =   750
            Width           =   2415
         End
         Begin VB.Label Label28 
            Caption         =   "Exchange Type"
            Height          =   225
            Left            =   11288
            TabIndex        =   96
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label32 
            Caption         =   "Client Name"
            Height          =   210
            Left            =   120
            TabIndex        =   95
            Top             =   765
            Width           =   885
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   13845
         TabIndex        =   69
         Top             =   8475
         Width           =   13905
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   10725
            TabIndex        =   79
            Top             =   465
            Width           =   2655
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   375
            Left            =   8085
            TabIndex        =   78
            Top             =   465
            Width           =   2640
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   375
            Left            =   5445
            TabIndex        =   77
            Top             =   465
            Width           =   2640
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   2790
            TabIndex        =   76
            Top             =   465
            Width           =   2655
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   375
            Left            =   135
            TabIndex        =   75
            Top             =   465
            Width           =   2655
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   5445
            TabIndex        =   74
            Top             =   90
            Width           =   2640
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   375
            Left            =   2790
            TabIndex        =   73
            Top             =   90
            Width           =   2655
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New "
            Height          =   375
            Left            =   135
            TabIndex        =   72
            Top             =   90
            Width           =   2655
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   8085
            TabIndex        =   71
            Top             =   75
            Width           =   2640
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "&Search"
            Height          =   375
            Left            =   10725
            TabIndex        =   70
            Top             =   90
            Width           =   2655
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   3345
         Left            =   120
         ScaleHeight     =   3285
         ScaleWidth      =   13920
         TabIndex        =   10
         Top             =   120
         Width           =   13980
         Begin VB.TextBox TxtWHRNo 
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1380
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.ComboBox CmbPreFix 
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
            TabIndex        =   1
            Top             =   480
            Visible         =   0   'False
            Width           =   2670
         End
         Begin VB.TextBox TxtReleaseOrderNo 
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
            MaxLength       =   30
            TabIndex        =   0
            Top             =   45
            Width           =   2670
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
            Height          =   375
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   2235
            Width           =   2430
         End
         Begin VB.TextBox TxtReleaseID 
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
            Left            =   9840
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   30
            Width           =   1950
         End
         Begin VB.TextBox TxtWHRNoNew 
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
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1380
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
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1800
            Width           =   7710
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
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2235
            Width           =   2670
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2670
         End
         Begin VB.TextBox TxtPledgeID 
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
            Top             =   945
            Width           =   2670
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
            Left            =   9840
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   2235
            Width           =   4005
         End
         Begin VB.TextBox TxtPledgeNoOfBags 
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
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1950
         End
         Begin VB.TextBox TxtPledgeQuantity 
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
            Left            =   9840
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1950
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
            Height          =   555
            Left            =   1905
            MaxLength       =   254
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   2655
            Width           =   11940
         End
         Begin MSComCtl2.DTPicker PledgeDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   111804417
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker WHRDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   19
            Top             =   1335
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   111804417
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker ReleaseDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   3
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   661
            _Version        =   393216
            Format          =   111804417
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtPreFix 
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
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label62 
            Caption         =   "Location ID "
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
            TabIndex        =   62
            Top             =   540
            Width           =   1575
         End
         Begin VB.Label Label22 
            Caption         =   "Release Order No"
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
            TabIndex        =   61
            Top             =   105
            Width           =   1920
         End
         Begin VB.Label Label23 
            Caption         =   "Exchange Type"
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
            Left            =   4785
            TabIndex        =   60
            Top             =   2160
            Width           =   1200
         End
         Begin VB.Label Label21 
            Caption         =   "Release Date"
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
            Left            =   4785
            TabIndex        =   58
            Top             =   105
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Release ID"
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
            Left            =   8355
            TabIndex        =   57
            Top             =   90
            Width           =   1335
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   1335
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
            TabIndex        =   29
            Top             =   1860
            Width           =   1365
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
            Left            =   120
            TabIndex        =   28
            Top             =   2295
            Width           =   1455
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
            Left            =   4800
            TabIndex        =   27
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Client ID"
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
            TabIndex        =   26
            Top             =   1860
            Width           =   1590
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
            Left            =   120
            TabIndex        =   25
            Top             =   1005
            Width           =   1335
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
            Left            =   8760
            TabIndex        =   24
            Top             =   2295
            Width           =   1095
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
            TabIndex        =   23
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Pledged No of Bags"
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
            Left            =   7860
            TabIndex        =   22
            Top             =   1020
            Width           =   1830
         End
         Begin VB.Label Label17 
            Caption         =   "Pledged Qty."
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
            Left            =   10200
            TabIndex        =   21
            Top             =   1020
            Width           =   1215
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
            TabIndex        =   20
            Top             =   2745
            Width           =   1215
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4920
         Left            =   120
         TabIndex        =   31
         Top             =   3510
         Width           =   14040
         _ExtentX        =   24765
         _ExtentY        =   8678
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "GSL Detail"
         TabPicture(0)   =   "FrmTxn_PledgeRelease.frx":0004
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Pledge Detail"
         TabPicture(1)   =   "FrmTxn_PledgeRelease.frx":0020
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "History Detail"
         TabPicture(2)   =   "FrmTxn_PledgeRelease.frx":003C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TxtCreated"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "TxtUpdated"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label46"
         Tab(2).Control(3)=   "Label47"
         Tab(2).ControlCount=   4
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
            Left            =   -73635
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   720
            Width           =   10890
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
            Left            =   -73635
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1485
            Width           =   10875
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   13770
            Begin VB.TextBox TxtUnpledgeWeight 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   8280
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   420
               Width           =   1965
            End
            Begin VB.TextBox TxtUnpledgeNoOfBags 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   6315
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   420
               Width           =   1905
            End
            Begin VB.TextBox TxtCMP 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   420
               Width           =   3315
            End
            Begin VB.TextBox TxtGSL 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   420
               Width           =   2775
            End
            Begin VB.CommandButton CmdUpdate 
               Caption         =   "Update List"
               Height          =   360
               Left            =   12105
               TabIndex        =   7
               Top             =   870
               Width           =   1620
            End
            Begin VB.TextBox TxtReleaseNoOfBags 
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
               Left            =   10290
               TabIndex        =   5
               Top             =   420
               Width           =   1785
            End
            Begin VB.TextBox TxtReleaseQuantity 
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
               Left            =   12112
               TabIndex        =   6
               Top             =   420
               Width           =   1620
            End
            Begin VB.Label Label15 
               Caption         =   "Pledged No Of Bags"
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
               Left            =   6360
               TabIndex        =   55
               Top             =   120
               Width           =   1875
            End
            Begin VB.Label Label10 
               Caption         =   "CMP"
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
               Left            =   1267
               TabIndex        =   54
               Top             =   165
               Width           =   1020
            End
            Begin VB.Label Label12 
               Caption         =   "G-S-L"
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
               Left            =   4537
               TabIndex        =   53
               Top             =   120
               Width           =   660
            End
            Begin VB.Label Label13 
               Caption         =   "Pledged  Qty."
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
               Left            =   8520
               TabIndex        =   52
               Top             =   120
               Width           =   1260
            End
            Begin VB.Label Label3 
               Caption         =   "Released No Of Bags"
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
               Left            =   10200
               TabIndex        =   51
               Top             =   120
               Width           =   1995
            End
            Begin VB.Label Label4 
               Caption         =   "Released Qty."
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
               Left            =   12300
               TabIndex        =   50
               Top             =   120
               Width           =   1365
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3555
            Left            =   -74880
            TabIndex        =   32
            Top             =   390
            Width           =   13815
            Begin VB.TextBox TxtBranchName 
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
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   660
               Width           =   6825
            End
            Begin VB.TextBox TxtBankName 
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
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   240
               Width           =   6825
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
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   1545
               Width           =   2040
            End
            Begin VB.TextBox TxtPledgeReqNo 
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
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   1095
               Width           =   6825
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
               Left            =   3120
               Sorted          =   -1  'True
               TabIndex        =   34
               Top             =   240
               Width           =   6825
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
               Left            =   3120
               Sorted          =   -1  'True
               TabIndex        =   33
               Top             =   660
               Width           =   6825
            End
            Begin MSComCtl2.DTPicker PledgeRequestDate 
               Height          =   375
               Left            =   3120
               TabIndex        =   39
               Top             =   1995
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   111804417
               CurrentDate     =   39884
            End
            Begin VB.Label Label11 
               Caption         =   "Pledge Fee Receviable"
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
               TabIndex        =   44
               Top             =   1590
               Width           =   2535
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
               Left            =   120
               TabIndex        =   43
               Top             =   1125
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
               TabIndex        =   42
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
               TabIndex        =   41
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
               Left            =   120
               TabIndex        =   40
               Top             =   2055
               Width           =   2895
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2985
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   5265
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
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
            Left            =   -74880
            TabIndex        =   68
            Top             =   780
            Width           =   1695
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
            Left            =   -74880
            TabIndex        =   67
            Top             =   1545
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_PledgeRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mPledgeNo As Variant
Dim mGSLID, mLicenseID, mClientIDRow As Double
Dim mCMPID As Double
Dim mCommodityID As Double
Dim mExchangeTypeID As Double
Dim mSBankID As Double
Dim mROBags As Double
Dim mROQty As Double
Dim mROCheck As Boolean
Dim mPreMsg  As Variant
Dim mReleaseBags, mReleaseQty, mSM_Prefix As Variant
Dim mSCMFlag, mNCMLFinance As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
ReleaseDate.Value = Date
Call ClearData
TxtReleaseOrderNo.Locked = False
CmbPreFix.Visible = True
TxtPreFix.Visible = False
CmbPreFix.Text = ""
Frame1.Enabled = True
If RsTxn_PledgeRelease.RecordCount > 0 Then
   ReleaseDate.SetFocus
End If
Call GButtonLock(Me, False)
End Sub
Private Sub ClearData()
TxtReleaseID = ""
TxtPledgeID = ""
TxtReleaseOrderNo = ""
TxtWHRNo = ""
TxtClientID = ""
TxtClientName = ""
TxtCommodity = ""
TxtLicenseNo = ""
TxtExchangeTypeID = ""
TxtPledgeNoOfBags = ""
TxtPledgeQuantity = ""
TxtRemarks = ""
TxtCreated = ""
TxtUpdated = ""
TxtWHRNoNew = ""
TxtReleaseOrderNo.Locked = False
TxtPledgeID.Locked = False
TxtRemarks.Locked = False
ReleaseDate.Enabled = True
Call Grid_Head
End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text
Call TableMst_Location("")

CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbPreFix.SetFocus
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
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location ID selection !!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub CmbSPreFix_GotFocus()

tmp = CmbSPreFix.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If
CmbSPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSPreFix.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSPreFix.Text = tmp

End Sub

Private Sub CmbSPreFix_LostFocus()

If CmbSPreFix.Text = "" Then
   mLocationID = Null
   mSM_Prefix = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSPreFix.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSPreFix.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdClear_Click()
SFrom.Value = Date - 365
STo.Value = Date
CmbSPreFix.Text = ""
CmbSBankName.Text = ""
CmbSBranchName.Text = ""
CmbSearchExchange.Text = ""
TxtSearchPledge = ""
TxtSearchPledge.Locked = True
Label25.Caption = "WHR/ISIN"
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

tmp = "Select TPR.Txn_PledgeRelease_ID,ML.LocationName,TPR.ReleaseOrderNo,TPR.PledgeReleaseDate,TPC.Txn_PledgeCreation_ID,ME.ExchangeType,WD.WHRNo,WD.CMPID,MI.Code, "
tmp = tmp & " MBank.BankName,MBranch.BranchName,MC.ClientName,TPR.Remarks, "
tmp = tmp & " Sum(ReleasedNoOfBags) as 'ReleaseBags',Sum(ReleasedQuantity) as 'ReleaseQty',TPR.G_UID "
tmp = tmp & " From Txn_PledgeRelease TPR"
tmp = tmp & " Inner Join Txn_PledgeCreation TPC On TPR.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID And TPR.SM_Prefix = TPC.SM_Prefix "
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPC.ExchangeTypeID = ME.ExchangeTypeID "
tmp = tmp & " Left Join Txn_WHR_Detail WD on TPC.WHRNo = WD.WHRNo "
tmp = tmp & " Inner Join Mst_Bank MBank On TPC.BankID = MBank.BankID "
tmp = tmp & " Inner Join Mst_Branch MBranch On TPC.BranchID = MBranch.BranchID "
tmp = tmp & " Left  Join Mst_Client MC On TPC.ClientID = MC.ClientID "
tmp = tmp & " Left Join Txn_ImportNcdexLotPleadge MI On TPC.ISINID = MI.Code "
tmp = tmp & " Inner Join Mst_Location ML On TPR.SM_Prefix = ML.SM_Prefix "
tmp = tmp & " Left Join Txn_PledgeRelease_GSL TPRG On TPR.Txn_PledgeRelease_ID = TPRG.Txn_PledgeRelease_ID And TPR.SM_Prefix = TPRG.SM_Prefix "

If Gen_DBType = "MDB" Then
   wc = " Where TPR.PledgeReleaseDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   wc = " Where TPR.PledgeReleaseDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSPreFix.Text <> "" Then
   wc = GenWc(wc, CmbSPreFix.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbSBankName.Text <> "" Then
   wc = GenWc(wc, CmbSBankName.Text, "MBank.BankName", "S", "N")
End If

If CmbSBranchName.Text <> "" Then
   wc = GenWc(wc, CmbSBranchName.Text, "MBranch.BranchName", "S", "N")
End If

If CmbSearchExchange.Text <> "" Then
   wc = GenWc(wc, CmbSearchExchange.Text, "ME.ExchangeType", "S", "N")
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
      wc = GenWc(wc, TxtSearchPledge, "MI.Code", "S", "L")
   End If
   
End If


If TxtSClientName <> "" Then
   wc = GenWc(wc, TxtSClientName, "MC.ClientName", "S", "L")
End If

tmp = tmp & " " & wc & " "

tmp = tmp & " Group By TPR.Txn_PledgeRelease_ID,ML.LocationName,TPR.ReleaseOrderNo,TPR.PledgeReleaseDate,TPC.Txn_PledgeCreation_ID,ME.ExchangeType,WD.WHRNo,WD.CMPID,MI.Code,MBank.BankName,MBranch.BranchName,MC.ClientName,TPR.Remarks,TPR.G_UID "

tmp = tmp & " Order By TPR.Txn_PledgeRelease_ID,ML.LocationName  "

Call TmpTable


Call Grid_Head1
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    MSHFlexGrid1.TextMatrix(iu, 0) = TmpRs!Txn_PledgeRelease_ID
    MSHFlexGrid1.TextMatrix(iu, 1) = TmpRs!Locationname
    MSHFlexGrid1.TextMatrix(iu, 2) = IIf(IsNull(TmpRs!ReleaseOrderNo), "", TmpRs!ReleaseOrderNo)
    MSHFlexGrid1.TextMatrix(iu, 3) = Format(TmpRs!PledgeReleaseDate, "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(iu, 4) = TmpRs!Txn_PledgeCreation_ID
    MSHFlexGrid1.TextMatrix(iu, 5) = TmpRs!ExchangeType
    If IsNull(TmpRs!WHRNo) <> True Then
       MSHFlexGrid1.TextMatrix(iu, 6) = TmpRs!CMPID & "-" & ((TmpRs!WHRNo) - TmpRs!CMPID * 1000000)
    Else
       MSHFlexGrid1.TextMatrix(iu, 6) = IIf(IsNull(TmpRs!Code), "", TmpRs!Code)
    End If
    MSHFlexGrid1.TextMatrix(iu, 7) = TmpRs!BankName
    MSHFlexGrid1.TextMatrix(iu, 8) = TmpRs!BranchName
    MSHFlexGrid1.TextMatrix(iu, 9) = TmpRs!ClientName
    MSHFlexGrid1.TextMatrix(iu, 10) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
    MSHFlexGrid1.TextMatrix(iu, 11) = IIf(IsNull(TmpRs!ReleaseBags), 0, TmpRs!ReleaseBags)
    MSHFlexGrid1.TextMatrix(iu, 12) = IIf(IsNull(TmpRs!ReleaseQty), 0, TmpRs!ReleaseQty)
    MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs!G_UID
    TmpRs.MoveNext
Next


Label24.Refresh
Label24.Caption = TmpRs.RecordCount


End Sub

Private Sub CmdUpdate_Click()
Dim strSplitString() As String
If TxtReleaseNoOfBags = "" Then
   MsgBox "Please provide the No of Bags to be released!!"
   TxtReleaseNoOfBags.SetFocus
   Exit Sub
End If
If TxtReleaseQuantity = "" Then
   MsgBox "Please provide Quantity to be released!!!"
   TxtReleaseQuantity.SetFocus
   Exit Sub
End If


If Val(TxtUnpledgeNoOfBags) = TxtReleaseNoOfBags Then
   If Val(TxtUnpledgeWeight) <> Val(TxtReleaseQuantity) Then
      MsgBox "When all the bags are released, all the quantity are also required to be released !!!! "
      TxtReleaseNoOfBags.SetFocus
      Exit Sub
   End If
End If

If Val(TxtUnpledgeWeight) = TxtReleaseQuantity Then
   If Val(TxtUnpledgeNoOfBags) <> Val(TxtReleaseNoOfBags) Then
      MsgBox "When all the quantity are released, all the  bags are also required to be released !!!! "
      TxtReleaseQuantity.SetFocus
      Exit Sub
   End If
End If

If TxtReleaseNoOfBags = 0 And TxtReleaseQuantity = 0 Then

Else
    If Val(TxtUnpledgeNoOfBags) <> Val(TxtReleaseNoOfBags) Then
       tmp = Validate_ComBags(mCommodityID, Val(TxtReleaseNoOfBags), Val(TxtReleaseQuantity))
       If Len(tmp) > 1 Then
          MsgBox tmp
          TxtReleaseQuantity.SetFocus
          Exit Sub
       End If
    End If
End If
MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))

If Val(TxtReleaseNoOfBags) = 0 And Val(TxtReleaseQuantity) = 0 Then
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
Else
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
End If

MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = TxtCmp
mCMPID = GetCMPID(TxtCmp)
strSplitString = Split(TxtGSL, "~")
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) = Trim(strSplitString(0))
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3) = Trim(strSplitString(1))
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4) = Trim(strSplitString(2))
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5) = Val(TxtUnpledgeNoOfBags) - Val(TxtReleaseNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6) = Format(Val(TxtUnpledgeWeight) - Val(TxtReleaseQuantity), "########0.000")
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7) = Val(TxtReleaseNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8) = Val(TxtReleaseQuantity)


MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))

TxtCmp = ""
TxtGSL = ""
TxtUnpledgeNoOfBags = ""
TxtUnpledgeWeight = ""
TxtReleaseNoOfBags = ""
TxtReleaseQuantity = ""
TxtReleaseNoOfBags.Locked = True
TxtReleaseQuantity.Locked = True
CmdUpdate.Enabled = False

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   'tmp = "Select * from Mst_GSL where sm_prefix = '" & TxtPreFix & "' And GSLID = " & MSHFlexGrid2.TextMatrix(2, 9) & " "
   For I = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
          RsMst_GSL.MoveFirst
          RsMst_GSL.Filter = "Mst_GSL_ID = " & MSHFlexGrid2.TextMatrix(I, 9) & " And SM_Prefix = '" & mSM_Prefix & "' And ExchangeTypeID = " & mExchangeTypeID & " "
          If RsMst_GSL.RecordCount > 0 Then
             RsMst_GSL!PledgeBags = Val(IIf(IsNull(RsMst_GSL!PledgeBags), "0", RsMst_GSL!PledgeBags)) + Val(MSHFlexGrid2.TextMatrix(I, 7))
             RsMst_GSL!PledgeWeight = Val(IIf(IsNull(RsMst_GSL!PledgeWeight), "0", RsMst_GSL!PledgeWeight)) + Val(MSHFlexGrid2.TextMatrix(I, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
             RsMst_GSL.Update
          End If
          Call TableTxn_PledgeCreation_GSL(" Where Txn_PledgeCreation_ID = " & TxtPledgeID)
          'RsTxn_PledgeCreation_GSL.MoveFirst
          RsTxn_PledgeCreation_GSL.Find "Mst_GSL_ID =  " & MSHFlexGrid2.TextMatrix(I, 9)
          If Not RsTxn_PledgeCreation_GSL.EOF Then
             RsTxn_PledgeCreation_GSL!PledgeNoOfBags = Val(RsTxn_PledgeCreation_GSL!PledgeNoOfBags) + Val(MSHFlexGrid2.TextMatrix(I, 7))
             RsTxn_PledgeCreation_GSL!PledgeQuantity = Val(RsTxn_PledgeCreation_GSL!PledgeQuantity) + Val(MSHFlexGrid2.TextMatrix(I, 8))
          End If
          RsTxn_PledgeCreation_GSL.Update
          RsTxn_PledgeRelease_GSL.Requery
       RsMst_GSL.Filter = "Mst_GSL_ID <> Null "
   Next
 
   
   tmp = "Delete From Txn_PledgeRelease_GSL where Txn_PledgeRelease_ID = " & TxtReleaseID
   Call TmpTable
   RsTxn_PledgeRelease.Delete
   RsTxn_PledgeRelease.Update
   
   If RsTxn_PledgeRelease.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_PledgeRelease.MoveNext
   If RsTxn_PledgeRelease.EOF() Then
      RsTxn_PledgeRelease.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_PledgeRelease.CancelUpdate
End Sub

Private Sub ExitCmd_Click()
Unload Me
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
SSTab1.Tab = 0
Call Grid_Head1
Call CmdClear_Click
Call TableTxn_PledgeRelease
If RsTxn_PledgeRelease.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call TableTxn_PledgeCreation
Call Indata

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
MSHFlexGrid2.TextMatrix(1, 5) = "Pledged  No Of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Pledged  Qty"
MSHFlexGrid2.TextMatrix(1, 7) = "Released No Of Bags"
MSHFlexGrid2.TextMatrix(1, 8) = "Released Qty"
MSHFlexGrid2.TextMatrix(1, 9) = "ReleaseGSLID"


MSHFlexGrid2.ColWidth(0) = 400
MSHFlexGrid2.ColWidth(1) = 3200
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 1000
MSHFlexGrid2.ColWidth(5) = 1600
MSHFlexGrid2.ColWidth(6) = 1600
MSHFlexGrid2.ColWidth(7) = 1400
MSHFlexGrid2.ColWidth(8) = 1400
MSHFlexGrid2.ColWidth(9) = 0

MSHFlexGrid2.RowHeight(1) = 400
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
 TxtCmp = ""
TxtGSL = ""
TxtUnpledgeNoOfBags = ""
TxtUnpledgeWeight = ""
TxtReleaseNoOfBags = ""
TxtReleaseQuantity = ""
TxtReleaseNoOfBags.Locked = True
TxtReleaseQuantity.Locked = True
CmdUpdate.Enabled = False

 
If MSHFlexGrid2.Rows > 3 Then
   If MSHFlexGrid2.RowSel <> "0" Then
      If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then Exit Sub
      TxtReleaseNoOfBags.Locked = False
      TxtReleaseQuantity.Locked = False
      TxtCmp = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1)
      TxtGSL = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) & "~" & MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3) & "~" & MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4)
      TxtUnpledgeNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
      TxtUnpledgeWeight = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
      If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7)) = 0 Then
         TxtReleaseNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5))
      Else
         TxtReleaseNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
      End If
      If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8)) = 0 Then
         TxtReleaseQuantity = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6))
      Else
         TxtReleaseQuantity = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
      End If
       
      CmdUpdate.Enabled = True
   End If
End If
End Sub

Private Sub ReleaseDate_LostFocus()
If ReleaseDate.Value > Date Then
   MsgBox "Future Release Date is not allowed !!!! "
   ReleaseDate.SetFocus
End If


End Sub

Private Sub SaveCmd_Click()

If ReleaseDate.Value > Date Then
   MsgBox "Release Date should not be a future date!!!"
   ReleaseDate.SetFocus
   Exit Sub
End If

If ReleaseDate.Value < PledgeDate.Value Then
   MsgBox "Release Date must be greater than Pledge Date !!!"
   ReleaseDate.SetFocus
   Exit Sub
End If

If TxtReleaseOrderNo = "" Then
   MsgBox "Blank Release Order not allowed!!!"
   TxtReleaseOrderNo.SetFocus
   Exit Sub
End If

If CmbPreFix.Text = "" Then
   MsgBox "Blank Location ID not allowed!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If

If TxtPledgeID = "" Then
   MsgBox "Blank Pledge ID not allowed!!!"
   TxtPledgeID.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "No data is available for release!!!"
   TxtPledgeID.SetFocus
   Exit Sub
End If

If ChkSelection() = False Then
   MsgBox "Please select atleast one lot!!"
   MSHFlexGrid2.SetFocus
   Exit Sub
End If

If mNCMLFinance = 1 Then GoTo E

If mExchangeTypeID = 1 Then
'   If Val(MSHFlexGrid2.TextMatrix(0, 7)) > mROBags Then
'
'      'If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "head cmg" Then
'
''      If LCase(Gen_UserRole) <> "power" Then
''         MsgBox " Total Release bags must be less than or equal to RO Bags " & mROBags & " from CM-PRo!!" 'MsgBox "No relase made in CM Pro for this WHR . Can't Release!!!"
''         Exit Sub
''      Else
'       MsgBox "Total Release bags must be less than or equal to RO Bags " & mROBags & " from CM-PRo. "
''         If ans = vbNo Then
'            Exit Sub
''         End If
''      End If
'   End If
'
'   If Val(MSHFlexGrid2.TextMatrix(0, 8)) > mROQty Then
''      MsgBox " Total Release Qty must be less than or equal to RO Qty  " & mROQty & " from CM-PRo!!"
''      Exit Sub
'      If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "head cmg" Then
'         MsgBox " Total Release Qty must be less than or equal to RO Qty  " & mROQty & " from CM-PRo!!" 'MsgBox "No relase made in CM Pro for this WHR . Can't Release!!!"
'         Exit Sub
'      Else
'         ans = MsgBox(" Total Release Qty must be less than or equal to RO Qty  " & mROQty & " from CM-PRo. Still Release ?", vbYesNo)
'         If ans = vbNo Then
'            Exit Sub
'         End If
'      End If
'   End If

   If Format(PledgeDate.Value, "yyyy-mm-dd") >= "2012-01-31" Then
   
      mReleaseBags = 0
      mReleaseQty = 0
        
      tmp = "Select sum(ReleasedNoOfBags)ReleasedNoOfBags,sum(ReleasedQuantity)ReleasedQuantity from Txn_PledgeRelease_GSL TPRG"
      tmp = tmp & " Where Txn_PledgeCreation_ID=" & Val(TxtPledgeID) & " And SM_Prefix='" & mSM_Prefix & "'"
      
      Call Tmp2Table
      
      If TmpRs2.RecordCount > 0 Then
         mReleaseBags = IIf(IsNull(TmpRs2!ReleasedNoOfBags), 0, TmpRs2!ReleasedNoOfBags)
         mReleaseQty = IIf(IsNull(TmpRs2!ReleasedQuantity), 0, TmpRs2!ReleasedQuantity)
      End If
      mReleaseBags = mReleaseBags + Val(MSHFlexGrid2.TextMatrix(0, 7))
      mReleaseQty = mReleaseQty + Val(MSHFlexGrid2.TextMatrix(0, 8))
      
      If mReleaseBags > mROBags Then
   '      MsgBox " Total Release bags must be less than or equal to RO Bags!!"
         If mSCMFlag = 0 Then
            MsgBox "Total Release bags must be less than or equal to RO Bags " & mROBags & " from CM-PRo. "
         Else
            MsgBox "Total Release bags must be less than or equal to DO Bags " & mROBags & " from Store-Man. "
         End If
         Exit Sub
      End If
       
      If mReleaseQty > mROQty Then
   '      MsgBox " Total Release Qty must be less than or equal to RO Qty!!"
         If mSCMFlag = 0 Then
            MsgBox " Total Release Qty must be less than or equal to RO Qty  " & mROQty & " from CM-PRo!!" 'MsgBox "No relase made in CM Pro for this WHR . Can't Release!!!"
         Else
            MsgBox " Total Release Qty must be less than or equal to DO Qty  " & mROQty & " from Store-Man!!" 'MsgBox "No relase made in CM Pro for this WHR . Can't Release!!!"
         End If
         Exit Sub
      End If
   End If

End If

E:

If Edit_Flg = "A" Then
   RsTxn_PledgeRelease.AddNew
   RsTxn_PledgeRelease!PledgeReleaseID = GetPledgeReleaseID
   RsTxn_PledgeRelease!G_UID = GetGUID
   mPRG_UID = RsTxn_PledgeRelease!G_UID
   'TxtReleaseID = RsTxn_PledgeRelease!PledgeReleaseID
Else

End If

RsTxn_PledgeRelease!PledgeReleaseDate = ReleaseDate.Value
RsTxn_PledgeRelease!Txn_PledgeCreation_ID = TxtPledgeID
RsTxn_PledgeRelease!ReleaseOrderNo = IIf(IsNull(TxtReleaseOrderNo), "", TxtReleaseOrderNo)
RsTxn_PledgeRelease!Remarks = IIf(IsNull(TxtRemarks), "", TxtRemarks)
RsTxn_PledgeRelease!SM_Prefix = mSM_Prefix
RsTxn_PledgeRelease!SCMFlag = mSCMFlag

If IsNull(RsTxn_PledgeRelease!CreatedBy) = True Or RsTxn_PledgeRelease!CreatedBy = "" Then
   RsTxn_PledgeRelease!CreatedBy = Gen_UserLoginID
   RsTxn_PledgeRelease!CreatedDate = Now
Else
   RsTxn_PledgeRelease!UpdatedBy = Gen_UserLoginID
   RsTxn_PledgeRelease!UpdatedDate = Now
End If

RsTxn_PledgeRelease.Update

tmp = "Select * from Txn_PledgeRelease Where G_UID = '" & mPRG_UID & "'"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   TxtReleaseID = TmpRs1!Txn_PledgeRelease_ID
End If

Call SavePledgeRelease_GSL
Call Gen_Email


End Sub
Private Sub SavePledgeRelease_GSL()
Dim strSplitString() As String
Dim mPledegeNoNew As Variant
Call TableTxn_PledgeRelease_GSL("Where sm_prefix = '" & mSM_Prefix & "' And Txn_PledgeRelease_ID = " & TxtReleaseID)
Call TableMst_GSL(" Where sm_prefix = '" & mSM_Prefix & "'")
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(I, 0) = "*" Then
       RsTxn_PledgeRelease_GSL.AddNew
       RsTxn_PledgeRelease_GSL!G_UID = GetGUID
       RsTxn_PledgeRelease_GSL!PledgeReleaseGSLID = GetPledgeReleaseGSLID
       RsTxn_PledgeRelease_GSL!Txn_PledgeRelease_ID = TxtReleaseID
       RsTxn_PledgeRelease_GSL!Txn_PledgeCreation_ID = TxtPledgeID
       RsTxn_PledgeRelease_GSL!Mst_GSL_ID = MSHFlexGrid2.TextMatrix(I, 9)
       RsTxn_PledgeRelease_GSL!ReleasedNoOfBags = MSHFlexGrid2.TextMatrix(I, 7)
       RsTxn_PledgeRelease_GSL!ReleasedQuantity = MSHFlexGrid2.TextMatrix(I, 8)
       RsTxn_PledgeRelease_GSL!SM_Prefix = mSM_Prefix

       RsTxn_PledgeRelease_GSL.Update
       RsTxn_PledgeRelease_GSL.Requery
       If RsMst_GSL.RecordCount > 0 Then
          RsMst_GSL.MoveFirst
       End If
       RsMst_GSL.Find "Mst_GSL_ID = " & MSHFlexGrid2.TextMatrix(I, 9)
       If Not RsMst_GSL.EOF Then
'          mPledgeNo = RsMst_GSL!PledgeNo
'          strSplitString = Split(mPledgeNo, "~")
'          For k = 0 To UBound(strSplitString)
'              If TxtPledgeID <> strSplitString(k) Then
'                 mPledegeNoNew = mPledegeNoNew & " ~ " & strSplitString(k)
'              End If
'          Next
          RsMst_GSL!PledgeBags = Val(IIf(IsNull(RsMst_GSL!PledgeBags), "0", RsMst_GSL!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(I, 7))
          RsMst_GSL!PledgeWeight = Val(IIf(IsNull(RsMst_GSL!PledgeWeight), "0", RsMst_GSL!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(I, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'          If RsMst_GSL!PledgeBags = 0 And RsMst_GSL!PledgeWeight = 0 Then
'             RsMst_GSL!PledgeNo = ""
'          Else
'             RsMst_GSL!PledgeNo = IIf(IsNull(mPledegeNoNew), "", mPledegeNoNew)
'          End If
          RsMst_GSL.Update
          If IsNull(RsMst_GSL!FGSLID_Mst_GSL_ID) = False Then
             Call UpdateFGSLID(I, RsMst_GSL!FGSLID_Mst_GSL_ID)
          End If
       End If
       RsTxn_PledgeCreation_GSL.MoveFirst
       RsTxn_PledgeCreation_GSL.Find "Mst_GSL_ID =  " & MSHFlexGrid2.TextMatrix(I, 9)
       If Not RsTxn_PledgeCreation_GSL.EOF Then
          RsTxn_PledgeCreation_GSL!PledgeNoOfBags = Val(RsTxn_PledgeCreation_GSL!PledgeNoOfBags) - Val(MSHFlexGrid2.TextMatrix(I, 7))
          RsTxn_PledgeCreation_GSL!PledgeQuantity = Val(RsTxn_PledgeCreation_GSL!PledgeQuantity) - Val(MSHFlexGrid2.TextMatrix(I, 8))
       End If
       RsTxn_PledgeCreation_GSL.Update
       RsTxn_PledgeRelease_GSL.Requery
    End If
Next
RsTxn_PledgeRelease_GSL.Close
RsTxn_PledgeCreation_GSL.Close
cn.Close


Dim pBag, PQty As Variant
If mExchangeTypeID = 2 Then
   Call Indata
   Exit Sub
End If
If mExchangeTypeID = 1 Then
'   tmp = "SELECT Max(Txn_WHR_Detail.WHRNoOfBags) AS MaxOfWHRNoOfBags, Max(Txn_WHR_Detail.WHRQuantity)"
'   tmp = tmp & " AS MaxOfWHRQuantity, "
'   tmp = tmp & " max(Txn_WHR_Detail.DeliveredQuantity)AS MaxOfDeliveredQty,max(Txn_WHR_Detail.DeliveredNoOfBags)As MaxOfDeliveredOfBags, "
'   tmp = tmp & " Sum(Txn_PledgeRelease_GSL.ReleasedNoOfBags) AS SumOfReleasedNoOfBags,"
'   tmp = tmp & " Sum(Txn_PledgeRelease_GSL.ReleasedQuantity) As SumOfReleasedQuantity "
'   tmp = tmp & " FROM ((Txn_WHR_Detail INNER JOIN Txn_PledgeCreation ON"
'   tmp = tmp & " Txn_WHR_Detail.PledgeID = Txn_PledgeCreation.PledgeID)"
'   tmp = tmp & " INNER JOIN Txn_PledgeRelease ON Txn_PledgeCreation.PledgeID ="
'   tmp = tmp & " Txn_PledgeRelease.PledgeID) INNER JOIN Txn_PledgeRelease_GSL ON"
'   tmp = tmp & " Txn_PledgeRelease.PledgeReleaseID = Txn_PledgeRelease_GSL.PledgeReleaseID "
'   tmp = tmp & " WHERE Txn_PledgeCreation.sm_prefix = '" & CmbPreFix.Text & "' And (((Txn_PledgeCreation.WHRNo)=" & Val(TxtWHRNo) & "))"



   tmp = "SELECT Max(Txn_WHR_Detail.WHRNoOfBags) AS MaxOfWHRNoOfBags, Max(Txn_WHR_Detail.WHRQuantity)"
   tmp = tmp & " AS MaxOfWHRQuantity, "
   tmp = tmp & " max(Txn_WHR_Detail.DeliveredQuantity)AS MaxOfDeliveredQty,max(Txn_WHR_Detail.DeliveredNoOfBags)As MaxOfDeliveredOfBags, "
   tmp = tmp & " Sum(Txn_PledgeRelease_GSL.ReleasedNoOfBags) AS SumOfReleasedNoOfBags,"
   tmp = tmp & " Sum(Txn_PledgeRelease_GSL.ReleasedQuantity) As SumOfReleasedQuantity "
   tmp = tmp & " FROM ((Txn_WHR_Detail INNER JOIN Txn_PledgeCreation ON"
   tmp = tmp & " Txn_WHR_Detail.Txn_PledgeCreation_ID = Txn_PledgeCreation.Txn_PledgeCreation_ID And Txn_WHR_Detail.SM_Prefix = Txn_PledgeCreation.SM_Prefix)"
   tmp = tmp & " INNER JOIN Txn_PledgeRelease ON Txn_PledgeCreation.Txn_PledgeCreation_ID ="
   tmp = tmp & " Txn_PledgeRelease.Txn_PledgeCreation_ID And Txn_PledgeCreation.SM_Prefix = Txn_PledgeRelease.SM_Prefix) INNER JOIN Txn_PledgeRelease_GSL ON"
   tmp = tmp & " Txn_PledgeRelease.Txn_PledgeRelease_ID = Txn_PledgeRelease_GSL.Txn_PledgeRelease_ID And Txn_PledgeRelease.SM_Prefix =    Txn_PledgeRelease_GSL.SM_Prefix"
   tmp = tmp & " WHERE Txn_PledgeCreation.sm_prefix = '" & mSM_Prefix & "' And (((Txn_PledgeCreation.WHRNo)=" & Val(TxtWHRNo) & "))"



Else
   tmp = "" ' tmp & " WHERE (((Txn_PledgeCreation.ISINID)= '" & (TxtWHRNo) & "'))"
End If

Call TmpTable

pBag = IIf(IsNull(TmpRs!MaxOfWHRNoOfBags), 0, TmpRs!MaxOfWHRNoOfBags)
PQty = IIf(IsNull(TmpRs!MaxOfWHRQuantity), 0, TmpRs!MaxOfWHRQuantity)

'pBag = pBag - IIf(IsNull(TmpRs!MaxOfDeliveredOfBags), 0, TmpRs!MaxOfDeliveredOfBags)
'PQty = PQty - IIf(IsNull(TmpRs!MaxOfDeliveredQty), 0, TmpRs!MaxOfDeliveredQty)

pBag = pBag - IIf(IsNull(TmpRs!SumOfReleasedNoOfBags), 0, TmpRs!SumOfReleasedNoOfBags)
PQty = PQty - IIf(IsNull(TmpRs!SumOfReleasedQuantity), 0, TmpRs!SumOfReleasedQuantity)


'tmp = "SELECT Sum(Txn_PledgeRelease_GSL.ReleasedNoOfBags) AS SumOfReleasedNoOfBags, "
'tmp = tmp & " Sum(Txn_PledgeRelease_GSL.ReleasedQuantity) AS SumOfReleasedQuantity"
'tmp = tmp & " FROM (Txn_PledgeCreation INNER JOIN Txn_PledgeRelease ON "
'tmp = tmp & " Txn_PledgeCreation.PledgeID = Txn_PledgeRelease.PledgeID) INNER JOIN Txn_PledgeRelease_GSL "
'tmp = tmp & " ON Txn_PledgeRelease.PledgeReleaseID = Txn_PledgeRelease_GSL.PledgeReleaseID"
'tmp = tmp & " WHERE (((Txn_PledgeCreation.WHRNo)=" & Val(TxtWHRNo) & "));"

'Commented on 23 March 2015 by Parag for solving SCM Pledge issue also above code for max deliveredofbags is commented informed Mayur Sir on same day morning
'tmp = "SELECT (Txn_PledgeRelease_GSL.ReleasedNoOfBags) AS SumOfReleasedNoOfBags,"
'tmp = tmp & " (Txn_PledgeRelease_GSL.ReleasedQuantity) As SumOfReleasedQuantity"
'tmp = tmp & " FROM Txn_PledgeRelease_GSL "
'
'tmp = tmp & " WHERE Txn_PledgeRelease_GSL.sm_prefix = '" & mSM_Prefix & "' And (((Txn_PledgeRelease_GSL.PledgeReleaseID)=" & Val(TxtReleaseID) & "));"
'
'Call TmpTable
'
'
'For i = 1 To TmpRs.RecordCount
'    pBag = pBag - TmpRs!SumOfReleasedNoOfBags
'    PQty = PQty - TmpRs!SumOfReleasedQuantity
'   TmpRs.MoveNext
'Next

If pBag <= 0 Or PQty <= 0 Then
   Call TableTxn_WHRDetail(" Where WHRNo = " & TxtWHRNo)
   If Not RsTxn_WHRDetail.EOF Then
      RsTxn_WHRDetail!Txn_PledgeCreation_ID = Null
      RsTxn_WHRDetail.Update
      RsTxn_WHRDetail.Close
   End If
End If

TmpRs.Close
Call Indata

End Sub
Private Sub UpdateFGSLID(x As Variant, mFGSLID As Variant)


k: tmp = "Select * from Mst_GSL Where Mst_GSL_ID  = " & mFGSLID & ""
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then

    If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
       TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
       TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
       TmpRs1.Update
    Else
       mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
       GoTo k
    End If
End If



'mNewSM_Prefix = ""
'If LCase(mSM_Prefix) = "bkner" Then
'   mNewSM_Prefix = "bknrp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "ujjan" Then
'   mNewSM_Prefix = "ujjap"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "dewas" Then
'   mNewSM_Prefix = "dewsp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "vdisa" Then
'   mNewSM_Prefix = "vdshp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "adlbd" Then
'   mNewSM_Prefix = "adlbp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "akola" Then
'   mNewSM_Prefix = "aklap"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "sgngr" Then
'   mNewSM_Prefix = "sngrp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "nzmbd" Then
'   mNewSM_Prefix = "nzmdp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "unjha" Then
'   mNewSM_Prefix = "unjap"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "sngli" Then
'   mNewSM_Prefix = "sglip"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "guntr" Then
'   mNewSM_Prefix = "gntrp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "dvngr" Then
'   mNewSM_Prefix = "dvgrp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "jdhpr" Then
'   mNewSM_Prefix = "jdhrp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "hghat" Then
'   mNewSM_Prefix = "hngtp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "kadi" Then
'   mNewSM_Prefix = "kadip"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "latur" Then
'   mNewSM_Prefix = "lturp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "kota" Then
'   mNewSM_Prefix = "ktap"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "rsngr" Then
'   mNewSM_Prefix = "rsgrp"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'ElseIf LCase(mSM_Prefix) = "rsgrp" Then
'   mNewSM_Prefix = "rsngr"
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "'"
'   Call Tmp1Table
'Else
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "'"
'   Call Tmp1Table
'End If
'
'If mNewSM_Prefix <> "" Then
'   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "'"
'   Call Tmp2Table
'End If
'
''If RsMst_GSL.RecordCount > 0 Then
''   RsMst_GSL.MoveFirst
''End If
'
''RsMst_GSL.Find "GSLID = " & MSHFlexGrid2.TextMatrix(x, 9)
'
''If Not RsMst_GSL.EOF Then
''   mFGSLID = RsMst_GSL!FGSLID
'If TmpRs1.RecordCount > 0 Then
'   If LCase(mSM_Prefix) <> "bkner" And LCase(mSM_Prefix) <> "ujjan" And LCase(mSM_Prefix) <> "dewas" And LCase(mSM_Prefix) <> "vdisa" And LCase(mSM_Prefix) <> "adlbd" And LCase(mSM_Prefix) <> "akola" And LCase(mSM_Prefix) <> "sgngr" And LCase(mSM_Prefix) <> "nzmbd" And LCase(mSM_Prefix) <> "unjha" And LCase(mSM_Prefix) <> "sngli" And LCase(mSM_Prefix) <> "guntr" And LCase(mSM_Prefix) <> "dvngr" And LCase(mSM_Prefix) <> "jdhpr" And LCase(mSM_Prefix) <> "hghat" And LCase(mSM_Prefix) <> "kadi" And LCase(mSM_Prefix) <> "latur" And LCase(mSM_Prefix) <> "kota" And LCase(mSM_Prefix) <> "rsngr" And LCase(mSM_Prefix) <> "rsgrp" Then
'l:   TmpRs1.MoveFirst
'     TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'     If TmpRs1.EOF = False Then
'        If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'           If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'              TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'              TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'              TmpRs1.Update
'           Else
'              mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'              GoTo l
'           End If
'        End If
'     'Else
'     End If
'   End If
''Bikaner
'   If LCase(mSM_Prefix) = "bkner" Then
'      If TmpRs1.RecordCount > 0 Then
'k:         TmpRs1.MoveFirst
'         TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'         If TmpRs1.EOF = False Then
'            If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'               If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                  TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                  TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                  TmpRs1.Update
'               Else
'                   If TmpRs2.RecordCount > 0 Then
'                      TmpRs2.MoveFirst
'                      TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                      If TmpRs2.EOF = False Then
'                         If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                            If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                               TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                               TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                               TmpRs2.Update
'                            Else
'                               mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                               GoTo k
'                            End If
'                         End If
'                      End If
'                   End If
'                   mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                   GoTo k
'               End If
'            End If
'         Else
'            If TmpRs2.RecordCount > 0 Then
'               TmpRs2.MoveFirst
'               TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'               If TmpRs2.EOF = False Then
'                  If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                     If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                        TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                        TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                        TmpRs2.Update
'                     Else
'                        mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                        GoTo k
'                     End If
'                  End If
'               End If
'            End If
'         End If
'      End If
'
'
'
'
'   End If
''Ujjain
'           If LCase(mSM_Prefix) = "ujjan" Then
'              If TmpRs1.RecordCount > 0 Then
'u:               TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
''                          mFGSLID = TmpRs1!FGSLID
''                          GoTo u
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo u
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo u
'                       End If
'                    End If
'              Else
'                 If TmpRs2.RecordCount > 0 Then
'                    TmpRs2.MoveFirst
'                    TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                    If TmpRs2.EOF = False Then
'                       If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                          If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                             TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                             TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                             TmpRs2.Update
'                          Else
'                             mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                             GoTo u
'                          End If
'                       End If
'                    End If
'                 End If
'              End If
'           End If
'         End If
''Dewas
'           If LCase(mSM_Prefix) = "dewas" Then
'              If TmpRs1.RecordCount > 0 Then
'd:               TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo d
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo d
'                       End If
'                    End If
'                 Else
'                       If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo d
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'           End If
'
''VIDISHA
'           If LCase(mSM_Prefix) = "vdisa" Then
'              If TmpRs1.RecordCount > 0 Then
'v:               TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo v
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo v
'                       End If
'
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo v
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'           End If
'
''ADLABAD
'           If LCase(mSM_Prefix) = "adlbd" Then
'              If TmpRs1.RecordCount > 0 Then
'a:               TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo a
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo a
'                       End If
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo a
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'           End If
'
''AKOLA
'           If LCase(mSM_Prefix) = "akola" Then
'              If TmpRs1.RecordCount > 0 Then
'ak:              TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo ak
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo ak
'                       End If
'                    End If
'                 Else
'                         If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo ak
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'           End If
'
''Sriganganagar
'           If LCase(mSM_Prefix) = "sgngr" Then
'              If TmpRs1.RecordCount > 0 Then
'S:               TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo S
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo S
'                       End If
'                    End If
'                 Else
'
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo S
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'           End If
'
''Nizamabad
'           If LCase(mSM_Prefix) = "nzmbd" Then
'              If TmpRs1.RecordCount > 0 Then
'N:               TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo N
'                                   End If
'                                End If
'                             End If
'                          End If
'
'
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo N
'                       End If
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo N
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'           End If
'
''Unjha
'           If LCase(mSM_Prefix) = "unjha" Then
'              If TmpRs1.RecordCount > 0 Then
'Un:               TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo Un
'                                   End If
'                                End If
'                             End If
'                          End If
'
'
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo Un
'                       End If
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo Un
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'           End If
''Sangli
'           If LCase(mSM_Prefix) = "sngli" Then
'              If TmpRs1.RecordCount > 0 Then
'Sn:              TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo Sn
'                                   End If
'                                End If
'                             End If
'                          End If
'
'
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo Sn
'                       End If
'                    End If
'                 Else
'
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo Sn
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'           End If
''Guntur
'           If LCase(mSM_Prefix) = "guntr" Then
'              If TmpRs1.RecordCount > 0 Then
'Gun:             TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo Gun
'                                   End If
'                                End If
'                             End If
'                          End If
'
'
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo Gun
'                       End If
'                     End If
'                 Else
'
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo Gun
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'           End If
''Davangeere
'           If LCase(mSM_Prefix) = "dvngr" Then
'              If TmpRs1.RecordCount > 0 Then
'DVN:             TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo DVN
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo DVN
'                       End If
'                    End If
'                 Else
'
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo DVN
'                                   End If
'                                End If
'                             End If
'                          End If
'
'                 End If
'              End If
'           End If
'
''JOdhpur
'           If LCase(mSM_Prefix) = "jdhpr" Then
'              If TmpRs1.RecordCount > 0 Then
'JD:              TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo JD
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo JD
'                       End If
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo JD
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'            End If
'
''Hinganghat
'         If LCase(mSM_Prefix) = "hghat" Then
'              If TmpRs1.RecordCount > 0 Then
'HG:              TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo HG
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo HG
'                       End If
'                    End If
'                 Else
'
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo HG
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'         End If
'
''          End If
''Kadi
'          If LCase(mSM_Prefix) = "kadi" Then
'              If TmpRs1.RecordCount > 0 Then
'KD:              TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo KD
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo KD
'                       End If
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo KD
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'         End If
'
''Latur
'          If LCase(mSM_Prefix) = "latur" Then
'              If TmpRs1.RecordCount > 0 Then
'LTR:              TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo LTR
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo LTR
'                       End If
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo LTR
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'         End If
'
'
''KOTA P
'          If LCase(mSM_Prefix) = "kota" Then
'              If TmpRs1.RecordCount > 0 Then
'KTA:              TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo KTA
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo KTA
'                       End If
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo KTA
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'         End If
'
'
''Raisingnagar P
'          If LCase(mSM_Prefix) = "rsngr" Then
'              If TmpRs1.RecordCount > 0 Then
'RASP:              TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo RASP
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo RASP
'                       End If
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo RASP
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'         End If
'     '   End If
''     End If
''Raisingnagar
'          If LCase(mSM_Prefix) = "rsgrp" Then
'              If TmpRs1.RecordCount > 0 Then
'RASNG:              TmpRs1.MoveFirst
'                 TmpRs1.Find "Mst_GSL_ID = " & mFGSLID
'                 If TmpRs1.EOF = False Then
'                    If InStr(1, TmpRs1!PledgeNo, TxtPledgeID) > 0 Then
'                       If IsNull(TmpRs1!FGSLID_Mst_GSL_ID) = True Then
'                          TmpRs1!PledgeBags = Val(IIf(IsNull(TmpRs1!PledgeBags), "0", TmpRs1!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                          TmpRs1!PledgeWeight = Val(IIf(IsNull(TmpRs1!PledgeWeight), "0", TmpRs1!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                          TmpRs1.Update
'                       Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo RASNG
'                                   End If
'                                End If
'                             End If
'                          End If
'                          mFGSLID = TmpRs1!FGSLID_Mst_GSL_ID
'                          GoTo RASNG
'                       End If
'                    End If
'                 Else
'                          If TmpRs2.RecordCount > 0 Then
'                             TmpRs2.MoveFirst
'                             TmpRs2.Find "Mst_GSL_ID = " & mFGSLID
'                             If TmpRs2.EOF = False Then
'                                If InStr(1, TmpRs2!PledgeNo, TxtPledgeID) > 0 And TmpRs2!PledgeBags >= Val(MSHFlexGrid2.TextMatrix(x, 7)) Then
'                                   If IsNull(TmpRs2!FGSLID_Mst_GSL_ID) = True Then
'                                      TmpRs2!PledgeBags = Val(IIf(IsNull(TmpRs2!PledgeBags), "0", TmpRs2!PledgeBags)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
'                                      TmpRs2!PledgeWeight = Val(IIf(IsNull(TmpRs2!PledgeWeight), "0", TmpRs2!PledgeWeight)) - Val(MSHFlexGrid2.TextMatrix(x, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
'                                      TmpRs2.Update
'                                   Else
'                                      mFGSLID = TmpRs2!FGSLID_Mst_GSL_ID
'                                      GoTo RASNG
'                                   End If
'                                End If
'                             End If
'                          End If
'                 End If
'              End If
'          End If
'     '   End If
'     'End If
'  End If
''End If

End Sub

Private Sub txtPledgeID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If CmbPreFix.Text = "" Then Exit Sub

If TxtPledgeID = "" Then Exit Sub
If Not IsNumeric(TxtPledgeID) Then
   MsgBox "Enter numeric value !!! "
   TxtPledgeID.SetFocus
   Exit Sub
End If

Call TableTxn_PledgeCreation(" Where sm_prefix = '" & mSM_Prefix & "' And  Txn_PledgeCreation_ID = " & TxtPledgeID)

If RsTxn_PledgeCreation.RecordCount = 0 Then
   MsgBox "No record found!!!"
   TxtPledgeID.SetFocus
   Exit Sub
End If

PledgeDate.Value = RsTxn_PledgeCreation!PledgeDate
mExchangeTypeID = RsTxn_PledgeCreation!ExchangeTypeID
mNCMLFinance = IIf(IsNull(RsTxn_PledgeCreation!NCMLFinance), 0, RsTxn_PledgeCreation!NCMLFinance)

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   If mNCMLFinance = 1 Then
      MsgBox "Selected WHR is Pledge to NFIN , Contact CMG Team for release of same!!!"
      TxtPledgeID.SetFocus
      Exit Sub
   End If
End If

TxtWHRNoNew = ""
If mExchangeTypeID = 1 Then
   Call EnDn("Non NCDEX", False)
   TxtWHRNo = RsTxn_PledgeCreation!WHRNo
   
   If LCase(mSM_Prefix) = "bknrp" Then
      mNewSM_Prefix = "BKNER"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "ujjap" Then
      mNewSM_Prefix = "Ujjan"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "dewsp" Then
      mNewSM_Prefix = "DEWAS"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "vdshp" Then
      mNewSM_Prefix = "VDISA"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "adlbp" Then
      mNewSM_Prefix = "ADLBD"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "aklap" Then
      mNewSM_Prefix = "akola"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "sngrp" Then
      mNewSM_Prefix = "sgngr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "nzmdp" Then
      mNewSM_Prefix = "nzmbd"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "unjap" Then
      mNewSM_Prefix = "unjha"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "sglip" Then
      mNewSM_Prefix = "sngli"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "gntrp" Then
      mNewSM_Prefix = "guntr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "dvgrp" Then
      mNewSM_Prefix = "dvngr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "jdhrp" Then
      mNewSM_Prefix = "jdhpr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "hngtp" Then
      mNewSM_Prefix = "hghat"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "kadip" Then
      mNewSM_Prefix = "kadi"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "ktap" Then
      mNewSM_Prefix = "kota"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "rsgrp" Then
      mNewSM_Prefix = "rsngr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   ElseIf LCase(mSM_Prefix) = "rsngr" Then
      mNewSM_Prefix = "rsgrp"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & mNewSM_Prefix & "' Or sm_prefix = '" & mSM_Prefix & "') And WHRNo = " & TxtWHRNo)
   Else
      Call TableTxn_WHRDetail("Where sm_prefix = '" & mSM_Prefix & "' And WHRNo = " & TxtWHRNo)
   End If

   TxtWHRNo = RsTxn_PledgeCreation!WHRNo
   TxtWHRNoNew = IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) & "-" & Val(TxtWHRNo) - (IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) * 1000000)
  
   WHRDate.Value = RsTxn_WHRDetail!WHRDate
   mClientIDRow = RsTxn_WHRDetail!ClientID
   ans = GetClientNameByRow(mClientIDRow)
   If ans(0) = "" Then
      MsgBox "Not found!!!"
      Exit Sub
   End If
   TxtClientID.Text = ans(0)
   TxtClientName = ans(1)
   mCommodityID = IIf(IsNull(RsTxn_WHRDetail!CommodityID), "", RsTxn_WHRDetail!CommodityID)
   TxtLicenseNo = GetLicenseNo(RsTxn_WHRDetail!LicenseID)
   mCARNo = IIf(IsNull(RsTxn_WHRDetail!CARNo), 0, RsTxn_WHRDetail!CARNo)
   mSCMFlag = IIf(IsNull(RsTxn_PledgeCreation!SCMFlag), 0, RsTxn_PledgeCreation!SCMFlag)
   If mNCMLFinance = 0 Then
      If mSCMFlag = 0 Then
         If Format(PledgeDate.Value, "yyyy-mm-dd") >= "2012-01-31" Then

            mROBags = 0
            mROQty = 0
      
            tmp = "Select sum(RocarBags)RocarBags,sum(ROCARQty)ROCARQty"
            tmp = tmp & " from Txn_ROCARMap TRO"
            tmp = tmp & " Inner join Txn_CAR TC on TRO.CARNo=TC.CARNo"
            tmp = tmp & " Inner join Txn_ReleaseOrder RO on RO.ROSRNo=TRO.ROSRNo"
            'tmp = tmp & " Where RO.Flag = 'P' And TC.WHRNo = '" & TxtWHRNoNew & "' And TC.CARNo = " & mCARNo & " "
      
            If mCARNo <> 0 Then
               tmp = tmp & " Where RocarBags is not null And RO.Flag = 'P' And TC.WHRNo = '" & TxtWHRNoNew & "' And TC.CARNo = " & mCARNo & ""
            Else
               tmp = tmp & " Where RocarBags is not null And RO.Flag = 'P' And TC.WHRNo = '" & TxtWHRNoNew & "' "
            End If
      
            Call TmpTableCMPro
       
            If TmpRsCMPro.RecordCount > 0 Then
               mROBags = IIf(IsNull(TmpRsCMPro!RocarBags), 0, TmpRsCMPro!RocarBags)
               mROQty = IIf(IsNull(TmpRsCMPro!ROCARQty), 0, TmpRsCMPro!ROCARQty)
            End If
       
            If mROBags = 0 Or mROQty = 0 Then
   '           If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "head cmg" Then
               MsgBox "No relase made in CM Pro for this WHR. Can't Release. Contact CMG/CCG for further help !!!"
               TxtPledgeID.SetFocus
               Exit Sub
   '        Else
   '           ans = MsgBox("No relase made in CM Pro for this WHR. Still want to Release Pledge?", vbYesNo)
   '           If ans = vbNo Then
   '              TxtPledgeID.SetFocus
   '              Exit Sub
   '           End If
   '        End If
          End If
        End If
     Else
        mROBags = 0
        mROQty = 0
      
        tmp = "SELECT  Sum(TSCM.DOBags) 'PrevDOBags' ,Sum(TSCM.DOQty) 'PrevDOQty'"
        tmp = tmp & " FROM  Txn_SCMDO TSCM"
        tmp = tmp & " Inner Join Txn_WHR_Detail TWD On TSCM.WHRNo = TWD.WHRNo"
        tmp = tmp & " Where TWD.SCMFlag = '1' And TWD.WHRNo = " & TxtWHRNo & " And TSCM.Flag in ('A')"
        tmp = tmp & " And TSCM.DOBags is not null "

        Call Tmp1Table
       
        If TmpRs1.RecordCount > 0 Then
           mROBags = IIf(IsNull(TmpRs1!PrevDOBags), 0, TmpRs1!PrevDOBags)
           mROQty = IIf(IsNull(TmpRs1!PrevDOQty), 0, TmpRs1!PrevDOQty)
        End If
       
        If mROBags = 0 Or mROQty = 0 Then
           MsgBox "No DO is approved in Store-Man for this WHR. Can't Release. Contact CMG/SCM for further help !!!"
           TxtPledgeID.SetFocus
           Exit Sub
        End If
      End If
   End If
ElseIf mExchangeTypeID = 2 Then
   Call EnDn("NCDEX", False)
   TxtWHRNo = RsTxn_PledgeCreation!ISINID
   Call TableMst_ISIN("Where ISINID = '" & TxtWHRNo & "'")
   mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
   mCMPID = IIf(IsNull(RsMst_ISIN!CMPID), "", RsMst_ISIN!CMPID)
   TxtLicenseNo = GetCMPName(mCMPID)
   mClientIDRow = RsTxn_PledgeCreation!ClientID
   ans = GetClientNameByRow(mClientIDRow)
   If ans(0) = "" Then
      MsgBox "Not found!!!"
      Exit Sub
   End If
   TxtClientID.Text = ans(0)
   TxtClientName = ans(1)
End If
TxtCommodity = GetCommodityName(mCommodityID)
TxtExchangeTypeID = GetExchangeName(mExchangeTypeID)
TxtRemarks = IIf(IsNull(RsTxn_PledgeCreation!Remarks), "", RsTxn_PledgeCreation!Remarks)
TxtBankName = GetBankName(RsTxn_PledgeCreation!BankID, "N")
TxtBranchName = GetBranchName(RsTxn_PledgeCreation!BranchID, "N")
TxtPledgeReqNo = IIf(IsNull(RsTxn_PledgeCreation!PledgeReqNo), "", RsTxn_PledgeCreation!PledgeReqNo)
TxtPledgeFee = IIf(IsNull(RsTxn_PledgeCreation!PledgeFees), "", RsTxn_PledgeCreation!PledgeFees)
PledgeRequestDate.Value = RsTxn_PledgeCreation!PledgeReqDate
TxtPledgeNoOfBags = 0
TxtPledgeQuantity = 0
Call Grid_Head

Call TableMst_GSL(" Where sm_prefix = '" & mSM_Prefix & "'")
Call TableTxn_PledgeCreation_GSL("Where sm_prefix = '" & mSM_Prefix & "' And Txn_PledgeCreation_ID = " & TxtPledgeID & " And PledgeNoOfBags > 0  And PledgeQuantity >= 0 ")

If RsTxn_PledgeCreation_GSL.RecordCount = 0 Then
   MsgBox "No data available for pledge release!!!"
   TxtPledgeID.SetFocus
   Exit Sub
End If

MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_PledgeCreation_GSL.RecordCount

k = 1

For I = 0 To RsTxn_PledgeCreation_GSL.RecordCount - 1
    k = k + 1
    RsMst_GSL.Find "Mst_GSL_ID = " & RsTxn_PledgeCreation_GSL!Mst_GSL_ID
    If Not RsMst_GSL.EOF Then
       If RsTxn_PledgeCreation_GSL!PledgeNoOfBags > 0 And RsTxn_PledgeCreation_GSL!PledgeQuantity >= 0 Then
          MSHFlexGrid2.TextMatrix(k, 1) = GetCMPName(RsMst_GSL!CMPID)
          MSHFlexGrid2.TextMatrix(k, 2) = GetGodownName(RsMst_GSL!GodownID, "N")
          MSHFlexGrid2.TextMatrix(k, 3) = RsMst_GSL!StackNo
          MSHFlexGrid2.TextMatrix(k, 4) = RsMst_GSL!LotNo
          MSHFlexGrid2.TextMatrix(k, 5) = RsTxn_PledgeCreation_GSL!PledgeNoOfBags
          MSHFlexGrid2.TextMatrix(k, 6) = RsTxn_PledgeCreation_GSL!PledgeQuantity
          MSHFlexGrid2.TextMatrix(k, 7) = 0 'RsTxn_PledgeCreation_GSL!PledgeNoOfBags
          MSHFlexGrid2.TextMatrix(k, 8) = 0  'RsTxn_PledgeCreation_GSL!PledgeQuantity
          MSHFlexGrid2.TextMatrix(k, 9) = RsTxn_PledgeCreation_GSL!Mst_GSL_ID
          RsMst_GSL.MoveFirst
       End If
       TxtPledgeNoOfBags = Val(TxtPledgeNoOfBags) + Val(RsTxn_PledgeCreation_GSL!PledgeNoOfBags)  'Val(TxtPledgeNoOfBags) +
       TxtPledgeQuantity = Val(TxtPledgeQuantity) + Val(RsTxn_PledgeCreation_GSL!PledgeQuantity) 'Val(TxtPledgeQuantity) +
    End If
    RsTxn_PledgeCreation_GSL.MoveNext
Next I

End Sub

Private Sub EnDn(Exg, Yn)

If Exg = "Non NCDEX" Then
   Label1.Caption = "WHR No"
   Label2.Caption = "License No"
ElseIf Exg = "NCDEX" Then
   Label1.Caption = "ISIN ID"
   Label2.Caption = "CMP"
   TxtClientID.Text = ""
   TxtClientName.Text = ""
   TxtPledgeNoOfBags.Text = ""
   TxtPledgeQuantity.Text = ""
End If

End Sub
Private Sub TxtReleaseNoOfBags_LostFocus()
If TxtReleaseNoOfBags = "" Then Exit Sub

If IsNumeric(TxtReleaseNoOfBags) = False Then
   MsgBox "Invalid numeric format!!!"
   TxtReleaseNoOfBags.SetFocus
   Exit Sub
End If

If InStr(1, TxtReleaseNoOfBags, ".") > 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtReleaseNoOfBags.SetFocus
   Exit Sub
End If

If Val(TxtReleaseNoOfBags) < 0 Then
   MsgBox "Released No of Bags can't be negetive!!!"
   TxtReleaseNoOfBags.SetFocus
   Exit Sub
End If
If Val(TxtReleaseNoOfBags) > Val(TxtUnpledgeNoOfBags) Then
   MsgBox "Released No of Bags should be less than or equal to Pledged No of Bags"
   TxtReleaseNoOfBags.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtReleaseQuantity_LostFocus()
If TxtReleaseQuantity = "" Then Exit Sub
If IsNumeric(TxtReleaseQuantity) = False Then
   MsgBox "Invalid numeric format!!!"
   TxtReleaseQuantity.SetFocus
   Exit Sub
End If
If Val(TxtReleaseQuantity) < 0 Then
   MsgBox "Released Quantity can't be negetive!!!"
   TxtReleaseQuantity.SetFocus
   Exit Sub
End If
If Val(TxtReleaseQuantity) > Val(TxtUnpledgeWeight) Then
   MsgBox "Released Quantity should be less than or equal to Pledged Quantity"
   TxtReleaseQuantity.SetFocus
   Exit Sub
End If
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

Private Function GetPledgeReleaseID() As Double
Dim Retval As Double
tmp = "Select max(PledgeReleaseID) from Txn_PledgeRelease where sm_prefix = '" & mSM_Prefix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetPledgeReleaseID = Retval
End Function
Private Function GetPledgeReleaseGSLID() As Double
Dim Retval As Double
tmp = "Select max(PledgeReleaseGSLID) from Txn_PledgeRelease_GSL where sm_prefix = '" & mSM_Prefix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetPledgeReleaseGSLID = Retval
End Function
Private Sub Indata()
CmbPreFix.Visible = False
TxtPreFix.Visible = True
CmbPreFix.Text = ""
Frame1.Enabled = False
TxtReleaseOrderNo.Locked = True
TxtReleaseID = IIf(IsNull(RsTxn_PledgeRelease!Txn_PledgeRelease_ID), "0", RsTxn_PledgeRelease!Txn_PledgeRelease_ID)
ReleaseDate.Value = RsTxn_PledgeRelease!PledgeReleaseDate
TxtPledgeID = RsTxn_PledgeRelease!Txn_PledgeCreation_ID
TxtReleaseOrderNo = IIf(IsNull(RsTxn_PledgeRelease!ReleaseOrderNo), "", RsTxn_PledgeRelease!ReleaseOrderNo)

tmp = " select * from Mst_Location Where SM_Prefix='" & RsTxn_PledgeRelease!SM_Prefix & "'"
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   TxtPreFix.Text = TmpRs2!Locationname
End If
mSM_Prefix = RsTxn_PledgeRelease!SM_Prefix
RsTxn_PledgeCreation.Filter = "sm_prefix = '" & mSM_Prefix & "'"
RsTxn_PledgeCreation.Find "Txn_PledgeCreation_ID = " & TxtPledgeID
mNCMLFinance = IIf(IsNull(RsTxn_PledgeCreation!NCMLFinance), 0, RsTxn_PledgeCreation!NCMLFinance)
If Not RsTxn_PledgeCreation.EOF Then
   PledgeDate.Value = RsTxn_PledgeCreation!PledgeDate
   mExchangeTypeID = RsTxn_PledgeCreation!ExchangeTypeID
   'TxtPreFix.Text = RsTxn_PledgeCreation!SM_Prefix
   If mExchangeTypeID = 1 Then
      Call EnDn("Non NCDEX", False)
      TxtWHRNo = RsTxn_PledgeCreation!WHRNo
      
'      If LCase(mSM_Prefix) = "bknrp" Then
'         mNewSM_Prefix = "BKNER"
         
      If LCase(mSM_Prefix) = "bknrp" Then
         mNewSM_Prefix = "BKNER"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "ujjap" Then
         mNewSM_Prefix = "Ujjan"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "dewsp" Then
         mNewSM_Prefix = "DEWAS"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "vdshp" Then
         mNewSM_Prefix = "VDISA"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "adlbp" Then
         mNewSM_Prefix = "ADLBD"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "aklap" Then
         mNewSM_Prefix = "AKOLA"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "sngrp" Then
         mNewSM_Prefix = "SGNGR"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "nzmdp" Then
         mNewSM_Prefix = "NZMBD"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "unjap" Then
         mNewSM_Prefix = "unjha"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "sglip" Then
         mNewSM_Prefix = "sngli"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "gntrp" Then
         mNewSM_Prefix = "guntr"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "dvgrp" Then
         mNewSM_Prefix = "dvngr"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "jdhrp" Then
         mNewSM_Prefix = "jdhpr"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "hngtp" Then
         mNewSM_Prefix = "hghat"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      ElseIf LCase(mSM_Prefix) = "kadip" Then
         mNewSM_Prefix = "kadi"
         Call TableTxn_WHRDetail("Where sm_prefix = '" & mNewSM_Prefix & "' And WHRNo = " & TxtWHRNo)
         If RsTxn_WHRDetail.RecordCount = 0 Then
            GoTo T
         End If
      Else
T:         Call TableTxn_WHRDetail("Where sm_prefix = '" & mSM_Prefix & "' And WHRNo = " & TxtWHRNo)
      End If
      
      'Call TableTxn_WHRDetail("Where sm_prefix = '" & mSM_Prefix & "' And WHRNo = " & TxtWHRNo)
      TxtWHRNo = TxtWHRNo
      
      
      
      TxtWHRNoNew = RsTxn_WHRDetail!CMPID & "-" & (Val(TxtWHRNo) - RsTxn_WHRDetail!CMPID * 1000000)
      WHRDate.Value = RsTxn_WHRDetail!WHRDate
      mClientIDRow = RsTxn_WHRDetail!ClientID
      ans = GetClientNameByRow(mClientIDRow)
      If ans(0) = "" Then
         MsgBox "Critical Error : Client Detail not found contact technology !!!"
         Exit Sub
      End If
      TxtClientID.Text = ans(0)
      TxtClientName = ans(1)
      mCommodityID = IIf(IsNull(RsTxn_WHRDetail!CommodityID), "", RsTxn_WHRDetail!CommodityID)
      TxtLicenseNo = GetLicenseNo(RsTxn_WHRDetail!LicenseID)
    
   ElseIf mExchangeTypeID = 2 Then
      Call EnDn("NCDEX", False)
      TxtWHRNo = RsTxn_PledgeCreation!ISINID
'      Call TableMst_ISIN("Where  ISINID = '" & TxtWHRNo & "'")
      TxtWHRNoNew = RsTxn_PledgeCreation!ISINID

      tmp = "Select Distinct mgsl.CommodityID,mgsl.ClientID,mgsl.CMPID"
      tmp = tmp & " from Txn_PledgeRelease a"
      tmp = tmp & " inner join Txn_ImportNcdexLotPleadge b on a.SM_Prefix = b.SM_Prefix and a.Txn_PledgeCreation_ID = b.Txn_PledgeCreation_ID"
      tmp = tmp & " inner join Mst_GSL mgsl on b.SM_Prefix = mgsl.SM_Prefix and b.Mst_GSL_ID = mgsl.Mst_GSL_ID"
      tmp = tmp & " inner join Txn_PledgeCreation TPC On a.SM_Prefix = TPC.SM_Prefix and a.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID"
      tmp = tmp & " Where a.SM_Prefix = '" & mSM_Prefix & "' and  a.Txn_PledgeRelease_ID = " & TxtReleaseID & ""
      
      Call TmpTable
      
      mCommodityID = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
      mCMPID = IIf(IsNull(TmpRs!CMPID), "", TmpRs!CMPID)
      TxtLicenseNo = GetCMPName(mCMPID)
      
      mClientIDRow = TmpRs!ClientID
      ans = GetClientNameByRow(mClientIDRow)
      If ans(0) = "" Then
         MsgBox "Not found!!!"
         Exit Sub
      End If
      TxtClientID.Text = ans(0)
      TxtClientName = ans(1)
      
      
   End If
   TxtCommodity = GetCommodityName(mCommodityID)
   TxtExchangeTypeID = GetExchangeName(mExchangeTypeID)
   TxtBankName = GetBankName(RsTxn_PledgeCreation!BankID, "N")
   TxtBranchName = GetBranchName(RsTxn_PledgeCreation!BranchID, "N")
   TxtPledgeReqNo = IIf(IsNull(RsTxn_PledgeCreation!PledgeReqNo), "", RsTxn_PledgeCreation!PledgeReqNo)
   TxtPledgeFee = IIf(IsNull(RsTxn_PledgeCreation!PledgeFees), "", RsTxn_PledgeCreation!PledgeFees)
   PledgeRequestDate.Value = RsTxn_PledgeCreation!PledgeReqDate
   TxtRemarks = IIf(IsNull(RsTxn_PledgeRelease!Remarks), "", RsTxn_PledgeRelease!Remarks)
   TxtCreated = IIf(IsNull(RsTxn_PledgeRelease!CreatedBy), "", RsTxn_PledgeRelease!CreatedBy) & " :- " & IIf(IsNull(RsTxn_PledgeRelease!CreatedDate), "", RsTxn_PledgeRelease!CreatedDate)
   TxtUpdated = IIf(IsNull(RsTxn_PledgeRelease!UpdatedBy), "", RsTxn_PledgeRelease!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_PledgeRelease!UpdatedDate), "", RsTxn_PledgeRelease!UpdatedDate)

End If

TxtRemarks.Locked = True
CmdUpdate.Enabled = False
Call Grid_Head
Call TableMst_GSL(" Where sm_prefix = '" & mSM_Prefix & "'")
Call TableTxn_PledgeRelease_GSL("Where sm_prefix = '" & mSM_Prefix & "' And Txn_PledgeRelease_ID = " & TxtReleaseID)
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_PledgeRelease_GSL.RecordCount
k = 1
For I = 0 To RsTxn_PledgeRelease_GSL.RecordCount - 1
    k = k + 1
    RsMst_GSL.Find "Mst_GSL_ID = " & RsTxn_PledgeRelease_GSL!Mst_GSL_ID
    If Not RsMst_GSL.EOF Then
       MSHFlexGrid2.TextMatrix(k, 1) = GetCMPName(RsMst_GSL!CMPID)
       MSHFlexGrid2.TextMatrix(k, 2) = GetGodownName(RsMst_GSL!GodownID, "N")
       MSHFlexGrid2.TextMatrix(k, 3) = RsMst_GSL!StackNo
       MSHFlexGrid2.TextMatrix(k, 4) = RsMst_GSL!LotNo
       MSHFlexGrid2.TextMatrix(k, 5) = "0" 'RsMst_GSL!WHRBags
       MSHFlexGrid2.TextMatrix(k, 6) = "0" 'RsMst_GSL!WHRWeight
       MSHFlexGrid2.TextMatrix(k, 7) = RsTxn_PledgeRelease_GSL!ReleasedNoOfBags
       MSHFlexGrid2.TextMatrix(k, 8) = RsTxn_PledgeRelease_GSL!ReleasedQuantity
       MSHFlexGrid2.TextMatrix(k, 9) = RsTxn_PledgeRelease_GSL!Mst_GSL_ID
       MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(k, 7)) 'RsTxn_PledgeRelease_GSL!ReleasedNoOfBags
       MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(k, 8))
       
       RsMst_GSL.MoveFirst
    End If
    RsTxn_PledgeRelease_GSL.MoveNext
Next I
TxtPledgeID.Locked = True
ReleaseDate.Enabled = False
Call GButtonLock(Me, True)
'DeleteCmd.Enabled = False
EditCmd.Enabled = False
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_PledgeRelease.MoveNext
If RsTxn_PledgeRelease.EOF Then
   RsTxn_PledgeRelease.MoveLast
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
RsTxn_PledgeRelease.MovePrevious
If RsTxn_PledgeRelease.BOF Then
   RsTxn_PledgeRelease.MoveFirst
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
RsTxn_PledgeRelease.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_PledgeRelease.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SearchCmd_Click()
Dim ans As Variant
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_PledgeRelease.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If

Frame0.Height = Frame4.Height - 200
Frame0.Left = Picture2.Left ' + 1400
Frame0.Top = Picture2.Top '+ 5000
Frame0.Width = Frame4.Width - 150
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 500
MSHFlexGrid1.Height = Frame0.Height - 1700

RsTxn_PledgeRelease.MoveFirst

'Call TableTxn_PledgeCreation
''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head1
'Call CmdClear_Click
''===========comment on 7 july 2012 for not clearing search grid=============

'MSHFlexGrid1.Rows = RsTxn_PledgeRelease.RecordCount + 1
'For i = 1 To RsTxn_PledgeRelease.RecordCount
'    MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(RsTxn_PledgeRelease!PledgeReleaseID), "", RsTxn_PledgeRelease!PledgeReleaseID)
'    MSHFlexGrid1.TextMatrix(i, 1) = IIf(IsNull(RsTxn_PledgeRelease!PledgeReleaseDate), "", RsTxn_PledgeRelease!PledgeReleaseDate)
'    MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(RsTxn_PledgeRelease!PledgeID), "", RsTxn_PledgeRelease!PledgeID)
'    MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(RsTxn_PledgeRelease!Remarks), "", RsTxn_PledgeRelease!Remarks)
'    MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(RsTxn_PledgeRelease!ReleaseOrderNo), "", RsTxn_PledgeRelease!ReleaseOrderNo)
'    MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(RsTxn_PledgeRelease!G_UID), "", RsTxn_PledgeRelease!G_UID)
'    RsTxn_PledgeRelease.MoveNext
'Next
End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_PledgeRelease.MoveFirst
'RsTxn_PledgeRelease.Find "PledgeReleaseID = " & Val(MSHFlexGrid1.TextMatrix(i, 0))
RsTxn_PledgeRelease.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 13) & "'"
If Not RsTxn_PledgeRelease.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Public Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 14
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1


MSHFlexGrid1.TextMatrix(0, 0) = "Pledge Release ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Release Order No"
MSHFlexGrid1.TextMatrix(0, 3) = "Pledge Release Date"
MSHFlexGrid1.TextMatrix(0, 4) = "Pledge ID"
MSHFlexGrid1.TextMatrix(0, 5) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 6) = "WHR No / ISIN ID"
MSHFlexGrid1.TextMatrix(0, 7) = "Bank Name"
MSHFlexGrid1.TextMatrix(0, 8) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 9) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 10) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 11) = "Release Bags"
MSHFlexGrid1.TextMatrix(0, 12) = "Release Qty"
MSHFlexGrid1.TextMatrix(0, 13) = "G_UID"


MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 800
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1600
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1600
MSHFlexGrid1.ColWidth(10) = 1600
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 0


MSHFlexGrid1.RowHeight(0) = 600

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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_PledgeRelease.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_PledgeRelease.xls")
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
           If C = "6" Then
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
tmp = tmp & " Where ML.SM_Prefix = '" & mSM_Prefix & "'"

Call TmpTable

mLocationName = TmpRs!Locationname
mStateName = TmpRs!StateName
mStateID = TmpRs!StateID

'mPreMsg = "<P> Dear DRC , <br> "
'mPreMsg = mPreMsg & " The below mentioned WHR/ISIN is released today in mentioned CMP warehouse. The details of the pledge release are given below. Kindly generate CAR No for the mentioned WHR to the lending Bank at the earliest."
'mPreMsg = mPreMsg & "<br>"

mMsg = mMsg & "<tr><td> Release ID </td><td> " & mSM_Prefix & " - " & TxtReleaseID & " </td></tr>"
mMsg = mMsg & "<tr><td> Release Order No </td><td> " & TxtReleaseOrderNo & " </td></tr>"
mMsg = mMsg & "<tr><td> Pledge ID </td><td> " & mSM_Prefix & " - " & TxtPledgeID & " </td></tr>"
mMsg = mMsg & "<tr><td> Location Name </td><td> " & mLocationName & "  </td></tr>"
mMsg = mMsg & "<tr><td> State Name </td><td> " & mStateName & "  </td></tr>"

If LCase(TxtExchangeTypeID.Text) = "non ncdex" Then
   mMsg = mMsg & "<tr><td> WHR No </td><td> " & TxtWHRNoNew & "   </td></tr>"
   mMsg = mMsg & "<tr><td> WHR Date </td><td> " & Format(WHRDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
ElseIf LCase(TxtExchangeTypeID.Text) = "ncdex" Then
   mMsg = mMsg & "<tr><td> ISIN </td><td> TxtWHRNo </td></tr>"
   mMsg = mMsg & "<tr><td> ISIN Valid Date </td><td> " & Format(WHRDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
End If

mMsg = mMsg & "<tr><td> Borrower Name and ID </td><td> " & TxtClientName & " - " & TxtClientID & " </td></tr>"
mMsg = mMsg & "<tr><td> Commodity </td><td> " & TxtCommodity & " </td></tr>"

'If LCase(TxtExchangeTypeID.Text) = "non ncdex" Then
'   mMsg = mMsg & "<tr><td> Value of Commodity </td><td> " & Val(mValueOfCommodity) & " </td></tr>"
'End If


mMsg = mMsg & "<tr><td> Bank Name </td><td> " & TxtBankName & " </td></tr>"

mMsg = mMsg & "<tr><td> Branch Name </td><td> " & TxtBranchName & " </td></tr>"
mMsg = mMsg & "<tr><td> Pledge Released by  </td><td> " & Gen_UserName & " - " & Gen_UserLoginID & " (" & GetDesignaion(Gen_DesignationID) & ")</td></tr>"

If LCase(TxtExchangeTypeID.Text) = "non ncdex" Then

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
       mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid2.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 7) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 8) & " </td></tr>"
    Else
       mGridMsg = mGridMsg & "<tr><td> &nbsp </td><td> &nbsp </td><td> &nbsp </td><td> Total Pledge </td><td> " & MSHFlexGrid2.TextMatrix(0, 7) & " </td><td> " & MSHFlexGrid2.TextMatrix(0, 8) & " </td></tr>"
    End If
Next

mMsg = mMsg & "<tr><td> Details </td><td><table border = 1>" & mGridMsg & "</table></td></tr>"

'mMsg = mMsg & "<br><br><br>"

'mPostMsg = "<br><br><br>"
'mPostMsg = mPostMsg & "<P>Please issue SR to the Bank immediately."


'TOADD = "sr.ccg@ncml.com"
'TOADD = mForwardToEmailID ' mEmployeeEmailID

If mNCMLFinance = 0 Then
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
Else

   TOADD = ""
   TOCC = ""

   TOADD = Gen_ReadNFinMailIDINI("NFinCC")
   TOCC = Gen_ReadNFinMailIDINI("NFinTo")

End If
'TOCC = "rakesh.go@ncml.com"
'TOADD = "sagar.k@ncml.com"

mMsgSubject = "New Pledge Released in StoreMan - Txn_PledgeRelease_ID is " & TxtReleaseID & " in Location  " & mLocationName & ""

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)


End Sub
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT> " & mPreMsg & " <table border = 1>"
htmlStringtail = "</table></Font></body></html>"

'htmlStringtail = "</table> " & mPostMsg & " </Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  .MailFrom = Gen_UserEmailID
'  'message
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "manish.g@ncml.com" '"sr.ccg@ncml.com"
'  End If
'  .SendTo = ToADD_
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
