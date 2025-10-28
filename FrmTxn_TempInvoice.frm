VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_TempInvoice 
   Caption         =   "Ad-hoc Invoice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   10095
      Left            =   14925
      TabIndex        =   20
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   10920
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   9615
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   15135
      Begin VB.Frame Frame6 
         BackColor       =   &H008080FF&
         Caption         =   "Reverse Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   3720
         TabIndex        =   97
         Top             =   2880
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox TxtReverseRemark 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   1560
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   100
            Top             =   840
            Width           =   4680
         End
         Begin VB.CommandButton CmdSaveReverse 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1575
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   1440
            Width           =   1440
         End
         Begin VB.CommandButton CmdCancelReverse 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3015
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   1440
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker TxtReverseDate 
            Height          =   375
            Left            =   1560
            TabIndex        =   101
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   54788097
            CurrentDate     =   39884
         End
         Begin VB.Label Label37 
            BackColor       =   &H008080FF&
            Caption         =   "Remark"
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
            TabIndex        =   103
            Top             =   930
            Width           =   1260
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Reverse Date"
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
            TabIndex        =   102
            Top             =   390
            Width           =   1800
         End
      End
      Begin VB.Frame Frame0 
         Height          =   1965
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   15360
         Begin VB.ComboBox CmbClient 
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
            Left            =   5400
            Sorted          =   -1  'True
            TabIndex        =   32
            Top             =   800
            Width           =   9555
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PI Date"
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
            TabIndex        =   109
            Top             =   240
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Invoice Date"
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
            Left            =   1320
            TabIndex        =   108
            Top             =   240
            Width           =   1440
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4830
            TabIndex        =   95
            Top             =   1248
            Width           =   225
         End
         Begin VB.TextBox TxtSAINo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   870
            MaxLength       =   12
            TabIndex        =   93
            Top             =   1195
            Width           =   2025
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   345
            Left            =   12750
            TabIndex        =   28
            Top             =   1203
            Width           =   1125
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13905
            TabIndex        =   27
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
            Left            =   10560
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   390
            Width           =   2520
         End
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_TempInvoice.frx":0000
            Left            =   13080
            List            =   "FrmTxn_TempInvoice.frx":0010
            TabIndex        =   25
            Top             =   375
            Width           =   1845
         End
         Begin VB.ComboBox CmbSInvoiceType 
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
            ItemData        =   "FrmTxn_TempInvoice.frx":0051
            Left            =   2955
            List            =   "FrmTxn_TempInvoice.frx":0053
            TabIndex        =   24
            Top             =   375
            Width           =   2460
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
            ItemData        =   "FrmTxn_TempInvoice.frx":0055
            Left            =   5445
            List            =   "FrmTxn_TempInvoice.frx":0057
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   375
            Width           =   2040
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
            ItemData        =   "FrmTxn_TempInvoice.frx":0059
            Left            =   7515
            List            =   "FrmTxn_TempInvoice.frx":005B
            Sorted          =   -1  'True
            TabIndex        =   22
            Top             =   375
            Width           =   2985
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   855
            Left            =   90
            TabIndex        =   29
            Top             =   1680
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1508
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
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   345
            Left            =   120
            TabIndex        =   30
            Top             =   735
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
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
            Format          =   54788097
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   360
            Left            =   1575
            TabIndex        =   31
            Top             =   735
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   54788097
            CurrentDate     =   39884
         End
         Begin VB.Label Label48 
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
            Left            =   3000
            TabIndex        =   150
            Top             =   840
            Width           =   1230
         End
         Begin VB.Label Label35 
            Caption         =   "Approved By FAG"
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
            Left            =   3000
            TabIndex        =   96
            Top             =   1248
            Width           =   1710
         End
         Begin VB.Label Label33 
            Caption         =   "AI No"
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
            Left            =   195
            TabIndex        =   94
            Top             =   1248
            Width           =   615
         End
         Begin VB.Label Label16 
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
            Left            =   8400
            TabIndex        =   41
            Top             =   1248
            Width           =   2415
         End
         Begin VB.Label Label15 
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
            Left            =   10920
            TabIndex        =   40
            Top             =   1255
            Width           =   135
         End
         Begin VB.Label Label12 
            Caption         =   "From Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   532
            TabIndex        =   39
            Top             =   480
            Width           =   510
         End
         Begin VB.Label Label34 
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
            Height          =   225
            Left            =   11400
            TabIndex        =   38
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label14 
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
            Height          =   225
            Left            =   2107
            TabIndex        =   37
            Top             =   480
            Width           =   270
         End
         Begin VB.Label Label13 
            Caption         =   "Status"
            Height          =   240
            Left            =   13800
            TabIndex        =   36
            Top             =   195
            Width           =   525
         End
         Begin VB.Label Label29 
            Caption         =   "Invoice Type"
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
            Left            =   3600
            TabIndex        =   35
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Label30 
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
            Height          =   225
            Left            =   5760
            TabIndex        =   34
            Top             =   120
            Width           =   1545
         End
         Begin VB.Label Label31 
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
            Height          =   225
            Left            =   8280
            TabIndex        =   33
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1140
         Left            =   120
         TabIndex        =   85
         Top             =   8400
         Width           =   13320
         Begin VB.CommandButton CmdViewInvoice 
            Caption         =   "View Invoice"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   11820
            TabIndex        =   107
            Top             =   600
            Width           =   1380
         End
         Begin VB.CommandButton CmdReverse 
            Caption         =   "Reverse"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   11820
            TabIndex        =   104
            Top             =   195
            Width           =   1380
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   105
            TabIndex        =   0
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2445
            TabIndex        =   15
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4785
            TabIndex        =   91
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   7125
            TabIndex        =   90
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   9465
            TabIndex        =   17
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   105
            TabIndex        =   89
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2445
            TabIndex        =   88
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4785
            TabIndex        =   87
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   7125
            TabIndex        =   86
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   9465
            TabIndex        =   16
            Top             =   600
            Width           =   2340
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8295
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   13305
         Begin TabDlg.SSTab SSTab1 
            Height          =   2565
            Left            =   120
            TabIndex        =   118
            Top             =   5640
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   4524
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "TI/ Log Details"
            TabPicture(0)   =   "FrmTxn_TempInvoice.frx":005D
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label41"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label40"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label39"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label38"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label28"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label23"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label5"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label10"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label9"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label49"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "TxtBalAmt"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "TxtTIAmount"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "TxtTIDate"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "TxtTINo"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "TxtTISTAmt"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "TxtRecAmt"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "TxtEntryNo"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "TxtCreated"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "TxtUpdated"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "TxtPostInvoiceNo"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).ControlCount=   20
            TabCaption(1)   =   "Godown details"
            TabPicture(1)   =   "FrmTxn_TempInvoice.frx":0079
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame2"
            Tab(1).Control(1)=   "MSHFlexGrid3"
            Tab(1).Control(2)=   "Check1"
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Reservation Godown details"
            TabPicture(2)   =   "FrmTxn_TempInvoice.frx":0095
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "TxtDistributedAmt"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "MSHFlexGrid2"
            Tab(2).ControlCount=   2
            Begin VB.TextBox TxtPostInvoiceNo 
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
               IMEMode         =   3  'DISABLE
               Left            =   2055
               Locked          =   -1  'True
               TabIndex        =   151
               Top             =   1170
               Width           =   5415
            End
            Begin VB.Frame Frame2 
               Height          =   765
               Left            =   -74880
               TabIndex        =   141
               Top             =   420
               Width           =   12855
               Begin VB.CommandButton CmdRemoveFromList 
                  Caption         =   "Remove From List"
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
                  Left            =   10575
                  TabIndex        =   149
                  Top             =   450
                  Width           =   1935
               End
               Begin VB.CommandButton CmdAddToList 
                  Caption         =   "Add To List"
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
                  Left            =   8865
                  TabIndex        =   148
                  Top             =   450
                  Width           =   1695
               End
               Begin VB.TextBox TxtDistIncomeAmount 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   9120
                  TabIndex        =   144
                  TabStop         =   0   'False
                  Top             =   135
                  Width           =   2055
               End
               Begin VB.ComboBox CmbGodownNo 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   6270
                  Sorted          =   -1  'True
                  TabIndex        =   143
                  Top             =   120
                  Width           =   2145
               End
               Begin VB.ComboBox CmbCMPName 
                  Height          =   315
                  Left            =   720
                  Sorted          =   -1  'True
                  TabIndex        =   142
                  Top             =   120
                  Width           =   4560
               End
               Begin VB.Label lblGodownNo 
                  Caption         =   "Godown No"
                  Height          =   180
                  Left            =   5400
                  TabIndex        =   147
                  Top             =   180
                  Width           =   1095
               End
               Begin VB.Label Label46 
                  Caption         =   "CMP"
                  Height          =   240
                  Left            =   120
                  TabIndex        =   146
                  Top             =   150
                  Width           =   480
               End
               Begin VB.Label Label47 
                  Caption         =   "Amount"
                  Height          =   180
                  Left            =   8520
                  TabIndex        =   145
                  Top             =   195
                  Width           =   1095
               End
            End
            Begin VB.TextBox TxtDistributedAmt 
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
               Left            =   -73665
               TabIndex        =   137
               TabStop         =   0   'False
               Top             =   1065
               Visible         =   0   'False
               Width           =   2175
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
               Left            =   7920
               Locked          =   -1  'True
               TabIndex        =   134
               TabStop         =   0   'False
               Top             =   1800
               Width           =   4890
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
               Left            =   2055
               Locked          =   -1  'True
               TabIndex        =   133
               TabStop         =   0   'False
               Top             =   1800
               Width           =   3810
            End
            Begin VB.TextBox TxtEntryNo 
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
               IMEMode         =   3  'DISABLE
               Left            =   11010
               Locked          =   -1  'True
               TabIndex        =   125
               Top             =   750
               Width           =   1815
            End
            Begin VB.TextBox TxtRecAmt 
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
               IMEMode         =   3  'DISABLE
               Left            =   7380
               Locked          =   -1  'True
               TabIndex        =   124
               Top             =   750
               Width           =   1815
            End
            Begin VB.TextBox TxtTISTAmt 
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
               IMEMode         =   3  'DISABLE
               Left            =   5685
               Locked          =   -1  'True
               TabIndex        =   123
               Top             =   750
               Width           =   1695
            End
            Begin VB.TextBox TxtTINo 
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
               IMEMode         =   3  'DISABLE
               Left            =   240
               Locked          =   -1  'True
               TabIndex        =   122
               Top             =   750
               Width           =   1815
            End
            Begin VB.TextBox TxtTIDate 
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
               IMEMode         =   3  'DISABLE
               Left            =   2055
               Locked          =   -1  'True
               TabIndex        =   121
               Top             =   750
               Width           =   1815
            End
            Begin VB.TextBox TxtTIAmount 
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
               IMEMode         =   3  'DISABLE
               Left            =   3870
               Locked          =   -1  'True
               TabIndex        =   120
               Top             =   750
               Width           =   1815
            End
            Begin VB.TextBox TxtBalAmt 
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
               IMEMode         =   3  'DISABLE
               Left            =   9195
               Locked          =   -1  'True
               TabIndex        =   119
               Top             =   750
               Width           =   1815
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   2055
               Left            =   -74880
               TabIndex        =   138
               Top             =   360
               Width           =   12855
               _ExtentX        =   22675
               _ExtentY        =   3625
               _Version        =   393216
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
               Height          =   1290
               Left            =   -74880
               TabIndex        =   139
               Top             =   1200
               Width           =   12855
               _ExtentX        =   22675
               _ExtentY        =   2275
               _Version        =   393216
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Clear Amount"
               Height          =   195
               Left            =   -63480
               TabIndex        =   140
               Top             =   1200
               Width           =   1380
            End
            Begin VB.Label Label49 
               Caption         =   "Post Invoice No"
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
               Left            =   240
               TabIndex        =   152
               Top             =   1230
               Width           =   1455
            End
            Begin VB.Label Label9 
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
               Left            =   6420
               TabIndex        =   136
               Top             =   1860
               Width           =   975
            End
            Begin VB.Label Label10 
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
               Left            =   360
               TabIndex        =   135
               Top             =   1860
               Width           =   1695
            End
            Begin VB.Label Label5 
               Caption         =   "Entry No"
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
               Left            =   11505
               TabIndex        =   132
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label23 
               Caption         =   "Received Amount"
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
               Left            =   7485
               TabIndex        =   131
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label28 
               Caption         =   "TI ST Amount"
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
               Left            =   5865
               TabIndex        =   130
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label38 
               Caption         =   "TI No"
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
               Left            =   915
               TabIndex        =   129
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Label39 
               Caption         =   "TI Amount"
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
               Left            =   4185
               TabIndex        =   128
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label40 
               Caption         =   "TI Date"
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
               Left            =   2505
               TabIndex        =   127
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label41 
               Caption         =   "Balance Amount"
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
               Left            =   9345
               TabIndex        =   126
               Top             =   480
               Width           =   1695
            End
         End
         Begin VB.ComboBox CmbMSC 
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00B1
            Left            =   7680
            List            =   "FrmTxn_TempInvoice.frx":00B3
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   4680
            Width           =   4935
         End
         Begin VB.TextBox TxtRNAVCode 
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
            Left            =   11430
            Locked          =   -1  'True
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   2925
            Width           =   1200
         End
         Begin VB.TextBox TxtNAVCode 
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
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   2925
            Width           =   1200
         End
         Begin VB.TextBox TxtNAVClientName 
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   5130
            Width           =   10770
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Show DP "
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
            Left            =   1815
            TabIndex        =   106
            Top             =   1620
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.ComboBox CmbDPName 
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
            Sorted          =   -1  'True
            TabIndex        =   105
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TxtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   4545
            MaxLength       =   12
            TabIndex        =   12
            Top             =   4680
            Width           =   1455
         End
         Begin VB.ComboBox CmbClientDP 
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
            Sorted          =   -1  'True
            TabIndex        =   45
            Top             =   1920
            Visible         =   0   'False
            Width           =   5805
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00B5
            Left            =   1815
            List            =   "FrmTxn_TempInvoice.frx":00B7
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TxtRemark 
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
            Left            =   1815
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   3390
            Width           =   10800
         End
         Begin VB.TextBox TxtPIID 
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   240
            Width           =   1515
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00B9
            Left            =   7680
            List            =   "FrmTxn_TempInvoice.frx":00BB
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   705
            Width           =   4935
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00BD
            Left            =   1800
            List            =   "FrmTxn_TempInvoice.frx":00BF
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   1170
            Width           =   3855
         End
         Begin VB.TextBox TxtAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1815
            MaxLength       =   12
            TabIndex        =   9
            Top             =   2925
            Width           =   3840
         End
         Begin VB.ComboBox CmbMonth 
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00C1
            Left            =   7680
            List            =   "FrmTxn_TempInvoice.frx":00C3
            TabIndex        =   4
            Top             =   1170
            Width           =   2055
         End
         Begin VB.ComboBox CmbYear 
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00C5
            Left            =   10680
            List            =   "FrmTxn_TempInvoice.frx":00C7
            TabIndex        =   5
            Top             =   1170
            Width           =   1935
         End
         Begin VB.ComboBox CmbInvoiceType 
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00C9
            Left            =   7680
            List            =   "FrmTxn_TempInvoice.frx":00CB
            TabIndex        =   1
            Top             =   240
            Width           =   4920
         End
         Begin VB.TextBox TxtBatchNo 
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
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2925
            Width           =   960
         End
         Begin VB.ComboBox CmbServiceTaxID 
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00CD
            Left            =   6840
            List            =   "FrmTxn_TempInvoice.frx":00CF
            TabIndex        =   83
            Top             =   2400
            Visible         =   0   'False
            Width           =   5790
         End
         Begin VB.TextBox TxtFAGRemarks 
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
            Left            =   1815
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   4050
            Width           =   10800
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
            Left            =   12720
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   705
            Width           =   375
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
            Left            =   12720
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   375
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
            Height          =   360
            Left            =   12720
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1920
            Width           =   375
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
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1920
            Width           =   4590
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
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1920
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
            Left            =   1815
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1920
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
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3495
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
            Left            =   5685
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1170
            Width           =   375
         End
         Begin VB.CommandButton CmdInvoiceType 
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
            Left            =   12720
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   240
            Width           =   375
         End
         Begin MSComCtl2.DTPicker TxtPIDate 
            Height          =   375
            Left            =   4320
            TabIndex        =   54
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   54788097
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtFAGPostingDate 
            Height          =   375
            Left            =   1815
            TabIndex        =   14
            Top             =   4680
            Width           =   1440
            _ExtentX        =   2540
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
            Format          =   54788097
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtInvoiceType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   240
            Width           =   4920
         End
         Begin VB.TextBox TxtExchangeTypeID 
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
            Left            =   1815
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   705
            Width           =   3810
         End
         Begin VB.TextBox TxtLocationID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   705
            Width           =   4920
         End
         Begin VB.TextBox TxtCommodityID 
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1170
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox TxtMonth 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   1170
            Width           =   2040
         End
         Begin VB.TextBox TxtYear 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   1170
            Width           =   1920
         End
         Begin VB.TextBox TxtServiceTaxID 
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
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   2400
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00D1
            Left            =   1815
            List            =   "FrmTxn_TempInvoice.frx":00D3
            Sorted          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   615
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
            ItemData        =   "FrmTxn_TempInvoice.frx":00D5
            Left            =   1815
            List            =   "FrmTxn_TempInvoice.frx":00E5
            TabIndex        =   8
            Top             =   2400
            Width           =   3855
         End
         Begin VB.TextBox TxtStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   2400
            Width           =   3855
         End
         Begin VB.TextBox TxtMSC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   7680
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   117
            Top             =   4680
            Width           =   4920
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Contract"
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
            Left            =   6180
            TabIndex        =   116
            Top             =   4740
            Width           =   735
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "RNAVCode"
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
            Left            =   10320
            TabIndex        =   114
            Top             =   2985
            Width           =   1065
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "NAVCode"
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
            Left            =   7920
            TabIndex        =   112
            Top             =   2985
            Width           =   915
         End
         Begin VB.Label Label42 
            Caption         =   "NAVISION Client Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            TabIndex        =   111
            Top             =   5085
            Width           =   1695
         End
         Begin VB.Label Label32 
            Caption         =   "Quantity/No. of Bags"
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
            Left            =   3345
            TabIndex        =   92
            Top             =   4620
            Width           =   1140
         End
         Begin VB.Label Label17 
            Caption         =   "Remark"
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
            TabIndex        =   84
            Top             =   3540
            Width           =   1800
         End
         Begin VB.Label Label8 
            Caption         =   "AI Date"
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
            Left            =   3525
            TabIndex        =   82
            Top             =   285
            Width           =   765
         End
         Begin VB.Label Label4 
            Caption         =   "AI No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   81
            Top             =   315
            Width           =   1800
         End
         Begin VB.Label LblCMPName 
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
            Height          =   210
            Left            =   6180
            TabIndex        =   80
            Top             =   780
            Width           =   900
         End
         Begin VB.Label Label1 
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
            Left            =   120
            TabIndex        =   79
            Top             =   1223
            Width           =   1215
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   758
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Amount excl. Tax"
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
            Left            =   120
            TabIndex        =   77
            Top             =   2858
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Month"
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
            Left            =   6180
            TabIndex        =   76
            Top             =   1223
            Width           =   645
         End
         Begin VB.Label Label18 
            Caption         =   "Year"
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
            Left            =   9960
            TabIndex        =   75
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Status"
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
            TabIndex        =   74
            Top             =   2430
            Width           =   1800
         End
         Begin VB.Label Label19 
            Caption         =   "Invoice Type"
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
            Left            =   6180
            TabIndex        =   73
            Top             =   270
            Width           =   1380
         End
         Begin VB.Label Label20 
            Caption         =   "Batch No"
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
            Left            =   5820
            TabIndex        =   72
            Top             =   2985
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Tax Code"
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
            Left            =   5820
            TabIndex        =   71
            Top             =   2460
            Width           =   1110
         End
         Begin VB.Label Label22 
            Caption         =   "FAG Remark"
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
            TabIndex        =   70
            Top             =   4200
            Width           =   1800
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
            Left            =   7920
            TabIndex        =   69
            Top             =   1680
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label24 
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
            Left            =   9480
            TabIndex        =   68
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label25 
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
            Left            =   7132
            TabIndex        =   67
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label26 
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
            Left            =   4560
            TabIndex        =   66
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label27 
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
            Left            =   120
            TabIndex        =   65
            Top             =   1980
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "FAG Posting Date"
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
            TabIndex        =   64
            Top             =   4710
            Width           =   1620
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_TempInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mStatus As Variant
Dim mLocationID As Double
Dim mExchangeTypeID As Double
Dim mCommodityID As Double
Dim mClientIDRow As Double
Dim mFromDate, mToDate, mSM_Prefix, mInvoiceTypeID, mMonthListindex, mServiceTaxID, mContractID As Variant
Dim mDPID As String
Dim mEntryNo As Variant
Dim mSelRow, mSelCol As Variant
Dim mCMPID, mGodownID As Variant

Dim mSExchangeTypeID, mSClientID, mSClientRowID As Variant


Private Sub AddCmd_Click()
Edit_Flg = "A"

Call GButtonLock(Me, False)
Call Grid_Head3

TxtPIID = ""
'TxtPIDate.Enabled = True
TxtPIDate.Value = Date

CmbInvoiceType.Visible = True
TxtInvoiceType.Visible = False
CmbInvoiceType.Text = ""
mInvoiceTypeID = Null
CmdInvoiceType.Visible = True

'CmbExchangeTypeID.Visible = True
'TxtExchangeTypeID.Visible = False
'CmbExchangeTypeID.Text = ""
TxtExchangeTypeID = ""
mExchangeTypeID = 0

CmbLocationID.Visible = True
TxtLocationID.Visible = False
CmbLocationID.Text = ""
CmdLocationID.Visible = True
mLocationID = 0

CmbCommodityID.Visible = True
TxtCommodityID.Visible = False
CmbCommodityID.Text = ""
CmdCommodityID.Visible = True
mCommodityID = 0

CmbMonth.Visible = True
TxtMonth.Visible = False
CmbMonth.Text = ""


CmbYear.Visible = True
TxtYear.Visible = False
CmbYear.Text = ""

'CmbServiceTaxID.Visible = True
'TxtServiceTaxID.Visible = False
'CmbServiceTaxID.Text = ""
'CmdServiceTaxID.Visible = True

TxtServiceTaxID = ""
mServiceTaxID = Null

TxtMSC.Text = ""
CmbMSC.Text = ""
mContractID = Null

CmbStatus.Visible = False
TxtStatus.Visible = True
CmbStatus.Text = "Pending"
TxtStatus.Text = "Pending"
mStatusFlag = "P"

TxtClientID.Locked = False

TxtClientID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
TxtNAVClientName = ""
CmbClientDP.Text = ""
CmbClientDP.Visible = False
CmdClient.Visible = True
TxtQty.Locked = False
TxtAmount.Locked = False
TxtRemark.Locked = False
TxtFAGRemarks.Locked = True
CmbDPName.Text = ""
CmbDPName.Visible = False
CmbDPName.Enabled = True
TxtNAVCode.Text = ""
TxtRNAVCode.Text = ""


TxtFAGPostingDate.Visible = False
Label11.Visible = False

TxtRemark = ""
TxtAmount = ""
TxtFAGRemarks = ""
TxtQty = ""
TxtTINo = ""
TxtTIDate = ""
TxtTIAmount = ""
TxtTISTAmt = ""
TxtEntryNo = ""
TxtRecAmt = ""
TxtBalAmt = ""
mEntryNo = 0

TxtBatchNo = ""

TxtCreated = ""
TxtUpdated = ""
CmdReverse.Enabled = False
Frame6.Visible = False
Call Grid_Head2
If RsTxn_TempInvoice.RecordCount > 0 Then
   CmbInvoiceType.SetFocus
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   If MSHFlexGrid3.Rows > 3 Then
      For I = 2 To MSHFlexGrid3.Rows - 2
          MSHFlexGrid3.TextMatrix(I, 6) = 0
      Next
      MSHFlexGrid3.TextMatrix(0, 6) = 0
   End If
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   CmbDPName.Visible = True
   CmbDPName.Width = 10800 '9400 'CmbNonNcdexClient.Width
   CmbDPName.Top = TxtClientID.Top 'CmbNonNcdexClient.Top
   CmbDPName.Left = TxtClientID.Left 'CmbNonNcdexClient.Left
   CmbNonNcdexClient.Visible = False
'   SSTab1.Tab = 0
Else
   Label25.Visible = True
   Label26.Visible = True
   Label28.Visible = True
'   CmbNonNcdexClient.Visible = True
'   CmbNonNcdexClient.Width = CmbDPName.Width
'   CmbNonNcdexClient.Top = CmbDPName.Top
'   CmbNonNcdexClient.Left = CmbDPName.Left
   CmbDPName.Visible = False
End If

End Sub


Private Sub CmbInvoiceType_GotFocus()

tmp = CmbInvoiceType.Text
Call TableMst_IncomeExpenseType(" where IEFlag = 'I'")
CmbInvoiceType.Clear
If RsMst_IncomeExpenseType.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_IncomeExpenseType.RecordCount
    CmbInvoiceType.AddItem RsMst_IncomeExpenseType!IEType
    RsMst_IncomeExpenseType.MoveNext
Next
CmbInvoiceType.Text = tmp

End Sub

Private Sub CmbInvoiceType_LostFocus()

If CmbInvoiceType.Text = "" Then
   Exit Sub
End If
RsMst_IncomeExpenseType.MoveFirst
RsMst_IncomeExpenseType.Find "IEType = '" & CmbInvoiceType.Text & "'"
If RsMst_IncomeExpenseType.EOF Then
   MsgBox "Invalid selection "
   CmbInvoiceType.SetFocus
   Exit Sub
End If
If tmp <> CmbInvoiceType.Text Then
   CmbCommodityID.Text = ""
End If
mInvoiceTypeID = RsMst_IncomeExpenseType!IETypeID
If IsNull(RsMst_IncomeExpenseType!ExchangeTypeID) = True Then
   MsgBox "Please Map Exchange Type to selected Income Type!!"
   CmbInvoiceType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_IncomeExpenseType!ExchangeTypeID
TxtExchangeTypeID.Text = GetExchangeName(Val(mExchangeTypeID))

If mExchangeTypeID = 1 Then
   CmbNonNcdexClient.Width = 10680
   CmbNonNcdexClient.Top = TxtClientID.Top
   CmbNonNcdexClient.Left = TxtClientID.Left
   Label25.Visible = False
   Label24.Visible = False
   Label27.Visible = False
   CmbNonNcdexClient.Visible = True
   If Edit_Flg = "E" Then
      CmbNonNcdexClient.Text = TxtClientName
   Else
      If tmp <> CmbInvoiceType.Text Then
         CmbNonNcdexClient.Text = ""
      End If
   End If
Else
   Label25.Visible = True
   Label24.Visible = True
   Label27.Visible = True
   CmbNonNcdexClient.Visible = False
   If LCase(CmbInvoiceType.Text) = "dp-client" Or LCase(CmbInvoiceType.Text) = "commtrack income" Then
      Check3.Value = 1
   Else
      Check3.Value = 0
   End If
   Call Check3_Click
End If
If mInvoiceTypeID = 4 Then
   If tmp <> CmbInvoiceType.Text Then
      CmbMSC.Text = ""
   End If
   CmbMSC.Visible = True
   TxtMSC.Visible = False
   Call Grid_Head3
   Frame2.Enabled = False
   Check1.Visible = False
Else
   CmbMSC.Text = ""
   CmbMSC.Visible = False
   TxtMSC.Visible = True
   TxtMSC = ""
   Call Grid_Head2
   Frame2.Enabled = True
   Check1.Visible = True
End If

End Sub

Private Sub CmbMonth_GotFocus()
tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I), I - 1
Next
CmbMonth.Text = tmp
End Sub
Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
       CmbMonth.Text = CmbMonth.List(I)
       mMonthListindex = I
       mStatus = True
       Exit For
    End If
Next
If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If

'Call GetConnectedGodowns
'Call DistributeGodownIncome

End Sub

''-- Ullhas Patil 27 Sep 2017
Private Sub CmbClient_GotFocus()
If CmbSExchangeTypeID.Text = "" Then
    MsgBox "Select Exchange Type"
    Exit Sub
End If
tmp = CmbClient.Text
Call TableMst_Client(" Where ExchangeTypeID = " & mSExchangeTypeID & " Order By ClientName ")
CmbClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Client found"
   CmbClient.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClient.AddItem UCase(RsMst_Client!ClientName) & "#" & RsMst_Client!ClientIDRow & "#" & RsMst_Client!ClientID
    RsMst_Client.MoveNext
Next
CmbClient.Text = tmp
End Sub
Private Sub CmbClient_LostFocus()
Dim mSplit As Variant
If CmbClient.Text = "" Then
   Exit Sub
End If
mSplit = Split(CmbClient.Text, "#")
If UBound(mSplit) <> 2 Then
   MsgBox "Invalid selection "
   CmbClient.SetFocus
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientID = " & mSplit(2)
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbClient.SetFocus
   Exit Sub
Else
   If LCase(RsMst_Client!ClientName) <> LCase(mSplit(0)) Then
    MsgBox "Invalid selection "
    CmbClient.SetFocus
    Exit Sub
   End If
End If

mSClientRowID = IIf(IsNull(RsMst_Client!ClientID), "", RsMst_Client!ClientID)
mSClientID = IIf(IsNull(RsMst_Client!ClientID), "", RsMst_Client!ClientID)
''mDepositoryID = IIf(IsNull(RsMst_Client!DepositoryID), "", RsMst_Client!DepositoryID)
''mNAVCode = IIf(IsNull(RsMst_Client!NavCode), "", RsMst_Client!NavCode)
End Sub
''-----

Private Sub CmbSCommodityID_GotFocus()
tmp = CmbSCommodityID.Text
Call TableMst_Commodity
CmbSCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbSCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbSCommodityID.Text = tmp

End Sub

Private Sub CmbSCommodityID_LostFocus()
If CmbSCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodityID.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbsExchangeTypeID_GotFocus()

tmp = CmbSExchangeTypeID.Text
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

Private Sub CmbSInvoiceType_GotFocus()

tmp = CmbSInvoiceType.Text
Call TableMst_IncomeExpenseType(" where IEFlag = 'I'")
CmbSInvoiceType.Clear
If RsMst_IncomeExpenseType.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_IncomeExpenseType.RecordCount
    CmbSInvoiceType.AddItem RsMst_IncomeExpenseType!IEType
    RsMst_IncomeExpenseType.MoveNext
Next
CmbSInvoiceType.Text = tmp

End Sub

Private Sub CmbSInvoiceType_LostFocus()
If CmbSInvoiceType.Text = "" Then
   Exit Sub
End If
RsMst_IncomeExpenseType.MoveFirst
RsMst_IncomeExpenseType.Find "IEType = '" & CmbSInvoiceType.Text & "'"
If RsMst_IncomeExpenseType.EOF Then
   MsgBox "Invalid selection "
   CmbSInvoiceType.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
End Sub

Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbSLocationID.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then
   Exit Sub
End If

If LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "rejected" And LCase(CmbSStatus.Text) <> "freeze - dont generate invoice" Then
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then
   TxtFAGRemarks.Locked = True
   TxtFAGPostingDate.Visible = False
   Label11.Visible = False
   Exit Sub
End If

If LCase(CmbStatus.Text) = "approved" Then
   mStatus = "A"
   TxtFAGRemarks.Locked = False
   TxtFAGPostingDate.Visible = True
   Label11.Visible = True
   TxtFAGPostingDate.Enabled = True
   mMonth = mMonthListindex + 1
   If mMonth < 10 Then
      mMonth = "0" & mMonth
   End If
   mTotalDays = Day(DateSerial(CmbYear, mMonth + 1, 0))
   mFAGToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
    '"" 'Date
   If CDate(mFAGToDate) > Date Then
      TxtFAGPostingDate.Value = Date
   Else
      TxtFAGPostingDate.Value = CDate(mFAGToDate)
   End If
   

ElseIf LCase(CmbStatus.Text) = "pending" Then
   mStatus = "P"
   TxtFAGRemarks.Locked = True
   TxtFAGPostingDate.Visible = False
   Label11.Visible = False
ElseIf LCase(CmbStatus.Text) = "rejected" Then
   mStatus = "R"
   TxtFAGRemarks.Locked = False
   TxtFAGPostingDate.Visible = True
   Label11.Visible = True
   TxtFAGPostingDate.Enabled = True
   TxtFAGPostingDate.Value = Date
ElseIf LCase(CmbStatus.Text) = "freeze - dont generate invoice" Then
   mStatus = "F"
   TxtFAGRemarks.Locked = False
   TxtFAGPostingDate.Visible = True
   Label11.Visible = True
   TxtFAGPostingDate.Enabled = True
   TxtFAGPostingDate.Value = Date
Else
   MsgBox "Invalid selection!!!"
   CmbStatus.SetFocus
   TxtFAGRemarks.Locked = True
   TxtFAGPostingDate.Visible = False
   Label11.Visible = False
   Exit Sub
End If


'If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fag admin" And LCase(Gen_UserRole) <> "power" Then
'   If mStatus = "A" Or mStatus = "F" Then
'      MsgBox "Only FAG user have access to approve/freeze Ad-Hoc Invoice '!!!"
'      CmbStatus.SetFocus
'      Exit Sub
'   End If
'End If

If mStatus = "A" Then
   If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
      MsgBox "Only FAG user have access to approve Ad-Hoc Invoice '!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

If mStatus = "F" Then
   If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" Then
      MsgBox "Only Head CMG / Super Admin user have access to freeze Ad-Hoc Invoice '!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If


End Sub

Private Sub CmdCancelReverse_Click()
TxtReverseDate.Value = Date
TxtReverseRemark = ""
Frame6.Visible = False
End Sub

Private Sub CmdClient_Click()
FrmMst_Client.Show
End Sub
Private Sub CmdCommodityID_Click()
FrmMst_Commodity.Show
End Sub

Private Sub CmdGo_Click()

Dim wc As Variant

wc = ""

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' should be less than 'To Date' !!"
   SFrom.SetFocus
   Exit Sub
End If

Call Grid_Head

mTotalDays = Day(DateSerial(Year(STo.Value), Month(STo.Value) + 1, 0))

If Option2.Value = True Then
   mSFromDate = Format(Year(SFrom.Value) & "-" & Month(SFrom.Value) & "-" & "01", "yyyy-mm-dd")
   mSToDate = Format(Year(STo.Value) & "-" & Month(STo.Value) & "-" & mTotalDays, "yyyy-mm-dd")
Else
   mSFromDate = SFrom.Value  'Format(Year(SFrom.Value) & "-" & Month(SFrom.Value) & "-" & "01", "yyyy-mm-dd")
   mSToDate = STo.Value  'Format(Year(STo.Value) & "-" & Month(STo.Value) & "-" & "01", "yyyy-mm-dd")
End If


tmp = "Select TI.TIID,TI.TINo,TI.TIDate,TI.TIAmount,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode,TI.TISTAmount ,TI.PIID,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+isnull(TI.AdjAmountFAG,0) RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID ) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDS,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID )+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID  )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Where TI.PIFLag in ('ADHOC')  "

Call Tmp1Table



tmp = "Select "
tmp = tmp & " TTI.PINo,TTI.PIDate,TTI.InvoiceDate,MI.IEType,ML.LocationName,ME.ExchangeType,MComm.Commodity,"
tmp = tmp & " MC.ClientName + '~' + MC.ClientIDRow as 'ClientName',"
tmp = tmp & " MD.DepositoryID,MD.DepositoryName,TTI.Amount,MST.Description,"
tmp = tmp & " TTI.BatchID, TTI.Remarks, TTI.FAGRemarks, TTI.FAGPostingDate, TTI.G_UID, TTI.Flag, "
tmp = tmp & " TTD.TIID,TTD.TINo,TTD.TIDate,TTD.TIAmount,TTD.TISTAmount,TTD.EntryNo,'' ReceivedAmount,''ReceivedTDS,''REceivedOtherDed,''PostInvoiceNo,TTI.ReverseFlag,TTI.ReverseRemark,TTI.ReverseDate,TTI.RBatchID,MD1.DepositoryName+ '~'+ MD1.DepositoryID as DPID,MSC.ContractName,MSI.StateIndiaName ,MSI.GSTIN 'NCMLGSTNo',MCSG.GSTNo 'ClientGSTNo',MCSG.GSTNAVCode,MDSG.GSTNo 'DPGSTNo',MDSG.GSTNAVCode 'DPGSTNAVCode','' 'TaxApplied',MCSG.RGSTNAVCode,TTI.IETypeID "
tmp = tmp & " From Txn_TempInvoice TTI"
tmp = tmp & " Inner Join Mst_IncomeExpenseType MI On TTI.IETypeID = MI.IETypeID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TTI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Left Join Mst_Client MC On TTI.ClientID = MC.ClientID"
tmp = tmp & " Left Join Mst_Depository MD On MC.DepositoryID = MD.DepositoryID"
tmp = tmp & " Left Join Mst_Depository MD1 On TTI.DPID = MD1.DepositoryID"
tmp = tmp & " Inner Join Mst_Commodity MComm On TTI.CommodityID = MComm.CommodityID"
tmp = tmp & " Inner Join Mst_Location ML On TTI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ServiceTax MST On TTI.ServiceTaxID = MST.ServiceTaxID"
tmp = tmp & " Left Join Txn_TIDetails TTD On TTI.PINO =  TTD.PIID And PIFlag = 'ADHoc' "
tmp = tmp & " Left Join Mst_StorageContract MSC On TTI.ContractID =  MSC.ContractID "
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID "
tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID "
tmp = tmp & " Left  Join Mst_ClientStateWiseGST MCSG On MC.ClientID = MCSG.ClientID ANd MSI.StateIndiaID = MCSG.StateINdiaID "
tmp = tmp & " Left  Join Mst_DepositoryStateWiseGST MDSG On MD1.DepositoryID = MDSG.DepositoryID ANd MSI.StateIndiaID = MDSG.StateINdiaID "

'If Gen_DBType = "MDB" Then
'   wc = " Where TTI.PIDate >= #" & SFrom.Value & "# And TTI.PIDate <= #" & STo.Value & "# "
'Else
'   wc = " Where TTI.PIDate >= '" & SFrom.Value & "' And TTI.PIDate <= '" & STo.Value & "' "
'End If
            

If Gen_DBType = "MDB" Then
   If Option1.Value = True Then
      wc = " Where TTI.PIDate >= #" & SFrom.Value & "# And TTI.PIDate <= #" & STo.Value & "# "
   Else
      wc = " Where TTI.InvoiceDate >= #" & mSFromDate & "# And TTI.InvoiceDate <= #" & mSToDate & "# "
   End If
Else
'   wc = " Where TTI.PIDate >= '" & SFrom.Value & "' And TTI.PIDate <= '" & STo.Value & "' "
   If Option1.Value = True Then
      wc = " Where TTI.PIDate >= '" & SFrom.Value & "' And TTI.PIDate <= '" & STo.Value & "' "
   Else
      wc = " Where TTI.InvoiceDate >= '" & mSFromDate & "' And TTI.InvoiceDate <= '" & mSToDate & "' "
   End If

End If
            

If CmbSLocationID.Text <> "" Then
   wc = wc & " And  ML.LocationName = '" & CmbSLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And ML." & Gen_WcLocation
   End If
End If
      
If CmbSInvoiceType.Text <> "" Then
   wc = wc & " And MI.IEType = '" & CmbSInvoiceType.Text & "'"
End If

If CmbSExchangeTypeID.Text <> "" Then
   wc = wc & " And ME.ExchangeType = '" & CmbSExchangeTypeID.Text & "'"
End If

If CmbSCommodityID.Text <> "" Then
   wc = wc & " And MComm.Commodity = '" & CmbSCommodityID.Text & "'"
End If

If CmbSStatus.Text <> "" Then
   wc = wc & " And  TTI.Flag = '" & Left(CmbSStatus.Text, 1) & "'"
End If

''-- Ullhas Patil 27 Sep 2017
If CmbClient.Text <> "" Then
    wc = wc & " And MC.ClientID = " & mSClientRowID
End If
''------

If TxtSAINo <> "" Then
    wc = wc & " And TTI.PINo=" & Val(TxtSAINo)
End If
            
If ChkBatchNo.Value = 0 Then
   wc = wc & " And TTI.BatchID = 0 " '" & Val(TxtSAINo)
End If
            
            
tmp1 = " Order By ML.LocationName"
tmp = tmp & " " & wc & " " & tmp1

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
       For C = 0 To TmpRs.Fields.Count - 3
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       If TmpRs!IETypeID = 4 Or TmpRs!IETypeID = 11 Then
          MSHFlexGrid1.TextMatrix(I, 37) = IIf(IsNull(TmpRs!RGSTNavCode), "", TmpRs!RGSTNavCode)
       End If
       MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 2) = Format(MSHFlexGrid1.TextMatrix(I, 2), "mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 20) = Format(MSHFlexGrid1.TextMatrix(I, 20), "dd-mmm-yyyy")
       If MSHFlexGrid1.TextMatrix(I, 17) = "P" Then
          MSHFlexGrid1.TextMatrix(I, 17) = "Pending"
       ElseIf MSHFlexGrid1.TextMatrix(I, 17) = "A" Then
          MSHFlexGrid1.TextMatrix(I, 17) = "Approved"
       ElseIf MSHFlexGrid1.TextMatrix(I, 17) = "R" Then
          MSHFlexGrid1.TextMatrix(I, 17) = "Rejected"
       ElseIf MSHFlexGrid1.TextMatrix(I, 17) = "F" Then
          MSHFlexGrid1.TextMatrix(I, 17) = "Freeze - dont generate invoice"
       End If
       If MSHFlexGrid1.TextMatrix(I, 28) = "Y" Then
          MSHFlexGrid1.TextMatrix(I, 28) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 28) = "No"
       End If
       If MSHFlexGrid1.TextMatrix(I, 30) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 30) = Format(MSHFlexGrid1.TextMatrix(I, 30), "dd-mm-yyyy")
       End If
       
       If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
       TmpRs1.Find "PIID=" & TmpRs!PINo
       If TmpRs1.EOF = False Then
          
          MSHFlexGrid1.TextMatrix(I, 24) = TmpRs1!RecAmount
          MSHFlexGrid1.TextMatrix(I, 25) = TmpRs1!RecTDS
          MSHFlexGrid1.TextMatrix(I, 26) = TmpRs1!RecOtherDed
          
          MSHFlexGrid1.TextMatrix(I, 27) = IIf(IsNull(TmpRs1!PostInvoiceNo), "", TmpRs1!PostInvoiceNo)
       End If
       If Trim(MSHFlexGrid1.TextMatrix(I, 35)) <> "" And Trim(MSHFlexGrid1.TextMatrix(I, 36)) <> "" Then
          If Left(MSHFlexGrid1.TextMatrix(I, 35), 2) = Left(MSHFlexGrid1.TextMatrix(I, 36), 2) Then
             MSHFlexGrid1.TextMatrix(I, 40) = "Intra"
          Else
             MSHFlexGrid1.TextMatrix(I, 40) = "Inter"
          End If
       End If
       If Trim(MSHFlexGrid1.TextMatrix(I, 35)) <> "" And Trim(MSHFlexGrid1.TextMatrix(I, 39)) <> "" Then
          If Left(MSHFlexGrid1.TextMatrix(I, 35), 2) = Left(MSHFlexGrid1.TextMatrix(I, 39), 2) Then
             MSHFlexGrid1.TextMatrix(I, 40) = "Intra"
          Else
             MSHFlexGrid1.TextMatrix(I, 40) = "Inter"
          End If
       End If
       
       TmpRs.MoveNext
   Next
End If
Label15.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdInvoiceType_Click()
FrmMst_IncomeExpenseType.Show
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdReverse_Click()
Frame6.Visible = True
TxtReverseDate.Value = Date
TxtReverseRemark = ""
End Sub

Private Sub CmdSaveReverse_Click()

If TxtReverseDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtReverseDate.SetFocus
   Exit Sub
End If
If TxtReverseDate.Value < TxtFAGPostingDate.Value Then
   MsgBox "'Reverse Date' should not less than 'FAG Posting Date' !!"
   TxtReverseDate.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtReverseDate.Value, "yyyy-mm") Then
   MsgBox "Reverse Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtReverseDate.SetFocus
   Exit Sub
End If

If TxtReverseRemark.Text = "" Then
   MsgBox "Blank Remark not allowed!!"
   TxtReverseRemark.SetFocus
   Exit Sub
End If
TxtReverseRemark = Replace(TxtReverseRemark, ",", "-")
RsTxn_TempInvoice!ReverseDate = TxtReverseDate.Value
RsTxn_TempInvoice!ReverseFlag = "Y"
RsTxn_TempInvoice!UpdatedBy = Gen_UserLoginID
RsTxn_TempInvoice!UpdatedDate = Now
RsTxn_TempInvoice!ReverseRemark = TxtReverseRemark.Text

RsTxn_TempInvoice.Update

tmp = " Update Txn_TIDetails Set TIAmount=0,ReverseFlag='Y',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where PIID=" & RsTxn_TempInvoice!PINo & " and PIFLag='ADHOC'"

Call TmpTable

Frame6.Visible = False
CmdReverse.Enabled = False

End Sub

Private Sub CmdServiceTaxID_Click()
FrmMst_ServiceTax.Show
End Sub


Private Sub CmdViewInvoice_Click()

Dim fso As New FileSystemObject

On Error GoTo Error:
'Shell (App.Path & "\Invoices\" & TxtEntryNo & " ")
If mEntryNo <> "" Then
   wsPathAdobe = Gen_ReadPathConfigINI("PDFPATH") '"\\10.0.0.8\Reader\AcroRD32.exe" '"C:\Program Files\Adobe\Acrobat 6.0\Reader\AcroRd32.exe" '"\\10.0.0.8\Reader\AcroRD32.exe"
   mFileName = App.Path & "\NavisionInvoice\" & mEntryNo & ".pdf"
   'mFileName = App.Path & "\Invoices\" & Replace(TxtPostInvoiceNo, "/", "") & ".pdf"
'   If wsPathAdobe <> "" Then
'      iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
'
'
'   End If
   
   If fso.FileExists(mFileName) = True Then
      If wsPathAdobe <> "" Then
         iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
         Call Log_Print_History(mSM_Prefix, "ADHOC", TxtPIID, mEntryNo)
      End If
   Else
      MsgBox "Invoice not found. Please contact CMG Department for same!!!"
   End If
   
End If
Exit Sub

Error:
      MsgBox "Please contact Technology department!!!"

End Sub

Private Sub EditCmd_Click()

If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "freeze - dont generate invoice" Then
   MsgBox "Ad-Hoc Invoice " & TxtStatus & ". Can't edit!!!"
   Exit Sub
End If

Edit_Flg = "E"

Call GButtonLock(Me, False)

CmbInvoiceType.Text = TxtInvoiceType
TxtInvoiceType.Visible = False
CmbInvoiceType.Visible = True
'CmbExchangeTypeID.Text = TxtExchangeTypeID
'TxtExchangeTypeID.Visible = False
CmbExchangeTypeID.Visible = True
CmbLocationID.Text = TxtLocationID
TxtLocationID.Visible = False
CmbLocationID.Visible = True
CmbCommodityID.Text = TxtCommodityID
TxtCommodityID.Visible = False
CmbCommodityID.Visible = True
CmbMonth.Text = TxtMonth
TxtMonth.Visible = False
CmbMonth.Visible = True
CmbYear.Text = TxtYear
TxtYear.Visible = False
CmbYear.Visible = True
CmdReverse.Enabled = False
CmbDPName.Enabled = True
Frame6.Visible = False
If LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "superadmin" Then
   CmbServiceTaxID.Visible = True
   TxtServiceTaxID.Visible = False
   CmbServiceTaxID.Text = TxtServiceTaxID
   CmdServiceTaxID.Visible = True
End If


'CmbServiceTaxID.Text = TxtServiceTaxID
'TxtServiceTaxID.Visible = False
'CmbServiceTaxID.Visible = True

CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = TxtStatus
TxtFAGRemarks.Locked = True
TxtRemark.Locked = False
TxtAmount.Locked = False
TxtClientID.Locked = False
TxtQty.Locked = False
If mInvoiceTypeID = 4 Then
   CmbMSC.Visible = True
   TxtMSC.Visible = False
   CmbMSC.Text = TxtMSC.Text
   Frame2.Enabled = False
Else
'   Check1.Visible = True
'   Check1.Value = 0
'   If MSHFlexGrid3.Rows < 4 Then
'      Call GetConnectedGodowns
'      Call DistributeGodownIncome
'   End If
   Frame2.Enabled = True
End If

CmdViewInvoice.Enabled = False

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
SFrom.Value = Date - 365
STo.Value = Date
CmbSInvoiceType.Text = ""
CmbSExchangeTypeID.Text = ""
CmbSLocationID.Text = ""
CmbSStatus.Text = ""
CmbSCommodityID.Text = ""
TxtSAINo = ""
Label15.Caption = ""
Option1.Value = True

If Gen_WcLocation <> "" Then
   Call TableTxn_TempInvoice("Where " & Gen_WcLocation)
Else
   Call TableTxn_TempInvoice(" Order by PINo")
End If
If RsTxn_TempInvoice.RecordCount = 0 Then
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
Private Sub Indata()

Call GButtonLock(Me, True)
CmdReverse.Enabled = False
Frame6.Visible = False
CmdViewInvoice.Enabled = True
Check1.Visible = False
Frame2.Enabled = False

tmp = " Select TTI.PINo,TTI.PIDate,TTI.IETypeID,TTI.ExchangeTypeID,TTI.LocationID,TTI.CommodityID,TTI.ClientID, "
tmp = tmp & " TTI.InvoiceDate,TTI.ServiceTaxID,TTI.Amount,TTI.Flag,TTI.BatchID,TTI.Remarks,TTI.FAGRemarks,TTI.FAGPostingDate,"
tmp = tmp & " MI.IEType , Me.ExchangeType, ML.LocationName, MC.ClientName, MComm.Commodity, TTI.G_UID,MST.Description,TTI.Quantity,TTI.DPID, MD.NAVCode,MSC.ContractName  + '~' + convert(varchar(30),MSC.ContractID)  as ContractName ,TTI.ContractID,ML.SM_Prefix"
tmp = tmp & " From Txn_TempInvoice TTI"
tmp = tmp & " Inner Join Mst_IncomeExpenseType MI On TTI.IETypeID = MI.IETypeID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TTI.ExchangeTypeID = ME.ExchangeTypeID"
'tmp = tmp & " Inner Join Mst_Client MC On TTI.ClientIDRow = MC.ClientIDRow"
tmp = tmp & " Left Join Mst_Client MC On TTI.ClientID = MC.ClientID"
tmp = tmp & " Left Join Mst_depository MD On TTI.DPID = MD.DepositoryID"
tmp = tmp & " Inner Join Mst_Commodity MComm On TTI.CommodityID = MComm.CommodityID"
tmp = tmp & " Inner Join Mst_Location ML On TTI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ServiceTax MST On TTI.ServiceTaxID = MST.ServiceTaxID"
tmp = tmp & " Left Join Mst_StorageContract MSC On TTI.ContractID = MSC.ContractID"
tmp = tmp & " Where TTI.PINo = " & RsTxn_TempInvoice!PINo & ""

Call TmpTable

CmbMSC.Visible = False
TxtMSC.Visible = True

TxtPIID = TmpRs!PINo
TxtPIDate.Value = TmpRs!PIDate
TxtPIDate.Enabled = False
mSM_Prefix = TmpRs!SM_Prefix
CmbInvoiceType.Visible = False
TxtInvoiceType.Visible = True
TxtInvoiceType.Text = TmpRs!IEType
mInvoiceTypeID = TmpRs!IETypeID
CmdInvoiceType.Visible = False
If mInvoiceTypeID = 4 Then
   mContractID = IIf(IsNull(TmpRs!ContractID), Null, TmpRs!ContractID)
   TxtMSC.Text = IIf(IsNull(TmpRs!ContractName), "", TmpRs!ContractName)
Else
   TxtMSC.Text = ""
End If
'CmbExchangeTypeID.Visible = False
'TxtExchangeTypeID.Visible = True
TxtExchangeTypeID.Text = TmpRs!ExchangeType
mExchangeTypeID = TmpRs!ExchangeTypeID

CmbLocationID.Visible = False
TxtLocationID.Visible = True
TxtLocationID.Text = TmpRs!Locationname
CmdLocationID.Visible = False
mLocationID = TmpRs!LocationID

CmbCommodityID.Visible = False
TxtCommodityID.Visible = True
TxtCommodityID.Text = TmpRs!Commodity
CmdCommodityID.Visible = False
mCommodityID = TmpRs!CommodityID

CmbMonth.Visible = False
TxtMonth.Visible = True

mMonthListindex = Month(TmpRs!InvoiceDate) - 1
TxtMonth.Text = MonthName(Month(TmpRs!InvoiceDate))

CmbYear.Visible = False
TxtYear.Visible = True
TxtYear.Text = Year(TmpRs!InvoiceDate)

CmbServiceTaxID.Visible = False
TxtServiceTaxID.Visible = True
TxtServiceTaxID.Text = TmpRs!Description
CmdServiceTaxID.Visible = False
mServiceTaxID = TmpRs!ServiceTaxID

CmbStatus.Visible = False
TxtStatus.Visible = True
mStatusFlag = TmpRs!Flag

If LCase(mStatusFlag) = "p" Then
   TxtStatus.Text = "Pending"
ElseIf LCase(mStatusFlag) = "a" Then
   TxtStatus.Text = "Approved"
ElseIf LCase(mStatusFlag) = "r" Then
   TxtStatus.Text = "Rejected"
ElseIf LCase(mStatusFlag) = "f" Then
   TxtStatus.Text = "Freeze - dont generate invoice"
End If

TxtNAVClientName = ""
TxtClientID.Locked = True
CmbDPName.Visible = False
CmdClient.Visible = False
If LCase(TxtInvoiceType.Text) <> "dp-client" And LCase(TxtInvoiceType.Text) <> "commtrack income" Then

   mClientIDRow = IIf(IsNull(TmpRs!ClientID), 0, TmpRs!ClientID)
   ans = GetClientNameByRow(mClientIDRow)
   TxtClientID = ans(0)
   TxtClientName = ans(1) 'TmpRs!ClientName
   TxtDPID = ans(2)  'TmpRs!DepositoryID
   TxtDPName = ans(3) 'TmpRs!DepositoryName
   TxtNAVClientName = ans(5)
   TxtNAVCode.Text = ans(6)
   TxtRNAVCode.Text = ans(7)
   CmbClientDP.Text = ""
   CmbClientDP.Visible = False
   CmbNonNcdexClient.Visible = False
   If mExchangeTypeID = 1 Then
      CmbNonNcdexClient.Text = TxtClientName
   End If
Else
   CmbDPName.Visible = True
   mDPID = TmpRs!DPID
   CmbDPName.Text = GetDepositoryName(mDPID, 2) & "~" & mDPID
   CmbDPName.Enabled = False
   CmbDPName.Width = 10800 'CmbNonNcdexClient.Width
   CmbDPName.Top = TxtClientID.Top 'CmbNonNcdexClient.Top
   CmbDPName.Left = TxtClientID.Left 'CmbNonNcdexClient.Left
   CmbNonNcdexClient.Visible = False
   TxtNAVClientName = GetNAVDepositoryName(mDPID)
   TxtNAVCode.Text = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
End If

TxtQty.Locked = True
TxtAmount.Locked = True
TxtRemark.Locked = True
TxtFAGRemarks.Locked = True

TxtRemark = TmpRs!Remarks
TxtAmount = TmpRs!Amount
TxtQty = IIf(IsNull(TmpRs!Quantity), 0, TmpRs!Quantity)
TxtFAGRemarks = IIf(IsNull(TmpRs!FAGRemarks), "", TmpRs!FAGRemarks)
If IsNull(TmpRs!FAGPostingDate) = False Then
   TxtFAGPostingDate.Value = TmpRs!FAGPostingDate
   TxtFAGPostingDate.Enabled = False
Else
   TxtFAGPostingDate.Visible = False
   Label11.Visible = False
End If

'TxtTINo = IIf(IsNull(TmpRs!TINo), "", TmpRs!TINo)
'TxtTIDate = IIf(IsNull(TmpRs!TIDate), "", TmpRs!TIDate)
'TxtAmountInclTax = IIf(IsNull(TmpRs!AmountIncl), "", TmpRs!AmountIncl)
TxtBatchNo = IIf(IsNull(TmpRs!BatchID), 0, TmpRs!BatchID)

TxtCreated = IIf(IsNull(RsTxn_TempInvoice!CreatedBy), "", RsTxn_TempInvoice!CreatedBy) & " :- " & IIf(IsNull(RsTxn_TempInvoice!CreatedDate), "", Format(RsTxn_TempInvoice!CreatedDate, "dd-mmm-yyyy hh:mm:ss"))
TxtUpdated = IIf(IsNull(RsTxn_TempInvoice!UpdatedBy), "", RsTxn_TempInvoice!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_TempInvoice!UpdatedDate), "", Format(RsTxn_TempInvoice!UpdatedDate, "dd-mmm-yyyy hh:mm:ss"))

Call InDataTIDetails

If mInvoiceTypeID = 4 And IsNull(mContractID) = False Then
   Call InDataReservationGodowns
Else
   Call InDataGodowns
End If
Frame2.Enabled = False

If RsTxn_TempInvoice.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub


Private Sub MSHFlexGrid1_Click()

Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_TempInvoice.MoveFirst
RsTxn_TempInvoice.Find "PINo = " & MSHFlexGrid1.TextMatrix(I, 0) & ""
If Not RsTxn_TempInvoice.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub
Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid2.RowSel
If r < 2 Then Exit Sub
If MSHFlexGrid2.TextMatrix(r, 1) = "" Then Exit Sub
If mInvoiceTypeID <> 4 Then Exit Sub
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "admin" Then
   mSelRow = MSHFlexGrid2.row
   mSelCol = MSHFlexGrid2.Col
   If mSelCol <> 6 Then Exit Sub
   TxtDistributedAmt.Visible = True
   TxtDistributedAmt.Text = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.row, 6)
   With TxtDistributedAmt
        .Left = (MSHFlexGrid2.Left + MSHFlexGrid2.CellLeft) - 30
        .Top = (MSHFlexGrid2.Top + MSHFlexGrid2.CellTop) - 30
        .Width = MSHFlexGrid2.CellWidth + 35
        .Height = MSHFlexGrid2.CellHeight '- 200
        .ZOrder
        .SetFocus
   End With
End If
End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid3.RowSel
If r < 2 Then Exit Sub
If MSHFlexGrid3.TextMatrix(r, 0) = "" Then Exit Sub
If mInvoiceTypeID = 4 Then Exit Sub

CmbCMPName.Text = MSHFlexGrid3.TextMatrix(r, 0)
mCMPID = GetCMPID(MSHFlexGrid3.TextMatrix(r, 0))
CmbGodownNo.Text = MSHFlexGrid3.TextMatrix(r, 1)
TxtDistIncomeAmount = Val(MSHFlexGrid3.TextMatrix(r, 6))
mGodownID = MSHFlexGrid3.TextMatrix(r, 7)

CmdAddtoList.Caption = "Update List"
CmdRemoveFromList.Enabled = True
CmbCMPName.SetFocus

End Sub

Private Sub TxtAmount_GotFocus()
tmp1 = TxtAmount
End Sub

Private Sub TxtAmount_LostFocus()

If TxtAmount = "" Then Exit Sub
If IsNumeric(TxtAmount) = False Or Val(TxtAmount) < 0 Then
   MsgBox "Invalid Numeric format!!!"
   TxtAmount.SetFocus
   Exit Sub
End If

If mInvoiceTypeID = 4 Then
   If TxtAmount <> tmp1 Then
      Call DistributeReservationIncome
   End If
Else
   If TxtAmount <> tmp1 Then
      Call DistributeGodownIncome
   End If
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
   CmbCommodityID.Text = ""
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   CmbNonNcdexClient.Width = 10680
   CmbNonNcdexClient.Top = TxtClientID.Top
   CmbNonNcdexClient.Left = TxtClientID.Left
   Label25.Visible = False
   Label24.Visible = False
   Label27.Visible = False
   CmbNonNcdexClient.Visible = True
Else
   Label25.Visible = True
   Label24.Visible = True
   Label27.Visible = True
   CmbNonNcdexClient.Visible = False
End If
TxtNAVCode.Text = ""
TxtRNAVCode.Text = ""
End Sub

Private Sub CmbServiceTaxID_GotFocus()
tmp = CmbServiceTaxID.Text
Call TableMst_ServiceTax(" Where Flag = 'A' ")
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

End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   Call Grid_Head3
   mLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix


Call ClearFrame

'Call GetConnectedGodowns
'Call DistributeGodownIncome


End Sub
Private Sub ClearFrame()

Call Grid_Head3
CmbCMPName.Text = ""
CmbGodownNo.Text = ""
CmdAddtoList.Caption = "Add To List"

End Sub
Private Sub CmbCommodityID_GotFocus()

If TxtExchangeTypeID.Text = "" Then
   MsgBox "Please Select Invoice Type!!!"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

tmp1 = CmbCommodityID.Text
'Call TableMst_Commodity
tmp = "Select MCD.CommodityID,Commodity,MC.ServiceTaxID from Mst_CommodityDetail MCD"
tmp = tmp & " Inner Join Mst_Commodity MC on MCD.CommodityID=MC.CommodityID Where ExchangeTypeID=" & mExchangeTypeID

Call Tmp1Table

CmbCommodityID.Clear
If TmpRs1.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
    CmbCommodityID.AddItem TmpRs1!Commodity
    TmpRs1.MoveNext
Next
CmbCommodityID.Text = tmp1
    
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If

TmpRs1.MoveFirst
TmpRs1.Find "Commodity = '" & CmbCommodityID.Text & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If

mCommodityID = TmpRs1!CommodityID
mServiceTaxID = IIf(IsNull(TmpRs1!ServiceTaxID), 0, TmpRs1!ServiceTaxID)
TxtServiceTaxID.Text = GetServiceTaxDescription(IIf(IsNull(TmpRs1!ServiceTaxID), 0, TmpRs1!ServiceTaxID))

If LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "superadmin" Then
   CmbServiceTaxID.Visible = True
   TxtServiceTaxID.Visible = False
   CmbServiceTaxID.Text = TxtServiceTaxID 'CmbServiceTaxID.Text = ""
   CmdServiceTaxID.Visible = True
End If


'Call TableMst_CommodityDetail("Where CommodityId = " & mCommodityID & " and ExchangeTypeId = " & mExchangeTypeID)
'If RsMst_CommodityDetail.RecordCount = 0 Then
'   MsgBox "No Record Found For Exchange" & CmbExchangeTypeID.Text & " - " & CmbCommodityID.Text
'   CmbCommodityID.SetFocus
'   Exit Sub
'End If

End Sub

Private Sub TxtClientID_GotFocus()
If SaveCmd.Enabled = False Then Exit Sub

If TxtExchangeTypeID.Text = "" Then
   MsgBox "Blank Exchange Type not allowed!!!"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtClientID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub

CmbClientDP.Visible = False
If TxtClientID = "" Then
   TxtNAVCode.Text = ""
   TxtRNAVCode.Text = ""
   Exit Sub
End If

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
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow,  Mst_Client.NAVClientName, Mst_Client.NAVCode, Mst_Client.RNAVCode " & _
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
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
TxtNAVClientName = IIf(IsNull(TmpRs1!NAVClientName), "", TmpRs1!NAVClientName)
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientID
TxtNAVCode.Text = IIf(IsNull(TmpRs1!NavCode), "", TmpRs1!NavCode)
TxtRNAVCode.Text = IIf(IsNull(TmpRs1!RNAVCode), "", TmpRs1!RNAVCode)
End Sub

Private Sub CmbNonNcdexClient_GotFocus()
If TxtExchangeTypeID.Text = "" Then Exit Sub
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where Status = 'A' And ExchangeTypeId = " & mExchangeTypeID)
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
Dim ans As Variant
ans = GetClientNameByRow(RsMst_Client!ClientID)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)
TxtNAVClientName = ans(5)
TxtNAVCode.Text = ans(6)
TxtRNAVCode.Text = ans(7)
End Sub

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For I = 2011 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
flg = "N"
For I = 2011 To Year(Now)
    If CmbYear.Text = I Then
       flg = "Y"
       Exit For
    End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbYear.SetFocus
   Exit Sub
End If

'Call GetConnectedGodowns


End Sub

Private Sub SaveCmd_Click()

If CmbInvoiceType.Text = "" Then
   MsgBox "Blank Invoice Type not allowed!!!"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

If TxtExchangeTypeID.Text = "" Then
   MsgBox "Blank Exchange Type not allowed!!!"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Blank Commodity not allowed!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If CmbNonNcdexClient.Visible = True Then
   If CmbNonNcdexClient.Text = "" Then
      MsgBox "Blank Client not allowed!!!"
      CmbNonNcdexClient.SetFocus
      Exit Sub
   End If
Else
   If CmbDPName.Visible = True And CmbDPName.Text = "" Then
      MsgBox "Blank DP Name not allowed!!!"
      CmbDPName.SetFocus
      Exit Sub
   ElseIf CmbDPName.Visible = False Then
      If TxtClientID = "" Then
         MsgBox "Blank Client not allowed!!!"
         TxtClientID.SetFocus
         Exit Sub
      End If
   End If
End If

If CmbMonth.Text = "" Then
   MsgBox "Blank Month not allowed!!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox "Blank Year not allowed!!!"
   CmbYear.SetFocus
   Exit Sub
End If

If Edit_Flg = "E" Then
   If CmbStatus.Text = "" Then
      MsgBox "Blank Status not allowed!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

If TxtServiceTaxID.Text = "" Then
   MsgBox "Blank GST Group Code not allowed for selected Commodity!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If TxtAmount = "" Then
   MsgBox "Blank Amount not allowed!!!"
   TxtAmount.SetFocus
   Exit Sub
End If

If LCase(CmbStatus.Text) <> "pending" Then
   If TxtFAGRemarks = "" Then
      MsgBox " Blank FAG Remarks not allowed!!! "
      TxtFAGRemarks.SetFocus
      Exit Sub
   End If
      
     
   If TxtFAGPostingDate.Value > Date Then
      MsgBox "Future date not allowed!!"
      TxtFAGPostingDate.SetFocus
      Exit Sub
   End If
     
'Commented on 05 Sept 2011 as required by Nishikant and order by Mayur Sir 12 22 PM
     
'   If TxtFAGPostingDate.Value < TxtPIDate.Value Then
'      MsgBox "'FAG Posting Date' must be greater than 'PI Date'"
'      TxtFAGPostingDate.SetFocus
'      Exit Sub
'   End If

End If
   
mMonth = mMonthListindex + 1
If mMonth < 10 Then
   mMonth = "0" & mMonth
End If
mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
If CDate(Format(mFromDate, "yyyymm")) > Format(Date, "yyyymm") Then
   MsgBox "Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If


'COmmenetd on 18th April 2013
'If LCase(CmbInvoiceType.Text) = "reservation income" Then
'   If Format(CDate(mFromDate), "yyyymm") > 201303 Then
'      MsgBox "Invoicing from April 2013 not allowed via this module!!!"
'      CmbMonth.SetFocus
'      Exit Sub
'   End If
'End If


If LCase(CmbStatus.Text) <> "rejected" Then
   If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(mFromDate), "yyyy-mm") Then
      MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
      CmbMonth.SetFocus
      Exit Sub
   End If
End If

If CDate(Format(mFromDate, "yyyymm")) < "201104" Then
   MsgBox "Invoicing for months previous to april 2011 not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If mInvoiceTypeID = 4 Then
   If LCase(CmbStatus.Text) <> "rejected" Then
      If CmbMSC.Text = "" Then
         MsgBox " Please select contract!!"
         If CmbMSC.Visible = True Then
            CmbMSC.SetFocus
         End If
         Exit Sub
      End If
      If Val(TxtAmount) <> Val(MSHFlexGrid2.TextMatrix(0, 6)) Then
         MsgBox "Invoice Amount and Distributed amount mismatch. Please check !!!"
         Exit Sub
      End If
   End If
Else
   If LCase(CmbStatus.Text) <> "rejected" Then
      If CDate(Format(mFromDate, "yyyymm")) > "201402" Then
         If Val(TxtAmount) <> Val(MSHFlexGrid3.TextMatrix(0, 6)) Then
            MsgBox "Invoice Amount and Distributed amount mismatch. Please check !!!"
            Exit Sub
         End If
      End If
   End If
End If

If LCase(CmbInvoiceType.Text) = "nnnr" Then
   tmp = "Select * from Txn_ProInvoice Where SM_Prefix = '" & mSM_Prefix & "' And  BillingMonth >= '" & Format(mFromDate, Gen_DatFormat) & "' " ''ClientIDRow = " & mClientIDRow & " And
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      If MsgBox("Invoice is already generated for this location for the selected month. This will be taken as an additional Invoice. Press 'yes' to generate this adhoc invoice!!!", vbYesNo) = vbNo Then
         CmbMonth.SetFocus
         Exit Sub
      End If
   End If
Else
'   tmp = "Select * from Txn_OtherIncome Where LocationID = " & mLocationID & " And IETypeID = " & mInvoiceTypeID & " And TxnDate >= '" & Format(mFromDate, Gen_DatFormat) & "' "  ''And ClientName = '" & TxtClientName & "'
'   Call TmpTable
'   If TmpRs.RecordCount > 0 Then
'      If MsgBox("Income Uploaded by CMG Dept. for selected location, still you wan't to create Invoice!!!", vbYesNo) = vbNo Then
'         CmbMonth.SetFocus
'         Exit Sub
'      End If
'   End If
End If




If Edit_Flg = "A" Then
   RsTxn_TempInvoice.AddNew
   RsTxn_TempInvoice!PINo = GetMaxPINo
   RsTxn_TempInvoice!G_UID = GetGUID
   TxtPIID.Text = RsTxn_TempInvoice!PINo

Else
   RsTxn_TempInvoice.MoveFirst
   RsTxn_TempInvoice.Find " PINo = '" & Val(TxtPIID.Text) & "'"
End If
mMonth = mMonthListindex + 1
If mMonth < 10 Then
   mMonth = "0" & mMonth
End If
mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
RsTxn_TempInvoice!InvoiceDate = CDate(mFromDate)
RsTxn_TempInvoice!PIDate = Format(TxtPIDate.Value, Gen_DatFormat)
RsTxn_TempInvoice!IETypeID = mInvoiceTypeID
RsTxn_TempInvoice!LocationID = mLocationID
RsTxn_TempInvoice!ExchangeTypeID = mExchangeTypeID
RsTxn_TempInvoice!CommodityID = mCommodityID
If CmbDPName.Visible = True Then
   RsTxn_TempInvoice!DPID = mDPID
   RsTxn_TempInvoice!ClientID = Null
Else
   RsTxn_TempInvoice!ClientID = mClientIDRow
   RsTxn_TempInvoice!DPID = Null
End If
If mInvoiceTypeID = 4 Then
   RsTxn_TempInvoice!ContractID = mContractID
Else
   RsTxn_TempInvoice!ContractID = Null
End If

RsTxn_TempInvoice!BatchID = 0
RsTxn_TempInvoice!RBatchID = 0
RsTxn_TempInvoice!ServiceTaxID = mServiceTaxID
RsTxn_TempInvoice!Amount = Format(Val(TxtAmount.Text), "#####0.00")

RsTxn_TempInvoice!Quantity = Format(Val(TxtQty.Text), "#####0.000")

RsTxn_TempInvoice!Flag = Left(CmbStatus.Text, 1)
RsTxn_TempInvoice!Remarks = TxtRemark.Text

If LCase(CmbStatus) <> "pending" Then
   RsTxn_TempInvoice!FAGRemarks = TxtFAGRemarks.Text
   RsTxn_TempInvoice!FAGPostingDate = TxtFAGPostingDate
Else
   RsTxn_TempInvoice!FAGRemarks = Null
   RsTxn_TempInvoice!FAGPostingDate = Null
End If

If IsNull(RsTxn_TempInvoice!CreatedBy) = True Or RsTxn_TempInvoice!CreatedBy = "" Then
   RsTxn_TempInvoice!CreatedBy = Gen_UserLoginID
   RsTxn_TempInvoice!CreatedDate = Now
Else
   RsTxn_TempInvoice!UpdatedBy = Gen_UserLoginID
   RsTxn_TempInvoice!UpdatedDate = Now
End If

RsTxn_TempInvoice.Update

If mInvoiceTypeID = 4 Then
   Call SaveReservationGodownDetails
   tmp = "Delete from Txn_TempInvGodownDetails Where PINo = " & TxtPIID & ""
   Call TmpTable
Else
   Call SaveGodownDetails
   tmp = "Delete from Txn_TempInvSRGodownDetails Where PINo = " & TxtPIID & ""
   Call TmpTable
End If

Call GButtonLock(Me, True)
Call Indata

End Sub

Private Function GetMaxPINo() As Double
Dim Retval As Double
tmp = "Select max(PINo) from Txn_TempInvoice"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxPINo = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_TempInvoice.MoveNext
If RsTxn_TempInvoice.EOF Then
   RsTxn_TempInvoice.MoveLast
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
RsTxn_TempInvoice.MovePrevious
If RsTxn_TempInvoice.BOF Then
   RsTxn_TempInvoice.MoveFirst
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

Private Sub LastCmd_Click()
RsTxn_TempInvoice.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub FirstCmd_Click()
RsTxn_TempInvoice.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_TempInvoice.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_TempInvoice.MoveFirst
   RsTxn_TempInvoice.Find "PINo = " & Val(TxtPIID) & ""
   If RsTxn_TempInvoice.EOF Then
      RsTxn_TempInvoice.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 50
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 350
MSHFlexGrid1.Height = Frame0.Height - 1700

RsTxn_TempInvoice.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'SFrom.Value = Date - 365
'STo.Value = Date
'CmbSInvoiceType.Text = ""
'CmbSExchangeTypeID.Text = ""
'CmbSLocationID.Text = ""
'CmbSStatus.Text = ""
'CmbSCommodityID.Text = ""
'TxtSAINo = ""
'Label15.Caption = ""
'Option1.Value = True
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 41 '39 '34 '30

MSHFlexGrid1.TextMatrix(0, 0) = "PI No"
MSHFlexGrid1.TextMatrix(0, 1) = "PI Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Invoice Date"
MSHFlexGrid1.TextMatrix(0, 3) = "Invoice Type"
MSHFlexGrid1.TextMatrix(0, 4) = "Location"
MSHFlexGrid1.TextMatrix(0, 5) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 6) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 7) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 8) = "DP ID"
MSHFlexGrid1.TextMatrix(0, 9) = "DP Name"
MSHFlexGrid1.TextMatrix(0, 10) = "Amount Excl. Tax"
MSHFlexGrid1.TextMatrix(0, 11) = "GST Group Code"
MSHFlexGrid1.TextMatrix(0, 12) = "Batch ID"
MSHFlexGrid1.TextMatrix(0, 13) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 14) = "FAG Remarks"
MSHFlexGrid1.TextMatrix(0, 15) = "FAGPosting Date"
MSHFlexGrid1.TextMatrix(0, 16) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 17) = "Status"
MSHFlexGrid1.TextMatrix(0, 18) = "TI ID"
MSHFlexGrid1.TextMatrix(0, 19) = "TI No"
MSHFlexGrid1.TextMatrix(0, 20) = "TI Date"
MSHFlexGrid1.TextMatrix(0, 21) = "TI Amount"
MSHFlexGrid1.TextMatrix(0, 22) = "TI Tax Amount"
MSHFlexGrid1.TextMatrix(0, 23) = "TI Entry No"

MSHFlexGrid1.TextMatrix(0, 24) = "Received Amount"
MSHFlexGrid1.TextMatrix(0, 25) = "Received TDS"
MSHFlexGrid1.TextMatrix(0, 26) = "Received Other Ded."
MSHFlexGrid1.TextMatrix(0, 27) = "Post Invoice No"

MSHFlexGrid1.TextMatrix(0, 28) = "Reverse Flag"

MSHFlexGrid1.TextMatrix(0, 29) = "Reverse Remark"
MSHFlexGrid1.TextMatrix(0, 30) = "Reverse Date"
MSHFlexGrid1.TextMatrix(0, 31) = "RBatch ID"

MSHFlexGrid1.TextMatrix(0, 32) = "DP Name and ID "
MSHFlexGrid1.TextMatrix(0, 33) = "Contract Name "

MSHFlexGrid1.TextMatrix(0, 34) = "State India Name"
MSHFlexGrid1.TextMatrix(0, 35) = "NCML GST No"
MSHFlexGrid1.TextMatrix(0, 36) = "Client GST No"
MSHFlexGrid1.TextMatrix(0, 37) = "Client GST NAVCode"
MSHFlexGrid1.TextMatrix(0, 38) = "DP GST No"
MSHFlexGrid1.TextMatrix(0, 39) = "DP GST NAVCode"
MSHFlexGrid1.TextMatrix(0, 40) = "Tax Applied (Inter/Intra)"


MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.ColWidth(14) = 1600
MSHFlexGrid1.ColWidth(15) = 1600
MSHFlexGrid1.ColWidth(16) = 0
MSHFlexGrid1.ColWidth(17) = 1600
MSHFlexGrid1.ColWidth(18) = 1600
MSHFlexGrid1.ColWidth(19) = 1600
MSHFlexGrid1.ColWidth(20) = 1600
MSHFlexGrid1.ColWidth(21) = 1600
MSHFlexGrid1.ColWidth(22) = 1600
MSHFlexGrid1.ColWidth(23) = 1600
MSHFlexGrid1.ColWidth(27) = 1600
'MSHFlexGrid1.ColWidth(18) = 1600
'MSHFlexGrid1.ColWidth(19) = 0
'MSHFlexGrid1.ColWidth(20) = 1600

End Sub

Private Sub TxtDistributedAmt_LostFocus()

If TxtDistributedAmt = "" Then Exit Sub
If IsNumeric(TxtDistributedAmt) = False Or Val(TxtDistributedAmt) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtDistributedAmt.SetFocus
   Exit Sub
End If

MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(mSelRow, 6))
MSHFlexGrid2.TextMatrix(mSelRow, mSelCol) = Val(TxtDistributedAmt)
MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(mSelRow, 6))

TxtDistributedAmt.Visible = False
TxtRemark.SetFocus

End Sub

Private Sub TxtFAGPostingDate_LostFocus()
If TxtFAGPostingDate.Value > Date Then
   MsgBox "Future date not allowed!!"
   TxtFAGPostingDate.SetFocus
   Exit Sub
End If

'Commented temprory on 03 Sept 2011 1232 PM As requested by Paresh mail n order given by Mayur Sir
'If TxtFAGPostingDate.Value < Date - 2 Then
'   MsgBox "Only One Day back date allowed!"
'   TxtFAGPostingDate.SetFocus
'   Exit Sub
'End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtFAGPostingDate.Value, "yyyy-mm") Then
   MsgBox "FAG Posting Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtFAGPostingDate.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtFAGRemarks_LostFocus()
TxtFAGRemarks = Replace(TxtFAGRemarks, ",", "-")
End Sub

Private Sub TxtRemark_LostFocus()
TxtRemark = Replace(TxtRemark, ",", "-")
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If Val(TxtBatchNo) <> 0 Then
   MsgBox "Batch generated for this Invoice!!!"
   Exit Sub
End If

If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "freeze" Then
   MsgBox "Ad-Hoc Invoice " & TxtStatus & ". Can't delete!!!"
   Exit Sub
End If
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   RsTxn_TempInvoice.Delete
   RsTxn_TempInvoice.Update
   If RsTxn_TempInvoice.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_TempInvoice.MoveNext
   If RsTxn_TempInvoice.EOF() Then
      RsTxn_TempInvoice.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_TempInvoice.CancelUpdate
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_TempInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_TempInvoice.xls")
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
VScroll1.Max = Me.ScaleHeight - Frame4.Height '4
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame1.Width '4
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True ''4
If Frame4.Height < Me.Height Then VScroll1.Enabled = False '4
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True '4
If Frame1.Width < Me.Width Then HScroll1.Enabled = False '4
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

Private Sub InDataTIDetails()
'Call TableTxn_TIDetails(" Where PIID= " & RsTxn_TempInvoice!PINO & " And PIFlag='ADHOC'")
'
'If RsTxn_TIDetails.RecordCount > 0 Then
'   TxtTIDate.Text = Format(RsTxn_TIDetails!TIDate, "dd-MMM-yyyy")
'   TxtTINo.Text = RsTxn_TIDetails!TINO
'   TxtAmountInclTax.Text = RsTxn_TIDetails!TIAmount
'   mEntryNo = RsTxn_TIDetails!EntryNo
'   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Then
'      If IIf(IsNull(RsTxn_TempInvoice!ReverseFlag), "N", RsTxn_TempInvoice!ReverseFlag) <> "Y" Then
'         If CheckTIInCR(RsTxn_TIDetails!TIID) = False Then
'            CmdReverse.Enabled = True
'         End If
'      End If
'   End If
'Else
'   TxtTIDate.Text = ""
'   TxtTINo.Text = ""
'   TxtAmountInclTax.Text = ""
'   mEntryNo = ""
'End If
''End If
Dim ans As Variant

ans = GetTIDetails(RsTxn_TempInvoice!PINo, "ADHOC")
TxtTINo = ans(1)
TxtTIDate = Format(ans(2), "dd-MMM-yyyy")
TxtTIAmount = ans(3)
TxtEntryNo = ans(5)
mEntryNo = TxtEntryNo
TxtTISTAmt = ans(7)
TxtPostInvoiceNo = ans(4)
TxtRecAmt = Val(ans(8)) + Val(ans(9)) + Val(ans(10))
If Val(TxtRecAmt) > 0 Then
   TxtBalAmt = Format(Val(TxtTIAmount) - Val(TxtRecAmt), "#####0.00")
Else
   TxtBalAmt = TxtTIAmount
End If

If TxtEntryNo <> "" And (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg") Then
   If IIf(IsNull(RsTxn_TempInvoice!ReverseFlag), "N", RsTxn_TempInvoice!ReverseFlag) <> "Y" Then
      If CheckTIInCR(Val(ans(0))) = False Then
         CmdReverse.Enabled = True
      End If
   End If
End If

End Sub

Private Sub TxtQty_LostFocus()
If TxtQty = "" Then Exit Sub
If IsNumeric(TxtQty) = False Or Val(TxtQty) < 0 Then
   MsgBox "Invalid Numeric format!!!"
   TxtQty.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbDPName_GotFocus()

tmp = CmbDPName.Text

Call TableMst_Depository("Where ExchangeTypeId = " & mExchangeTypeID & " ")

CmbDPName.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
    CmbDPName.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    RsMst_Depository.MoveNext
Next
CmbDPName.Text = tmp

End Sub

Private Sub CmbDPName_LostFocus()

Dim strSplitString() As String
Dim DPID_, DPName_ As String

If CmbDPName.Text = "" Then
   TxtNAVCode.Text = ""
   TxtRNAVCode.Text = ""
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
End If

strSplitString = Split(CmbDPName, "~")
DPName_ = Trim(strSplitString(0))
DPID_ = Trim(strSplitString(1))

RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID = '" & DPID_ & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbDPName.SetFocus
   Exit Sub
End If

mDPID = RsMst_Depository!DepositoryID
'mDPNAVCode = IIf(IsNull(RsMst_Depository!NAVCode), "", RsMst_Depository!NAVCode)
TxtNAVClientName = IIf(IsNull(RsMst_Depository!NAVDepositoryName), "", RsMst_Depository!NAVDepositoryName)
TxtNAVCode = IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)
'TxtRNAVCode = IIf(IsNull(RsMst_Depository!RNAVCode), "", RsMst_Depository!RNAVCode)
End Sub

Private Sub CmbMSC_GotFocus()
If CmbMonth.Text = "" Then
   MsgBox " Blank Month not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Blank Year not allowed  !!!"
   CmbYear.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox " Blank Location not allowed  !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox " Blank Commodity not allowed  !!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If CmbNonNcdexClient.Text = "" Then
   MsgBox " Blank Client not allowed  !!!"
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If CDate(Format(mFromDate, "yyyymm")) > Format(Date, "yyyymm") Then
   MsgBox "Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

tmp1 = CmbMSC.Text

tmp = "select  ContractName  + '~' + convert(varchar(30),ContractID)  as ContractName,LocationID ,ContractID from Mst_StorageContract Where Flag='A'"
tmp = tmp & " And BookedFrom <= '" & mToDate & "' And  ExtendedPeriod >= '" & mFromDate & "'"
tmp = tmp & " And LocationID=" & mLocationID
tmp = tmp & " And (CommodityID=" & mCommodityID & " or  CommodityID is null)"
tmp = tmp & " And ClientID=" & mClientIDRow

Call Tmp4Table

CmbMSC.Clear
If TmpRs4.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs4.RecordCount
    CmbMSC.AddItem TmpRs4!ContractName
    TmpRs4.MoveNext
Next
CmbMSC.Text = tmp1
End Sub
Private Sub CmbMSC_LostFocus()
If CmbMSC.Text = "" Then
   Exit Sub
End If
TmpRs4.MoveFirst
TmpRs4.Find "ContractName = '" & CmbMSC.Text & "'"
If TmpRs4.EOF Then
   MsgBox "Invalid selection "
   CmbMSC.SetFocus
   Exit Sub
End If
If tmp1 <> CmbMSC.Text Then
   mContractID = TmpRs4!ContractID
   Call GetReservationGodowns
   Call DistributeReservationIncome
End If

End Sub

Private Sub Grid_Head2()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 9
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "CMP"
MSHFlexGrid2.TextMatrix(1, 1) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 2) = "Unit Type"
MSHFlexGrid2.TextMatrix(1, 3) = "Capacity in Bags"
MSHFlexGrid2.TextMatrix(1, 4) = "Capacity in MT"
MSHFlexGrid2.TextMatrix(1, 5) = "Reservation Qty"
MSHFlexGrid2.TextMatrix(1, 6) = "Invoice Amount"
MSHFlexGrid2.TextMatrix(1, 7) = "GodownID"
'MSHFlexGrid2.TextMatrix(1, 7) = "SRGodownMapDetailID"
MSHFlexGrid2.TextMatrix(1, 8) = "Address"

MSHFlexGrid2.ColWidth(0) = 3000
MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 1600
MSHFlexGrid2.ColWidth(3) = 1400
MSHFlexGrid2.ColWidth(4) = 1400
MSHFlexGrid2.ColWidth(5) = 1400
MSHFlexGrid2.ColWidth(6) = 1400
'MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(8) = 0

MSHFlexGrid2.RowHeight(1) = 500
End Sub

Private Sub GetReservationGodowns()

Call Grid_Head2

tmp = " Select MC.CMPName,MG.GodownNo,MU.UnitType,MG.BagsCapacity,MG.GodownCapacity, " & _
      " TSRGD.ReservationQty,'' as InvoiceAmt , MG.GodownID, MG.Address " & _
      " From Txn_SRGodownMapDetails TSRGD " & _
      " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID " & _
      " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID " & _
      " Inner Join Mst_UnitType MU On MG.UnitTypeID = MU.UnitTypeID " & _
      " Inner Join Txn_SRGodownMapping TSRG On TSRGD.SRGodownMappingID = TSRG.SRGodownMappingID " & _
      " Where TSRG.ContractID = " & mContractID & " and (MG.CloseDate is null or MG.CloseDate >='" & mFromDate & "') And MG.StartDate<='" & mToDate & "'"
      
Call TmpTable

If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(x, 3))
       MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(MSHFlexGrid2.TextMatrix(x, 4))
       MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(x, 5))
       TmpRs.MoveNext
       x = x + 1
   Next
Else
   MsgBox "No godown mapped to this Contract!!!"
   CmbMSC.SetFocus
   Exit Sub
End If
End Sub

Private Sub SaveReservationGodownDetails()

'If MSHFlexGrid2.Rows > 2 Then
   tmp = "Delete from Txn_TempInvSRGodownDetails Where PINo = " & TxtPIID & ""
   Call TmpTable
      
   Call TableTxn_TempInvSRGodownDetails(" Where PINo = " & TxtPIID & "")
   x = 2
   For I = 1 To MSHFlexGrid2.Rows - 2
       If MSHFlexGrid2.TextMatrix(x, 0) = "" Then Exit For
'       If MSHFlexGrid2.TextMatrix(x, 7) <> "" Then
'          RsTxn_SRGodownMapDetails.Find "SRGodownMapDetailsID = " & MSHFlexGrid2.TextMatrix(x, 7)
'       Else
        RsTxn_TempInvSRGodownDetails.AddNew
        RsTxn_TempInvSRGodownDetails!TIGodownMapDetailsID = GetMaxDetailsID()
        RsTxn_TempInvSRGodownDetails!G_UID = GetGUID
'       End If
       RsTxn_TempInvSRGodownDetails!PINo = TxtPIID
       RsTxn_TempInvSRGodownDetails!GodownID = MSHFlexGrid2.TextMatrix(x, 7)
       RsTxn_TempInvSRGodownDetails!ReservationQty = Val(MSHFlexGrid2.TextMatrix(x, 5))
'       If Val(MSHFlexGrid2.TextMatrix(0, 5)) = 0 Then
'          RsTxn_TempInvSRGodownDetails!InvoiceAmt = Val(TxtAmount) '/ Val(MSHFlexGrid2.Rows - 3) '* Val(MSHFlexGrid2.TextMatrix(x, 5)) / 1 '"'Val(MSHFlexGrid2.TextMatrix(0, 5))
'       Else
'          RsTxn_TempInvSRGodownDetails!InvoiceAmt = Val(TxtAmount) * Val(MSHFlexGrid2.TextMatrix(x, 5)) / Val(MSHFlexGrid2.TextMatrix(0, 5))
'       End If
       RsTxn_TempInvSRGodownDetails!InvoiceAmt = Val(MSHFlexGrid2.TextMatrix(x, 6))
       If IsNull(RsTxn_TempInvSRGodownDetails!CreatedBy) = True Or RsTxn_TempInvSRGodownDetails!CreatedBy = "" Then
          RsTxn_TempInvSRGodownDetails!CreatedBy = Gen_UserLoginID
          RsTxn_TempInvSRGodownDetails!CreatedDate = Now
       Else
          RsTxn_TempInvSRGodownDetails!UpdatedBy = Gen_UserLoginID
          RsTxn_TempInvSRGodownDetails!UpdatedDate = Now
       End If
       x = x + 1
       RsTxn_TempInvSRGodownDetails.Update
Next
'End If
End Sub

Function GetMaxDetailsID() As Double
Dim Retval As Double
Retval = 0
tmp = "Select max(TIGodownMapDetailsID) from Txn_TempInvSRGodownDetails "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetMaxDetailsID = Retval

End Function

Private Sub InDataReservationGodowns()

Call Grid_Head2

tmp = " Select MC.CMPName,MG.GodownNo,MU.UnitType,MG.BagsCapacity,MG.GodownCapacity, " & _
      " TSRGD.ReservationQty, TSRGD.InvoiceAmt, MG.GodownID, MG.Address " & _
      " From Txn_TempInvSRGodownDetails TSRGD " & _
      " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID " & _
      " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID " & _
      " Inner Join Mst_UnitType MU On MG.UnitTypeID = MU.UnitTypeID " & _
      " Inner Join Txn_tempInvoice TSRG On TSRGD.PINO = TSRG.PINO " & _
      " Where TSRG.PINO = " & TxtPIID & " "
      
Call TmpTable

If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(x, 3))
       MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(MSHFlexGrid2.TextMatrix(x, 4))
       MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(x, 5))
       MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(x, 6))
       
       TmpRs.MoveNext
       x = x + 1
   Next
End If
End Sub

Private Sub DistributeReservationIncome()
x = 2
If CmbMSC.Text = "" Then
   Call Grid_Head2
   Exit Sub
End If
MSHFlexGrid2.TextMatrix(0, 6) = ""
For I = 1 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid2.TextMatrix(x, 0) = "" Then Exit For
    If Val(MSHFlexGrid2.TextMatrix(0, 5)) = 0 Then
       MSHFlexGrid2.TextMatrix(x, 6) = Val(TxtAmount) / Val(MSHFlexGrid2.Rows - 3) '* Val(MSHFlexGrid2.TextMatrix(x, 5)) / 1 '"'Val(MSHFlexGrid2.TextMatrix(0, 5))
    Else
       MSHFlexGrid2.TextMatrix(x, 6) = Val(TxtAmount) * Val(MSHFlexGrid2.TextMatrix(x, 5)) / Val(MSHFlexGrid2.TextMatrix(0, 5))
    End If
    MSHFlexGrid2.TextMatrix(x, 6) = Format(MSHFlexGrid2.TextMatrix(x, 6), "#######0.00")
    MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(x, 6))
    x = x + 1
Next

End Sub

Private Sub Grid_Head3()
MSHFlexGrid3.Clear
MSHFlexGrid3.Cols = 9
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 2

MSHFlexGrid3.TextMatrix(1, 0) = "CMP"
MSHFlexGrid3.TextMatrix(1, 1) = "Godown No"
MSHFlexGrid3.TextMatrix(1, 2) = "Unit Type"
MSHFlexGrid3.TextMatrix(1, 3) = "Capacity in Bags"
MSHFlexGrid3.TextMatrix(1, 4) = "Capacity in MT"
MSHFlexGrid3.TextMatrix(1, 5) = "Reservation Qty"
MSHFlexGrid3.TextMatrix(1, 6) = "Invoice Amount"
MSHFlexGrid3.TextMatrix(1, 7) = "GodownID"
'MSHFlexGrid3.TextMatrix(1, 7) = "SRGodownMapDetailID"
MSHFlexGrid3.TextMatrix(1, 8) = "Address"

MSHFlexGrid3.ColWidth(0) = 3000
MSHFlexGrid3.ColWidth(1) = 3000
MSHFlexGrid3.ColWidth(2) = 0
MSHFlexGrid3.ColWidth(3) = 0
MSHFlexGrid3.ColWidth(4) = 0
MSHFlexGrid3.ColWidth(5) = 0
MSHFlexGrid3.ColWidth(6) = 2000
'MSHFlexGrid3.ColWidth(7) = 0
MSHFlexGrid3.ColWidth(7) = 0
MSHFlexGrid3.ColWidth(8) = 0

MSHFlexGrid3.RowHeight(1) = 500
End Sub


Private Sub GetConnectedGodowns()

Call Grid_Head3

If CmbMonth.Text = "" Then
   Call Grid_Head3
   Exit Sub
End If

If CmbYear.Text = "" Then
   Call Grid_Head3
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   Call Grid_Head3
   Exit Sub
End If

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If CDate(Format(mFromDate, "yyyymm")) > Format(Date, "yyyymm") Then
   MsgBox "Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

tmp = " Select MC.CMPName,MG.GodownNo,MU.UnitType,MG.BagsCapacity,MG.GodownCapacity, " & _
      " '0' ReservationQty,'' as InvoiceAmt , MG.GodownID, MG.Address " & _
      " from Mst_Godown MG " & _
      " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID " & _
      " Inner Join Mst_UnitType MU On MG.UnitTypeID = MU.UnitTypeID " & _
      " Where MC.LocationID = " & mLocationID & " and (MG.CloseDate is null or MG.CloseDate >='" & mFromDate & "') And MG.StartDate<='" & mToDate & "'"
      
Call TmpTable

If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       For C = 0 To MSHFlexGrid3.Cols - 1
           MSHFlexGrid3.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(x, 3))
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(x, 4))
       MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + Val(MSHFlexGrid3.TextMatrix(x, 5))
       TmpRs.MoveNext
       x = x + 1
   Next
Else
   MsgBox "No godown mapped to this Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
End Sub

Private Sub SaveGodownDetails()

      
   tmp = "Delete from Txn_TempInvGodownDetails Where PINo = " & TxtPIID & " "
   Call TmpTable
      
   Call TableTxn_TempInvGodownDetails(" Where PINo = " & TxtPIID & "")
   x = 2
   For I = 1 To MSHFlexGrid3.Rows - 2
       If MSHFlexGrid3.TextMatrix(x, 0) = "" Then Exit For
'       If MSHFlexGrid3.TextMatrix(x, 7) <> "" Then
'          RsTxn_GodownMapDetails.Find "GodownMapDetailsID = " & MSHFlexGrid3.TextMatrix(x, 7)
'       Else
        RsTxn_TempInvGodownDetails.AddNew
        'RsTxn_TempInvGodownDetails!TIGodownMapDetailsID = GetMaxGodownDetailsID()
        RsTxn_TempInvGodownDetails!G_UID = GetGUID
'       End If
       RsTxn_TempInvGodownDetails!PINo = TxtPIID
       RsTxn_TempInvGodownDetails!GodownID = MSHFlexGrid3.TextMatrix(x, 7)
       RsTxn_TempInvGodownDetails!ReservationQty = Val(MSHFlexGrid3.TextMatrix(x, 5))
'       If Val(MSHFlexGrid3.TextMatrix(0, 5)) = 0 Then
'          RsTxn_TempInvGodownDetails!InvoiceAmt = Val(TxtAmount) '/ Val(MSHFlexGrid3.Rows - 3) '* Val(MSHFlexGrid3.TextMatrix(x, 5)) / 1 '"'Val(MSHFlexGrid3.TextMatrix(0, 5))
'       Else
'          RsTxn_TempInvGodownDetails!InvoiceAmt = Val(TxtAmount) * Val(MSHFlexGrid3.TextMatrix(x, 5)) / Val(MSHFlexGrid3.TextMatrix(0, 5))
'       End If
       RsTxn_TempInvGodownDetails!InvoiceAmt = Val(MSHFlexGrid3.TextMatrix(x, 6))
       If IsNull(RsTxn_TempInvGodownDetails!CreatedBy) = True Or RsTxn_TempInvGodownDetails!CreatedBy = "" Then
          RsTxn_TempInvGodownDetails!CreatedBy = Gen_UserLoginID
          RsTxn_TempInvGodownDetails!CreatedDate = Now
       Else
          RsTxn_TempInvGodownDetails!UpdatedBy = Gen_UserLoginID
          RsTxn_TempInvGodownDetails!UpdatedDate = Now
       End If
       x = x + 1
       RsTxn_TempInvGodownDetails.Update
Next
'End If
End Sub

Function GetMaxGodownDetailsID() As Double
Dim Retval As Double
Retval = 0
tmp = "Select max(TIGodownMapDetailsID) from Txn_TempInvGodownDetails "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetMaxGodownDetailsID = Retval

End Function

Private Sub InDataGodowns()

Call Grid_Head3


tmp = " Select MC.CMPName,MG.GodownNo,'0' 'UnitType','0' 'BagsCapacity' ,'0' 'GodownCapacity', " & _
      " '0' ReservationQty, TSRGD.InvoiceAmt, MG.GodownID, '' 'Address' " & _
      " From Txn_TempInvGodownDetails TSRGD " & _
      " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID " & _
      " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID " & _
      " Inner Join Txn_tempInvoice TSRG On TSRGD.PINO = TSRG.PINO " & _
      " Where TSRG.PINO = " & TxtPIID & " "
      
Call TmpTable

If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       For C = 0 To MSHFlexGrid3.Cols - 1
           MSHFlexGrid3.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(x, 6))
       TmpRs.MoveNext
       x = x + 1
   Next
End If
End Sub

Private Sub DistributeGodownIncome()
x = 2
If CmbLocationID.Text = "" Then
   Call Grid_Head3
   Exit Sub
End If

If CmbMonth.Text = "" Then
   Call Grid_Head3
   Exit Sub
End If

If CmbYear.Text = "" Then
   Call Grid_Head3
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   Call Grid_Head3
   Exit Sub
End If

MSHFlexGrid3.TextMatrix(0, 6) = ""
For I = 1 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(x, 0) = "" Then Exit For
    If Val(MSHFlexGrid3.TextMatrix(0, 5)) = 0 Then
       MSHFlexGrid3.TextMatrix(x, 6) = Val(TxtAmount) / Val(MSHFlexGrid3.Rows - 3) '* Val(MSHFlexGrid3.TextMatrix(x, 5)) / 1 '"'Val(MSHFlexGrid3.TextMatrix(0, 5))
    Else
       MSHFlexGrid3.TextMatrix(x, 6) = Val(TxtAmount) * Val(MSHFlexGrid3.TextMatrix(x, 5)) / Val(MSHFlexGrid3.TextMatrix(0, 5))
    End If
    MSHFlexGrid3.TextMatrix(x, 6) = Format(MSHFlexGrid3.TextMatrix(x, 6), "#######0.00")
    MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(x, 6))
    x = x + 1
Next

End Sub


Private Sub TxtDistIncomeAmount_LostFocus()

If TxtDistIncomeAmount = "" Then Exit Sub
If IsNumeric(TxtDistIncomeAmount) = False Or Val(TxtDistIncomeAmount) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtDistIncomeAmount.SetFocus
   Exit Sub
End If

If Val(TxtDistIncomeAmount) > Val(TxtAmount) Then
   MsgBox "Distributed amount should be less than or equal to Invoive Amount always!!!"
   TxtDistIncomeAmount.SetFocus
   Exit Sub
End If


End Sub

Private Sub CmdAddtoList_Click()

If CmbCMPName.Text = "" Then
   MsgBox "Blank CMP Not Allowed!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If

If CmbGodownNo.Text = "" Then
   MsgBox "Blank Godown Not Allowed!!!"
   CmbGodownNo.SetFocus
   Exit Sub
End If

If TxtDistIncomeAmount = "" Then
   MsgBox "Blank Amount Not Allowed!!!"
   
   TxtDistIncomeAmount.SetFocus
   Exit Sub
End If

Dim I, ans As Variant

ans = Validate_DuplicateGrid()
If ans = False Then
   MsgBox "Duplicate Entry !!!! "
   Exit Sub
End If

If CmdAddtoList.Caption = "Add To List" Then
   I = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
Else
   I = MSHFlexGrid3.RowSel
   MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) - Val(MSHFlexGrid3.TextMatrix(I, 6))
End If


MSHFlexGrid3.TextMatrix(I, 0) = CmbCMPName.Text  '"CMP"
MSHFlexGrid3.TextMatrix(I, 1) = CmbGodownNo.Text  '"Godown"
MSHFlexGrid3.TextMatrix(I, 2) = 0 'TxtUnitTypeID '"Unit Type"
MSHFlexGrid3.TextMatrix(I, 3) = 0 'Val(TxtBagsCapacity) '"Capacity in Bags"
MSHFlexGrid3.TextMatrix(I, 4) = 0 'Val(TxtGodownCapacity) '  "Capacity in MT"
MSHFlexGrid3.TextMatrix(I, 5) = 0 '"Reservation Qty"
MSHFlexGrid3.TextMatrix(I, 6) = Val(TxtDistIncomeAmount) '"GodownID"
'MSHFlexGrid3.TextMatrix(i, 7) = TxtDetailsID '"" '"SRGodownMapDetailID"
MSHFlexGrid3.TextMatrix(I, 7) = mGodownID
MSHFlexGrid3.TextMatrix(I, 8) = "" 'TxtGodownAddress"


MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(I, 6))

CmbCMPName.Text = ""
CmbGodownNo.Text = ""
TxtDistIncomeAmount = ""
CmdAddtoList.Caption = "Add To List"
CmbCMPName.SetFocus

End Sub


Private Function Validate_DuplicateGrid() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid3.Rows = 2 Then
   Retval = True
Else
    For I = 1 To MSHFlexGrid3.Rows - 1
        If MSHFlexGrid3.TextMatrix(I, 0) <> "" Then
           If CmdAddtoList.Caption = "Update List" Then
              If I <> MSHFlexGrid3.RowSel Then
                 If Val(MSHFlexGrid3.TextMatrix(I, 7)) = mGodownID Then
                    Retval = False
                 End If
              End If
           Else
              If Val(MSHFlexGrid3.TextMatrix(I, 7)) = mGodownID Then
                 Retval = False
              End If
           End If
        End If
    Next
End If
Validate_DuplicateGrid = Retval
End Function



Private Sub CmdRemoveFromList_Click()
If MSHFlexGrid3.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbCMPName.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   I = MSHFlexGrid3.RowSel
   If I > 0 Then
      If MSHFlexGrid3.TextMatrix(I, 0) <> "" Then
         
         MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) - Val(MSHFlexGrid3.TextMatrix(I, 6))

         For RO = I To MSHFlexGrid3.Rows - 2
             For C = 0 To MSHFlexGrid3.Cols - 1
                 MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
             Next
         Next
         MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
      End If
   End If
   CmbCMPName.Text = ""
   'Call ClearFrame
   CmbGodownNo.Text = ""
   TxtDistIncomeAmount = ""
   'CmdAddToGrid.Caption = "Add To Grid"
   CmdRemoveFromList.Enabled = False
   CmbCMPName.SetFocus
Else
   CmbCMPName.SetFocus
End If


End Sub
Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbMonth.Text = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox "Select Year!!!"
   CmbYear.SetFocus
   Exit Sub
End If

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

tmp = CmbCMPName.Text
Call TableMst_CMP(" Where (CloseDate is null or CloseDate >='" & mFromDate & "') And StartDate<='" & mToDate & "' And LocationID = " & mLocationID)
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   CmbGodownNo.Text = ""
   mCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbGodownNo.Text = ""
   CmbCMPName.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   CmbGodownNo.Text = ""
End If
mCMPID = RsMst_CMP!CMPID

End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If

If CmbMonth.Text = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox "Select Year!!!"
   CmbYear.SetFocus
   Exit Sub
End If

tmp = CmbGodownNo.Text
Call TableMst_Godown("Where (CloseDate is null or CloseDate >='" & mFromDate & "') And StartDate<='" & mToDate & "' And CMPID = " & mCMPID)
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
End Sub

Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
   mGodownID = Null
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
mGodownID = RsMst_Godown!GodownID

End Sub




