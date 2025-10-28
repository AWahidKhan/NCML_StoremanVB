VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_7DaysInvoice 
   Caption         =   "7 Days / Rejected Invoice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      Caption         =   "Log Details"
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
      Left            =   4680
      TabIndex        =   120
      Top             =   3360
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton Command2 
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
         Left            =   3375
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton CmdSaveLog 
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
         Left            =   1935
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   1560
         Width           =   1440
      End
      Begin VB.TextBox TxtRemark 
         Appearance      =   0  'Flat
         Height          =   1080
         Left            =   1200
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   121
         Top             =   360
         Width           =   5040
      End
      Begin VB.Label Label41 
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
         Left            =   240
         TabIndex        =   124
         Top             =   450
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Caption         =   "Log Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   4680
      TabIndex        =   125
      Top             =   2640
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox TxtLogRemark 
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
         Height          =   2640
         Left            =   1080
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   128
         Top             =   360
         Width           =   5160
      End
      Begin VB.CommandButton CmdDeleteLog 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete"
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
         Left            =   3375
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   3120
         Width           =   1440
      End
      Begin VB.CommandButton CmdCancelLog 
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
         Left            =   4815
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label42 
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
         TabIndex        =   129
         Top             =   1530
         Width           =   1260
      End
   End
   Begin VB.Frame Frame4 
      Height          =   9645
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   14955
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
         Left            =   4560
         TabIndex        =   8
         Top             =   3120
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox TxtReverseRemark 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   1560
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   1440
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker TxtReverseDate 
            Height          =   375
            Left            =   1560
            TabIndex        =   12
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
            Format          =   116981761
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   390
            Width           =   1800
         End
      End
      Begin VB.Frame Frame0 
         Height          =   1845
         Left            =   840
         TabIndex        =   15
         Top             =   6000
         Visible         =   0   'False
         Width           =   14685
         Begin VB.ComboBox CmbSReverse 
            Height          =   315
            ItemData        =   "FrmTxn_7DaysInvoice.frx":0000
            Left            =   6975
            List            =   "FrmTxn_7DaysInvoice.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   1170
            Width           =   1125
         End
         Begin VB.ComboBox Cmbsoutstanding 
            Height          =   315
            ItemData        =   "FrmTxn_7DaysInvoice.frx":0017
            Left            =   3953
            List            =   "FrmTxn_7DaysInvoice.frx":0021
            TabIndex        =   25
            Top             =   1170
            Width           =   1365
         End
         Begin VB.ComboBox CmbSInvoiceType 
            Height          =   315
            ItemData        =   "FrmTxn_7DaysInvoice.frx":0035
            Left            =   2445
            List            =   "FrmTxn_7DaysInvoice.frx":0048
            TabIndex        =   24
            Top             =   1170
            Width           =   1365
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   375
            TabIndex        =   16
            Top             =   443
            Width           =   225
         End
         Begin VB.TextBox TxtSCDTFNo 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1080
            TabIndex        =   23
            Top             =   1177
            Width           =   1305
         End
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_7DaysInvoice.frx":0080
            Left            =   11685
            List            =   "FrmTxn_7DaysInvoice.frx":008A
            TabIndex        =   21
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox TxtSBOEID 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5535
            TabIndex        =   26
            Top             =   1177
            Width           =   1305
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   12780
            TabIndex        =   29
            Top             =   1127
            Width           =   885
         End
         Begin VB.CommandButton CmdSearch 
            Caption         =   "Search"
            Height          =   350
            Left            =   11895
            TabIndex        =   28
            Top             =   1127
            Width           =   885
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "Clear"
            Height          =   350
            Left            =   13665
            TabIndex        =   30
            Top             =   1127
            Width           =   885
         End
         Begin VB.ComboBox CmbPreFix 
            Height          =   315
            Left            =   3870
            Sorted          =   -1  'True
            TabIndex        =   19
            Top             =   420
            Width           =   2400
         End
         Begin VB.ComboBox CmbNonNcdexClient 
            Height          =   315
            Left            =   6270
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   420
            Width           =   5385
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   1200
            Width           =   225
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1215
            Left            =   120
            TabIndex        =   31
            Top             =   1680
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   2143
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   300
            Left            =   1080
            TabIndex        =   17
            Top             =   420
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
            Format          =   116981761
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   300
            Left            =   2460
            TabIndex        =   18
            Top             =   420
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
            Format          =   116981761
            CurrentDate     =   39884
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Reverse"
            Height          =   195
            Left            =   7230
            TabIndex        =   119
            Top             =   855
            Width           =   600
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Outstanding Amount"
            Height          =   195
            Left            =   3915
            TabIndex        =   118
            Top             =   855
            Width           =   1440
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Invoice Type"
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
            Left            =   2535
            TabIndex        =   117
            Top             =   855
            Width           =   1125
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Invoice Date"
            Height          =   195
            Left            =   0
            TabIndex        =   114
            Top             =   195
            Width           =   915
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "CxTF No"
            Height          =   195
            Left            =   1395
            TabIndex        =   113
            Top             =   855
            Width           =   630
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            Height          =   195
            Left            =   12375
            TabIndex        =   40
            Top             =   195
            Width           =   450
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "To Date"
            Height          =   195
            Left            =   2835
            TabIndex        =   39
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Invoice No"
            Height          =   195
            Left            =   5790
            TabIndex        =   38
            Top             =   855
            Width           =   780
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Location"
            Height          =   195
            Left            =   4770
            TabIndex        =   37
            Top             =   195
            Width           =   615
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Client Name"
            Height          =   195
            Left            =   8535
            TabIndex        =   36
            Top             =   195
            Width           =   855
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "From Date"
            Height          =   195
            Left            =   1380
            TabIndex        =   35
            Top             =   195
            Width           =   735
         End
         Begin VB.Label LblTotalRecord 
            AutoSize        =   -1  'True
            Caption         =   "-"
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
            Left            =   12150
            TabIndex        =   34
            Top             =   870
            Width           =   75
         End
         Begin VB.Label Label24 
            Caption         =   "Total Record Found :-"
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
            Left            =   10110
            TabIndex        =   33
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Batch No"
            Height          =   195
            Left            =   60
            TabIndex        =   32
            Top             =   855
            Width           =   675
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1155
         Left            =   0
         TabIndex        =   92
         Top             =   8400
         Width           =   14880
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Print"
            Height          =   400
            Left            =   10845
            TabIndex        =   105
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   8745
            TabIndex        =   104
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   8745
            TabIndex        =   103
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   6645
            TabIndex        =   102
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   6645
            TabIndex        =   101
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   4545
            TabIndex        =   100
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   4545
            TabIndex        =   99
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2445
            TabIndex        =   98
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2445
            TabIndex        =   97
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   345
            TabIndex        =   96
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   345
            TabIndex        =   95
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton CmdReverse 
            Caption         =   "Reverse"
            Height          =   400
            Left            =   10845
            TabIndex        =   94
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton CmdViewInvoice 
            Caption         =   "View Invoice"
            Enabled         =   0   'False
            Height          =   810
            Left            =   12960
            TabIndex        =   93
            Top             =   195
            Width           =   1740
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8445
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   14880
         Begin VB.TextBox TxtInvoiceType 
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
            Left            =   12600
            Locked          =   -1  'True
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   2595
            Width           =   2040
         End
         Begin VB.TextBox TxtFAGPostingDate 
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
            Left            =   5190
            Locked          =   -1  'True
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   2610
            Width           =   2025
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
            Height          =   360
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   1114
            Width           =   5505
         End
         Begin VB.TextBox TxtCDTFDate 
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
            Left            =   12615
            Locked          =   -1  'True
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   1100
            Width           =   2025
         End
         Begin VB.TextBox TxtCDTFNo 
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
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   1100
            Width           =   2025
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
            ItemData        =   "FrmTxn_7DaysInvoice.frx":00A2
            Left            =   1680
            List            =   "FrmTxn_7DaysInvoice.frx":00AC
            TabIndex        =   4
            Top             =   2606
            Width           =   2055
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
            Height          =   540
            Left            =   1680
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1988
            Width           =   13020
         End
         Begin VB.TextBox TxtOriginalBOEID 
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
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   225
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox TxtG_UID 
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
            Left            =   14535
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.TextBox TxtFAGRemarks 
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
            Height          =   480
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   3045
            Width           =   13050
         End
         Begin VB.TextBox TxtBatchID 
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
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   2606
            Width           =   2040
         End
         Begin VB.TextBox TxtServiceTaxNAVCode 
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
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1551
            Width           =   2025
         End
         Begin VB.TextBox TxtExchangeType 
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
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   240
            Width           =   2025
         End
         Begin VB.TextBox TxtBillingMonth 
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
            Left            =   12585
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   2025
         End
         Begin VB.TextBox TxtPremium 
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
            Left            =   12600
            TabIndex        =   2
            Top             =   1551
            Width           =   2025
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
            Height          =   360
            Left            =   8880
            TabIndex        =   1
            Top             =   1551
            Width           =   2025
         End
         Begin VB.TextBox TxtAmount 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1551
            Width           =   2025
         End
         Begin VB.TextBox TxtPrefix 
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
            Left            =   13560
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   2115
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.TextBox TxtInvoiceDate 
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
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   240
            Width           =   2025
         End
         Begin VB.TextBox TxtBOEID 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   240
            Width           =   2025
         End
         Begin VB.TextBox TxtLocationID 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   677
            Width           =   5505
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
            Height          =   360
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   675
            Width           =   5745
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   4815
            Left            =   105
            TabIndex        =   55
            Top             =   3540
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   8493
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "GSL Details"
            TabPicture(0)   =   "FrmTxn_7DaysInvoice.frx":00C4
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "MSHFlexGrid2"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "TI Details/Log History"
            TabPicture(1)   =   "FrmTxn_7DaysInvoice.frx":00E0
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Label28"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label27"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "Label22"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "Label21"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "Label20"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "Label25"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "Label26"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "Label47"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "Label46"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "Label43"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "TxtBalAmt"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "TxtTIAmount"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "TxtTIDate"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "TxtTINo"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).Control(14)=   "TxtTISTAmt"
            Tab(1).Control(14).Enabled=   0   'False
            Tab(1).Control(15)=   "TxtRecAmt"
            Tab(1).Control(15).Enabled=   0   'False
            Tab(1).Control(16)=   "TxtEntryNo"
            Tab(1).Control(16).Enabled=   0   'False
            Tab(1).Control(17)=   "TxtUpdated"
            Tab(1).Control(17).Enabled=   0   'False
            Tab(1).Control(18)=   "TxtCreated"
            Tab(1).Control(18).Enabled=   0   'False
            Tab(1).Control(19)=   "TxtPostInvoiceNo"
            Tab(1).Control(19).Enabled=   0   'False
            Tab(1).ControlCount=   20
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
               Left            =   1485
               Locked          =   -1  'True
               TabIndex        =   130
               Top             =   1230
               Width           =   3690
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
               Left            =   1485
               Locked          =   -1  'True
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   2040
               Width           =   10965
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
               Left            =   1485
               Locked          =   -1  'True
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   2565
               Width           =   10950
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
               Left            =   12450
               Locked          =   -1  'True
               TabIndex        =   62
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
               Left            =   8820
               Locked          =   -1  'True
               TabIndex        =   61
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
               Left            =   7005
               Locked          =   -1  'True
               TabIndex        =   60
               Top             =   750
               Width           =   1815
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
               Left            =   1485
               Locked          =   -1  'True
               TabIndex        =   59
               Top             =   750
               Width           =   1890
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
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   58
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
               Left            =   5190
               Locked          =   -1  'True
               TabIndex        =   57
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
               Left            =   10635
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   750
               Width           =   1815
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   4335
               Left            =   -74880
               TabIndex        =   65
               Top             =   360
               Width           =   14580
               _ExtentX        =   25718
               _ExtentY        =   7646
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
            Begin VB.Label Label43 
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
               Height          =   495
               Left            =   120
               TabIndex        =   131
               Top             =   1170
               Width           =   1335
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
               Left            =   240
               TabIndex        =   74
               Top             =   2100
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
               Left            =   240
               TabIndex        =   73
               Top             =   2625
               Width           =   975
            End
            Begin VB.Label Label26 
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
               Left            =   12945
               TabIndex        =   72
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label25 
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
               Left            =   8925
               TabIndex        =   71
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label20 
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
               Left            =   7305
               TabIndex        =   70
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label21 
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
               Left            =   2235
               TabIndex        =   69
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Label22 
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
               Left            =   5625
               TabIndex        =   68
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label27 
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
               Left            =   3945
               TabIndex        =   67
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label28 
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
               Left            =   10785
               TabIndex        =   66
               Top             =   480
               Width           =   1695
            End
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
            Height          =   360
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   2606
            Width           =   2025
         End
         Begin VB.Label Label38 
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
            Left            =   11160
            TabIndex        =   116
            Top             =   2625
            Width           =   1440
         End
         Begin VB.Label Label32 
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
            Height          =   240
            Left            =   120
            TabIndex        =   111
            Top             =   1140
            Width           =   1440
         End
         Begin VB.Label Label31 
            Caption         =   "CxTF Date"
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
            Left            =   11160
            TabIndex        =   109
            Top             =   1140
            Width           =   1440
         End
         Begin VB.Label Label29 
            Caption         =   "CxTF No"
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
            Left            =   7320
            TabIndex        =   107
            Top             =   1140
            Width           =   1440
         End
         Begin VB.Label Label18 
            Caption         =   "Status"
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
            TabIndex        =   91
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Remarks"
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
            TabIndex        =   90
            Top             =   2100
            Width           =   1440
         End
         Begin VB.Label Label13 
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
            Height          =   240
            Left            =   13920
            TabIndex        =   89
            Top             =   2220
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Label12 
            Caption         =   "FAG Remarks"
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
            TabIndex        =   88
            Top             =   3120
            Width           =   1440
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
            Height          =   480
            Left            =   3915
            TabIndex        =   87
            Top             =   2546
            Width           =   1560
         End
         Begin VB.Label Label10 
            Caption         =   "Batch ID"
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
            Left            =   7440
            TabIndex        =   86
            Top             =   2640
            Width           =   1440
         End
         Begin VB.Label Label9 
            Caption         =   "Service Tax NAV Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3960
            TabIndex        =   85
            Top             =   1506
            Width           =   1440
         End
         Begin VB.Label Label8 
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
            Height          =   240
            Left            =   7305
            TabIndex        =   84
            Top             =   300
            Width           =   1440
         End
         Begin VB.Label Label7 
            Caption         =   "Billing Month"
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
            Left            =   11160
            TabIndex        =   83
            Top             =   300
            Width           =   1440
         End
         Begin VB.Label Label6 
            Caption         =   "Premium"
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
            Left            =   11160
            TabIndex        =   82
            Top             =   1611
            Width           =   855
         End
         Begin VB.Label Label5 
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
            Height          =   240
            Left            =   7320
            TabIndex        =   81
            Top             =   1611
            Width           =   1440
         End
         Begin VB.Label Label4 
            Caption         =   "Amount"
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
            TabIndex        =   80
            Top             =   1515
            Width           =   1440
         End
         Begin VB.Label Label3 
            Caption         =   "SM_Prefix"
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
            Left            =   12615
            TabIndex        =   79
            Top             =   2175
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label2 
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
            Height          =   240
            Left            =   3960
            TabIndex        =   78
            Top             =   270
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Invoice No"
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
            TabIndex        =   77
            Top             =   270
            Width           =   1440
         End
         Begin VB.Label LblLocationID 
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
            Left            =   120
            TabIndex        =   76
            Top             =   735
            Width           =   1440
         End
         Begin VB.Label LblClientID 
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
            Height          =   240
            Left            =   7305
            TabIndex        =   75
            Top             =   735
            Width           =   1440
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   10095
      Left            =   15060
      TabIndex        =   6
      Top             =   -120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   135
      TabIndex        =   0
      Top             =   9765
      Width           =   10935
   End
End
Attribute VB_Name = "FrmTxn_7DaysInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID As Double
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mMonth, mTotalDays, mYear As Variant
Dim mBOEID As Variant
Dim mFlag As Variant
Dim mEntryNo As Variant
Dim mFAGToDate As Variant
Dim mSInvoiceType As Variant
Dim mHSNNoNAGRI, mHSNNoAGRI As Variant
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 17 '16
MSHFlexGrid2.TextMatrix(1, 0) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 1) = "GSLID"
MSHFlexGrid2.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 3) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 4) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 5) = "No Of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Quantity"
MSHFlexGrid2.TextMatrix(1, 7) = "Billing Cycle"
MSHFlexGrid2.TextMatrix(1, 8) = "Billing Unit"
MSHFlexGrid2.TextMatrix(1, 9) = "First CDTF Date"
MSHFlexGrid2.TextMatrix(1, 10) = "No of Cycle"
MSHFlexGrid2.TextMatrix(1, 11) = "Rate per Cycle"
MSHFlexGrid2.TextMatrix(1, 12) = "Amount"

MSHFlexGrid2.TextMatrix(1, 13) = "RRN Date"
MSHFlexGrid2.TextMatrix(1, 14) = "FED Date"
MSHFlexGrid2.TextMatrix(1, 15) = "RRN no"
MSHFlexGrid2.TextMatrix(1, 16) = "NCDEX Space Charges"

MSHFlexGrid2.RowHeight(1) = 700
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.ColWidth(0) = 1800
MSHFlexGrid2.ColWidth(1) = 0
MSHFlexGrid2.ColWidth(2) = 1000
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 1000
MSHFlexGrid2.ColWidth(5) = 1600
MSHFlexGrid2.ColWidth(6) = 1000
MSHFlexGrid2.ColWidth(7) = 1000
MSHFlexGrid2.ColWidth(8) = 1000
MSHFlexGrid2.ColWidth(9) = 1400
MSHFlexGrid2.ColWidth(10) = 1000
MSHFlexGrid2.ColWidth(11) = 1000
MSHFlexGrid2.ColWidth(12) = 1000
MSHFlexGrid2.ColWidth(13) = 1400
MSHFlexGrid2.ColWidth(14) = 1400
MSHFlexGrid2.ColWidth(15) = 1400
MSHFlexGrid2.ColWidth(16) = 1400
End Sub


Private Sub AddCmd_Click()

Call GButtonLock(Me, False)
CmdPrint.Enabled = False
CmdReverse.Enabled = False
End Sub

Private Sub Check1_Click()
Call TxtFromDate_LostFocus
Call TxtToDate_LostFocus
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
   mLocationID = 0
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

End Sub


Private Sub Cmbsoutstanding_Lostfocus()
If Cmbsoutstanding.Text = "" Then Exit Sub

If LCase(Cmbsoutstanding.Text) <> "zero" And LCase(Cmbsoutstanding.Text) <> "non zero" Then
   MsgBox "Invalid Selection!!!"
   Cmbsoutstanding.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSReverse_Lostfocus()
If CmbSReverse.Text = "" Then Exit Sub
If LCase(CmbSReverse.Text) <> "yes" And LCase(CmbSReverse.Text) <> "no" Then
   MsgBox "Invalid Selection"
   CmbSReverse.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSStatus_LostFocus()

If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "confirmed" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "confirmed" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdCancelLog_Click()
Frame3.Visible = False
TxtLogRemark.Text = ""
End Sub

Private Sub CmdClear_Click()
Dim mSMonth, mSYear, mSTotalDays As Variant
Check1.Value = 0
TxtFromDate.Value = Date
TxtToDate.Value = Date
Call TxtFromDate_LostFocus
Call TxtToDate_LostFocus
CmbPreFix.Text = ""
CmbNonNcdexClient.Text = ""
TxtSBOEID.Text = ""
TxtSCDTFNo = ""
LblTotalRecord.Caption = "-"
Call Grid_Head
End Sub

Private Sub CmdDeleteLog_Click()

If Trim(TxtLogRemark.Text) = "" Then
   MsgBox "Log Remark not allowed!!!"
   TxtLogRemark.SetFocus
   Exit Sub
End If

Call TableLog_InvoiceDeletetion(" Where DocumentType = ''")

RsLog_InvoiceDeletetion.AddNew
If RsTxn_CDTFBOE!TxnType = "W" Then
   RsLog_InvoiceDeletetion!DocumentType = "RJ"
ElseIf RsTxn_CDTFBOE!TxnType = "D" Then
   RsLog_InvoiceDeletetion!DocumentType = "SD"
ElseIf RsTxn_CDTFBOE!TxnType = "R" Then
   RsLog_InvoiceDeletetion!DocumentType = "RRN"
End If
RsLog_InvoiceDeletetion!DcoumentID = RsTxn_CDTFBOE!BOEID
RsLog_InvoiceDeletetion!SM_Prefix = RsTxn_CDTFBOE!SM_Prefix
RsLog_InvoiceDeletetion!ClientID = RsTxn_CDTFBOE!ClientID
RsLog_InvoiceDeletetion!Amount = RsTxn_CDTFBOE!Amount
RsLog_InvoiceDeletetion!DeletetionRemark = Trim(TxtLogRemark.Text)
RsLog_InvoiceDeletetion!DeletedBy = Gen_UserLoginID
RsLog_InvoiceDeletetion!DeletedOn = Now
RsLog_InvoiceDeletetion.Update

Call DeleteChildTableDetail
RsTxn_CDTFBOE.Delete
RsTxn_CDTFBOE.Update
'RsTxn_CDTFBOE.Requery
If RsTxn_CDTFBOE.RecordCount = 0 Then
   Call ClearData
   Call GButtonLockAD(Me)
   'Call AddCmd_Click
   Exit Sub
End If
RsTxn_CDTFBOE.MoveNext
If RsTxn_CDTFBOE.EOF() Then
   RsTxn_CDTFBOE.MoveLast
End If
Call Indata
Frame3.Visible = False
TxtLogRemark.Text = ""
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

'Commented on 21 Aug 2013
'If TxtReverseDate.Value < CDate(TxtFAGPostingDate.Text) Then

If TxtReverseDate.Value < CDate(TxtTIDate.Text) Then
   MsgBox "'Reverse Date' should not less than 'TI Date' !!"
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

RsTxn_CDTFBOE!ReverseDate = TxtReverseDate.Value
RsTxn_CDTFBOE!ReverseFlag = "Y"
RsTxn_CDTFBOE!ReverseRemark = TxtReverseRemark.Text
RsTxn_CDTFBOE!UpdatedBy = Gen_UserLoginID
RsTxn_CDTFBOE!UpdatedDate = Now
RsTxn_CDTFBOE.Update

If RsTxn_CDTFBOE!TxnType = "D" Then
   tmp = " Update Txn_TIDetails Set TIAmount=0,ReverseFlag='Y',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where PIID=" & RsTxn_CDTFBOE!BOEID & " and PIFLag='SD'"
Else
   If LCase(TxtInvoiceType) = "rrn withdrawal" Then
      tmp = " Update Txn_TIDetails Set TIAmount=0,ReverseFlag='Y',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where PIID=" & RsTxn_CDTFBOE!BOEID & " and PIFLag='RRN'"
   ElseIf LCase(TxtInvoiceType) = "nnr yearly" Then
      tmp = " Update Txn_TIDetails Set TIAmount=0,ReverseFlag='Y',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where PIID=" & RsTxn_CDTFBOE!BOEID & " and PIFLag='NNNRY'"
   ElseIf LCase(TxtInvoiceType) = "pwr" Then
      tmp = " Update Txn_TIDetails Set TIAmount=0,ReverseFlag='Y',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where PIID=" & RsTxn_CDTFBOE!BOEID & " and PIFLag='PWR'"
   Else
      tmp = " Update Txn_TIDetails Set TIAmount=0,ReverseFlag='Y',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where PIID=" & RsTxn_CDTFBOE!BOEID & " and PIFLag='RJ'"
   End If
End If

Call TmpTable

Frame6.Visible = False
CmdReverse.Enabled = False
End Sub

Private Sub CmdSearch_Click()
Dim wc As Variant

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date Should Be Less Than To Date !!"
   TxtFromDate.SetFocus
   Exit Sub
End If

Call Grid_Head
If CmbSInvoiceType.Text = "" Then
   MsgBox "Please select Invoice Type!!"
   CmbSInvoiceType.SetFocus
   Exit Sub
End If
'Call TableTxn_TIDetails(" Where PIFlag in ('SD','RJ')")

tmp = "Select TI.TIID,TI.TINo,TI.TIDate,TI.TIAmount,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode,TI.TISTAmount ,TI.PIID,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+isnull(TI.AdjAmountFAG,0) RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID ) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDS,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID )+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID  )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Where TI.PIFLag in ('SD','RJ','RRN','NNNRY','PWR')  "

Call Tmp1Table

If LCase(CmbSInvoiceType.Text) <> "pwr" Then
   tmp = " Select TPI.BOEID,TPI.InvoiceDate,TPI.BillingMonth,ML.LocationName,ME.ExchangeType,TPI.CxTF_DetailsID,MC.ClientName+'~'+MC.ClientIDRow 'ClientName',"
Else
   tmp = " Select TPI.BOEID,TPI.InvoiceDate,TPI.BillingMonth,ML.LocationName,'NCDEX' 'ExchangeType',TPI.CxTF_DetailsID,MC.ClientName+'~'+MC.ClientIDRow 'ClientName', "
End If
tmp = tmp & " TPI.Amount , TPI.Discount, TPI.Premium, TPI.BatchID, MST.NAVCode, TPI.FAGPostingDate, TPI.G_UID,TPI.Remarks,TPI.Flag, "
tmp = tmp & "  '' TIID,'' TINo,'' TIDate,'' TIAmount, '' TISTAmount,''RecAmt,''RecTDS,''RecOtherDed,'' PostInvoiceNo,'' EntryNo,TPI.ReverseFlag,TPI.ReverseDate,TPI.ReverseRemark,TPI.RBatchID, MComm.COmmodity, sum(TIMGD.CDTFQty) Quantity, count(distinct TIMGD.Rate) RateCount, max(TIMGD.Rate) Rate,TPI.TxnType,Sum(isnull(TIMGD.NCDEXSpaceCharges,0)) 'NCDEXSpaceCharges','' 'HSNNo',MComm.AccountType,MSI.StateIndiaName ,MSI.GSTIN 'NCMLGSTNo',MCSG.GSTNo 'ClientGSTNo',MCSG.GSTNAVCode,'' 'TaxApplied',TPI.CreatedDate 'InvoiceCreatedOn',TC.CreatedDate 'TransactionCreatedOn' "
tmp = tmp & "  From Txn_CDTFBOE TPI "
tmp = tmp & "  Inner Join Txn_CDTFBOE_GSL TIMGD On TIMGD.BOEID = TPI.BOEID And TIMGD.SM_Prefix = TPI.SM_Prefix "
If LCase(CmbSInvoiceType.Text) = "7 days" Then
   tmp = tmp & "  left Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType='D' and TPI.TxnType='D'"
ElseIf LCase(CmbSInvoiceType.Text) = "rrn withdrawal" Then
   tmp = tmp & "  left Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType='W' and TPI.TxnType='R'"
ElseIf LCase(CmbSInvoiceType.Text) = "nnnr yearly" Then
   tmp = tmp & "  left Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType='W' and TPI.TxnType='Y'"
ElseIf LCase(CmbSInvoiceType.Text) = "pwr" Then
   tmp = tmp & "  left Join Txn_RRNPWR TC On TPI.CxTF_DetailsID = TC.ConvertedID And  TPI.TxnType='P'"
Else
   tmp = tmp & "  left Join Txn_CxTF_Details TC On TPI.CxTF_DetailsID = TC.CxTF_DetailsID And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType='W' and TPI.TxnType='W'"
End If
tmp = tmp & " left join Mst_Commodity MComm On TC.CommodityID = MComm.CommodityID "
tmp = tmp & "  left join Mst_ServiceTax MST On MComm.ServiceTaxID = MST.ServiceTaxID "
tmp = tmp & "  Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID "
If LCase(CmbSInvoiceType.Text) <> "pwr" Then
   tmp = tmp & "  left Join Mst_ExchangeType ME On TC.ExchangeTypeID = ME.ExchangeTypeID "
End If
tmp = tmp & "  Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID "
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID "
tmp = tmp & "  Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID "
tmp = tmp & " Left  Join Mst_ClientStateWiseGST MCSG On MC.ClientID = MCSG.ClientID ANd MSI.StateIndiaID = MCSG.StateINdiaID "

If Gen_DBType = "MDB" Then
   If Check1.Value = 1 Then
      wc = " Where TPI.InvoiceDate >= #" & TxtFromDate.Value & "# And TPI.InvoiceDate <= #" & TxtToDate.Value & "# "
   Else
      wc = " Where TPI.BillingMonth >= #" & TxtFromDate.Value & "# And TPI.BillingMonth <= #" & TxtToDate.Value & "# "
   End If
Else
   If Check1.Value = 1 Then
      wc = " Where TPI.InvoiceDate >= '" & TxtFromDate.Value & "' And TPI.InvoiceDate <= '" & TxtToDate.Value & "' "
   Else
      wc = " Where TPI.BillingMonth >= '" & TxtFromDate.Value & "' And TPI.BillingMonth <= '" & TxtToDate.Value & "' "
   End If
End If
            
If CmbPreFix.Text <> "" Then
   wc = wc & " And  ML.LocationName = '" & CmbPreFix.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And ML." & Gen_WcLocation
   End If
End If
      
If CmbNonNcdexClient.Text <> "" Then
   wc = wc & " And  MC.ClientName = '" & CmbNonNcdexClient.Text & "'"
End If

If CmbSStatus.Text <> "" Then
   wc = wc & " And  TPI.Flag = '" & Left(CmbSStatus.Text, 1) & "'"
End If

If CmbSInvoiceType.Text <> "" Then
   If LCase(CmbSInvoiceType.Text) = "rejected" Then
      wc = wc & " And  TPI.TxnType = 'W'"
   ElseIf LCase(CmbSInvoiceType.Text) = "7 days" Then
      wc = wc & " And  TPI.TxnType = 'D'"
   ElseIf LCase(CmbSInvoiceType.Text) = "rrn withdrawal" Then
      wc = wc & " And  TPI.TxnType = 'R'"
   ElseIf LCase(CmbSInvoiceType.Text) = "nnnr yearly" Then
      wc = wc & " And  TPI.TxnType = 'Y'"
   ElseIf LCase(CmbSInvoiceType.Text) = "pwr" Then
      wc = wc & " And  TPI.TxnType = 'P'"
   End If
End If


If TxtSBOEID.Text <> "" Then
   wc = wc & " And  TPI.BOEID = '" & TxtSBOEID.Text & "'"
End If

If ChkBatchNo.Value = 0 Then
   wc = wc & " And  TPI.BatchID = 0 " '" & TxtSBOEID.Text & ""
End If
            
If TxtSCDTFNo.Text <> "" Then
   wc = wc & " And  TPI.CxTF_DetailsID = " & Val(TxtSCDTFNo.Text)
End If
If LCase(Cmbsoutstanding.Text) = "zero" Then
   wc = wc & " And TPI.Amount - TPI.Discount + TPI.Premium = 0"
ElseIf LCase(Cmbsoutstanding.Text) = "non zero" Then
   wc = wc & " And TPI.Amount - TPI.Discount + TPI.Premium > 0"
End If
If LCase(CmbSReverse.Text) = "yes" Then
   wc = wc & " And TPI.ReverseFlag = 'Y'"
ElseIf LCase(Cmbsoutstanding.Text) = "no" Then
   wc = wc & " And isnull(TPI.ReverseFlag, 'N') = 'N'"
End If
            
tmp1 = " Order By TPI.BOEID "

tmp = tmp & " " & wc
If LCase(CmbSInvoiceType.Text) <> "pwr" Then
   tmp = tmp & " group by TPI.BOEID, TPI.InvoiceDate,TPI.BillingMonth,ML.LocationName,ME.ExchangeType,TPI.CxTF_DetailsID,MC.ClientName,MC.ClientIDRow,"
Else
   tmp = tmp & " group by TPI.BOEID, TPI.InvoiceDate,TPI.BillingMonth,ML.LocationName,TPI.CxTF_DetailsID,MC.ClientName,MC.ClientID,"
End If
tmp = tmp & " TPI.Amount , TPI.Discount, TPI.Premium, TPI.BatchID, MST.NAVCode, TPI.FAGPostingDate, TPI.G_UID,TPI.Remarks,TPI.Flag,"
tmp = tmp & " TPI.ReverseFlag, TPI.ReverseDate, TPI.ReverseRemark, TPI.RBatchID,TPI.TxnType"
tmp = tmp & " , MComm.COmmodity,MComm.AccountType,MSI.StateIndiaName ,MSI.GSTIN ,MCSG.GSTNo ,MCSG.GSTNAVCode,TPI.CreatedDate,TC.CreatedDate "

tmp = tmp & " " & tmp1

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
       TmpRs1.Find "PIID=" & TmpRs!BOEID
       If TmpRs1.EOF = False Then
       '' TIID,'' TINo,'' TIDate,'' TIAmount, '' TISTAmount,'' EntryNo
          MSHFlexGrid1.TextMatrix(I, 16) = TmpRs1!TIID
          MSHFlexGrid1.TextMatrix(I, 17) = TmpRs1!TINo
          MSHFlexGrid1.TextMatrix(I, 18) = TmpRs1!TIDate
          MSHFlexGrid1.TextMatrix(I, 19) = TmpRs1!TIAmount
          MSHFlexGrid1.TextMatrix(I, 20) = TmpRs1!TISTAmount
          
          MSHFlexGrid1.TextMatrix(I, 21) = TmpRs1!RecAmount
          MSHFlexGrid1.TextMatrix(I, 22) = TmpRs1!RecTDS
          MSHFlexGrid1.TextMatrix(I, 23) = TmpRs1!RecOtherDed
          
          MSHFlexGrid1.TextMatrix(I, 24) = IIf(IsNull(TmpRs1!PostInvoiceNo), "", TmpRs1!PostInvoiceNo)
          MSHFlexGrid1.TextMatrix(I, 25) = TmpRs1!EntryNo
       End If
       MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 2) = Format(MSHFlexGrid1.TextMatrix(I, 2), "dd-mmm-yyyy")
       
       If MSHFlexGrid1.TextMatrix(I, 18) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 18) = Format(MSHFlexGrid1.TextMatrix(I, 18), "dd-mmm-yyyy")
       End If
       
       If LCase(MSHFlexGrid1.TextMatrix(I, 37)) = "nagri" Then
          MSHFlexGrid1.TextMatrix(I, 36) = mHSNNoNAGRI
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 37)) = "agri" Then
          MSHFlexGrid1.TextMatrix(I, 36) = mHSNNoAGRI
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 15) = "P" Then
          MSHFlexGrid1.TextMatrix(I, 15) = "Pending"
       ElseIf MSHFlexGrid1.TextMatrix(I, 15) = "C" Then
          MSHFlexGrid1.TextMatrix(I, 15) = "Confirmed"
       ElseIf MSHFlexGrid1.TextMatrix(I, 15) = "A" Then
          MSHFlexGrid1.TextMatrix(I, 15) = "Approved"
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 26) = "Y" Then
          MSHFlexGrid1.TextMatrix(I, 26) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 26) = "No"
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 27) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 27) = Format(MSHFlexGrid1.TextMatrix(I, 27), "dd-mmm-yyyy")
       End If
       
       If Val(MSHFlexGrid1.TextMatrix(I, 32)) <> 1 Then
          MSHFlexGrid1.TextMatrix(I, 33) = ""
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 34) = "W" Then
          MSHFlexGrid1.TextMatrix(I, 34) = "Rejected"
          MSHFlexGrid1.TextMatrix(I, 0) = "RJ" & MSHFlexGrid1.TextMatrix(I, 0)
       ElseIf MSHFlexGrid1.TextMatrix(I, 34) = "D" Then
          MSHFlexGrid1.TextMatrix(I, 34) = "7 Days"
          MSHFlexGrid1.TextMatrix(I, 0) = "SD" & MSHFlexGrid1.TextMatrix(I, 0)
       ElseIf MSHFlexGrid1.TextMatrix(I, 34) = "R" Then
          MSHFlexGrid1.TextMatrix(I, 34) = "RRN Withdrawal"
          MSHFlexGrid1.TextMatrix(I, 0) = "RRN" & MSHFlexGrid1.TextMatrix(I, 0)
       ElseIf MSHFlexGrid1.TextMatrix(I, 34) = "Y" Then
          MSHFlexGrid1.TextMatrix(I, 34) = "NNNR Yearly"
          MSHFlexGrid1.TextMatrix(I, 0) = "NNNRY" & MSHFlexGrid1.TextMatrix(I, 0)
       ElseIf MSHFlexGrid1.TextMatrix(I, 34) = "P" Then
          MSHFlexGrid1.TextMatrix(I, 34) = "PWR"
          MSHFlexGrid1.TextMatrix(I, 0) = "PWR" & MSHFlexGrid1.TextMatrix(I, 0)
       End If
       

       If Trim(MSHFlexGrid1.TextMatrix(I, 39)) <> "" And Trim(MSHFlexGrid1.TextMatrix(I, 40)) <> "" Then
          If Left(MSHFlexGrid1.TextMatrix(I, 39), 2) = Left(MSHFlexGrid1.TextMatrix(I, 40), 2) Then
             MSHFlexGrid1.TextMatrix(I, 42) = "Intra"
          Else
             MSHFlexGrid1.TextMatrix(I, 42) = "Inter"
          End If
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 43) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 43) = Format(MSHFlexGrid1.TextMatrix(I, 43), "dd-mmm-yyyy hh:mm:ss")
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 44) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 44) = Format(MSHFlexGrid1.TextMatrix(I, 44), "dd-mmm-yyyy hh:mm:ss")
       End If
       
       TmpRs.MoveNext
   Next
End If

LblTotalRecord.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdCancelReverse_Click()
TxtReverseDate.Value = Date
TxtReverseRemark = ""
Frame6.Visible = False
End Sub

Private Sub CmdViewInvoice_Click()

On Error GoTo Error:
'Shell (App.Path & "\Invoices\" & TxtEntryNo & " ")
If mEntryNo <> "" Then
   wsPathAdobe = Gen_ReadPathConfigINI("PDFPATH") '"\\10.0.0.8\Reader\AcroRD32.exe" '"C:\Program Files\Adobe\Acrobat 6.0\Reader\AcroRd32.exe" '"\\10.0.0.8\Reader\AcroRD32.exe"
   mFileName = App.Path & "\NavisionInvoice\" & mEntryNo & ".pdf"
   'mFileName = App.Path & "\Invoices\" & Replace(TxtPostBOEID, "/", "") & ".pdf"
   If wsPathAdobe <> "" Then
      iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
   End If
   
End If
Exit Sub

Error:
      MsgBox "Please contact Technology department!!!"

End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
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
RsTxn_CDTFBOE.MoveFirst
RsTxn_CDTFBOE.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 13) & "'"
If Not RsTxn_CDTFBOE.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub EditCmd_Click()

If Val(TxtBatchID) <> 0 Then
   MsgBox "Batch generated for this Bill of Estimation!!!"
   Exit Sub
End If

If LCase(TxtStatus) = "confirmed" Then
   If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "headcmg" Then
      MsgBox "Bill of Estimation is confirmed. For any modification please contact CMG Dept.!!!"
      Exit Sub
   End If
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(TxtBillingMonth.Text), "yyyy-mm") Then
   MsgBox "Billing month is previous month from " & Format(Gen_BlockingDate, "MMM-yyyy") & ", you can not edit!!! "
   Exit Sub
End If



'Frame1.Enabled = True

TxtDiscount.Locked = False
TxtPremium.Locked = False
TxtRemarks.Locked = False
CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = TxtStatus

Call GButtonLock(Me, False)
CmdReverse.Enabled = False
CmdPrint.Enabled = False
TxtDiscount.SetFocus
CmdViewInvoice.Enabled = False

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
   CmdPrint.Enabled = False
   Exit Sub
End If
If Gen_WcLocation <> "" Then
   Call TableTxn_CDTFBOE(" Where " & Gen_WcLocation)
Else
   Call TableTxn_CDTFBOE
End If
MSHFlexGrid1.Clear
Call CmdClear_Click
Call Grid_Head1
CmdReverse.Enabled = False
If RsTxn_CDTFBOE.RecordCount = 0 Then
   Call GButtonLockAD(Me)
   Exit Sub
End If

Call Indata

End Sub


Public Sub Indata()

'Frame1.Enabled = False
CmdReverse.Enabled = False
Frame6.Visible = False

SSTab1.Tab = 0

mLocationID = IIf(IsNull(RsTxn_CDTFBOE!LocationID), 0, RsTxn_CDTFBOE!LocationID)
TxtLocationID = GetLocationName(mLocationID)
TxtPreFix = RsTxn_CDTFBOE!SM_Prefix
mBOEID = RsTxn_CDTFBOE!BOEID
TxtOriginalBOEID = RsTxn_CDTFBOE!BOEID

If RsTxn_CDTFBOE!TxnType = "W" Then
   TxtBOEID = "RJ" & RsTxn_CDTFBOE!BOEID
   TxtInvoiceType = "Rejected"
ElseIf RsTxn_CDTFBOE!TxnType = "D" Then
   TxtBOEID = "SD" & RsTxn_CDTFBOE!BOEID
   TxtInvoiceType = "7 Days"
ElseIf RsTxn_CDTFBOE!TxnType = "R" Then
   TxtBOEID = "RRN" & RsTxn_CDTFBOE!BOEID
   TxtInvoiceType = "RRN Withdrawal"
ElseIf RsTxn_CDTFBOE!TxnType = "Y" Then
   TxtBOEID = "NNNRY" & RsTxn_CDTFBOE!BOEID
   TxtInvoiceType = "NNNR Yearly"
ElseIf RsTxn_CDTFBOE!TxnType = "P" Then
   TxtBOEID = "PWR" & RsTxn_CDTFBOE!BOEID
   TxtInvoiceType = "PWR"
End If

TxtInvoiceDate = Format(RsTxn_CDTFBOE!InvoiceDate, "dd-mmm-yyyy")
mClientIDRow = RsTxn_CDTFBOE!ClientID
TxtClientID = GetClientName(IIf(IsNull(RsTxn_CDTFBOE!ClientID), "0", RsTxn_CDTFBOE!ClientID))
TxtCDTFNo = IIf(IsNull(RsTxn_CDTFBOE!CxTF_DetailsID), "", RsTxn_CDTFBOE!CxTF_DetailsID)
TxtCDTFDate = Format(RsTxn_CDTFBOE!CDTFDate, "dd-MMM-yyyy")
TxtBillingMonth = Format(RsTxn_CDTFBOE!BillingMonth, "mmm-yyyy")
mTotalDays = Day(DateSerial(Year(RsTxn_CDTFBOE!BillingMonth), Month(RsTxn_CDTFBOE!BillingMonth) + 1, 0))
mFAGToDate = (Year(RsTxn_CDTFBOE!BillingMonth) & "-" & Month(RsTxn_CDTFBOE!BillingMonth) & "-" & mTotalDays)
'TxtInvoiceType = IIf(RsTxn_CDTFBOE!TxnType = "W", "Rejected", "7 Days")
TxtAmount = IIf(IsNull(RsTxn_CDTFBOE!Amount), 0, RsTxn_CDTFBOE!Amount)
If IsNull(RsTxn_CDTFBOE!FAGPostingDate) = False Then
   TxtFAGPostingDate.Text = Format(RsTxn_CDTFBOE!FAGPostingDate, "dd-MMM-yyyy")
Else
   TxtFAGPostingDate.Text = ""
End If

TxtFAGRemarks = IIf(IsNull(RsTxn_CDTFBOE!FAGRemarks), "", RsTxn_CDTFBOE!FAGRemarks)
''TxtServiceTaxNAVCode = IIf(IsNull(RsTxn_CDTFBOE!ServiceTaxNAVCode), "", RsTxn_CDTFBOE!ServiceTaxNAVCode)

If RsTxn_CDTFBOE!TxnType = "D" Then
   tmp = "Select TCD.CxTF_DetailsID,TCD.CxTFDate,Commodity,ExchangeType,TCD.ExchangeTypeID,MS.NavCode From  Txn_CxTF_Details TCD"
   tmp = tmp & " inner join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
   tmp = tmp & " inner join Mst_ServiceTax MS On MComm.ServiceTaxID = MS.ServiceTaxID"
   tmp = tmp & " inner join Mst_ExchangeType ME On TCD.ExchangeTypeID = ME.ExchangeTypeID"
   tmp = tmp & " Where TxnType = 'D' and CxTF_DetailsID=" & Val(RsTxn_CDTFBOE!CxTF_DetailsID) & "And SM_Prefix='" & TxtPreFix & "'"
ElseIf RsTxn_CDTFBOE!TxnType = "W" Or RsTxn_CDTFBOE!TxnType = "R" Or RsTxn_CDTFBOE!TxnType = "Y" Then
   tmp = "Select TCD.CxTF_DetailsID,TCD.CxTFDate,Commodity,ExchangeType,TCD.ExchangeTypeID,MS.NavCode From  Txn_CxTF_Details TCD"
   tmp = tmp & " inner join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
   tmp = tmp & " inner join Mst_ServiceTax MS On MComm.ServiceTaxID = MS.ServiceTaxID"
   tmp = tmp & " inner join Mst_ExchangeType ME On TCD.ExchangeTypeID = ME.ExchangeTypeID"
   tmp = tmp & " Where TxnType = 'W' and CxTF_DetailsID=" & Val(RsTxn_CDTFBOE!CxTF_DetailsID) & "And SM_Prefix='" & TxtPreFix & "'"
ElseIf RsTxn_CDTFBOE!TxnType = "P" Then
   tmp = "Select TCD.ConvertedID,TCD.PWRDate,Commodity,'2' 'ExchangeType',MS.NavCode From  Txn_RRNPWR TCD"
   tmp = tmp & "  inner join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
   tmp = tmp & " inner join Mst_ServiceTax MS On MComm.ServiceTaxID = MS.ServiceTaxID"
   tmp = tmp & " Where ConvertedID = " & Val(RsTxn_CDTFBOE!CxTF_DetailsID) & ""
End If
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   If RsTxn_CDTFBOE!TxnType <> "P" Then
      mExchangeTypeID = IIf(IsNull(TmpRs1!ExchangeTypeID), 0, TmpRs1!ExchangeTypeID)
      TxtExchangeType = IIf(IsNull(TmpRs1!ExchangeType), "", TmpRs1!ExchangeType)
   Else
      mExchangeTypeID = 2
      TxtExchangeType = "NCDEX"
   End If
   TxtCommodity = IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
   TxtServiceTaxNAVCode = IIf(IsNull(TmpRs1!NavCode), "", TmpRs1!NavCode)
Else
   mExchangeTypeID = 0
   TxtExchangeType = ""
   TxtCommodity = ""
   TxtServiceTaxNAVCode = ""
End If

TxtBatchID = IIf(IsNull(RsTxn_CDTFBOE!BatchID), 0, RsTxn_CDTFBOE!BatchID)
TxtDiscount = IIf(IsNull(RsTxn_CDTFBOE!Discount), 0, RsTxn_CDTFBOE!Discount)
TxtPremium = IIf(IsNull(RsTxn_CDTFBOE!Premium), 0, RsTxn_CDTFBOE!Premium)

TxtDiscount.Locked = True
TxtPremium.Locked = True
TxtRemarks.Locked = True
CmbStatus.Visible = False
TxtStatus.Visible = True

mFlag = IIf(IsNull(RsTxn_CDTFBOE!Flag), "P", RsTxn_CDTFBOE!Flag)

If mFlag = "P" Then
   TxtStatus = "Pending"
ElseIf mFlag = "C" Then
   TxtStatus = "Confirmed"
End If

TxtRemarks = IIf(IsNull(RsTxn_CDTFBOE!Remarks), "", RsTxn_CDTFBOE!Remarks)
TxtG_UID = IIf(IsNull(RsTxn_CDTFBOE!G_UID), "", RsTxn_CDTFBOE!G_UID)

TxtCreated = IIf(IsNull(RsTxn_CDTFBOE!CreatedBy), "", RsTxn_CDTFBOE!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CDTFBOE!CreatedDate), "", Format(RsTxn_CDTFBOE!CreatedDate, "dd-mmm-yyyy hh:mm:ss"))
TxtUpdated = IIf(IsNull(RsTxn_CDTFBOE!UpdatedBy), "", RsTxn_CDTFBOE!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CDTFBOE!UpdatedDate), "", Format(RsTxn_CDTFBOE!UpdatedDate, "dd-mmm-yyyy hh:mm:ss"))

Call IndataGSLDetails
Call IndataTIDetail

CmdPrint.Enabled = True
CmdViewInvoice.Enabled = True
'
'If RsTxn_CDTFBOE.RecordCount = 1 Then
'   Call GButtonLock_1(Me, True)
'Else
   Call GButtonLock(Me, True)
'End If

End Sub
Private Sub IndataGSLDetails()
Dim mYear, mMonth, mFromDate As Variant
Call Grid_Head1

''mYear = Year(RsTxn_CDTFBOE!InvoiceDate)
''mMonth = Month(RsTxn_CDTFBOE!InvoiceDate)
''mFromDate = Format(mYear & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

If RsTxn_CDTFBOE!TxnType = "R" Or RsTxn_CDTFBOE!TxnType = "P" Then
   tmp = "Select MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,GSL.StackNo,GSL.LotNo,"
   tmp = tmp & " TIMGD.CDTFBags,TIMGD.CDTFQty,BillingCycle,TIMGD.BillingUnit,GSL.CDTFDate,TIMGD.NoofDays,TIMGD.Rate,"
   tmp = tmp & " TIMGD.Amount,MRRN.RRNDate,MIS.FED,MRRN.RRNNO,TIMGD.NCDEXSpaceCharges"
   tmp = tmp & " From Txn_CDTFBOE_GSL TIMGD"
   tmp = tmp & " Inner Join Txn_CDTFBOE TPI On TIMGD.BOEID = TPI.BOEID And TIMGD.SM_Prefix = TPI.SM_Prefix"
   tmp = tmp & " Inner Join Mst_GSL GSL On TIMGD.Mst_GSL_ID = GSL.Mst_GSL_ID And TIMGD.SM_Prefix = GSL.SM_Prefix"
   tmp = tmp & " Inner Join Mst_CMP MC On GSL.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_BillingCycle MBC On TIMGD.BillingCycleID = MBC.BillingCycleID"
   tmp = tmp & " Left Join Mst_ISIN MIS on GSL.ISINID=MIS.ISINID"
   tmp = tmp & " Left Join Mst_RRN_LotDetail MRRND on MIS.ISINID=MRRND.ISINID"
   tmp = tmp & " Left Join Mst_RRN MRRN on MRRND.RRNID=MRRN.RRNID"
   tmp = tmp & " Where TPI.BOEID = " & mBOEID & ""
   tmp = tmp & " Order By MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,GSL.StackNo,GSL.LotNo "
Else
   tmp = "Select MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,GSL.StackNo,GSL.LotNo,"
   tmp = tmp & " TIMGD.CDTFBags,TIMGD.CDTFQty,BillingCycle,TIMGD.BillingUnit,GSL.CDTFDate,TIMGD.NoofDays,TIMGD.Rate,"
   tmp = tmp & " TIMGD.Amount,'' as RRNDate,'' as FED,'' as RRNNO,TIMGD.NCDEXSpaceCharges"
   tmp = tmp & " From Txn_CDTFBOE_GSL TIMGD"
   tmp = tmp & " Inner Join Txn_CDTFBOE TPI On TIMGD.BOEID = TPI.BOEID And TIMGD.SM_Prefix = TPI.SM_Prefix"
   tmp = tmp & " Inner Join Mst_GSL GSL On TIMGD.Mst_GSL_ID = GSL.Mst_GSL_ID And TIMGD.SM_Prefix = GSL.SM_Prefix"
   tmp = tmp & " Inner Join Mst_CMP MC On GSL.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_BillingCycle MBC On TIMGD.BillingCycleID = MBC.BillingCycleID"
   tmp = tmp & " Where TPI.BOEID = " & mBOEID & ""
   tmp = tmp & " Order By MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,GSL.StackNo,GSL.LotNo "
End If

Call TmpTable

If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       If LCase(MSHFlexGrid2.TextMatrix(x, 8)) = "q" Then
          MSHFlexGrid2.TextMatrix(x, 8) = "Quantity"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 8)) = "b" Then
          MSHFlexGrid2.TextMatrix(x, 8) = "Bags"
       End If
       
       MSHFlexGrid2.TextMatrix(x, 9) = Format(MSHFlexGrid2.TextMatrix(x, 9), "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(0, 12) = Val(MSHFlexGrid2.TextMatrix(0, 12)) + Val(MSHFlexGrid2.TextMatrix(x, 12))
       MSHFlexGrid2.TextMatrix(0, 16) = Val(MSHFlexGrid2.TextMatrix(0, 16)) + Val(MSHFlexGrid2.TextMatrix(x, 16))
       MSHFlexGrid2.TextMatrix(x, 13) = Format(MSHFlexGrid2.TextMatrix(x, 13), "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(x, 14) = Format(MSHFlexGrid2.TextMatrix(x, 14), "dd-MMM-yyyy")
       x = x + 1
       TmpRs.MoveNext
   Next
End If

End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub


End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CDTFBOE.MoveNext
If RsTxn_CDTFBOE.EOF Then
   RsTxn_CDTFBOE.MoveLast
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
RsTxn_CDTFBOE.MovePrevious
If RsTxn_CDTFBOE.BOF Then
   RsTxn_CDTFBOE.MoveFirst
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
RsTxn_CDTFBOE.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CDTFBOE.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SaveCmd_Click()

If TxtBOEID = "" Then
   MsgBox "No Invoice No found to update!!!"
   Exit Sub
End If

If Val(TxtPremium) <> 0 And Val(TxtDiscount) <> 0 Then
   MsgBox "Either of both 'Discount' or 'Premium' is allowed!!!"
   TxtDiscount.SetFocus
   Exit Sub
End If

If CmbStatus.Text = "" Then
   MsgBox "Blank Status not allowed!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If TxtRemarks.Text = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtRemarks.SetFocus
   Exit Sub
End If

RsTxn_CDTFBOE.MoveFirst
RsTxn_CDTFBOE.Find "G_UID = '" & TxtG_UID & "'"

If RsTxn_CDTFBOE.EOF = False Then
   RsTxn_CDTFBOE!Discount = Val(TxtDiscount)
   RsTxn_CDTFBOE!Premium = Val(TxtPremium)
   RsTxn_CDTFBOE!Remarks = Trim(TxtRemarks)
   RsTxn_CDTFBOE!Flag = Left(CmbStatus.Text, 1)
   
''   If LCase(CmbStatus) <> "pending" Then
''      RsTxn_CDTFBOE!FAGRemarks = TxtFAGRemarks.Text
''      RsTxn_CDTFBOE!FAGPostingDate = TxtFAGPostingDate
''   Else
''      RsTxn_CDTFBOE!FAGRemarks = Null
''      RsTxn_CDTFBOE!FAGPostingDate = Null
''   End If

   RsTxn_CDTFBOE!UpdatedBy = Gen_UserLoginID
   RsTxn_CDTFBOE!UpdatedDate = Now
   RsTxn_CDTFBOE.Update
End If

RsTxn_CDTFBOE.MoveFirst
RsTxn_CDTFBOE.Find "G_UID = '" & TxtG_UID & "'"

Call Indata

End Sub

Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CDTFBOE.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsTxn_CDTFBOE.EOF Then
      RsTxn_CDTFBOE.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 100
Frame0.Left = Frame1.Left '- 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width '- 100
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1800

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call CmdClear_Click
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtDiscount_LostFocus()
If TxtDiscount = "" Then Exit Sub
If IsNumeric(TxtDiscount) = False Or Val(TxtDiscount) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtDiscount.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtPremium_LostFocus()
If TxtPremium = "" Then Exit Sub
If IsNumeric(TxtPremium) = False Or Val(TxtPremium) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtPremium.SetFocus
   Exit Sub
End If
End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 45 '43 '36 '35 '34

MSHFlexGrid1.TextMatrix(0, 0) = "Invoice No"
MSHFlexGrid1.TextMatrix(0, 1) = "Invoice Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Billing Month"
MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(0, 4) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 5) = "CxTF No"
MSHFlexGrid1.TextMatrix(0, 6) = "Client"
MSHFlexGrid1.TextMatrix(0, 7) = "Amount"
MSHFlexGrid1.TextMatrix(0, 8) = "Discount"
MSHFlexGrid1.TextMatrix(0, 9) = "Premium"
MSHFlexGrid1.TextMatrix(0, 10) = "Batch ID"
MSHFlexGrid1.TextMatrix(0, 11) = "GST Group Code" '"Service Tax NAVCode"
MSHFlexGrid1.TextMatrix(0, 12) = "FAGPosting Date"
MSHFlexGrid1.TextMatrix(0, 13) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 14) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 15) = "Status"
MSHFlexGrid1.TextMatrix(0, 16) = "TI ID"
MSHFlexGrid1.TextMatrix(0, 17) = "TI No"
MSHFlexGrid1.TextMatrix(0, 18) = "TI Date"
MSHFlexGrid1.TextMatrix(0, 19) = "TI Amount"
MSHFlexGrid1.TextMatrix(0, 20) = "TI Service Tax Amount"

MSHFlexGrid1.TextMatrix(0, 21) = "Received Amount"
MSHFlexGrid1.TextMatrix(0, 22) = "Received TDS"
MSHFlexGrid1.TextMatrix(0, 23) = "Received Other Ded."
MSHFlexGrid1.TextMatrix(0, 24) = "Post Invoice No"
MSHFlexGrid1.TextMatrix(0, 25) = "TI Entry No"
MSHFlexGrid1.TextMatrix(0, 26) = "Reverse Flag"
MSHFlexGrid1.TextMatrix(0, 27) = "Reverse Date"
MSHFlexGrid1.TextMatrix(0, 28) = "Reverse Remark"
MSHFlexGrid1.TextMatrix(0, 29) = "Reverse Batch ID"
MSHFlexGrid1.TextMatrix(0, 30) = "Commodity "
MSHFlexGrid1.TextMatrix(0, 31) = "Quantity"
MSHFlexGrid1.TextMatrix(0, 32) = "Rate Count"
MSHFlexGrid1.TextMatrix(0, 33) = "Rate"
MSHFlexGrid1.TextMatrix(0, 34) = "Invoice Type"
MSHFlexGrid1.TextMatrix(0, 35) = "NCDEX Space Charges"
MSHFlexGrid1.TextMatrix(0, 36) = "HSN Code"
MSHFlexGrid1.TextMatrix(0, 37) = "Account Type"

MSHFlexGrid1.TextMatrix(0, 38) = "State India Name"
MSHFlexGrid1.TextMatrix(0, 39) = "NCML GST No"
MSHFlexGrid1.TextMatrix(0, 40) = "Client GST No"
MSHFlexGrid1.TextMatrix(0, 41) = "Client GST NAVCode"
MSHFlexGrid1.TextMatrix(0, 42) = "Tax Applied (Inter/Intra)"
MSHFlexGrid1.TextMatrix(0, 43) = "Invoice Created Date"
MSHFlexGrid1.TextMatrix(0, 44) = "Transaction Created Date"

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 3200
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 0
MSHFlexGrid1.ColWidth(14) = 3000
MSHFlexGrid1.ColWidth(15) = 1600
MSHFlexGrid1.ColWidth(16) = 1600
MSHFlexGrid1.ColWidth(17) = 1600
MSHFlexGrid1.ColWidth(18) = 1600
MSHFlexGrid1.ColWidth(19) = 1600
MSHFlexGrid1.ColWidth(20) = 1600
MSHFlexGrid1.ColWidth(21) = 1600
MSHFlexGrid1.ColWidth(26) = 2500
MSHFlexGrid1.ColWidth(27) = 1200
MSHFlexGrid1.ColWidth(28) = 1200
MSHFlexGrid1.ColWidth(29) = 1200
MSHFlexGrid1.ColWidth(30) = 1200

End Sub

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For I = 2010 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
Call Grid_Head

If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
flg = "N"
For I = 2010 To Year(Now)
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

End Sub
Private Sub CmbMonth_LostFocus()
Call Grid_Head
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
       CmbMonth.Text = CmbMonth.List(I)
       mStatus = True
       Exit For
    End If
Next

If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = 2 ") '" & mExchangeTypeID)
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
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_CDTFBOE.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_CDTFBOE.xls")
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

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If Val(TxtBatchID) <> 0 Then
   MsgBox "Batch generated for this Bill of Estimation!!!"
   Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   Frame3.Visible = True
   TxtLogRemark.Text = ""
'   Call DeleteChildTableDetail
'   RsTxn_CDTFBOE.Delete
'   RsTxn_CDTFBOE.Update
'   'RsTxn_CDTFBOE.Requery
'   If RsTxn_CDTFBOE.RecordCount = 0 Then
'      Call ClearData
'      Call GButtonLockAD(Me)
'      'Call AddCmd_Click
'      Exit Sub
'   End If
'   RsTxn_CDTFBOE.MoveNext
'   If RsTxn_CDTFBOE.EOF() Then
'      RsTxn_CDTFBOE.MoveLast
'   End If
'   Call InData
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_CDTFBOE.CancelUpdate
End Sub
Private Sub ClearData()

TxtLocationID = ""
TxtPreFix = ""
TxtBOEID = ""
TxtInvoiceDate = ""
TxtClientID = ""

TxtBillingMonth = ""
TxtAmount = 0
TxtFAGPostingDate = ""
TxtFAGRemarks = ""
TxtServiceTaxNAVCode = ""

TxtExchangeType = ""
TxtBatchID = ""
TxtDiscount = 0
TxtPremium = 0

TxtG_UID = ""

TxtCreated = ""
TxtUpdated = ""

Call Grid_Head1



End Sub
Private Sub DeleteChildTableDetail()

Dim mSMonth, mSYear, mSFromDate As Variant

mSMonth = Month(CDate(TxtInvoiceDate))
mSYear = Year(CDate(TxtInvoiceDate))

mSFromDate = Format(mSYear & "-" & mSMonth & "-" & "01", "yyyy-mm-dd")

If LCase(TxtInvoiceType) <> "rejected" And LCase(TxtInvoiceType) <> "rrn withdrawal" Then
   tmp = "Update Txn_CxTF_GSL Set BillFlag=Null Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & TxtCDTFNo & " And TxnType='D' And BillFlag='I' And Mst_GSL_ID in (Select Mst_GSL_ID From Txn_CDTFBOE_GSL Where BOEID=" & Val(TxtOriginalBOEID) & ")"
Else
   tmp = "Update Txn_CxTF_GSL Set BillFlag=Null Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & TxtCDTFNo & " And TxnType='W' And BillFlag='I' And Mst_GSL_ID in (Select Mst_GSL_ID From Txn_CDTFBOE_GSL Where BOEID=" & Val(TxtOriginalBOEID) & ")"
End If

Call TmpTable

tmp = "Delete From Txn_CDTFBOE_GSL Where BOEID=" & TxtOriginalBOEID
Call TmpTable

End Sub


Private Sub Findwithrdrawal(SM_Prefix_ As Variant, mFromDate_ As Variant, x_ As Variant, gslID_ As Variant)
Dim Retval As Double

tmp = "Select Sum(TI.BalanceBags) 'With_Bal_Bags',Sum(TI.BalanceWeight) 'With_Bal_Wt' "
tmp = tmp & " From Txn_InvMonGSLDetail TI"
tmp = tmp & " Where  TI.YearMonth='" & mFromDate_ & "' and  TI.SM_Prefix='" & SM_Prefix_ & "' And TI.Mst_GSL_ID = " & gslID_ & " And TI.TFlag='W' " 'And TI.CMPID =" & CMPID_ & " And TI.CommodityID=" & CommodityID_ & " "
    
    
Call Tmp2Table
With MSHFlexGrid2
   If TmpRs2.RecordCount > 0 Then
      If IsNull(TmpRs2!With_Bal_Bags) = False Then
         .TextMatrix(x_, 6) = Val(.TextMatrix(x_, 6)) - TmpRs2!With_Bal_Bags
         .TextMatrix(x_, 7) = Val(.TextMatrix(x_, 7)) - TmpRs2!With_Bal_Wt
      End If
   End If
End With
End Sub

Private Sub TxtFromDate_LostFocus()
If Check1.Value = 0 Then
   TxtFromDate.Value = CDate(Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-" & "01")
End If
End Sub

Private Sub TxtToDate_LostFocus()
If Check1.Value = 0 Then
   TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
End If
End Sub
Private Function GetLastDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function
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
VScroll1.Max = Me.ScaleHeight - Frame1.Height '4
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame1.Width '4
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame1.Height >= Me.Height Then VScroll1.Enabled = True ''4
If Frame1.Height < Me.Height Then VScroll1.Enabled = False '4
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

Private Sub CmdPrint_Click()

Gen_mTimeStamp = GetTimeStamp

'Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_CDTFBOEGSLwiseDetails.xls")

Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "_BOE-" & Trim(TxtPreFix) & "" & Trim(TxtBOEID) & ".xls")

Call Xls_Print_Rpt

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
 
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_BOE-" & Trim(TxtPreFix) & "" & Trim(TxtBOEID) & ".xls")

I = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
oWS.Name = TxtPreFix & "-" & TxtBOEID  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 2

'oWS.Range("a1:g1").Merge
'oWS.Range("a:a").ColumnWidth = 10.86
'oWS.Range("b:b").ColumnWidth = 12.86
'oWS.Range("c:c").ColumnWidth = 12.86
'oWS.Range("d:d").ColumnWidth = 12.86
'oWS.Range("e:e").ColumnWidth = 12.86 '32.14
'oWS.Range("f:f").ColumnWidth = 12.86
'oWS.Range("g:g").ColumnWidth = 12.86
'

mRow = mRow + 1
oWS.Cells(mRow, C + 1) = "Invoice No"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtBOEID.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 4) = "Invoice Date"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 5) = TxtInvoiceDate.Text
oWS.Cells(mRow, C + 5).HorizontalAlignment = 2
oWS.Cells(mRow, C + 5).Font.Size = 7


oWS.Cells(mRow, C + 7) = "Exchange Type"
oWS.Cells(mRow, C + 7).Font.Bold = True
oWS.Cells(mRow, C + 7).WrapText = True
tmp = "h" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 7).Font.Size = 7
oWS.Cells(mRow, C + 8) = TxtExchangeType.Text
oWS.Cells(mRow, C + 8).HorizontalAlignment = 2
oWS.Cells(mRow, C + 8).Font.Size = 7

oWS.Cells(mRow, C + 10) = "Billing Month"
oWS.Cells(mRow, C + 10).Font.Bold = True
oWS.Cells(mRow, C + 10).WrapText = True
tmp = "k" & mRow & ":" & "p" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 10).Font.Size = 7
oWS.Cells(mRow, C + 11) = TxtBillingMonth.Text
oWS.Cells(mRow, C + 11).HorizontalAlignment = 2
oWS.Cells(mRow, C + 11).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "CxTF No" '"CDTF No"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtCDTFNo.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 7) = "CxTF Date" '"CDTF Date"
oWS.Cells(mRow, C + 7).Font.Bold = True
oWS.Cells(mRow, C + 7).WrapText = True
tmp = "h" & mRow & ":" & "p" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 7).Font.Size = 7
oWS.Cells(mRow, C + 8) = TxtCDTFDate.Text
oWS.Cells(mRow, C + 8).HorizontalAlignment = 2
oWS.Cells(mRow, C + 8).Font.Size = 7


mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Location"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtLocationID.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 7) = "Client Name"
oWS.Cells(mRow, C + 7).Font.Bold = True
oWS.Cells(mRow, C + 7).WrapText = True
tmp = "h" & mRow & ":" & "p" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 7).Font.Size = 7
oWS.Cells(mRow, C + 8) = TxtClientID.Text
oWS.Cells(mRow, C + 8).HorizontalAlignment = 2
oWS.Cells(mRow, C + 8).Font.Size = 7


mRow = mRow + 1
oWS.Cells(mRow, C + 1) = "Amount"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtAmount.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 4) = "ST Nav Code"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 5) = TxtServiceTaxNAVCode.Text
oWS.Cells(mRow, C + 5).HorizontalAlignment = 2
oWS.Cells(mRow, C + 5).Font.Size = 7


oWS.Cells(mRow, C + 7) = "FAG Posting Date"
oWS.Cells(mRow, C + 7).Font.Bold = True
oWS.Cells(mRow, C + 7).WrapText = True
tmp = "h" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 7).Font.Size = 7
oWS.Cells(mRow, C + 8) = TxtFAGPostingDate.Text
oWS.Cells(mRow, C + 8).HorizontalAlignment = 2
oWS.Cells(mRow, C + 8).Font.Size = 7


oWS.Cells(mRow, C + 10) = "Batch ID"
oWS.Cells(mRow, C + 10).Font.Bold = True
oWS.Cells(mRow, C + 10).WrapText = True
tmp = "k" & mRow & ":" & "p" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 10).Font.Size = 7
oWS.Cells(mRow, C + 11) = TxtBatchID.Text
oWS.Cells(mRow, C + 11).HorizontalAlignment = 2
oWS.Cells(mRow, C + 11).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "FAG Remark"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "p" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtFAGRemarks.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Discount"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtDiscount.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 4) = "Premium"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 5) = TxtPremium.Text
oWS.Cells(mRow, C + 5).HorizontalAlignment = 2
oWS.Cells(mRow, C + 5).Font.Size = 7

tmp = "g" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 7) = "Total BOE Amount(Amt + Premium) - Discount"
oWS.Cells(mRow, C + 7).Font.Bold = True
oWS.Cells(mRow, C + 7).WrapText = True
oWS.Cells(mRow, C + 7).Font.Size = 7

tmp = "k" & mRow & ":" & "p" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 11) = (Val(TxtAmount) + Val(TxtPremium)) - Val(TxtDiscount)
oWS.Cells(mRow, C + 11).HorizontalAlignment = 2
oWS.Cells(mRow, C + 11).Font.Size = 7


mRow = mRow + 1

tmp = "a" & mRow & ":" & "p" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "GSL Details"
oWS.Cells(mRow, C + 1).Font.Size = 9
oWS.Cells(mRow, C + 1).Font.Bold = True

mRow = mRow + 1

For k = 1 To MSHFlexGrid2.Rows - 1
    For Col = 0 To MSHFlexGrid2.Cols - 1
        oWS.Cells(mRow, Col + 1) = MSHFlexGrid2.TextMatrix(k, Col)
        oWS.Cells(mRow, Col + 1).Font.Size = 7
        If k = 1 Then
           oWS.Cells(mRow, Col + 1).Font.Bold = True
           oWS.Cells(mRow, Col + 1).Font.Size = 7
        End If
    Next
    mRow = mRow + 1
Next k
   
'mRow = mRow + 1

tmp = "a" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Total"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True

oWS.Cells(mRow, C + 13) = MSHFlexGrid2.TextMatrix(0, 12)
oWS.Cells(mRow, C + 13).Font.Size = 7


For Cell = 3 To mRow
    tmp = "a" & Cell & ":" & "p" & Cell
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub
Private Sub IndataTIDetail()

''tmp = " Select TDI.TINO,TDI.TIAmount,TDI.TIDate,TDI.TIID,TDI.EntryNo From Txn_TIDetails TDI"
''tmp = tmp & " Where PIID=" & RsTxn_CDTFBOE!BOEID & " and TDI.PIFLag='NNNR'"
''
''Call TmpTable
''
''If TmpRs.RecordCount > 0 Then
''   TxtTINo = IIf(IsNull(TmpRs!TINO), "", TmpRs!TINO)
''   TxtTIDate = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-mm-yyyy"))
''   TxtTIAmount = IIf(IsNull(TmpRs!TIAmount), "", TmpRs!TIAmount)
''   mEntryNo = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
''   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "power" Then
''      If IIf(IsNull(RsTxn_CDTFBOE!ReverseFlag), "N", RsTxn_CDTFBOE!ReverseFlag) <> "Y" Then
''         If CheckTIInCR(TmpRs!TIID) = False Then
''            CmdReverse.Enabled = True
''         End If
''      End If
''   End If
''Else
''   TxtTINo = ""
''   TxtTIDate = ""
''   TxtTIAmount = ""
''   mEntryNo = ""
''End If

Dim ans As Variant

If RsTxn_CDTFBOE!TxnType = "W" Then
   ans = GetTIDetails(RsTxn_CDTFBOE!BOEID, "RJ")
ElseIf RsTxn_CDTFBOE!TxnType = "D" Then
   ans = GetTIDetails(RsTxn_CDTFBOE!BOEID, "SD")
ElseIf RsTxn_CDTFBOE!TxnType = "R" Then
   ans = GetTIDetails(RsTxn_CDTFBOE!BOEID, "RRN")
ElseIf RsTxn_CDTFBOE!TxnType = "Y" Then
   ans = GetTIDetails(RsTxn_CDTFBOE!BOEID, "NNNRY")
ElseIf RsTxn_CDTFBOE!TxnType = "P" Then
   ans = GetTIDetails(RsTxn_CDTFBOE!BOEID, "PWR")
End If

TxtTINo = ans(1)
TxtTIDate = Format(ans(2), "dd-MMM-yyyy")
TxtTIAmount = ans(3)
TxtEntryNo = ans(5)
mEntryNo = TxtEntryNo
TxtTISTAmt = ans(7)
TxtRecAmt = Val(ans(8)) + Val(ans(9)) + Val(ans(10))
If Val(TxtRecAmt) > 0 Then
   TxtBalAmt = Format(Val(TxtTIAmount) - Val(TxtRecAmt), "#####0.00")
Else
   TxtBalAmt = TxtTIAmount
End If
TxtPostInvoiceNo = ans(4)
If TxtEntryNo <> "" And (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg") Then
   If IIf(IsNull(RsTxn_CDTFBOE!ReverseFlag), "N", RsTxn_CDTFBOE!ReverseFlag) <> "Y" Then
      If CheckTIInCR(Val(ans(0))) = False Then
         CmdReverse.Enabled = True
      End If
   End If
End If

End Sub

Private Sub CmbSInvoiceType_LostFocus()
If CmbSInvoiceType.Text = "" Then Exit Sub
If LCase(CmbSInvoiceType.Text) <> "7 days" And LCase(CmbSInvoiceType.Text) <> "rejected" And LCase(CmbSInvoiceType.Text) <> "rrn withdrawal" And LCase(CmbSInvoiceType.Text) <> "nnnr yearly" And LCase(CmbSInvoiceType.Text) <> "pwr" Then
   MsgBox "Invalid Selection!!!"
   CmbSInvoiceType.SetFocus
   Exit Sub
End If

If LCase(CmbSInvoiceType.Text) = "7 days" Then
   mSInvoiceType = 1
ElseIf LCase(CmbSInvoiceType.Text) = "rrn withdrawal" Then
   mSInvoiceType = 13
ElseIf LCase(CmbSInvoiceType.Text) = "nnnr yearly" Then
   mSInvoiceType = 7
ElseIf LCase(CmbSInvoiceType.Text) = "pwr" Then
   mSInvoiceType = 13
Else
   mSInvoiceType = 2
End If

mHSNNoNAGRI = GetGLAccNAVCodeFromInvoiceType(mSInvoiceType, "NAGRI", "H")
mHSNNoAGRI = GetGLAccNAVCodeFromInvoiceType(mSInvoiceType, "AGRI", "H")


End Sub




