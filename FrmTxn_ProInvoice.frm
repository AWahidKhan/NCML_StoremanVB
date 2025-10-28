VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_ProInvoice 
   Caption         =   "Bill of Estimation (BOE) Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
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
      Height          =   3375
      Left            =   4560
      TabIndex        =   107
      Top             =   3120
      Visible         =   0   'False
      Width           =   6375
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
         TabIndex        =   110
         Top             =   2880
         Width           =   1455
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
         TabIndex        =   109
         Top             =   2880
         Width           =   1440
      End
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
         Height          =   2400
         Left            =   1080
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   108
         Top             =   360
         Width           =   5160
      End
      Begin VB.Label Label33 
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
         TabIndex        =   111
         Top             =   1410
         Width           =   1260
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   30
      TabIndex        =   8
      Top             =   9720
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   10095
      Left            =   14955
      TabIndex        =   0
      Top             =   -165
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9525
      Left            =   15
      TabIndex        =   9
      Top             =   15
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
         TabIndex        =   80
         Top             =   3240
         Visible         =   0   'False
         Width           =   6375
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
            TabIndex        =   83
            Top             =   1440
            Width           =   1455
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
            TabIndex        =   82
            Top             =   1440
            Width           =   1440
         End
         Begin VB.TextBox TxtReverseRemark 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   1560
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   81
            Top             =   840
            Width           =   4680
         End
         Begin MSComCtl2.DTPicker TxtReverseDate 
            Height          =   375
            Left            =   1560
            TabIndex        =   84
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
            Format          =   120324097
            CurrentDate     =   39884
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
            TabIndex        =   86
            Top             =   390
            Width           =   1800
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
            TabIndex        =   85
            Top             =   930
            Width           =   1260
         End
      End
      Begin VB.Frame Frame0 
         Height          =   1845
         Left            =   0
         TabIndex        =   53
         Top             =   6000
         Visible         =   0   'False
         Width           =   14010
         Begin VB.CheckBox ChkConverted 
            Appearance      =   0  'Flat
            Caption         =   "Converted"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6210
            TabIndex        =   106
            Top             =   150
            Width           =   1425
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   78
            Top             =   840
            Width           =   225
         End
         Begin VB.ComboBox CmbNonNcdexClient 
            Height          =   315
            Left            =   5550
            Sorted          =   -1  'True
            TabIndex        =   60
            Top             =   420
            Width           =   5745
         End
         Begin VB.ComboBox CmbPreFix 
            Height          =   315
            Left            =   3030
            Sorted          =   -1  'True
            TabIndex        =   59
            Top             =   420
            Width           =   2520
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "Clear"
            Height          =   350
            Left            =   13665
            TabIndex        =   58
            Top             =   792
            Width           =   885
         End
         Begin VB.CommandButton CmdSearch 
            Caption         =   "Search"
            Height          =   350
            Left            =   11895
            TabIndex        =   57
            Top             =   792
            Width           =   885
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   12780
            TabIndex        =   56
            Top             =   792
            Width           =   885
         End
         Begin VB.TextBox TxtSInvoiceNo 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   13215
            TabIndex        =   55
            Top             =   420
            Width           =   1305
         End
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_ProInvoice.frx":0000
            Left            =   11325
            List            =   "FrmTxn_ProInvoice.frx":0010
            TabIndex        =   54
            Top             =   420
            Width           =   1845
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1215
            Left            =   90
            TabIndex        =   61
            Top             =   1185
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
            Left            =   120
            TabIndex        =   62
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
            Format          =   120324097
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   300
            Left            =   1575
            TabIndex        =   63
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
            Format          =   120324097
            CurrentDate     =   39884
         End
         Begin VB.Label Label30 
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
            Left            =   120
            TabIndex        =   79
            Top             =   840
            Width           =   1650
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
            Left            =   6750
            TabIndex        =   71
            Top             =   840
            Width           =   1935
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
            Left            =   8790
            TabIndex        =   70
            Top             =   840
            Width           =   75
         End
         Begin VB.Label Label23 
            Caption         =   "From Date"
            Height          =   225
            Left            =   412
            TabIndex        =   69
            Top             =   165
            Width           =   750
         End
         Begin VB.Label Label16 
            Caption         =   "Client Name"
            Height          =   225
            Left            =   8100
            TabIndex        =   68
            Top             =   165
            Width           =   885
         End
         Begin VB.Label Label34 
            Caption         =   "Location"
            Height          =   225
            Left            =   3945
            TabIndex        =   67
            Top             =   165
            Width           =   690
         End
         Begin VB.Label Label15 
            Caption         =   "Invoice No"
            Height          =   240
            Left            =   13440
            TabIndex        =   66
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label14 
            Caption         =   "To Date"
            Height          =   225
            Left            =   1927
            TabIndex        =   65
            Top             =   165
            Width           =   630
         End
         Begin VB.Label Label19 
            Caption         =   "Status"
            Height          =   240
            Left            =   11985
            TabIndex        =   64
            Top             =   240
            Width           =   525
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
         Left            =   15
         TabIndex        =   10
         Top             =   -45
         Width           =   14880
         Begin VB.TextBox TxtType 
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
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   3690
            Width           =   2025
         End
         Begin VB.TextBox TxtHSNNo 
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
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   3690
            Width           =   2025
         End
         Begin VB.TextBox TxtAdHocAmt 
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
            Left            =   12585
            Locked          =   -1  'True
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   2160
            Width           =   2025
         End
         Begin VB.TextBox TxtRejectionRemark 
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
            Height          =   465
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   3165
            Width           =   13020
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
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   675
            Width           =   5745
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
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   675
            Width           =   5505
         End
         Begin VB.TextBox TxtInvoiceNo 
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
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   240
            Width           =   2025
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
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
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
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1140
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
            Left            =   1680
            TabIndex        =   1
            Top             =   2160
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
            Left            =   5160
            TabIndex        =   2
            Top             =   2160
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
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   210
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
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   225
            Width           =   2025
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
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1140
            Width           =   2025
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
            Left            =   12585
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1140
            Width           =   2040
         End
         Begin VB.TextBox TxtFAGRemarks 
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
            Height          =   480
            Left            =   1680
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1605
            Width           =   13050
         End
         Begin VB.TextBox TxtOriginalInvoiceNo 
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
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   225
            Visible         =   0   'False
            Width           =   345
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
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   2580
            Width           =   13020
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
            ItemData        =   "FrmTxn_ProInvoice.frx":0038
            Left            =   8880
            List            =   "FrmTxn_ProInvoice.frx":0048
            TabIndex        =   3
            Top             =   2160
            Width           =   2055
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
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   2160
            Width           =   2025
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   4335
            Left            =   120
            TabIndex        =   72
            Top             =   4080
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7646
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "GSL Details"
            TabPicture(0)   =   "FrmTxn_ProInvoice.frx":0070
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "MSHFlexGrid2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "TI Details/Log History"
            TabPicture(1)   =   "FrmTxn_ProInvoice.frx":008C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "TxtPostInvoiceNo"
            Tab(1).Control(1)=   "TxtBalAmt"
            Tab(1).Control(2)=   "TxtTIAmount"
            Tab(1).Control(3)=   "TxtTIDate"
            Tab(1).Control(4)=   "TxtTINo"
            Tab(1).Control(5)=   "TxtTISTAmt"
            Tab(1).Control(6)=   "TxtRecAmt"
            Tab(1).Control(7)=   "TxtEntryNo"
            Tab(1).Control(8)=   "TxtUpdated"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "TxtCreated"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "Label38"
            Tab(1).Control(11)=   "Label28"
            Tab(1).Control(12)=   "Label27"
            Tab(1).Control(13)=   "Label22"
            Tab(1).Control(14)=   "Label21"
            Tab(1).Control(15)=   "Label20"
            Tab(1).Control(16)=   "Label25"
            Tab(1).Control(17)=   "Label26"
            Tab(1).Control(18)=   "Label47"
            Tab(1).Control(19)=   "Label46"
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
               Left            =   -73515
               Locked          =   -1  'True
               TabIndex        =   116
               Top             =   1230
               Width           =   5610
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
               Left            =   -64365
               Locked          =   -1  'True
               TabIndex        =   101
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
               Left            =   -69810
               Locked          =   -1  'True
               TabIndex        =   94
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
               Left            =   -71625
               Locked          =   -1  'True
               TabIndex        =   93
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
               Left            =   -73515
               Locked          =   -1  'True
               TabIndex        =   92
               Top             =   750
               Width           =   1890
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
               Left            =   -67995
               Locked          =   -1  'True
               TabIndex        =   91
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
               Left            =   -66180
               Locked          =   -1  'True
               TabIndex        =   90
               Top             =   750
               Width           =   1815
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
               Left            =   -62550
               Locked          =   -1  'True
               TabIndex        =   89
               Top             =   750
               Width           =   1815
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
               Left            =   -73515
               Locked          =   -1  'True
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   2325
               Width           =   10950
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
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   1800
               Width           =   10965
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   3735
               Left            =   120
               TabIndex        =   75
               Top             =   480
               Width           =   14460
               _ExtentX        =   25506
               _ExtentY        =   6588
               _Version        =   393216
               Rows            =   3
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
            Begin VB.Label Label38 
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
               Left            =   -74880
               TabIndex        =   117
               Top             =   1170
               Width           =   1335
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
               Left            =   -64215
               TabIndex        =   102
               Top             =   480
               Width           =   1695
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
               Left            =   -71055
               TabIndex        =   100
               Top             =   480
               Width           =   735
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
               Left            =   -69375
               TabIndex        =   99
               Top             =   480
               Width           =   975
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
               Left            =   -72765
               TabIndex        =   98
               Top             =   480
               Width           =   495
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
               Left            =   -67695
               TabIndex        =   97
               Top             =   480
               Width           =   1335
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
               Left            =   -66075
               TabIndex        =   96
               Top             =   480
               Width           =   1695
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
               Left            =   -62055
               TabIndex        =   95
               Top             =   480
               Width           =   855
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
               TabIndex        =   77
               Top             =   2385
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
               TabIndex        =   76
               Top             =   1860
               Width           =   1695
            End
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
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1140
            Width           =   2025
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
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   2520
            Visible         =   0   'False
            Width           =   225
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
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2520
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label35 
            Caption         =   "Type"
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
            Left            =   4080
            TabIndex        =   115
            Top             =   3735
            Width           =   480
         End
         Begin VB.Label Label32 
            Caption         =   "HSN No"
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
            Left            =   135
            TabIndex        =   114
            Top             =   3750
            Width           =   1440
         End
         Begin VB.Label Label31 
            Caption         =   "Adhoc Amount"
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
            Left            =   11280
            TabIndex        =   105
            Top             =   2108
            Width           =   1035
         End
         Begin VB.Label Label29 
            Caption         =   "Rejection/Freeze Remarks"
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
            Left            =   120
            TabIndex        =   103
            Top             =   3157
            Width           =   1560
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
            Left            =   7425
            TabIndex        =   42
            Top             =   690
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
            TabIndex        =   41
            Top             =   735
            Width           =   1440
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
            TabIndex        =   40
            Top             =   270
            Width           =   1440
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
            TabIndex        =   39
            Top             =   270
            Width           =   1200
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
            TabIndex        =   38
            Top             =   2535
            Visible         =   0   'False
            Width           =   960
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
            TabIndex        =   37
            Top             =   1155
            Width           =   1440
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
            Left            =   120
            TabIndex        =   36
            Top             =   2220
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
            Left            =   3960
            TabIndex        =   35
            Top             =   2190
            Width           =   855
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
            Left            =   11340
            TabIndex        =   34
            Top             =   270
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
            Left            =   7425
            TabIndex        =   33
            Top             =   285
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
            TabIndex        =   32
            Top             =   1095
            Width           =   1440
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
            Left            =   11340
            TabIndex        =   31
            Top             =   1185
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
            Left            =   7425
            TabIndex        =   30
            Top             =   1110
            Width           =   1440
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
            TabIndex        =   29
            Top             =   1680
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
            TabIndex        =   28
            Top             =   2580
            Visible         =   0   'False
            Width           =   600
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
            TabIndex        =   27
            Top             =   2700
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
            Left            =   7455
            TabIndex        =   26
            Top             =   2190
            Width           =   855
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1155
         Left            =   15
         TabIndex        =   43
         Top             =   8340
         Width           =   14880
         Begin VB.CommandButton CmdViewInvoice 
            Caption         =   "View Invoice"
            Enabled         =   0   'False
            Height          =   810
            Left            =   12960
            TabIndex        =   88
            Top             =   195
            Width           =   1740
         End
         Begin VB.CommandButton CmdReverse 
            Caption         =   "Reverse"
            Height          =   400
            Left            =   10845
            TabIndex        =   87
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   345
            TabIndex        =   52
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   345
            TabIndex        =   51
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2445
            TabIndex        =   6
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2445
            TabIndex        =   50
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   4545
            TabIndex        =   49
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   4545
            TabIndex        =   48
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   6645
            TabIndex        =   47
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   6645
            TabIndex        =   46
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   8745
            TabIndex        =   45
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   8745
            TabIndex        =   7
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Print"
            Height          =   400
            Left            =   10845
            TabIndex        =   44
            Top             =   195
            Width           =   2100
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_ProInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID As Double
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mMonth, mTotalDays, mYear As Variant
Dim mInvoiceNo As Variant
Dim mFlag As Variant
Dim mEntryNo As Variant
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 21 '20 '15
MSHFlexGrid2.TextMatrix(1, 0) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 1) = "GSLID"
MSHFlexGrid2.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 3) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 4) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 5) = "Commodity"
MSHFlexGrid2.TextMatrix(1, 6) = "No Of Bags"
MSHFlexGrid2.TextMatrix(1, 7) = "Quantity"
MSHFlexGrid2.TextMatrix(1, 8) = "Transaction Flag"
MSHFlexGrid2.TextMatrix(1, 9) = "CDTF No/CITF No"
MSHFlexGrid2.TextMatrix(1, 10) = "Billing Cycle"
MSHFlexGrid2.TextMatrix(1, 11) = "Billing Unit"
MSHFlexGrid2.TextMatrix(1, 12) = "No of Cycle"
MSHFlexGrid2.TextMatrix(1, 13) = "Rate per Cycle"
'-- Old
MSHFlexGrid2.TextMatrix(1, 14) = "Old Billing Cycle"
MSHFlexGrid2.TextMatrix(1, 15) = "Old Billing Unit"
MSHFlexGrid2.TextMatrix(1, 16) = "Old Rate per Cycle"

MSHFlexGrid2.TextMatrix(1, 17) = "Amount"
MSHFlexGrid2.TextMatrix(1, 18) = "Updated By"
MSHFlexGrid2.TextMatrix(1, 19) = "Updated On"
MSHFlexGrid2.TextMatrix(1, 20) = "Transaction Date"

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
MSHFlexGrid2.ColWidth(9) = 1000
MSHFlexGrid2.ColWidth(10) = 1000
MSHFlexGrid2.ColWidth(11) = 1000
MSHFlexGrid2.ColWidth(12) = 1000
MSHFlexGrid2.ColWidth(13) = 1000
MSHFlexGrid2.ColWidth(14) = 1000
MSHFlexGrid2.ColWidth(15) = 1000
MSHFlexGrid2.ColWidth(16) = 1000
MSHFlexGrid2.ColWidth(17) = 1000
MSHFlexGrid2.ColWidth(18) = 1000
MSHFlexGrid2.ColWidth(19) = 1000

End Sub

Private Sub AddCmd_Click()

Call GButtonLock(Me, False)
CmdPRint.Enabled = False
CmdReverse.Enabled = False
End Sub

Private Sub ChkConverted_Click()
CmbNonNcdexClient.Text = ""
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


Private Sub CmbSStatus_LostFocus()

If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "confirmed" And LCase(CmbSStatus.Text) <> "reject" And LCase(CmbSStatus.Text) <> "freeze" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbStatus_LostFocus()

TxtRejectionRemark.Locked = True
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "confirmed" And LCase(CmbStatus.Text) <> "reject" And LCase(CmbStatus.Text) <> "freeze" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If


If LCase(CmbStatus.Text) = "reject" Or LCase(CmbStatus.Text) = "freeze" Then
   TxtRejectionRemark.Locked = False
End If

End Sub

Private Sub CmdCancelLog_Click()
Frame2.Visible = False
TxtLogRemark.Text = ""
End Sub

Private Sub CmdClear_Click()
Dim mSMonth, mSYear, mSTotalDays As Variant
TxtFromDate.Value = Date
TxtToDate.Value = Date
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
Call TxtFromDate_LostFocus
Call TxtToDate_LostFocus
CmbPreFix.Text = ""
CmbNonNcdexClient.Text = ""
TxtSInvoiceNo.Text = ""
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
RsLog_InvoiceDeletetion!DocumentType = "NNNR"
RsLog_InvoiceDeletetion!DcoumentID = RsTxn_ProInvoice!InvoiceNo
RsLog_InvoiceDeletetion!SM_Prefix = RsTxn_ProInvoice!SM_Prefix
RsLog_InvoiceDeletetion!ClientID = RsTxn_ProInvoice!ClientID
RsLog_InvoiceDeletetion!Amount = RsTxn_ProInvoice!Amount
RsLog_InvoiceDeletetion!DeletetionRemark = Trim(TxtLogRemark.Text)
RsLog_InvoiceDeletetion!DeletedBy = Gen_UserLoginID
RsLog_InvoiceDeletetion!DeletedOn = Now
RsLog_InvoiceDeletetion.Update

Call DeleteChildTableDetail

RsTxn_ProInvoice.Delete
RsTxn_ProInvoice.Update
'RsTxn_ProInvoice.Requery
If RsTxn_ProInvoice.RecordCount = 0 Then
   Call ClearData
   Call GButtonLockAD(Me)
   'Call AddCmd_Click
   Exit Sub
End If
RsTxn_ProInvoice.MoveNext
If RsTxn_ProInvoice.EOF() Then
   RsTxn_ProInvoice.MoveLast
End If
Call Indata
Frame2.Visible = False
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

'If TxtReverseDate.Value < CDate(TxtFAGPostingDate.Text) Then
   
If TxtReverseDate.Value < CDate(TxtTIDate.Text) Then
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
RsTxn_ProInvoice!ReverseDate = TxtReverseDate.Value
RsTxn_ProInvoice!ReverseFlag = "Y"
RsTxn_ProInvoice!ReverseRemark = TxtReverseRemark.Text
RsTxn_ProInvoice!UpdatedBy = Gen_UserLoginID
RsTxn_ProInvoice!UpdatedDate = Now
RsTxn_ProInvoice.Update

tmp = " Update Txn_TIDetails Set TIAmount=0,ReverseFlag='Y',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where PIID=" & RsTxn_ProInvoice!InvoiceNo & " and PIFLag='NNNR'"
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


tmp = " Select TPI.InvoiceNo,TPI.InvoiceDate,TPI.BillingMonth,ML.LocationName,ME.ExchangeType,MC.ClientName, "
tmp = tmp & " TPI.Amount , TPI.Discount, TPI.Premium, TPI.BatchID, TPI.ServiceTaxNAVCode, TPI.FAGPostingDate, TPI.G_UID,TPI.Remarks,TPI.Flag, "
tmp = tmp & " TTD.TIID,TTD.TINo,TTD.TIDate,TTD.TIAmount,TTD.TISTAmount,TTD.EntryNo,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TTD.TIID)+isnull(TTD.AdjAmountFAG,0) + "
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TTD.TIID ) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TTD.TIID ) +"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TTD.TIID )+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TTD.TIID) 'ReceivedAmount','' 'BalanceAmount' ,TPI.ReverseFlag,ReverseRemark,ReverseDate,RBatchID,MC.Converted,TPI.HSNNo,TPI.AccountType,MSI.StateIndiaName ,MSI.GSTIN 'NCMLGSTNo',MCSG.GSTNo 'ClientGSTNo',MCSG.GSTNAVCode,'' 'TaxApplied'  "
tmp = tmp & " From Txn_ProInvoice TPI "
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID "
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID "
tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID "
tmp = tmp & " Left Join Txn_TIDetails TTD On TPI.InvoiceNo =  TTD.PIID And TTD.PIFlag = 'NNNR' " 'And TTD.TIAmount <> 0 "
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID "
tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID "
tmp = tmp & " Left  Join Mst_ClientStateWiseGST MCSG On MC.ClientID = MCSG.ClientID ANd MSI.StateIndiaID = MCSG.StateINdiaID "

If Gen_DBType = "MDB" Then
   wc = " Where TPI.InvoiceDate >= #" & TxtFromDate.Value & "# And TPI.InvoiceDate <= #" & TxtToDate.Value & "# "
Else
   wc = " Where TPI.InvoiceDate >= '" & TxtFromDate.Value & "' And TPI.InvoiceDate <= '" & TxtToDate.Value & "' "
End If
            
If CmbPreFix.Text <> "" Then
   wc = wc & " And  ML.LocationName = '" & CmbPreFix.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And ML." & Gen_WcLocation
   End If
End If
      
If ChkConverted.Value = 1 Then
   wc = wc & " And  MC.Converted = '1'"
End If
      
If CmbNonNcdexClient.Text <> "" Then
   wc = wc & " And  MC.ClientName = '" & CmbNonNcdexClient.Text & "'"
End If

If CmbSStatus.Text <> "" Then
   wc = wc & " And  TPI.Flag = '" & Left(CmbSStatus.Text, 1) & "'"
End If

If TxtSInvoiceNo.Text <> "" Then
   wc = wc & " And  TPI.InvoiceNo = '" & Replace(TxtSInvoiceNo.Text, "NN", "") & "'"
End If

If ChkBatchNo.Value = 0 Then
   wc = wc & " And  TPI.BatchID = 0 " '" & TxtSInvoiceNo.Text & ""
End If
            
tmp1 = " Order By ML.LocationName,TPI.InvoiceNo "

tmp = tmp & " " & wc & " " & tmp1

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(I, 0) = "NN" & MSHFlexGrid1.TextMatrix(I, 0)
       MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-mmm-yyyy")
       If MSHFlexGrid1.TextMatrix(I, 17) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 17) = Format(MSHFlexGrid1.TextMatrix(I, 17), "dd-mmm-yyyy")
       End If
       MSHFlexGrid1.TextMatrix(I, 2) = Format(MSHFlexGrid1.TextMatrix(I, 2), "mmm-yyyy")
       If MSHFlexGrid1.TextMatrix(I, 25) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 25) = Format(MSHFlexGrid1.TextMatrix(I, 25), "dd-mmm-yyyy")
       End If
       If MSHFlexGrid1.TextMatrix(I, 14) = "P" Then
          MSHFlexGrid1.TextMatrix(I, 14) = "Pending"
       ElseIf MSHFlexGrid1.TextMatrix(I, 14) = "C" Then
          MSHFlexGrid1.TextMatrix(I, 14) = "Confirmed"
       ElseIf MSHFlexGrid1.TextMatrix(I, 14) = "R" Then
          MSHFlexGrid1.TextMatrix(I, 14) = "Reject"
       ElseIf MSHFlexGrid1.TextMatrix(I, 14) = "F" Then
          MSHFlexGrid1.TextMatrix(I, 14) = "Freeze"
       End If
       If MSHFlexGrid1.TextMatrix(I, 23) = "Y" Then
          MSHFlexGrid1.TextMatrix(I, 23) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 23) = "No"
       End If
       If MSHFlexGrid1.TextMatrix(I, 27) = "1" Then
          MSHFlexGrid1.TextMatrix(I, 27) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 27) = "No"
       End If
       
       If Trim(MSHFlexGrid1.TextMatrix(I, 31)) <> "" And Trim(MSHFlexGrid1.TextMatrix(I, 32)) <> "" Then
          If Left(MSHFlexGrid1.TextMatrix(I, 31), 2) = Left(MSHFlexGrid1.TextMatrix(I, 32), 2) Then
             MSHFlexGrid1.TextMatrix(I, 34) = "Intra"
          Else
             MSHFlexGrid1.TextMatrix(I, 34) = "Inter"
          End If
       End If
       MSHFlexGrid1.TextMatrix(I, 22) = Format(Val(MSHFlexGrid1.TextMatrix(I, 18)) - Val(MSHFlexGrid1.TextMatrix(I, 21)), "#####0.00")
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
Dim fso As New FileSystemObject
On Error GoTo Error:
'Shell (App.Path & "\Invoices\" & TxtEntryNo & " ")
If mEntryNo <> "" Then
   wsPathAdobe = Gen_ReadPathConfigINI("PDFPATH") '"\\10.0.0.8\Reader\AcroRD32.exe" '"C:\Program Files\Adobe\Acrobat 6.0\Reader\AcroRd32.exe" '"\\10.0.0.8\Reader\AcroRD32.exe"
   mFileName = App.Path & "\NavisionInvoice\" & mEntryNo & ".pdf"
   'mFileName = App.Path & "\Invoices\" & Replace(TxtPostInvoiceNo, "/", "") & ".pdf"
   If fso.FileExists(mFileName) = True Then
      If wsPathAdobe <> "" Then
         iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
         Call Log_Print_History(TxtPreFix, "NNNR", TxtOriginalInvoiceNo, mEntryNo)
      End If
   Else
      MsgBox "Invoice not found. Please contact CMG Department for same!!!"
   End If
End If
Exit Sub

Error:
      MsgBox "Please contact Technology department!!!"

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
RsTxn_ProInvoice.MoveFirst
RsTxn_ProInvoice.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 12) & "'"
If Not RsTxn_ProInvoice.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub EditCmd_Click()
If Val(TxtBatchID) <> 0 Then
   MsgBox "Batch generated for this Bill of Estimation!!!"
   Exit Sub
End If

'LCase(TxtStatus) = "reject"
If LCase(TxtStatus) = "reject" Or LCase(TxtStatus) = "freeze" Then
   MsgBox "Bill of Estimation is " & CmbStatus.Text & ". You can not edit!!!"
   Exit Sub
End If

If LCase(TxtStatus) = "confirmed" Then
   If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "headcmg" Then
      MsgBox "Bill of Estimation is confirmed. For any modification please contact CMG Dept.!!!"
      Exit Sub
   End If
End If


If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtBillingMonth.Text, "yyyy-mm") Then
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
CmdPRint.Enabled = False
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
   CmdPRint.Enabled = False
   Exit Sub
End If


Call CmdClear_Click
If Gen_WcLocation <> "" Then
   Call TableTxn_ProInvoice(" Where " & Gen_WcLocation)
Else
   Call TableTxn_ProInvoice
End If

Call Grid_Head1

If RsTxn_ProInvoice.RecordCount = 0 Then
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
mLocationID = IIf(IsNull(RsTxn_ProInvoice!LocationID), 0, RsTxn_ProInvoice!LocationID)
TxtLocationID = GetLocationName(mLocationID)
TxtPreFix = RsTxn_ProInvoice!SM_Prefix
mInvoiceNo = RsTxn_ProInvoice!InvoiceNo
TxtOriginalInvoiceNo = RsTxn_ProInvoice!InvoiceNo
TxtInvoiceNo = "NN" & RsTxn_ProInvoice!InvoiceNo
TxtInvoiceDate = Format(RsTxn_ProInvoice!InvoiceDate, "dd-mmm-yyyy")
mClientIDRow = RsTxn_ProInvoice!ClientID
TxtClientID = GetClientName(IIf(IsNull(RsTxn_ProInvoice!ClientID), "0", RsTxn_ProInvoice!ClientID))

TxtBillingMonth = Format(RsTxn_ProInvoice!BillingMonth, "mmm-yyyy")
TxtAmount = RsTxn_ProInvoice!Amount
TxtAdHocAmt = GetADhocInvoiceAmount(mLocationID, mClientIDRow)

If IsNull(RsTxn_ProInvoice!FAGPostingDate) = False Then
   TxtFAGPostingDate = Format(RsTxn_ProInvoice!FAGPostingDate, "dd-mmm-yyyy")
Else
   TxtFAGPostingDate = ""
End If

TxtFAGRemarks = ""  'RsTxn_ProInvoice!Remark
TxtServiceTaxNAVCode = IIf(IsNull(RsTxn_ProInvoice!ServiceTaxNAVCode), "", RsTxn_ProInvoice!ServiceTaxNAVCode)

mExchangeTypeID = IIf(IsNull(RsTxn_ProInvoice!ExchangeTypeID), 0, RsTxn_ProInvoice!ExchangeTypeID)
If mExchangeTypeID <> 0 Then
   TxtExchangeType = GetExchangeName(mExchangeTypeID)
Else
   TxtExchangeType = ""
End If
TxtBatchID = IIf(IsNull(RsTxn_ProInvoice!BatchID), 0, RsTxn_ProInvoice!BatchID)

TxtDiscount = IIf(IsNull(RsTxn_ProInvoice!Discount), 0, RsTxn_ProInvoice!Discount)
TxtPremium = IIf(IsNull(RsTxn_ProInvoice!Premium), 0, RsTxn_ProInvoice!Premium)

TxtHSNNo = IIf(IsNull(RsTxn_ProInvoice!HSNNo), 0, RsTxn_ProInvoice!HSNNo)
TxtType = IIf(IsNull(RsTxn_ProInvoice!AccountType), "", RsTxn_ProInvoice!AccountType)


TxtDiscount.Locked = True
TxtPremium.Locked = True
TxtRemarks.Locked = True
CmbStatus.Visible = False
TxtStatus.Visible = True


mFlag = IIf(IsNull(RsTxn_ProInvoice!Flag), "P", RsTxn_ProInvoice!Flag)


If mFlag = "P" Then
   TxtStatus = "Pending"
ElseIf mFlag = "C" Then
   TxtStatus = "Confirmed"
ElseIf mFlag = "R" Then
   TxtStatus = "Reject"
ElseIf mFlag = "F" Then
   TxtStatus = "Freeze"
End If


TxtRejectionRemark.Locked = True
TxtRemarks = IIf(IsNull(RsTxn_ProInvoice!Remarks), "", RsTxn_ProInvoice!Remarks)
TxtRejectionRemark = IIf(IsNull(RsTxn_ProInvoice!RejectionRemarks), "", RsTxn_ProInvoice!RejectionRemarks)
TxtG_UID = IIf(IsNull(RsTxn_ProInvoice!G_UID), "", RsTxn_ProInvoice!G_UID)

TxtCreated = IIf(IsNull(RsTxn_ProInvoice!CreatedBy), "", RsTxn_ProInvoice!CreatedBy) & " :- " & IIf(IsNull(RsTxn_ProInvoice!CreatedDate), "", Format(RsTxn_ProInvoice!CreatedDate, "dd-mmm-yyyy hh:mm:ss"))
TxtUpdated = IIf(IsNull(RsTxn_ProInvoice!UpdatedBy), "", RsTxn_ProInvoice!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_ProInvoice!UpdatedDate), "", Format(RsTxn_ProInvoice!UpdatedDate, "dd-mmm-yyyy hh:mm:ss"))

Call IndataGSLDetails
Call IndataTIDetail
CmdPRint.Enabled = True
CmdViewInvoice.Enabled = True

If RsTxn_ProInvoice.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub
Private Sub IndataGSLDetails()
Dim mYear, mMonth, mFromDate As Variant
Call Grid_Head1

mYear = Year(RsTxn_ProInvoice!InvoiceDate)
mMonth = Month(RsTxn_ProInvoice!InvoiceDate)
mFromDate = Format(mYear & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

'tmp = "Select MC.CMPName,TIMGD.GSLID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo,MComm.Commodity, " & _
'      " TIMGD.BalanceBags,TIMGD.BalanceWeight,TIMGD.TFlag,TIMGD.CxTFNo,MBC.BillingCycle,TIMGD.BillingUnit,TIMGD.NoofDays,TIMGD.Rate,MBC1.BillingCycle,TIMGD.OldBillingUnit,TIMGD.OldRate, " & _
'      " TIMGD.InvoiceAmount,TIMGD.UpdatedBy ,TIMGD.UpdatedDate,TIMGD.CxTFDate " & _
'      " From Txn_InvMonGSLDetail TIMGD " & _
'      " Inner Join Txn_ProInvoice TPI On TIMGD.PINo = TPI.InvoiceNo And TIMGD.SM_Prefix = TPI.SM_Prefix " & _
'      " Inner Join Mst_CMP MC On TIMGD.CMPID = MC.CMPID " & _
'      " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID " & _
'      " Inner Join Mst_Commodity MComm On TIMGD.CommodityID = MComm.CommodityID " & _
'      " Inner Join Mst_BillingCycle MBC On TIMGD.BillingCycleID = MBC.BillingCycleID " & _
'      " Left Join Mst_BillingCycle MBC1 On TIMGD.OldBillingCycleID = MBC1.BillingCycleID " & _
'      " Where TIMGD.PINo = " & mInvoiceNo & " And TIMGD.SM_prefix = '" & TxtPreFix & "' And TIMGD.ClientIDRow = " & mClientIDRow & " " & _
'      " Order By MC.CMPName,MComm.Commodity,TIMGD.GSLID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo "

tmp = "Select CMPName,Mst_GSL_ID,GodownNo,StackNo,LotNo,Commodity,"
tmp = tmp & " BalanceBags,BalanceWeight,TFlag,CxTFNo,BillingCycle,BillingUnit,NoofDays,Rate,OldBillingCycle,OldBillingUnit,OldRate,"
tmp = tmp & " InvoiceAmount,UpdatedBy ,UpdatedDate,CxTFDate From ("
tmp = tmp & " Select MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo,MComm.Commodity,"
tmp = tmp & " TIMGD.BalanceBags,TIMGD.BalanceWeight,TIMGD.TFlag,TIMGD.CxTFNo,MBC.BillingCycle,TIMGD.BillingUnit,TIMGD.NoofDays,TIMGD.Rate,MBC1.BillingCycle 'OldBillingCycle',TIMGD.OldBillingUnit,TIMGD.OldRate,"
tmp = tmp & " TIMGD.InvoiceAmount , TIMGD.UpdatedBy, TIMGD.UpdatedDate, TIMGD.CxTFDate"
tmp = tmp & " From Txn_InvMonGSLDetail TIMGD"
tmp = tmp & " Inner Join Txn_ProInvoice TPI On TIMGD.PINo = TPI.InvoiceNo And TIMGD.SM_Prefix = TPI.SM_Prefix"
tmp = tmp & " Inner Join Mst_CMP MC On TIMGD.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Commodity MComm On TIMGD.CommodityID = MComm.CommodityID"
tmp = tmp & " Inner Join Mst_BillingCycle MBC On TIMGD.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Left Join Mst_BillingCycle MBC1 On TIMGD.OldBillingCycleID = MBC1.BillingCycleID"
tmp = tmp & " Where TIMGD.PINo = " & mInvoiceNo & " And TIMGD.SM_prefix = '" & TxtPreFix & "' And TIMGD.ClientID = " & mClientIDRow & ""
tmp = tmp & " Union All"
tmp = tmp & " Select MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo,MComm.Commodity,"
tmp = tmp & " TIMGD.BalanceBags,TIMGD.BalanceWeight,TIMGD.TFlag,TIMGD.CxTFNo,MBC.BillingCycle,TIMGD.BillingUnit,TIMGD.NoofDays,TIMGD.Rate,MBC1.BillingCycle 'OldBillingCycle',TIMGD.OldBillingUnit,TIMGD.OldRate,"
tmp = tmp & " TIMGD.InvoiceAmount , TIMGD.UpdatedBy, TIMGD.UpdatedDate, TIMGD.CxTFDate"
tmp = tmp & " From Txn_InvMonGSLDetail2014to2015 TIMGD"
tmp = tmp & " Inner Join Txn_ProInvoice TPI On TIMGD.PINo = TPI.InvoiceNo And TIMGD.SM_Prefix = TPI.SM_Prefix"
tmp = tmp & " Inner Join Mst_CMP MC On TIMGD.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Commodity MComm On TIMGD.CommodityID = MComm.CommodityID"
tmp = tmp & " Inner Join Mst_BillingCycle MBC On TIMGD.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Left Join Mst_BillingCycle MBC1 On TIMGD.OldBillingCycleID = MBC1.BillingCycleID"
tmp = tmp & " Where TIMGD.PINo = " & mInvoiceNo & " And TIMGD.SM_prefix = '" & TxtPreFix & "' And TIMGD.ClientID = " & mClientIDRow & ""
tmp = tmp & " Union All"
tmp = tmp & " Select MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo,MComm.Commodity,"
tmp = tmp & " TIMGD.BalanceBags,TIMGD.BalanceWeight,TIMGD.TFlag,TIMGD.CxTFNo,MBC.BillingCycle,TIMGD.BillingUnit,TIMGD.NoofDays,TIMGD.Rate,MBC1.BillingCycle 'OldBillingCycle',TIMGD.OldBillingUnit,TIMGD.OldRate,"
tmp = tmp & " TIMGD.InvoiceAmount , TIMGD.UpdatedBy, TIMGD.UpdatedDate, TIMGD.CxTFDate"
tmp = tmp & " From Txn_InvMonGSLDetail2016 TIMGD"
tmp = tmp & " Inner Join Txn_ProInvoice TPI On TIMGD.PINo = TPI.InvoiceNo And TIMGD.SM_Prefix = TPI.SM_Prefix"
tmp = tmp & " Inner Join Mst_CMP MC On TIMGD.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Commodity MComm On TIMGD.CommodityID = MComm.CommodityID"
tmp = tmp & " Inner Join Mst_BillingCycle MBC On TIMGD.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Left Join Mst_BillingCycle MBC1 On TIMGD.OldBillingCycleID = MBC1.BillingCycleID"
tmp = tmp & " Where TIMGD.PINo = " & mInvoiceNo & " And TIMGD.SM_prefix = '" & TxtPreFix & "' And TIMGD.ClientID = " & mClientIDRow & ""
tmp = tmp & " Union All"
tmp = tmp & " Select MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo,MComm.Commodity,"
tmp = tmp & " TIMGD.BalanceBags,TIMGD.BalanceWeight,TIMGD.TFlag,TIMGD.CxTFNo,MBC.BillingCycle,TIMGD.BillingUnit,TIMGD.NoofDays,TIMGD.Rate,MBC1.BillingCycle 'OldBillingCycle',TIMGD.OldBillingUnit,TIMGD.OldRate,"
tmp = tmp & " TIMGD.InvoiceAmount , TIMGD.UpdatedBy, TIMGD.UpdatedDate, TIMGD.CxTFDate"
tmp = tmp & " From Txn_InvMonGSLDetailUpto2013 TIMGD"
tmp = tmp & " Inner Join Txn_ProInvoice TPI On TIMGD.PINo = TPI.InvoiceNo And TIMGD.SM_Prefix = TPI.SM_Prefix"
tmp = tmp & " Inner Join Mst_CMP MC On TIMGD.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Commodity MComm On TIMGD.CommodityID = MComm.CommodityID"
tmp = tmp & " Inner Join Mst_BillingCycle MBC On TIMGD.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Left Join Mst_BillingCycle MBC1 On TIMGD.OldBillingCycleID = MBC1.BillingCycleID"
tmp = tmp & " Where TIMGD.PINo = " & mInvoiceNo & " And TIMGD.SM_prefix = '" & TxtPreFix & "' And TIMGD.ClientID = " & mClientIDRow & ""
tmp = tmp & " ) A"
tmp = tmp & " Order By CMPName,Commodity,Mst_GSL_ID,GodownNo,StackNo,LotNo"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs.RecordCount + 3
   x = 2
   For I = 1 To TmpRs.RecordCount
       'MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       If LCase(MSHFlexGrid2.TextMatrix(x, 8)) = "o" Then
          MSHFlexGrid2.TextMatrix(x, 8) = "Op. Balance"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 8)) = "d" Then
          MSHFlexGrid2.TextMatrix(x, 8) = "Deposit"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 8)) = "w" Then
          MSHFlexGrid2.TextMatrix(x, 8) = "Withdrawal"
       End If
       
       If LCase(MSHFlexGrid2.TextMatrix(x, 11)) = "q" Then
          MSHFlexGrid2.TextMatrix(x, 11) = "Quantity"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 11)) = "b" Then
          MSHFlexGrid2.TextMatrix(x, 11) = "Bags"
       End If
       
       If LCase(MSHFlexGrid2.TextMatrix(x, 15)) = "q" Then
          MSHFlexGrid2.TextMatrix(x, 15) = "Quantity"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 15)) = "b" Then
          MSHFlexGrid2.TextMatrix(x, 15) = "Bags"
       End If
       
       MSHFlexGrid2.TextMatrix(0, 17) = Val(MSHFlexGrid2.TextMatrix(0, 17)) + Val(MSHFlexGrid2.TextMatrix(x, 17))
       
       If MSHFlexGrid2.TextMatrix(x, 19) <> "" Then
          MSHFlexGrid2.TextMatrix(x, 19) = Format(MSHFlexGrid2.TextMatrix(x, 19), "dd-mmm-yyyy hh:mm:ss")
       End If
       
       If MSHFlexGrid2.TextMatrix(x, 20) <> "" Then
          MSHFlexGrid2.TextMatrix(x, 20) = Format(MSHFlexGrid2.TextMatrix(x, 20), "dd-mmm-yyyy")
       End If
       
       If MSHFlexGrid2.TextMatrix(x, 14) <> "" Or MSHFlexGrid2.TextMatrix(x, 15) <> "" Or MSHFlexGrid2.TextMatrix(x, 16) <> "" Then
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.Col = C
              MSHFlexGrid2.row = x
              MSHFlexGrid2.CellBackColor = vbRed
          Next
       End If
       
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
RsTxn_ProInvoice.MoveNext
If RsTxn_ProInvoice.EOF Then
   RsTxn_ProInvoice.MoveLast
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
RsTxn_ProInvoice.MovePrevious
If RsTxn_ProInvoice.BOF Then
   RsTxn_ProInvoice.MoveFirst
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
RsTxn_ProInvoice.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_ProInvoice.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SaveCmd_Click()

If TxtInvoiceNo = "" Then
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

If LCase(CmbStatus.Text) <> "pending" Then ' Or LCase(CmbStatus.Text) = "confirm" Then
   If LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "power" Then
      MsgBox "For modification in " & CmbStatus.Text & " Invoices contact CMG Dept.!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

If LCase(CmbStatus.Text) = "reject" Or LCase(CmbStatus.Text) = "freeze" Then
   If Trim(TxtRejectionRemark) = "" Then
      MsgBox "Blank rejection/freeze remarks not allowed!!!"
      TxtRejectionRemark.SetFocus
      Exit Sub
   End If
End If

If TxtRemarks.Text = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtRemarks.SetFocus
   Exit Sub
End If

If LCase(CmbStatus.Text) = "confirmed" Then
   If MSHFlexGrid2.Rows <= 3 Then
      tmp = "Select * from Txn_MonthwiseInvoiceData Where PINo = " & Val(TxtOriginalInvoiceNo) & " "
      Call Tmp7Table
      If TmpRs7.RecordCount = 0 Then
         MsgBox "GSL Detail not found!!!"
         CmbStatus.SetFocus
         Exit Sub
      End If
   'ElseIf Format(Val(MSHFlexGrid2.TextMatrix(0, 17)), "#######") <> Format(Val(TxtAmount.Text), "#######") Then
   ElseIf Abs(Val(MSHFlexGrid2.TextMatrix(0, 17)) - Val(TxtAmount.Text)) > 1 Then
      MsgBox "GSL amount and Invoice amount does not matched!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If

    If (Val(TxtAmount.Text) - Val(TxtDiscount)) + Val(TxtPremium) > 0 Then
       If GetStateWiseNAVCode(mClientIDRow, mLocationID, "N") = "" Then
          MsgBox "Nav code mapping missing in State Wise GST mapping screen!!!"
          CmbStatus.SetFocus
          Exit Sub
       End If
    End If
End If


RsTxn_ProInvoice.MoveFirst
RsTxn_ProInvoice.Find "G_UID = '" & TxtG_UID & "'"

If RsTxn_ProInvoice.EOF = False Then
   RsTxn_ProInvoice!Discount = Val(TxtDiscount)
   RsTxn_ProInvoice!Premium = Val(TxtPremium)
   RsTxn_ProInvoice!Remarks = Trim(TxtRemarks)
   RsTxn_ProInvoice!Flag = Left(CmbStatus.Text, 1)
   RsTxn_ProInvoice!RejectionRemarks = Trim(TxtRejectionRemark)
   RsTxn_ProInvoice!UpdatedBy = Gen_UserLoginID
   RsTxn_ProInvoice!UpdatedDate = Now
   RsTxn_ProInvoice.Update
End If

RsTxn_ProInvoice.MoveFirst
RsTxn_ProInvoice.Find "G_UID = '" & TxtG_UID & "'"


Call Indata


End Sub

Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_ProInvoice.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsTxn_ProInvoice.EOF Then
      RsTxn_ProInvoice.MoveFirst
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
MSHFlexGrid1.Height = Frame0.Height - 1300
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
MSHFlexGrid1.Cols = 35 '33 '32 '31 '28 '26


MSHFlexGrid1.TextMatrix(0, 0) = "Invoice No"
MSHFlexGrid1.TextMatrix(0, 1) = "Invoice Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Billing Month"
MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(0, 4) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 5) = "Client"
MSHFlexGrid1.TextMatrix(0, 6) = "Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "Discount"
MSHFlexGrid1.TextMatrix(0, 8) = "Premium"
MSHFlexGrid1.TextMatrix(0, 9) = "Batch ID"
MSHFlexGrid1.TextMatrix(0, 10) = "GST Group Code" '"Service Tax NAVCode"
MSHFlexGrid1.TextMatrix(0, 11) = "FAGPosting Date"
MSHFlexGrid1.TextMatrix(0, 12) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 13) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 14) = "Status"
MSHFlexGrid1.TextMatrix(0, 15) = "TI ID"
MSHFlexGrid1.TextMatrix(0, 16) = "TI No"
MSHFlexGrid1.TextMatrix(0, 17) = "TI Date"
MSHFlexGrid1.TextMatrix(0, 18) = "TI Amount"
MSHFlexGrid1.TextMatrix(0, 19) = "TI Service Tax Amount"
MSHFlexGrid1.TextMatrix(0, 20) = "TI Entry No"
MSHFlexGrid1.TextMatrix(0, 21) = "Received Amount"
MSHFlexGrid1.TextMatrix(0, 22) = "Balance Amount"
MSHFlexGrid1.TextMatrix(0, 23) = "Reverse Flag"
MSHFlexGrid1.TextMatrix(0, 24) = "Reverse Remark"
MSHFlexGrid1.TextMatrix(0, 25) = "Reverse Date"
MSHFlexGrid1.TextMatrix(0, 26) = "RBatch ID"
MSHFlexGrid1.TextMatrix(0, 27) = "Converted Client"
MSHFlexGrid1.TextMatrix(0, 28) = "HSN No"
MSHFlexGrid1.TextMatrix(0, 29) = "Account Type"
MSHFlexGrid1.TextMatrix(0, 30) = "State India Name"
MSHFlexGrid1.TextMatrix(0, 31) = "NCML GST No"
MSHFlexGrid1.TextMatrix(0, 32) = "Client GST No"
MSHFlexGrid1.TextMatrix(0, 33) = "Client GST NAVCode"
MSHFlexGrid1.TextMatrix(0, 34) = "Tax Applied (Inter/Intra)"

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 3200
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 0
MSHFlexGrid1.ColWidth(13) = 3000
MSHFlexGrid1.ColWidth(14) = 1600
MSHFlexGrid1.ColWidth(15) = 1600
MSHFlexGrid1.ColWidth(16) = 1600
MSHFlexGrid1.ColWidth(17) = 1600
MSHFlexGrid1.ColWidth(18) = 1600
MSHFlexGrid1.ColWidth(19) = 1600
MSHFlexGrid1.ColWidth(20) = 1600

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
If ChkConverted.Value = 0 Then
   Call TableMst_Client("where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
Else
   Call TableMst_Client("where ExchangeTypeId = 1 And COnverted = '1' ") '" & mExchangeTypeID)
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_ProInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_ProInvoice.xls")
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
           If C = 29 Or C = 30 Then
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

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
If Val(TxtBatchID) <> 0 Then
   MsgBox "Batch generated for this Bill of Estimation!!!"
   Exit Sub
End If

'If TxtStatus <> "Pending" Then
'   MsgBox "Batch generated for this Bill of Estimation!!!"
'   Exit Sub
'End If


ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   Frame2.Visible = True
   TxtLogRemark.Text = ""

'   Call DeleteChildTableDetail
'   RsTxn_ProInvoice.Delete
'   RsTxn_ProInvoice.Update
'   'RsTxn_ProInvoice.Requery
'   If RsTxn_ProInvoice.RecordCount = 0 Then
'      Call ClearData
'      Call GButtonLockAD(Me)
'      'Call AddCmd_Click
'      Exit Sub
'   End If
'   RsTxn_ProInvoice.MoveNext
'   If RsTxn_ProInvoice.EOF() Then
'      RsTxn_ProInvoice.MoveLast
'   End If
'   Call InData
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_ProInvoice.CancelUpdate
End Sub
Private Sub ClearData()

TxtLocationID = ""
TxtPreFix = ""
TxtInvoiceNo = ""
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

'tmp = "Update Txn_InvMonGSLDetail Set PINo = Null,InvoiceDate = Null Where PINo = " & TxtOriginalInvoiceNo & " And SM_Prefix = '" & TxtPreFix & "'"
'Call TmpTable
tmp = "Select * From Txn_InvMonGSLDetail Where YearMOnth = '" & mSFromDate & "' And PINo = " & TxtOriginalInvoiceNo & " And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
For j = 1 To TmpRs.RecordCount
    TmpRs!PINo = Null
    TmpRs!InvoiceDate = Null
    TmpRs.Update
    TmpRs.MoveNext
Next

tmp = "Delete From Txn_MonthwiseInvoiceData Where PINo = " & TxtOriginalInvoiceNo & " "
Call TmpTable

End Sub


Private Sub Findwithrdrawal(SM_Prefix_ As Variant, mFromDate_ As Variant, x_ As Variant, gslID_ As Variant)
Dim Retval As Double

tmp = "Select Sum(TI.BalanceBags) 'With_Bal_Bags',Sum(TI.BalanceWeight) 'With_Bal_Wt' "
tmp = tmp & " From Txn_InvMonGSLDetail TI"
tmp = tmp & " Where  TI.YearMonth='" & mFromDate_ & "' and  TI.SM_Prefix='" & SM_Prefix_ & "' And TI.GSLID = " & gslID_ & " And TI.TFlag='W' " 'And TI.CMPID =" & CMPID_ & " And TI.CommodityID=" & CommodityID_ & " "
    
    
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
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
End Sub

Private Sub TxtToDate_LostFocus()
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
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

'Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_ProInvoiceGSLwiseDetails.xls")

Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "_BOE-" & Trim(TxtPreFix) & "" & Trim(TxtInvoiceNo) & ".xls")

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
 
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_BOE-" & Trim(TxtPreFix) & "" & Trim(TxtInvoiceNo) & ".xls")

I = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
oWS.Name = TxtPreFix & "-" & TxtInvoiceNo  'MSHFlexGrid2.TextMatrix(i, 23)
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
tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtInvoiceNo.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 5) = "Invoice Date"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).WrapText = True
tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 6) = TxtInvoiceDate.Text
oWS.Cells(mRow, C + 6).HorizontalAlignment = 2
oWS.Cells(mRow, C + 6).Font.Size = 7


oWS.Cells(mRow, C + 8) = "Exchange Type"
oWS.Cells(mRow, C + 8).Font.Bold = True
oWS.Cells(mRow, C + 8).WrapText = True
tmp = "i" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 8).Font.Size = 7
oWS.Cells(mRow, C + 9) = TxtExchangeType.Text
oWS.Cells(mRow, C + 9).HorizontalAlignment = 2
oWS.Cells(mRow, C + 9).Font.Size = 7

oWS.Cells(mRow, C + 12) = "Billing Month"
oWS.Cells(mRow, C + 12).Font.Bold = True
oWS.Cells(mRow, C + 12).WrapText = True
tmp = "m" & mRow & ":" & "o" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 12).Font.Size = 7
oWS.Cells(mRow, C + 13) = TxtBillingMonth.Text
oWS.Cells(mRow, C + 13).HorizontalAlignment = 2
oWS.Cells(mRow, C + 13).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Location"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtLocationID.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 8) = "Client Name"
oWS.Cells(mRow, C + 8).Font.Bold = True
oWS.Cells(mRow, C + 8).WrapText = True
tmp = "i" & mRow & ":" & "o" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 8).Font.Size = 7
oWS.Cells(mRow, C + 9) = TxtClientID.Text
oWS.Cells(mRow, C + 9).HorizontalAlignment = 2
oWS.Cells(mRow, C + 9).Font.Size = 7


mRow = mRow + 1
oWS.Cells(mRow, C + 1) = "Amount"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtAmount.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 5) = "ST Nav Code"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).WrapText = True
tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 6) = TxtServiceTaxNAVCode.Text
oWS.Cells(mRow, C + 6).HorizontalAlignment = 2
oWS.Cells(mRow, C + 6).Font.Size = 7


oWS.Cells(mRow, C + 8) = "FAG Posting Date"
oWS.Cells(mRow, C + 8).Font.Bold = True
oWS.Cells(mRow, C + 8).WrapText = True
tmp = "i" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 8).Font.Size = 7
oWS.Cells(mRow, C + 9) = TxtFAGPostingDate.Text
oWS.Cells(mRow, C + 9).HorizontalAlignment = 2
oWS.Cells(mRow, C + 9).Font.Size = 7


oWS.Cells(mRow, C + 12) = "Batch ID"
oWS.Cells(mRow, C + 12).Font.Bold = True
oWS.Cells(mRow, C + 12).WrapText = True
tmp = "m" & mRow & ":" & "o" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 12).Font.Size = 7
oWS.Cells(mRow, C + 13) = TxtBatchID.Text
oWS.Cells(mRow, C + 13).HorizontalAlignment = 2
oWS.Cells(mRow, C + 13).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "FAG Remark"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "o" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtFAGRemarks.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7


mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Discount"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtDiscount.Text
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 5) = "Premium"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).WrapText = True
tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 6) = TxtPremium.Text
oWS.Cells(mRow, C + 6).HorizontalAlignment = 2
oWS.Cells(mRow, C + 6).Font.Size = 7

tmp = "h" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 8) = "Total BOE Amount(Amt + Premium) - Discount"
oWS.Cells(mRow, C + 8).Font.Bold = True
oWS.Cells(mRow, C + 8).WrapText = True
oWS.Cells(mRow, C + 8).Font.Size = 7

tmp = "l" & mRow & ":" & "o" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 12) = (Val(TxtAmount) + Val(TxtPremium)) - Val(TxtDiscount)
oWS.Cells(mRow, C + 12).HorizontalAlignment = 2
oWS.Cells(mRow, C + 12).Font.Size = 7


mRow = mRow + 1

tmp = "a" & mRow & ":" & "o" & mRow
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

tmp = "a" & mRow & ":" & "n" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Total"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True

oWS.Cells(mRow, C + 15) = MSHFlexGrid2.TextMatrix(0, 14)
oWS.Cells(mRow, C + 15).Font.Size = 7


For Cell = 3 To mRow
    tmp = "a" & Cell & ":" & "o" & Cell
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
'oWB.Close
'oXL.Quit
MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub
Private Sub IndataTIDetail()

''tmp = " Select TDI.TINO,TDI.TIAmount,TDI.TIDate,TDI.TIID,TDI.EntryNo From Txn_TIDetails TDI"
''tmp = tmp & " Where PIID=" & RsTxn_ProInvoice!InvoiceNo & " and TDI.PIFLag='NNNR'"
''
''Call TmpTable
''
''If TmpRs.RecordCount > 0 Then
''   TxtTINo = IIf(IsNull(TmpRs!TINO), "", TmpRs!TINO)
''   TxtTIDate = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-mm-yyyy"))
''   TxtTIAmount = IIf(IsNull(TmpRs!TIAmount), "", TmpRs!TIAmount)
''   mEntryNo = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
''   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fag admin" Or LCase(Gen_UserRole) = "power" Then
''      If IIf(IsNull(RsTxn_ProInvoice!ReverseFlag), "N", RsTxn_ProInvoice!ReverseFlag) <> "Y" Then
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

ans = GetTIDetails(RsTxn_ProInvoice!InvoiceNo, "NNNR")
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
If TxtEntryNo <> "" And (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "admin") Or LCase(Gen_UserRole) = "headcmg" Then
   If IIf(IsNull(RsTxn_ProInvoice!ReverseFlag), "N", RsTxn_ProInvoice!ReverseFlag) <> "Y" Then
      If CheckTIInCR(Val(ans(0))) = False Then
         CmdReverse.Enabled = True
      End If
   End If
End If

End Sub


Private Function GetADhocInvoiceAmount(mLocID_ As Double, mClientID_ As Double)
Dim Retval As Variant
Retval = 0

tmp = "Select LocationID,ClientID,Sum(Amount) As 'Amount' From Txn_TempInvoice "
tmp = tmp & " Where IETypeID = 7 And LocationID = " & mLocID_ & " And ClientID = " & mClientID_ & " And Month(InvoiceDate) = " & Month(CDate(TxtBillingMonth)) & "  And Year(InvoiceDate) = " & Year(CDate(TxtBillingMonth)) & " And Flag <> 'R' "
tmp = tmp & " Group By LocationID,ClientID "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Retval = TmpRs1!Amount
End If

GetADhocInvoiceAmount = Retval

End Function


